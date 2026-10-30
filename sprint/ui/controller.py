"""The window's state and the work behind it: reading files, page pictures, printing.

The controller lives on the UI thread. Workers run on their own threads and hand
their results back through ``Bridge`` signals, which Qt delivers on the UI thread.
"""

from __future__ import annotations

import logging
import os
import queue
import threading
from dataclasses import dataclass, field
from typing import Callable, Optional

from PySide6.QtCore import QObject, QTimer, Signal
from PySide6.QtGui import QImage, QPixmap

from .. import jobs, printing
from ..i18n import count, join_and, t
from ..pdfdocs import read_document, render_preview
from ..plan import (A3_DEST, A4_DEST, PLOT_DEST, Entry, Options, Part, Plan, build_plan, mark_done,
                    reset_done)
from ..paper import ROLL_WIDTHS_MM
from ..settings import Settings

log = logging.getLogger(__name__)

DESTS = (A4_DEST, A3_DEST, PLOT_DEST)
MAX_COPIES = 99
BIG_PREVIEW_PX = 400     # pictures this large take memory: only the latest few are kept
BIG_PREVIEWS_KEPT = 8
DONE_SHOW_MS = 2200


def _is_pdf(path: str) -> bool:
    return path.lower().endswith(".pdf") and os.path.isfile(path)


def expand_paths(paths) -> list[str]:
    """PDF files among ``paths``; folders contribute the PDFs directly inside them."""
    result = []
    for path in paths:
        path = os.path.normpath(path)
        if os.path.isdir(path):
            try:
                names = sorted(os.listdir(path), key=str.casefold)
            except OSError:
                continue
            result += [os.path.join(path, n) for n in names if _is_pdf(os.path.join(path, n))]
        elif _is_pdf(path):
            result.append(path)
    return result


def file_stamp(path: str) -> Optional[tuple]:
    try:
        stat = os.stat(path)
    except OSError:
        return None
    return stat.st_mtime_ns, stat.st_size


@dataclass
class Services:
    """What the controller needs from Windows; tests pass fakes."""
    list_printers: Callable[[], list[str]] = printing.list_printers
    detect_printers: Callable[[list[str]], dict] = printing.detect_printers
    printer_exists: Callable[[str], bool] = printing.printer_exists
    supports_duplex: Callable[[str], bool] = printing.supports_duplex
    default_duplex: Callable[[str], bool] = printing.default_duplex
    open_preferences: Callable[[str], None] = printing.open_printer_preferences
    backend: printing.Backend = field(default_factory=printing.DirectBackend)
    plot_backend: printing.Backend = field(default_factory=printing.DirectBackend)


class Bridge(QObject):
    """Signals the worker threads emit."""
    doc_read = Signal(object, object)          # entry, DocInfo
    preview = Signal(str, int, int, QImage)    # path, page index, size, picture
    progress = Signal(str, int, int)           # destination, pages handled, pages
    part_sent = Signal(object)                 # Part
    log = Signal(str, str)                     # message, level
    job_done = Signal(int, bool, int)          # pages sent, cancelled, problems
    call = Signal(object)                      # run this callable on the UI thread


class Controller(QObject):
    changed = Signal()                 # anything shown may have changed
    entries_changed = Signal()         # files were added or removed
    preview_ready = Signal(str, int, int)   # path, page index, size: picture in the cache
    toast = Signal(str, str)           # message, kind ("ok", "info", "error")
    log_line = Signal(str, str)        # message, level
    printers_changed = Signal()        # the list of installed printers

    def __init__(self, settings: Settings, services: Services, temp_dir: str, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.services = services
        self.temp_dir = temp_dir
        self.entries: list[Entry] = []
        roll = settings.roll if settings.roll in ROLL_WIDTHS_MM else 610
        self.opts = Options(a3_to_plotter=settings.a3_to_plotter, roll=roll)
        self.printers = jobs.Printers(settings.printer_a4, settings.printer_a3, settings.printer_plotter)
        self.copies = 1
        self.duplex = {A4_DEST: False, A3_DEST: False}
        self.duplex_supported = {A4_DEST: False, A3_DEST: False}
        self.installed: list[str] = []
        self.running: Optional[str] = None   # the destination being printed
        self.waiting: list[str] = []         # destinations to print after it
        self.just_sent: set[str] = set()     # printed a moment ago
        self.totals = dict.fromkeys(DESTS, 0)
        self.handled = dict.fromkeys(DESTS, 0)
        self.previews: dict[tuple[str, int, int], QPixmap] = {}   # (path, page index, size)
        self.log_lines: list[tuple[str, str]] = []
        self._plan: Optional[Plan] = None
        self._cancel = threading.Event()

        self.bridge = Bridge()
        self.bridge.doc_read.connect(self._on_doc_read)
        self.bridge.preview.connect(self._on_preview)
        self.bridge.progress.connect(self._on_progress)
        self.bridge.part_sent.connect(self._on_part_sent)
        self.bridge.log.connect(self.log)
        self.bridge.job_done.connect(self._on_job_done)
        self.bridge.call.connect(lambda fn: fn())

        self._reads: "queue.Queue[Entry]" = queue.Queue()
        self._pending_reads = 0
        threading.Thread(target=self._read_worker, name="read", daemon=True).start()
        self._preview_requests: queue.PriorityQueue = queue.PriorityQueue()
        self._preview_wanted: set[tuple[str, int, int]] = set()
        self._preview_count = 0
        threading.Thread(target=self._preview_worker, name="preview", daemon=True).start()
        self._job_thread: Optional[threading.Thread] = None

    # ── State ────────────────────────────────────────────────────────────────

    @property
    def busy(self) -> bool:
        return self.running is not None or bool(self.waiting)

    @property
    def reading(self) -> int:
        return self._pending_reads

    def plan(self) -> Plan:
        if self._plan is None:
            self.opts.duplex_a4 = self.duplex[A4_DEST] and self.duplex_supported[A4_DEST]
            self._plan = build_plan(self.entries, self.opts)
        return self._plan

    def _changed(self) -> None:
        self._plan = None
        self.changed.emit()

    def printer_for(self, dest: str) -> str:
        return {A4_DEST: self.printers.a4, A3_DEST: self.printers.a3, PLOT_DEST: self.printers.plotter}[dest]

    @staticmethod
    def pages_for(plan: Plan, dest: str) -> int:
        """The pages ``dest``'s print button sends."""
        return {A4_DEST: plan.a4.pending, A3_DEST: plan.a3.pending, PLOT_DEST: plan.plot_included}[dest]

    def status(self, dest: str) -> str:
        """'printing', 'waiting' (for the one printing), 'sent' (a moment ago) or 'idle'."""
        if dest == self.running:
            return "printing"
        if dest in self.waiting:
            return "waiting"
        return "sent" if dest in self.just_sent else "idle"

    def log(self, message: str, level: str = "info") -> None:
        getattr(log, {"error": "error", "warn": "warning"}.get(level, "info"))(message)
        self.log_lines = (self.log_lines + [(message, level)])[-300:]
        self.log_line.emit(message, level)

    def run_in_thread(self, work: Callable, done: Callable) -> None:
        """Run ``work()`` on a worker thread, then ``done(result)`` on the UI thread."""
        def run():
            try:
                result = work()
            except printing.PrintError as exc:
                log.warning("%s", exc)
                result = None
            except Exception:
                log.exception("Background task failed")
                result = None
            self.bridge.call.emit(lambda: done(result))
        threading.Thread(target=run, daemon=True).start()

    # ── Files ────────────────────────────────────────────────────────────────

    def add_paths(self, paths) -> int:
        """Add PDFs (folders: the PDFs in them); returns how many were added."""
        paths = list(paths)
        pdfs = expand_paths(paths)
        if paths and not pdfs:
            self.toast.emit(t("toast_not_pdf"), "info")
        known = {os.path.normcase(e.path) for e in self.entries}
        added = 0
        for path in pdfs:
            if os.path.normcase(path) in known:
                continue
            known.add(os.path.normcase(path))
            entry = Entry(path)
            self.entries.append(entry)
            self._queue_read(entry)
            added += 1
        if added:
            self.entries_changed.emit()
            self._changed()
        return added

    def _queue_read(self, entry: Entry) -> None:
        self._pending_reads += 1
        self._reads.put(entry)

    def _read_worker(self) -> None:
        while True:
            entry = self._reads.get()
            try:
                doc = read_document(entry.path)
            except Exception as exc:   # read_document should not raise; be safe
                from ..paper import DocInfo
                doc = DocInfo(entry.path, error=str(exc))
            self.bridge.doc_read.emit(entry, doc)

    def _on_doc_read(self, entry: Entry, doc) -> None:
        self._pending_reads = max(0, self._pending_reads - 1)
        if entry not in self.entries:
            self._changed()
            return
        entry.set_doc(doc)
        entry.stamp = file_stamp(entry.path)
        self._forget_previews(entry.path)
        if not doc.ok:
            self.log(t("read_error", name=doc.name, error=doc.error), "error")
        self._changed()

    def remove(self, entry: Entry) -> None:
        if self.busy or entry not in self.entries:
            return
        self.entries.remove(entry)
        self._forget_previews(entry.path)
        self.entries_changed.emit()
        self._changed()

    def clear(self) -> None:
        if self.busy or not self.entries:
            return
        self.entries.clear()
        self.previews.clear()
        self.entries_changed.emit()
        self._changed()

    def toggle_page(self, entry: Entry, pos: int) -> None:
        if self.busy or not entry.ready:
            return
        state = entry.states[pos]
        state.skip = not state.skip
        self._changed()

    def _refresh_changed_files(self) -> list[str]:
        """Read files again that changed on disk since they were added."""
        changed = []
        for entry in self.entries:
            if entry.doc is not None and entry.stamp is not None and file_stamp(entry.path) != entry.stamp:
                entry.doc = None
                self._queue_read(entry)
                changed.append(os.path.basename(entry.path))
        if changed:
            self.entries_changed.emit()
            self._changed()
        return changed

    # ── Page pictures ────────────────────────────────────────────────────────

    def preview(self, path: str, page_index: int, px: int, urgent: bool = False) -> Optional[QPixmap]:
        """The page's picture, ``px`` pixels on its long side, if it is ready; otherwise
        it is made in the background (``urgent``: before the others waiting)."""
        key = (path, page_index, px)
        if key in self.previews:
            return self.previews[key]
        if key not in self._preview_wanted:
            self._preview_wanted.add(key)
            self._preview_count += 1
            self._preview_requests.put((0 if urgent else 1, self._preview_count, key))
        return None

    def any_preview(self, path: str, page_index: int) -> Optional[QPixmap]:
        """The largest picture of the page made so far, to show while a bigger one is made."""
        sizes = [key for key in self.previews if key[:2] == (path, page_index)]
        return self.previews[max(sizes, key=lambda key: key[2])] if sizes else None

    def _preview_worker(self) -> None:
        while True:
            _urgency, _order, (path, index, px) = self._preview_requests.get()
            try:
                width, height, data = render_preview(path, index, px)
                image = QImage(data, width, height, width * 4, QImage.Format.Format_RGBA8888).copy()
            except Exception:
                log.debug("No preview for %s page %d", path, index, exc_info=True)
                image = QImage()
            self.bridge.preview.emit(path, index, px, image)

    def _on_preview(self, path: str, index: int, px: int, image: QImage) -> None:
        key = (path, index, px)
        self._preview_wanted.discard(key)
        if image.isNull() or not any(entry.path == path for entry in self.entries):
            return
        self.previews[key] = QPixmap.fromImage(image)
        if px >= BIG_PREVIEW_PX:
            big = [k for k in self.previews if k[2] >= BIG_PREVIEW_PX]
            for old in big[:-BIG_PREVIEWS_KEPT]:
                del self.previews[old]
        self.preview_ready.emit(path, index, px)

    def _forget_previews(self, path: str) -> None:
        for key in [k for k in self.previews if k[0] == path]:
            del self.previews[key]

    # ── Options ──────────────────────────────────────────────────────────────

    def set_a3_to_plotter(self, on: bool) -> None:
        self.opts.a3_to_plotter = on
        self.settings.a3_to_plotter = on
        self._save()
        self._changed()

    def set_roll(self, width: int) -> None:
        if width in ROLL_WIDTHS_MM and width != self.opts.roll:
            self.opts.roll = width
            self.settings.roll = width
            self._save()
            self._changed()

    def set_copies(self, copies: int) -> None:
        self.copies = max(1, min(MAX_COPIES, copies))
        self._changed()

    def set_duplex(self, dest: str, on: bool) -> None:
        self.duplex[dest] = on
        self._changed()

    def reset_done(self, dest: str, roll: Optional[int] = None) -> None:
        if not self.busy:
            reset_done(self.entries, self.opts, dest, roll)
            self._changed()

    # ── Printers ─────────────────────────────────────────────────────────────

    def load_printers(self, detect_if_unset: bool = True) -> None:
        def work():
            names = self.services.list_printers()
            found = None
            if detect_if_unset and not self.settings.has_printers:
                found = self.services.detect_printers(names)
            return names, found

        def done(result):
            if result is None:
                return
            names, found = result
            self.installed = names
            self.printers_changed.emit()
            if found:
                self._use_detected(found)
            for dest in (A4_DEST, A3_DEST):
                self._load_printer_info(dest)
            self._changed()
        self.run_in_thread(work, done)

    def detect_printers(self) -> None:
        def work():
            names = self.services.list_printers()
            return names, self.services.detect_printers(names)

        def done(result):
            if result is None:
                self.toast.emit(t("toast_detect_none"), "info")
                return
            self.installed = result[0]
            self.printers_changed.emit()
            self._use_detected(result[1])
        self.run_in_thread(work, done)

    def _use_detected(self, found: dict) -> None:
        names = [found.get(k) for k in ("a4", "a3", "plotter") if found.get(k)]
        if not names:
            self.toast.emit(t("toast_detect_none"), "info")
            return
        for dest, key in ((A4_DEST, "a4"), (A3_DEST, "a3"), (PLOT_DEST, "plotter")):
            if found.get(key):
                self.set_printer(dest, found[key])
        self.toast.emit(t("toast_detected", names=join_and(dict.fromkeys(names))), "ok")

    def set_printer(self, dest: str, name: str) -> None:
        attr = {A4_DEST: "a4", A3_DEST: "a3", PLOT_DEST: "plotter"}[dest]
        if getattr(self.printers, attr) == name:
            return
        setattr(self.printers, attr, name)
        setattr(self.settings, f"printer_{attr}", name)
        self._save()
        if dest != PLOT_DEST:
            self._load_printer_info(dest)
        self._changed()

    def _load_printer_info(self, dest: str) -> None:
        name = self.printer_for(dest)
        if not name:
            self.duplex_supported[dest] = False
            return

        def work():
            return self.services.supports_duplex(name), self.services.default_duplex(name)

        def done(result):
            if result is None or self.printer_for(dest) != name:
                return
            self.duplex_supported[dest], self.duplex[dest] = result
            self._changed()
        self.run_in_thread(work, done)

    def open_preferences(self, dest: str) -> None:
        name = self.printer_for(dest)
        if name:
            self.services.open_preferences(name)

    def _save(self) -> None:
        try:
            self.settings.save()
        except OSError:
            log.warning("Could not save the settings", exc_info=True)

    # ── Printing ─────────────────────────────────────────────────────────────

    def start_print(self, dest: str) -> bool:
        """Print ``dest``'s pages now, or after the print that is running."""
        if self.reading or self.status(dest) in ("printing", "waiting"):
            return False
        if not self.busy:
            changed = self._refresh_changed_files()
            if changed:
                self.toast.emit(t("toast_changed", name=changed[0]), "info")
                return False
        if not self.pages_for(self.plan(), dest):
            return False
        if not self.printer_for(dest):
            self.toast.emit(t("toast_pick_printer", target=t(f"target_{dest}")), "error")
            return False
        self.just_sent.discard(dest)
        if self.running is None:
            self._run(dest)
        else:
            self.waiting.append(dest)
            self._changed()
        return True

    def _run(self, dest: str) -> None:
        duplex = self.duplex[dest] if self.duplex_supported.get(dest) else None
        job = jobs.Job.from_plan(
            self.plan(), dest, jobs.Printers(self.printers.a4, self.printers.a3, self.printers.plotter),
            self.copies, duplex)
        self.running = dest
        self.totals[dest] = job.pages(dest) * self.copies
        self.handled[dest] = 0
        self._cancel = threading.Event()
        ctx = jobs.JobContext(
            self.services.backend, self.services.plot_backend, self.temp_dir,
            lambda message, level="info": self.bridge.log.emit(message, level),
            cancel=self._cancel, printer_exists=self.services.printer_exists,
            progress=lambda dest, handled, total: self.bridge.progress.emit(dest, handled, total),
            sent=self.bridge.part_sent.emit)
        entries = self.entries

        def run():
            sent, cancelled = 0, False
            try:
                sent = jobs.run(ctx, entries, job)
            except jobs.Cancelled:
                cancelled = True
                ctx.log(t("cancelled"), "warn")
            except Exception as exc:
                log.exception("Print job failed")
                ctx.log(t("unexpected", error=exc), "error")
                ctx.problems += 1
            self.bridge.job_done.emit(sent, cancelled, ctx.problems)

        self._job_thread = threading.Thread(target=run, name="print", daemon=True)
        self._job_thread.start()
        self._changed()
        return True

    def stop_print(self, dest: Optional[str] = None) -> None:
        """Stop ``dest``'s print or take it out of the queue; without one, stop everything."""
        if dest is None:
            self.waiting.clear()
        if dest in self.waiting:
            self.waiting.remove(dest)
        elif self.running is not None and dest in (None, self.running):
            self._cancel.set()
        self._changed()

    def _on_progress(self, dest: str, handled: int, total: int) -> None:
        self.handled[dest] = handled * self.copies
        self.changed.emit()

    def _on_part_sent(self, part: Part) -> None:
        mark_done(self.entries, [part])
        self._changed()

    def _on_job_done(self, sent: int, cancelled: bool, problems: int) -> None:
        dest, self.running = self.running, None
        self._job_thread = None
        plan = self.plan()
        if cancelled:
            self.toast.emit(t("toast_stopped"), "info")
        elif problems:
            self.toast.emit(t("toast_problems"), "error")
        else:
            self.just_sent.add(dest)
            QTimer.singleShot(DONE_SHOW_MS, self, lambda sent_dest=dest: self._sent_shown(sent_dest))
            pages = count(sent * self.copies, "page")
            nxt = plan.next_roll(self.opts.roll) if dest == PLOT_DEST and not plan.plot_included else None
            self.toast.emit(t("toast_sent_next", pages=pages, w=nxt) if nxt else t("toast_sent", pages=pages), "ok")
        while self.waiting:
            queued = self.waiting.pop(0)
            if self.pages_for(plan, queued) and self.printer_for(queued):
                self._run(queued)
                return
        if self.settings.clear_when_done and not plan.pending and not problems and not cancelled:
            self.entries.clear()
            self.previews.clear()
            self.entries_changed.emit()
        self._changed()

    def _sent_shown(self, dest: str) -> None:
        self.just_sent.discard(dest)
        self._changed()

    def wait_for_job(self, timeout: float = 5.0) -> None:
        thread = self._job_thread
        if thread is not None:
            thread.join(timeout)
