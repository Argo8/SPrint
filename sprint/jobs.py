"""Print workflows: send the right pages of the selected PDFs to each printer.

Runs on a worker thread. Progress goes through ``log(message, level)``;
levels are "info", "ok", "warn" and "error".
"""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from typing import Callable

from . import printing
from .i18n import count, t
from .paper import PLOT_A3, plot_groups
from .pdfdocs import file_for_pages, read_document

LogFn = Callable[..., None]   # log(message, level="info")


class Cancelled(Exception):
    """The user asked to stop printing."""


@dataclass
class JobContext:
    backend: printing.Backend
    temp_dir: str
    log: LogFn
    cancel: threading.Event = field(default_factory=threading.Event)
    paper_cls: type = printing.PlotterPaper   # replaced in tests
    printer_exists: Callable[[str], bool] = printing.printer_exists

    def check_cancel(self) -> None:
        if self.cancel.is_set():
            raise Cancelled


def roll_label(key: str) -> str:
    return t("roll_a3") if key == PLOT_A3.key else t("roll_mm", w=key)


def _check_printer(ctx: JobContext, printer: str, target: str) -> bool:
    if not printer:
        ctx.log(t("no_printer", target=target), "warn")
        return False
    if not ctx.printer_exists(printer):
        ctx.log(t("printer_missing", printer=printer), "error")
        return False
    return True


def _read_all(ctx: JobContext, paths: list[str]):
    """Read every file again: it may have changed since it was added."""
    docs = []
    for path in paths:
        ctx.check_cancel()
        doc = read_document(path)
        if doc.ok:
            docs.append(doc)
        else:
            ctx.log(t("read_error", name=doc.name, error=doc.error), "error")
    return docs


def _send(ctx: JobContext, printer: str, doc, path: str, copies: int) -> bool:
    """Print one file; logs problems and returns whether it was sent."""
    ctx.check_cancel()
    try:
        confirmed = ctx.backend.print_file(printer, path, copies)
    except Exception as exc:
        ctx.log(t("send_error", name=doc.name, error=exc), "error")
        return False
    if not confirmed:
        ctx.log(t("unconfirmed"), "warn")
    return True


def print_size(ctx: JobContext, kind: str, printer: str, copies: int, paths: list[str]) -> int:
    """Print all pages of one size (A4 or A3); returns the number of pages sent."""
    if not paths:
        ctx.log(t("no_files"), "warn")
        return 0
    if not _check_printer(ctx, printer, kind):
        return 0

    ctx.log(t("print_start", kind=kind, printer=printer, copies=copies))
    sent = 0
    try:
        for doc in _read_all(ctx, paths):
            indices = doc.indices_of_kind(kind)
            if not indices:
                continue
            try:
                path = file_for_pages(doc, indices, ctx.temp_dir, kind)
            except Exception as exc:
                ctx.log(t("send_error", name=doc.name, error=exc), "error")
                continue
            if _send(ctx, printer, doc, path, copies):
                ctx.log(t("print_sent", name=doc.name, pages=count(len(indices), "page")))
                sent += len(indices)
    except Cancelled:
        ctx.log(t("cancelled"), "warn")
        return sent

    if sent:
        ctx.log(t("print_done", kind=kind), "ok")
    else:
        ctx.log(t("print_nothing", kind=kind), "warn")
    return sent


def plot(ctx: JobContext, key: str, printer: str, copies: int, paths: list[str]) -> int:
    """Plot the pages of one roll option, one paper length at a time.

    Copies are collated per paper length: every file of that length, then again.
    Returns the number of pages sent (each page counted once).
    """
    if not key:
        ctx.log(t("plot_no_roll"), "warn")
        return 0
    if not paths:
        ctx.log(t("no_files"), "warn")
        return 0
    if not _check_printer(ctx, printer, t("target_plotter")):
        return 0

    label = roll_label(key)
    try:
        groups = plot_groups(_read_all(ctx, paths), key)
    except Cancelled:
        ctx.log(t("cancelled"), "warn")
        return 0
    if not groups:
        ctx.log(t("plot_nothing", label=label), "warn")
        return 0

    ctx.log(t("plot_start", printer=printer, label=label,
              groups=count(len(groups), "group"), copies=copies))
    sent = 0
    paper = _open_paper(ctx, printer)
    try:
        for number, group in enumerate(groups, start=1):
            ctx.check_cancel()
            ctx.log(t("plot_group", i=number, total=len(groups), length=group.length_mm,
                      pages=count(group.page_count, "page")))
            if paper is not None:
                try:
                    w, h = paper.apply(group.paper_width_mm, group.length_mm)
                    ctx.log(t("paper_set", w=_mm(w), h=_mm(h)))
                except Exception as exc:
                    ctx.log(t("paper_fail", error=exc), "warn")

            files = []
            for doc, indices in group.parts:
                try:
                    path = file_for_pages(doc, indices, ctx.temp_dir,
                                          f"{group.paper_width_mm}x{group.length_mm}")
                except Exception as exc:
                    ctx.log(t("send_error", name=doc.name, error=exc), "error")
                    continue
                files.append((doc, indices, path))

            for copy_no in range(copies):
                for doc, indices, path in files:
                    if _send(ctx, printer, doc, path, 1) and copy_no == 0:
                        ctx.log(t("print_sent", name=doc.name, pages=count(len(indices), "page")))
                        sent += len(indices)
    except Cancelled:
        ctx.log(t("cancelled"), "warn")
        return sent
    finally:
        _close_paper(ctx, paper)

    if sent:
        ctx.log(t("plot_done"), "ok")
    return sent


def _open_paper(ctx: JobContext, printer: str):
    """Start overriding the plotter's paper size; None if that is not possible."""
    paper = ctx.paper_cls(printer)
    try:
        paper.__enter__()
    except Exception as exc:
        ctx.log(t("paper_fail", error=exc), "warn")
        return None
    return paper


def _close_paper(ctx: JobContext, paper) -> None:
    if paper is None:
        return
    paper.__exit__(None, None, None)
    if paper.restore_error is not None:
        ctx.log(t("paper_restore_fail", error=paper.restore_error), "warn")


def _mm(value: float) -> str:
    return str(int(value)) if float(value).is_integer() else f"{value:.1f}"
