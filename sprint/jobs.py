"""Print jobs: send the pages a plan includes to the printers.

Runs on a worker thread: A4 pages, then A3 pages, then the plotter pages for
the roll in the plotter, longest paper first. The context's callbacks report
what happens; ``log(message, level)`` levels are "info", "ok", "warn" and "error".
"""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from typing import Callable, Optional

from . import printing
from .i18n import count, t
from .pdfdocs import file_for_pages
from .plan import A3_DEST, A4_DEST, PLOT_DEST, Entry, Part, Plan, PlotRun

LogFn = Callable[..., None]   # log(message, level="info")


class Cancelled(Exception):
    """The user asked to stop printing."""


@dataclass
class Printers:
    a4: str = ""
    a3: str = ""
    plotter: str = ""


@dataclass
class Job:
    """What one press of a print button sends: the pages of one destination."""
    a4: list[Part]
    a3: list[Part]
    plot: list[PlotRun]
    printers: Printers
    copies: int = 1
    duplex_a4: Optional[bool] = None   # None: as the printer is set up
    duplex_a3: Optional[bool] = None

    @classmethod
    def from_plan(cls, plan: Plan, dest: str, printers: Printers, copies: int = 1,
                  duplex: Optional[bool] = None) -> "Job":
        return cls(plan.a4_parts if dest == A4_DEST else [],
                   plan.a3_parts if dest == A3_DEST else [],
                   plan.plot_runs if dest == PLOT_DEST else [],
                   printers, copies,
                   duplex if dest == A4_DEST else None,
                   duplex if dest == A3_DEST else None)

    def pages(self, dest: str) -> int:
        if dest == PLOT_DEST:
            return sum(run.page_count for run in self.plot)
        return sum(len(part.pages) for part in (self.a4 if dest == A4_DEST else self.a3))


def _ignore(*_args) -> None:
    pass


@dataclass
class JobContext:
    backend: printing.Backend         # A4 / A3 printers
    plot_backend: printing.Backend    # plotter: must be able to set any paper size
    temp_dir: str
    log: LogFn
    cancel: threading.Event = field(default_factory=threading.Event)
    printer_exists: Callable[[str], bool] = printing.printer_exists
    progress: Callable[[str, int, int], None] = _ignore   # (destination, pages handled, pages)
    sent: Callable[[Part], None] = _ignore                # these pages reached the printer
    problems: int = 0                                     # parts that could not be printed

    def check_cancel(self) -> None:
        if self.cancel.is_set():
            raise Cancelled

    def error(self, message: str) -> None:
        self.problems += 1
        self.log(message, "error")


_TARGETS = {A4_DEST: "target_a4", A3_DEST: "target_a3", PLOT_DEST: "target_plot"}


def run(ctx: JobContext, entries: list[Entry], job: Job) -> int:
    """Print the job; returns the number of pages sent (copies not counted)."""
    return (_print(ctx, ctx.backend, A4_DEST, job.printers.a4, entries, job.a4, job.copies,
                   "A4", job.duplex_a4)
            + _print(ctx, ctx.backend, A3_DEST, job.printers.a3, entries, job.a3, job.copies,
                     "A3", job.duplex_a3)
            + _plot(ctx, job.printers.plotter, entries, job.plot, job.copies))


def _printer_ready(ctx: JobContext, printer: str, dest: str) -> bool:
    target = t(_TARGETS[dest])
    if not printer:
        ctx.error(t("no_printer", target=target))
        return False
    if not ctx.printer_exists(printer):
        ctx.error(t("printer_missing", printer=printer))
        return False
    ctx.log(t("job_target", target=target, printer=printer))
    return True


def _print(ctx, backend, dest, printer, entries, parts, copies, paper, duplex) -> int:
    if not parts:
        return 0
    total = sum(len(part.pages) for part in parts)
    if not _printer_ready(ctx, printer, dest):
        return 0
    handled = sent = 0
    ctx.progress(dest, 0, total)
    for part in parts:
        if _send(ctx, backend, printer, entries[part.entry], part, copies, paper, paper, duplex):
            sent += len(part.pages)
        handled += len(part.pages)
        ctx.progress(dest, handled, total)
    return sent


def _plot(ctx, printer, entries, runs, copies) -> int:
    if not runs:
        return 0
    total = sum(run.page_count for run in runs)
    if not _printer_ready(ctx, printer, PLOT_DEST):
        return 0
    handled = sent = 0
    ctx.progress(PLOT_DEST, 0, total)
    for run in runs:
        ctx.log(t("plot_run", w=run.width_mm, length=run.length_mm,
                  pages=count(run.page_count, "page")))
        paper = (run.width_mm, run.length_mm)
        for part in run.parts:
            label = f"{run.width_mm}x{run.length_mm}"
            if _send(ctx, ctx.plot_backend, printer, entries[part.entry], part, copies, paper,
                     label, None):
                sent += len(part.pages)
            handled += len(part.pages)
            ctx.progress(PLOT_DEST, handled, total)
    return sent


def _send(ctx, backend, printer, entry: Entry, part: Part, copies, paper, label, duplex) -> bool:
    """Print some pages of one file and log the outcome; returns whether they were sent."""
    ctx.check_cancel()
    doc = entry.doc
    indices = [doc.pages[pos].index for pos in part.pages]
    try:
        path = file_for_pages(doc, indices, ctx.temp_dir, label)
    except Exception as exc:
        ctx.error(t("temp_error", name=doc.name, error=exc))
        return False
    ctx.check_cancel()
    try:
        confirmed = backend.print_file(printer, path, copies, paper, duplex)
    except Exception as exc:
        ctx.error(t("send_error", name=doc.name, error=exc))
        return False
    ctx.log(t("print_sent", name=doc.name, pages=count(len(indices), "page")), "ok")
    if not confirmed:
        ctx.log(t("unconfirmed"), "warn")
    ctx.sent(part)
    return True
