"""What goes where: the plan for the files in the list.

Built again after every change. The window shows it, and each print button
sends exactly the pages the plan includes for its destination:

  * A4 and A3 pages that are not printed yet;
  * plotter pages that need the roll which is in the plotter now.

Pages the user skipped are left out; pages printed in this session are marked
done, so pressing a button again prints only what is left (for the plotter:
after the roll has been changed). Pure Python, unit-tested.
"""

from __future__ import annotations

import math
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from .paper import A3, A4, DocInfo, PageInfo, exceeds_rolls, plot_paper

A4_DEST, A3_DEST, PLOT_DEST = "a4", "a3", "plot"


@dataclass
class PageState:
    skip: bool = False   # left out by the user
    done: bool = False   # printed in this session


@dataclass
class Entry:
    """A file in the list. ``doc`` is None until the file has been read."""
    path: str
    doc: Optional[DocInfo] = None
    states: list[PageState] = field(default_factory=list)   # one per doc.pages
    stamp: Optional[tuple] = None                           # file (mtime, size) when read

    @property
    def ready(self) -> bool:
        return self.doc is not None and self.doc.ok

    def set_doc(self, doc: DocInfo) -> None:
        self.doc = doc
        self.states = [PageState() for _ in doc.pages]


@dataclass
class Options:
    a3_to_plotter: bool = False   # A3 pages on the plotter's 297 mm roll instead of the A3 printer
    roll: int = 610               # the roll in the plotter now
    duplex_a4: bool = False       # only changes the number of sheets shown


@dataclass
class PageRoute:
    dest: str                    # A4_DEST, A3_DEST or PLOT_DEST
    roll: Optional[int] = None   # plotter pages: the roll they need
    length_mm: int = 0           # plotter pages: paper length
    oversize: bool = False       # wider than every roll: printed smaller
    included: bool = False       # sent by its print button


@dataclass
class Part:
    """Pages of one file printed in one job."""
    entry: int          # position in the list
    pages: list[int]    # positions in entry.doc.pages


@dataclass
class PlotRun:
    """Plotter pages printed on one paper size: the roll width by one length."""
    width_mm: int
    length_mm: int
    parts: list[Part]

    @property
    def page_count(self) -> int:
        return sum(len(part.pages) for part in self.parts)


@dataclass
class DestCount:
    total: int = 0      # pages going there (not skipped)
    pending: int = 0    # of those, not printed yet: what its print button sends


@dataclass
class RollCount:
    total: int = 0
    pending: int = 0
    pending_mm: int = 0                                   # paper the pending pages take
    lengths: Counter = field(default_factory=Counter)     # pending: length -> pages


@dataclass
class EntryPlan:
    routes: list[PageRoute]   # one per doc.pages
    kinds: Counter            # A4 / A3 / OTHER -> pages, skipped ones included
    active: int = 0           # pages not skipped
    done: int = 0             # of those, printed


@dataclass
class Plan:
    entries: list[Optional[EntryPlan]] = field(default_factory=list)  # None: not read (yet)
    a4: DestCount = field(default_factory=DestCount)
    a3: DestCount = field(default_factory=DestCount)      # pages for the A3 printer
    a3_pages: int = 0                                     # A3 pages wherever they go
    rolls: dict[int, RollCount] = field(default_factory=dict)
    oversize: int = 0
    a4_sheets: int = 0
    a4_parts: list[Part] = field(default_factory=list)
    a3_parts: list[Part] = field(default_factory=list)
    plot_runs: list[PlotRun] = field(default_factory=list)   # longest paper first

    @property
    def plot_included(self) -> int:
        return sum(run.page_count for run in self.plot_runs)

    @property
    def plot_included_mm(self) -> int:
        return sum(run.length_mm * run.page_count for run in self.plot_runs)

    @property
    def plot_total(self) -> int:
        return sum(roll.total for roll in self.rolls.values())

    @property
    def plot_pending(self) -> int:
        return sum(roll.pending for roll in self.rolls.values())

    @property
    def pending(self) -> int:
        return self.a4.pending + self.a3.pending + self.plot_pending

    def pending_rolls(self) -> list[int]:
        return sorted(width for width, roll in self.rolls.items() if roll.pending)

    def next_roll(self, current: int) -> Optional[int]:
        """The roll to load after ``current``: the narrowest one still needed."""
        return next((width for width in self.pending_rolls() if width != current), None)


def route_of(page: PageInfo, opts: Options) -> PageRoute:
    kind = page.kind
    if kind == A4:
        return PageRoute(A4_DEST)
    if kind == A3 and not opts.a3_to_plotter:
        return PageRoute(A3_DEST)
    roll, length = plot_paper(page.width_mm, page.height_mm, kind)
    return PageRoute(PLOT_DEST, roll, length,
                     oversize=kind != A3 and exceeds_rolls(page.width_mm, page.height_mm))


def build_plan(entries: list[Entry], opts: Options) -> Plan:
    plan = Plan()
    runs: dict[int, dict[int, list[int]]] = {}   # length -> entry -> page positions
    for entry_no, entry in enumerate(entries):
        if not entry.ready:
            plan.entries.append(None)
            continue
        info = EntryPlan([], Counter())
        a4_here: list[int] = []
        a3_here: list[int] = []
        a4_pending = 0
        for pos, (page, state) in enumerate(zip(entry.doc.pages, entry.states)):
            route = route_of(page, opts)
            info.routes.append(route)
            info.kinds[page.kind] += 1
            if state.skip:
                continue
            info.active += 1
            info.done += state.done
            if page.kind == A3:
                plan.a3_pages += 1
            if route.dest == A4_DEST:
                plan.a4.total += 1
                if not state.done:
                    plan.a4.pending += 1
                    a4_pending += 1
                    route.included = True
                    a4_here.append(pos)
            elif route.dest == A3_DEST:
                plan.a3.total += 1
                if not state.done:
                    plan.a3.pending += 1
                    route.included = True
                    a3_here.append(pos)
            else:
                plan.oversize += route.oversize
                roll = plan.rolls.setdefault(route.roll, RollCount())
                roll.total += 1
                if not state.done:
                    roll.pending += 1
                    roll.pending_mm += route.length_mm
                    roll.lengths[route.length_mm] += 1
                    if route.roll == opts.roll:
                        route.included = True
                        runs.setdefault(route.length_mm, {}).setdefault(entry_no, []).append(pos)
        plan.a4_sheets += math.ceil(a4_pending / 2) if opts.duplex_a4 else a4_pending
        if a4_here:
            plan.a4_parts.append(Part(entry_no, a4_here))
        if a3_here:
            plan.a3_parts.append(Part(entry_no, a3_here))
        plan.entries.append(info)
    plan.plot_runs = [
        PlotRun(opts.roll, length, [Part(e, pages) for e, pages in sorted(parts.items())])
        for length, parts in sorted(runs.items(), reverse=True)
    ]
    return plan


def mark_done(entries: list[Entry], parts: list[Part]) -> None:
    for part in parts:
        states = entries[part.entry].states
        for pos in part.pages:
            states[pos].done = True


def reset_done(entries: list[Entry], opts: Options, dest: str, roll: Optional[int] = None) -> None:
    """Make printed pages of one destination (or one roll) printable again."""
    for entry in entries:
        if not entry.ready:
            continue
        for page, state in zip(entry.doc.pages, entry.states):
            route = route_of(page, opts)
            if route.dest == dest and (roll is None or route.roll == roll):
                state.done = False
