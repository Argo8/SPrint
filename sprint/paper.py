"""Page-size classification and plotter roll logic.

Pure Python with no I/O, so every routing rule can be unit-tested.

Routing rules:
  * A4 pages (either orientation)  -> A4 printer
  * A3 pages (either orientation)  -> A3 printer, or the plotter via the "A3" option
  * every other size               -> plotter, on the narrowest roll the page fits on
"""

from __future__ import annotations

import os
from collections import Counter
from dataclasses import dataclass, field
from typing import Iterable, Optional

PT_TO_MM = 25.4 / 72

A4 = "A4"
A3 = "A3"
OTHER = "OTHER"

STANDARD_SIZES_MM = {A4: (210.0, 297.0), A3: (297.0, 420.0)}

#: How far (in mm) a page may deviate from A4/A3 and still count as that size.
#: Scanners and PDF printers rarely produce exact sizes.
SIZE_TOLERANCE_MM = 3.0

#: Roll widths available on the plotter, in mm.
ROLL_WIDTHS_MM = (297, 420, 610, 914)

#: A page this much wider than a roll still counts as fitting it (rounding noise).
_ROLL_SLACK_MM = 0.5


def classify(width_mm: float, height_mm: float,
             tolerance: float = SIZE_TOLERANCE_MM) -> str:
    """Return A4, A3 or OTHER, ignoring orientation."""
    short, long = sorted((width_mm, height_mm))
    for name, (std_short, std_long) in STANDARD_SIZES_MM.items():
        if abs(short - std_short) <= tolerance and abs(long - std_long) <= tolerance:
            return name
    return OTHER


def roll_for(width_mm: float, height_mm: float) -> int:
    """Narrowest roll that fits the page's short side; the widest roll if none does."""
    short = min(width_mm, height_mm)
    for roll in sorted(ROLL_WIDTHS_MM):
        if short <= roll + _ROLL_SLACK_MM:
            return roll
    return max(ROLL_WIDTHS_MM)


def exceeds_rolls(width_mm: float, height_mm: float) -> bool:
    """True when even the widest roll is narrower than the page."""
    return min(width_mm, height_mm) > max(ROLL_WIDTHS_MM) + _ROLL_SLACK_MM


@dataclass(frozen=True)
class PlotOption:
    """One choice in the plotter's paper-width selector."""
    key: str                                # "A3" or the roll width as text, e.g. "610"
    paper_width_mm: int
    fixed_length_mm: Optional[int] = None   # A3 sheets are always 420 mm long

    def length_for(self, width_mm: float, height_mm: float) -> int:
        """Paper length needed for a page: its long side, rounded to whole mm."""
        return self.fixed_length_mm or round(max(width_mm, height_mm))


PLOT_A3 = PlotOption(A3, 297, 420)
PLOT_OPTIONS = (PLOT_A3,) + tuple(PlotOption(str(w), w) for w in ROLL_WIDTHS_MM)
_OPTIONS_BY_KEY = {opt.key: opt for opt in PLOT_OPTIONS}


def plot_option(key: str) -> PlotOption:
    return _OPTIONS_BY_KEY[key]


def plot_option_for(width_mm: float, height_mm: float) -> Optional[PlotOption]:
    """The plotter option a page belongs to; None for A4 pages."""
    kind = classify(width_mm, height_mm)
    if kind == A3:
        return PLOT_A3
    if kind == OTHER:
        return _OPTIONS_BY_KEY[str(roll_for(width_mm, height_mm))]
    return None


# ── Documents ────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class PageInfo:
    index: int          # 0-based position in the document
    width_mm: float     # as displayed, i.e. with /Rotate applied
    height_mm: float

    @property
    def kind(self) -> str:
        return classify(self.width_mm, self.height_mm)


@dataclass
class DocInfo:
    path: str
    page_count: int = 0                     # pages in the file, readable or not
    pages: list[PageInfo] = field(default_factory=list)
    error: Optional[str] = None             # set when the file cannot be read at all
    encrypted: bool = False                 # the file needs a password
    page_errors: list[tuple[int, str]] = field(default_factory=list)  # (page no., reason)

    @property
    def name(self) -> str:
        return os.path.basename(self.path)

    @property
    def ok(self) -> bool:
        return self.error is None

    def indices_of_kind(self, kind: str) -> list[int]:
        return [p.index for p in self.pages if p.kind == kind]


@dataclass
class PlotGroup:
    """Pages printed on the same paper size (one paper-size change on the plotter)."""
    paper_width_mm: int
    length_mm: int
    parts: list[tuple[DocInfo, list[int]]]  # page indices per document, in list order

    @property
    def page_count(self) -> int:
        return sum(len(indices) for _, indices in self.parts)


def plot_groups(docs: Iterable[DocInfo], key: str) -> list[PlotGroup]:
    """Group the pages for a plotter option by paper length, longest first."""
    option = plot_option(key)
    docs = list(docs)
    by_length: dict[int, dict[int, list[int]]] = {}
    for doc_no, doc in enumerate(docs):
        for page in doc.pages:
            if plot_option_for(page.width_mm, page.height_mm) != option:
                continue
            length = option.length_for(page.width_mm, page.height_mm)
            by_length.setdefault(length, {}).setdefault(doc_no, []).append(page.index)
    return [
        PlotGroup(option.paper_width_mm, length,
                  [(docs[doc_no], indices) for doc_no, indices in sorted(parts.items())])
        for length, parts in sorted(by_length.items(), reverse=True)
    ]


@dataclass
class Summary:
    files: int = 0
    total_pages: int = 0
    kinds: Counter = field(default_factory=Counter)        # A4 / A3 / OTHER -> pages
    other_sizes: Counter = field(default_factory=Counter)  # (w, h) in whole mm -> pages
    plot_counts: Counter = field(default_factory=Counter)  # plot option key -> pages
    oversize: int = 0                                      # pages wider than every roll
    unreadable: list[DocInfo] = field(default_factory=list)
    page_errors: list[tuple[DocInfo, int, str]] = field(default_factory=list)


def summarize(docs: Iterable[DocInfo]) -> Summary:
    s = Summary()
    for doc in docs:
        s.files += 1
        if not doc.ok:
            s.unreadable.append(doc)
            continue
        s.page_errors.extend((doc, no, reason) for no, reason in doc.page_errors)
        for page in doc.pages:
            s.total_pages += 1
            kind = page.kind
            s.kinds[kind] += 1
            if kind == OTHER:
                s.other_sizes[(round(page.width_mm), round(page.height_mm))] += 1
                if exceeds_rolls(page.width_mm, page.height_mm):
                    s.oversize += 1
            option = plot_option_for(page.width_mm, page.height_mm)
            if option is not None:
                s.plot_counts[option.key] += 1
    return s
