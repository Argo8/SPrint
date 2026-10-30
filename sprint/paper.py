"""Page sizes and plotter rolls.

Pure Python with no I/O, so every routing rule can be unit-tested.

Routing rules:
  * A4 pages (either orientation)  -> A4 printer
  * A3 pages (either orientation)  -> A3 printer, or the plotter's 297 mm roll
  * every other size               -> plotter, on the narrowest roll the page fits on

The plotter holds one roll at a time, so its pages are printed roll by roll.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

PT_TO_MM = 25.4 / 72

A4 = "A4"
A3 = "A3"
OTHER = "OTHER"

STANDARD_SIZES_MM = {A4: (210.0, 297.0), A3: (297.0, 420.0)}

#: How far (in mm) a page may deviate from A4/A3 and still count as that size.
#: Scanners and PDF printers rarely produce exact sizes.
SIZE_TOLERANCE_MM = 3.0

#: Roll widths the plotter takes, in mm.
ROLL_WIDTHS_MM = (297, 420, 610, 914)

#: A3 pages on the plotter: always a whole A3 sheet off the 297 mm roll.
A3_ROLL_MM, A3_LENGTH_MM = 297, 420

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


def plot_paper(width_mm: float, height_mm: float, kind: str) -> tuple[int, int]:
    """Roll width and paper length (mm) a page takes on the plotter.

    The page's short side goes across the roll, so the paper is as long as its
    long side. A page wider than every roll is printed smaller to fit the widest
    one, and the paper is cut to its reduced length.
    """
    if kind == A3:
        return A3_ROLL_MM, A3_LENGTH_MM
    short, long = sorted((width_mm, height_mm))
    roll = roll_for(width_mm, height_mm)
    if short > roll + _ROLL_SLACK_MM:
        return roll, round(long * roll / short)
    return roll, round(long)


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
