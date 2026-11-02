"""Windows printing: printers, paper sizes and the ways a PDF reaches the printer.

By default SPrint renders the pages itself with PDFium straight into a Windows
print job ("direct"). The job carries its own settings, so any paper size the
driver accepts works - custom plot lengths included - and the printer's
settings are never changed. A4/A3 pages can instead go through an external
program (settings.ini, [printing] program = sumatra | adobe):
  * SumatraPDF - silent, waits until the job is spooled
  * Adobe Reader / Acrobat ("/t") - returns at once, so we watch the print queue
"""

from __future__ import annotations

import ctypes
import glob
import logging
import math
import os
import subprocess
import time
from typing import Optional, Union

try:
    import pywintypes
    import win32gui
    import win32print
except ImportError:  # not on Windows (tests, development)
    pywintypes = win32gui = win32print = None

from .paper import MAX_PAPER_MM, PT_TO_MM
from .pdfdocs import PDFIUM_LOCK, load_flat_page, read_document
from .settings import app_dir

log = logging.getLogger(__name__)

CREATE_NO_WINDOW = 0x08000000
PRINTER_ENUM_LOCAL = 0x2
PRINTER_ENUM_CONNECTIONS = 0x4  # printers shared from other computers / print servers

# DEVMODE fields and values (wingdi.h)
DM_ORIENTATION = 0x1
DM_PAPERSIZE = 0x2
DM_PAPERLENGTH = 0x4
DM_PAPERWIDTH = 0x8
DM_COPIES = 0x100
DM_DUPLEX = 0x1000
DM_COLLATE = 0x8000
DM_OUT_BUFFER = 2
DM_IN_BUFFER = 8
DMORIENT_PORTRAIT = 1
DMCOLLATE_TRUE = 1
DMDUP_SIMPLEX, DMDUP_VERTICAL = 1, 2    # one-sided; two-sided, turned on the long edge
DMPAPER_A3 = 8
DMPAPER_USER = 256

# DeviceCapabilities / GetDeviceCaps indexes (wingdi.h)
DC_PAPERS, DC_PAPERSIZE, DC_DUPLEX, DC_COPIES, DC_COLLATE = 2, 3, 7, 18, 22
HORZRES, VERTRES = 8, 10
LOGPIXELSX, LOGPIXELSY = 88, 90
PHYSICALWIDTH, PHYSICALHEIGHT, PHYSICALOFFSETX, PHYSICALOFFSETY = 110, 111, 112, 113

# FPDF_RenderPage flags: annotations, print quality, text as outlines (no GDI font lookups)
FPDF_PRINTING = 0x800
RENDER_FLAGS = 0x01 | FPDF_PRINTING | 0x04

#: Finding a page's empty margins: a grey darker than this is ink.
INK_BELOW = 250
_INK = bytes(1 if value < INK_BELOW else 0 for value in range(256))

#: Paper sizes drivers know by name (portrait, mm) -> Windows paper ID. A4/A3
#: printers get these; a plot asks for its exact size first and falls back to
#: one of these only when the driver has no custom sizes.
STANDARD_PAPER = {(210, 297): 9, (297, 420): 8, (420, 594): 66}   # A4, A3, A2
NAMED_PAPER_MM = {"A4": (210, 297), "A3": (297, 420)}

SUMATRA_TIMEOUT_S = 15 * 60   # one job; large plots can take a while to spool
QUEUE_TIMEOUT_S = 45          # Adobe may need a while to start

Paper = Union[str, tuple]     # "A4", "A3" or (width_mm, length_mm)

#: Printers that make files instead of paper; never picked automatically.
VIRTUAL_PRINTER_WORDS = ("print to pdf", "xps document writer", "onenote", "fax", "pdf24",
                         "adobe pdf", "cutepdf", "pdfcreator", "bullzip", "foxit", "send to")
PLOTTER_WORDS = ("designjet", "plotter", "imageprograf", "surecolor", "plotwave",
                 "colorwave", "kip ", "t650", "t630", "t830", "t1700")
#: A printer that takes paper at least this wide (mm) is a plotter.
PLOTTER_MIN_WIDTH_MM = 594


class PrintError(Exception):
    pass


def _require_windows() -> None:
    if win32print is None:
        raise PrintError("printing requires Windows (pywin32)")


# ── Printers ─────────────────────────────────────────────────────────────────

def list_printers() -> list[str]:
    """Installed printers, including ones connected from a print server."""
    _require_windows()
    flags = PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS
    names = {info[2] for info in win32print.EnumPrinters(flags, None, 1)}
    return sorted(names, key=str.casefold)


def default_printer() -> Optional[str]:
    try:
        return win32print.GetDefaultPrinter() or None
    except Exception:
        return None


def printer_exists(name: str) -> bool:
    try:
        handle = win32print.OpenPrinter(name)
    except Exception:
        return False
    win32print.ClosePrinter(handle)
    return True


def _port(printer: str) -> str:
    handle = win32print.OpenPrinter(printer)
    try:
        return win32print.GetPrinter(handle, 2)["pPortName"]
    finally:
        win32print.ClosePrinter(handle)


def _capability(printer: str, capability: int, port: Optional[str] = None):
    """DeviceCapabilities, or None when the driver cannot tell."""
    try:
        return win32print.DeviceCapabilities(printer, port or _port(printer), capability)
    except Exception:
        log.debug("Cannot read capability %s of %s", capability, printer, exc_info=True)
        return None


def supports_duplex(printer: str) -> bool:
    return _capability(printer, DC_DUPLEX) == 1


def _user_settings(handle):
    """This user's printing preferences for the printer (what its settings dialog
    changes), or the printer's defaults when the user has not changed any."""
    try:
        devmode = win32print.GetPrinter(handle, 9)["pDevMode"]
    except Exception:
        devmode = None
    return devmode if devmode is not None else win32print.GetPrinter(handle, 2)["pDevMode"]


def _driver_sized(handle, printer: str, devmode):
    """``devmode`` in a buffer as large as the driver writes now: settings saved by
    an older version of the driver can be smaller, and the driver would write past
    their end (pywin32 does not check)."""
    try:
        needed = win32print.DocumentProperties(None, handle, printer, None, None, 0)
    except Exception:
        log.debug("Cannot ask %s for the size of its settings", printer, exc_info=True)
        return devmode
    if needed <= devmode.Size + devmode.DriverExtra:
        return devmode
    buffer = pywintypes.DEVMODEType(needed - devmode.Size)
    try:
        win32print.DocumentProperties(None, handle, printer, buffer, devmode, DM_IN_BUFFER | DM_OUT_BUFFER)
    except Exception:   # the driver does not take the old settings: start from its own
        log.warning("%s: saved printing preferences no longer fit its driver; using its defaults", printer)
        win32print.DocumentProperties(None, handle, printer, buffer, None, DM_OUT_BUFFER)
    return buffer


def default_duplex(printer: str) -> bool:
    """Whether the printer prints two-sided by default (its printing preferences)."""
    try:
        handle = win32print.OpenPrinter(printer)
        try:
            devmode = _user_settings(handle)
        finally:
            win32print.ClosePrinter(handle)
        return bool(devmode and devmode.Fields & DM_DUPLEX and devmode.Duplex > DMDUP_SIMPLEX)
    except Exception:
        return False


def _driver_papers(printer: str) -> tuple[list[int], list[tuple[float, float]]]:
    """The driver's paper IDs and their sizes in mm, in the same order."""
    port = None
    try:
        port = _port(printer)
    except Exception:
        pass
    ids = list(_capability(printer, DC_PAPERS, port) or ())
    sizes = []
    for size in _capability(printer, DC_PAPERSIZE, port) or ():
        x, y = (size["x"], size["y"]) if isinstance(size, dict) else tuple(size)[:2]
        sizes.append((x / 10, y / 10))   # tenths of a millimetre
    return ids, sizes


def _paper_capabilities(printer: str) -> tuple[set[int], float]:
    """Paper IDs the driver offers and the widest paper's short side in mm."""
    ids, sizes = _driver_papers(printer)
    return set(ids), max((min(size) for size in sizes), default=0.0)


def _is_paper_of_size(printer: str, paper_id: int, size_mm: tuple) -> bool:
    """Whether the driver's paper ``paper_id`` has this size (some drivers swap a
    custom size for a paper of theirs that has the same size)."""
    ids, sizes = _driver_papers(printer)
    if paper_id not in ids or len(sizes) != len(ids):
        return False
    have = sorted(sizes[ids.index(paper_id)])
    return all(abs(a - b) <= 1.0 for a, b in zip(have, sorted(size_mm)))


def is_virtual_printer(name: str) -> bool:
    lowered = name.casefold()
    return any(word in lowered for word in VIRTUAL_PRINTER_WORDS)


def detect_printers(printers: Optional[list[str]] = None) -> dict[str, str]:
    """Best guess for the A4 printer, the A3 printer and the plotter ("" when none fits).

    The plotter takes paper at least A1 wide (or is named like one), the A4
    printer is the Windows default printer (else one without A3, else any), and
    the A3 printer is another printer that offers A3 (else the A4 printer, if it can).
    """
    names = list_printers() if printers is None else list(printers)
    candidates = [n for n in names if not is_virtual_printer(n)] or names
    caps = {name: _paper_capabilities(name) for name in candidates}
    plotter = next((n for n in candidates
                    if any(w in n.casefold() for w in PLOTTER_WORDS)
                    or caps[n][1] >= PLOTTER_MIN_WIDTH_MM), "")
    others = [n for n in candidates if n != plotter]
    a4_only = [n for n in others if DMPAPER_A3 not in caps[n][0]]
    default = default_printer()
    a4 = default if default in others else (a4_only or others or [""])[0]
    a3 = next((n for n in others if DMPAPER_A3 in caps[n][0] and n != a4), "")
    if not a3 and a4 and DMPAPER_A3 in caps[a4][0]:
        a3 = a4
    return {"a4": a4, "a3": a3, "plotter": plotter}


def open_printer_preferences(name: str) -> None:
    """Show the driver's "Printing preferences" dialog without blocking the caller."""
    subprocess.Popen(["rundll32", "printui.dll,PrintUIEntry", "/e", "/n", name],
                     creationflags=CREATE_NO_WINDOW)


def _job_ids(printer: str) -> Optional[set[int]]:
    try:
        handle = win32print.OpenPrinter(printer)
        try:
            return {job["JobId"] for job in win32print.EnumJobs(handle, 0, 9999, 1)}
        finally:
            win32print.ClosePrinter(handle)
    except Exception:
        log.debug("Cannot read the queue of %s", printer, exc_info=True)
        return None


def _wait_for_new_job(printer: str, before: Optional[set[int]],
                      timeout: float = QUEUE_TIMEOUT_S) -> bool:
    """Wait until a job that was not in ``before`` shows up in the queue."""
    if before is None:
        return False
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        time.sleep(0.3)
        current = _job_ids(printer)
        if current is None:
            return False
        if current - before:
            return True
    return False


# ── Paper size and page placement ────────────────────────────────────────────

def job_settings(printer: str, paper: Optional[Paper] = None, copies: int = 1,
                 duplex: Optional[bool] = None):
    """The printer's settings for this user (its "printing preferences") with the
    paper size, copies and (unless None) one- or two-sided printing of one job.
    Nothing is saved to the printer.

    A plot (``paper`` as width and length) asks for exactly that size, and for a
    standard paper of that size if the driver has no custom sizes. Raises
    PrintError if the driver accepts neither.
    """
    _require_windows()
    if paper is None:
        papers = [None]
    elif isinstance(paper, str):
        papers = [(STANDARD_PAPER[NAMED_PAPER_MM[paper]], NAMED_PAPER_MM[paper])]
    else:
        width, length = paper
        if round(max(width, length) * 10) > MAX_PAPER_MM * 10:
            raise PrintError(f"{width:g}×{length:g} mm: Windows cannot send paper longer than {MAX_PAPER_MM} mm")
        papers = [(DMPAPER_USER, paper)]
        standard = STANDARD_PAPER.get((round(width), round(length)))
        if standard is not None:
            papers.append((standard, paper))
    handle = win32print.OpenPrinter(printer)
    try:
        for choice in papers:
            stored = _user_settings(handle)   # a fresh copy for every try
            if stored is None:
                raise PrintError("the printer driver returned no settings")
            devmode = _driver_sized(handle, printer, stored)
            if _ask_driver(handle, printer, devmode, choice, copies, duplex):
                return devmode
        width, length = papers[0][1]
        raise PrintError(f"the driver does not accept {width:g}×{length:g} mm paper")
    finally:
        win32print.ClosePrinter(handle)


def _ask_driver(handle, printer: str, devmode, paper: Optional[tuple], copies: int,
                duplex: Optional[bool]) -> bool:
    """Put the job's settings into ``devmode`` and let the driver check them; False
    if it does not take the paper (``paper``: Windows paper ID and size in mm)."""
    if paper is not None:
        paper_id, (width, length) = paper
        devmode.Orientation = DMORIENT_PORTRAIT
        devmode.PaperSize = paper_id
        devmode.Fields |= DM_ORIENTATION | DM_PAPERSIZE
        if paper_id == DMPAPER_USER:
            devmode.PaperWidth = round(width * 10)     # tenths of a millimetre
            devmode.PaperLength = round(length * 10)
            devmode.Fields |= DM_PAPERWIDTH | DM_PAPERLENGTH
        else:
            devmode.Fields &= ~(DM_PAPERWIDTH | DM_PAPERLENGTH)
    devmode.Copies = copies
    devmode.Collate = DMCOLLATE_TRUE
    devmode.Fields |= DM_COPIES | DM_COLLATE
    if duplex is not None:
        devmode.Duplex = DMDUP_VERTICAL if duplex else DMDUP_SIMPLEX
        devmode.Fields |= DM_DUPLEX
    # The driver may adjust the request to what it supports.
    win32print.DocumentProperties(None, handle, printer, devmode, devmode, DM_IN_BUFFER | DM_OUT_BUFFER)
    return paper is None or devmode.PaperSize == paper_id or _is_paper_of_size(printer, devmode.PaperSize,
                                                                            (width, length))


def place_page(size_mm, sheet_px, offset_px, printable_px, dpi,
               blank_mm=(0.0, 0.0, 0.0, 0.0)) -> tuple[int, int, int, int, int]:
    """Where to draw a page on the sheet: (x, y, width, height, rotate) in device pixels,
    relative to the printable area as GDI expects.

    The page is turned a quarter if its orientation differs from the sheet's and
    centred on the sheet. It prints at its real size when the edge the printer
    cannot reach holds only empty margin (``blank_mm``: left, top, right, bottom of
    the page as shown), so drawings keep their scale; otherwise it is shrunk just
    enough for its ink to fit. It is never enlarged.
    """
    w_mm, h_mm = size_mm
    left, top, right, bottom = blank_mm
    sheet_w, sheet_h = sheet_px
    across, down = sheet_w / dpi[0], sheet_h / dpi[1]   # inches: a printer's dpi may differ by axis
    rotate = 0
    if (w_mm > h_mm and across < down) or (w_mm < h_mm and across > down):
        w_mm, h_mm, rotate = h_mm, w_mm, 1          # 1 = 90° clockwise
        left, top, right, bottom = bottom, left, top, right   # the left edge goes to the top
    per_mm_x, per_mm_y = dpi[0] / 25.4, dpi[1] / 25.4
    width, height = w_mm * per_mm_x, h_mm * per_mm_y
    ink_w = max(1.0, width - (left + right) * per_mm_x)
    ink_h = max(1.0, height - (top + bottom) * per_mm_y)
    scale = min(1.0, printable_px[0] / ink_w, printable_px[1] / ink_h)
    x, width = _along(sheet_w, width * scale, offset_px[0], printable_px[0],
                      left * per_mm_x * scale, right * per_mm_x * scale)
    y, height = _along(sheet_h, height * scale, offset_px[1], printable_px[1],
                       top * per_mm_y * scale, bottom * per_mm_y * scale)
    return x, y, width, height, rotate


def _along(sheet: int, size: float, offset: int, printable: int, blank_before: float,
           blank_after: float) -> tuple[int, int]:
    """Start (relative to the printable area) and length of the page along one axis:
    centred on the sheet, moved as little as needed to keep its ink printable."""
    before, after = math.floor(blank_before), math.floor(blank_after)
    size = min(round(size), printable + before + after)   # rounding must not push ink out
    start = (sheet - size) // 2 - offset
    return max(-before, min(start, printable - size + after)), size


def blank_margins_mm(page) -> tuple[float, float, float, float]:
    """Empty margin (mm) between each edge of a pypdfium2 page and its nearest ink, as
    it prints: (left, top, right, bottom) of the page as shown. Zeros when the page
    has no ink or cannot be drawn, which means ink up to the edges."""
    import pypdfium2.raw as pdfium_c

    try:
        scale = min(2 * PT_TO_MM, 2400 / max(page.get_size()))   # 2 pixels a mm, at most 2400 a side
        bitmap = page.render(scale=scale, grayscale=True, may_draw_forms=False, extra_flags=FPDF_PRINTING,
                             force_bitmap_format=pdfium_c.FPDFBitmap_Gray)
        try:
            width, height, stride = bitmap.width, bitmap.height, bitmap.stride
            data = bytes(bitmap.buffer)
        finally:
            bitmap.close()
    except Exception:
        log.debug("Cannot find the margins of a page", exc_info=True)
        return 0.0, 0.0, 0.0, 0.0
    rows = [data[row * stride:row * stride + width].translate(_INK) for row in range(height)]
    inked = [row for row, ink in enumerate(rows) if 1 in ink]
    if not inked:
        return 0.0, 0.0, 0.0, 0.0
    first = min(rows[row].find(1) for row in inked)
    last = max(rows[row].rfind(1) for row in inked)
    mm = PT_TO_MM / scale   # one pixel; a pixel less on each side, ink may start anywhere in it
    return (max(0, first - 1) * mm, max(0, inked[0] - 1) * mm,
            max(0, width - 2 - last) * mm, max(0, height - 2 - inked[-1]) * mm)


# ── PDF print programs ───────────────────────────────────────────────────────

class Backend:
    """Sends one PDF file to a printer."""
    key = ""

    def print_file(self, printer: str, path: str, copies: int = 1,
                   paper: Optional[Paper] = None, duplex: Optional[bool] = None) -> bool:
        """Print and return True once the job is known to be in the print queue.

        ``paper`` asks for that paper instead of the printer's default, and
        ``duplex`` for one- or two-sided printing, where the program supports it.
        False means the job was handed over but could not be confirmed. Raises on
        failure.
        """
        raise NotImplementedError


class DirectBackend(Backend):
    """Renders the pages with PDFium straight into a Windows print job (vector output)."""
    key = "direct"

    def print_file(self, printer, path, copies=1, paper=None, duplex=None):
        devmode = job_settings(printer, paper, copies, duplex)
        if isinstance(paper, tuple):
            devmode = _room_for_margins(printer, paper, copies, devmode)
        sizes = {p.index: (p.width_mm, p.height_mm) for p in read_document(path).pages}
        with PDFIUM_LOCK:
            return self._print(printer, path, copies, devmode, sizes)

    def _print(self, printer, path, copies, devmode, sizes):
        import pypdfium2 as pdfium
        import pypdfium2.raw as pdfium_c

        name = os.path.basename(path)
        pdf = pdfium.PdfDocument(path)
        blanks = {}   # page -> empty margins, worked out once for all copies
        jobs = 1
        if copies > 1 and not _driver_copies(printer, copies, collate=len(pdf) > 1):
            jobs, devmode.Copies = copies, 1    # print the job again for every copy
        try:
            for _ in range(jobs):
                hdc = win32gui.CreateDC("WINSPOOL", printer, devmode)
                try:
                    win32print.StartDoc(hdc, (name, None, None, 0))
                    try:
                        for index in range(len(pdf)):
                            page = load_flat_page(pdf, index, for_print=True)
                            try:
                                size = sizes.get(index) or tuple(v * PT_TO_MM for v in page.get_size())
                                if index not in blanks:
                                    blanks[index] = blank_margins_mm(page)
                                win32print.StartPage(hdc)
                                x, y, w, h, rotate = place_page(size, *_sheet_metrics(hdc), blanks[index])
                                pdfium_c.FPDF_RenderPage(ctypes.c_void_p(hdc), page.raw,
                                                         x, y, w, h, rotate, RENDER_FLAGS)
                                win32print.EndPage(hdc)
                            finally:
                                page.close()
                    except BaseException:
                        win32print.AbortDoc(hdc)
                        raise
                    win32print.EndDoc(hdc)
                finally:
                    win32gui.DeleteDC(hdc)
        finally:
            pdf.close()
        return True   # EndDoc returns once the job is spooled


def _driver_copies(printer: str, copies: int, collate: bool) -> bool:
    """Whether the driver itself can print this many (collated) copies."""
    try:
        port = _port(printer)
        if win32print.DeviceCapabilities(printer, port, DC_COPIES) < copies:
            return False
        return not collate or win32print.DeviceCapabilities(printer, port, DC_COLLATE) == 1
    except Exception:
        log.debug("Cannot read the capabilities of %s", printer, exc_info=True)
        return False


def _room_for_margins(printer: str, paper: tuple, copies: int, devmode):
    """Make custom (roll) paper longer by the printer's top and bottom margins,
    so a page as long as the paper still prints at its real size."""
    hdc = win32gui.CreateDC("WINSPOOL", printer, devmode)
    try:
        (_, sheet_h), _, (_, printable_h), (_, dpi_y) = _sheet_metrics(hdc)
    finally:
        win32gui.DeleteDC(hdc)
    margins_mm = (sheet_h - printable_h) / dpi_y * 25.4
    if margins_mm < 0.5:
        return devmode
    width, length = paper
    try:
        return job_settings(printer, (width, length + math.ceil(margins_mm)), copies)
    except PrintError:   # longer than the driver allows: keep the length, the page shrinks a little
        return devmode


def _sheet_metrics(hdc):
    caps = lambda index: win32print.GetDeviceCaps(hdc, index)  # noqa: E731
    return ((caps(PHYSICALWIDTH), caps(PHYSICALHEIGHT)),
            (caps(PHYSICALOFFSETX), caps(PHYSICALOFFSETY)),
            (caps(HORZRES), caps(VERTRES)),
            (caps(LOGPIXELSX), caps(LOGPIXELSY)))


class SumatraBackend(Backend):
    key = "sumatra"

    def __init__(self, exe: str):
        self.exe = exe

    def print_file(self, printer, path, copies=1, paper=None, duplex=None):
        settings = []
        if copies > 1:
            settings.append(f"{copies}x")
        if paper in NAMED_PAPER_MM:   # SumatraPDF 3.6 knows no custom sizes
            settings.append(f"paper={paper}")
        if duplex is not None:
            settings.append("duplexlong" if duplex else "simplex")
        args = [self.exe, "-print-to", printer, "-silent"]
        if settings:
            args += ["-print-settings", ",".join(settings)]
        args.append(path)
        try:
            result = subprocess.run(args, creationflags=CREATE_NO_WINDOW,
                                    stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                    stderr=subprocess.DEVNULL, timeout=SUMATRA_TIMEOUT_S)
        except subprocess.TimeoutExpired:
            raise PrintError("SumatraPDF did not finish in time") from None
        if result.returncode != 0:
            raise PrintError(f"SumatraPDF exit code {result.returncode}")
        return True  # SumatraPDF exits only after the job has been spooled


class _QueueWatchingBackend(Backend):
    """For programs that return before printing; one launch per copy."""

    def _launch(self, printer: str, path: str) -> None:
        raise NotImplementedError

    def print_file(self, printer, path, copies=1, paper=None, duplex=None):
        confirmed = True
        for _ in range(copies):
            before = _job_ids(printer)
            self._launch(printer, path)
            # Waiting also keeps the next launch (or paper change) from racing this one.
            confirmed = _wait_for_new_job(printer, before) and confirmed
        return confirmed


class AdobeBackend(_QueueWatchingBackend):
    key = "adobe"

    def __init__(self, exe: str):
        self.exe = exe

    def _launch(self, printer, path):
        subprocess.Popen([self.exe, "/h", "/t", path, printer], creationflags=CREATE_NO_WINDOW,
                         stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL)


def _registered_app(exe_name: str) -> Optional[str]:
    """Path registered under ...\\CurrentVersion\\App Paths\\<exe_name>."""
    try:
        import winreg
    except ImportError:
        return None
    key = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}"
    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(root, key, 0, winreg.KEY_READ | view) as k:
                    value, _ = winreg.QueryValueEx(k, "")
            except OSError:
                continue
            path = os.path.expandvars(str(value).strip().strip('"'))
            if os.path.isfile(path):
                return path
    return None


def _first_file(paths) -> Optional[str]:
    for path in paths:
        if path and os.path.isfile(path):
            return path
    return None


def find_sumatra() -> Optional[str]:
    env = os.environ.get
    return _first_file([
        os.path.join(app_dir(), "SumatraPDF.exe"),
        *sorted(glob.glob(os.path.join(app_dir(), "SumatraPDF*.exe"))),  # portable builds
        _registered_app("SumatraPDF.exe"),
        os.path.join(env("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(env("ProgramFiles", ""), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(env("ProgramFiles(x86)", ""), "SumatraPDF", "SumatraPDF.exe"),
    ])


def find_adobe() -> Optional[str]:
    env = os.environ.get
    pf, pf86 = env("ProgramFiles", r"C:\Program Files"), env("ProgramFiles(x86)", r"C:\Program Files (x86)")
    return _first_file([
        _registered_app("AcroRd32.exe"),
        _registered_app("Acrobat.exe"),
        os.path.join(pf86, r"Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"),
        os.path.join(pf, r"Adobe\Acrobat DC\Acrobat\Acrobat.exe"),
        os.path.join(pf, r"Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"),
        os.path.join(pf86, r"Adobe\Acrobat DC\Acrobat\Acrobat.exe"),
        os.path.join(pf86, r"Adobe\Reader 11.0\Reader\AcroRd32.exe"),
    ])


def find_backend(program: str = "") -> Backend:
    """The backend for A4/A3 printing: an external program if configured and found, else direct."""
    if program == "sumatra":
        exe = find_sumatra()
        if exe:
            return SumatraBackend(exe)
    elif program == "adobe":
        exe = find_adobe()
        if exe:
            return AdobeBackend(exe)
    return DirectBackend()
