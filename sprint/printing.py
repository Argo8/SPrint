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
    import win32gui
    import win32print
except ImportError:  # not on Windows (tests, development)
    win32gui = win32print = None

from .paper import PT_TO_MM
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
DM_COLLATE = 0x8000
DM_OUT_BUFFER = 2
DM_IN_BUFFER = 8
DMORIENT_PORTRAIT = 1
DMCOLLATE_TRUE = 1
DMPAPER_USER = 256

# DeviceCapabilities / GetDeviceCaps indexes (wingdi.h)
DC_COPIES, DC_COLLATE = 18, 22
HORZRES, VERTRES = 8, 10
LOGPIXELSX, LOGPIXELSY = 88, 90
PHYSICALWIDTH, PHYSICALHEIGHT, PHYSICALOFFSETX, PHYSICALOFFSETY = 110, 111, 112, 113

# FPDF_RenderPage flags: annotations, print quality, text as outlines (no GDI font lookups)
RENDER_FLAGS = 0x01 | 0x800 | 0x04

#: Paper sizes drivers know by name (portrait, mm) -> Windows paper ID.
#: Named sizes are understood by more drivers than a custom width/length.
STANDARD_PAPER = {(210, 297): 9, (297, 420): 8, (420, 594): 66}   # A4, A3, A2
NAMED_PAPER_MM = {"A4": (210, 297), "A3": (297, 420)}

SUMATRA_TIMEOUT_S = 15 * 60   # one job; large plots can take a while to spool
QUEUE_TIMEOUT_S = 45          # Adobe may need a while to start

Paper = Union[str, tuple]     # "A4", "A3" or (width_mm, length_mm)


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

def job_settings(printer: str, paper: Optional[Paper] = None, copies: int = 1):
    """The printer's settings for this user (its "printing preferences") with the
    paper size and copies of one job. Nothing is saved to the printer.

    Raises PrintError if the driver does not accept the paper size.
    """
    _require_windows()
    handle = win32print.OpenPrinter(printer)
    try:
        devmode = win32print.GetPrinter(handle, 2)["pDevMode"]
        if devmode is None:
            raise PrintError("the printer driver returned no settings")
        paper_id = None
        if paper is not None:
            width, length = NAMED_PAPER_MM.get(paper, paper)
            paper_id = STANDARD_PAPER.get((round(width), round(length)), DMPAPER_USER)
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
        # Let the driver validate the request; it may adjust it to what it supports.
        win32print.DocumentProperties(None, handle, printer, devmode, devmode,
                                      DM_IN_BUFFER | DM_OUT_BUFFER)
        if paper_id is not None and devmode.PaperSize != paper_id:
            raise PrintError(f"the driver does not accept {width:g}×{length:g} mm paper")
        return devmode
    finally:
        win32print.ClosePrinter(handle)


def place_page(size_mm, sheet_px, offset_px, printable_px, dpi) -> tuple[int, int, int, int, int]:
    """Where to draw a page on the sheet: (x, y, width, height, rotate) in device pixels,
    relative to the printable area as GDI expects.

    The page is turned a quarter if its orientation differs from the sheet's,
    printed at its real size unless that exceeds the printable area (then it is
    shrunk to fit), and centred on the sheet.
    """
    w_mm, h_mm = size_mm
    sheet_w, sheet_h = sheet_px
    rotate = 0
    if (w_mm > h_mm and sheet_w < sheet_h) or (w_mm < h_mm and sheet_w > sheet_h):
        w_mm, h_mm, rotate = h_mm, w_mm, 1          # 1 = 90° clockwise
    width, height = w_mm / 25.4 * dpi[0], h_mm / 25.4 * dpi[1]
    scale = min(1.0, printable_px[0] / width, printable_px[1] / height)
    width = min(round(width * scale), printable_px[0])
    height = min(round(height * scale), printable_px[1])
    x = (sheet_w - width) // 2 - offset_px[0]
    y = (sheet_h - height) // 2 - offset_px[1]
    x = max(0, min(x, printable_px[0] - width))
    y = max(0, min(y, printable_px[1] - height))
    return x, y, width, height, rotate


# ── PDF print programs ───────────────────────────────────────────────────────

class Backend:
    """Sends one PDF file to a printer."""
    key = ""

    def print_file(self, printer: str, path: str, copies: int = 1,
                   paper: Optional[Paper] = None) -> bool:
        """Print and return True once the job is known to be in the print queue.

        ``paper`` asks for that paper instead of the printer's default, where the
        program supports it. False means the job was handed over but could not be
        confirmed. Raises on failure.
        """
        raise NotImplementedError


class DirectBackend(Backend):
    """Renders the pages with PDFium straight into a Windows print job (vector output)."""
    key = "direct"

    def print_file(self, printer, path, copies=1, paper=None):
        import pypdfium2 as pdfium
        import pypdfium2.raw as pdfium_c

        from .pdfdocs import read_document

        devmode = job_settings(printer, paper, copies)
        if isinstance(paper, tuple) and devmode.PaperSize == DMPAPER_USER:
            devmode = _room_for_margins(printer, paper, copies, devmode)
        sizes = {p.index: (p.width_mm, p.height_mm) for p in read_document(path).pages}
        name = os.path.basename(path)
        pdf = pdfium.PdfDocument(path)
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
                            page = pdf[index]
                            try:
                                size = sizes.get(index) or tuple(v * PT_TO_MM for v in page.get_size())
                                win32print.StartPage(hdc)
                                x, y, w, h, rotate = place_page(size, *_sheet_metrics(hdc))
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
        handle = win32print.OpenPrinter(printer)
        try:
            port = win32print.GetPrinter(handle, 2)["pPortName"]
        finally:
            win32print.ClosePrinter(handle)
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

    def print_file(self, printer, path, copies=1, paper=None):
        settings = []
        if copies > 1:
            settings.append(f"{copies}x")
        if paper in NAMED_PAPER_MM:   # SumatraPDF 3.6 knows no custom sizes
            settings.append(f"paper={paper}")
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

    def print_file(self, printer, path, copies=1, paper=None):
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
