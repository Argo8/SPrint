"""Windows printing: printer list, plotter paper size and the program that prints PDFs.

PDFs are printed by an external program, in order of preference:
  * SumatraPDF  — silent, waits until the job is spooled, supports copies
  * Adobe Reader / Acrobat ("/t") — returns at once, so we watch the print queue
  * the default PDF app via the Windows "printto" verb (same caveat)
"""

from __future__ import annotations

import glob
import logging
import os
import subprocess
import time
from typing import Optional

try:
    import win32api
    import win32print
except ImportError:  # not on Windows (tests, development)
    win32api = win32print = None

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
DM_OUT_BUFFER = 2
DM_IN_BUFFER = 8
DMORIENT_PORTRAIT = 1
DMPAPER_USER = 256

#: Paper sizes drivers know by name (portrait, mm) -> Windows paper ID.
#: Named sizes are understood by more drivers than a custom width/length.
STANDARD_PAPER = {(210, 297): 9, (297, 420): 8, (420, 594): 66}   # A4, A3, A2

SUMATRA_TIMEOUT_S = 15 * 60   # one job; large plots can take a while to spool
QUEUE_TIMEOUT_S = 45          # Adobe may need a while to start


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


# ── Plotter paper size ───────────────────────────────────────────────────────

def _devmode(handle, level: int):
    try:
        return win32print.GetPrinter(handle, level)["pDevMode"]
    except Exception:
        return None


def _first_devmode(handle, levels):
    for level in levels:
        devmode = _devmode(handle, level)
        if devmode is not None:
            return devmode
    return None


class PlotterPaper:
    """Temporarily makes a custom paper size the printer's default for this user.

    Programs started with "print to" use the printer's default settings, so this
    is how every group of plot pages gets its paper length. The previous
    per-user settings are restored on exit; nothing changes for other users.

        with PlotterPaper(name) as paper:
            paper.apply(610, 841)
            ...print...
        if paper.restore_error: ...
    """

    def __init__(self, printer: str):
        self.printer = printer
        self.restore_error: Optional[Exception] = None
        self._handle = None
        self._original = None
        self._changed = False

    def __enter__(self) -> "PlotterPaper":
        _require_windows()
        self._handle = win32print.OpenPrinter(self.printer)
        self._original = _devmode(self._handle, 9)
        return self

    def apply(self, width_mm: float, length_mm: float) -> tuple[float, float]:
        """Set the paper size; returns the size the driver accepted, in mm."""
        # The user's current defaults: per-user (level 9), else global (8 / 2).
        devmode = _first_devmode(self._handle, (9, 8, 2))
        if devmode is None:
            raise PrintError("the printer driver returned no settings")
        paper_id = STANDARD_PAPER.get((round(width_mm), round(length_mm)), DMPAPER_USER)
        devmode.Orientation = DMORIENT_PORTRAIT
        devmode.PaperSize = paper_id
        devmode.Fields |= DM_ORIENTATION | DM_PAPERSIZE
        if paper_id == DMPAPER_USER:
            devmode.PaperWidth = round(width_mm * 10)     # tenths of a millimetre
            devmode.PaperLength = round(length_mm * 10)
            devmode.Fields |= DM_PAPERWIDTH | DM_PAPERLENGTH
        else:
            devmode.Fields &= ~(DM_PAPERWIDTH | DM_PAPERLENGTH)
        # Let the driver validate the request; it may adjust it to what it supports.
        win32print.DocumentProperties(None, self._handle, self.printer, devmode, devmode,
                                      DM_IN_BUFFER | DM_OUT_BUFFER)
        if devmode.PaperSize != paper_id:
            raise PrintError(f"the driver does not accept {width_mm:g}×{length_mm:g} mm paper")
        win32print.SetPrinter(self._handle, 9, {"pDevMode": devmode}, 0)
        self._changed = True
        if paper_id == DMPAPER_USER:
            return devmode.PaperWidth / 10, devmode.PaperLength / 10
        return width_mm, length_mm

    def __exit__(self, *exc_info) -> bool:
        try:
            if self._changed:
                original = self._original
                if original is None:  # there were no per-user settings: copy the global ones
                    original = _first_devmode(self._handle, (8, 2))
                if original is not None:
                    win32print.SetPrinter(self._handle, 9, {"pDevMode": original}, 0)
        except Exception as exc:
            self.restore_error = exc
            log.warning("Could not restore settings of %s", self.printer, exc_info=True)
        finally:
            win32print.ClosePrinter(self._handle)
        return False


# ── PDF print programs ───────────────────────────────────────────────────────

class Backend:
    """Sends one PDF file to a printer."""
    key = ""

    def print_file(self, printer: str, path: str, copies: int = 1,
                   paper: Optional[str] = None) -> bool:
        """Print and return True once the job is known to be in the print queue.

        ``paper`` ("A4" or "A3") asks for that paper instead of the printer's
        default where the program supports it. False means the job was handed
        over but could not be confirmed. Raises on failure.
        """
        raise NotImplementedError


class SumatraBackend(Backend):
    key = "sumatra"

    def __init__(self, exe: str):
        self.exe = exe

    def print_file(self, printer, path, copies=1, paper=None):
        settings = []
        if copies > 1:
            settings.append(f"{copies}x")
        if paper:
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


class ShellBackend(_QueueWatchingBackend):
    key = "shell"

    def _launch(self, printer, path):
        win32api.ShellExecute(0, "printto", path, f'"{printer}"', os.path.dirname(path), 0)


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


def find_backend() -> Backend:
    exe = find_sumatra()
    if exe:
        return SumatraBackend(exe)
    exe = find_adobe()
    if exe:
        return AdobeBackend(exe)
    return ShellBackend()
