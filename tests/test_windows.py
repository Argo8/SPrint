"""Tests against the real Windows printing system (skipped elsewhere).

The end-to-end test needs a printer that writes to a file and SumatraPDF;
CI sets both up and passes them in SPRINT_TEST_PRINTER / SPRINT_TEST_OUTPUT.
"""

import os
import sys
import time

import pytest

from sprint import printing
from sprint.pdfdocs import read_document

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="needs Windows")

if sys.platform == "win32":
    import win32print

BUILTIN_PRINTER = "Microsoft Print to PDF"


@pytest.fixture
def builtin_printer():
    if BUILTIN_PRINTER not in printing.list_printers():
        pytest.skip(f"{BUILTIN_PRINTER} is not installed")
    return BUILTIN_PRINTER


def paper_settings(printer):
    handle = win32print.OpenPrinter(printer)
    try:
        dm = printing._first_devmode(handle, (9, 8, 2))
        return dm.PaperSize, dm.PaperWidth, dm.PaperLength, dm.Orientation
    finally:
        win32print.ClosePrinter(handle)


def test_printers(builtin_printer):
    assert printing.printer_exists(builtin_printer)
    assert not printing.printer_exists("No such printer 4711")
    assert isinstance(printing._job_ids(builtin_printer), set)


def test_paper_override_is_restored(builtin_printer):
    before = paper_settings(builtin_printer)
    with printing.PlotterPaper(builtin_printer) as paper:
        assert paper.apply(297, 420) == (297, 420)
        assert paper_settings(builtin_printer)[0] == 8   # A3
    assert paper.restore_error is None
    assert paper_settings(builtin_printer) == before


def test_rejected_custom_size_leaves_settings_alone(builtin_printer):
    before = paper_settings(builtin_printer)
    with printing.PlotterPaper(builtin_printer) as paper:
        try:
            paper.apply(610, 1234)
        except printing.PrintError:
            pass
    assert paper_settings(builtin_printer) == before


@pytest.mark.skipif(not os.environ.get("SPRINT_TEST_PRINTER"), reason="no file printer configured")
def test_print_with_paper_override_end_to_end(make_pdf):
    printer = os.environ["SPRINT_TEST_PRINTER"]
    output = os.environ["SPRINT_TEST_OUTPUT"]
    exe = printing.find_sumatra()
    assert exe, "SumatraPDF not found"
    backend = printing.SumatraBackend(exe)
    source = make_pdf("a3.pdf", [(297, 420)])

    def print_and_measure(**paper_size):
        if os.path.exists(output):
            os.remove(output)
        if paper_size:
            with printing.PlotterPaper(printer) as paper:
                paper.apply(**paper_size)
                assert backend.print_file(printer, source)
        else:
            assert backend.print_file(printer, source)
        deadline = time.monotonic() + 60
        while not (os.path.exists(output) and os.path.getsize(output) > 0):
            assert time.monotonic() < deadline, "no output from the printer"
            time.sleep(0.5)
        time.sleep(2)
        doc = read_document(output)
        assert doc.ok, doc.error
        page = doc.pages[0]
        return round(min(page.width_mm, page.height_mm)), round(max(page.width_mm, page.height_mm))

    assert print_and_measure(width_mm=297, length_mm=420) == (297, 420)
    assert print_and_measure() != (297, 420)   # default paper again once restored
