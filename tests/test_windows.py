"""Tests against the real Windows printing system (skipped elsewhere).

The end-to-end tests need a printer that writes its output to a file (CI adds
one with the "Microsoft Print To PDF" driver) and pass it in
SPRINT_TEST_PRINTER / SPRINT_TEST_OUTPUT. The output PDF shows which paper
the job used and where the page landed on it.
"""

import os
import sys
import time

import pytest
from pypdf.generic import DecodedStreamObject, NameObject

from conftest import MM
from sprint import printing
from sprint.pdfdocs import read_document

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="needs Windows")

if sys.platform == "win32":
    import win32print

BUILTIN_PRINTER = "Microsoft Print to PDF"
FILE_PRINTER = os.environ.get("SPRINT_TEST_PRINTER")
needs_file_printer = pytest.mark.skipif(not FILE_PRINTER, reason="no file printer configured")


@pytest.fixture
def builtin_printer():
    if BUILTIN_PRINTER not in printing.list_printers():
        pytest.skip(f"{BUILTIN_PRINTER} is not installed")
    return BUILTIN_PRINTER


def printer_paper(printer):
    handle = win32print.OpenPrinter(printer)
    try:
        dm = win32print.GetPrinter(handle, 2)["pDevMode"]
        return dm.PaperSize, dm.PaperWidth, dm.PaperLength, dm.Orientation
    finally:
        win32print.ClosePrinter(handle)


def test_printers(builtin_printer):
    assert printing.printer_exists(builtin_printer)
    assert not printing.printer_exists("No such printer 4711")
    assert isinstance(printing._job_ids(builtin_printer), set)


def test_job_settings_leave_the_printer_alone(builtin_printer):
    before = printer_paper(builtin_printer)
    assert printing.job_settings(builtin_printer, "A3", copies=2).PaperSize == 8
    with pytest.raises(printing.PrintError):   # this driver has no custom sizes
        printing.job_settings(builtin_printer, (610, 1234))
    assert printer_paper(builtin_printer) == before


def test_job_settings_use_this_users_printing_preferences(builtin_printer):
    """What the printer's settings dialog (the ⚙ button) changes reaches SPrint's jobs."""
    handle = win32print.OpenPrinter(builtin_printer, {"DesiredAccess": win32print.PRINTER_ACCESS_USE})
    try:
        before = win32print.GetPrinter(handle, 9)["pDevMode"]
        mine = win32print.GetPrinter(handle, 2)["pDevMode"]
        landscape = 3 - mine.Orientation   # the other orientation
        mine.Orientation = landscape
        mine.Fields |= printing.DM_ORIENTATION
        win32print.SetPrinter(handle, 9, {"pDevMode": mine}, 0)
        try:
            assert printing.job_settings(builtin_printer).Orientation == landscape
        finally:
            if before is not None:
                win32print.SetPrinter(handle, 9, {"pDevMode": before}, 0)
            else:
                mine.Orientation = 3 - landscape
                win32print.SetPrinter(handle, 9, {"pDevMode": mine}, 0)
    finally:
        win32print.ClosePrinter(handle)


def boxed_pdf(make_pdf, name, size_mm):
    """A page with a black square near its top-left corner."""
    def draw(_, page, writer):
        top = size_mm[1] * MM
        box = DecodedStreamObject()
        box.set_data(f"0 g {20 * MM:.2f} {top - 80 * MM:.2f} {60 * MM:.2f} {60 * MM:.2f} re f".encode())
        page[NameObject("/Contents")] = writer._add_object(box)
    return make_pdf(name, [size_mm], draw)


def print_and_inspect(backend, path, **kwargs):
    """Print to the file printer; return the paper size (mm) and where the ink is (0..1)."""
    import pypdfium2 as pdfium

    output = os.environ["SPRINT_TEST_OUTPUT"]
    if os.path.exists(output):
        os.remove(output)
    assert backend.print_file(FILE_PRINTER, path, **kwargs)
    deadline = time.monotonic() + 60
    while not (os.path.exists(output) and os.path.getsize(output) > 0):
        assert time.monotonic() < deadline, "no output from the printer"
        time.sleep(0.5)
    time.sleep(2)
    doc = read_document(output)
    assert doc.ok, doc.error
    paper = round(doc.pages[0].width_mm), round(doc.pages[0].height_mm)
    pdf = pdfium.PdfDocument(output)
    image = pdf[0].render(scale=0.25).to_pil().convert("L")
    pdf.close()
    ink = image.point(lambda v: 255 if v < 128 else 0).getbbox()
    assert ink, "the page came out blank"
    centre = ((ink[0] + ink[2]) / 2 / image.width, (ink[1] + ink[3]) / 2 / image.height)
    return paper, centre


@needs_file_printer
def test_direct_print_uses_the_requested_paper(make_pdf):
    backend = printing.DirectBackend()
    portrait = boxed_pdf(make_pdf, "portrait.pdf", (297, 420))
    paper, (x, y) = print_and_inspect(backend, portrait, paper=(297, 420))
    assert paper == (297, 420)
    assert x < 0.5 and y < 0.5                      # the square stays top-left

    a4 = boxed_pdf(make_pdf, "a4.pdf", (210, 297))
    assert print_and_inspect(backend, a4, paper="A4")[0] == (210, 297)


@needs_file_printer
def test_direct_print_turns_landscape_pages(make_pdf):
    landscape = boxed_pdf(make_pdf, "landscape.pdf", (420, 297))
    paper, (x, y) = print_and_inspect(printing.DirectBackend(), landscape, paper="A3")
    assert paper == (297, 420)
    assert x > 0.5 and y < 0.5                      # turned clockwise: top-left goes top-right


@needs_file_printer
def test_sumatra_paper_end_to_end(make_pdf):
    exe = printing.find_sumatra()
    if not exe:
        pytest.skip("SumatraPDF not installed")
    backend = printing.SumatraBackend(exe)
    a3 = boxed_pdf(make_pdf, "a3.pdf", (297, 420))
    assert print_and_inspect(backend, a3)[0] == (297, 420)              # matching paper picked
    assert print_and_inspect(backend, a3, paper="A4")[0] == (210, 297)  # explicit paper wins
