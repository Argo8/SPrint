"""Reading page sizes from PDF files and writing page subsets for printing."""

from __future__ import annotations

import logging
import os

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.errors import FileNotDecryptedError
    from pypdf.generic import NameObject
except ImportError:  # older installs: PyPDF2 >= 3 has the same API for what is used here
    from PyPDF2 import PdfReader, PdfWriter
    from PyPDF2.errors import FileNotDecryptedError
    from PyPDF2.generic import NameObject

from .paper import PT_TO_MM, DocInfo, PageInfo

# Malformed PDFs make pypdf log a warning per oddity; they are harmless here.
logging.getLogger("pypdf").setLevel(logging.ERROR)
logging.getLogger("PyPDF2").setLevel(logging.ERROR)


def _corners(rect) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = (float(v) for v in list(rect)[:4])
    return min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)


def page_size_mm(page) -> tuple[float, float]:
    """Visible size of a page in mm, as viewers display and print it.

    That is the crop box clipped to the media box (neither has to start at 0,0),
    scaled by /UserUnit and turned by /Rotate.
    """
    mx0, my0, mx1, my1 = _corners(page.mediabox)
    cx0, cy0, cx1, cy1 = _corners(page.cropbox)
    x0, y0, x1, y1 = max(mx0, cx0), max(my0, cy0), min(mx1, cx1), min(my1, cy1)
    if x1 <= x0 or y1 <= y0:  # crop box outside the media box: ignore it
        x0, y0, x1, y1 = mx0, my0, mx1, my1

    unit = page.get("/UserUnit", 1)
    unit = float(unit.get_object() if hasattr(unit, "get_object") else unit)
    scale = (unit if unit > 0 else 1.0) * PT_TO_MM
    width, height = (x1 - x0) * scale, (y1 - y0) * scale

    rotate = page.get("/Rotate", 0)
    rotate = int(rotate.get_object() if hasattr(rotate, "get_object") else rotate)
    if rotate % 180 == 90:
        width, height = height, width
    return width, height


def _describe(exc: Exception) -> str:
    return str(exc) or type(exc).__name__


def read_document(path: str) -> DocInfo:
    """Read the size of every page. Never raises; problems are recorded in the result."""
    doc = DocInfo(path)
    try:
        pages = PdfReader(path).pages
        doc.page_count = len(pages)
    except FileNotDecryptedError as exc:
        doc.encrypted = True
        doc.error = _describe(exc)
        return doc
    except Exception as exc:
        doc.error = _describe(exc)
        return doc

    for index in range(doc.page_count):
        try:
            doc.pages.append(PageInfo(index, *page_size_mm(pages[index])))
        except FileNotDecryptedError as exc:
            doc.encrypted = True
            doc.error = _describe(exc)
            doc.pages.clear()
            return doc
        except Exception as exc:
            doc.page_errors.append((index + 1, _describe(exc)))
    return doc


def _copy_optional_content(reader, writer) -> None:
    """Carry over layer settings (/OCProperties) so hidden or non-printing CAD layers stay hidden.

    pypdf copies the pages (and the layer objects they use) but not the document-level
    list that says which layers are switched off.
    """
    try:
        source_root = reader.trailer["/Root"]
        if "/OCProperties" not in source_root:
            return
        cloned = source_root["/OCProperties"].clone(writer)
        ref = getattr(cloned, "indirect_reference", None)
        target_root = getattr(writer, "root_object", None) or writer._root_object
        target_root[NameObject("/OCProperties")] = ref if ref is not None else cloned
    except Exception:
        logging.getLogger(__name__).warning("Could not copy layer settings", exc_info=True)


def write_subset(source_path: str, indices: list[int], target_path: str) -> None:
    """Write the given pages (0-based) of a PDF to a new file."""
    reader = PdfReader(source_path)
    writer = PdfWriter()
    for index in indices:
        writer.add_page(reader.pages[index])
    _copy_optional_content(reader, writer)
    with open(target_path, "wb") as fh:
        writer.write(fh)


def _unique_path(directory: str, stem: str, suffix: str) -> str:
    path = os.path.join(directory, f"{stem}{suffix}")
    n = 2
    while os.path.exists(path):
        path = os.path.join(directory, f"{stem} ({n}){suffix}")
        n += 1
    return path


def file_for_pages(doc: DocInfo, indices: list[int], temp_dir: str, label: str) -> str:
    """Path of a PDF containing exactly ``indices`` of ``doc``.

    When that is the whole document the original file is used as is, which keeps
    it untouched (layers, forms, annotations) and avoids a temporary copy.
    """
    if indices == list(range(doc.page_count)):
        return doc.path
    stem = os.path.splitext(doc.name)[0]
    target = _unique_path(temp_dir, f"{stem} [{label}]", ".pdf")
    write_subset(doc.path, indices, target)
    return target
