import os
import sys

import pytest
from pypdf import PdfWriter
from pypdf.generic import (ArrayObject, DecodedStreamObject, DictionaryObject, FloatObject, NameObject,
                           NumberObject, TextStringObject)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")   # window tests need no screen

from sprint import i18n  # noqa: E402

MM = 72 / 25.4


@pytest.fixture(autouse=True)
def english():
    i18n.set_language("en")
    yield
    i18n.set_language("en")


@pytest.fixture
def make_pdf(tmp_path):
    """make_pdf("name.pdf", [(w_mm, h_mm), ...]) -> path of a PDF with blank pages."""
    def make(name, sizes, tweak=None):
        writer = PdfWriter()
        for index, (w, h) in enumerate(sizes):
            page = writer.add_blank_page(w * MM, h * MM)
            if tweak:
                tweak(index, page, writer)
        path = tmp_path / name
        with open(path, "wb") as fh:
            writer.write(fh)
        return str(path)
    return make


def add_signature(page, writer, rect_mm=(30, 30, 100, 60)):
    """A visible signature on ``page``: a form field whose appearance is a black box.

    PDFium draws form fields only when asked to, so this is what goes missing if
    a page is drawn without them."""
    x, y, w, h = (v * MM for v in rect_mm)
    look = DecodedStreamObject()
    look.set_data(f"0 g 0 0 {w:.2f} {h:.2f} re f".encode())
    look.update({NameObject("/Type"): NameObject("/XObject"), NameObject("/Subtype"): NameObject("/Form"),
                 NameObject("/BBox"): ArrayObject([FloatObject(0), FloatObject(0), FloatObject(w), FloatObject(h)])})
    field = DictionaryObject({
        NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Widget"),
        NameObject("/FT"): NameObject("/Sig"), NameObject("/T"): TextStringObject("Signature1"),
        NameObject("/Rect"): ArrayObject([FloatObject(x), FloatObject(y), FloatObject(x + w), FloatObject(y + h)]),
        NameObject("/F"): NumberObject(4),   # printed
        NameObject("/AP"): DictionaryObject({NameObject("/N"): writer._add_object(look)}),
        NameObject("/P"): page.indirect_reference,
    })
    reference = writer._add_object(field)
    page[NameObject("/Annots")] = ArrayObject([reference])
    form = writer._root_object.setdefault(NameObject("/AcroForm"), DictionaryObject({NameObject("/Fields"): ArrayObject()}))
    form["/Fields"].append(reference)
