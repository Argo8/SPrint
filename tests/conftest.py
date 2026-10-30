import os
import sys

import pytest
from pypdf import PdfWriter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

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
