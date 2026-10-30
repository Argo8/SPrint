import pytest
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, DictionaryObject, NameObject, NumberObject,
                           RectangleObject, TextStringObject)

from conftest import MM
from sprint.paper import A3, A4, OTHER
from sprint.pdfdocs import file_for_pages, read_document, write_subset


def sizes(doc):
    return [(round(p.width_mm), round(p.height_mm), p.kind) for p in doc.pages]


def test_reads_page_sizes(make_pdf):
    doc = read_document(make_pdf("mixed.pdf", [(210, 297), (420, 297), (841, 1189)]))
    assert doc.ok and doc.page_count == 3
    assert sizes(doc) == [(210, 297, A4), (420, 297, A3), (841, 1189, OTHER)]


def test_media_box_not_at_origin(make_pdf):
    def shift(_, page, __):
        page.mediabox = RectangleObject([-40, -60, 210 * MM - 40, 297 * MM - 60])
    assert sizes(read_document(make_pdf("shifted.pdf", [(210, 297)], shift))) == [(210, 297, A4)]


def test_crop_box_is_what_prints(make_pdf):
    def crop(_, page, __):
        page.cropbox = RectangleObject([20, 20, 20 + 210 * MM, 20 + 297 * MM])
    assert sizes(read_document(make_pdf("crop.pdf", [(297, 420)], crop))) == [(210, 297, A4)]


def test_rotate_and_user_unit(make_pdf):
    def tweak(_, page, __):
        page[NameObject("/Rotate")] = NumberObject(90)
        page[NameObject("/UserUnit")] = NumberObject(2)
    doc = read_document(make_pdf("big.pdf", [(594.5, 420.5)], tweak))
    assert sizes(doc) == [(841, 1189, OTHER)]


def test_unreadable_file(tmp_path):
    path = tmp_path / "broken.pdf"
    path.write_bytes(b"%PDF-1.4 garbage")
    doc = read_document(str(path))
    assert not doc.ok and doc.error and doc.pages == []


def test_missing_file(tmp_path):
    assert not read_document(str(tmp_path / "missing.pdf")).ok


def test_password_protected(make_pdf, tmp_path):
    reader = PdfReader(make_pdf("plain.pdf", [(210, 297)]))
    writer = PdfWriter(clone_from=reader)
    writer.encrypt(user_password="secret", owner_password="owner", algorithm="RC4-128")
    path = tmp_path / "locked.pdf"
    with open(path, "wb") as fh:
        writer.write(fh)
    doc = read_document(str(path))
    assert doc.encrypted and not doc.ok


def test_file_for_pages_uses_original_when_complete(make_pdf, tmp_path):
    doc = read_document(make_pdf("all_a4.pdf", [(210, 297)] * 3))
    assert file_for_pages(doc, [0, 1, 2], str(tmp_path), "A4") == doc.path


def test_file_for_pages_writes_subset(make_pdf, tmp_path):
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    doc = read_document(make_pdf("mixed.pdf", [(210, 297), (420, 297), (210, 297)]))
    first = file_for_pages(doc, [0, 2], str(out_dir), "A4")
    second = file_for_pages(doc, [0, 2], str(out_dir), "A4")
    assert first != second  # never overwrite a file a PDF app may still be reading
    assert sizes(read_document(first)) == [(210, 297, A4), (210, 297, A4)]


def test_subset_keeps_layer_visibility(tmp_path):
    writer = PdfWriter()
    writer.add_blank_page(210 * MM, 297 * MM)
    page = writer.add_blank_page(420 * MM, 297 * MM)
    layer = writer._add_object(DictionaryObject({
        NameObject("/Type"): NameObject("/OCG"),
        NameObject("/Name"): TextStringObject("Hidden layer"),
    }))
    page[NameObject("/Resources")] = DictionaryObject({
        NameObject("/Properties"): DictionaryObject({NameObject("/OC1"): layer})})
    writer._root_object[NameObject("/OCProperties")] = DictionaryObject({
        NameObject("/OCGs"): ArrayObject([layer]),
        NameObject("/D"): DictionaryObject({NameObject("/OFF"): ArrayObject([layer])}),
    })
    source = tmp_path / "layers.pdf"
    with open(source, "wb") as fh:
        writer.write(fh)

    target = tmp_path / "subset.pdf"
    write_subset(str(source), [1], str(target))

    reader = PdfReader(str(target))
    assert len(reader.pages) == 1
    off = reader.trailer["/Root"]["/OCProperties"]["/D"]["/OFF"][0].get_object()
    used = reader.pages[0]["/Resources"]["/Properties"]["/OC1"].get_object()
    assert off["/Name"] == "Hidden layer"
    assert off.indirect_reference.idnum == used.indirect_reference.idnum


@pytest.mark.parametrize("indices", [[0], [1, 0]])
def test_write_subset_page_order(make_pdf, tmp_path, indices):
    source = make_pdf("two.pdf", [(210, 297), (420, 297)])
    target = tmp_path / "t.pdf"
    write_subset(source, indices, str(target))
    expected = [[(210, 297, A4)], [(420, 297, A3)]]
    assert sizes(read_document(str(target))) == [expected[i][0] for i in indices]
