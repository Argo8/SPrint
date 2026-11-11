import pytest
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, DictionaryObject, NameObject, NumberObject,
                           RectangleObject, TextStringObject)

from conftest import MM
from sprint.paper import A3, A4, OTHER
from sprint.pdfdocs import file_for_pages, read_document, render_preview, write_subset


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


def test_signatures_and_form_fields_are_drawn(make_pdf, tmp_path):
    from conftest import add_signature
    path = make_pdf("signed.pdf", [(210, 297), (210, 297)],
                    lambda index, page, writer: index == 1 and add_signature(page, writer))

    def ink(path, index):
        width, height, rgba = render_preview(path, index, 300)
        return sum(1 for i in range(0, len(rgba), 4) if rgba[i] < 100)
    assert ink(path, 0) == 0
    assert ink(path, 1) > 500                      # the 70×30 mm box
    # also when the page is taken out of the file for printing
    subset = file_for_pages(read_document(path), [1], str(tmp_path), "A4")
    assert ink(subset, 0) > 500


def test_blank_margins_of_a_drawing(make_pdf):
    import pypdfium2 as pdfium
    from pypdf.generic import DecodedStreamObject

    from sprint.printing import blank_margins_mm

    def frame(inset_left, inset_top, inset_right, inset_bottom, rotate=0):
        def draw(_, page, writer):
            w, h = 297, 420
            box = DecodedStreamObject()
            x0, y0 = inset_left * MM, inset_bottom * MM
            x1, y1 = (w - inset_right) * MM, (h - inset_top) * MM
            box.set_data(f"0.7 w 0 G {x0:.2f} {y0:.2f} {x1 - x0:.2f} {y1 - y0:.2f} re S".encode())
            page[NameObject("/Contents")] = writer._add_object(box)
            if rotate:
                page[NameObject("/Rotate")] = NumberObject(rotate)
        return draw

    def margins(name, tweak):
        pdf = pdfium.PdfDocument(make_pdf(name, [(297, 420)], tweak))
        try:
            return blank_margins_mm(pdf[0])
        finally:
            pdf.close()

    left, top, right, bottom = margins("iso.pdf", frame(20, 10, 10, 10))
    assert 18.9 <= left <= 20 and all(8.9 <= m <= 10 for m in (top, right, bottom))
    assert margins("edge.pdf", frame(0, 10, 10, 10))[0] < 1
    # shown turned a quarter clockwise: the wide left margin is now at the top
    left, top, right, bottom = margins("turned.pdf", frame(20, 10, 10, 10, rotate=90))
    assert 18.9 <= top <= 20 and all(8.9 <= m <= 10 for m in (left, right, bottom))
    assert margins("blank.pdf", None) == (0.0, 0.0, 0.0, 0.0)


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
