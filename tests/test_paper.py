import pytest

from sprint.paper import (A3, A4, OTHER, PLOT_A3, DocInfo, PageInfo, classify, exceeds_rolls,
                          plot_groups, plot_option_for, roll_for, summarize)


@pytest.mark.parametrize("w, h, kind", [
    (210, 297, A4), (297, 210, A4), (209.9, 297.0, A4), (212.9, 294.1, A4),
    (297, 420, A3), (420, 297, A3), (296.9, 420.2, A3),
    (213.5, 297, OTHER),                 # beyond the 3 mm tolerance
    (215.9, 279.4, OTHER),               # US Letter
    (148, 210, OTHER),                   # A5
    (594, 841, OTHER), (297, 630, OTHER),
])
def test_classify(w, h, kind):
    assert classify(w, h) == kind


@pytest.mark.parametrize("w, h, roll", [
    (216, 279, 297), (297.4, 630, 297), (298, 600, 420), (420, 1189, 420),
    (594, 841, 610), (841, 594, 610), (841, 1189, 914), (1000, 1500, 914),
])
def test_roll_for(w, h, roll):
    assert roll_for(w, h) == roll


def test_exceeds_rolls():
    assert not exceeds_rolls(841, 1189)
    assert not exceeds_rolls(914.3, 2000)
    assert exceeds_rolls(1000, 1500)


def test_plot_option_for():
    assert plot_option_for(210, 297) is None
    assert plot_option_for(420, 297) == PLOT_A3
    assert plot_option_for(594, 841).key == "610"
    assert plot_option_for(1189, 841).key == "914"


def doc(name, *sizes):
    return DocInfo(name, page_count=len(sizes),
                   pages=[PageInfo(i, w, h) for i, (w, h) in enumerate(sizes)])


def test_plot_groups_by_length_longest_first():
    a = doc("a.pdf", (594, 841), (210, 297), (610, 1000))
    b = doc("b.pdf", (841, 594), (600, 1000.4))
    groups = plot_groups([a, b], "610")
    assert [(g.paper_width_mm, g.length_mm) for g in groups] == [(610, 1000), (610, 841)]
    assert [(d.path, idx) for d, idx in groups[0].parts] == [("a.pdf", [2]), ("b.pdf", [1])]
    assert [(d.path, idx) for d, idx in groups[1].parts] == [("a.pdf", [0]), ("b.pdf", [0])]
    assert groups[1].page_count == 2


def test_plot_groups_a3_always_420_long():
    groups = plot_groups([doc("a.pdf", (297, 420), (298, 422), (420, 297))], "A3")
    assert len(groups) == 1
    assert (groups[0].paper_width_mm, groups[0].length_mm, groups[0].page_count) == (297, 420, 3)


def test_plot_groups_empty():
    assert plot_groups([doc("a.pdf", (210, 297))], "914") == []


def test_summarize():
    broken = DocInfo("broken.pdf", error="not a PDF")
    with_errors = doc("e.pdf", (210, 297))
    with_errors.page_errors.append((2, "bad page"))
    s = summarize([doc("a.pdf", (210, 297), (420, 297), (841, 1189), (1000, 1500)),
                   broken, with_errors])
    assert s.files == 3
    assert s.total_pages == 5
    assert s.kinds == {A4: 2, A3: 1, OTHER: 2}
    assert s.other_sizes == {(841, 1189): 1, (1000, 1500): 1}
    assert s.plot_counts == {"A3": 1, "914": 2}
    assert s.oversize == 1
    assert s.unreadable == [broken]
    assert s.page_errors == [(with_errors, 2, "bad page")]
