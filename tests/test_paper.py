import pytest

from sprint.paper import A3, A4, OTHER, SMALL, classify, exceeds_rolls, plot_paper, roll_for


@pytest.mark.parametrize("w, h, kind", [
    (210, 297, A4), (297, 210, A4), (209.9, 297.0, A4), (212.9, 294.1, A4),
    (297, 420, A3), (420, 297, A3), (296.9, 420.2, A3),
    (213.5, 297, OTHER),                 # beyond the 3 mm tolerance: wider than an A4 sheet
    (215.9, 279.4, OTHER),               # US Letter: wider than an A4 sheet
    (148, 210, SMALL), (210, 148, SMALL),  # A5
    (105, 148, SMALL), (99, 210, SMALL),   # A6, DL
    (205, 290, SMALL),                   # a little smaller than A4
    (200, 310, OTHER),                   # narrower than A4 but longer
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


@pytest.mark.parametrize("w, h, kind, paper", [
    (841, 594, OTHER, (610, 841)),         # short side across the roll
    (594, 420, OTHER, (420, 594)),
    (297, 630.4, OTHER, (297, 630)),
    (298, 422, A3, (297, 420)),            # A3 is always a whole A3 sheet off the 297 roll
    (420, 297, A3, (297, 420)),
    (1000, 1500, OTHER, (914, 1371)),      # too wide: printed smaller, paper cut to match
    (148, 210, SMALL, (297, 210)),         # smaller than A4, sent to the plotter
])
def test_plot_paper(w, h, kind, paper):
    assert plot_paper(w, h, kind) == paper
