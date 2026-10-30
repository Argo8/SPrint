import string

import pytest

from sprint import i18n
from sprint.i18n import _NOUNS, _TEXTS, count, plural_index


def placeholders(text):
    return {name for _, name, _, _ in string.Formatter().parse(text) if name}


def test_languages_have_the_same_texts():
    assert set(_TEXTS["en"]) == set(_TEXTS["hr"])
    for key, text in _TEXTS["en"].items():
        assert placeholders(text) == placeholders(_TEXTS["hr"][key]), key
    assert set(_NOUNS["en"]) == set(_NOUNS["hr"])


@pytest.mark.parametrize("n, form", [
    (1, "stranica"), (2, "stranice"), (4, "stranice"), (5, "stranica"), (11, "stranica"),
    (12, "stranica"), (14, "stranica"), (21, "stranica"), (22, "stranice"), (25, "stranica"),
    (101, "stranica"), (111, "stranica"), (112, "stranica"), (0, "stranica"),
])
def test_croatian_plurals(n, form):
    i18n.set_language("hr")
    assert count(n, "page") == f"{n} {form}"


def test_croatian_plural_forms_differ_by_noun():
    i18n.set_language("hr")
    assert count(3, "file") == "3 datoteke"
    assert count(1, "group") == "1 grupa"
    assert count(5, "group") == "5 grupa"


def test_english_plurals():
    assert count(1, "page") == "1 page"
    assert count(0, "page") == "0 pages"
    assert plural_index("en", 2) == 1


def test_t_formats_and_falls_back():
    assert i18n.t("roll_mm", w=610) == "610 mm"
    assert i18n.t("no_such_key") == "no_such_key"
    with pytest.raises(ValueError):
        i18n.set_language("de")
