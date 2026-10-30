"""English / Croatian UI texts."""

from __future__ import annotations

import locale
import os

LANGUAGES = ("en", "hr")

# Nouns with plural forms: English (one, other), Croatian (one, few, many).
_NOUNS = {
    "en": {
        "page": ("page", "pages"),
        "page_acc": ("page", "pages"),   # "print 1 page"
        "file": ("file", "files"),
        "sheet": ("sheet", "sheets"),
        "copy": ("copy", "copies"),
        "printer": ("printer", "printers"),
        "skipped": ("skipped", "skipped"),
        "bad_page": ("unreadable page", "unreadable pages"),
    },
    "hr": {
        "page": ("stranica", "stranice", "stranica"),
        "page_acc": ("stranicu", "stranice", "stranica"),   # "ispiši 1 stranicu"
        "file": ("datoteka", "datoteke", "datoteka"),
        "sheet": ("list", "lista", "listova"),
        "copy": ("kopija", "kopije", "kopija"),
        "printer": ("pisač", "pisača", "pisača"),
        "skipped": ("preskočena", "preskočene", "preskočenih"),
        "bad_page": ("nečitljiva stranica", "nečitljive stranice", "nečitljivih stranica"),
    },
}

_TEXTS = {
    "en": {
        # header
        "settings":          "Settings",
        # files
        "files_title":       "Files",
        "files_none":        "No files yet",
        "files_summary":     "{files} · {pages}",
        "files_reading":     "{files} · reading…",
        "add_pdf":           "Add PDF",
        "clear":             "Clear",
        "drop_title":        "Drag PDF files here",
        "drop_hint":         "or right-click a PDF in Explorer → Send to → SPrint",
        "choose_files":      "Choose files",
        "drop_release":      "Release to add",
        "list_hint":         "Drop more PDFs here · click a file to see its pages",
        "pages_hint":        "Click a page to skip it or bring it back.",
        "file_pages":        "{n} p.",
        "chip_a4":           "{n}× A4",
        "chip_a3":           "{n}× A3",
        "chip_large":        "{n}× large format",
        "chip_done_part":    "{done} of {total} printed",
        "chip_oversize":     "too wide · printed smaller",
        "status_reading":    "Reading…",
        "status_printed":    "Printed",
        "status_error":      "Can't open",
        "status_locked":     "Password protected",
        "tip_remove":        "Remove {name}",
        "tip_show_pages":    "Show the pages of {name}",
        "tip_hide_pages":    "Hide the pages of {name}",
        "dest_a4":           "A4 printer",
        "dest_a3":           "A3 printer",
        "dest_roll":         "Roll {w}",
        "dest_skipped":      "Skipped",
        "page_tip":          "Page {n} · {size} · {dest}",
        # destinations
        "where_title":       "Where it goes",
        "a4_title":          "A4 pages",
        "a3_title":          "A3 pages",
        "plot_title":        "Plotter",
        "printer":           "Printer",
        "choose_printer":    "Choose a printer",
        "duplex":            "Two-sided",
        "reprint":           "Print again",
        "replot":            "Plot again",
        "printer_settings":  "Printer settings: trays, quality, paper…",
        "sum_no_a4":         "No A4 pages",
        "sum_no_a3":         "No A3 pages",
        "sum_done":          "{pages} · printed ✓",
        "sum_sheets":        "{pages} · {sheets}",
        "sum_a3_plotter":    "{pages} · on the plotter, roll 297",
        "a3_target":         "Print on",
        "a3_on_printer":     "A3 printer",
        "a3_on_plotter":     "Plotter (roll 297)",
        "a3_plotter_note":   "A3 pages go on the plotter's 297 mm roll: choose roll 297 below and plot.",
        "plot_none":         "No large formats",
        "plot_all_done":     "{pages} · all plotted ✓",
        "plot_needs_one":    "{pages} · needs the {rolls} mm roll",
        "plot_needs_many":   "{pages} · needs rolls {rolls} mm",
        "roll_label":        "Roll in the plotter now",
        "roll_chip_none":    "none",
        "roll_chip_pages":   "{n} p.",
        "roll_chip_done":    "✓ done",
        "roll_title":        "Roll {w} mm",
        "roll_info_empty":   "No pages need the {w} mm roll.",
        "roll_info_done":    "Roll {w} mm: everything is plotted.",
        "roll_info":         "Roll {w} mm: {pages} · ≈ {meters} of paper ({lengths})",
        "roll_later":        "Then still needed: {rolls} mm",
        "roll_next":         "Next: load the {w} mm roll and choose it above ({pages}).",
        "roll_oversize":     "{pages} wider than the widest roll: printed smaller.",
        # print buttons
        "btn_print":         "Print {pages}",
        "btn_plot":          "Plot {pages}",
        "btn_print_empty":   "Print",
        "btn_plot_empty":    "Plot",
        "btn_reading":       "Reading the files…",
        "btn_pick_printer":  "Choose a printer first",
        "btn_all_done":      "Everything printed",
        "btn_roll_done":     "Roll {w} is done",
        "btn_roll_empty":    "Nothing for roll {w}",
        "copies":            "Copies",
        "tip_fewer":         "Fewer copies",
        "tip_more":          "More copies",
        "sending":           "Sending…",
        "waiting":           "Waiting its turn…",
        "stop":              "Stop",
        "cancel":            "Cancel",
        "card_sent":         "Sent",
        # toasts
        "toast_sent":        "Sent: {pages}",
        "toast_sent_next":   "Sent: {pages} · next roll: {w} mm",
        "toast_stopped":     "Printing stopped",
        "toast_pick_printer": "Choose a printer for {target}",
        "toast_problems":    "Some pages were not sent, see the log in Settings",
        "toast_detected":    "Printers found: {names}",
        "toast_detect_none": "No printers found",
        "toast_not_pdf":     "Only PDF files can be added",
        "toast_changed":     "{name} changed and was read again: check it and print again",
        # settings
        "close_settings":    "Close settings",
        "sec_print":         "Printing",
        "clear_when_done":   "Clear the list when everything is printed",
        "clear_when_done_hint": "Ready for the next project.",
        "sec_printers":      "Printers",
        "detect":            "Find printers automatically",
        "language":          "Language",
        "sec_log":           "Log",
        "open_log_folder":   "Open the log folder",
        "about":             "SPrint {version} · {author}",
        # dialogs
        "dlg_add_title":     "Choose PDF files",
        "dlg_pdf_files":     "PDF files",
        "dlg_close_title":   "Printing in progress",
        "dlg_close_msg":     "Pages are still being sent to the printer.\nStop after the current file and close?",
        "dlg_stop_close":    "Stop and close",
        "dlg_keep_printing": "Keep printing",
        # print log
        "backend_direct":    "Printing directly with the built-in PDF renderer.",
        "backend_sumatra":   "A4/A3 printing via SumatraPDF, plotting directly.",
        "backend_adobe":     "A4/A3 printing via Adobe Reader, plotting directly.",
        "program_missing":   "Print program \"{program}\" not found, printing directly.",
        "target_a4":         "A4 pages",
        "target_a3":         "A3 pages",
        "target_plot":       "the plotter",
        "job_target":        "{target} → {printer}",
        "no_printer":        "No printer chosen for {target}.",
        "printer_missing":   "Printer not found: {printer}",
        "plot_run":          "Roll {w} mm, paper {length} mm long: {pages}",
        "print_sent":        "Sent {name} ({pages})",
        "send_error":        "{name}: {error}",
        "temp_error":        "Could not prepare {name}: {error}",
        "unconfirmed":       "Could not confirm that the job reached the print queue.",
        "cancelled":         "Printing stopped.",
        "unexpected":        "Unexpected error: {error}",
        "read_error":        "{name}: {error}",
    },
    "hr": {
        # header
        "settings":          "Postavke",
        # files
        "files_title":       "Datoteke",
        "files_none":        "Još nema datoteka",
        "files_summary":     "{files} · {pages}",
        "files_reading":     "{files} · još se čita…",
        "add_pdf":           "Dodaj PDF",
        "clear":             "Očisti",
        "drop_title":        "Povuci PDF datoteke ovdje",
        "drop_hint":         "ili desni klik na PDF u Exploreru → Pošalji u → SPrint",
        "choose_files":      "Odaberi datoteke",
        "drop_release":      "Pusti za dodavanje",
        "list_hint":         "Povuci još PDF-ova ovdje · klikni datoteku za pregled stranica",
        "pages_hint":        "Klikni stranicu da je preskočiš ili vratiš.",
        "file_pages":        "{n} str.",
        "chip_a4":           "{n}× A4",
        "chip_a3":           "{n}× A3",
        "chip_large":        "{n}× veliki format",
        "chip_done_part":    "{done} od {total} ispisano",
        "chip_oversize":     "preširoko · ispisat će se umanjeno",
        "status_reading":    "Čitam…",
        "status_printed":    "Ispisano",
        "status_error":      "Ne može se otvoriti",
        "status_locked":     "Zaštićeno lozinkom",
        "tip_remove":        "Ukloni {name}",
        "tip_show_pages":    "Prikaži stranice: {name}",
        "tip_hide_pages":    "Sakrij stranice: {name}",
        "dest_a4":           "A4 pisač",
        "dest_a3":           "A3 pisač",
        "dest_roll":         "Rola {w}",
        "dest_skipped":      "Preskočeno",
        "page_tip":          "Stranica {n} · {size} · {dest}",
        # destinations
        "where_title":       "Kamo ide",
        "a4_title":          "A4 stranice",
        "a3_title":          "A3 stranice",
        "plot_title":        "Ploter",
        "printer":           "Pisač",
        "choose_printer":    "Odaberi pisač",
        "duplex":            "Obostrano",
        "reprint":           "Ispiši ponovno",
        "replot":            "Plotaj ponovno",
        "printer_settings":  "Postavke pisača: ladice, kvaliteta, papir…",
        "sum_no_a4":         "Nema A4 stranica",
        "sum_no_a3":         "Nema A3 stranica",
        "sum_done":          "{pages} · ispisano ✓",
        "sum_sheets":        "{pages} · {sheets}",
        "sum_a3_plotter":    "{pages} · na ploter, rola 297",
        "a3_target":         "Ispiši na",
        "a3_on_printer":     "A3 pisač",
        "a3_on_plotter":     "Ploter (rola 297)",
        "a3_plotter_note":   "A3 stranice idu na ploter, na rolu 297 mm: dolje odaberi rolu 297 i plotaj.",
        "plot_none":         "Nema velikih formata",
        "plot_all_done":     "{pages} · sve isplotano ✓",
        "plot_needs_one":    "{pages} · treba rola {rolls} mm",
        "plot_needs_many":   "{pages} · trebaju role {rolls} mm",
        "roll_label":        "Rola koja je sada u ploteru",
        "roll_chip_none":    "nema",
        "roll_chip_pages":   "{n} str.",
        "roll_chip_done":    "✓ gotovo",
        "roll_title":        "Rola {w} mm",
        "roll_info_empty":   "Za rolu {w} mm nema stranica.",
        "roll_info_done":    "Rola {w} mm: sve je isplotano.",
        "roll_info":         "Rola {w} mm: {pages} · ≈ {meters} papira ({lengths})",
        "roll_later":        "Poslije trebaju još role: {rolls} mm",
        "roll_next":         "Sljedeće: stavi rolu {w} mm i odaberi je gore ({pages}).",
        "roll_oversize":     "{pages} šire od najšire role: ispisat će se umanjeno.",
        # print buttons
        "btn_print":         "Ispiši {pages}",
        "btn_plot":          "Plotaj {pages}",
        "btn_print_empty":   "Ispiši",
        "btn_plot_empty":    "Plotaj",
        "btn_reading":       "Čitam datoteke…",
        "btn_pick_printer":  "Prvo odaberi pisač",
        "btn_all_done":      "Sve je ispisano",
        "btn_roll_done":     "Rola {w} je gotova",
        "btn_roll_empty":    "Nema stranica za rolu {w}",
        "copies":            "Kopije",
        "tip_fewer":         "Manje kopija",
        "tip_more":          "Više kopija",
        "sending":           "Šaljem…",
        "waiting":           "Čeka na red…",
        "stop":              "Zaustavi",
        "cancel":            "Odustani",
        "card_sent":         "Poslano",
        # toasts
        "toast_sent":        "Poslano: {pages}",
        "toast_sent_next":   "Poslano: {pages} · sljedeća rola: {w} mm",
        "toast_stopped":     "Ispis je zaustavljen",
        "toast_pick_printer": "Odaberi pisač za {target}",
        "toast_problems":    "Neke stranice nisu poslane, pogledaj dnevnik u Postavkama",
        "toast_detected":    "Prepoznati pisači: {names}",
        "toast_detect_none": "Nije pronađen nijedan pisač",
        "toast_not_pdf":     "Mogu se dodati samo PDF datoteke",
        "toast_changed":     "{name} se promijenila pa je ponovno pročitana: provjeri je i ispiši opet",
        # settings
        "close_settings":    "Zatvori postavke",
        "sec_print":         "Ispis",
        "clear_when_done":   "Kad je sve ispisano, očisti popis",
        "clear_when_done_hint": "Spremno za sljedeći projekt.",
        "sec_printers":      "Pisači",
        "detect":            "Prepoznaj pisače automatski",
        "language":          "Jezik",
        "sec_log":           "Dnevnik",
        "open_log_folder":   "Otvori mapu dnevnika",
        "about":             "SPrint {version} · {author}",
        # dialogs
        "dlg_add_title":     "Odaberite PDF datoteke",
        "dlg_pdf_files":     "PDF datoteke",
        "dlg_close_title":   "Ispis je u tijeku",
        "dlg_close_msg":     "Stranice se još šalju na pisač.\nZaustaviti nakon trenutne datoteke i zatvoriti?",
        "dlg_stop_close":    "Zaustavi i zatvori",
        "dlg_keep_printing": "Nastavi ispis",
        # print log
        "backend_direct":    "Ispis izravno, ugrađenim PDF prikazom.",
        "backend_sumatra":   "A4/A3 ispis preko SumatraPDF-a, plotanje izravno.",
        "backend_adobe":     "A4/A3 ispis preko Adobe Readera, plotanje izravno.",
        "program_missing":   "Program za ispis „{program}” nije pronađen, ispisujem izravno.",
        "target_a4":         "A4 stranice",
        "target_a3":         "A3 stranice",
        "target_plot":       "ploter",
        "job_target":        "{target} → {printer}",
        "no_printer":        "Nije odabran pisač za {target}.",
        "printer_missing":   "Pisač nije pronađen: {printer}",
        "plot_run":          "Rola {w} mm, papir dug {length} mm: {pages}",
        "print_sent":        "Poslano: {name} ({pages})",
        "send_error":        "{name}: {error}",
        "temp_error":        "Ne mogu pripremiti {name}: {error}",
        "unconfirmed":       "Nije potvrđeno da je posao stigao u red ispisa.",
        "cancelled":         "Ispis je zaustavljen.",
        "unexpected":        "Neočekivana greška: {error}",
        "read_error":        "{name}: {error}",
    },
}

_current = "en"


def set_language(lang: str) -> None:
    global _current
    if lang not in LANGUAGES:
        raise ValueError(f"unsupported language: {lang}")
    _current = lang


def get_language() -> str:
    return _current


def t(key: str, **kwargs) -> str:
    text = _TEXTS[_current].get(key) or _TEXTS["en"].get(key, key)
    return text.format(**kwargs) if kwargs else text


def plural_index(lang: str, n: int) -> int:
    n = abs(n)
    if lang == "hr":
        if n % 10 == 1 and n % 100 != 11:
            return 0
        if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
            return 1
        return 2
    return 0 if n == 1 else 1


def count(n: int, noun: str) -> str:
    """'5 pages' / '5 stranica' with the right plural form."""
    forms = _NOUNS[_current][noun]
    return f"{n} {forms[plural_index(_current, n)]}"


def meters(mm: float) -> str:
    """Paper length in metres, one decimal: '2.1 m' / '2,1 m'."""
    text = f"{mm / 1000:.1f}"
    return (text.replace(".", ",") if _current == "hr" else text) + " m"


def join_and(items) -> str:
    """'420, 610 and 914' / '420, 610 i 914'."""
    items = [str(item) for item in items]
    if len(items) < 2:
        return "".join(items)
    return ", ".join(items[:-1]) + (" i " if _current == "hr" else " and ") + items[-1]


def system_language() -> str:
    """'hr' on Croatian (and related) Windows/locales, otherwise 'en'."""
    if os.name == "nt":
        try:
            import ctypes
            lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
            # Primary language 0x1A covers Croatian, Serbian and Bosnian.
            return "hr" if lang_id & 0x3FF == 0x1A else "en"
        except Exception:
            pass
    try:
        name = (locale.getlocale()[0] or "").lower()
    except ValueError:
        name = ""
    return "hr" if name.startswith(("hr", "croatian", "bs", "sr")) else "en"
