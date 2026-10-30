"""English / Croatian UI texts."""

from __future__ import annotations

import locale
import os

LANGUAGES = ("en", "hr")

# Nouns with plural forms: English (one, other), Croatian (one, few, many).
_NOUNS = {
    "en": {
        "page": ("page", "pages"),
        "file": ("file", "files"),
        "group": ("group", "groups"),
    },
    "hr": {
        "page": ("stranica", "stranice", "stranica"),
        "file": ("datoteka", "datoteke", "datoteka"),
        "group": ("grupa", "grupe", "grupa"),
    },
}

_TEXTS = {
    "en": {
        # window
        "subtitle":         "automatic print router",
        "card_files":       "FILES",
        "card_print":       "PRINT SETTINGS",
        "card_plotter":     "PLOTTER — PAPER WIDTH",
        "card_log":         "ACTION LOG",
        "placeholder":      "Drag PDF files or folders here  or click  + Add",
        "btn_add":          "+ Add",
        "btn_remove":       "− Remove",
        "btn_clear":        "Clear",
        "copies_lbl":       "Copies:",
        "lbl_a4":           "Printer A4:",
        "lbl_a3":           "Printer A3:",
        "lbl_plotter":      "Plotter:",
        "btn_print_a4":     "Print A4",
        "btn_print_a3":     "Print A3",
        "btn_plot":         "Plot",
        "author":           "Author: {author}",
        "status_ready":     "Ready",
        "status_printing":  "Printing…",
        "status_analyzing": "Reading files…",
        "status_stopping":  "Stopping…",
        "dlg_add_title":    "Select PDF files",
        "dlg_pdf_files":    "PDF files",
        "dlg_close_title":  "Printing in progress",
        "dlg_close_msg":    "Pages are still being sent to the printer.\nStop after the current file and close?",
        "tip_settings":     "Printing preferences",
        "file_pending":     "{name}  (…)",
        "file_pages":       "{name}  ({pages})",
        "file_error":       "{name}  (error)",
        # plotter selector
        "roll_a3":          "A3",
        "roll_mm":          "{w} mm",
        "roll_hint":        "Select the width of the paper roll loaded in the plotter.",
        "roll_selected":    "Selected roll: {label} — {pages}",
        "roll_lengths":     "Lengths: {lengths}",
        "roll_check":       "Check that this paper width is loaded in the plotter.",
        # analysis summary
        "sum_file":         "Selected file: {name} ({pages})",
        "sum_files":        "Selected files: {n} ({pages} total)",
        "sum_line":         "  {label}: {pages}",
        "size_other":       "Other",
        "sum_size":         "    {w}×{h} mm: {pages}",
        "sum_unreadable":   "WARNING — cannot read {name}: {error}",
        "sum_encrypted":    "WARNING — {name} is password protected and cannot be read.",
        "sum_page_errors":  "WARNING — unreadable pages: {n}",
        "sum_page_error":   "  {name}, page {n}: {reason}",
        "sum_oversize":     "WARNING — wider than the widest roll ({w} mm): {pages}",
        "sum_all_ok":       "All pages are assigned for printing.",
        # print log
        "backend_sumatra":  "Printing via SumatraPDF.",
        "backend_adobe":    "Printing via Adobe Reader.",
        "backend_shell":    "Printing via the default PDF app. Installing SumatraPDF makes printing faster and more reliable.",
        "saved_missing":    "Saved printer \"{printer}\" was not found.",
        "no_files":         "No PDF files selected.",
        "no_printer":       "Select a printer for {target} first.",
        "target_plotter":   "the plotter",
        "printer_missing":  "Printer \"{printer}\" is not available.",
        "print_start":      "Printing {kind} pages on {printer} (copies: {copies})...",
        "print_sent":       "  {name} — sent: {pages}",
        "print_done":       "Finished printing {kind} pages.",
        "print_nothing":    "There are no {kind} pages in the selected files.",
        "read_error":       "Cannot read {name}: {error}",
        "send_error":       "  Error sending {name}: {error}",
        "unconfirmed":      "  Warning: could not confirm the job reached the print queue.",
        "plot_no_roll":     "Select the roll width before plotting.",
        "plot_nothing":     "No pages for roll {label}.",
        "plot_start":       "Plot on {printer}, roll {label}: {groups} by length (copies: {copies}).",
        "plot_group":       "  [{i}/{total}] length {length} mm — {pages}",
        "paper_set":        "    Paper: {w}×{h} mm",
        "paper_fail":       "    Warning: paper size not set automatically ({error}).",
        "paper_restore_fail": "Warning: plotter paper settings were not restored ({error}).",
        "plot_done":        "Plotting finished.",
        "cancelled":        "Printing stopped.",
        "unexpected":       "Unexpected error: {error}",
        "temp_cleanup":     "Temporary files not deleted: {error}",
        "settings_fail":    "Settings not saved: {error}",
    },
    "hr": {
        "subtitle":         "automatski usmjerivač tiskanja",
        "card_files":       "DATOTEKE",
        "card_print":       "POSTAVKE TISKANJA",
        "card_plotter":     "PLOTER — ŠIRINA PAPIRA",
        "card_log":         "ZAPIS RADNJI",
        "placeholder":      "Povucite PDF datoteke ili mape ovdje  ili kliknite  + Dodaj",
        "btn_add":          "+ Dodaj",
        "btn_remove":       "− Ukloni",
        "btn_clear":        "Očisti",
        "copies_lbl":       "Broj kopija:",
        "lbl_a4":           "Printer A4:",
        "lbl_a3":           "Printer A3:",
        "lbl_plotter":      "Ploter:",
        "btn_print_a4":     "Print A4",
        "btn_print_a3":     "Print A3",
        "btn_plot":         "Plot",
        "author":           "Autor: {author}",
        "status_ready":     "Spremno",
        "status_printing":  "Slanje na ispis…",
        "status_analyzing": "Čitanje datoteka…",
        "status_stopping":  "Zaustavljanje…",
        "dlg_add_title":    "Odaberite PDF datoteke",
        "dlg_pdf_files":    "PDF datoteke",
        "dlg_close_title":  "Ispis u tijeku",
        "dlg_close_msg":    "Stranice se još šalju na printer.\nZaustaviti nakon trenutne datoteke i zatvoriti?",
        "tip_settings":     "Postavke ispisa",
        "file_pending":     "{name}  (…)",
        "file_pages":       "{name}  ({pages})",
        "file_error":       "{name}  (greška)",
        "roll_a3":          "A3",
        "roll_mm":          "{w} mm",
        "roll_hint":        "Odaberite širinu role papira koja je u ploteru.",
        "roll_selected":    "Odabrana rola: {label} — {pages}",
        "roll_lengths":     "Dužine: {lengths}",
        "roll_check":       "Provjerite nalazi li se ta širina papira u ploteru.",
        "sum_file":         "Odabrana datoteka: {name} ({pages})",
        "sum_files":        "Odabrane datoteke: {n} (ukupno {pages})",
        "sum_line":         "  {label}: {pages}",
        "size_other":       "Ostalo",
        "sum_size":         "    {w}×{h} mm: {pages}",
        "sum_unreadable":   "UPOZORENJE — datoteka {name} se ne može pročitati: {error}",
        "sum_encrypted":    "UPOZORENJE — datoteka {name} je zaštićena lozinkom i ne može se pročitati.",
        "sum_page_errors":  "UPOZORENJE — nečitljive stranice: {n}",
        "sum_page_error":   "  {name}, stranica {n}: {reason}",
        "sum_oversize":     "UPOZORENJE — šire od najšire role ({w} mm): {pages}",
        "sum_all_ok":       "Sve stranice su raspoređene za ispis.",
        "backend_sumatra":  "Ispis preko SumatraPDF-a.",
        "backend_adobe":    "Ispis preko Adobe Readera.",
        "backend_shell":    "Ispis preko zadane PDF aplikacije. Instalacija SumatraPDF-a čini ispis bržim i pouzdanijim.",
        "saved_missing":    "Spremljeni printer „{printer}” nije pronađen.",
        "no_files":         "Nije odabrana nijedna PDF datoteka.",
        "no_printer":       "Prvo odaberite printer za {target}.",
        "target_plotter":   "ploter",
        "printer_missing":  "Printer „{printer}” nije dostupan.",
        "print_start":      "Ispis {kind} stranica na {printer} (kopija: {copies})...",
        "print_sent":       "  {name} — poslano: {pages}",
        "print_done":       "Završen ispis {kind} stranica.",
        "print_nothing":    "U odabranim datotekama nema {kind} stranica.",
        "read_error":       "Greška pri čitanju datoteke {name}: {error}",
        "send_error":       "  Greška pri slanju datoteke {name}: {error}",
        "unconfirmed":      "  Upozorenje: nije potvrđeno da je posao stigao u red ispisa.",
        "plot_no_roll":     "Odaberite širinu role prije plotanja.",
        "plot_nothing":     "Nema stranica za rolu {label}.",
        "plot_start":       "Plot na {printer}, rola {label}: {groups} po dužini (kopija: {copies}).",
        "plot_group":       "  [{i}/{total}] dužina {length} mm — {pages}",
        "paper_set":        "    Papir: {w}×{h} mm",
        "paper_fail":       "    Upozorenje: veličina papira nije postavljena automatski ({error}).",
        "paper_restore_fail": "Upozorenje: postavke papira na ploteru nisu vraćene ({error}).",
        "plot_done":        "Plotanje završeno.",
        "cancelled":        "Ispis zaustavljen.",
        "unexpected":       "Neočekivana greška: {error}",
        "temp_cleanup":     "Privremene datoteke nisu obrisane: {error}",
        "settings_fail":    "Postavke nisu spremljene: {error}",
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
