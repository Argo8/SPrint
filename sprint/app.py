"""SPrint main window."""

from __future__ import annotations

import logging
import logging.handlers
import math
import os
import queue
import shutil
import sys
import tempfile
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as tb
from tkinterdnd2 import DND_FILES, TkinterDnD

from . import APP_NAME, AUTHOR, __version__, i18n, jobs, printing
from .i18n import count, t
from .paper import A3, A4, OTHER, PLOT_OPTIONS, ROLL_WIDTHS_MM, DocInfo, plot_groups, summarize
from .pdfdocs import read_document
from .settings import Settings, data_dir

logger = logging.getLogger(__name__)

PAGE_BG = "#F5F7FA"
CARD_BG = "#FFFFFF"
FIELD_BG = "#F8F9FA"
LOG_BG = "#FAFAFA"
FG = "#212121"
FG2 = "#757575"
PLACEHOLDER_FG = "#9E9E9E"
PRIMARY = "#1565C0"
BORDER = "#E0E0E0"
LEVEL_COLORS = {"info": "#37474F", "ok": "#2E7D32", "warn": "#E65100", "error": "#C62828"}

FONT = "Segoe UI"
MONO = "Consolas"

MAX_COPIES = 999
NO_ROLL = "none"  # radio value when no roll is selected (ttk treats "" as "mixed")
TEMP_PREFIX = "SPrint_"
STALE_TEMP_AGE_S = 24 * 3600


def resource_path(name: str) -> str:
    """Bundled file: next to the sources, or unpacked by PyInstaller."""
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def _is_pdf(path: str) -> bool:
    return path.lower().endswith(".pdf") and os.path.isfile(path)


def expand_paths(paths) -> list[str]:
    """PDF files among ``paths``; folders contribute the PDFs directly inside them."""
    result = []
    for path in paths:
        path = os.path.normpath(path)
        if os.path.isdir(path):
            try:
                names = sorted(os.listdir(path), key=str.casefold)
            except OSError:
                continue
            result += [os.path.join(path, n) for n in names if _is_pdf(os.path.join(path, n))]
        elif _is_pdf(path):
            result.append(path)
    return result


def _auto_scrollbar(scrollbar):
    """Scroll command that shows the scrollbar only when there is something to scroll."""
    def set_(first, last):
        if float(first) <= 0.0 and float(last) >= 1.0:
            scrollbar.grid_remove()
        else:
            scrollbar.grid()
        scrollbar.set(first, last)
    return set_


class SPrintApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.settings = Settings.load()
        i18n.set_language(self.settings.language or i18n.system_language())

        self.paths: list[str] = []            # files in list order
        self.docs: dict[str, DocInfo] = {}    # page analysis, for display only
        self.plotted: set[str] = set()        # roll options plotted since the list changed
        self.pending = 0                      # files still being analysed
        self.job: threading.Thread | None = None
        self.cancel = threading.Event()
        self._closing = False   # waiting for the print job to stop before closing
        self._closed = False
        self._ui_calls: queue.Queue = queue.Queue()
        self._analysis: queue.Queue = queue.Queue()
        self._translated: list[tuple[tk.Widget, str]] = []

        _clean_stale_temp_dirs()
        self.temp_dir = tempfile.mkdtemp(prefix=TEMP_PREFIX)
        self.backend = printing.find_backend(self.settings.print_program)
        self.plot_backend = printing.DirectBackend()

        self.var_a4 = tk.StringVar()
        self.var_a3 = tk.StringVar()
        self.var_plotter = tk.StringVar()
        self.var_copies = tk.StringVar(value="1")
        self.var_roll = tk.StringVar(value=NO_ROLL)

        self._build()
        self._load_printers()
        for var in (self.var_a4, self.var_a3, self.var_plotter):
            var.trace_add("write", lambda *_: self._save_settings())
        self._apply_language()
        self._fit_to_screen()

        root.protocol("WM_DELETE_WINDOW", self._on_close)
        root.report_callback_exception = self._on_tk_error
        threading.Thread(target=self._analysis_worker, name="analysis", daemon=True).start()
        root.after(50, self._drain_ui_calls)
        program = self.settings.print_program
        if program and program != self.backend.key:
            self.log(t("program_missing", program=program), "warn")
        self.log(t(f"backend_{self.backend.key}"))
        logger.info("SPrint %s, print program: %s", __version__,
                    getattr(self.backend, "exe", self.backend.key))

    # ── Building the window ──────────────────────────────────────────────────

    def _tr(self, widget, key: str):
        """Give ``widget`` the translated text ``key``, now and after language changes."""
        self._translated.append((widget, key))
        return widget

    def _build(self) -> None:
        root = self.root
        root.title(APP_NAME)
        try:
            root.iconbitmap(resource_path("icon.ico"))
        except tk.TclError:
            pass
        root.configure(bg=PAGE_BG)

        style = tb.Style(theme="litera")
        style.configure(".", font=(FONT, 10))
        style.configure("TRadiobutton", background=CARD_BG, foreground=FG, font=(FONT, 10))

        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        outer = tk.Frame(root, bg=PAGE_BG, padx=22, pady=16)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(5, weight=1)
        self.outer = outer

        self._build_header(outer)
        self._build_files(self._card(2, "card_files"))
        self._build_printers(self._card(3, "card_print"))
        self._build_plotter(self._card(4, "card_plotter"))
        self._build_log(self._card(5, "card_log"))
        self._build_footer(outer)

        root.bind("<Control-o>", lambda e: self.ask_files())

    def _build_header(self, outer) -> None:
        hdr = tk.Frame(outer, bg=PAGE_BG)
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        hdr.columnconfigure(1, weight=1)
        tk.Label(hdr, text=APP_NAME, bg=PAGE_BG, fg=PRIMARY,
                 font=(FONT, 20, "bold")).grid(row=0, column=0, sticky="w")
        self._tr(tk.Label(hdr, bg=PAGE_BG, fg=FG2, font=(FONT, 10)),
                 "subtitle").grid(row=0, column=1, sticky="sw", padx=(8, 0), pady=(0, 6))
        langs = tk.Frame(hdr, bg=PAGE_BG)
        langs.grid(row=0, column=2, sticky="e")
        self.lang_buttons = {}
        for code in i18n.LANGUAGES:
            btn = tb.Button(langs, text=code.upper(), width=3,
                            command=lambda c=code: self.set_language(c))
            btn.pack(side=tk.LEFT, padx=(3, 0))
            self.lang_buttons[code] = btn
        tk.Frame(outer, bg=BORDER, height=1).grid(row=1, column=0, sticky="ew", pady=(2, 12))

    def _card(self, row: int, title_key: str) -> tk.Frame:
        frame = tk.Frame(self.outer, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER)
        frame.grid(row=row, column=0, sticky="nsew", pady=(0, 10))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        tk.Frame(frame, bg=PRIMARY, height=3).grid(row=0, column=0, sticky="ew")
        self._tr(tk.Label(frame, bg=CARD_BG, fg=FG2, font=(FONT, 8, "bold"), anchor="w",
                          padx=14, pady=5), title_key).grid(row=1, column=0, sticky="ew")
        tk.Frame(frame, bg=BORDER, height=1).grid(row=2, column=0, sticky="ew")
        inner = tk.Frame(frame, bg=CARD_BG, padx=14, pady=10)
        inner.grid(row=3, column=0, sticky="nsew")
        return inner

    def _build_files(self, card) -> None:
        card.columnconfigure(0, weight=1)
        box = tk.Frame(card, bg=CARD_BG)
        box.grid(row=0, column=0, sticky="nsew")
        box.columnconfigure(0, weight=1)
        self.listbox = tk.Listbox(
            box, height=5, selectmode=tk.EXTENDED, activestyle="none",
            bg=FIELD_BG, fg=FG, font=(FONT, 10), relief="flat", borderwidth=0,
            selectbackground=PRIMARY, selectforeground="white",
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=PRIMARY)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        vsb = tb.Scrollbar(box, orient=tk.VERTICAL, command=self.listbox.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = tb.Scrollbar(box, orient=tk.HORIZONTAL, command=self.listbox.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.listbox.configure(yscrollcommand=_auto_scrollbar(vsb),
                               xscrollcommand=_auto_scrollbar(hsb))

        self.placeholder = self._tr(tk.Label(box, bg=FIELD_BG, fg=PLACEHOLDER_FG,
                                             font=(FONT, 9, "italic")), "placeholder")
        for widget in (self.listbox, self.placeholder):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop)
        self.listbox.bind("<Delete>", lambda e: self.remove_selected())
        self.listbox.bind("<Control-a>", self._select_all)
        self.listbox.bind("<Double-Button-1>", self._open_file)

        buttons = tk.Frame(card, bg=CARD_BG)
        buttons.grid(row=0, column=1, sticky="n", padx=(10, 0))
        for key, style, command in (("btn_add", "primary-outline", self.ask_files),
                                    ("btn_remove", "danger-outline", self.remove_selected),
                                    ("btn_clear", "dark-outline", self.clear_files)):
            self._tr(tb.Button(buttons, width=9, bootstyle=style, command=command),
                     key).pack(fill="x", pady=(0, 6))

    def _build_printers(self, card) -> None:
        card.columnconfigure(1, weight=1)
        self._tr(tk.Label(card, bg=CARD_BG, fg=FG, font=(FONT, 10)),
                 "copies_lbl").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=(0, 8))
        check = (self.root.register(self._valid_copies), "%P")
        tb.Spinbox(card, from_=1, to=MAX_COPIES, width=6, textvariable=self.var_copies,
                   validate="key", validatecommand=check).grid(row=0, column=1, sticky="w", pady=(0, 8))
        tk.Frame(card, bg=BORDER, height=1).grid(row=1, column=0, columnspan=4, sticky="ew", pady=(0, 6))

        self.comboboxes = []
        self.action_buttons = []
        rows = (("lbl_a4", self.var_a4, "btn_print_a4", "primary", self.print_a4),
                ("lbl_a3", self.var_a3, "btn_print_a3", "primary", self.print_a3),
                ("lbl_plotter", self.var_plotter, "btn_plot", "warning", self.plot))
        for row, (label_key, var, button_key, style, command) in enumerate(rows, start=2):
            self._tr(tk.Label(card, bg=CARD_BG, fg=FG, font=(FONT, 10)),
                     label_key).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=4)
            combo = tb.Combobox(card, textvariable=var, state="readonly")
            combo.configure(postcommand=lambda c=combo: self._refresh_printer_list(c))
            combo.grid(row=row, column=1, sticky="ew", pady=4)
            self.comboboxes.append(combo)
            tb.Button(card, text="⚙", width=3, bootstyle="secondary-outline",
                      command=lambda v=var: self._printer_preferences(v.get())
                      ).grid(row=row, column=2, padx=8)
            button = self._tr(tb.Button(card, width=9, bootstyle=style, command=command), button_key)
            button.grid(row=row, column=3)
            self.action_buttons.append(button)

    def _build_plotter(self, card) -> None:
        card.columnconfigure(1, weight=1)
        radios = tk.Frame(card, bg=CARD_BG)
        radios.grid(row=0, column=0, sticky="nw", padx=(0, 16))
        self.radios = {}
        for option in PLOT_OPTIONS:
            radio = tb.Radiobutton(radios, variable=self.var_roll, value=option.key,
                                   command=self._refresh_plot_info, cursor="hand2")
            radio.pack(anchor="w", pady=2)
            self.radios[option.key] = radio
        self.plot_info = tk.Text(
            card, width=40, height=5, wrap="word", state="disabled", cursor="arrow",
            bg=FIELD_BG, fg=FG2, font=(FONT, 9), relief="flat", borderwidth=0,
            highlightthickness=1, highlightbackground=BORDER, padx=10, pady=8)
        self.plot_info.grid(row=0, column=1, sticky="nsew")

    def _build_log(self, card) -> None:
        card.columnconfigure(0, weight=1)
        card.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            card, wrap="word", width=82, height=12, state="disabled",
            bg=LOG_BG, fg=LEVEL_COLORS["info"], font=(MONO, 9), relief="flat", borderwidth=0,
            highlightthickness=1, highlightbackground=BORDER, padx=10, pady=8)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        vsb = tb.Scrollbar(card, orient=tk.VERTICAL, command=self.log_text.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=_auto_scrollbar(vsb))
        for level, color in LEVEL_COLORS.items():
            self.log_text.tag_configure(level, foreground=color)

    def _build_footer(self, outer) -> None:
        footer = tk.Frame(outer, bg=PAGE_BG)
        footer.grid(row=6, column=0, sticky="ew", pady=(2, 0))
        footer.columnconfigure(1, weight=1)
        self.progress = tb.Progressbar(footer, mode="indeterminate", length=80, bootstyle="primary")
        self.progress.grid(row=0, column=0, padx=(0, 8))
        self.progress.grid_remove()
        self.status = tk.Label(footer, bg=PAGE_BG, fg=FG2, font=(FONT, 9))
        self.status.grid(row=0, column=1, sticky="w")
        self.author = tk.Label(footer, bg=PAGE_BG, fg=FG2, font=(FONT, 9))
        self.author.grid(row=0, column=2, sticky="e")

    def _fit_to_screen(self) -> None:
        """Shrink the log if the window would not fit on the screen, then centre it."""
        root = self.root
        root.update_idletasks()
        available = root.winfo_screenheight() - 90   # taskbar and title bar
        lines = int(self.log_text.cget("height"))
        line_px = max(1, self.log_text.winfo_reqheight() // lines)
        excess = root.winfo_reqheight() - available
        if excess > 0:
            lines = max(4, lines - math.ceil(excess / line_px))
            self.log_text.configure(height=lines)
            root.update_idletasks()
        width, height = root.winfo_reqwidth(), root.winfo_reqheight()
        root.minsize(width, height - (lines - 4) * line_px)   # the log may shrink to 4 lines
        x = max(0, (root.winfo_screenwidth() - width) // 2)
        y = max(0, (available - height) // 2)
        root.geometry(f"+{x}+{y}")

    # ── Language ─────────────────────────────────────────────────────────────

    def set_language(self, code: str) -> None:
        i18n.set_language(code)
        self.settings.language = code
        self._save_settings()
        self._apply_language()

    def _apply_language(self) -> None:
        for widget, key in self._translated:
            widget.configure(text=t(key))
        current = i18n.get_language()
        for code, button in self.lang_buttons.items():
            button.configure(bootstyle="primary" if code == current else "secondary-outline")
        self.author.configure(text=f"{t('author', author=AUTHOR)}  ·  v{__version__}")
        self._refresh_files()
        self._update_status()

    # ── Files ────────────────────────────────────────────────────────────────

    def ask_files(self) -> None:
        selected = filedialog.askopenfilenames(
            parent=self.root, title=t("dlg_add_title"),
            filetypes=[(t("dlg_pdf_files"), "*.pdf"), ("*", "*.*")])
        self.add_paths(selected)

    def _on_drop(self, event):
        self.add_paths(self.root.tk.splitlist(event.data))
        return event.action

    def add_paths(self, paths) -> None:
        known = {os.path.normcase(p) for p in self.paths}
        added = False
        for path in expand_paths(paths):
            path = os.path.abspath(path)
            if os.path.normcase(path) in known:
                continue
            known.add(os.path.normcase(path))
            self.paths.append(path)
            self.pending += 1
            self._analysis.put(path)
            added = True
        if added:
            self.plotted.clear()
            self._refresh_files()
            self._update_status()

    def remove_selected(self) -> None:
        selected = set(self.listbox.curselection())
        if not selected:
            return
        removed = [p for i, p in enumerate(self.paths) if i in selected]
        self.paths = [p for i, p in enumerate(self.paths) if i not in selected]
        for path in removed:
            self.docs.pop(path, None)
        self.plotted.clear()
        self.listbox.selection_clear(0, tk.END)  # indices shift; keep nothing selected
        self._refresh_files()

    def clear_files(self) -> None:
        self.paths.clear()
        self.docs.clear()
        self.plotted.clear()
        self._refresh_files()

    def _select_all(self, _event=None):
        self.listbox.selection_set(0, tk.END)
        return "break"

    def _open_file(self, event) -> None:
        """Double-click opens the file in the default PDF viewer."""
        index = self.listbox.nearest(event.y)
        if not 0 <= index < len(self.paths) or not hasattr(os, "startfile"):
            return
        try:
            os.startfile(self.paths[index])
        except OSError as exc:
            self.log(str(exc), "error")

    def _analysis_worker(self) -> None:
        while True:
            path = self._analysis.get()
            doc = read_document(path)
            self.call_in_ui(self._on_analyzed, doc)

    def _on_analyzed(self, doc: DocInfo) -> None:
        self.pending = max(0, self.pending - 1)
        if doc.path in self.paths:
            self.docs[doc.path] = doc
            self._refresh_files()
        self._update_status()

    def _analysed_docs(self) -> list[DocInfo]:
        return [self.docs[p] for p in self.paths if p in self.docs]

    def _refresh_files(self) -> None:
        """Redraw everything that depends on the file list."""
        selection = set(self.listbox.curselection())
        self.listbox.delete(0, tk.END)
        for path in self.paths:
            name = os.path.basename(path)
            doc = self.docs.get(path)
            if doc is None:
                text = t("file_pending", name=name)
            elif not doc.ok:
                text = t("file_error", name=name)
            else:
                text = t("file_pages", name=name, pages=count(len(doc.pages), "page"))
            self.listbox.insert(tk.END, text)
        for index in selection:
            if index < len(self.paths):
                self.listbox.selection_set(index)
        if self.paths:
            self.placeholder.place_forget()
        else:
            self.placeholder.place(in_=self.listbox, relx=0.5, rely=0.5, anchor="center")
        self._refresh_summary()
        self._refresh_radios()
        self._refresh_plot_info()

    # ── Analysis summary, roll selector ──────────────────────────────────────

    def _summary_lines(self) -> list[tuple[str, str]]:
        docs = self._analysed_docs()
        if not docs:
            return []
        s = summarize(docs)
        pages = count(s.total_pages, "page")
        if len(self.paths) == 1:
            lines = [(t("sum_file", name=docs[0].name, pages=pages), "info")]
        else:
            lines = [(t("sum_files", n=len(self.paths), pages=pages), "info")]
        for kind, label in ((A4, "A4"), (A3, "A3"), (OTHER, t("size_other"))):
            if s.kinds[kind]:
                lines.append((t("sum_line", label=label, pages=count(s.kinds[kind], "page")), "info"))
        for (w, h), n in sorted(s.other_sizes.items()):
            lines.append((t("sum_size", w=w, h=h, pages=count(n, "page")), "info"))
        if s.oversize:
            lines.append((t("sum_oversize", w=max(ROLL_WIDTHS_MM),
                            pages=count(s.oversize, "page")), "warn"))
        for doc in s.unreadable:
            lines.append((t("sum_encrypted", name=doc.name) if doc.encrypted else
                          t("sum_unreadable", name=doc.name, error=doc.error), "error"))
        if s.page_errors:
            lines.append((t("sum_page_errors", n=len(s.page_errors)), "warn"))
            lines += [(t("sum_page_error", name=doc.name, n=no, reason=reason), "warn")
                      for doc, no, reason in s.page_errors]
        if s.total_pages and not (s.oversize or s.unreadable or s.page_errors):
            lines.append((t("sum_all_ok"), "ok"))
        return lines

    def _refresh_summary(self) -> None:
        """The analysis summary sits at the top of the log, above the action messages."""
        text = self.log_text
        text.configure(state="normal")
        ranges = text.tag_ranges("summary")
        if ranges:
            text.delete(ranges[0], ranges[-1])
        lines = self._summary_lines()
        if lines:
            position = "1.0"
            for line, level in lines + [("", "info")]:
                text.insert(position, line + "\n", ("summary", level))
                position = text.index(f"{position} +1 line")
        text.configure(state="disabled")

    def _refresh_radios(self) -> None:
        counts = summarize(self._analysed_docs()).plot_counts
        for option in PLOT_OPTIONS:
            text = jobs.roll_label(option.key)
            if counts[option.key]:
                text += f"  —  {count(counts[option.key], 'page')}"
            if option.key in self.plotted:
                text += "  ✓"
            self.radios[option.key].configure(text=text)

    def _refresh_plot_info(self) -> None:
        key = self.var_roll.get()
        if key not in self.radios:
            lines = [t("roll_hint")]
        else:
            groups = plot_groups(self._analysed_docs(), key)
            pages = sum(g.page_count for g in groups)
            lines = [t("roll_selected", label=jobs.roll_label(key), pages=count(pages, "page"))]
            if groups:
                lengths = ", ".join(f"{g.length_mm} mm ×{g.page_count}" for g in groups)
                lines.append(t("roll_lengths", lengths=lengths))
            lines.append(t("roll_check"))
        self.plot_info.configure(state="normal")
        self.plot_info.delete("1.0", tk.END)
        self.plot_info.insert("1.0", "\n".join(lines))
        self.plot_info.configure(state="disabled")

    # ── Printers ─────────────────────────────────────────────────────────────

    def _load_printers(self) -> None:
        try:
            self.printers = printing.list_printers()
        except Exception:
            logger.exception("Cannot list printers")
            self.printers = []
        default = printing.default_printer()
        for combo in self.comboboxes:
            combo.configure(values=self.printers)
        saved = (self.settings.printer_a4, self.settings.printer_a3, self.settings.printer_plotter)
        for var, name in zip((self.var_a4, self.var_a3, self.var_plotter), saved, strict=True):
            if name:
                var.set(name)
                if name not in self.printers:
                    self.log(t("saved_missing", printer=name), "warn")
            elif default in self.printers:
                var.set(default)
            elif self.printers:
                var.set(self.printers[0])

    def _refresh_printer_list(self, combo) -> None:
        try:
            self.printers = printing.list_printers()
        except Exception:
            logger.exception("Cannot list printers")
            return
        combo.configure(values=self.printers)

    def _printer_preferences(self, name: str) -> None:
        if not name:
            return
        try:
            printing.open_printer_preferences(name)
        except Exception as exc:
            self.log(t("unexpected", error=exc), "error")

    def _save_settings(self) -> None:
        self.settings.printer_a4 = self.var_a4.get()
        self.settings.printer_a3 = self.var_a3.get()
        self.settings.printer_plotter = self.var_plotter.get()
        try:
            self.settings.save()
        except OSError as exc:
            logger.exception("Cannot save settings")
            self.log(t("settings_fail", error=exc), "error")

    # ── Printing ─────────────────────────────────────────────────────────────

    def _valid_copies(self, proposed: str) -> bool:
        return proposed == "" or (proposed.isdigit() and int(proposed) <= MAX_COPIES)

    def _copies(self) -> int:
        try:
            copies = min(max(int(self.var_copies.get()), 1), MAX_COPIES)
        except ValueError:
            copies = 1
        self.var_copies.set(str(copies))
        return copies

    def print_a4(self) -> None:
        self._start_job(jobs.print_size, A4, self.var_a4.get())

    def print_a3(self) -> None:
        self._start_job(jobs.print_size, A3, self.var_a3.get())

    def plot(self) -> None:
        key = self.var_roll.get()
        self._start_job(jobs.plot, key if key in self.radios else "", self.var_plotter.get())

    def _start_job(self, func, target: str, printer: str) -> None:
        if self.job is not None:
            return
        ctx = jobs.JobContext(self.backend, self.plot_backend, self.temp_dir, self.log,
                              cancel=self.cancel)
        args = (ctx, target, printer, self._copies(), list(self.paths))

        def run():
            try:
                sent = func(*args)
            except Exception as exc:
                logger.exception("Print job failed")
                self.log(t("unexpected", error=exc), "error")
                sent = 0
            self.call_in_ui(self._job_finished, func, target, sent)

        self.cancel.clear()
        self.job = threading.Thread(target=run, name="print", daemon=True)
        self.job.start()
        self._update_status()

    def _job_finished(self, func, target: str, sent: int) -> None:
        self.job = None
        if func is jobs.plot and sent:
            self.plotted.add(target)
            self._refresh_radios()
        self._update_status()
        if self._closing:
            self._close()

    def _update_status(self) -> None:
        busy = self.job is not None
        for button in self.action_buttons:
            button.configure(state="disabled" if busy else "normal")
        if self._closing:
            key = "status_stopping"
        elif busy:
            key = "status_printing"
        elif self.pending:
            key = "status_analyzing"
        else:
            key = "status_ready"
        self.status.configure(text=t(key))
        if busy or self.pending:
            if not self.progress.winfo_ismapped():
                self.progress.grid()
                self.progress.start(15)
        else:
            self.progress.stop()
            self.progress.grid_remove()

    # ── Log and thread plumbing ──────────────────────────────────────────────

    def call_in_ui(self, func, *args) -> None:
        """Run ``func(*args)`` on the Tk thread (safe to call from any thread)."""
        self._ui_calls.put((func, args))

    def _drain_ui_calls(self) -> None:
        try:
            while True:
                func, args = self._ui_calls.get_nowait()
                try:
                    func(*args)
                except Exception:
                    logger.exception("UI update failed")
        except queue.Empty:
            pass
        if not self._closed:
            self.root.after(50, self._drain_ui_calls)

    def log(self, message: str, level: str = "info") -> None:
        """Append a line to the action log (safe to call from any thread)."""
        getattr(logger, {"warn": "warning", "ok": "info"}.get(level, level))(message)
        if threading.current_thread() is threading.main_thread():
            self._append_log(message, level)
        else:
            self.call_in_ui(self._append_log, message, level)

    def _append_log(self, message: str, level: str) -> None:
        text = self.log_text
        text.configure(state="normal")
        text.insert(tk.END, f"{time.strftime('%H:%M:%S')}  {message}\n", (level,))
        text.configure(state="disabled")
        text.see(tk.END)

    def _on_tk_error(self, exc_type, exc, tb_) -> None:
        logger.error("Unhandled error", exc_info=(exc_type, exc, tb_))
        self.log(t("unexpected", error=exc), "error")

    # ── Closing ──────────────────────────────────────────────────────────────

    def _on_close(self) -> None:
        if self._closing:
            return
        if self.job is not None:
            if not messagebox.askyesno(t("dlg_close_title"), t("dlg_close_msg"), parent=self.root):
                return
            # Let the job stop after the current file so it can restore the plotter settings.
            self._closing = True
            self.cancel.set()
            self._update_status()
            return
        self._close()

    def _close(self) -> None:
        self._closed = True
        self._save_settings()
        shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.root.destroy()


def _clean_stale_temp_dirs() -> None:
    """Remove temp folders of earlier runs (files can stay locked while a PDF app prints them)."""
    base = tempfile.gettempdir()
    try:
        names = os.listdir(base)
    except OSError:
        return
    now = time.time()
    for name in names:
        path = os.path.join(base, name)
        try:
            if name.startswith(TEMP_PREFIX) and now - os.path.getmtime(path) > STALE_TEMP_AGE_S:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def _setup_logging() -> None:
    handlers: list[logging.Handler] = []
    try:
        os.makedirs(data_dir(), exist_ok=True)
        handlers.append(logging.handlers.RotatingFileHandler(
            os.path.join(data_dir(), "sprint.log"), maxBytes=1_000_000, backupCount=2,
            encoding="utf-8"))
    except OSError:
        pass
    if sys.stderr is not None:  # None in the windowed .exe
        handlers.append(logging.StreamHandler())
    logging.basicConfig(level=logging.INFO, handlers=handlers,
                        format="%(asctime)s %(levelname)s %(name)s: %(message)s")


def _prepare_windows_process() -> None:
    """Sharp text on high-DPI screens and our own taskbar icon when run as a script."""
    if os.name != "nt":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"{AUTHOR}.{APP_NAME}")
    except Exception:
        pass


def main(argv: list[str] | None = None) -> None:
    """Start SPrint. PDF files or folders given as arguments are added to the list
    (so it works with "Send to" and dropping files on SPrint.exe).

    ``--smoke-test`` closes the window once the files are read and exits with
    status 1 if any could not be read; the build pipeline uses it.
    """
    argv = sys.argv[1:] if argv is None else argv
    smoke_test = "--smoke-test" in argv
    _setup_logging()
    _prepare_windows_process()
    root = TkinterDnD.Tk()
    root.withdraw()
    app = SPrintApp(root)
    app.add_paths([arg for arg in argv if not arg.startswith("--")])
    root.deiconify()

    failed = []
    if smoke_test:
        def finish():
            if app.pending:
                root.after(200, finish)
                return
            failed.extend(doc.name for doc in app.docs.values() if not doc.ok)
            try:   # the page renderer used for printing must be bundled too
                import pypdfium2 as pdfium
                for path in app.paths[:1]:
                    pdf = pdfium.PdfDocument(path)
                    pdf[0].render(scale=0.1)
                    pdf.close()
            except Exception:
                logger.exception("PDF renderer failed")
                failed.append("PDF renderer")
            logger.info("Smoke test: %d files read, failed: %s", len(app.docs), failed or "none")
            app._on_close()
        root.after(1000, finish)

    root.mainloop()
    if failed:
        sys.exit(1)
