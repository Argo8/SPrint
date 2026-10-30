"""The SPrint window: files on the left, where they go on the right, one print button.

The window shows the controller's state: ``refresh`` rebuilds every text and
switch from a fresh plan whenever something changes.
"""

from __future__ import annotations

import os
from html import escape
from typing import Optional

from PySide6.QtCore import QPropertyAnimation, QRectF, QSize, Qt, QTimer, QUrl, QVariantAnimation
from PySide6.QtGui import QColor, QDesktopServices, QKeySequence, QPainter, QShortcut
from PySide6.QtWidgets import (QApplication, QButtonGroup, QComboBox, QFileDialog, QFrame,
                               QGraphicsOpacityEffect, QGridLayout, QLabel, QMainWindow, QMessageBox,
                               QPlainTextEdit, QScrollArea, QStackedWidget, QWidget)

from .. import APP_NAME, AUTHOR, __version__
from ..i18n import count, get_language, join_and, meters, set_language, t
from ..paper import A3, A4, ROLL_WIDTHS_MM
from ..plan import A3_DEST, A4_DEST, PLOT_DEST, Entry
from ..settings import data_dir
from . import icons, theme
from .controller import Controller
from .theme import colors, font
from .widgets import (EASE, Collapsible, DropOverlay, Drawer, ElidedLabel, FlowLayout, HoverButton,
                      PageShape, PageTile, PrintButton, ProgressBar, RollChip, Segmented, Spinner,
                      Switch, Toast, ToggleChip, button, card, fade_in, hbox, icon_button,
                      label, refresh_icons, set_icon, set_prop, vbox)

RIGHT_WIDTH = 410


def page_label(page) -> str:
    if page.kind == A4:
        return "A4"
    if page.kind == A3:
        return "A3"
    return f"{round(page.width_mm)}×{round(page.height_mm)}"


class Chevron(QWidget):
    """A small arrow that turns to point up when its row opens."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(20, 20)
        self._angle = 0.0
        self._anim = QVariantAnimation(self)
        self._anim.setDuration(260)
        self._anim.setEasingCurve(EASE)
        self._anim.valueChanged.connect(lambda v: (setattr(self, "_angle", float(v)), self.update()))

    def set_open(self, open_: bool) -> None:
        self._anim.stop()
        self._anim.setStartValue(self._angle)
        self._anim.setEndValue(180.0 if open_ else 0.0)
        self._anim.start()

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        p.translate(10, 10)
        p.rotate(self._angle)
        p.drawPixmap(-9, -9, 18, 18, icons.pixmap("chevron-down", colors().muted, 18))


class FloatingIcon(QWidget):
    """The big icon of the empty file list, gently floating."""

    def __init__(self, name: str, parent=None):
        super().__init__(parent)
        self.name = name
        self.setFixedSize(64, 72)
        self._offset = 0.0
        self._anim = QVariantAnimation(self)
        self._anim.setDuration(3200)
        self._anim.setLoopCount(-1)
        self._anim.setKeyValueAt(0, 0.0)
        self._anim.setKeyValueAt(0.5, -5.0)
        self._anim.setKeyValueAt(1, 0.0)
        self._anim.valueChanged.connect(lambda v: (setattr(self, "_offset", float(v)), self.update()))
        self._anim.start()

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(2, 6 + self._offset, 60, 60)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(c.soft(c.accent, .12)))
        p.drawRoundedRect(rect, 18, 18)
        p.drawPixmap(rect.adjusted(16, 16, -16, -16).toRect(), icons.pixmap(self.name, c.accent_text, 28))


class RowHeader(HoverButton):
    """The clickable top line of a file: picture, name and what is in it."""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = hbox(self, (6, 4, 6, 4), 12)
        self.shape = PageShape(34)
        self.name = ElidedLabel()
        self.name.setFont(font(10.5, True))
        self.chips = QWidget()
        self.chips_layout = hbox(self.chips, spacing=6)
        texts = vbox(spacing=4)
        texts.addWidget(self.name)
        texts.addWidget(self.chips)
        self.chevron = Chevron()
        layout.addWidget(self.shape)
        layout.addLayout(texts, 1)
        layout.addWidget(self.chevron)
        for child in (self.shape, self.name, self.chips, self.chevron):
            child.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

    def set_chips(self, chips: list[tuple[str, str]]) -> None:
        current = [(w.text(), w.property("pill") or "") for w in self._chip_labels()]
        if current == chips:
            return
        while self.chips_layout.count():
            item = self.chips_layout.takeAt(0)
            if item.widget():
                item.widget().hide()
                item.widget().deleteLater()
        for text, kind in chips:
            chip = QLabel(text)
            chip.setFont(font(8.5, True) if kind else font(9))
            if kind:
                chip.setProperty("pill", kind)
            else:
                chip.setProperty("muted", True)
            chip.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
            self.chips_layout.addWidget(chip)
        self.chips_layout.addStretch(1)

    def _chip_labels(self) -> list[QLabel]:
        return [self.chips_layout.itemAt(i).widget() for i in range(self.chips_layout.count())
                if self.chips_layout.itemAt(i).widget() is not None]

    def sizeHint(self):
        return QSize(200, 56)

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        if self.hover.value > 0:
            fill = QColor(c.surface2)
            fill.setAlphaF(self.hover.value)
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(fill)
            p.drawRoundedRect(QRectF(self.rect()), 9, 9)
        self.focus_ring(p, QRectF(self.rect()), 9)


class FileRow(QFrame):
    """One file in the list; opens to show its pages."""

    def __init__(self, window: "MainWindow", entry: Entry, fresh: bool):
        super().__init__()
        self.window = window
        self.entry = entry
        self.setProperty("row", True)
        layout = vbox(self, spacing=0)
        top = hbox(margins=(6, 6, 8, 6), spacing=8)
        self.header = RowHeader()
        self.header.clicked.connect(self.toggle)
        self.spinner = Spinner()
        self.status = QLabel()
        self.status.setFont(font(8.5, True))
        self.remove_button = icon_button("x", size=36, icon_size=16)
        self.remove_button.clicked.connect(lambda: window.c.remove(entry))
        self.status.setFixedHeight(26)
        top.addWidget(self.header, 1)
        top.addWidget(self.spinner, 0, Qt.AlignmentFlag.AlignVCenter)
        top.addWidget(self.status, 0, Qt.AlignmentFlag.AlignVCenter)
        top.addWidget(self.remove_button, 0, Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(top)
        self.pages = Collapsible()
        pages_layout = vbox(self.pages, (62, 0, 14, 14), 10)
        self.hint = label(t("pages_hint"), muted=True, size=9)
        self.tiles_box = QWidget()
        self.tiles_layout = FlowLayout(self.tiles_box, spacing=8)
        pages_layout.addWidget(self.hint)
        pages_layout.addWidget(self.tiles_box)
        layout.addWidget(self.pages)
        self.tiles: list[PageTile] = []
        if fresh:
            fade_in(self)

    def toggle(self) -> None:
        if not self.entry.ready:
            return
        opening = not self.pages.is_open
        if opening:
            self._build_tiles()
            self.window.refresh()
        self.pages.set_open(opening)
        self.header.chevron.set_open(opening)

    def _build_tiles(self) -> None:
        if self.tiles or not self.entry.ready:
            return
        for pos, page in enumerate(self.entry.doc.pages):
            tile = PageTile((page.width_mm, page.height_mm), page.kind, page_label(page))
            tile.clicked.connect(lambda _=False, pos=pos: self.window.c.toggle_page(self.entry, pos))
            self.tiles_layout.addWidget(tile)
            self.tiles.append(tile)

    def update_from(self, info, locked: bool) -> None:
        entry, c = self.entry, self.window.c
        doc = entry.doc
        self.header.name.setText(os.path.basename(entry.path))
        reading = doc is None
        self.spinner.setVisible(reading)
        self.remove_button.setEnabled(not locked)
        self.remove_button.setToolTip(t("tip_remove", name=os.path.basename(entry.path)))
        self.remove_button.setAccessibleName(self.remove_button.toolTip())
        self.header.setToolTip(t("tip_hide_pages" if self.pages.is_open else "tip_show_pages",
                                 name=os.path.basename(entry.path)))
        if reading:
            self._status(t("status_reading"), "accent")
            self.header.set_chips([])
            self.header.shape.set_page((210, 297), A4, pending=True)
            return
        if not doc.ok:
            self._status(t("status_locked" if doc.encrypted else "status_error"), "error")
            self.header.set_chips([(doc.error or "", "")])
            self.header.shape.set_page((210, 297), A4, pending=True)
            return
        first = doc.pages[0] if doc.pages else None
        if first is not None:
            self.header.shape.set_page((first.width_mm, first.height_mm), first.kind)
            picture = c.preview(entry.path, first.index)
            if picture is not None:
                self.header.shape.set_picture(picture)
        chips = [(t("file_pages", n=doc.page_count), "")]
        kinds = info.kinds
        if kinds[A4]:
            chips.append((t("chip_a4", n=kinds[A4]), "a4"))
        if kinds[A3]:
            chips.append((t("chip_a3", n=kinds[A3]), "a3"))
        if kinds["OTHER"]:
            chips.append((t("chip_large", n=kinds["OTHER"]), "plot"))
        skipped = len(doc.pages) - info.active
        if skipped:
            chips.append((count(skipped, "skipped"), "muted"))
        if 0 < info.done < info.active:
            chips.append((t("chip_done_part", done=info.done, total=info.active), "done"))
        if any(route.oversize for route in info.routes):
            chips.append((t("chip_oversize"), "warn"))
        if doc.page_errors:
            chips.append((count(len(doc.page_errors), "bad_page"), "warn"))
        self.header.set_chips(chips)
        printed = info.active > 0 and info.done == info.active
        self._status(t("status_printed") if printed else "", "ok")
        self.hint.setText(t("pages_hint"))
        if self.tiles:
            for pos, (page, state, route, tile) in enumerate(zip(doc.pages, entry.states, info.routes, self.tiles)):
                if state.skip:
                    dest = t("dest_skipped")
                elif route.dest == A4_DEST:
                    dest = t("dest_a4")
                elif route.dest == A3_DEST:
                    dest = t("dest_a3")
                else:
                    dest = t("dest_roll", w=route.roll)
                if state.done and not state.skip:
                    dest += " ✓"
                tile.set_state(dest, state.skip, state.done,
                               t("page_tip", n=page.index + 1, size=page_label(page), dest=dest))
                tile.setEnabled(not locked)
                if self.pages.is_open and tile.picture is None:
                    picture = c.preview(entry.path, page.index)
                    if picture is not None:
                        tile.set_picture(picture)

    def _status(self, text: str, kind: str) -> None:
        self.status.setVisible(bool(text))
        self.status.setText(("✓ " if kind == "ok" else "") + text)
        set_prop(self.status, "pill", kind)

    def picture_ready(self, page_index: int) -> None:
        if not self.entry.ready:
            return
        picture = self.window.c.previews.get((self.entry.path, page_index))
        if picture is None:
            return
        doc = self.entry.doc
        if doc.pages and doc.pages[0].index == page_index:
            self.header.shape.set_picture(picture)
        for page, tile in zip(doc.pages, self.tiles):
            if page.index == page_index:
                tile.set_picture(picture)


class DestCard(QFrame):
    """Header of a destination card: badge, title, summary and on/off switch."""

    def __init__(self, badge_kind: str, badge_text: str = "", badge_icon: Optional[str] = None):
        super().__init__()
        self.setProperty("card", True)
        self.layout_ = vbox(self, (14, 14, 14, 14), 10)
        head = hbox(spacing=12)
        badge = QLabel(badge_text)
        badge.setFixedSize(44, 44)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setFont(font(11, True))
        badge.setProperty("badge", badge_kind)
        if badge_icon:
            set_icon(badge, badge_icon, "plot_text", 22)
        texts = vbox(spacing=1)
        self.title = label(size=11, bold=True)
        self.summary = label(muted=True, size=9.5)
        self.reprint = button("", "link")
        self.reprint.setFont(font(9, True))
        self.reprint.setMinimumHeight(20)
        line = hbox(spacing=8)
        line.addWidget(self.summary)
        line.addWidget(self.reprint)
        line.addStretch(1)
        texts.addWidget(self.title)
        texts.addLayout(line)
        self.switch = Switch()
        head.addWidget(badge)
        head.addLayout(texts, 1)
        head.addWidget(self.switch)
        self.layout_.addLayout(head)
        self.body = QWidget()
        self.body_layout = vbox(self.body, spacing=10)
        self.layout_.addWidget(self.body)
        self.progress_row = QWidget()
        row = hbox(self.progress_row, spacing=10)
        self.bar = ProgressBar(badge_kind)
        self.bar_text = label(muted=True, size=9, bold=True)
        self.bar_text.setMinimumWidth(64)
        self.bar_text.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        row.addWidget(self.bar, 1)
        row.addWidget(self.bar_text)
        self.progress_row.hide()
        self.layout_.addWidget(self.progress_row)

    def printer_row(self, chip: Optional[ToggleChip] = None) -> tuple[QWidget, QComboBox]:
        """The printer list, with a toggle (two-sided) beside it."""
        box = QWidget()
        layout = hbox(box, spacing=8)
        combo = QComboBox()
        combo.setFocusPolicy(Qt.FocusPolicy.TabFocus)
        combo.setFont(font(10))
        combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        combo.setMinimumContentsLength(14)
        layout.addWidget(combo, 1)
        if chip is not None:
            layout.addWidget(chip)
        return box, combo

    def set_enabled_look(self, on: bool) -> None:
        effect = self.body.graphicsEffect()
        if on and effect is not None:
            self.body.setGraphicsEffect(None)
        elif not on and effect is None:
            dim = QGraphicsOpacityEffect(self.body)
            dim.setOpacity(0.45)
            self.body.setGraphicsEffect(dim)

    def set_progress(self, show: bool, value: float, text: str) -> None:
        if show and not self.progress_row.isVisible():
            self.progress_row.show()
            self.bar.set_value(0, animate=False)
        self.progress_row.setVisible(show)
        if show:
            self.bar.set_value(value)
            self.bar_text.setText(text)


class MainWindow(QMainWindow):
    def __init__(self, controller: Controller, asset_dir: str):
        super().__init__()
        self.c = controller
        self.asset_dir = asset_dir
        self.rows: dict[int, FileRow] = {}
        self._close_pending = False
        self.setWindowTitle(APP_NAME)
        self.setAcceptDrops(True)

        root = QWidget()
        root.setObjectName("root")
        root.setFocusPolicy(Qt.FocusPolicy.ClickFocus)   # clicking empty space clears focus
        self.setCentralWidget(root)
        layout = vbox(root, spacing=0)
        layout.addWidget(self._build_header())
        body = QWidget()
        body.setObjectName("body")
        body_layout = hbox(body, (20, 18, 20, 18), 18)
        body_layout.addWidget(self._build_files(), 1)
        body_layout.addWidget(self._build_right())
        layout.addWidget(body, 1)

        self.drop_overlay = DropOverlay(self.files_card)
        self.toast = Toast(root)
        self.drawer = Drawer(root)
        self._build_drawer()

        QShortcut(QKeySequence(QKeySequence.StandardKey.Print), self, activated=self._print)
        QShortcut(QKeySequence(QKeySequence.StandardKey.Open), self, activated=self.ask_files)
        QShortcut(QKeySequence(Qt.Key.Key_Escape), self, activated=self.drawer.close_drawer)

        c = controller
        c.changed.connect(self.refresh)
        c.entries_changed.connect(self._sync_rows)
        c.preview_ready.connect(self._picture_ready)
        c.toast.connect(lambda text, kind: self.toast.show_message(text, kind))
        c.log_line.connect(self._append_log)
        c.printers_changed.connect(self._fill_printers)

        self.retranslate()
        self._fill_printers()
        self._sync_rows()
        self.refresh()

    # ── Building ─────────────────────────────────────────────────────────────

    def _build_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("header")
        layout = hbox(header, (20, 12, 16, 12), 12)
        logo = QLabel()
        logo.setFixedSize(38, 38)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo.setProperty("logo", True)
        set_icon(logo, "printer", "white", 20)
        name = label(APP_NAME, size=13.5, bold=True)
        version = QLabel(".".join(__version__.split(".")[:2]))
        version.setFont(font(8.5, True))
        version.setProperty("pill", "version")
        layout.addWidget(logo)
        layout.addWidget(name)
        layout.addWidget(version)
        layout.addStretch(1)
        self.theme_button = icon_button("moon")
        self.theme_button.clicked.connect(self._toggle_theme)
        self.settings_button = button("", "ghost", "sliders")
        self.settings_button.clicked.connect(self.open_settings)
        layout.addWidget(self.theme_button)
        layout.addWidget(self.settings_button)
        return header

    def _build_files(self) -> QFrame:
        self.files_card = card()
        layout = vbox(self.files_card, (20, 18, 20, 16), 14)
        top = hbox(spacing=10)
        titles = vbox(spacing=2)
        self.files_title = label(size=13, bold=True)
        self.files_summary = label(muted=True, size=9.5)
        titles.addWidget(self.files_title)
        titles.addWidget(self.files_summary)
        top.addLayout(titles, 1)
        self.add_button = button("", "soft", "plus")
        self.add_button.clicked.connect(self.ask_files)
        self.clear_button = button("", "ghost", "trash")
        self.clear_button.clicked.connect(self.c.clear)
        top.addWidget(self.add_button)
        top.addWidget(self.clear_button)
        layout.addLayout(top)

        self.files_stack = QStackedWidget()
        self.dropzone = QFrame()
        self.dropzone.setProperty("dropzone", True)
        drop = vbox(self.dropzone, (24, 40, 24, 40), 10)
        drop.addStretch(1)
        icon = FloatingIcon("upload")
        drop.addWidget(icon, 0, Qt.AlignmentFlag.AlignHCenter)
        self.drop_title = label(size=13, bold=True)
        self.drop_hint = label(muted=True)
        self.drop_hint.setWordWrap(True)
        for lbl in (self.drop_title, self.drop_hint):
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            drop.addWidget(lbl)
        self.choose_button = button("", "primary")
        self.choose_button.clicked.connect(self.ask_files)
        drop.addSpacing(6)
        drop.addWidget(self.choose_button, 0, Qt.AlignmentFlag.AlignHCenter)
        drop.addStretch(1)
        self.files_stack.addWidget(self.dropzone)

        self.list_area = QScrollArea()
        self.list_area.setWidgetResizable(True)
        self.list_area.setFrameShape(QFrame.Shape.NoFrame)
        holder = QWidget()
        self.list_layout = vbox(holder, (0, 0, 6, 0), 8)
        self.list_layout.addStretch(1)
        self.list_area.setWidget(holder)
        self.files_stack.addWidget(self.list_area)
        layout.addWidget(self.files_stack, 1)

        hint_row = hbox(spacing=8)
        info = QLabel()
        set_icon(info, "info", "muted", 16)
        self.list_hint = label(muted=True, size=9)
        hint_row.addWidget(info)
        hint_row.addWidget(self.list_hint, 1)
        self.hint_box = QWidget()
        self.hint_box.setLayout(hint_row)
        layout.addWidget(self.hint_box)
        return self.files_card

    def _build_right(self) -> QWidget:
        right = QWidget()
        right.setFixedWidth(RIGHT_WIDTH)
        layout = vbox(right, spacing=10)
        self.where_title = label(size=8.5, bold=True)
        self.where_title.setProperty("section", True)
        self.where_title.setContentsMargins(4, 0, 0, 0)
        layout.addWidget(self.where_title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.right_scroll = scroll
        column = QWidget()
        cards = vbox(column, (0, 0, 4, 2), 12)
        cards.addWidget(self._build_a4())
        cards.addWidget(self._build_a3())
        cards.addWidget(self._build_plot())
        cards.addStretch(1)
        scroll.setWidget(column)
        layout.addWidget(scroll, 1)
        layout.addWidget(self._build_print_panel())
        return right

    def _build_a4(self) -> DestCard:
        self.a4 = DestCard("a4", "A4")
        self.a4.switch.clicked.connect(lambda on: self.c.set_switch(A4_DEST, on))
        self.a4.reprint.clicked.connect(lambda: self.c.reset_done(A4_DEST))
        self.a4_duplex = ToggleChip()
        self.a4_duplex.clicked.connect(lambda on: self.c.set_duplex(A4_DEST, on))
        box, self.a4_printer = self.a4.printer_row(self.a4_duplex)
        self.a4_printer.activated.connect(lambda _i: self._pick_printer(A4_DEST, self.a4_printer))
        self.a4.body_layout.addWidget(box)
        return self.a4

    def _build_a3(self) -> DestCard:
        self.a3 = DestCard("a3", "A3")
        self.a3.switch.clicked.connect(lambda on: self.c.set_switch(A3_DEST, on))
        self.a3.reprint.clicked.connect(lambda: self.c.reset_done(A3_DEST))
        self.a3_target = Segmented(["", ""])
        self.a3_target.picked.connect(lambda index: self.c.set_a3_to_plotter(index == 1))
        self.a3.body_layout.addWidget(self.a3_target, 0, Qt.AlignmentFlag.AlignLeft)
        self.a3_duplex = ToggleChip()
        self.a3_duplex.clicked.connect(lambda on: self.c.set_duplex(A3_DEST, on))
        self.a3_printer_box, self.a3_printer = self.a3.printer_row(self.a3_duplex)
        self.a3_printer.activated.connect(lambda _i: self._pick_printer(A3_DEST, self.a3_printer))
        self.a3.body_layout.addWidget(self.a3_printer_box)
        self.a3_note = QFrame()
        self.a3_note.setProperty("note", True)
        note = hbox(self.a3_note, (12, 10, 12, 10), 10)
        note_icon = QLabel()
        set_icon(note_icon, "roll", "plot_text", 18)
        note_icon.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.a3_note_text = label(size=9.5, wrap=True)
        note.addWidget(note_icon)
        note.addWidget(self.a3_note_text, 1)
        self.a3.body_layout.addWidget(self.a3_note)
        return self.a3

    def _build_plot(self) -> DestCard:
        self.plot = DestCard("plot", badge_icon="roll")
        self.plot.switch.clicked.connect(lambda on: self.c.set_switch(PLOT_DEST, on))
        box, self.plot_printer = self.plot.printer_row()
        self.plot_printer.activated.connect(lambda _i: self._pick_printer(PLOT_DEST, self.plot_printer))
        self.plot.body_layout.addWidget(box)

        rolls = vbox(spacing=6)
        self.roll_caption = label(muted=True, size=9, bold=True)
        rolls.addWidget(self.roll_caption)
        grid = QGridLayout()
        grid.setSpacing(4)
        self.roll_group = QButtonGroup(self)
        self.roll_group.setExclusive(True)
        self.roll_chips: dict[int, RollChip] = {}
        for column, width in enumerate(ROLL_WIDTHS_MM):
            chip = RollChip(width)
            self.roll_group.addButton(chip, width)
            grid.addWidget(chip, 0, column)
            self.roll_chips[width] = chip
        self.roll_group.idClicked.connect(self.c.set_roll)
        rolls.addLayout(grid)
        self.plot.body_layout.addLayout(rolls)

        self.roll_box = QFrame()
        self.roll_box.setProperty("box", True)
        info = vbox(self.roll_box, (12, 10, 12, 10), 4)
        self.roll_title = label(size=10.5, bold=True)
        self.roll_info = label(muted=True, size=9.5, wrap=True)
        self.roll_later = label(muted=True, size=9, wrap=True)
        self.roll_oversize = label(size=9, wrap=True)
        self.roll_oversize.setProperty("pill", "warn")
        self.replot = button("", "link")
        self.replot.clicked.connect(lambda: self.c.reset_done(PLOT_DEST, self.c.opts.roll))
        for widget in (self.roll_title, self.roll_info, self.roll_later, self.roll_oversize):
            info.addWidget(widget)
        info.addWidget(self.replot, 0, Qt.AlignmentFlag.AlignLeft)
        self.plot.body_layout.addWidget(self.roll_box)

        self.callout = QFrame()
        self.callout.setProperty("callout", True)
        callout = hbox(self.callout, (14, 12, 12, 12), 12)
        texts = vbox(spacing=2)
        self.next_title = label(size=10.5, bold=True)
        self.next_text = label(muted=True, size=9.5, wrap=True)
        texts.addWidget(self.next_title)
        texts.addWidget(self.next_text)
        self.next_button = button("", "plot")
        self.next_button.clicked.connect(self._use_next_roll)
        callout.addLayout(texts, 1)
        callout.addWidget(self.next_button)
        self.plot.body_layout.addWidget(self.callout)
        self._next_roll: Optional[int] = None

        self.plot_done = QLabel()
        self.plot_done.setFont(font(9.5, True))
        self.plot_done.setProperty("pill", "ok")
        self.plot.body_layout.addWidget(self.plot_done, 0, Qt.AlignmentFlag.AlignLeft)
        return self.plot

    def _build_print_panel(self) -> QFrame:
        panel = card()
        layout = hbox(panel, (12, 12, 12, 12), 12)
        copies = vbox(spacing=4)
        self.copies_label = label(muted=True, size=8.5, bold=True)
        self.copies_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        stepper = QFrame()
        stepper.setProperty("box", True)
        steps = hbox(stepper, (2, 2, 2, 2), 0)
        self.fewer = icon_button("minus", size=36, icon_size=16, role="text")
        self.more = icon_button("plus", size=36, icon_size=16, role="text")
        self.copies_value = label(size=11, bold=True)
        self.copies_value.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.copies_value.setMinimumWidth(36)
        self.fewer.clicked.connect(lambda: self.c.set_copies(self.c.copies - 1))
        self.more.clicked.connect(lambda: self.c.set_copies(self.c.copies + 1))
        steps.addWidget(self.fewer)
        steps.addWidget(self.copies_value)
        steps.addWidget(self.more)
        copies.addStretch(1)
        copies.addWidget(self.copies_label)
        copies.addWidget(stepper)
        copies.addStretch(1)
        layout.addLayout(copies)

        self.print_stack = QStackedWidget()
        self.print_button = PrintButton()
        self.print_button.clicked.connect(self._print)
        self.print_stack.addWidget(self.print_button)

        progress = QWidget()
        prog = hbox(progress, (4, 0, 0, 0), 10)
        texts = vbox(spacing=8)
        line = hbox(spacing=8)
        line.addWidget(Spinner())
        self.sending_label = label(size=10.5, bold=True)
        self.sending_count = label(muted=True, size=10, bold=True)
        line.addWidget(self.sending_label, 1)
        line.addWidget(self.sending_count)
        texts.addStretch(1)
        texts.addLayout(line)
        self.total_bar = ProgressBar("accent", 8)
        texts.addWidget(self.total_bar)
        texts.addStretch(1)
        prog.addLayout(texts, 1)
        self.stop_button = button("", "ghost", "stop")
        self.stop_button.clicked.connect(self.c.stop_print)
        prog.addWidget(self.stop_button)
        self.print_stack.addWidget(progress)

        self.done_box = QLabel()
        self.done_box.setTextFormat(Qt.TextFormat.RichText)
        self.done_box.setProperty("pill", "ok")
        self.done_box.setMinimumHeight(66)
        self.done_box.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.print_stack.addWidget(self.done_box)
        self.print_stack.setFixedHeight(68)
        layout.addWidget(self.print_stack, 1)
        return panel

    def _build_drawer(self) -> None:
        panel = self.drawer.panel
        outer = vbox(panel, spacing=0)
        head = hbox(margins=(22, 18, 16, 14), spacing=12)
        self.settings_title = label(size=14, bold=True)
        close = icon_button("x")
        close.clicked.connect(self.drawer.close_drawer)
        self.close_settings_button = close
        head.addWidget(self.settings_title, 1)
        head.addWidget(close)
        outer.addLayout(head)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        content = QWidget()
        layout = vbox(content, (22, 0, 22, 22), 22)

        def section(title_attr: str) -> tuple:
            box = vbox(spacing=10)
            title = label(size=8.5, bold=True)
            title.setProperty("section", True)
            setattr(self, title_attr, title)
            box.addWidget(title)
            layout.addLayout(box)
            return box

        printing_box = section("sec_print_title")
        row = hbox(spacing=16)
        texts = vbox(spacing=2)
        self.clear_done_title = label(size=10, bold=True, wrap=True)
        self.clear_done_hint = label(muted=True, size=9, wrap=True)
        texts.addWidget(self.clear_done_title)
        texts.addWidget(self.clear_done_hint)
        self.clear_done = Switch()
        self.clear_done.clicked.connect(self._set_clear_when_done)
        row.addLayout(texts, 1)
        row.addWidget(self.clear_done)
        printing_box.addLayout(row)

        printers_box = section("sec_printers_title")
        self.detect_button = button("", "ghost", "refresh")
        self.detect_button.clicked.connect(self.c.detect_printers)
        printers_box.addWidget(self.detect_button, 0, Qt.AlignmentFlag.AlignLeft)
        self.prefs_title = label(size=10, bold=True)
        self.prefs_hint = label(muted=True, size=9, wrap=True)
        printers_box.addWidget(self.prefs_title)
        printers_box.addWidget(self.prefs_hint)
        prefs = hbox(spacing=8)
        self.pref_buttons = {}
        for dest in (A4_DEST, A3_DEST, PLOT_DEST):
            btn = button("", "ghost", "external")
            btn.clicked.connect(lambda _=False, dest=dest: self.c.open_preferences(dest))
            prefs.addWidget(btn)
            self.pref_buttons[dest] = btn
        prefs.addStretch(1)
        printers_box.addLayout(prefs)

        look_box = section("sec_look_title")
        self.theme_caption = label(muted=True, size=9, bold=True)
        self.theme_choice = Segmented(["", "", ""])
        self.theme_choice.picked.connect(lambda i: self.set_theme(("", theme.LIGHT, theme.DARK)[i]))
        self.language_caption = label(muted=True, size=9, bold=True)
        self.language_choice = Segmented(["Hrvatski", "English"])
        self.language_choice.picked.connect(lambda i: self.set_language(("hr", "en")[i]))
        for widget in (self.theme_caption, self.theme_choice, self.language_caption, self.language_choice):
            look_box.addWidget(widget, 0, Qt.AlignmentFlag.AlignLeft)

        log_box = section("sec_log_title")
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(font(8.5))
        self.log_view.setFixedHeight(170)
        log_box.addWidget(self.log_view)
        self.log_folder = button("", "link", "folder")
        self.log_folder.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(data_dir())))
        log_box.addWidget(self.log_folder, 0, Qt.AlignmentFlag.AlignLeft)
        self.about = label(muted=True, size=9)
        layout.addWidget(self.about)
        layout.addStretch(1)
        scroll.setWidget(content)
        outer.addWidget(scroll, 1)

    # ── Texts ────────────────────────────────────────────────────────────────

    def retranslate(self) -> None:
        self.theme_button.setToolTip(t("theme_to_light" if colors().dark else "theme_to_dark"))
        self.theme_button.setAccessibleName(self.theme_button.toolTip())
        self.settings_button.setText(t("settings"))
        self.files_title.setText(t("files_title"))
        self.add_button.setText(t("add_pdf"))
        self.clear_button.setText(t("clear"))
        self.drop_title.setText(t("drop_title"))
        self.drop_hint.setText(t("drop_hint"))
        self.choose_button.setText(t("choose_files"))
        self.list_hint.setText(t("list_hint"))
        self.drop_overlay.text = t("drop_release")
        self.where_title.setText(t("where_title").upper())
        self.a4.title.setText(t("a4_title"))
        self.a3.title.setText(t("a3_title"))
        self.plot.title.setText(t("plot_title"))
        self.a4.switch.setToolTip(t("tip_switch_a4"))
        self.a3.switch.setToolTip(t("tip_switch_a3"))
        self.plot.switch.setToolTip(t("tip_switch_plot"))
        for combo in (self.a4_printer, self.a3_printer, self.plot_printer):
            combo.setPlaceholderText(t("choose_printer"))
            combo.setToolTip(t("printer"))
            combo.setAccessibleName(t("printer"))
        self.a4_duplex.setText(t("duplex"))
        self.a3_duplex.setText(t("duplex"))
        self.a4.reprint.setText(t("reprint"))
        self.a3.reprint.setText(t("reprint"))
        self.plot.reprint.setText(t("replot"))
        self.a3_target.setToolTip(t("a3_target"))
        self.a3_target.set_labels([t("a3_on_printer"), t("a3_on_plotter")])
        self.a3_note_text.setText(t("a3_plotter_note"))
        self.roll_caption.setText(t("roll_label"))
        self.replot.setText(t("replot"))
        self.next_title.setText(t("next_title"))
        self.copies_label.setText(t("copies"))
        self.fewer.setToolTip(t("tip_fewer"))
        self.more.setToolTip(t("tip_more"))
        self.sending_label.setText(t("sending"))
        self.stop_button.setText(t("stop"))
        self.settings_title.setText(t("settings"))
        self.close_settings_button.setToolTip(t("close_settings"))
        self.sec_print_title.setText(t("sec_print").upper())
        self.clear_done_title.setText(t("clear_when_done"))
        self.clear_done_hint.setText(t("clear_when_done_hint"))
        self.sec_printers_title.setText(t("sec_printers").upper())
        self.detect_button.setText(t("detect"))
        self.prefs_title.setText(t("win_prefs"))
        self.prefs_hint.setText(t("win_prefs_hint"))
        for dest, key in ((A4_DEST, "A4"), (A3_DEST, "A3"), (PLOT_DEST, "plot_title")):
            self.pref_buttons[dest].setText(t(key) if key == "plot_title" else key)
        self.sec_look_title.setText(t("sec_look").upper())
        self.theme_caption.setText(t("theme"))
        self.theme_choice.set_labels([t("theme_system"), t("theme_light"), t("theme_dark")])
        self.language_caption.setText(t("language"))
        self.sec_log_title.setText(t("sec_log").upper())
        self.log_folder.setText(t("open_log_folder"))
        self.about.setText(t("about", version=__version__, author=AUTHOR))

    # ── State → widgets ──────────────────────────────────────────────────────

    def _sync_rows(self) -> None:
        entries = self.c.entries
        wanted = {id(e) for e in entries}
        for key in [k for k in self.rows if k not in wanted]:
            row = self.rows.pop(key)
            self.list_layout.removeWidget(row)
            row.deleteLater()
        fresh = bool(self.rows)
        for index, entry in enumerate(entries):
            row = self.rows.get(id(entry))
            if row is None:
                row = FileRow(self, entry, fresh)
                self.rows[id(entry)] = row
            if self.list_layout.indexOf(row) != index:
                self.list_layout.insertWidget(index, row)
        self.files_stack.setCurrentIndex(1 if entries else 0)
        self.hint_box.setVisible(bool(entries))

    def refresh(self) -> None:
        c = self.c
        plan = c.plan()
        opts = c.opts
        busy = c.busy
        entries = c.entries

        # files
        pages = sum(len(e.doc.pages) for e in entries if e.ready)
        if not entries:
            self.files_summary.setText(t("files_none"))
        else:
            key = "files_reading" if c.reading else "files_summary"
            self.files_summary.setText(t(key, files=count(len(entries), "file"), pages=count(pages, "page")))
        self.clear_button.setVisible(bool(entries))
        self.clear_button.setEnabled(not busy)
        for entry, info in zip(entries, plan.entries):
            row = self.rows.get(id(entry))
            if row is not None:
                row.update_from(info, busy)

        # A4
        self.a4.switch.setChecked(opts.a4_on)
        self.a4.switch.setEnabled(not busy)
        self.a4.set_enabled_look(opts.a4_on)
        if not plan.a4.total:
            summary = t("sum_no_a4")
        elif not plan.a4.pending:
            summary = t("sum_done", pages=count(plan.a4.total, "page"))
        elif not opts.a4_on:
            summary = t("sum_off", pages=count(plan.a4.pending, "page"))
        elif c.duplex[A4_DEST] and c.duplex_supported[A4_DEST]:
            summary = t("sum_sheets", pages=count(plan.a4.pending, "page"),
                        sheets=count(plan.a4_sheets, "sheet"))
        else:
            summary = count(plan.a4.pending, "page")
        self.a4.summary.setText(summary)
        self._set_printer_combo(self.a4_printer, c.printers.a4, busy)
        self.a4_duplex.setVisible(c.duplex_supported[A4_DEST])
        self.a4_duplex.setChecked(c.duplex[A4_DEST])
        self.a4_duplex.setEnabled(not busy)
        self.a4.reprint.setVisible(plan.a4.total > 0 and not plan.a4.pending)
        self.a4.reprint.setEnabled(not busy)

        # A3
        self.a3.switch.setChecked(opts.a3_on)
        self.a3.switch.setEnabled(not busy)
        self.a3.set_enabled_look(opts.a3_on)
        if not plan.a3_pages:
            summary = t("sum_no_a3")
        elif opts.a3_to_plotter:
            summary = t("sum_a3_plotter" if opts.a3_on else "sum_off", pages=count(plan.a3_pages, "page"))
        elif not plan.a3.pending:
            summary = t("sum_done", pages=count(plan.a3.total, "page"))
        elif not opts.a3_on:
            summary = t("sum_off", pages=count(plan.a3.pending, "page"))
        else:
            summary = count(plan.a3.pending, "page")
        self.a3.summary.setText(summary)
        self.a3_target.set_index(1 if opts.a3_to_plotter else 0)
        self.a3_target.setEnabled(not busy)
        self.a3_printer_box.setVisible(not opts.a3_to_plotter)
        self.a3_note.setVisible(opts.a3_to_plotter)
        self._set_printer_combo(self.a3_printer, c.printers.a3, busy)
        self.a3_duplex.setVisible(c.duplex_supported[A3_DEST])
        self.a3_duplex.setChecked(c.duplex[A3_DEST])
        self.a3_duplex.setEnabled(not busy)
        self.a3.reprint.setVisible(not opts.a3_to_plotter and plan.a3.total > 0 and not plan.a3.pending)
        self.a3.reprint.setEnabled(not busy)
        self.plot.reprint.hide()

        # plotter
        self.plot.switch.setChecked(opts.plot_on)
        self.plot.switch.setEnabled(not busy)
        self.plot.set_enabled_look(opts.plot_on)
        pending_rolls = plan.pending_rolls()
        if not plan.plot_total:
            summary = t("plot_none")
        elif not plan.plot_pending:
            summary = t("plot_all_done", pages=count(plan.plot_total, "page"))
        elif not opts.plot_on:
            summary = t("sum_off", pages=count(plan.plot_pending, "page"))
        else:
            summary = t("plot_needs_one" if len(pending_rolls) == 1 else "plot_needs_many",
                        pages=count(plan.plot_pending, "page"), rolls=join_and(pending_rolls))
        self.plot.summary.setText(summary)
        self._set_printer_combo(self.plot_printer, c.printers.plotter, busy)
        for width, chip in self.roll_chips.items():
            info = plan.rolls.get(width)
            total, pending = (info.total, info.pending) if info else (0, 0)
            sub = (t("roll_chip_none") if not total else
                   t("roll_chip_pages", n=pending) if pending else t("roll_chip_done"))
            chip.set_info(sub, not total, bool(total) and not pending, f"{t('roll_title', w=width)}, {sub}")
            chip.setChecked(width == opts.roll)
            chip.setEnabled(not busy)
        current = plan.rolls.get(opts.roll)
        current_pending = current.pending if current else 0
        self.roll_title.setText(t("roll_title", w=opts.roll))
        self.replot.setVisible(False)
        if not opts.plot_on:
            info_text = t("roll_info_off")
        elif not current or not current.total:
            info_text = t("roll_info_empty")
        elif not current.pending:
            info_text = t("roll_info_done")
            self.replot.setVisible(True)
        else:
            lengths = ", ".join(f"{length} mm ×{n}" for length, n in sorted(current.lengths.items(), reverse=True))
            info_text = t("roll_info", pages=count(current.pending, "page"),
                          meters=meters(current.pending_mm), lengths=lengths)
        self.roll_info.setText(info_text)
        self.replot.setEnabled(not busy)
        others = [w for w in pending_rolls if w != opts.roll]
        self.roll_later.setVisible(bool(opts.plot_on and current_pending and others))
        self.roll_later.setText(t("roll_later", rolls=join_and(others)))
        self.roll_oversize.setVisible(plan.oversize > 0)
        self.roll_oversize.setText(t("roll_oversize", pages=count(plan.oversize, "page")))
        self._next_roll = plan.next_roll(opts.roll) if opts.plot_on and not current_pending else None
        show_next = self._next_roll is not None
        if show_next and not self.callout.isVisible():
            self.callout.show()
            fade_in(self.callout)
            QTimer.singleShot(60, self, lambda: self.right_scroll.ensureWidgetVisible(self.callout, 0, 16))
        self.callout.setVisible(show_next)
        if show_next:
            self.next_text.setText(t("next_text", w=self._next_roll,
                                     pages=count(plan.rolls[self._next_roll].pending, "page")))
            self.next_button.setText(t("next_button", w=self._next_roll))
            self.next_button.setEnabled(not busy)
        self.plot_done.setVisible(bool(opts.plot_on and plan.plot_total and not plan.plot_pending))
        self.plot_done.setText("✓ " + t("plot_done_pill"))

        # progress on the cards
        for dest, card_ in ((A4_DEST, self.a4), (A3_DEST, self.a3), (PLOT_DEST, self.plot)):
            total = c.totals.get(dest, 0)
            show = c.job_state in ("printing", "done") and total > 0
            done = c.job_state == "done" or c.handled[dest] >= total
            value = 1.0 if c.job_state == "done" else (c.handled[dest] / total if total else 0)
            card_.set_progress(show, value, t("card_sent") if done else t("card_sending"))

        # print panel
        self.copies_value.setText(str(c.copies))
        self.fewer.setEnabled(not busy and c.copies > 1)
        self.more.setEnabled(not busy and c.copies < 99)
        title, sub, enabled = self._print_texts(plan, show_next)
        self.print_button.set_texts(title, sub)
        self.print_button.setEnabled(enabled)
        if c.job_state == "printing":
            self.print_stack.setCurrentIndex(1)
            total = sum(c.totals.values())
            handled = sum(c.handled.values())
            self.sending_count.setText(f"{handled} / {total}")
            self.total_bar.set_value(handled / total if total else 0)
        elif c.job_state == "done":
            self.print_stack.setCurrentIndex(2)
            title = t("card_sent") if self._next_roll else t("sent_all")
            text = f"<div style='font-size:12.5pt'>✓ {escape(title)}</div>"
            if self._next_roll:
                text += f"<div style='font-size:9.5pt; font-weight:500'>{escape(t('sent_next', w=self._next_roll))}</div>"
            self.done_box.setText(text)
        else:
            self.print_stack.setCurrentIndex(0)
            self.total_bar.set_value(0, animate=False)

        self.clear_done.setChecked(c.settings.clear_when_done)
        self.theme_choice.set_index({"": 0, theme.LIGHT: 1, theme.DARK: 2}.get(c.settings.theme, 0))
        self.language_choice.set_index(0 if get_language() == "hr" else 1)
        if self._close_pending and not busy:
            QTimer.singleShot(0, self, self.close)

    def _print_texts(self, plan, show_next: bool) -> tuple[str, str, bool]:
        c = self.c
        copies = f" · {count(c.copies, 'copy')}" if c.copies > 1 else ""
        if not c.entries:
            return t("print"), t("sub_add"), False
        if c.reading:
            return t("print"), t("sub_reading"), False
        if not plan.included:
            if not plan.pending:
                return t("all_printed"), t("sub_all_done"), False
            if show_next:
                return t("print"), t("sub_change_roll", w=self._next_roll), False
            return t("print"), t("sub_nothing"), False
        missing = c.missing_printer(plan)
        if missing:
            return t("print"), t("sub_pick_printer", target=t(f"target_{missing}")), False
        if not plan.a4.included and not plan.a3.included:
            return (t("plot_on_roll", w=c.opts.roll),
                    t("sub_plot", pages=count(plan.plot_included, "page"),
                      meters=meters(plan.plot_included_mm)) + copies, True)
        parts = []
        if plan.a4.included:
            parts.append(f"A4 {plan.a4.included}")
        if plan.a3.included:
            parts.append(f"A3 {plan.a3.included}")
        if plan.plot_included:
            parts.append(t("sub_part_roll", w=c.opts.roll, n=plan.plot_included))
        return t("print"), " · ".join(parts) + copies, True

    def _fill_printers(self) -> None:
        for combo, current in ((self.a4_printer, self.c.printers.a4), (self.a3_printer, self.c.printers.a3),
                               (self.plot_printer, self.c.printers.plotter)):
            names = list(self.c.installed)
            for name in (self.c.printers.a4, self.c.printers.a3, self.c.printers.plotter):
                if name and name not in names:
                    names.append(name)
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(names)
            combo.setCurrentIndex(names.index(current) if current in names else -1)
            combo.blockSignals(False)

    def _set_printer_combo(self, combo: QComboBox, name: str, busy: bool) -> None:
        combo.setEnabled(not busy)
        if combo.currentText() != name:
            combo.blockSignals(True)
            index = combo.findText(name)
            if index < 0 and name:
                combo.addItem(name)
                index = combo.count() - 1
            combo.setCurrentIndex(index)
            combo.blockSignals(False)

    def _pick_printer(self, dest: str, combo: QComboBox) -> None:
        self.c.set_printer(dest, combo.currentText())

    def _picture_ready(self, path: str, page_index: int) -> None:
        for row in self.rows.values():
            if row.entry.path == path:
                row.picture_ready(page_index)

    def _append_log(self, message: str, level: str) -> None:
        mark = {"error": "✖ ", "warn": "! ", "ok": "✓ "}.get(level, "  ")
        self.log_view.appendPlainText(mark + message)

    # ── Actions ──────────────────────────────────────────────────────────────

    def ask_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, t("dlg_add_title"), "",
                                                f"{t('dlg_pdf_files')} (*.pdf *.PDF)")
        if paths:
            self.c.add_paths(paths)

    def _print(self) -> None:
        if not self.drawer.is_open and self.print_button.isEnabled():
            self.c.start_print()

    def _use_next_roll(self) -> None:
        if self._next_roll is not None:
            self.c.set_roll(self._next_roll)

    def _set_clear_when_done(self, on: bool) -> None:
        self.c.settings.clear_when_done = on
        self.c._save()

    def open_settings(self) -> None:
        self.drawer.open_drawer()

    def _toggle_theme(self) -> None:
        self.set_theme(theme.LIGHT if colors().dark else theme.DARK)

    def set_theme(self, mode: str) -> None:
        """Switch theme with a short cross-fade."""
        self.c.settings.theme = mode
        self.c._save()
        root = self.centralWidget()
        snapshot = QLabel(root)
        snapshot.setPixmap(root.grab())
        snapshot.setGeometry(root.rect())
        snapshot.show()
        snapshot.raise_()
        self.apply_theme()
        effect = QGraphicsOpacityEffect(snapshot)
        snapshot.setGraphicsEffect(effect)
        anim = QPropertyAnimation(effect, b"opacity", snapshot)
        anim.setDuration(280)
        anim.setStartValue(1.0)
        anim.setEndValue(0.0)
        anim.finished.connect(snapshot.deleteLater)
        anim.start()

    def apply_theme(self) -> None:
        theme.apply(QApplication.instance(), self.c.settings.theme, self.asset_dir)
        set_icon(self.theme_button, "sun" if colors().dark else "moon")
        refresh_icons(self)
        self.retranslate()
        self.refresh()
        for widget in self.findChildren(QWidget):
            widget.update()

    def set_language(self, lang: str) -> None:
        set_language(lang)
        self.c.settings.language = lang
        self.c._save()
        self.retranslate()
        self.refresh()

    def bring_to_front(self) -> None:
        self.setWindowState((self.windowState() & ~Qt.WindowState.WindowMinimized) | Qt.WindowState.WindowActive)
        self.show()
        self.raise_()
        self.activateWindow()

    # ── Drag and drop, closing ───────────────────────────────────────────────

    def _drop_paths(self, event) -> list[str]:
        mime = event.mimeData()
        if not mime.hasUrls():
            return []
        return [url.toLocalFile() for url in mime.urls() if url.isLocalFile()]

    def dragEnterEvent(self, event):
        if self._drop_paths(event):
            event.acceptProposedAction()
            self._show_drop(True)

    def dragMoveEvent(self, event):
        if self._drop_paths(event):
            event.acceptProposedAction()

    def dragLeaveEvent(self, event):
        self._show_drop(False)

    def dropEvent(self, event):
        paths = self._drop_paths(event)
        self._show_drop(False)
        if paths:
            event.acceptProposedAction()
            self.c.add_paths(paths)

    def _show_drop(self, on: bool) -> None:
        set_prop(self.files_card, "drag", on)
        set_prop(self.dropzone, "drag", on)
        if on and self.c.entries:
            self.drop_overlay.setGeometry(self.files_card.rect())
            self.drop_overlay.show()
            self.drop_overlay.raise_()
            fade_in(self.drop_overlay, 180)
        else:
            self.drop_overlay.hide()

    def closeEvent(self, event):
        if self.c.busy:
            box = QMessageBox(QMessageBox.Icon.Question, t("dlg_close_title"), t("dlg_close_msg"), parent=self)
            stop = box.addButton(t("dlg_stop_close"), QMessageBox.ButtonRole.AcceptRole)
            box.addButton(t("dlg_keep_printing"), QMessageBox.ButtonRole.RejectRole)
            box.exec()
            if box.clickedButton() is stop:
                self._close_pending = True
                self.c.stop_print()
            event.ignore()
            return
        self.c._save()
        event.accept()

    def showEvent(self, event):
        super().showEvent(event)
        self.centralWidget().setFocus(Qt.FocusReason.OtherFocusReason)   # no ring on a button at start

    def resizeEvent(self, event):
        super().resizeEvent(event)
        root = self.centralWidget()
        if self.drawer.isVisible():
            self.drawer.setGeometry(root.rect())
        if self.toast.isVisible():
            self.toast.move((root.width() - self.toast.width()) // 2, root.height() - self.toast.height() - 24)
