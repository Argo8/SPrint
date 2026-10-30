"""The SPrint window: files on the left, where they go on the right.

Each destination (A4, A3, plotter) has its own card with its printer, the
printer's settings and its own print button. The window shows the controller's
state: ``refresh`` rebuilds every text and button from a fresh plan whenever
something changes.
"""

from __future__ import annotations

import math
import os
from typing import Optional

from PySide6.QtCore import QEvent, QObject, QPoint, QRect, QRectF, QSize, Qt, QTimer, QUrl, QVariantAnimation
from PySide6.QtGui import QColor, QDesktopServices, QKeySequence, QPainter, QShortcut
from PySide6.QtWidgets import (QButtonGroup, QComboBox, QFileDialog, QFrame, QGridLayout, QLabel,
                               QMainWindow, QMessageBox, QPlainTextEdit, QScrollArea, QStackedWidget,
                               QToolButton, QVBoxLayout, QWidget)

from .. import APP_NAME, AUTHOR, __version__
from ..i18n import count, get_language, join_and, meters, set_language, t
from ..paper import A3, A4, ROLL_WIDTHS_MM
from ..plan import A3_DEST, A4_DEST, PLOT_DEST, Entry
from ..settings import data_dir
from . import icons
from .controller import Controller
from .theme import colors, font
from .widgets import (EASE, ActionButton, Collapsible, DropOverlay, Drawer, ElidedLabel, FlowLayout,
                      HoverButton, PageShape, PageTile, PreviewPopup, ProgressBar, RollChip, Segmented,
                      Spinner, Switch, Toast, ToggleChip, button, card, fade_in, hbox, icon_button, label,
                      set_icon, set_prop, vbox)

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
    """The clickable top line of a file: picture of its first page, name and what is in it."""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = hbox(self, (6, 4, 6, 4), 12)
        self.shape = PageShape(46)
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
        # the picture takes the mouse for the large preview; clicks on it still reach the row
        for child in (self.name, self.chips, self.chevron):
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
        return QSize(200, 64)

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
        self.dest_texts: list[str] = []   # where each page goes, as shown
        self.setProperty("row", True)
        layout = vbox(self, spacing=0)
        top = hbox(margins=(6, 6, 8, 6), spacing=8)
        self.header = RowHeader()
        self.header.clicked.connect(self.toggle)
        window.peek.watch(self.header.shape, self, 0)
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
        pages_layout = vbox(self.pages, (74, 0, 14, 14), 10)
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
            self.window.peek.watch(tile, self, pos)
            self.tiles_layout.addWidget(tile)
            self.tiles.append(tile)

    def page_detail(self, pos: int) -> str:
        page = self.entry.doc.pages[pos]
        dest = self.dest_texts[pos] if pos < len(self.dest_texts) else ""
        return t("page_tip", n=page.index + 1, size=page_label(page), dest=dest)

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
            picture = c.preview(entry.path, first.index, self.header.shape.picture_px())
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
        self.dest_texts = []
        for state, route in zip(entry.states, info.routes):
            if state.skip:
                dest = t("dest_skipped")
            elif route.dest == A4_DEST:
                dest = t("dest_a4")
            elif route.dest == A3_DEST:
                dest = t("dest_a3")
            else:
                dest = t("dest_roll", w=route.roll)
            self.dest_texts.append(dest + (" ✓" if state.done and not state.skip else ""))
        for pos, (page, state, tile) in enumerate(zip(doc.pages, entry.states, self.tiles)):
            tile.set_state(self.dest_texts[pos], state.skip, state.done, self.page_detail(pos))
            tile.setEnabled(not locked)
            if self.pages.is_open and tile.picture is None:
                picture = c.preview(entry.path, page.index, tile.picture_px())
                if picture is not None:
                    tile.set_picture(picture)

    def _status(self, text: str, kind: str) -> None:
        self.status.setVisible(bool(text))
        self.status.setText(("✓ " if kind == "ok" else "") + text)
        set_prop(self.status, "pill", kind)

    def picture_ready(self, page_index: int, px: int) -> None:
        if not self.entry.ready:
            return
        picture = self.window.c.previews.get((self.entry.path, page_index, px))
        if picture is None:
            return
        doc = self.entry.doc
        if doc.pages and doc.pages[0].index == page_index and self.header.shape.picture_px() == px:
            self.header.shape.set_picture(picture)
        for page, tile in zip(doc.pages, self.tiles):
            if page.index == page_index and tile.picture_px() == px:
                tile.set_picture(picture)


class PagePeek(QObject):
    """A large picture of a page when the mouse rests on one of its small pictures."""
    DELAY_MS = 550    # resting this long opens it
    LINGER_MS = 150   # moving to the next page within this keeps it open

    def __init__(self, window: "MainWindow"):
        super().__init__(window)
        self.window = window
        self.popup = PreviewPopup(window.centralWidget())
        self.current: Optional[QWidget] = None
        self.shown: Optional[tuple[str, int, int]] = None   # (path, page index, size) wanted
        self._wait = QTimer(self)
        self._wait.setSingleShot(True)
        self._wait.setInterval(self.DELAY_MS)
        self._wait.timeout.connect(self._show)
        self._linger = QTimer(self)
        self._linger.setSingleShot(True)
        self._linger.setInterval(self.LINGER_MS)
        self._linger.timeout.connect(self.hide)
        window.c.preview_ready.connect(self._picture_ready)

    def watch(self, widget: QWidget, row: FileRow, pos: int) -> None:
        """Show page ``pos`` of ``row``'s file when the mouse rests on ``widget``."""
        widget.peek_target = (row, pos)
        widget.installEventFilter(self)

    def eventFilter(self, obj, event):
        kind = event.type()
        if kind == QEvent.Type.Enter:
            self.current = obj
            self._linger.stop()
            if self.popup.isVisible():
                self._show()
            else:
                self._wait.start()
        elif kind == QEvent.Type.Leave:
            if obj is self.current:
                self.current = None
                self._wait.stop()
                self._linger.start()
        elif kind in (QEvent.Type.MouseButtonPress, QEvent.Type.Wheel):
            self.hide()
        elif kind == QEvent.Type.ToolTip:   # the large picture says more
            return True
        return False

    def hide(self) -> None:
        self._wait.stop()
        self._linger.stop()
        self.shown = None
        self.popup.hide()

    def _show(self) -> None:
        widget = self.current
        target = getattr(widget, "peek_target", None)
        if target is None or not widget.isVisible():
            return
        row, pos = target
        entry = row.entry
        if not entry.ready or pos >= len(entry.doc.pages):
            return
        page = entry.doc.pages[pos]
        root = self.window.centralWidget()
        edge = 2 * (PreviewPopup.SHADOW + PreviewPopup.PAD)
        scale = min((min(root.width() * 0.46, 900) - edge) / page.width_mm,
                    (root.height() - 24 - edge - PreviewPopup.CAPTION) / page.height_mm)
        size = QSize(max(60, round(page.width_mm * scale)), max(60, round(page.height_mm * scale)))
        px = math.ceil(max(size.width(), size.height()) * root.devicePixelRatioF())
        c = self.window.c
        picture = c.preview(entry.path, page.index, px, urgent=True)
        self.shown = (entry.path, page.index, px)
        self.popup.show_page(picture or c.any_preview(entry.path, page.index), size,
                             os.path.basename(entry.path), row.page_detail(pos), loading=picture is None)
        self._place(widget)
        self.popup.appear()

    def _place(self, widget: QWidget) -> None:
        """Beside the small picture, on the side with room, inside the window."""
        root = self.window.centralWidget()
        area = QRect(widget.mapTo(root, QPoint(0, 0)), widget.size())
        width, height = self.popup.width(), self.popup.height()
        if area.right() + width <= root.width():
            x = area.right()
        elif area.left() - width >= 0:
            x = area.left() - width
        else:
            x = max(0, min(root.width() - width, area.center().x() - width // 2))
        y = max(0, min(root.height() - height, area.center().y() - height // 2))
        self.popup.move(x, y)

    def _picture_ready(self, path: str, index: int, px: int) -> None:
        if self.shown == (path, index, px) and self.popup.isVisible():
            picture = self.window.c.previews.get(self.shown)
            if picture is not None:
                self.popup.set_picture(picture)


class DestCard(QFrame):
    """A destination: badge, title and summary, its printer and settings, its print button."""

    def __init__(self, badge_kind: str, badge_text: str = "", badge_icon: Optional[str] = None,
                 role: str = "accent"):
        super().__init__()
        self.setProperty("card", True)
        self.layout_ = vbox(self, (14, 12, 14, 12), 8)
        self.head = hbox(spacing=12)
        badge = QLabel(badge_text)
        badge.setFixedSize(38, 38)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setFont(font(10.5, True))
        badge.setProperty("badge", badge_kind)
        if badge_icon:
            set_icon(badge, badge_icon, "plot_text", 20)
        texts = vbox(spacing=1)
        self.title = label(size=10.5, bold=True)
        self.title_row = hbox(spacing=8)   # a toggle (two-sided) can go at its end
        self.title_row.addWidget(self.title, 1)
        self.summary = label(muted=True, size=9.5)
        self.summary.setMinimumWidth(1)   # a long text is cut instead of widening the column
        self.reprint = button("", "link")
        self.reprint.setFont(font(9, True))
        self.reprint.setMinimumHeight(20)
        line = hbox(spacing=8)
        line.addWidget(self.summary)
        line.addWidget(self.reprint)
        line.addStretch(1)
        texts.addLayout(self.title_row)
        texts.addLayout(line)
        self.head.addWidget(badge)
        self.head.addLayout(texts, 1)
        self.layout_.addLayout(self.head)
        self.body = QWidget()
        self.body_layout = vbox(self.body, spacing=8)
        self.layout_.addWidget(self.body)

        self.action = QStackedWidget()
        self.action.setFixedHeight(46)
        self.print_button = ActionButton(role)
        self.action.addWidget(self.print_button)
        progress = QWidget()
        row = hbox(progress, (4, 0, 0, 0), 10)
        texts = vbox(spacing=7)
        line = hbox(spacing=8)
        self.spinner = Spinner()
        self.status_text = label(size=10, bold=True)
        self.status_count = label(muted=True, size=9.5, bold=True)
        line.addWidget(self.spinner)
        line.addWidget(self.status_text, 1)
        line.addWidget(self.status_count)
        texts.addStretch(1)
        texts.addLayout(line)
        self.bar = ProgressBar(badge_kind)
        texts.addWidget(self.bar)
        texts.addStretch(1)
        row.addLayout(texts, 1)
        self.stop_button = button("", "ghost", "stop")
        row.addWidget(self.stop_button)
        self.action.addWidget(progress)
        self.sent_label = QLabel()
        self.sent_label.setFont(font(10.5, True))
        self.sent_label.setProperty("pill", "ok")
        self.sent_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.action.addWidget(self.sent_label)
        self.layout_.addWidget(self.action)

    def printer_row(self) -> tuple[QWidget, QComboBox, QToolButton]:
        """The printer list and a button for the printer's settings."""
        box = QWidget()
        layout = hbox(box, spacing=6)
        combo = QComboBox()
        combo.setFocusPolicy(Qt.FocusPolicy.TabFocus)
        combo.setFont(font(10))
        combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        combo.setMinimumContentsLength(12)
        settings = icon_button("settings", size=38, icon_size=18)
        layout.addWidget(combo, 1)
        layout.addWidget(settings)
        return box, combo, settings

    def show_status(self, status: str, handled: int, total: int) -> None:
        """The print button, or how its print is going: printing, waiting or just sent."""
        if status in ("printing", "waiting"):
            printing = status == "printing"
            if self.action.currentIndex() != 1:
                self.bar.set_value(0, animate=False)
            self.action.setCurrentIndex(1)
            self.spinner.setVisible(printing)
            self.bar.setVisible(printing)
            self.status_text.setText(t("sending") if printing else t("waiting"))
            self.status_count.setText(f"{handled} / {total}" if printing and total else "")
            if printing:
                self.bar.set_value(handled / total if total else 0)
            self.stop_button.setText(t("stop") if printing else t("cancel"))
        elif status == "sent":
            self.action.setCurrentIndex(2)
            self.sent_label.setText("✓ " + t("card_sent"))
        else:
            self.action.setCurrentIndex(0)


class MainWindow(QMainWindow):
    def __init__(self, controller: Controller):
        super().__init__()
        self.c = controller
        self.rows: dict[int, FileRow] = {}
        self._close_pending = False
        self.setWindowTitle(APP_NAME)
        self.setAcceptDrops(True)

        root = QWidget()
        root.setObjectName("root")
        root.setFocusPolicy(Qt.FocusPolicy.ClickFocus)   # clicking empty space clears focus
        self.setCentralWidget(root)
        self.peek = PagePeek(self)
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
        self.settings_button = button("", "ghost", "sliders")
        self.settings_button.clicked.connect(self.open_settings)
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
        self.list_area.verticalScrollBar().valueChanged.connect(lambda _v: self.peek.hide())
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
        layout = vbox(right, spacing=8)
        top = hbox(margins=(4, 0, 4, 0), spacing=8)
        self.where_title = label(size=8.5, bold=True)
        self.where_title.setProperty("section", True)
        top.addWidget(self.where_title, 1)
        self.copies_label = label(size=8.5, bold=True)
        self.copies_label.setProperty("section", True)
        stepper = QFrame()
        stepper.setProperty("box", True)
        steps = hbox(stepper, (2, 2, 2, 2), 0)
        self.fewer = icon_button("minus", size=30, icon_size=14, role="text")
        self.more = icon_button("plus", size=30, icon_size=14, role="text")
        self.copies_value = label(size=10.5, bold=True)
        self.copies_value.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.copies_value.setMinimumWidth(30)
        self.fewer.clicked.connect(lambda: self.c.set_copies(self.c.copies - 1))
        self.more.clicked.connect(lambda: self.c.set_copies(self.c.copies + 1))
        steps.addWidget(self.fewer)
        steps.addWidget(self.copies_value)
        steps.addWidget(self.more)
        top.addWidget(self.copies_label)
        top.addWidget(stepper)
        layout.addLayout(top)

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
        return right

    def _wire_card(self, card_: DestCard, dest: str, settings: QToolButton) -> None:
        card_.print_button.clicked.connect(lambda: self.c.start_print(dest))
        card_.stop_button.clicked.connect(lambda: self.c.stop_print(dest))
        settings.clicked.connect(lambda: self.c.open_preferences(dest))

    def _build_a4(self) -> DestCard:
        self.a4 = DestCard("a4", "A4")
        self.a4.reprint.clicked.connect(lambda: self.c.reset_done(A4_DEST))
        self.a4_duplex = ToggleChip()
        self.a4_duplex.clicked.connect(lambda on: self.c.set_duplex(A4_DEST, on))
        self.a4_duplex.setFixedHeight(30)
        self.a4.title_row.addWidget(self.a4_duplex)
        box, self.a4_printer, self.a4_settings = self.a4.printer_row()
        self.a4_printer.activated.connect(lambda _i: self._pick_printer(A4_DEST, self.a4_printer))
        self.a4.body_layout.addWidget(box)
        self._wire_card(self.a4, A4_DEST, self.a4_settings)
        return self.a4

    def _build_a3(self) -> DestCard:
        self.a3 = DestCard("a3", "A3")
        self.a3.reprint.clicked.connect(lambda: self.c.reset_done(A3_DEST))
        self.a3_target = Segmented(["", ""])
        self.a3_target.picked.connect(lambda index: self.c.set_a3_to_plotter(index == 1))
        self.a3.body_layout.addWidget(self.a3_target, 0, Qt.AlignmentFlag.AlignLeft)
        self.a3_duplex = ToggleChip()
        self.a3_duplex.clicked.connect(lambda on: self.c.set_duplex(A3_DEST, on))
        self.a3_duplex.setFixedHeight(30)
        self.a3.title_row.addWidget(self.a3_duplex)
        self.a3_printer_box, self.a3_printer, self.a3_settings = self.a3.printer_row()
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
        self._wire_card(self.a3, A3_DEST, self.a3_settings)
        return self.a3

    def _build_plot(self) -> DestCard:
        self.plot = DestCard("plot", badge_icon="roll", role="plot")
        self.plot.reprint.hide()
        box, self.plot_printer, self.plot_settings = self.plot.printer_row()
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
        info = vbox(self.roll_box, (12, 8, 12, 8), 3)
        self.roll_info = label(size=9.5, wrap=True)
        self.roll_later = label(muted=True, size=9, wrap=True)
        self.roll_oversize = label(size=9, wrap=True)
        self.roll_oversize.setProperty("pill", "warn")
        self.replot = button("", "link")
        self.replot.setMinimumHeight(20)
        self.replot.clicked.connect(lambda: self.c.reset_done(PLOT_DEST, self.c.opts.roll))
        first = hbox(spacing=8)
        first.addWidget(self.roll_info, 1)
        first.addWidget(self.replot, 0, Qt.AlignmentFlag.AlignTop)
        info.addLayout(first)
        for widget in (self.roll_later, self.roll_oversize):
            info.addWidget(widget)
        self.plot.body_layout.addWidget(self.roll_box)
        self._wire_card(self.plot, PLOT_DEST, self.plot_settings)
        return self.plot

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

        def section(title_attr: str) -> QVBoxLayout:
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

        language_box = section("sec_language_title")
        self.language_choice = Segmented(["Hrvatski", "English"])
        self.language_choice.picked.connect(lambda i: self.set_language(("hr", "en")[i]))
        language_box.addWidget(self.language_choice, 0, Qt.AlignmentFlag.AlignLeft)

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
        self.copies_label.setText(t("copies").upper())
        self.fewer.setToolTip(t("tip_fewer"))
        self.more.setToolTip(t("tip_more"))
        self.a4.title.setText(t("a4_title"))
        self.a3.title.setText(t("a3_title"))
        self.plot.title.setText(t("plot_title"))
        for combo in (self.a4_printer, self.a3_printer, self.plot_printer):
            combo.setPlaceholderText(t("choose_printer"))
            combo.setToolTip(t("printer"))
            combo.setAccessibleName(t("printer"))
        for settings in (self.a4_settings, self.a3_settings, self.plot_settings):
            settings.setToolTip(t("printer_settings"))
            settings.setAccessibleName(t("printer_settings"))
        self.a4_duplex.setText(t("duplex"))
        self.a3_duplex.setText(t("duplex"))
        self.a4.reprint.setText(t("reprint"))
        self.a3.reprint.setText(t("reprint"))
        self.a3_target.setToolTip(t("a3_target"))
        self.a3_target.set_labels([t("a3_on_printer"), t("a3_on_plotter")])
        self.a3_note_text.setText(t("a3_plotter_note"))
        self.roll_caption.setText(t("roll_label"))
        self.replot.setText(t("replot"))
        self.settings_title.setText(t("settings"))
        self.close_settings_button.setToolTip(t("close_settings"))
        self.sec_print_title.setText(t("sec_print").upper())
        self.clear_done_title.setText(t("clear_when_done"))
        self.clear_done_hint.setText(t("clear_when_done_hint"))
        self.sec_printers_title.setText(t("sec_printers").upper())
        self.detect_button.setText(t("detect"))
        self.sec_language_title.setText(t("language").upper())
        self.sec_log_title.setText(t("sec_log").upper())
        self.log_folder.setText(t("open_log_folder"))
        self.about.setText(t("about", version=__version__, author=AUTHOR))

    # ── State → widgets ──────────────────────────────────────────────────────

    def _sync_rows(self) -> None:
        self.peek.hide()
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

        self.copies_value.setText(str(c.copies))
        self.fewer.setEnabled(not busy and c.copies > 1)
        self.more.setEnabled(not busy and c.copies < 99)

        # A4
        if not plan.a4.total:
            summary = t("sum_no_a4")
        elif not plan.a4.pending:
            summary = t("sum_done", pages=count(plan.a4.total, "page"))
        elif c.duplex[A4_DEST] and c.duplex_supported[A4_DEST]:
            summary = t("sum_sheets", pages=count(plan.a4.pending, "page"),
                        sheets=count(plan.a4_sheets, "sheet"))
        else:
            summary = count(plan.a4.pending, "page")
        self.a4.summary.setText(summary)
        self._set_printer_combo(self.a4_printer, c.printers.a4, busy)
        self.a4_settings.setEnabled(bool(c.printers.a4))
        self.a4_duplex.setVisible(c.duplex_supported[A4_DEST])
        self.a4_duplex.setChecked(c.duplex[A4_DEST])
        self.a4_duplex.setEnabled(not busy)
        self.a4.reprint.setVisible(plan.a4.total > 0 and not plan.a4.pending)
        self.a4.reprint.setEnabled(not busy)
        self._update_action(self.a4, A4_DEST, plan)

        # A3
        if not plan.a3_pages:
            summary = t("sum_no_a3")
        elif opts.a3_to_plotter:
            summary = t("sum_a3_plotter", pages=count(plan.a3_pages, "page"))
        elif not plan.a3.pending:
            summary = t("sum_done", pages=count(plan.a3.total, "page"))
        else:
            summary = count(plan.a3.pending, "page")
        self.a3.summary.setText(summary)
        self.a3_target.set_index(1 if opts.a3_to_plotter else 0)
        self.a3_target.setEnabled(not busy)
        self.a3_printer_box.setVisible(not opts.a3_to_plotter)
        self.a3_note.setVisible(opts.a3_to_plotter)
        self.a3.action.setVisible(not opts.a3_to_plotter)
        self._set_printer_combo(self.a3_printer, c.printers.a3, busy)
        self.a3_settings.setEnabled(bool(c.printers.a3))
        self.a3_duplex.setVisible(c.duplex_supported[A3_DEST] and not opts.a3_to_plotter)
        self.a3_duplex.setChecked(c.duplex[A3_DEST])
        self.a3_duplex.setEnabled(not busy)
        self.a3.reprint.setVisible(not opts.a3_to_plotter and plan.a3.total > 0 and not plan.a3.pending)
        self.a3.reprint.setEnabled(not busy)
        self._update_action(self.a3, A3_DEST, plan)

        # plotter
        pending_rolls = plan.pending_rolls()
        if not plan.plot_total:
            summary = t("plot_none")
        elif not plan.plot_pending:
            summary = t("plot_all_done", pages=count(plan.plot_total, "page"))
        else:
            summary = t("plot_needs_one" if len(pending_rolls) == 1 else "plot_needs_many",
                        pages=count(plan.plot_pending, "page"), rolls=join_and(pending_rolls))
        self.plot.summary.setText(summary)
        self._set_printer_combo(self.plot_printer, c.printers.plotter, busy)
        self.plot_settings.setEnabled(bool(c.printers.plotter))
        for width, chip in self.roll_chips.items():
            info = plan.rolls.get(width)
            total, pending = (info.total, info.pending) if info else (0, 0)
            sub = (t("roll_chip_none") if not total else
                   t("roll_chip_pages", n=pending) if pending else t("roll_chip_done"))
            chip.set_info(sub, not total, bool(total) and not pending, f"{t('roll_title', w=width)}, {sub}")
            chip.setChecked(width == opts.roll)
            chip.setEnabled(not busy)
        current = plan.rolls.get(opts.roll)
        self.replot.setVisible(False)
        if not current or not current.total:
            info_text = t("roll_info_empty", w=opts.roll)
        elif not current.pending:
            info_text = t("roll_info_done", w=opts.roll)
            self.replot.setVisible(True)
        else:
            lengths = ", ".join(f"{length} mm ×{n}" for length, n in sorted(current.lengths.items(), reverse=True))
            info_text = t("roll_info", w=opts.roll, pages=count(current.pending, "page"),
                          meters=meters(current.pending_mm), lengths=lengths)
        self.roll_info.setText(info_text)
        self.replot.setEnabled(not busy)
        others = [w for w in pending_rolls if w != opts.roll]
        if current and current.pending:
            later = t("roll_later", rolls=join_and(others)) if others else ""
        else:
            nxt = plan.next_roll(opts.roll)
            later = t("roll_next", w=nxt, pages=count(plan.rolls[nxt].pending, "page")) if nxt else ""
        self.roll_later.setVisible(bool(later))
        self.roll_later.setText(later)
        self.roll_oversize.setVisible(plan.oversize > 0)
        self.roll_oversize.setText(t("roll_oversize", pages=count(plan.oversize, "page")))
        self._update_action(self.plot, PLOT_DEST, plan)

        self.clear_done.setChecked(c.settings.clear_when_done)
        self.language_choice.set_index(0 if get_language() == "hr" else 1)
        if self._close_pending and not busy:
            QTimer.singleShot(0, self, self.close)

    def _update_action(self, card_: DestCard, dest: str, plan) -> None:
        """The card's print button with what it will print, or why it cannot."""
        c = self.c
        status = c.status(dest)
        card_.show_status(status, c.handled[dest], c.totals[dest])
        if status != "idle":
            return
        plot = dest == PLOT_DEST
        pages = c.pages_for(plan, dest)
        icon = "printer"
        if c.reading:
            text = t("btn_reading")
        elif pages and not c.printer_for(dest):
            text = t("btn_pick_printer")
        elif pages:
            text = t("btn_plot" if plot else "btn_print", pages=count(pages, "page_acc"))
            if c.copies > 1:
                text += f" · {count(c.copies, 'copy')}"
        elif plot:
            current = plan.rolls.get(c.opts.roll)
            if current and current.total:
                text, icon = t("btn_roll_done", w=c.opts.roll), "check"
            elif plan.plot_pending:
                text = t("btn_roll_empty", w=c.opts.roll)
            elif plan.plot_total:
                text, icon = t("btn_all_done"), "check"
            else:
                text = t("btn_plot_empty")
        else:
            done = (plan.a4 if dest == A4_DEST else plan.a3).total
            text, icon = (t("btn_all_done"), "check") if done else (t("btn_print_empty"), "printer")
        card_.print_button.set_content(text, icon)
        card_.print_button.setEnabled(bool(pages) and not c.reading and bool(c.printer_for(dest)))

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

    def _picture_ready(self, path: str, page_index: int, px: int) -> None:
        for row in self.rows.values():
            if row.entry.path == path:
                row.picture_ready(page_index, px)

    def _append_log(self, message: str, level: str) -> None:
        mark = {"error": "✖ ", "warn": "! ", "ok": "✓ "}.get(level, "  ")
        self.log_view.appendPlainText(mark + message)

    # ── Actions ──────────────────────────────────────────────────────────────

    def ask_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, t("dlg_add_title"), "",
                                                f"{t('dlg_pdf_files')} (*.pdf *.PDF)")
        if paths:
            self.c.add_paths(paths)

    def _set_clear_when_done(self, on: bool) -> None:
        self.c.settings.clear_when_done = on
        self.c._save()

    def open_settings(self) -> None:
        self.peek.hide()
        self.drawer.open_drawer()
        self.toast.raise_()   # a message still showing stays readable above the drawer

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
        self.peek.hide()
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

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QEvent.Type.ActivationChange and not self.isActiveWindow():
            self.peek.hide()

    def showEvent(self, event):
        super().showEvent(event)
        self.centralWidget().setFocus(Qt.FocusReason.OtherFocusReason)   # no ring on a button at start

    def resizeEvent(self, event):
        super().resizeEvent(event)
        root = self.centralWidget()
        self.peek.hide()
        if self.drawer.isVisible():
            self.drawer.setGeometry(root.rect())
        if self.toast.isVisible():
            self.toast.move((root.width() - self.toast.width()) // 2, root.height() - self.toast.height() - 24)
