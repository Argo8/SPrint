"""Building blocks of the window: switches, chips, page tiles, toasts, the drawer…

Custom-painted widgets take their colours from ``theme.colors()`` when they
paint. Animations are short (150-400 ms) with an ease-out curve.
"""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import (QEasingCurve, QEvent, QObject, QPoint, QPropertyAnimation, QRect, QRectF,
                            QSize, Qt, QTimer, QVariantAnimation, Signal)
from PySide6.QtGui import QColor, QFontMetrics, QPainter, QPainterPath, QPen, QPixmap
from PySide6.QtWidgets import (QAbstractButton, QButtonGroup, QFrame, QGraphicsOpacityEffect,
                               QHBoxLayout, QLabel, QLayout, QPushButton, QSizePolicy, QToolButton,
                               QVBoxLayout, QWidget)

from . import icons
from .theme import colors, font, mix

EASE = QEasingCurve.Type.OutCubic
MAX_SIZE = 16777215


# ── Helpers ──────────────────────────────────────────────────────────────────

def repolish(widget: QWidget) -> None:
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()


def set_prop(widget: QWidget, name: str, value) -> None:
    """Set a dynamic property the style sheet matches on, and restyle if it changed."""
    if widget.property(name) != value:
        widget.setProperty(name, value)
        repolish(widget)


def label(text: str = "", *, muted: bool = False, size: float = 10, bold: bool = False,
          wrap: bool = False, parent: Optional[QWidget] = None) -> QLabel:
    lbl = QLabel(text, parent)
    lbl.setFont(font(size, bold))
    if muted:
        lbl.setProperty("muted", True)
    lbl.setWordWrap(wrap)
    return lbl


def color(value, alpha: float = 1.0) -> QColor:
    c = QColor(value)
    c.setAlphaF(alpha)
    return c


def blend(a: str, b: str, amount: float) -> QColor:
    return QColor(mix(a, b, max(0.0, min(1.0, amount))))


def set_icon(widget, name: str, role: str = "muted", size: int = 18) -> None:
    """Give a button or label a line icon in one of the theme colours."""
    tint = "#FFFFFF" if role == "white" else getattr(colors(), role)
    if isinstance(widget, QLabel):
        widget.setPixmap(icons.pixmap(name, tint, size))
    else:
        widget.setIcon(icons.icon(name, tint, size))
        widget.setIconSize(QSize(size, size))


def icon_button(name: str, tooltip: str = "", size: int = 40, icon_size: int = 18,
                role: str = "muted") -> QToolButton:
    button = QToolButton()
    button.setFixedSize(size, size)
    button.setCursor(Qt.CursorShape.PointingHandCursor)
    button.setFocusPolicy(Qt.FocusPolicy.TabFocus)
    button.setToolTip(tooltip)
    button.setAccessibleName(tooltip)
    set_icon(button, name, role, icon_size)
    return button


def button(text: str, variant: str = "ghost", icon: Optional[str] = None,
           icon_role: Optional[str] = None) -> QPushButton:
    btn = QPushButton(text)
    btn.setProperty("variant", variant)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setFocusPolicy(Qt.FocusPolicy.TabFocus)
    btn.setFont(font(10, True))
    if icon:
        role = icon_role or {"primary": "white", "plot": "white", "soft": "accent_text",
                             "link": "accent_text"}.get(variant, "text")
        set_icon(btn, icon, role, 16)
    return btn


class FocusVisible(QObject):
    """Focus rings only after keyboard navigation, like CSS ``:focus-visible``.

    The focused widget gets the property ``focusvisible`` when the focus came
    from the keyboard; the style sheet and the painted widgets draw the ring then.
    """
    KEYS = (Qt.Key.Key_Tab, Qt.Key.Key_Backtab, Qt.Key.Key_Left, Qt.Key.Key_Right,
            Qt.Key.Key_Up, Qt.Key.Key_Down)

    def __init__(self, app):
        super().__init__(app)
        self.keyboard = False
        app.installEventFilter(self)
        app.focusChanged.connect(self._changed)

    def eventFilter(self, _obj, event):
        kind = event.type()
        if kind == QEvent.Type.KeyPress and event.key() in self.KEYS:
            self.keyboard = True
        elif kind == QEvent.Type.MouseButtonPress:
            self.keyboard = False
        return False

    def _changed(self, old, new):
        if old is not None and old.property("focusvisible"):
            set_prop(old, "focusvisible", False)
        if new is not None and self.keyboard:
            set_prop(new, "focusvisible", True)


class Hover:
    """Animated 0..1 hover amount for a custom-painted widget."""

    def __init__(self, widget: QWidget, duration: int = 160):
        self.value = 0.0
        self._anim = QVariantAnimation(widget)
        self._anim.setDuration(duration)
        self._anim.setEasingCurve(EASE)
        self._anim.valueChanged.connect(lambda v: (setattr(self, "value", float(v)), widget.update()))

    def to(self, target: float) -> None:
        self._anim.stop()
        self._anim.setStartValue(self.value)
        self._anim.setEndValue(float(target))
        self._anim.start()


class HoverButton(QAbstractButton):
    """Base for painted buttons: hover animation, pointer cursor, focus ring."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.hover = Hover(self)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFocusPolicy(Qt.FocusPolicy.TabFocus)   # focus ring for keyboard use only

    def enterEvent(self, event):
        if self.isEnabled():
            self.hover.to(1)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.hover.to(0)
        super().leaveEvent(event)

    def focus_ring(self, painter: QPainter, rect: QRectF, radius: float) -> None:
        if self.hasFocus() and self.property("focusvisible"):
            painter.setPen(QPen(QColor(colors().accent), 2))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(rect.adjusted(1, 1, -1, -1), radius, radius)


def fade_in(widget: QWidget, duration: int = 320) -> None:
    """Fade a widget in (then drop the effect, which slows painting)."""
    effect = QGraphicsOpacityEffect(widget)
    effect.setOpacity(0.0)
    widget.setGraphicsEffect(effect)
    anim = QPropertyAnimation(effect, b"opacity", widget)
    anim.setDuration(duration)
    anim.setStartValue(0.0)
    anim.setEndValue(1.0)
    anim.setEasingCurve(EASE)
    anim.finished.connect(lambda: widget.setGraphicsEffect(None))
    anim.start(QPropertyAnimation.DeletionPolicy.DeleteWhenStopped)


# ── Layout ───────────────────────────────────────────────────────────────────

class FlowLayout(QLayout):
    """Lays widgets out in rows, wrapping like text."""

    def __init__(self, parent=None, spacing: int = 8):
        super().__init__(parent)
        self._items = []
        self._spacing = spacing
        self.setContentsMargins(0, 0, 0, 0)

    def addItem(self, item):
        self._items.append(item)

    def count(self):
        return len(self._items)

    def itemAt(self, index):
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index):
        return self._items.pop(index) if 0 <= index < len(self._items) else None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._do_layout(QRect(0, 0, width, 0), apply=False)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do_layout(rect, apply=True)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        margins = self.contentsMargins()
        return size + QSize(margins.left() + margins.right(), margins.top() + margins.bottom())

    def _do_layout(self, rect, apply):
        margins = self.contentsMargins()
        area = rect.adjusted(margins.left(), margins.top(), -margins.right(), -margins.bottom())
        x, y, line = area.x(), area.y(), 0
        for item in self._items:
            hint = item.sizeHint()
            if x + hint.width() > area.right() + 1 and line > 0:
                x, y, line = area.x(), y + line + self._spacing, 0
            if apply:
                item.setGeometry(QRect(QPoint(x, y), hint))
            x += hint.width() + self._spacing
            line = max(line, hint.height())
        return y + line - rect.y() + margins.bottom()


def clear_layout(layout: QLayout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        if item.widget():
            item.widget().deleteLater()
        elif item.layout():
            clear_layout(item.layout())


# ── Small pieces ─────────────────────────────────────────────────────────────

class ElidedLabel(QLabel):
    """A one-line label that ends in "…" when it does not fit."""

    def __init__(self, text: str = "", parent=None):
        super().__init__(parent)
        self._full = text
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.setText(text)

    def setText(self, text):
        self._full = text
        self.setToolTip(text)
        self._elide()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._elide()

    def _elide(self):
        metrics = QFontMetrics(self.font())
        super().setText(metrics.elidedText(self._full, Qt.TextElideMode.ElideMiddle,
                                           max(10, self.width())))

    def minimumSizeHint(self):
        return QSize(40, super().minimumSizeHint().height())


class Spinner(QWidget):
    def __init__(self, size: int = 14, parent=None):
        super().__init__(parent)
        self.setFixedSize(size, size)
        self._angle = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._step)
        self._timer.start(16)

    def _step(self):
        if self.isVisible():
            self._angle = (self._angle + 8) % 360
            self.update()

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(2, 2, -2, -2)
        p.setPen(QPen(color(c.accent, .22), 2))
        p.drawEllipse(rect)
        pen = QPen(QColor(c.accent), 2)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        p.setPen(pen)
        p.drawArc(rect, -self._angle * 16, 100 * 16)


class ProgressBar(QWidget):
    """Thin bar whose fill glides to the new value."""

    def __init__(self, role: str = "accent", height: int = 6, parent=None):
        super().__init__(parent)
        self.role = role
        self.setFixedHeight(height)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._value = 0.0
        self._anim = QVariantAnimation(self)
        self._anim.setDuration(260)
        self._anim.setEasingCurve(EASE)
        self._anim.valueChanged.connect(lambda v: (setattr(self, "_value", float(v)), self.update()))

    def set_value(self, value: float, animate: bool = True) -> None:
        value = max(0.0, min(1.0, value))
        self._anim.stop()
        if animate and self.isVisible():
            self._anim.setStartValue(self._value)
            self._anim.setEndValue(value)
            self._anim.start()
        else:
            self._value = value
            self.update()

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect())
        radius = rect.height() / 2
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(c.surface2))
        p.drawRoundedRect(rect, radius, radius)
        if self._value > 0:
            fill = QRectF(rect.left(), rect.top(), max(rect.height(), rect.width() * self._value), rect.height())
            p.setBrush(QColor(getattr(c, self.role)))
            p.drawRoundedRect(fill, radius, radius)


class Switch(HoverButton):
    """An on/off switch. ``clicked`` fires only when the user flips it."""

    def __init__(self, tooltip: str = "", parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setToolTip(tooltip)
        self.setAccessibleName(tooltip)
        self.setFixedSize(46, 28)
        self._pos = 0.0
        self._anim = QVariantAnimation(self)
        self._anim.setDuration(200)
        self._anim.setEasingCurve(EASE)
        self._anim.valueChanged.connect(lambda v: (setattr(self, "_pos", float(v)), self.update()))
        self.toggled.connect(self._slide)

    def _slide(self, on: bool) -> None:
        self._anim.stop()
        if not self.isVisible():
            self._pos = 1.0 if on else 0.0
            return
        self._anim.setStartValue(self._pos)
        self._anim.setEndValue(1.0 if on else 0.0)
        self._anim.start()

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        track = QRectF(2, 2, 42, 24)
        fill = blend(c.switch_off, c.accent, self._pos)
        if not self.isEnabled():
            fill.setAlphaF(0.5)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(fill)
        p.drawRoundedRect(track, 12, 12)
        knob = QRectF(track.left() + 3 + self._pos * 18, track.top() + 3, 18, 18)
        p.setBrush(QColor(0, 0, 0, 40))
        p.drawEllipse(knob.translated(0, 1))
        p.setBrush(QColor("#FFFFFF"))
        p.drawEllipse(knob)
        self.focus_ring(p, QRectF(self.rect()), 14)


class ToggleChip(HoverButton):
    """A toggle with a tick box, e.g. "Two-sided"."""

    def __init__(self, text: str = "", parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setText(text)
        self.setFont(font(10, True))
        self.toggled.connect(self.update)

    def sizeHint(self):
        return QSize(QFontMetrics(self.font()).horizontalAdvance(self.text()) + 54, 38)

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(1, 1, -1, -1)
        on = self.isChecked()
        h = self.hover.value
        p.setOpacity(1.0 if self.isEnabled() else 0.5)
        p.setPen(QPen(QColor(c.accent) if on else blend(c.border, c.border_strong, h), 1.2))
        p.setBrush(QColor(c.soft(c.accent, .10)) if on else QColor(c.surface))
        p.drawRoundedRect(rect, 10, 10)
        tick = QRectF(rect.left() + 12, rect.center().y() - 8, 16, 16)
        ink = QColor(c.accent_text) if on else blend(c.muted, c.text, h)
        if on:
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(c.accent))
            p.drawRoundedRect(tick, 5, 5)
            p.drawPixmap(tick.adjusted(2, 2, -2, -2).toRect(), icons.pixmap("check", "#FFFFFF", 12, 3))
        else:
            p.setPen(QPen(ink, 1.5))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRoundedRect(tick, 5, 5)
        p.setPen(ink)
        p.setFont(self.font())
        p.drawText(rect.adjusted(36, 0, -10, 0), Qt.AlignmentFlag.AlignVCenter, self.text())
        self.focus_ring(p, QRectF(self.rect()), 10)


class Segmented(QFrame):
    """Two or three choices side by side; the highlight slides to the picked one."""
    picked = Signal(int)   # by the user

    def __init__(self, labels: list[str], parent=None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        self._layout = QHBoxLayout(self)
        self._layout.setContentsMargins(3, 3, 3, 3)
        self._layout.setSpacing(2)
        self._highlight = _SegmentHighlight(self)
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._group.idClicked.connect(self._clicked)
        self._buttons: list[QPushButton] = []
        self._anim = QPropertyAnimation(self._highlight, b"geometry", self)
        self._anim.setDuration(240)
        self._anim.setEasingCurve(EASE)
        self.set_labels(labels)

    def set_labels(self, labels: list[str]) -> None:
        if len(labels) == len(self._buttons):
            for btn, text in zip(self._buttons, labels):
                btn.setText(text)
            QTimer.singleShot(0, self, lambda: self._place(False))
            return
        for btn in self._buttons:
            self._group.removeButton(btn)
            btn.deleteLater()
        self._buttons = []
        for i, text in enumerate(labels):
            btn = QPushButton(text, self)
            btn.setProperty("seg", True)
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFocusPolicy(Qt.FocusPolicy.TabFocus)
            btn.setFont(font(9.5, True))
            self._group.addButton(btn, i)
            self._layout.addWidget(btn)
            self._buttons.append(btn)
        if self._buttons:
            self._buttons[0].setChecked(True)

    def index(self) -> int:
        return self._group.checkedId()

    def set_index(self, index: int) -> None:
        if 0 <= index < len(self._buttons) and index != self.index():
            self._buttons[index].setChecked(True)
            self._place(self.isVisible())

    def _clicked(self, index: int) -> None:
        self._place(True)
        self.picked.emit(index)

    def _place(self, animate: bool) -> None:
        btn = self._group.checkedButton()
        if btn is None:
            return
        target = btn.geometry()
        self._anim.stop()
        if animate and self._highlight.geometry().width() > 0:
            self._anim.setStartValue(self._highlight.geometry())
            self._anim.setEndValue(target)
            self._anim.start()
        else:
            self._highlight.setGeometry(target)
        self._highlight.lower()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._place(False)

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self, lambda: self._place(False))

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(QPen(QColor(c.border), 1))
        p.setBrush(QColor(c.surface2))
        p.drawRoundedRect(QRectF(self.rect()).adjusted(.5, .5, -.5, -.5), 11, 11)


class _SegmentHighlight(QWidget):
    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(.5, .5, -.5, -.5)
        p.setPen(QPen(QColor(c.border), 1))
        p.setBrush(QColor(c.surface))
        p.drawRoundedRect(rect, 8, 8)


class RollChip(HoverButton):
    """One roll width: how many pages need it, or ✓ when they are printed."""

    def __init__(self, width_mm: int, parent=None):
        super().__init__(parent)
        self.width_mm = width_mm
        self.setCheckable(True)
        self.setMinimumSize(68, 60)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.sub = ""
        self.empty = True
        self.done = False
        self.toggled.connect(self.update)

    def set_info(self, sub: str, empty: bool, done: bool, accessible: str) -> None:
        self.sub, self.empty, self.done = sub, empty, done
        self.setAccessibleName(accessible)
        self.setToolTip(accessible)
        self.update()

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setOpacity(1.0 if self.isEnabled() else 0.5)
        rect = QRectF(self.rect()).adjusted(3, 3, -3, -3)
        sel = self.isChecked()
        h = self.hover.value
        if sel:
            p.setPen(QPen(color(c.plot, .28), 3))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRoundedRect(rect.adjusted(-1.5, -1.5, 1.5, 1.5), 13, 13)
        p.setPen(QPen(QColor(c.plot) if sel else blend(c.border, c.border_strong, h), 1.5 if sel else 1))
        p.setBrush(QColor(c.soft(c.plot, .12)) if sel else blend(c.surface2, c.surface, h))
        p.drawRoundedRect(rect, 12, 12)
        top = QRectF(rect.left(), rect.top() + 7, rect.width(), 24)
        p.setFont(font(12.5, True))
        p.setPen(QColor(c.muted if self.empty and not sel else c.text))
        p.drawText(top, Qt.AlignmentFlag.AlignCenter, str(self.width_mm))
        p.setFont(font(8.5, True))
        p.setPen(QColor(c.ok if self.done else c.plot_text if sel else c.muted))
        p.drawText(QRectF(rect.left(), top.bottom(), rect.width(), 18), Qt.AlignmentFlag.AlignCenter, self.sub)
        self.focus_ring(p, rect, 12)


# ── Page pictures ────────────────────────────────────────────────────────────

KIND_ROLES = {"A4": "a4", "A3": "a3", "OTHER": "plot"}


def page_box(width_mm: float, height_mm: float, box: float) -> QSize:
    scale = box / max(width_mm, height_mm)
    return QSize(max(6, round(width_mm * scale)), max(6, round(height_mm * scale)))


def paint_page(p: QPainter, rect: QRectF, kind: str, skipped: bool,
               picture: Optional[QPixmap] = None) -> None:
    """A page outline in its destination colour, with its picture when there is one."""
    c = colors()
    role = getattr(c, KIND_ROLES.get(kind, "plot"))
    if picture is not None and not picture.isNull():
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor("#FFFFFF"))
        p.drawRoundedRect(rect, 2, 2)
        p.save()
        clip = QPainterPath()
        clip.addRoundedRect(rect, 2, 2)
        p.setClipPath(clip)
        p.drawPixmap(rect, picture, QRectF(picture.rect()))
        p.restore()
        fill = Qt.BrushStyle.NoBrush
    else:
        fill = QColor(c.soft(role, .14))
    pen = QPen(QColor(c.border_strong if skipped else role), 1.5)
    if skipped:
        pen.setStyle(Qt.PenStyle.DashLine)
    p.setPen(pen)
    p.setBrush(fill if not skipped else Qt.BrushStyle.NoBrush)
    p.drawRoundedRect(rect, 2, 2)


class PageShape(QWidget):
    """The first page of a file, as a small outline (or picture) in its colour."""

    def __init__(self, box: int = 36, parent=None):
        super().__init__(parent)
        self.box = box
        self.setFixedSize(box + 6, box + 8)
        self.size_mm = (210.0, 297.0)
        self.kind = "A4"
        self.picture: Optional[QPixmap] = None
        self.pending = False

    def set_page(self, size_mm, kind: str, pending: bool = False) -> None:
        self.size_mm, self.kind, self.pending = size_mm, kind, pending
        self.update()

    def set_picture(self, picture: Optional[QPixmap]) -> None:
        self.picture = picture
        self.update()

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        size = page_box(*self.size_mm, self.box)
        rect = QRectF((self.width() - size.width()) / 2, (self.height() - size.height()) / 2,
                      size.width(), size.height())
        paint_page(p, rect, self.kind, self.pending, None if self.pending else self.picture)


class PageTile(HoverButton):
    """One page: its picture, size and where it goes. Clicking skips it or brings it back."""
    WIDTH, HEIGHT = 100, 128

    def __init__(self, size_mm, kind: str, label_text: str, parent=None):
        super().__init__(parent)
        self.setFixedSize(self.WIDTH, self.HEIGHT)
        self.size_mm, self.kind, self.label_text = size_mm, kind, label_text
        self.dest = ""
        self.skipped = False
        self.done = False
        self.picture: Optional[QPixmap] = None

    def set_state(self, dest: str, skipped: bool, done: bool, tooltip: str) -> None:
        self.dest, self.skipped, self.done = dest, skipped, done
        self.setToolTip(tooltip)
        self.setAccessibleName(tooltip)
        self.update()

    def set_picture(self, picture: QPixmap) -> None:
        self.picture = picture
        self.update()

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        h = self.hover.value
        lift = 2 * h
        rect = QRectF(self.rect()).adjusted(1, 1 + 2 - lift, -1, -1 - lift)
        p.setPen(QPen(blend(c.border, c.border_strong, h), 1))
        p.setBrush(blend(c.surface2, c.surface, h))
        p.drawRoundedRect(rect, 12, 12)
        p.setOpacity(0.55 if self.skipped else (1.0 if self.isEnabled() else 0.6))
        long_side = max(self.size_mm)
        box = 40 + 22 * max(0.0, min(1.0, (long_side - 297) / (1189 - 297)))
        size = page_box(*self.size_mm, box)
        area = QRectF(rect.left(), rect.top() + 6, rect.width(), 66)
        page = QRectF(area.center().x() - size.width() / 2, area.center().y() - size.height() / 2,
                      size.width(), size.height())
        paint_page(p, page, self.kind, self.skipped, self.picture)
        if self.done and not self.skipped:
            badge = QRectF(page.right() - 9, page.top() - 7, 16, 16)
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(c.ok))
            p.drawEllipse(badge)
            p.drawPixmap(badge.adjusted(3, 3, -3, -3).toRect(), icons.pixmap("check", "#FFFFFF", 10, 3))
        label_font = font(9.5, True)
        label_font.setStrikeOut(self.skipped)
        p.setFont(label_font)
        p.setPen(QColor(c.text))
        p.drawText(QRectF(rect.left(), rect.top() + 76, rect.width(), 18), Qt.AlignmentFlag.AlignCenter,
                   self.label_text)
        p.setFont(font(8.5, self.done))
        p.setPen(QColor(c.ok if self.done and not self.skipped else c.muted))
        p.drawText(QRectF(rect.left() + 4, rect.top() + 94, rect.width() - 8, 30),
                   Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop | Qt.TextFlag.TextWordWrap,
                   self.dest)
        p.setOpacity(1.0)
        self.focus_ring(p, rect, 12)


# ── Print button ─────────────────────────────────────────────────────────────

class PrintButton(HoverButton):
    """The big button: what will be printed, in two lines."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(66)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.title = ""
        self.sub = ""
        self.shortcut_text = "Ctrl+P"

    def set_texts(self, title: str, sub: str) -> None:
        self.title, self.sub = title, sub
        self.setAccessibleName(f"{title}. {sub}")
        self.update()

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(1, 1, -1, -1)
        enabled = self.isEnabled()
        fill = (blend(c.accent, c.accent_hover, self.hover.value) if enabled
                else blend(c.surface, c.accent, .42))
        if self.isDown():
            fill = fill.darker(108)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(fill)
        p.drawRoundedRect(rect, 14, 14)
        ink = QColor("#FFFFFF")
        box = QRectF(rect.left() + 14, rect.center().y() - 20, 40, 40)
        p.setBrush(color("#FFFFFF", .16))
        p.drawRoundedRect(box, 12, 12)
        p.drawPixmap(box.adjusted(10, 10, -10, -10).toRect(), icons.pixmap("printer", "#FFFFFF", 20))
        right = rect.right() - 14
        if enabled and self.shortcut_text and rect.width() > 330:
            p.setFont(font(8.5, True))
            width = QFontMetrics(p.font()).horizontalAdvance(self.shortcut_text) + 16
            chip = QRectF(right - width, rect.center().y() - 11, width, 22)
            p.setBrush(color("#FFFFFF", .16))
            p.drawRoundedRect(chip, 6, 6)
            p.setPen(ink)
            p.drawText(chip, Qt.AlignmentFlag.AlignCenter, self.shortcut_text)
            right = chip.left() - 8
        text = QRectF(box.right() + 14, rect.top(), right - box.right() - 14, rect.height())
        p.setPen(ink)
        p.setFont(font(12.5, True))
        metrics = QFontMetrics(p.font())
        p.drawText(QRectF(text.left(), rect.center().y() - 21, text.width(), 22),
                   Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                   metrics.elidedText(self.title, Qt.TextElideMode.ElideRight, int(text.width())))
        p.setFont(font(9))
        metrics = QFontMetrics(p.font())
        p.setPen(color("#FFFFFF", .9))
        p.drawText(QRectF(text.left(), rect.center().y() + 1, text.width(), 20),
                   Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                   metrics.elidedText(self.sub, Qt.TextElideMode.ElideRight, int(text.width())))
        if self.hasFocus() and self.property("focusvisible"):
            p.setPen(QPen(QColor(c.text), 2))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRoundedRect(rect.adjusted(-0.5, -0.5, 0.5, 0.5), 15, 15)


# ── Overlays ─────────────────────────────────────────────────────────────────

class DropOverlay(QWidget):
    """Shown over the file list while PDFs are dragged over the window."""

    def __init__(self, parent: QWidget):
        super().__init__(parent)
        self.text = ""
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.hide()

    def paintEvent(self, _event):
        c = colors()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(10, 10, -10, -10)
        pen = QPen(QColor(c.accent), 2)
        pen.setStyle(Qt.PenStyle.DashLine)
        p.setPen(pen)
        p.setBrush(color(c.soft(c.accent, .12), .92))
        p.drawRoundedRect(rect, 12, 12)
        p.setFont(font(12, True))
        metrics = QFontMetrics(p.font())
        width = metrics.horizontalAdvance(self.text) + 32
        x = rect.center().x() - width / 2
        p.drawPixmap(QRect(int(x), int(rect.center().y() - 11), 22, 22), icons.pixmap("upload", c.accent_text, 22))
        p.setPen(QColor(c.accent_text))
        p.drawText(QRectF(x + 32, rect.center().y() - 14, width, 28), Qt.AlignmentFlag.AlignVCenter, self.text)


class Toast(QFrame):
    """A short message near the bottom of the window that slides in and fades out."""

    def __init__(self, parent: QWidget):
        super().__init__(parent)
        self.setObjectName("toast")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 10, 16, 10)
        layout.setSpacing(10)
        self.icon = QLabel()
        self.icon.setFixedSize(22, 22)
        self.icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text = QLabel()
        self.text.setFont(font(10, True))
        layout.addWidget(self.icon)
        layout.addWidget(self.text)
        self._effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self._effect)
        self._fade = QPropertyAnimation(self._effect, b"opacity", self)
        self._fade.setDuration(260)
        self._slide = QPropertyAnimation(self, b"pos", self)
        self._slide.setDuration(320)
        self._slide.setEasingCurve(EASE)
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.dismiss)
        self._fade.finished.connect(self._faded)
        self.hide()

    def show_message(self, text: str, kind: str = "ok", seconds: float = 3.2) -> None:
        c = colors()
        badge = {"ok": c.accent, "error": c.danger, "info": c.accent}.get(kind, c.accent)
        name = {"ok": "check", "error": "alert", "info": "info"}.get(kind, "check")
        pix = QPixmap(44, 44)
        pix.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(badge))
        painter.drawEllipse(0, 0, 44, 44)
        painter.drawPixmap(10, 10, 24, 24, icons.pixmap(name, "#FFFFFF", 12, 3))
        painter.end()
        pix.setDevicePixelRatio(2)
        self.icon.setPixmap(pix)
        self.text.setText(text)
        self.adjustSize()
        parent = self.parentWidget()
        x = (parent.width() - self.width()) // 2
        end = QPoint(x, parent.height() - self.height() - 24)
        self._slide.stop()
        self._fade.stop()
        if not self.isVisible():
            self.move(end + QPoint(0, 16))
            self._effect.setOpacity(0.0)
        self.show()
        self.raise_()
        self._slide.setStartValue(self.pos())
        self._slide.setEndValue(end)
        self._slide.start()
        self._fade.setStartValue(self._effect.opacity())
        self._fade.setEndValue(1.0)
        self._fade.start()
        self._timer.start(int(seconds * 1000))

    def dismiss(self) -> None:
        self._fade.stop()
        self._fade.setStartValue(self._effect.opacity())
        self._fade.setEndValue(0.0)
        self._fade.start()

    def _faded(self) -> None:
        if self._effect.opacity() == 0.0:
            self.hide()


class _Scrim(QWidget):
    clicked = Signal()

    def __init__(self, parent):
        super().__init__(parent)
        self.amount = 0.0

    def paintEvent(self, _event):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(0, 0, 0, round(75 * self.amount)))

    def mousePressEvent(self, event):
        self.clicked.emit()


class Drawer(QWidget):
    """A panel that slides in from the right over a dimmed window."""
    closed = Signal()

    def __init__(self, parent: QWidget, width: int = 420):
        super().__init__(parent)
        self.panel_width = width
        self.scrim = _Scrim(self)
        self.scrim.clicked.connect(self.close_drawer)
        self.panel = QFrame(self)
        self.panel.setObjectName("drawer")
        self._open = False
        self._anim = QVariantAnimation(self)
        self._anim.setDuration(360)
        self._anim.setEasingCurve(EASE)
        self._anim.valueChanged.connect(self._apply)
        self._anim.finished.connect(self._finished)
        self._t = 0.0
        self.hide()

    @property
    def is_open(self) -> bool:
        return self._open

    def open_drawer(self) -> None:
        self._open = True
        self.setGeometry(self.parentWidget().rect())
        self.show()
        self.raise_()
        self._run(1.0)
        self.panel.setFocus()

    def close_drawer(self) -> None:
        if self._open:
            self._open = False
            self._run(0.0)
            self.closed.emit()

    def _run(self, target: float) -> None:
        self._anim.stop()
        self._anim.setStartValue(self._t)
        self._anim.setEndValue(target)
        self._anim.start()

    def _apply(self, value) -> None:
        self._t = float(value)
        self.scrim.amount = self._t
        self.scrim.update()
        width = min(self.panel_width, self.width())
        self.panel.setGeometry(round(self.width() - width * self._t), 0, width, self.height())

    def _finished(self) -> None:
        if not self._open:
            self.hide()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.scrim.setGeometry(self.rect())
        self._apply(self._t)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close_drawer()
        else:
            super().keyPressEvent(event)


class Collapsible(QWidget):
    """A container that opens and closes by animating its height."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMaximumHeight(0)
        self._open = False
        self._anim = QPropertyAnimation(self, b"maximumHeight", self)
        self._anim.setDuration(300)
        self._anim.setEasingCurve(EASE)
        self._anim.finished.connect(self._finished)

    @property
    def is_open(self) -> bool:
        return self._open

    def set_open(self, open_: bool, animate: bool = True) -> None:
        if open_ == self._open:
            return
        self._open = open_
        target = self._content_height() if open_ else 0
        self._anim.stop()
        if animate and self.isVisible():
            self._anim.setStartValue(min(self.height(), self._content_height()) if not open_ else 0)
            self._anim.setEndValue(target)
            self._anim.start()
        else:
            self.setMaximumHeight(MAX_SIZE if open_ else 0)

    def _content_height(self) -> int:
        layout = self.layout()
        if layout is None:
            return self.sizeHint().height()
        width = self.width() or (self.parentWidget().width() if self.parentWidget() else 400)
        if layout.hasHeightForWidth():
            return layout.totalHeightForWidth(width)
        return layout.totalSizeHint().height()

    def _finished(self) -> None:
        if self._open:
            self.setMaximumHeight(MAX_SIZE)


def separator() -> QFrame:
    line = QFrame()
    line.setProperty("separator", True)
    line.setFrameShape(QFrame.Shape.NoFrame)
    return line


def vbox(widget: Optional[QWidget] = None, margins=(0, 0, 0, 0), spacing: int = 8) -> QVBoxLayout:
    layout = QVBoxLayout(widget) if widget is not None else QVBoxLayout()
    layout.setContentsMargins(*margins)
    layout.setSpacing(spacing)
    return layout


def hbox(widget: Optional[QWidget] = None, margins=(0, 0, 0, 0), spacing: int = 8) -> QHBoxLayout:
    layout = QHBoxLayout(widget) if widget is not None else QHBoxLayout()
    layout.setContentsMargins(*margins)
    layout.setSpacing(spacing)
    return layout


def card(parent=None) -> QFrame:
    frame = QFrame(parent)
    frame.setProperty("card", True)
    return frame
