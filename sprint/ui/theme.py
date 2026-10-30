"""Colours, fonts and the style sheet, for the light and the dark theme.

Custom-painted widgets read ``colors()`` when they paint, so switching the
theme only needs a new style sheet and a repaint.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QGuiApplication

from . import icons

LIGHT, DARK = "light", "dark"


@dataclass(frozen=True)
class Colors:
    bg: str
    surface: str
    surface2: str
    border: str
    border_strong: str
    text: str
    muted: str
    accent: str
    accent_hover: str
    accent_text: str
    a4: str
    a4_text: str
    a3: str
    a3_text: str
    plot: str
    plot_text: str
    plot_strong: str
    ok: str
    danger: str
    switch_off: str
    toast_bg: str
    toast_text: str
    dark: bool

    def soft(self, color: str, amount: float = 0.10) -> str:
        """``color`` faintly over the card background (pills, highlights)."""
        return mix(self.surface, color, amount)


_LIGHT = Colors(
    bg="#F2F4F7", surface="#FFFFFF", surface2="#F6F7F9", border="#E2E6EC", border_strong="#C7CED8",
    text="#151922", muted="#586172", accent="#2563EB", accent_hover="#1D4FD7", accent_text="#1D4ED8",
    a4="#2563EB", a4_text="#1D4ED8", a3="#0D9488", a3_text="#0F766E",
    plot="#EA580C", plot_text="#C2410C", plot_strong="#C2410C", ok="#166534", danger="#B42318",
    switch_off="#C7CED8", toast_bg="#151922", toast_text="#F7F8FA", dark=False)

_DARK = Colors(
    bg="#0D1015", surface="#151920", surface2="#1B2029", border="#262D38", border_strong="#3A4452",
    text="#E8EBF1", muted="#A2ABB9", accent="#3366F0", accent_hover="#4675F3", accent_text="#93AFFF",
    a4="#6E97FF", a4_text="#A8C0FF", a3="#2CC7B3", a3_text="#6EE7D6",
    plot="#FB923C", plot_text="#FDBA74", plot_strong="#C2410C", ok="#4ADE80", danger="#F97066",
    switch_off="#3A4452", toast_bg="#EEF1F5", toast_text="#151922", dark=True)

_current = _LIGHT


def mix(base: str, top: str, amount: float) -> str:
    """``top`` laid over ``base`` with the given opacity, as a solid colour."""
    a, b = QColor(base), QColor(top)
    blend = lambda x, y: round(x + (y - x) * amount)   # noqa: E731
    return QColor(blend(a.red(), b.red()), blend(a.green(), b.green()),
                  blend(a.blue(), b.blue())).name()


def colors() -> Colors:
    return _current


def system_is_dark() -> bool:
    hints = QGuiApplication.styleHints()
    return hasattr(hints, "colorScheme") and hints.colorScheme() == Qt.ColorScheme.Dark


def resolve(mode: str) -> str:
    """'light' or 'dark' for a setting of '', 'light' or 'dark' ('' = like Windows)."""
    if mode in (LIGHT, DARK):
        return mode
    return DARK if system_is_dark() else LIGHT


def apply(app, mode: str, asset_dir: str) -> str:
    """Switch the application to the light or dark theme; returns which one."""
    global _current
    name = resolve(mode)
    _current = _DARK if name == DARK else _LIGHT
    hints = QGuiApplication.styleHints()
    if hasattr(hints, "setColorScheme"):   # Qt 6.8+: also the window frame
        scheme = {LIGHT: Qt.ColorScheme.Light, DARK: Qt.ColorScheme.Dark}.get(mode, Qt.ColorScheme.Unknown)
        hints.setColorScheme(scheme)
    app.setStyleSheet(style_sheet(_current, asset_dir))
    return name


def font(size: float = 10, weight=None) -> QFont:
    """The UI font; ``weight`` True means semi-bold."""
    f = QFont()
    f.setFamilies(["Segoe UI Variable Text", "Segoe UI", "Inter", "Noto Sans", "DejaVu Sans"])
    f.setPointSizeF(size)
    if weight is True:
        weight = QFont.Weight.DemiBold
    f.setWeight(weight or QFont.Weight.Normal)
    return f


def style_sheet(c: Colors, asset_dir: str) -> str:
    chevron = icons.write_png(asset_dir, "chevron-down", c.muted, 16)
    chevron = chevron.replace(os.sep, "/")
    soft = c.soft
    return f"""
* {{ outline: 0; }}
QWidget#root, QWidget#body {{ background: {c.bg}; }}
QWidget {{ color: {c.text}; }}
QToolTip {{ background: {c.toast_bg}; color: {c.toast_text}; border: 0; padding: 6px 8px; }}

QFrame#header {{ background: {c.surface}; border-bottom: 1px solid {c.border}; }}
QFrame[card="true"] {{ background: {c.surface}; border: 1px solid {c.border}; border-radius: 14px; }}
QFrame[card="true"][drag="true"] {{ border: 2px solid {c.accent}; }}
QFrame[box="true"] {{ background: {c.surface2}; border: 1px solid {c.border}; border-radius: 12px; }}
QFrame[callout="true"] {{ background: {soft(c.plot, .12)}; border: 1.5px dashed {c.plot}; border-radius: 12px; }}
QFrame[note="true"] {{ background: {soft(c.plot, .10)}; border: 0; border-radius: 10px; }}
QFrame[dropzone="true"] {{ background: {c.surface2}; border: 1.5px dashed {c.border_strong}; border-radius: 14px; }}
QFrame[dropzone="true"][drag="true"] {{ background: {soft(c.accent, .10)}; border-color: {c.accent}; }}
QFrame[row="true"] {{ background: {c.surface}; border: 1px solid {c.border}; border-radius: 12px; }}
QFrame[row="true"]:hover {{ border-color: {c.border_strong}; }}
QFrame#drawer {{ background: {c.surface}; border-left: 1px solid {c.border}; }}
QFrame[separator="true"] {{ background: {c.border}; border: 0; max-height: 1px; min-height: 1px; }}
QFrame#toast {{ background: {c.toast_bg}; border-radius: 12px; }}
QFrame#toast QLabel {{ color: {c.toast_text}; }}

QLabel {{ background: transparent; }}
QLabel[muted="true"] {{ color: {c.muted}; }}
QLabel[section="true"] {{ color: {c.muted}; font-weight: 700; letter-spacing: 1px; }}
QLabel[pill="a4"] {{ background: {soft(c.a4)}; color: {c.a4_text}; border-radius: 10px; padding: 1px 8px; font-weight: 600; }}
QLabel[pill="a3"] {{ background: {soft(c.a3, .12)}; color: {c.a3_text}; border-radius: 10px; padding: 1px 8px; font-weight: 600; }}
QLabel[pill="plot"] {{ background: {soft(c.plot, .11)}; color: {c.plot_text}; border-radius: 10px; padding: 1px 8px; font-weight: 600; }}
QLabel[pill="muted"] {{ background: {c.surface2}; color: {c.muted}; border: 1px solid {c.border}; border-radius: 10px; padding: 0 8px; font-weight: 600; }}
QLabel[pill="ok"] {{ background: {soft(c.ok, .12)}; color: {c.ok}; border-radius: 12px; padding: 2px 10px; font-weight: 600; }}
QLabel[pill="done"] {{ background: {soft(c.ok, .12)}; color: {c.ok}; border-radius: 10px; padding: 1px 8px; font-weight: 600; }}
QLabel[pill="warn"] {{ background: {soft(c.plot, .12)}; color: {c.plot_text}; border-radius: 10px; padding: 1px 8px; font-weight: 600; }}
QLabel[pill="error"] {{ background: {soft(c.danger, .12)}; color: {c.danger}; border-radius: 12px; padding: 2px 10px; font-weight: 600; }}
QLabel[pill="accent"] {{ background: {soft(c.accent, .10)}; color: {c.accent_text}; border-radius: 12px; padding: 2px 10px; font-weight: 600; }}
QLabel[pill="version"] {{ background: {c.surface2}; color: {c.muted}; border: 1px solid {c.border}; border-radius: 10px; padding: 0 8px; font-weight: 600; }}
QLabel[badge="a4"] {{ background: {soft(c.a4)}; color: {c.a4_text}; border-radius: 12px; font-weight: 700; }}
QLabel[badge="a3"] {{ background: {soft(c.a3, .12)}; color: {c.a3_text}; border-radius: 12px; font-weight: 700; }}
QLabel[badge="plot"] {{ background: {soft(c.plot, .11)}; color: {c.plot_text}; border-radius: 12px; }}
QLabel[logo="true"] {{ background: {c.accent}; border-radius: 11px; }}
QLabel[ok="true"] {{ color: {c.ok}; font-weight: 600; }}

QPushButton {{ border-radius: 10px; padding: 0 14px; min-height: 36px; font-weight: 600; }}
QPushButton[variant="primary"] {{ background: {c.accent}; color: #FFFFFF; border: 0; }}
QPushButton[variant="primary"]:hover {{ background: {c.accent_hover}; }}
QPushButton[variant="soft"] {{ background: {soft(c.accent, .09)}; color: {c.accent_text}; border: 0; }}
QPushButton[variant="soft"]:hover {{ background: {soft(c.accent, .18)}; }}
QPushButton[variant="ghost"] {{ background: {c.surface}; color: {c.text}; border: 1px solid {c.border}; font-weight: 500; }}
QPushButton[variant="ghost"]:hover {{ background: {c.surface2}; border-color: {c.border_strong}; }}
QPushButton[variant="plot"] {{ background: {c.plot_strong}; color: #FFFFFF; border: 0; }}
QPushButton[variant="plot"]:hover {{ background: {mix(c.plot_strong, '#000000', .12)}; }}
QPushButton[variant="link"] {{ background: transparent; color: {c.accent_text}; border: 0; padding: 0; min-height: 28px; text-align: left; }}
QPushButton[variant="link"]:hover {{ text-decoration: underline; }}
QPushButton:disabled {{ color: {mix(c.surface, c.text, .45)}; }}
QPushButton[variant="primary"]:disabled, QPushButton[variant="plot"]:disabled {{ background: {mix(c.surface, c.accent, .45)}; color: {mix(c.accent, '#FFFFFF', .75)}; }}
QPushButton[focusvisible="true"] {{ border: 2px solid {c.accent}; }}
QPushButton[variant="link"][focusvisible="true"] {{ border: 0; text-decoration: underline; }}

QPushButton[seg="true"] {{ background: transparent; border: 0; color: {c.muted}; font-weight: 500;
            padding: 0 12px; min-height: 30px; border-radius: 8px; }}
QPushButton[seg="true"]:hover {{ color: {c.text}; }}
QPushButton[seg="true"]:checked {{ color: {c.text}; font-weight: 600; }}
QPushButton[seg="true"][focusvisible="true"] {{ border: 0; color: {c.accent_text}; text-decoration: underline; }}
QPushButton[seg="true"]:disabled {{ color: {mix(c.surface, c.text, .4)}; }}

QToolButton {{ background: transparent; border: 1px solid transparent; border-radius: 10px; padding: 0; }}
QToolButton:hover {{ background: {c.surface2}; }}
QToolButton[focusvisible="true"] {{ border-color: {c.accent}; }}
QToolButton:disabled {{ background: transparent; }}

QComboBox {{ background: {c.surface2}; color: {c.text}; border: 1px solid {c.border}; border-radius: 10px;
            padding: 0 30px 0 12px; min-height: 36px; font-weight: 500; }}
QComboBox:hover {{ border-color: {c.border_strong}; }}
QComboBox[focusvisible="true"] {{ border: 2px solid {c.accent}; }}
QComboBox:disabled {{ color: {c.muted}; }}
QComboBox::drop-down {{ border: 0; width: 28px; subcontrol-origin: padding; subcontrol-position: center right; }}
QComboBox::down-arrow {{ image: url("{chevron}"); width: 14px; height: 14px; }}
QComboBox QAbstractItemView {{ background: {c.surface}; color: {c.text}; border: 1px solid {c.border};
            selection-background-color: {soft(c.accent, .12)}; selection-color: {c.text}; padding: 4px; }}

QScrollArea, QScrollArea > QWidget > QWidget {{ background: transparent; border: 0; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {c.border_strong}; border-radius: 3px; min-height: 32px; }}
QScrollBar::handle:vertical:hover {{ background: {c.muted}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}

QPlainTextEdit {{ background: {c.surface2}; color: {c.text}; border: 1px solid {c.border}; border-radius: 10px;
                 padding: 6px; selection-background-color: {soft(c.accent, .25)}; }}
"""
