"""User settings, stored in %APPDATA%\\SPrint\\settings.ini.

Older versions kept ``printer_choices.ini`` in the working directory; it is read
once as a fallback so existing printer choices carry over.
"""

from __future__ import annotations

import configparser
import logging
import os
import sys
from dataclasses import dataclass, fields
from typing import Optional

log = logging.getLogger(__name__)

APP_DIR_NAME = "SPrint"
SETTINGS_FILE = "settings.ini"
LEGACY_FILE = "printer_choices.ini"


def data_dir() -> str:
    """Per-user folder for settings and the log file."""
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, APP_DIR_NAME)


def app_dir() -> str:
    """Folder of SPrint.exe, or of the sources when run as a script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@dataclass
class Settings:
    printer_a4: str = ""
    printer_a3: str = ""
    printer_plotter: str = ""
    roll: int = 610               # the roll in the plotter, remembered between sessions
    a3_to_plotter: bool = False   # A3 pages on the plotter's 297 mm roll
    language: str = ""            # "" until the user picks one
    print_program: str = ""       # "" = direct; "sumatra" or "adobe" for A4/A3 pages
    clear_when_done: bool = False  # empty the list once everything is printed

    _SECTIONS = {
        "printer_a4": ("printers", "a4"),
        "printer_a3": ("printers", "a3"),
        "printer_plotter": ("printers", "plotter"),
        "roll": ("plotter", "roll"),
        "a3_to_plotter": ("plotter", "a3"),
        "language": ("ui", "language"),
        "print_program": ("printing", "program"),
        "clear_when_done": ("printing", "clear_when_done"),
    }

    @property
    def has_printers(self) -> bool:
        return bool(self.printer_a4 or self.printer_a3 or self.printer_plotter)

    @classmethod
    def load(cls, path: Optional[str] = None) -> "Settings":
        path = path or os.path.join(data_dir(), SETTINGS_FILE)
        settings = cls()
        parser = _parser()
        try:
            if parser.read(path, encoding="utf-8"):
                for f in fields(cls):
                    settings._read(parser, f)
                return settings
        except (configparser.Error, UnicodeDecodeError, OSError):
            log.warning("Could not read %s", path, exc_info=True)
        settings._load_legacy()
        return settings

    def _read(self, parser: configparser.ConfigParser, f) -> None:
        section, key = self._SECTIONS[f.name]
        default = getattr(self, f.name)
        try:
            if isinstance(default, bool):
                value = parser.getboolean(section, key, fallback=default)
            elif isinstance(default, int):
                value = parser.getint(section, key, fallback=default)
            else:
                value = parser.get(section, key, fallback=default)
        except ValueError:
            log.warning("Ignoring invalid %s/%s in the settings", section, key)
            return
        setattr(self, f.name, value)

    def _load_legacy(self) -> None:
        for folder in dict.fromkeys((os.getcwd(), app_dir())):
            legacy = os.path.join(folder, LEGACY_FILE)
            if not os.path.isfile(legacy):
                continue
            # The old version wrote this file in the system code page.
            for encoding in ("utf-8", "mbcs" if os.name == "nt" else "latin-1"):
                parser = _parser()
                try:
                    parser.read(legacy, encoding=encoding)
                except (configparser.Error, UnicodeDecodeError, LookupError, OSError):
                    continue
                self.printer_a4 = parser.get("PRINTERS", "A4", fallback="")
                self.printer_a3 = parser.get("PRINTERS", "A3", fallback="")
                self.printer_plotter = parser.get("PRINTERS", "Large", fallback="")
                return

    def save(self, path: Optional[str] = None) -> None:
        path = path or os.path.join(data_dir(), SETTINGS_FILE)
        parser = _parser()
        for f in fields(self):
            section, key = self._SECTIONS[f.name]
            if not parser.has_section(section):
                parser.add_section(section)
            value = getattr(self, f.name)
            parser.set(section, key, ("yes" if value else "no") if isinstance(value, bool) else str(value))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            parser.write(fh)
        os.replace(tmp, path)


def _parser() -> configparser.ConfigParser:
    # No interpolation: printer names may contain "%".
    return configparser.ConfigParser(interpolation=None)
