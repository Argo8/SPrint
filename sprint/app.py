"""Starting SPrint: logging, Qt and the window.

PDF files or folders given as arguments are added to the list (so it works with
"Send to" and with dropping files on SPrint.exe); when SPrint is already open
they go to that window instead.
"""

from __future__ import annotations

import getpass
import logging
import logging.handlers
import os
import shutil
import sys
import tempfile
import time

from . import APP_NAME, AUTHOR, __version__
from .i18n import set_language, system_language, t
from .settings import Settings, data_dir

logger = logging.getLogger("sprint.app")

TEMP_PREFIX = "SPrint_"
STALE_TEMP_AGE_S = 24 * 3600


def resource_path(name: str) -> str:
    """Bundled file: next to the sources, or unpacked by PyInstaller."""
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


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


def _set_app_id() -> None:
    """Our own taskbar icon and grouping, also when run as a script."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"{AUTHOR}.{APP_NAME}")
        except Exception:
            pass


def _place_window(window) -> None:
    screen = window.screen().availableGeometry()
    width = min(1260, int(screen.width() * 0.92))
    height = min(880, int(screen.height() * 0.92))
    window.resize(width, height)
    window.move(screen.left() + (screen.width() - width) // 2, screen.top() + (screen.height() - height) // 2)


def main(argv: list[str] | None = None) -> None:
    """Start SPrint.

    ``--smoke-test`` closes the window once the files are read and exits with
    status 1 if any could not be read; the build pipeline uses it.
    """
    argv = sys.argv[1:] if argv is None else argv
    smoke_test = "--smoke-test" in argv
    paths = [arg for arg in argv if not arg.startswith("--")]
    _setup_logging()
    _set_app_id()

    from PySide6.QtCore import QTimer
    from PySide6.QtGui import QIcon
    from PySide6.QtWidgets import QApplication

    from . import printing
    from .pdfdocs import render_preview
    from .single_instance import SingleInstance
    from .ui import theme
    from .ui.controller import Controller, Services
    from .ui.main_window import MainWindow

    app = QApplication([sys.argv[0]])
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(__version__)
    app.setOrganizationName(AUTHOR)
    app.setStyle("Fusion")
    app.setFont(theme.font(10))
    app.setWindowIcon(QIcon(resource_path("icon.ico")))
    from .ui.widgets import FocusVisible
    app.focus_visible = FocusVisible(app)

    instance = None
    if not smoke_test:
        instance = SingleInstance(f"{APP_NAME}-{getpass.getuser()}")
        if instance.send(paths):
            logger.info("SPrint is already open; passed it %d paths", len(paths))
            return

    settings = Settings.load()
    set_language(settings.language or system_language())
    _clean_stale_temp_dirs()
    temp_dir = tempfile.mkdtemp(prefix=TEMP_PREFIX)
    asset_dir = os.path.join(temp_dir, "ui")
    theme.apply(app, asset_dir)

    backend = printing.find_backend(settings.print_program)
    controller = Controller(settings, Services(backend=backend), temp_dir)
    window = MainWindow(controller)
    _place_window(window)
    logger.info("SPrint %s, print program: %s", __version__, backend.key)
    controller.log(t(f"backend_{backend.key}"))
    if settings.print_program and settings.print_program != backend.key:
        controller.log(t("program_missing", program=settings.print_program), "warn")

    if instance is not None:
        def received(new_paths):
            controller.add_paths(new_paths)
            window.bring_to_front()
        instance.listen(received)

    controller.load_printers()
    controller.add_paths(paths)
    window.show()

    failed: list[str] = []
    if smoke_test:
        def finish():
            if controller.reading:
                QTimer.singleShot(200, finish)
                return
            failed.extend(os.path.basename(e.path) for e in controller.entries if e.doc is not None and not e.doc.ok)
            try:   # the page renderer used for printing must be bundled too
                for entry in controller.entries[:1]:
                    render_preview(entry.path, 0, 64)
                window.grab()
            except Exception:
                logger.exception("PDF renderer failed")
                failed.append("PDF renderer")
            logger.info("Smoke test: %d files read, failed: %s", len(controller.entries), failed or "none")
            window.close()
        QTimer.singleShot(800, finish)

    app.exec()
    shutil.rmtree(temp_dir, ignore_errors=True)
    if failed:
        sys.exit(1)
