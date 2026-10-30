"""The window, driven like a user would, with fake printers (Qt offscreen)."""

import os
import subprocess
import sys

import pytest

pytest.importorskip("PySide6")
pytest.importorskip("pytestqt")

from PySide6.QtCore import Qt  # noqa: E402

from sprint import printing  # noqa: E402
from sprint.settings import Settings  # noqa: E402
from sprint.ui import theme  # noqa: E402
from sprint.ui.controller import Controller, Services  # noqa: E402
from sprint.ui.main_window import MainWindow  # noqa: E402
from sprint.ui.widgets import FocusVisible  # noqa: E402

A4 = (210, 297)
PRINTERS = ["Canon iR-ADV C3530", "HP DesignJet T650", "HP LaserJet Pro M404"]


class FakeBackend(printing.Backend):
    key = "direct"

    def __init__(self):
        self.calls = []

    def print_file(self, printer, path, copies=1, paper=None, duplex=None):
        self.calls.append((printer, os.path.basename(path), copies, paper, duplex))
        return True


@pytest.fixture
def services():
    return Services(
        list_printers=lambda: list(PRINTERS),
        detect_printers=lambda names: {"a4": "HP LaserJet Pro M404", "a3": "Canon iR-ADV C3530",
                                       "plotter": "HP DesignJet T650"},
        printer_exists=lambda name: name in PRINTERS,
        supports_duplex=lambda name: name.startswith("HP LaserJet"),
        default_duplex=lambda name: False,
        open_preferences=lambda name: None,
        backend=FakeBackend(), plot_backend=FakeBackend())


@pytest.fixture
def window(qtbot, tmp_path, services, monkeypatch):
    settings = Settings(language="en")
    monkeypatch.setattr(Settings, "save", lambda self, path=None: None)
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance()
    if not hasattr(app, "focus_visible"):
        app.focus_visible = FocusVisible(app)
    theme.apply(app, "light", str(tmp_path / "ui"))
    controller = Controller(settings, services, str(tmp_path))
    win = MainWindow(controller, str(tmp_path / "ui"))
    win.resize(1280, 860)
    qtbot.addWidget(win)
    win.show()
    controller.load_printers()
    qtbot.waitUntil(lambda: controller.printers.plotter == "HP DesignJet T650")
    return win


@pytest.fixture
def project(make_pdf):
    return [make_pdf("plan.pdf", [(841, 594), (841, 594), (594, 420)]),
            make_pdf("bill.pdf", [A4] * 12),
            make_pdf("details.pdf", [(420, 297), (297, 420), (420, 297), (420, 297)]),
            make_pdf("site.pdf", [(1189, 841)]),
            make_pdf("offer.pdf", [A4, A4])]


def load(qtbot, window, paths):
    window.c.add_paths(paths)
    qtbot.waitUntil(lambda: window.c.reading == 0 and all(e.doc for e in window.c.entries))


def print_now(qtbot, window):
    assert window.print_button.isEnabled()
    qtbot.mouseClick(window.print_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: window.c.job_state == "done", timeout=10000)


def test_empty_window(window):
    assert window.files_stack.currentIndex() == 0
    assert not window.print_button.isEnabled()
    assert window.print_button.sub == "Add PDF files"
    assert window.a4.summary.text() == "No A4 pages"


def test_files_and_destinations(qtbot, window, project):
    load(qtbot, window, project)
    assert window.files_stack.currentIndex() == 1
    assert window.files_summary.text() == "5 files · 22 pages"
    assert window.a4.summary.text() == "14 pages"
    assert window.a3.summary.text() == "4 pages"
    assert window.plot.summary.text() == "4 pages · needs rolls 420, 610 and 914 mm"
    assert window.roll_chips[610].isChecked() and window.roll_chips[610].sub == "2 p."
    assert window.roll_chips[297].sub == "none"
    assert window.roll_info.text().startswith("2 pages · ≈ 1.7 m of paper")
    assert window.print_button.title == "Print"
    assert window.print_button.sub == "A4 14 · A3 4 · roll 610: 2"
    # adding the same files again changes nothing
    window.c.add_paths(project[:2])
    assert len(window.c.entries) == 5


def test_roll_by_roll(qtbot, window, project, services):
    load(qtbot, window, project)
    print_now(qtbot, window)
    assert [c[1] for c in services.backend.calls] == ["bill.pdf", "offer.pdf", "details.pdf"]
    assert services.plot_backend.calls == [("HP DesignJet T650", "plan [610x841].pdf", 1, (610, 841), None)]
    assert window.callout.isVisible()
    assert window.next_button.text() == "Roll 420 is loaded"
    assert window.a4.summary.text() == "14 pages · printed ✓"
    assert window.a4.reprint.isVisible()
    qtbot.waitUntil(lambda: window.c.job_state == "idle", timeout=5000)
    assert not window.print_button.isEnabled()
    assert window.print_button.sub == "Next: load the 420 mm roll"

    qtbot.mouseClick(window.next_button, Qt.MouseButton.LeftButton)
    assert window.c.opts.roll == 420 and window.c.settings.roll == 420
    assert window.print_button.title == "Plot on roll 420"
    print_now(qtbot, window)
    assert services.plot_backend.calls[-1] == ("HP DesignJet T650", "plan [420x594].pdf", 1, (420, 594), None)
    qtbot.waitUntil(lambda: window.c.job_state == "idle", timeout=5000)
    qtbot.mouseClick(window.next_button, Qt.MouseButton.LeftButton)
    print_now(qtbot, window)
    assert services.plot_backend.calls[-1][1:4] == ("site.pdf", 1, (914, 1189))
    assert window.plot_done.isVisible()
    qtbot.waitUntil(lambda: window.c.job_state == "idle", timeout=5000)
    assert window.print_button.title == "All printed"


def test_skip_a_page_and_switch_off(qtbot, window, project):
    load(qtbot, window, project[:2])
    row = window.rows[id(window.c.entries[1])]
    qtbot.mouseClick(row.header, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: len(row.tiles) == 12)
    qtbot.mouseClick(row.tiles[0], Qt.MouseButton.LeftButton)
    assert window.c.entries[1].states[0].skip
    assert window.a4.summary.text() == "11 pages"
    assert row.tiles[0].dest == "Skipped"
    qtbot.mouseClick(window.a4.switch, Qt.MouseButton.LeftButton)
    assert not window.c.opts.a4_on
    assert window.a4.summary.text() == "11 pages · switched off"
    assert window.print_button.title == "Plot on roll 610"
    assert window.print_button.sub == "2 pages · ≈ 1.7 m of paper"


def test_a3_on_the_plotter(qtbot, window, project, services):
    load(qtbot, window, [project[2]])
    window.a3_target._clicked(1)   # "Plotter (roll 297)"
    assert window.c.opts.a3_to_plotter and window.a3_note.isVisible()
    assert window.roll_chips[297].sub == "4 p."
    qtbot.mouseClick(window.roll_chips[297], Qt.MouseButton.LeftButton)
    print_now(qtbot, window)
    assert services.backend.calls == []
    assert [c[3] for c in services.plot_backend.calls] == [(297, 420)]


def test_missing_printer_blocks_printing(qtbot, window, project):
    load(qtbot, window, [project[3]])
    window.c.set_roll(914)
    assert window.print_button.isEnabled()
    window.c.set_printer("plot", "")
    assert not window.print_button.isEnabled()
    assert window.print_button.sub == "Choose a printer for the plotter"


def test_copies_and_duplex(qtbot, window, project, services):
    load(qtbot, window, [project[1]])
    qtbot.waitUntil(lambda: window.a4_duplex.isVisible())
    qtbot.mouseClick(window.a4_duplex, Qt.MouseButton.LeftButton)
    qtbot.mouseClick(window.more, Qt.MouseButton.LeftButton)
    assert window.copies_value.text() == "2"
    assert window.a4.summary.text() == "12 pages · 6 sheets"
    assert window.print_button.sub == "A4 12 · 2 copies"
    print_now(qtbot, window)
    assert services.backend.calls == [("HP LaserJet Pro M404", "bill.pdf", 2, "A4", True)]


def test_settings_drawer_theme_and_language(qtbot, window):
    qtbot.mouseClick(window.settings_button, Qt.MouseButton.LeftButton)
    assert window.drawer.is_open
    window.theme_choice._clicked(2)
    assert theme.colors().dark and window.c.settings.theme == "dark"
    window.language_choice._clicked(0)
    assert window.files_title.text() == "Datoteke"
    assert window.print_button.sub == "Dodaj PDF datoteke"
    window.drawer.close_drawer()
    assert not window.drawer.is_open
    window.set_theme("light")
    window.set_language("en")


def test_only_pdfs_are_added(qtbot, window, tmp_path):
    other = tmp_path / "notes.txt"
    other.write_text("x")
    messages = []
    window.c.toast.connect(lambda text, kind: messages.append(text))
    assert window.c.add_paths([str(other)]) == 0
    assert messages == ["Only PDF files can be added"]


def test_folders_add_the_pdfs_in_them(qtbot, window, make_pdf, tmp_path):
    make_pdf("b.pdf", [A4])
    make_pdf("a.pdf", [A4])
    load(qtbot, window, [str(tmp_path)])
    assert [os.path.basename(e.path) for e in window.c.entries] == ["a.pdf", "b.pdf"]


def test_remove_and_clear(qtbot, window, project):
    load(qtbot, window, project[:3])
    row = window.rows[id(window.c.entries[0])]
    qtbot.mouseClick(row.remove_button, Qt.MouseButton.LeftButton)
    assert len(window.c.entries) == 2
    qtbot.mouseClick(window.clear_button, Qt.MouseButton.LeftButton)
    assert window.c.entries == [] and window.files_stack.currentIndex() == 0


def test_changed_file_is_read_again(qtbot, window, make_pdf):
    path = make_pdf("plan.pdf", [A4])
    load(qtbot, window, [path])
    stamp = window.c.entries[0].stamp
    os.utime(path, ns=(stamp[0] + 5_000_000_000, stamp[0] + 5_000_000_000))
    assert not window.c.start_print()
    qtbot.waitUntil(lambda: window.c.entries[0].doc is not None)


def test_single_instance_hands_over_paths(qtbot):
    from sprint.single_instance import SingleInstance
    name = f"SPrint-test-{os.getpid()}"
    received = []
    server = SingleInstance(name)
    server.listen(received.extend)
    # the second start is another program, as when "Send to" starts SPrint again
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    code = ("import sys; sys.path.insert(0, sys.argv[1]); from PySide6.QtCore import QCoreApplication; "
            "app = QCoreApplication([]); from sprint.single_instance import SingleInstance; "
            "sys.exit(0 if SingleInstance(sys.argv[2]).send(['a.pdf', 'b.pdf']) else 1)")
    second = subprocess.Popen([sys.executable, "-c", code, root, name])
    qtbot.waitUntil(lambda: len(received) == 2, timeout=20000)
    assert second.wait(20) == 0
    assert [os.path.basename(p) for p in received] == ["a.pdf", "b.pdf"]
    assert not SingleInstance(name + "-none").send(["x.pdf"])
