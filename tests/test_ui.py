"""The window, driven like a user would, with fake printers (Qt offscreen)."""

import os
import subprocess
import sys
import threading

import pytest

pytest.importorskip("PySide6")
pytest.importorskip("pytestqt")

from PySide6.QtCore import QEvent, QPointF, Qt  # noqa: E402
from PySide6.QtGui import QEnterEvent  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from sprint import printing  # noqa: E402
from sprint.plan import A3_DEST, A4_DEST, PLOT_DEST  # noqa: E402
from sprint.settings import Settings  # noqa: E402
from sprint.ui import theme  # noqa: E402
from sprint.ui import controller as controller_module  # noqa: E402
from sprint.ui.controller import Controller, Services  # noqa: E402
from sprint.ui.main_window import MainWindow  # noqa: E402
from sprint.ui.widgets import FocusVisible  # noqa: E402

A4 = (210, 297)
PRINTERS = ["Canon iR-ADV C3530", "HP DesignJet T650", "HP LaserJet Pro M404"]


class FakeBackend(printing.Backend):
    key = "direct"

    def __init__(self):
        self.calls = []
        self.gate = None   # set: printing waits until it opens

    def print_file(self, printer, path, copies=1, paper=None, duplex=None):
        if self.gate is not None:
            self.gate.wait(10)
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
    monkeypatch.setattr(controller_module, "DONE_SHOW_MS", 30)
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance()
    if not hasattr(app, "focus_visible"):
        app.focus_visible = FocusVisible(app)
    theme.apply(app, str(tmp_path / "ui"))
    controller = Controller(settings, services, str(tmp_path))
    win = MainWindow(controller)
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


def print_card(qtbot, window, card):
    qtbot.waitUntil(lambda: card.action.currentIndex() == 0)   # not still showing "Sent"
    assert card.print_button.isEnabled()
    qtbot.mouseClick(card.print_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not window.c.busy, timeout=10000)


def test_empty_window(window):
    assert window.files_stack.currentIndex() == 0
    assert window.a4.summary.text() == "No A4 pages"
    assert [card.print_button.text() for card in (window.a4, window.a3, window.plot)] == ["Print", "Print", "Plot"]
    assert not any(card.print_button.isEnabled() for card in (window.a4, window.a3, window.plot))


def test_files_and_destinations(qtbot, window, project):
    load(qtbot, window, project)
    assert window.files_stack.currentIndex() == 1
    assert window.files_summary.text() == "5 files · 22 pages"
    assert window.a4.summary.text() == "14 pages"
    assert window.a3.summary.text() == "4 pages"
    assert window.plot.summary.text() == "4 pages · needs rolls 420, 610 and 914 mm"
    assert window.roll_chips[610].isChecked() and window.roll_chips[610].sub == "2 p."
    assert window.roll_chips[297].sub == "none"
    assert window.roll_info.text().startswith("Roll 610 mm: 2 pages · ≈ 1.7 m of paper")
    assert window.roll_later.text() == "Then still needed: 420 and 914 mm"
    assert window.a4.print_button.text() == "Print 14 pages"
    assert window.a3.print_button.text() == "Print 4 pages"
    assert window.plot.print_button.text() == "Plot 2 pages"
    # adding the same files again changes nothing
    window.c.add_paths(project[:2])
    assert len(window.c.entries) == 5


def test_each_card_prints_its_own_pages(qtbot, window, project, services):
    load(qtbot, window, project)
    print_card(qtbot, window, window.a3)
    assert [c[1] for c in services.backend.calls] == ["details.pdf"]
    assert services.plot_backend.calls == []
    assert window.a3.summary.text() == "4 pages · printed ✓"
    assert window.a3.reprint.isVisible()
    qtbot.waitUntil(lambda: window.a3.action.currentIndex() == 0)
    assert window.a3.print_button.text() == "Everything printed"
    assert not window.a3.print_button.isEnabled()
    assert window.a4.print_button.isEnabled() and window.a4.print_button.text() == "Print 14 pages"


def test_roll_by_roll(qtbot, window, project, services):
    load(qtbot, window, project)
    print_card(qtbot, window, window.plot)
    assert services.plot_backend.calls == [("HP DesignJet T650", "plan [610x841].pdf", 1, (610, 841), None)]
    assert services.backend.calls == []
    qtbot.waitUntil(lambda: window.plot.action.currentIndex() == 0)
    assert window.plot.print_button.text() == "Roll 610 is done"
    assert not window.plot.print_button.isEnabled()
    assert window.roll_later.text() == "Next: load the 420 mm roll and choose it above (1 page)."

    qtbot.mouseClick(window.roll_chips[420], Qt.MouseButton.LeftButton)   # the roll was changed
    assert window.c.opts.roll == 420 and window.c.settings.roll == 420
    assert window.plot.print_button.text() == "Plot 1 page"
    print_card(qtbot, window, window.plot)
    assert services.plot_backend.calls[-1] == ("HP DesignJet T650", "plan [420x594].pdf", 1, (420, 594), None)
    qtbot.mouseClick(window.roll_chips[914], Qt.MouseButton.LeftButton)
    print_card(qtbot, window, window.plot)
    assert services.plot_backend.calls[-1][1:4] == ("site.pdf", 1, (914, 1189))
    qtbot.waitUntil(lambda: window.plot.action.currentIndex() == 0)
    assert window.plot.summary.text() == "4 pages · all plotted ✓"
    assert window.plot.print_button.text() == "Roll 914 is done"


def test_a_second_print_waits_its_turn(qtbot, window, project, services):
    load(qtbot, window, project)
    services.backend.gate = threading.Event()
    qtbot.mouseClick(window.a4.print_button, Qt.MouseButton.LeftButton)
    qtbot.mouseClick(window.plot.print_button, Qt.MouseButton.LeftButton)
    assert (window.c.status(A4_DEST), window.c.status(PLOT_DEST)) == ("printing", "waiting")
    assert window.a4.stop_button.text() == "Stop" and window.plot.stop_button.text() == "Cancel"
    assert window.plot.status_text.text() == "Waiting its turn…"
    services.backend.gate.set()
    qtbot.waitUntil(lambda: not window.c.busy, timeout=10000)
    assert [c[1] for c in services.backend.calls] == ["bill.pdf", "offer.pdf"]
    assert [c[1] for c in services.plot_backend.calls] == ["plan [610x841].pdf"]


def test_a_waiting_print_can_be_cancelled(qtbot, window, project, services):
    load(qtbot, window, project)
    services.backend.gate = threading.Event()
    qtbot.mouseClick(window.a4.print_button, Qt.MouseButton.LeftButton)
    qtbot.mouseClick(window.a3.print_button, Qt.MouseButton.LeftButton)
    qtbot.mouseClick(window.a3.stop_button, Qt.MouseButton.LeftButton)
    assert window.c.status(A3_DEST) == "idle" and window.a3.print_button.text() == "Print 4 pages"
    services.backend.gate.set()
    qtbot.waitUntil(lambda: not window.c.busy, timeout=10000)
    assert [c[1] for c in services.backend.calls] == ["bill.pdf", "offer.pdf"]


def test_skip_a_page(qtbot, window, project):
    load(qtbot, window, project[:2])
    row = window.rows[id(window.c.entries[1])]
    qtbot.mouseClick(row.header, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: len(row.tiles) == 12)
    qtbot.mouseClick(row.tiles[0], Qt.MouseButton.LeftButton)
    assert window.c.entries[1].states[0].skip
    assert window.a4.summary.text() == "11 pages"
    assert row.tiles[0].dest == "Skipped"
    assert window.a4.print_button.text() == "Print 11 pages"


def test_a3_on_the_plotter(qtbot, window, project, services):
    load(qtbot, window, [project[2]])
    window.a3_target._clicked(1)   # "Plotter (roll 297)"
    assert window.c.opts.a3_to_plotter and window.a3_note.isVisible()
    assert not window.a3.action.isVisible()
    assert window.roll_chips[297].sub == "4 p."
    qtbot.mouseClick(window.roll_chips[297], Qt.MouseButton.LeftButton)
    print_card(qtbot, window, window.plot)
    assert services.backend.calls == []
    assert [c[3] for c in services.plot_backend.calls] == [(297, 420)]


def test_missing_printer_blocks_printing(qtbot, window, project):
    load(qtbot, window, [project[3]])
    window.c.set_roll(914)
    assert window.plot.print_button.isEnabled()
    window.c.set_printer("plot", "")
    assert not window.plot.print_button.isEnabled()
    assert window.plot.print_button.text() == "Choose a printer first"


def test_copies_and_duplex(qtbot, window, project, services):
    load(qtbot, window, [project[1]])
    qtbot.waitUntil(lambda: window.a4_duplex.isVisible())
    qtbot.mouseClick(window.a4_duplex, Qt.MouseButton.LeftButton)
    qtbot.mouseClick(window.more, Qt.MouseButton.LeftButton)
    assert window.copies_value.text() == "2"
    assert window.a4.summary.text() == "12 pages · 6 sheets"
    assert window.a4.print_button.text() == "Print 12 pages · 2 copies"
    print_card(qtbot, window, window.a4)
    assert services.backend.calls == [("HP LaserJet Pro M404", "bill.pdf", 2, "A4", True)]


def test_printer_settings_beside_each_printer(qtbot, window, services):
    opened = []
    services.open_preferences = opened.append
    for button in (window.a4_settings, window.a3_settings, window.plot_settings):
        qtbot.mouseClick(button, Qt.MouseButton.LeftButton)
    assert opened == ["HP LaserJet Pro M404", "Canon iR-ADV C3530", "HP DesignJet T650"]


def test_files_show_their_first_page_without_a_click(qtbot, window, project):
    load(qtbot, window, project)
    rows = list(window.rows.values())
    qtbot.waitUntil(lambda: all(row.header.shape.picture is not None for row in rows), timeout=10000)
    assert not any(row.pages.is_open for row in rows)
    shape = rows[0].header.shape
    assert max(shape.picture.width(), shape.picture.height()) == shape.picture_px()   # drawn 1:1


def test_resting_on_a_page_shows_it_large(qtbot, window, project):
    load(qtbot, window, [project[0]])
    row = window.rows[id(window.c.entries[0])]
    qtbot.mouseClick(row.header, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: len(row.tiles) == 3)
    popup = window.peek.popup
    QApplication.sendEvent(row.tiles[2], QEnterEvent(QPointF(5, 5), QPointF(5, 5), QPointF(5, 5)))
    qtbot.waitUntil(popup.isVisible, timeout=3000)
    assert (popup.title, popup.detail) == ("plan.pdf", "Page 3 · 594×420 · Roll 420")
    qtbot.waitUntil(lambda: not popup.spinner.isVisible(), timeout=10000)
    assert popup.picture.width() > 300
    QApplication.sendEvent(row.tiles[2], QEvent(QEvent.Type.Leave))
    qtbot.waitUntil(lambda: not popup.isVisible(), timeout=3000)


def test_settings_drawer_and_language(qtbot, window):
    qtbot.mouseClick(window.settings_button, Qt.MouseButton.LeftButton)
    assert window.drawer.is_open
    window.language_choice._clicked(0)
    assert window.files_title.text() == "Datoteke"
    assert window.a4.print_button.text() == "Ispiši"
    assert window.c.settings.language == "hr"
    window.drawer.close_drawer()
    assert not window.drawer.is_open
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
    assert not window.c.start_print(A4_DEST)
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
