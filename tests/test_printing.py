"""Printer logic against a fake win32print (the real one is exercised in test_windows.py)."""

import copy

import pytest

from sprint import printing


class FakeDevMode:
    def __init__(self, **values):
        self.Fields = 0
        self.Orientation = 2
        self.PaperSize = 9
        self.PaperWidth = 0
        self.PaperLength = 0
        self.__dict__.update(values)


class FakeWin32Print:
    def __init__(self, per_user=None, reject=False):
        self.per_user = per_user
        self.global_devmode = FakeDevMode(PaperSize=1)
        self.reject = reject
        self.saved = []        # (level, devmode) passed to SetPrinter
        self.open_handles = 0
        self.jobs = []

    def OpenPrinter(self, name, defaults=None):
        self.open_handles += 1
        return name

    def ClosePrinter(self, handle):
        self.open_handles -= 1

    def GetPrinter(self, handle, level):
        devmode = self.per_user if level == 9 else self.global_devmode
        return {"pDevMode": copy.copy(devmode)}

    def DocumentProperties(self, hwnd, handle, name, out, inp, mode):
        assert out is inp and mode == printing.DM_IN_BUFFER | printing.DM_OUT_BUFFER
        if self.reject:
            out.PaperSize = 1
        return 1

    def SetPrinter(self, handle, level, info, command):
        self.saved.append((level, copy.copy(info["pDevMode"])))
        if level == 9:
            self.per_user = copy.copy(info["pDevMode"])

    def EnumJobs(self, handle, first, count, level):
        return [{"JobId": j} for j in self.jobs.pop(0)] if self.jobs else []


@pytest.fixture
def fake(monkeypatch):
    def install(**kwargs):
        win32 = FakeWin32Print(**kwargs)
        monkeypatch.setattr(printing, "win32print", win32)
        return win32
    return install


def test_custom_paper_set_and_restored_from_global(fake):
    win32 = fake()
    with printing.PlotterPaper("Plotter") as paper:
        assert paper.apply(610, 841.4) == (610, 841.4)
        dm = win32.per_user
        assert (dm.PaperSize, dm.PaperWidth, dm.PaperLength, dm.Orientation) == (256, 6100, 8414, 1)
        assert dm.Fields & 0xF == 0xF
    assert paper.restore_error is None
    assert [level for level, _ in win32.saved] == [9, 9]
    assert win32.per_user.PaperSize == 1          # no per-user settings before: global copied
    assert win32.open_handles == 0


def test_named_paper_and_restore_of_user_settings(fake):
    original = FakeDevMode(PaperSize=9, Fields=printing.DM_PAPERWIDTH | printing.DM_PAPERLENGTH)
    win32 = fake(per_user=original)
    with printing.PlotterPaper("Plotter") as paper:
        paper.apply(297, 420)
        assert win32.per_user.PaperSize == 8                      # DMPAPER_A3
        assert not win32.per_user.Fields & (printing.DM_PAPERWIDTH | printing.DM_PAPERLENGTH)
    assert vars(win32.per_user) == vars(original)


def test_rejected_size_changes_nothing(fake):
    win32 = fake(reject=True)
    with printing.PlotterPaper("Plotter") as paper:
        with pytest.raises(printing.PrintError):
            paper.apply(610, 841)
    assert win32.saved == []


def test_wait_for_new_job(fake, monkeypatch):
    monkeypatch.setattr(printing.time, "sleep", lambda s: None)
    win32 = fake()
    win32.jobs = [[1], [1], [1, 2]]
    assert printing._wait_for_new_job("P", {1}, timeout=5)
    assert not printing._wait_for_new_job("P", None)


def test_backend_preference(monkeypatch):
    monkeypatch.setattr(printing, "find_sumatra", lambda: "C:/S/SumatraPDF.exe")
    monkeypatch.setattr(printing, "find_adobe", lambda: "C:/A/Acrobat.exe")
    assert isinstance(printing.find_backend(), printing.SumatraBackend)
    monkeypatch.setattr(printing, "find_sumatra", lambda: None)
    assert isinstance(printing.find_backend(), printing.AdobeBackend)
    monkeypatch.setattr(printing, "find_adobe", lambda: None)
    assert isinstance(printing.find_backend(), printing.ShellBackend)


def test_sumatra_command_line(monkeypatch):
    calls = []

    class Result:
        returncode = 0
    monkeypatch.setattr(printing.subprocess, "run", lambda args, **kw: calls.append(args) or Result())
    backend = printing.SumatraBackend("SumatraPDF.exe")
    assert backend.print_file("HP DesignJet", "C:/a b.pdf", 3)
    assert backend.print_file("Office", "C:/a.pdf", 1, paper="A3")
    assert backend.print_file("Office", "C:/a.pdf", 2, paper="A4")
    assert calls == [
        ["SumatraPDF.exe", "-print-to", "HP DesignJet", "-silent", "-print-settings", "3x", "C:/a b.pdf"],
        ["SumatraPDF.exe", "-print-to", "Office", "-silent", "-print-settings", "paper=A3", "C:/a.pdf"],
        ["SumatraPDF.exe", "-print-to", "Office", "-silent", "-print-settings", "2x,paper=A4", "C:/a.pdf"],
    ]
    Result.returncode = 1
    with pytest.raises(printing.PrintError):
        backend.print_file("HP DesignJet", "C:/a b.pdf")
