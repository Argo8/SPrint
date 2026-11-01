"""Printer logic against a fake win32print (the real one is exercised in test_windows.py)."""

import copy

import pytest

from sprint import printing
from sprint.printing import place_page


class FakeDevMode:
    def __init__(self, **values):
        self.Fields = 0
        self.Orientation = 2
        self.PaperSize = 9
        self.PaperWidth = 0
        self.PaperLength = 0
        self.Copies = 1
        self.Collate = 0
        self.Duplex = 1
        self.Size = 220          # the public part of DEVMODEW
        self.DriverExtra = 0     # the driver's own settings after it
        self.__dict__.update(values)


class FakeWin32Print:
    def __init__(self, reject_custom=False, max_copies=999):
        self.devmode = FakeDevMode(PaperSize=1, Fields=printing.DM_PAPERWIDTH)
        self.user_devmode = None   # this user's printing preferences, if changed
        self.devmode_size = 220    # what the driver writes now
        self.swaps = {}            # custom size (tenths of mm) -> paper ID the driver uses instead
        self.reject_custom = reject_custom
        self.max_copies = max_copies
        self.open_handles = 0
        self.jobs = []
        self.capabilities = {printing.DC_COPIES: 99, printing.DC_COLLATE: 1}
        self.per_printer = {}     # printer -> {capability: value}
        self.default = None

    def OpenPrinter(self, name, defaults=None):
        self.open_handles += 1
        return name

    def ClosePrinter(self, handle):
        self.open_handles -= 1

    def GetPrinter(self, handle, level):
        if level == 9:
            return {"pDevMode": copy.copy(self.user_devmode)}
        assert level == 2
        return {"pDevMode": copy.copy(self.devmode), "pPortName": "USB001"}

    def DeviceCapabilities(self, device, port, capability):
        if device in self.per_printer:
            return self.per_printer[device].get(capability)
        return self.capabilities[capability]

    def GetDefaultPrinter(self):
        return self.default

    def DocumentProperties(self, hwnd, handle, name, out, inp, mode):
        if mode == 0:
            return self.devmode_size
        assert mode == printing.DM_IN_BUFFER | printing.DM_OUT_BUFFER
        assert out.Size + out.DriverExtra >= self.devmode_size, "the driver would write past the buffer"
        if out is not inp:
            out.__dict__.update({k: v for k, v in inp.__dict__.items() if k != "DriverExtra"})
        if out.PaperSize == printing.DMPAPER_USER and (out.PaperWidth, out.PaperLength) in self.swaps:
            out.PaperSize = self.swaps[(out.PaperWidth, out.PaperLength)]
        if self.reject_custom and out.PaperSize == printing.DMPAPER_USER:
            out.PaperSize = 1
        out.Copies = min(out.Copies, self.max_copies)
        return 1

    def SetPrinter(self, *args):
        raise AssertionError("printer settings must not be changed")

    def EnumJobs(self, handle, first, count, level):
        return [{"JobId": j} for j in self.jobs.pop(0)] if self.jobs else []


@pytest.fixture
def fake(monkeypatch):
    def install(**kwargs):
        win32 = FakeWin32Print(**kwargs)
        monkeypatch.setattr(printing, "win32print", win32)
        return win32
    return install


def test_job_settings_custom_paper(fake):
    win32 = fake()
    dm = printing.job_settings("Plotter", (610, 841.4), copies=2)
    assert (dm.PaperSize, dm.PaperWidth, dm.PaperLength, dm.Orientation) == (256, 6100, 8414, 1)
    assert dm.Fields & 0xF == 0xF
    assert (dm.Copies, dm.Collate) == (2, 1)
    assert win32.open_handles == 0
    assert win32.devmode.PaperSize == 1          # the printer's own settings are untouched


@pytest.mark.parametrize("paper, paper_id", [("A4", 9), ("A3", 8)])
def test_job_settings_named_paper(fake, paper, paper_id):
    fake()
    dm = printing.job_settings("Office", paper)
    assert dm.PaperSize == paper_id
    assert not dm.Fields & (printing.DM_PAPERWIDTH | printing.DM_PAPERLENGTH)


def test_job_settings_keeps_paper_when_none_given(fake):
    fake()
    assert printing.job_settings("Office").PaperSize == 1


def test_job_settings_start_from_this_users_preferences(fake):
    win32 = fake()
    win32.user_devmode = FakeDevMode(PaperSize=5, Duplex=printing.DMDUP_VERTICAL, Fields=printing.DM_DUPLEX)
    assert printing.job_settings("Office").PaperSize == 5
    assert printing.default_duplex("Office")
    win32.user_devmode = None
    assert printing.job_settings("Office").PaperSize == 1
    assert not printing.default_duplex("Office")


def test_job_settings_duplex(fake):
    fake()
    dm = printing.job_settings("Office", "A4", duplex=True)
    assert dm.Duplex == printing.DMDUP_VERTICAL and dm.Fields & printing.DM_DUPLEX
    assert printing.job_settings("Office", duplex=False).Duplex == printing.DMDUP_SIMPLEX
    assert not printing.job_settings("Office").Fields & printing.DM_DUPLEX   # as the printer is set


def test_duplex_capability_and_default(fake):
    win32 = fake()
    win32.capabilities[printing.DC_DUPLEX] = 1
    assert printing.supports_duplex("Office")
    win32.capabilities[printing.DC_DUPLEX] = 0
    assert not printing.supports_duplex("Office")
    assert not printing.default_duplex("Office")
    win32.devmode.Duplex, win32.devmode.Fields = 2, printing.DM_DUPLEX
    assert printing.default_duplex("Office")


def test_detect_printers(fake):
    win32 = fake()
    a4_only = {printing.DC_PAPERS: [9, 11], printing.DC_PAPERSIZE: [{"x": 2100, "y": 2970}]}
    with_a3 = {printing.DC_PAPERS: [9, 8], printing.DC_PAPERSIZE: [(2100, 2970), (2970, 4200)]}
    plotter = {printing.DC_PAPERS: [9, 8, 256], printing.DC_PAPERSIZE: [(9140, 15000)]}
    win32.per_printer = {"HP LaserJet M404": a4_only, "Canon iR-ADV C3530": with_a3,
                         "Big Printer": plotter, "Microsoft Print to PDF": with_a3}
    win32.default = "Microsoft Print to PDF"
    names = ["Big Printer", "Canon iR-ADV C3530", "HP LaserJet M404", "Microsoft Print to PDF"]
    assert printing.detect_printers(names) == {
        "a4": "HP LaserJet M404", "a3": "Canon iR-ADV C3530", "plotter": "Big Printer"}
    win32.default = "Canon iR-ADV C3530"     # the default printer also takes A3
    assert printing.detect_printers(names) == {
        "a4": "Canon iR-ADV C3530", "a3": "Canon iR-ADV C3530", "plotter": "Big Printer"}
    # recognised by name even if the driver does not tell its paper sizes
    win32.per_printer["HP DesignJet T650"] = {}
    assert printing.detect_printers(["HP DesignJet T650", "HP LaserJet M404"])["plotter"] == \
        "HP DesignJet T650"
    assert printing.detect_printers([]) == {"a4": "", "a3": "", "plotter": ""}


def test_job_settings_rejected_size(fake):
    fake(reject_custom=True)
    with pytest.raises(printing.PrintError, match="610×1234"):
        printing.job_settings("PDF", (610, 1234))


def test_plot_asks_for_its_exact_size(fake):
    """Even a plot of a standard size, so its paper can be made longer for the margins."""
    fake()
    dm = printing.job_settings("Plotter", (297, 420))
    assert (dm.PaperSize, dm.PaperWidth, dm.PaperLength) == (printing.DMPAPER_USER, 2970, 4200)


def test_plot_on_a_driver_without_custom_sizes_gets_the_standard_paper(fake):
    fake(reject_custom=True)
    assert printing.job_settings("PDF", (297, 420)).PaperSize == 8


def test_driver_may_swap_a_custom_size_for_its_paper_of_that_size(fake):
    win32 = fake()
    win32.swaps = {(2970, 4200): 8}
    win32.capabilities.update({printing.DC_PAPERS: [9, 8], printing.DC_PAPERSIZE: [
        {"x": 2100, "y": 2970}, {"x": 2970, "y": 4200}]})
    assert printing.job_settings("Plotter", (297, 420)).PaperSize == 8
    win32.swaps = {(6100, 9140): 8}           # a different size is not accepted
    with pytest.raises(printing.PrintError):
        printing.job_settings("Plotter", (610, 914))


def test_paper_longer_than_windows_allows(fake):
    fake()
    with pytest.raises(printing.PrintError, match="3276"):
        printing.job_settings("Plotter", (914, 4000))
    assert printing.job_settings("Plotter", (914, 3276)).PaperLength == 32760


def test_settings_the_driver_no_longer_takes_are_replaced_by_its_own(fake, monkeypatch):
    win32 = fake()
    win32.devmode_size = 220 + 512
    monkeypatch.setattr(printing, "pywintypes", type("P", (), {"DEVMODEType": staticmethod(
        lambda extra: FakeDevMode(DriverExtra=extra))}))
    real = win32.DocumentProperties

    def refuse_old_settings(hwnd, handle, name, out, inp, mode):
        if mode == 0:
            return real(hwnd, handle, name, out, inp, mode)
        if inp is not None and inp.DriverExtra < 512:
            raise RuntimeError("DocumentProperties: invalid settings")
        if inp is None:   # the driver's own settings
            out.__dict__.update(FakeDevMode(DriverExtra=out.DriverExtra, PaperSize=9).__dict__)
            return 1
        return real(hwnd, handle, name, out, inp, mode)
    win32.DocumentProperties = refuse_old_settings
    assert printing._driver_sized("h", "Plotter", FakeDevMode()).DriverExtra == 512


def test_settings_from_an_older_driver_get_a_buffer_of_the_size_it_writes(fake, monkeypatch):
    win32 = fake()
    win32.devmode_size = 220 + 512            # the driver keeps 512 bytes of its own now
    monkeypatch.setattr(printing, "pywintypes", type("P", (), {"DEVMODEType": staticmethod(
        lambda extra: FakeDevMode(DriverExtra=extra))}))
    dm = printing.job_settings("Plotter", (610, 841))
    assert dm.DriverExtra == 512 and dm.PaperLength == 8410


def test_place_page_real_size_centred():
    # A4 page on A4 paper at 600 dpi with 4.2 mm margins: fits after shrinking a little
    x, y, w, h, rotate = place_page((210, 297), (4960, 7016), (100, 100), (4760, 6816), (600, 600))
    assert rotate == 0
    assert (w, h) == (4760, 6732)                 # shrunk to the printable width
    assert (x, y) == (0, 42)                      # centred on the sheet


def test_place_page_rotates_landscape_onto_portrait_roll():
    # 841×594 drawing on 610×841 roll paper, 300 dpi, 5 mm margins
    sheet, offset, area = (7205, 9933), (59, 59), (7087, 9815)
    x, y, w, h, rotate = place_page((841, 594), sheet, offset, area, (300, 300))
    assert rotate == 1
    assert h == 9815                              # long side fills the printable length
    assert w == round(594 / 25.4 * 300 * 9815 / (841 / 25.4 * 300))   # same scale across
    assert x == (sheet[0] - w) // 2 - offset[0] and y == 0


def test_place_page_real_size_when_paper_has_room():
    # the same drawing on roll paper made 10 mm longer for the margins prints 1:1
    sheet, offset, area = (7205, 10051), (59, 59), (7087, 9933)
    x, y, w, h, rotate = place_page((841, 594), sheet, offset, area, (300, 300))
    assert (w, h) == (round(594 / 25.4 * 300), round(841 / 25.4 * 300))


def test_place_page_with_different_resolution_across_and_down():
    # A4 page on A4 paper at 1200×600 dpi: more pixels across than down, still portrait
    sheet = (round(210 / 25.4 * 1200), round(297 / 25.4 * 600))
    x, y, w, h, rotate = place_page((210, 297), sheet, (0, 0), sheet, (1200, 600))
    assert rotate == 0 and (w, h) == sheet


def test_place_page_small_page_not_enlarged():
    x, y, w, h, rotate = place_page((100, 150), (2480, 3508), (0, 0), (2480, 3508), (300, 300))
    assert (w, h) == (1181, 1772) and rotate == 0
    assert (x, y) == ((2480 - 1181) // 2, (3508 - 1772) // 2)


def test_every_plot_gets_room_for_the_margins(fake, monkeypatch, make_pdf):
    fake()
    asked = []
    monkeypatch.setattr(printing, "_room_for_margins", lambda printer, paper, copies, dm: asked.append(paper) or dm)
    monkeypatch.setattr(printing.DirectBackend, "_print", lambda self, *args: True)
    path = make_pdf("a3.pdf", [(280, 420)])
    printing.DirectBackend().print_file("Plotter", path, paper=(297, 420))
    printing.DirectBackend().print_file("Office", path, paper="A3")
    assert asked == [(297, 420)]


def test_driver_copies(fake):
    win32 = fake()
    assert printing._driver_copies("Office", 2, collate=True)
    assert not printing._driver_copies("Office", 150, collate=False)       # more than the driver does
    win32.capabilities[printing.DC_COLLATE] = 0
    assert not printing._driver_copies("Office", 2, collate=True)         # would come out 1,1,2,2
    assert printing._driver_copies("Office", 2, collate=False)
    del win32.capabilities[printing.DC_COPIES]
    assert not printing._driver_copies("Office", 2, collate=False)        # unknown: print it again


def test_wait_for_new_job(fake, monkeypatch):
    monkeypatch.setattr(printing.time, "sleep", lambda s: None)
    win32 = fake()
    win32.jobs = [[1], [1], [1, 2]]
    assert printing._wait_for_new_job("P", {1}, timeout=5)
    assert not printing._wait_for_new_job("P", None)


def test_backend_choice(monkeypatch):
    monkeypatch.setattr(printing, "find_sumatra", lambda: "C:/S/SumatraPDF.exe")
    monkeypatch.setattr(printing, "find_adobe", lambda: None)
    assert isinstance(printing.find_backend(), printing.DirectBackend)
    assert isinstance(printing.find_backend("sumatra"), printing.SumatraBackend)
    assert isinstance(printing.find_backend("adobe"), printing.DirectBackend)   # not installed


def test_sumatra_command_line(monkeypatch):
    calls = []

    class Result:
        returncode = 0
    monkeypatch.setattr(printing.subprocess, "run", lambda args, **kw: calls.append(args) or Result())
    backend = printing.SumatraBackend("SumatraPDF.exe")
    assert backend.print_file("HP DesignJet", "C:/a b.pdf", 3)
    assert backend.print_file("Office", "C:/a.pdf", 1, paper="A3")
    assert backend.print_file("Office", "C:/a.pdf", 2, paper="A4")
    assert backend.print_file("Office", "C:/a.pdf", 1, paper=(610, 841))   # no custom sizes
    assert backend.print_file("Office", "C:/a.pdf", 1, "A4", duplex=True)
    assert backend.print_file("Office", "C:/a.pdf", 1, duplex=False)
    assert calls == [
        ["SumatraPDF.exe", "-print-to", "HP DesignJet", "-silent", "-print-settings", "3x", "C:/a b.pdf"],
        ["SumatraPDF.exe", "-print-to", "Office", "-silent", "-print-settings", "paper=A3", "C:/a.pdf"],
        ["SumatraPDF.exe", "-print-to", "Office", "-silent", "-print-settings", "2x,paper=A4", "C:/a.pdf"],
        ["SumatraPDF.exe", "-print-to", "Office", "-silent", "C:/a.pdf"],
        ["SumatraPDF.exe", "-print-to", "Office", "-silent", "-print-settings", "paper=A4,duplexlong", "C:/a.pdf"],
        ["SumatraPDF.exe", "-print-to", "Office", "-silent", "-print-settings", "simplex", "C:/a.pdf"],
    ]
    Result.returncode = 1
    with pytest.raises(printing.PrintError):
        backend.print_file("HP DesignJet", "C:/a b.pdf")
