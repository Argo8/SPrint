import configparser

from sprint import settings as settings_mod
from sprint.settings import Settings


def test_round_trip_with_awkward_names(tmp_path):
    path = str(tmp_path / "sub" / "settings.ini")
    Settings("Pisač 100% A4", "\\\\server\\Kopirka A3", "HP DesignJet T830", "hr").save(path)
    loaded = Settings.load(path)
    assert loaded == Settings("Pisač 100% A4", "\\\\server\\Kopirka A3", "HP DesignJet T830", "hr")


def test_missing_file_gives_defaults(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(settings_mod, "app_dir", lambda: str(tmp_path))
    assert Settings.load(str(tmp_path / "none.ini")) == Settings()


def test_reads_legacy_printer_choices(tmp_path, monkeypatch):
    # Written exactly like the old PRINT.py did (default ConfigParser lower-cases keys).
    legacy = configparser.ConfigParser()
    legacy["PRINTERS"] = {"A4": "Old A4", "A3": "Old A3", "Large": "Old plotter"}
    with open(tmp_path / "printer_choices.ini", "w") as fh:
        legacy.write(fh)
    monkeypatch.chdir(tmp_path)
    loaded = Settings.load(str(tmp_path / "new" / "settings.ini"))
    assert (loaded.printer_a4, loaded.printer_a3, loaded.printer_plotter) == \
        ("Old A4", "Old A3", "Old plotter")
    assert loaded.language == ""


def test_new_file_wins_over_legacy(tmp_path, monkeypatch):
    (tmp_path / "printer_choices.ini").write_text("[PRINTERS]\na4 = Old\na3 = Old\nlarge = Old\n")
    monkeypatch.chdir(tmp_path)
    path = str(tmp_path / "settings.ini")
    Settings(printer_a4="New").save(path)
    assert Settings.load(path).printer_a4 == "New"
