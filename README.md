# SPrint

Windows application for printing many PDF files of different page sizes on several printers and a plotter in one go.

[Hrvatski ↓](#sprint--hrvatski)

## Installation

No Python or anything else is needed.

1. Download **`SPrint-Setup-<version>.exe`** from [Releases](../../releases)
   (or from the artifacts of the latest successful [*Build* run](../../actions/workflows/build.yml)).
2. Run it. *Install for me only* needs no administrator rights; *Install for all users* puts SPrint in Program Files.
3. Start SPrint from the Start menu. Optionally the installer adds a desktop icon and an entry in Explorer's
   *Send to* menu (select PDF files or folders → right-click → *Send to* → SPrint).

Updating: run the newer installer, it replaces the old version and keeps your settings.
Uninstalling: *Settings → Apps → SPrint*.

Windows may warn that the installer comes from an unknown publisher (it is not code-signed):
click *More info* → *Run anyway*.

There is also a **`SPrint-portable-<version>.zip`**: unpack it anywhere (e.g. a USB stick) and run `SPrint.exe`.

## How it works

SPrint reads the size of every page and sends it to the right device:

| Page size | Goes to |
|---|---|
| A4 (either orientation) | **A4 printer** — button *Print A4* |
| A3 (either orientation) | **A3 printer** — button *Print A3*, or the plotter with the *A3* roll option |
| anything else | **plotter**, on the narrowest roll the page fits on: 297, 420, 610 or 914 mm |

Sizes within 3 mm of A4/A3 count as A4/A3 (scanners and PDF printers rarely produce exact sizes).
The visible page area is used (crop box, `/Rotate`, `/UserUnit`), i.e. what a PDF viewer shows and prints.

For the plotter you choose the roll that is loaded. SPrint groups that roll's pages by length and prints each group
on paper of *roll width × page length* (a few millimetres longer if the plotter needs room for its margins, so drawings
print at 1:1). The paper size travels with each print job, so the plotter's own settings are never changed.

## Features

- Drag & drop of PDF files **and folders**; files can also be passed on the command line,
  so SPrint works from Explorer's *Send to* menu or by dropping files onto `SPrint.exe`
- Summary of page sizes per file and in total, with warnings for unreadable files and pages wider than the widest roll
- Automatic paper length on the plotter, printed groups marked with ✓
- Number of copies (plotter copies are collated per paper length)
- Prints by itself with the built-in PDF renderer (PDFium, the engine of Google Chrome) — no Adobe Reader needed
- Landscape pages are turned to fit the paper; pages larger than the printable area are shrunk, smaller ones are not enlarged
- Only the needed pages of a file are printed; CAD layer visibility is kept
- Printing runs in the background — the window never freezes and can show progress; closing while printing
  stops after the current file
- Remembers the selected printers and language
- English / Croatian interface (Croatian by default on Croatian Windows)
- Log file for troubleshooting

## How SPrint prints

SPrint renders the pages itself (PDFium) straight into a Windows print job that carries its own paper size, copies and
orientation. Your printing preferences (tray, duplex, colour, quality) still apply; nothing in the printer's settings is
changed. This is what makes custom plot lengths possible.

The A4/A3 printers can use an external PDF program instead — add to `%APPDATA%\SPrint\settings.ini`:

```ini
[printing]
program = sumatra
```

`sumatra` uses [SumatraPDF](https://www.sumatrapdfreader.org/) (installed, or `SumatraPDF.exe` next to `SPrint.exe`),
`adobe` uses Adobe Acrobat / Acrobat Reader. The plotter always prints directly, because neither program can set a
custom paper length (SumatraPDF 3.6 ignores it, Adobe Reader takes it only from the printer's default settings).

## From source

### As a Python script

Python 3.10 or newer:

```bash
pip install -r requirements.txt
```

```bash
python PRINT.py
```

### Building the installer

```bash
pip install pyinstaller
```

```bash
pyinstaller SPrint.spec
```

That gives the program folder `dist\SPrint\`. Then, with [Inno Setup 6.5+](https://jrsoftware.org/isinfo.php):

```bash
iscc /DAppVersion=2.0.0 installer\SPrint.iss
```

which writes `dist\SPrint-Setup-2.0.0.exe`. For Croatian installer texts put Inno Setup's unofficial
[`Croatian.isl`](https://jrsoftware.org/files/istrans/) into `installer\` first.

GitHub Actions does all of this on every push: it runs the tests on Windows, builds the installer, installs it,
starts SPrint, uninstalls it and keeps the installer and the portable zip as artifacts. Pushing a tag such as
`v2.0.0` publishes them as a release.

### Tests

```bash
pip install pytest
```

```bash
python -m pytest
```

The tests of the page-size logic run on any OS; the printer tests run only on Windows.

## Files

| Location | Content |
|---|---|
| `%APPDATA%\SPrint\settings.ini` | selected printers, language, optional print program (older `printer_choices.ini` is imported once) |
| `%APPDATA%\SPrint\sprint.log` | log for troubleshooting |
| `%TEMP%\SPrint_*` | temporary PDFs with extracted pages, deleted on exit |

## Project structure

| File | Purpose |
|---|---|
| `PRINT.py` | entry point (`python PRINT.py`, PyInstaller) |
| `sprint/app.py` | window and user interaction |
| `sprint/jobs.py` | print and plot workflows |
| `sprint/paper.py` | page-size classification, rolls, grouping (pure logic) |
| `sprint/pdfdocs.py` | reading page sizes, extracting pages (`pypdf`) |
| `sprint/printing.py` | Windows printers, print jobs with their own paper size, external PDF programs |
| `sprint/settings.py` | settings file |
| `sprint/i18n.py` | English and Croatian texts |
| `tests/` | automated tests |
| `SPrint.spec`, `installer/SPrint.iss` | program build (PyInstaller) and installer (Inno Setup) |

## Requirements

- Windows 10 / 11
- Python 3.10+ (only to run from source)
- Printers installed in Windows

## Author

Filip Kozina

---

# SPrint — Hrvatski

Windows program za ispis više PDF datoteka različitih veličina stranica na više printera i ploter, u jednom koraku.

## Instalacija

Python ni ništa drugo nije potrebno.

1. Preuzmite **`SPrint-Setup-<verzija>.exe`** iz [Releases](../../releases)
   (ili iz artefakata zadnjeg uspješnog [*Build* pokretanja](../../actions/workflows/build.yml)).
2. Pokrenite ga. *Instalacija samo za mene* ne traži administratorska prava; *za sve korisnike* instalira u Program Files.
3. SPrint pokrenite iz izbornika Start. Instalacija po želji dodaje ikonu na radnu površinu i stavku u izborniku
   *Pošalji u* u Exploreru (označite PDF datoteke ili mape → desni klik → *Pošalji u* → SPrint).

Nadogradnja: pokrenite noviju instalaciju, ona zamjenjuje staru verziju i zadržava postavke.
Deinstalacija: *Postavke → Aplikacije → SPrint*.

Windows može upozoriti da je instalacija od nepoznatog izdavača (nije digitalno potpisana):
kliknite *Više informacija* → *Svejedno pokreni*.

Postoji i **`SPrint-portable-<verzija>.zip`**: raspakirajte ga bilo gdje (npr. na USB stick) i pokrenite `SPrint.exe`.

## Kako radi

SPrint očita veličinu svake stranice i šalje je na odgovarajući uređaj:

| Veličina stranice | Ide na |
|---|---|
| A4 (bilo koja orijentacija) | **A4 printer** — gumb *Print A4* |
| A3 (bilo koja orijentacija) | **A3 printer** — gumb *Print A3*, ili ploter s opcijom role *A3* |
| sve ostalo | **ploter**, na najužu rolu na koju stranica stane: 297, 420, 610 ili 914 mm |

Veličine do 3 mm odstupanja od A4/A3 smatraju se A4/A3 (skeneri i PDF printeri rijetko daju točne mjere).
Koristi se vidljivi dio stranice (crop box, `/Rotate`, `/UserUnit`), tj. ono što PDF preglednik prikazuje i ispisuje.

Za ploter odaberete rolu koja je umetnuta. SPrint grupira stranice te role po dužini i svaku grupu ispisuje na papir
*širina role × dužina stranice* (nekoliko milimetara dulji ako ploteru treba mjesta za margine, tako da se crteži ispisuju
u mjerilu 1:1). Veličina papira putuje sa svakim poslom ispisa, pa se postavke samog plotera nikad ne mijenjaju.

## Značajke

- Drag & drop PDF datoteka **i mapa**; datoteke se mogu zadati i u naredbenom retku,
  pa SPrint radi iz izbornika *Pošalji u* (Send to) u Exploreru ili ispuštanjem datoteka na `SPrint.exe`
- Pregled veličina stranica po datoteci i ukupno, s upozorenjima za nečitljive datoteke i stranice šire od najšire role
- Automatska dužina papira na ploteru, isplotane grupe označene s ✓
- Broj kopija (kopije na ploteru slažu se po dužini papira)
- Ispisuje sam, ugrađenim PDF prikazom (PDFium, pogon Google Chromea) — Adobe Reader nije potreban
- Položene stranice okreću se prema papiru; stranice veće od područja ispisa se smanjuju, manje se ne povećavaju
- Ispisuju se samo potrebne stranice datoteke; vidljivost CAD slojeva ostaje očuvana
- Ispis radi u pozadini — prozor se ne zamrzava i prikazuje napredak; zatvaranje tijekom ispisa
  zaustavlja ispis nakon trenutne datoteke
- Pamćenje odabranih printera i jezika
- Sučelje na engleskom / hrvatskom (na hrvatskim Windowsima hrvatski je zadani)
- Datoteka zapisa (log) za otklanjanje problema

## Kako SPrint ispisuje

SPrint sam iscrtava stranice (PDFium) izravno u Windows posao ispisa koji nosi vlastitu veličinu papira, broj kopija i
orijentaciju. Vaše postavke ispisa (ladica, obostrani ispis, boja, kvaliteta) i dalje vrijede; postavke printera se ne
mijenjaju. Upravo to omogućuje proizvoljne dužine papira na ploteru.

A4/A3 printeri mogu umjesto toga koristiti vanjski PDF program — dodajte u `%APPDATA%\SPrint\settings.ini`:

```ini
[printing]
program = sumatra
```

`sumatra` koristi [SumatraPDF](https://www.sumatrapdfreader.org/) (instaliran ili `SumatraPDF.exe` pokraj `SPrint.exe`),
`adobe` koristi Adobe Acrobat / Acrobat Reader. Ploter uvijek ispisuje izravno jer nijedan od tih programa ne može zadati
proizvoljnu dužinu papira (SumatraPDF 3.6 je zanemaruje, Adobe Reader je uzima samo iz zadanih postavki printera).

## Iz izvornog koda

### Kao Python skripta

Python 3.10 ili noviji:

```bash
pip install -r requirements.txt
```

```bash
python PRINT.py
```

### Izrada instalacije

```bash
pip install pyinstaller
```

```bash
pyinstaller SPrint.spec
```

To daje mapu programa `dist\SPrint\`. Zatim, uz [Inno Setup 6.5+](https://jrsoftware.org/isinfo.php):

```bash
iscc /DAppVersion=2.0.0 installer\SPrint.iss
```

što stvara `dist\SPrint-Setup-2.0.0.exe`. Za hrvatske tekstove instalacije prvo stavite neslužbeni
[`Croatian.isl`](https://jrsoftware.org/files/istrans/) iz Inno Setupa u `installer\`.

GitHub Actions sve to radi pri svakom pushu: pokreće testove na Windowsima, izrađuje instalaciju, instalira je,
pokreće SPrint, deinstalira ga i čuva instalaciju i prijenosni zip kao artefakte. Push taga poput `v2.0.0`
objavljuje ih kao release.

### Testovi

```bash
pip install pytest
```

```bash
python -m pytest
```

Testovi logike veličina stranica rade na svakom OS-u; testovi printera samo na Windowsima.

## Datoteke

| Lokacija | Sadržaj |
|---|---|
| `%APPDATA%\SPrint\settings.ini` | odabrani printeri, jezik, po želji program za ispis (stari `printer_choices.ini` uvozi se jednom) |
| `%APPDATA%\SPrint\sprint.log` | zapis za otklanjanje problema |
| `%TEMP%\SPrint_*` | privremeni PDF-ovi s izdvojenim stranicama, brišu se pri zatvaranju |

## Zahtjevi

- Windows 10 / 11
- Python 3.10+ (samo za pokretanje iz izvornog koda)
- Instalirani printeri u Windows sustavu

## Autor

Filip Kozina
