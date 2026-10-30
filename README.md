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

Drop PDF files (or folders) into the window. SPrint reads the size of every page and shows where it goes:

| Page size | Goes to |
|---|---|
| A4 (either orientation) | **A4 printer** |
| A3 (either orientation) | **A3 printer**, or the plotter's 297 mm roll (*Plotter (roll 297)* in the A3 card) |
| anything else | **plotter**, on the narrowest roll the page fits on: 297, 420, 610 or 914 mm |

Sizes within 3 mm of A4/A3 count as A4/A3 (scanners and PDF printers rarely produce exact sizes).
The visible page area is used (crop box, `/Rotate`, `/UserUnit`), i.e. what a PDF viewer shows and prints.

One **Print** button sends everything that is switched on: the A4 pages, the A3 pages and the plotter pages for the
roll that is in the plotter now. The plotter holds one roll at a time, so its pages are printed roll by roll:

1. In the *Plotter* card pick the roll that is loaded. Each roll shows how many pages need it.
2. Print. The roll gets a ✓ and SPrint says which roll is next (*Change the roll*).
3. Load that roll, click *Roll … is loaded*, print again — until everything is printed.

Printed pages are marked done, so printing again sends only what is left (*Print again* / *Plot again* resend a part).
On the plotter each page is printed on paper of *roll width × page length* (a few millimetres longer if the plotter
needs room for its margins, so drawings print at 1:1). The paper size travels with each print job, so the plotter's
own settings are never changed. SPrint remembers which roll is in the plotter.

## Features

- Drag & drop of PDF files **and folders**; *Send to → SPrint* in Explorer and files dropped onto `SPrint.exe` work too,
  and go into the SPrint window that is already open
- Click a file to see its pages with previews and where each one goes; click a page to leave it out
- Per printer: on/off switch, printer choice and two-sided printing (when the printer can do it)
- Paper length needed on the plotter, per roll, before you print
- Printers are found automatically the first time (and on request in *Settings*)
- Number of copies (collated)
- Prints by itself with the built-in PDF renderer (PDFium, the engine of Google Chrome), always as the PDF is —
  no Adobe Reader needed
- Landscape pages are turned to fit the paper; pages larger than the printable area are shrunk, smaller ones are not enlarged
- Only the needed pages of a file are printed; CAD layer visibility is kept
- Printing runs in the background; the window shows progress per printer
- Light and dark theme (or like Windows), English / Croatian interface
- Keyboard: *Ctrl+O* adds files, *Ctrl+P* prints, *Esc* closes the settings
- Log for troubleshooting (*Settings → Log*)

## How SPrint prints

SPrint renders the pages itself (PDFium) straight into a Windows print job that carries its own paper size, copies,
orientation and (if you switch it) two-sided printing. Your printing preferences (tray, colour, quality) still apply;
nothing in the printer's settings is changed. This is what makes custom plot lengths possible. *Settings → Windows
printer settings* opens the driver's own dialog for everything else.

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
iscc /DAppVersion=2.1.0 installer\SPrint.iss
```

which writes `dist\SPrint-Setup-2.1.0.exe`. For Croatian installer texts put Inno Setup's unofficial
[`Croatian.isl`](https://jrsoftware.org/files/istrans/) into `installer\` first.

GitHub Actions does all of this on every push: it runs the tests on Windows, builds the installer, installs it,
starts SPrint, uninstalls it and keeps the installer and the portable zip as artifacts.

To publish a release, raise `__version__` in `sprint/__init__.py`. Then either open *Actions → Build → Run
workflow* on GitHub, tick *Publish a release* and run it, or push a matching tag such as `v2.1.0` (the build
stops if the tag differs). The release gets the installer and the portable zip once the build passes; the text
on the release page comes from `.github/release-notes.md`.

### Tests

```bash
pip install pytest pytest-qt
```

```bash
python -m pytest
```

The logic and window tests run on any OS (the window without a screen, via Qt's offscreen mode);
the printer tests run only on Windows.

## Files

| Location | Content |
|---|---|
| `%APPDATA%\SPrint\settings.ini` | printers, the roll in the plotter, A3 on the plotter, language, theme, optional print program (older `printer_choices.ini` is imported once) |
| `%APPDATA%\SPrint\sprint.log` | log for troubleshooting |
| `%TEMP%\SPrint_*` | temporary PDFs with extracted pages, deleted on exit |

## Project structure

| File | Purpose |
|---|---|
| `PRINT.py` | entry point (`python PRINT.py`, PyInstaller) |
| `sprint/app.py` | start-up: logging, Qt, the window, `--smoke-test` |
| `sprint/ui/` | the window: `main_window.py`, `controller.py` (state and background work), `widgets.py`, `theme.py`, `icons.py` |
| `sprint/plan.py` | what goes where: pages per printer and roll, skipped and printed pages (pure logic) |
| `sprint/paper.py` | page sizes and rolls (pure logic) |
| `sprint/jobs.py` | print jobs |
| `sprint/pdfdocs.py` | reading page sizes, extracting pages (`pypdf`), page previews |
| `sprint/printing.py` | Windows printers, print jobs with their own paper size, external PDF programs |
| `sprint/single_instance.py` | one window: a second start hands its files over |
| `sprint/settings.py` | settings file |
| `sprint/i18n.py` | English and Croatian texts |
| `tests/` | automated tests |
| `SPrint.spec`, `installer/SPrint.iss` | program build (PyInstaller) and installer (Inno Setup) |

Icons follow the [Lucide](https://lucide.dev) set (ISC licence).

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
2. Pokrenite ga. Instalacija samo za sebe ne traži administratorska prava; instalacija za sve korisnike stavlja SPrint u Program Files.
3. SPrint pokrenite iz izbornika Start. Instalacija po želji dodaje ikonu na radnu površinu i stavku u izborniku
   *Pošalji u* u Exploreru (označite PDF datoteke ili mape → desni klik → *Pošalji u* → SPrint).

Nadogradnja: pokrenite noviju instalaciju, ona zamjenjuje staru verziju i zadržava postavke.
Deinstalacija: *Postavke → Aplikacije → SPrint*.

Windows može upozoriti da je instalacija od nepoznatog izdavača (nije digitalno potpisana):
kliknite *Više informacija* → *Svejedno pokreni*.

Postoji i **`SPrint-portable-<verzija>.zip`**: raspakirajte ga bilo gdje (npr. na USB stick) i pokrenite `SPrint.exe`.

## Kako radi

Povucite PDF datoteke (ili mape) u prozor. SPrint očita veličinu svake stranice i pokaže kamo ide:

| Veličina stranice | Ide na |
|---|---|
| A4 (bilo koja orijentacija) | **A4 pisač** |
| A3 (bilo koja orijentacija) | **A3 pisač**, ili ploter na rolu 297 mm (*Ploter (rola 297)* u kartici A3) |
| sve ostalo | **ploter**, na najužu rolu na koju stranica stane: 297, 420, 610 ili 914 mm |

Veličine do 3 mm odstupanja od A4/A3 smatraju se A4/A3 (skeneri i PDF pisači rijetko daju točne mjere).
Koristi se vidljivi dio stranice (crop box, `/Rotate`, `/UserUnit`), tj. ono što PDF preglednik prikazuje i ispisuje.

Jedan gumb **Ispiši** šalje sve što je uključeno: A4 stranice, A3 stranice i stranice za ploter za rolu koja je sada
u ploteru. U ploteru je jedna rola, pa se ploter ispisuje rolu po rolu:

1. U kartici *Ploter* odaberite rolu koja je stavljena. Uz svaku rolu piše koliko stranica je treba.
2. Ispišite. Rola dobije ✓ i SPrint kaže koja je sljedeća (*Promijeni rolu*).
3. Stavite tu rolu, kliknite *Rola … je stavljena* i ispišite opet — dok sve ne bude ispisano.

Ispisane stranice označe se kao gotove, pa ponovni ispis šalje samo ono što je ostalo (*Ispiši ponovno* / *Plotaj
ponovno* šalju dio opet). Na ploteru se svaka stranica ispisuje na papir *širina role × dužina stranice* (nekoliko
milimetara dulji ako ploteru treba mjesta za margine, tako da se crteži ispisuju u mjerilu 1:1). Veličina papira
putuje sa svakim poslom ispisa, pa se postavke samog plotera nikad ne mijenjaju. SPrint pamti koja je rola u ploteru.

## Značajke

- Drag & drop PDF datoteka **i mapa**; rade i *Pošalji u → SPrint* u Exploreru i ispuštanje datoteka na `SPrint.exe`,
  a datoteke idu u već otvoren prozor SPrinta
- Klik na datoteku pokazuje njezine stranice sa sličicama i kamo koja ide; klik na stranicu je preskače
- Za svaki pisač: prekidač, odabir pisača i obostrani ispis (ako ga pisač podržava)
- Dužina papira potrebna na ploteru, po roli, prije ispisa
- Pisači se prvi put prepoznaju sami (i na zahtjev u *Postavkama*)
- Broj kopija (složene)
- Ispisuje sam, ugrađenim PDF prikazom (PDFium, pogon Google Chromea), uvijek onako kako je u PDF-u —
  Adobe Reader nije potreban
- Položene stranice okreću se prema papiru; stranice veće od područja ispisa se smanjuju, manje se ne povećavaju
- Ispisuju se samo potrebne stranice datoteke; vidljivost CAD slojeva ostaje očuvana
- Ispis radi u pozadini; prozor pokazuje napredak za svaki pisač
- Svijetla i tamna tema (ili kao Windows), sučelje na hrvatskom / engleskom
- Tipkovnica: *Ctrl+O* dodaje datoteke, *Ctrl+P* ispisuje, *Esc* zatvara postavke
- Dnevnik za otklanjanje problema (*Postavke → Dnevnik*)

## Kako SPrint ispisuje

SPrint sam iscrtava stranice (PDFium) izravno u Windows posao ispisa koji nosi vlastitu veličinu papira, broj kopija,
orijentaciju i (ako ga uključite) obostrani ispis. Vaše postavke ispisa (ladica, boja, kvaliteta) i dalje vrijede;
postavke pisača se ne mijenjaju. Upravo to omogućuje proizvoljne dužine papira na ploteru. *Postavke → Windows postavke
pisača* otvara prozor upravljačkog programa za sve ostalo.

A4/A3 pisači mogu umjesto toga koristiti vanjski PDF program — dodajte u `%APPDATA%\SPrint\settings.ini`:

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
iscc /DAppVersion=2.1.0 installer\SPrint.iss
```

što stvara `dist\SPrint-Setup-2.1.0.exe`. Za hrvatske tekstove instalacije prvo stavite neslužbeni
[`Croatian.isl`](https://jrsoftware.org/files/istrans/) iz Inno Setupa u `installer\`.

GitHub Actions sve to radi pri svakom pushu: pokreće testove na Windowsima, izrađuje instalaciju, instalira je,
pokreće SPrint, deinstalira ga i čuva instalaciju i prijenosni zip kao artefakte.

Za objavu nove verzije povećajte `__version__` u `sprint/__init__.py`. Zatim na GitHubu otvorite *Actions →
Build → Run workflow*, označite *Publish a release* i pokrenite, ili pushajte isti tag, npr. `v2.1.0` (build
staje ako se tag razlikuje). Release dobiva instalaciju i prijenosni zip kad build prođe; tekst stranice
releasea nalazi se u `.github/release-notes.md`.

### Testovi

```bash
pip install pytest pytest-qt
```

```bash
python -m pytest
```

Testovi logike i prozora rade na svakom OS-u (prozor bez ekrana, u Qt-ovom offscreen načinu);
testovi pisača samo na Windowsima.

## Datoteke

| Lokacija | Sadržaj |
|---|---|
| `%APPDATA%\SPrint\settings.ini` | pisači, rola u ploteru, A3 na ploter, jezik, tema, po želji program za ispis (stari `printer_choices.ini` uvozi se jednom) |
| `%APPDATA%\SPrint\sprint.log` | zapis za otklanjanje problema |
| `%TEMP%\SPrint_*` | privremeni PDF-ovi s izdvojenim stranicama, brišu se pri zatvaranju |

## Zahtjevi

- Windows 10 / 11
- Python 3.10+ (samo za pokretanje iz izvornog koda)
- Instalirani pisači u Windows sustavu

## Autor

Filip Kozina
