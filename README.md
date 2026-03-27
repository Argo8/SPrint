# SPrint

Windows application for automatically printing multiple PDF files of different sizes on multiple printers and plotters.

## Description

SPrint analyzes the pages of PDF documents and automatically sends them to the appropriate printer based on size:
- **A4 printer** — standard A4 pages
- **A3 printer** — A3 pages
- **Plotter / large format** — all pages larger than A3 (technical drawings, blueprints, etc.)

Supports drag & drop of PDF files, copy count selection, page size analysis display, and remembered printer selection.

## Features

- Automatic detection of each page's size in a PDF
- Grouped printing by printer in one click
- Drag & drop interface
- Roll/plotter printer support with automatic print length calculation
- Remembers last printer selection (`printer_choices.ini`)
- Works as a `.py` script or standalone `.exe` (PyInstaller)
- English / Croatian UI language toggle

## Running

### As a Python script

Install dependencies:

```bash
pip install ttkbootstrap tkinterdnd2 pywin32 PyPDF2
```

Run:

```bash
python PRINT.py
```

### As .exe

Download `SPrint.exe` from [Releases](../../releases) or build it yourself:

```bash
pyinstaller SPrint.spec
```

## Dependencies

| Package | Purpose |
|---|---|
| `ttkbootstrap` | Modern tkinter theme |
| `tkinterdnd2` | Drag & drop support |
| `pywin32` | Windows printer API |
| `PyPDF2` | Reading and splitting PDFs |

## Requirements

- Windows 10 / 11
- Python 3.9+
- Printers installed in Windows

## Author

Filip Kozina

---

# SPrint — Hrvatski

Windows program za automatsko printanje više PDF datoteka različitih veličina na više printera i plotera.

## Opis

SPrint analizira stranice PDF dokumenata i automatski ih šalje na odgovarajući printer ovisno o veličini:
- **A4 printer** — standardne A4 stranice
- **A3 printer** — A3 stranice
- **Plotter / veliki format** — sve stranice veće od A3 (tehnički crteži, nacrti i sl.)

Podržava drag & drop PDF datoteka, odabir broja kopija, prikaz analize stranica po veličini i pohranu odabira printera.

## Značajke

- Automatska detekcija veličine svake stranice u PDF-u
- Grupiranje i printanje po printeru u jednom kliku
- Drag & drop sučelje
- Podrška za roll/ploter printere s automatskim izračunom duljine ispisa
- Pamćenje zadnjeg odabira printera (`printer_choices.ini`)
- Radi kao `.py` skripta ili kao standalone `.exe` (PyInstaller)
- Odabir jezika sučelja: engleski / hrvatski

## Pokretanje

### Kao Python skripta

Instaliraj ovisnosti:

```bash
pip install ttkbootstrap tkinterdnd2 pywin32 PyPDF2
```

Pokreni:

```bash
python PRINT.py
```

### Kao .exe

Preuzmi `SPrint.exe` iz [Releases](../../releases) ili buildi sam:

```bash
pyinstaller SPrint.spec
```

## Ovisnosti

| Paket | Svrha |
|---|---|
| `ttkbootstrap` | Moderni tkinter theme |
| `tkinterdnd2` | Drag & drop podrška |
| `pywin32` | Windows printer API |
| `PyPDF2` | Čitanje i splitanje PDF-a |

## Zahtjevi

- Windows 10 / 11
- Python 3.9+
- Instalirani printeri u Windows sustavu

## Autor

Filip Kozina
