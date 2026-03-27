# SPrint

Windows program za automatsko printanje više PDF datoteka različitih veličina na više printera i plotera.

## Opis

SPrint analizira stranice PDF dokumenata i automatski ih šalje na odgovarajući printer ovisno o veličini:
- **A4 printer** — standardne A4 stranice
- **A3 printer** — A3 stranice
- **Plotter / large format** — sve stranice veće od A3 (tehnički crteži, nacrty i sl.)

Podržava drag & drop PDF datoteka, odabir broja kopija, prikaz analize stranica po veličini i pohunu odabira printera.

## Značajke

- Automatska detekcija veličine svake stranice u PDF-u
- Grupiranje i printanje po printeru u jednom kliku
- Drag & drop sučelje
- Podrška za roll/ploter printere s automatskim izračunom duljine ispisa
- Pamćenje zadnjeg odabira printera (`printer_choices.ini`)
- Radi kao `.py` skripta ili kao standalone `.exe` (PyInstaller)

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
| `ttkbootstrap` | moderni tkinter theme |
| `tkinterdnd2` | drag & drop podrška |
| `pywin32` | Windows printer API |
| `PyPDF2` | čitanje i splitanje PDF-a |

## Zahtjevi

- Windows 10 / 11
- Python 3.9+
- Instalirani printeri u Windows sustavu
