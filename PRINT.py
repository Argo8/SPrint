import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, simpledialog
from ttkbootstrap import Style
from tkinterdnd2 import TkinterDnD, DND_FILES
import win32print
import win32api
import os
import sys
from PyPDF2 import PdfReader, PdfWriter
import configparser
import threading

def resource_path(relative_path):
    """Pronađi resurs — radi i iz .py i iz PyInstaller .exe"""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)
#import atexit
#from ttkthemes import ThemedTk

log_text = None  # Initialize log_text as a global variable
temp_pdf_files = []  # List to store temporary PDF files

def center_window(window):
    # Get the width and height of the window
    window_width = window.winfo_reqwidth()
    window_height = window.winfo_reqheight()

    # Calculate the position to center the window
    position_x = int(window.winfo_screenwidth()/2 - window_width/2)
    position_y = int(window.winfo_screenheight()/2 - window_height/2)

    # Set the geometry of the window
    window.geometry(f'{window_width}x{window_height}+{position_x}+{position_y}')

def open_printer_properties(printer_name):
    import subprocess
    subprocess.run(
        ['rundll32', 'printui.dll,PrintUIEntry', '/e', '/n', printer_name],
        creationflags=0x08000000  # CREATE_NO_WINDOW
    )

def set_printer_paper_dimensions(printer_name, width_mm, height_mm):
    """
    Postavi dimenzije papira printera (širina i dužina u mm).
    Koristi DocumentProperties s DM_UPDATE — ne traži admin prava.
    Vraća True ako uspije.
    """
    DM_OUT_BUFFER = 2   # dohvati trenutni DEVMODE
    DM_IN_BUFFER  = 8   # koristi ulazni DEVMODE
    DM_UPDATE     = 1   # spremi kao per-user postavke (ne treba admin)
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            devmode = win32print.DocumentProperties(
                None, hprinter, printer_name, None, None, DM_OUT_BUFFER)
            devmode.Fields  = devmode.Fields | 0x0002 | 0x0004 | 0x0008
            devmode.PaperSize   = 256                    # DMPAPER_USER
            devmode.PaperWidth  = round(width_mm  * 10)  # desetinke mm
            devmode.PaperLength = round(height_mm * 10)  # desetinke mm
            win32print.DocumentProperties(
                None, hprinter, printer_name, None, devmode, DM_IN_BUFFER | DM_UPDATE)
            return True
        finally:
            win32print.ClosePrinter(hprinter)
    except Exception as e:
        redirect_output(f"Automatska promjena dimenzija papira nije uspjela: {e}")
        return False

def _wait_for_new_print_job(printer_name, timeout=12):
    """Čeka dok se novi job ne pojavi u redu printera. Vraća True ako je detektiran."""
    import time
    try:
        h = win32print.OpenPrinter(printer_name)
        initial = set(j['JobId'] for j in win32print.EnumJobs(h, 0, 100, 1))
        win32print.ClosePrinter(h)
        start = time.time()
        while time.time() - start < timeout:
            time.sleep(0.4)
            h = win32print.OpenPrinter(printer_name)
            current = set(j['JobId'] for j in win32print.EnumJobs(h, 0, 100, 1))
            win32print.ClosePrinter(h)
            if current - initial:
                return True
        return False
    except Exception:
        import time
        time.sleep(3)  # fallback
        return False

ROLL_OPTIONS = [297, 420, 610, 914]

def _assigned_roll(w_mm, h_mm):
    """Vraća najmanji roll width koji može primiti stranicu."""
    needed = min(w_mm, h_mm)
    for opt in sorted(ROLL_OPTIONS):
        if opt >= needed:
            return opt
    return max(ROLL_OPTIONS)

_A_SIZES_MM = [(210, 297), (297, 210), (297, 420), (420, 297)]
_STD_THRESHOLD_PT = 5  # iste tolerancije kao u analyze_pdf_pages

def _is_standard_page(page):
    """Vraća True ako je stranica A4 ili A3 (isti kriterij kao analyze_pdf_pages)."""
    w_pt = float(page["/MediaBox"][2])
    h_pt = float(page["/MediaBox"][3])
    for sw_mm, sh_mm in _A_SIZES_MM:
        sw = sw_mm * 2.83465
        sh = sh_mm * 2.83465
        if ((abs(w_pt - sw) <= _STD_THRESHOLD_PT and abs(h_pt - sh) <= _STD_THRESHOLD_PT) or
                (abs(w_pt - sh) <= _STD_THRESHOLD_PT and abs(h_pt - sw) <= _STD_THRESHOLD_PT)):
            return True
    return False

def group_large_pages_by_print_length(input_pdf_paths, roll_width_mm):
    """
    Filtrira stranice za zadani roll_width_mm i grupira ih po dužini printanja.
    roll_width_mm == -1: A3 stranice na roli 297mm.
    Vraća dict: {print_length_mm: PdfWriter}
    """
    groups = {}
    a3_mode = (roll_width_mm == -1)

    for path in input_pdf_paths:
        try:
            reader = PdfReader(path)
            for page in reader.pages:
                w_mm = round(float(page["/MediaBox"][2]) / 2.83465)
                h_mm = round(float(page["/MediaBox"][3]) / 2.83465)
                if a3_mode:
                    # Uključi samo A3 stranice (297x420 ili 420x297)
                    is_a3 = any(
                        abs(w_mm - sw) <= 5 and abs(h_mm - sh) <= 5
                        for sw, sh in [(297, 420), (420, 297)]
                    )
                    if not is_a3:
                        continue
                else:
                    if _is_standard_page(page):
                        continue
                    if _assigned_roll(w_mm, h_mm) != roll_width_mm:
                        continue
                print_length = max(w_mm, h_mm)
                if print_length not in groups:
                    groups[print_length] = PdfWriter()
                groups[print_length].add_page(page)
        except Exception as e:
            redirect_output(f"Greška pri grupiranju stranica iz {os.path.basename(path)}: {e}")
    return groups

def print_large_pages_auto(printer_name, input_pdf_paths, copies, roll_width_mm):
    """
    Printa large-format stranice za odabrani roll_width_mm.
    PaperWidth = roll_width_mm (odabran radio buttonom).
    PaperLength = automatski po dužini svake grupe stranica.
    Mora se pokretati u background threadu.
    """
    if not roll_width_mm:
        redirect_output("Odaberite širinu role (radio button) prije plotanja.")
        return

    a3_mode = (roll_width_mm == -1)
    paper_width = 297 if a3_mode else roll_width_mm
    roll_label  = "A3 (297 mm)" if a3_mode else f"{roll_width_mm} mm"

    groups = group_large_pages_by_print_length(input_pdf_paths, roll_width_mm)
    if not groups:
        redirect_output(f"Nema stranica za rolu {roll_label}.")
        return

    total_groups = len(groups)
    redirect_output(f"Plot na roli {roll_label}: {total_groups} "
                    f"{'grupa' if total_groups == 1 else 'grupe'} po dužini.")

    sent = False
    for i, (print_length_mm, writer) in enumerate(
            sorted(groups.items(), reverse=True), start=1):
        n = len(writer.pages)
        tmp_path = f"temp_large_{paper_width}x{print_length_mm}.pdf"
        try:
            with open(tmp_path, 'wb') as f:
                writer.write(f)
            temp_pdf_files.append(tmp_path)
        except Exception as e:
            redirect_output(f"Greška pri kreiranju temp filea: {e}")
            continue

        redirect_output(f"  [{i}/{total_groups}] dužina {print_length_mm} mm "
                        f"— {n} {stranica_deklinacija(n)}")

        ok = set_printer_paper_dimensions(printer_name, paper_width, print_length_mm)
        if ok:
            redirect_output(f"    Papir: {paper_width}×{print_length_mm} mm")
        else:
            redirect_output(f"    Upozorenje: dimenzije nisu automatski postavljene.")

        try:
            print_pdf(printer_name, tmp_path, copies)
            sent = True
        except Exception as e:
            redirect_output(f"    Greška pri slanju na ploter: {e}")

        if i < total_groups:
            redirect_output("    Čekam potvrdu slanja u red...")
            _wait_for_new_print_job(printer_name)

    if sent:
        redirect_output("Plotanje završeno.")
    return sent

def delete_temp_pdf_files_and_exit(): # Function to delete temporary PDF files and exit the program
    global temp_pdf_files
    for temp_file in temp_pdf_files:
        try:
            os.remove(temp_file)
        except Exception as e:
            redirect_output("Privremeni dokumenti nisu obrisani: " + str(e))
    
    try:
        save_printer_choices(printer_choice_a4.get(), printer_choice_a3.get(), printer_choice_large.get())
        redirect_output("Printer choices saved.")
    except Exception as e:
        redirect_output(f"Failed to save printer choices. Error: {str(e)}")

    root.quit()  # Exit the program

def redirect_output(message):
    global log_text  # Access the global log_text variable

    if log_text is not None:
        log_text.insert(tk.END, message + '\n')  # Add the message to the widget
        log_text.see(tk.END)  # Scroll to the bottom
        print(message)  # Optional: retain this if you also want to keep output to the console

def load_printer_choices():
    try:
        config = configparser.ConfigParser()
        config.read('printer_choices.ini')
        choices = config['PRINTERS']
        printer_a4 = choices['A4']
        printer_a3 = choices['A3']
        printer_large = choices['Large']
        return printer_a4, printer_a3, printer_large
    except Exception:
        return None, None, None

def save_printer_choices(printer_a4, printer_a3, printer_large):
    try:
        config = configparser.ConfigParser()
        config['PRINTERS'] = {
            'A4': printer_a4,
            'A3': printer_a3,
            'Large': printer_large,
        }
        with open('printer_choices.ini', 'w') as configfile:
            config.write(configfile)
    except Exception as e:
        redirect_output("Error: Spremanje printera; " + str(e))

def print_filtered_document(printer_a4, printer_a3, printer_large, copies, file_paths):
    global log_text

    if not file_paths:
        redirect_output("Nije odabran dokument.")
        return

    redirect_output(f"Započet print {copies} primjeraka ({len(file_paths)} {'file' if len(file_paths) == 1 else 'fileova'})...")

    for file_path in file_paths:
        redirect_output(f"Obrađujem: {os.path.basename(file_path)}")
        a4_pdf_path = create_pdf_with_filtered_pages(file_path, "A4")
        a3_pdf_path = create_pdf_with_filtered_pages(file_path, "A3")
        large_pdf_path = create_pdf_with_filtered_pages(file_path, "large")

        if a4_pdf_path and printer_a4:
            redirect_output(f"Print A4 stranica na {printer_a4}...")
            print_pdf(printer_a4, a4_pdf_path, copies)
            redirect_output(f"Završen print A4 stranica.")

        if a3_pdf_path and printer_a3:
            redirect_output(f"Print A3 stranica na {printer_a3}...")
            print_pdf(printer_a3, a3_pdf_path, copies)
            redirect_output(f"Završen print A3 stranica.")

        if large_pdf_path and printer_large:
            redirect_output(f"Plot na {printer_large}...")
            print_pdf(printer_large, large_pdf_path, copies)
            redirect_output(f"Plotanje završeno.")

    redirect_output(f"Print gotov.")

def stranica_deklinacija(n):
    if n % 10 == 1 and n % 100 != 11:
        return "stranica"
    elif n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
        return "stranice"
    else:
        return "stranica"


def analyze_pdf_pages(input_pdf_paths, threshold=5):
    if isinstance(input_pdf_paths, str):
        input_pdf_paths = [input_pdf_paths]

    sizes_in_mm = {
        "A4": [(210, 297), (297, 210)],
        "A3": [(297, 420), (420, 297)],
    }

    def classify_page(page):
        page_width = float(page["/MediaBox"][2])
        page_height = float(page["/MediaBox"][3])
        for size_name, variants in sizes_in_mm.items():
            for w_mm, h_mm in variants:
                w_pt = w_mm * 2.83465
                h_pt = h_mm * 2.83465
                if (abs(page_width - w_pt) <= threshold and abs(page_height - h_pt) <= threshold) or \
                   (abs(page_width - h_pt) <= threshold and abs(page_height - w_pt) <= threshold):
                    return size_name
        return "Ostalo"

    counts = {}
    ostalo_dims = {}
    width_counts = {-1: 0, **{opt: 0 for opt in ROLL_OPTIONS}}  # -1 = A3 na ploteru
    failed_pages = []

    for input_pdf_path in input_pdf_paths:
        try:
            input_pdf = PdfReader(input_pdf_path)
            for page_num, page in enumerate(input_pdf.pages, start=1):
                try:
                    label = classify_page(page)
                    counts[label] = counts.get(label, 0) + 1
                    if label == "A3":
                        width_counts[-1] += 1
                    elif label == "Ostalo":
                        w_mm = round(float(page["/MediaBox"][2]) / 2.83465)
                        h_mm = round(float(page["/MediaBox"][3]) / 2.83465)
                        dim = f"{w_mm}x{h_mm} mm"
                        ostalo_dims[dim] = ostalo_dims.get(dim, 0) + 1
                        width_counts[_assigned_roll(w_mm, h_mm)] += 1
                except Exception as e:
                    failed_pages.append((page_num, str(e)))
        except Exception as e:
            redirect_output(f"Greška pri čitanju {os.path.basename(input_pdf_path)}: {e}")

    total = sum(counts.values())
    if len(input_pdf_paths) == 1:
        header = f"Odabran file: {os.path.basename(input_pdf_paths[0])} ({total} {stranica_deklinacija(total)})"
    else:
        n = len(input_pdf_paths)
        header = f"Odabrano {n} fileova ({total} {stranica_deklinacija(total)} ukupno)"
    lines = [header]
    for size in ["A4", "A3", "Ostalo"]:
        if size in counts:
            cnt = counts[size]
            lines.append(f"  {size}: {cnt} {stranica_deklinacija(cnt)}")
            if size == "Ostalo":
                for dim, c in sorted(ostalo_dims.items()):
                    lines.append(f"    {dim}: {c} {stranica_deklinacija(c)}")
    if failed_pages:
        lines.append(f"  UPOZORENJE: {len(failed_pages)} {stranica_deklinacija(len(failed_pages))} nije moglo biti pročitano:")
        for pnum, reason in failed_pages:
            lines.append(f"    Stranica {pnum}: {reason}")
    elif total > 0:
        lines.append(f"  Sve stranice su raspoređene za printanje.")
    msg = "\n".join(lines) + "\n\n"
    if log_text is not None:
        ranges = log_text.tag_ranges("file_info")
        if ranges:
            log_text.delete(ranges[0], ranges[1])
            log_text.insert("1.0", msg, "file_info")
        else:
            log_text.insert("1.0", msg, "file_info")
    return width_counts


def create_pdf_with_filtered_pages(input_pdf_path, page_size, page_size_threshold=5):
    try:
        input_pdf = PdfReader(input_pdf_path)
        output_pdf_writer = PdfWriter()

        # Define a function to check if the page is within acceptable A4, A3, or larger dimensions.
        def is_acceptable_size(page, target_size, threshold):
            # Convert dimensions from mm to points
            sizes_in_mm = {
                "A4": [(210, 297), (297, 210)],  # standard A4 sizes in mm (portrait, landscape)
                "A3": [(297, 420), (420, 297)],  # standard A3 sizes in mm (portrait, landscape)
            }

            # Extract the actual page dimensions
            page_width = float(page["/MediaBox"][2])
            page_height = float(page["/MediaBox"][3])

            # Check if the page size matches the target size
            for standard_size in sizes_in_mm[target_size]:
                standard_width_mm, standard_height_mm = standard_size
                standard_width = standard_width_mm * 2.83465
                standard_height = standard_height_mm * 2.83465

                if (
                    (abs(page_width - standard_width) <= threshold and abs(page_height - standard_height) <= threshold) or
                    (abs(page_width - standard_height) <= threshold and abs(page_height - standard_width) <= threshold)
                ):
                    return True

            return False

        def is_large(page):
            for size in ["A4", "A3"]:
                if is_acceptable_size(page, size, page_size_threshold):
                    return False
            return True

        pages_added = False

        for page in input_pdf.pages:
            if page_size == "large":
                if is_large(page):
                    output_pdf_writer.add_page(page)
                    pages_added = True
            elif is_acceptable_size(page, page_size, page_size_threshold):
                output_pdf_writer.add_page(page)
                pages_added = True

        if pages_added:  # Check if any pages have been added
            output_pdf_path = f"temp_{page_size}_{os.path.basename(input_pdf_path)}"
            with open(output_pdf_path, "wb") as output_pdf_file:
                output_pdf_writer.write(output_pdf_file)
            
            # Append the temporary PDF file path to the list
            temp_pdf_files.append(output_pdf_path)
            
            return output_pdf_path
    except Exception as e:
        redirect_output("Filtriranje stranica nije uspjelo. Error: " + str(e))

    return None

def _find_adobe():
    """Traži AcroRd32.exe / Acrobat.exe na standardnim lokacijama."""
    candidates = [
        r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
        r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
    ]
    return next((p for p in candidates if os.path.exists(p)), None)

def print_pdf(printer_name, file_name, copies):
    """Printa PDF što tišće — Adobe Reader /h → ShellExecute fallback."""
    import subprocess
    NO_WINDOW = 0x08000000

    adobe = _find_adobe()
    if adobe:
        try:
            for _ in range(copies):
                subprocess.Popen(
                    [adobe, '/h', '/t', file_name, printer_name],
                    creationflags=NO_WINDOW)
            return
        except Exception as e:
            redirect_output(f"Adobe Reader greška: {e}")

    # Fallback — može kratko prikazati prozor PDF preglednika
    try:
        for _ in range(copies):
            win32api.ShellExecute(0, "printto", file_name, f'"{printer_name}"', ".", 0)
    except Exception as e:
        redirect_output("Printanje PDF-a nije uspjelo. Error: " + str(e))

def main():
    global log_text, root

    root = TkinterDnD.Tk()
    root.withdraw()
    root.title("SPrint")
    root.resizable(False, False)
    try:
        root.iconbitmap(resource_path('icon.ico'))
    except Exception:
        pass

    style = Style(theme='litera')

    # ── Material-inspired palette ────────────────────────────
    PAGE_BG  = '#F5F7FA'   # page background
    CARD_BG  = '#FFFFFF'   # card surface
    FG       = '#212121'   # primary text
    FG2      = '#757575'   # secondary text
    PRIMARY  = '#1565C0'   # deep blue
    BORDER   = '#E0E0E0'   # card border / divider
    SEL_BG   = '#1565C0'   # selection blue
    LOG_BG   = '#FAFAFA'   # log background
    RBTN_BG  = CARD_BG     # radiobutton parent bg

    style.configure('TLabel',    font=('Segoe UI', 10), foreground=FG)
    style.configure('TButton',   font=('Segoe UI', 10))
    style.configure('TCombobox', font=('Segoe UI', 10))
    style.configure('TSpinbox',  font=('Segoe UI', 10))
    style.configure('Title.TLabel',
                    font=('Segoe UI', 20, 'bold'), foreground=PRIMARY)
    style.configure('Subtitle.TLabel',
                    font=('Segoe UI', 10), foreground=FG2)
    style.configure('CardTitle.TLabel',
                    font=('Segoe UI', 9, 'bold'), foreground=FG2)

    root.configure(bg=PAGE_BG)

    # ── Variables ───────────────────────────────────────────
    global printer_choice_a4, printer_choice_a3, printer_choice_large
    printer_choice_a4    = tk.StringVar()
    printer_choice_a3    = tk.StringVar()
    printer_choice_large = tk.StringVar()
    number_of_copies = tk.IntVar(value=1)
    chosen_width     = tk.IntVar()
    file_paths       = []
    current_width_counts = {}
    options = [-1, 297, 420, 610, 914]
    rbtns   = []

    # ── Outer frame ─────────────────────────────────────────
    outer = ttk.Frame(root, padding=(24, 18, 24, 18))
    outer.grid(column=0, row=0, sticky='nsew')
    outer.configure(style='TFrame')
    root.columnconfigure(0, weight=1)
    outer.columnconfigure(0, weight=1)

    # ── Header ──────────────────────────────────────────────
    hdr = ttk.Frame(outer)
    hdr.grid(row=0, column=0, sticky='ew', pady=(0, 6))
    ttk.Label(hdr, text="SPrint", style='Title.TLabel').pack(side=tk.LEFT)
    ttk.Label(hdr, text="   automatski usmjerivač tiskanja",
              style='Subtitle.TLabel').pack(side=tk.LEFT, pady=(9, 0))
    ttk.Separator(outer, orient='horizontal').grid(
        row=1, column=0, sticky='ew', pady=(2, 14))

    # ── Card helper ─────────────────────────────────────────
    def make_card(parent, row, label):
        """White card with a top-colored label strip."""
        frame = tk.Frame(parent, bg=CARD_BG, bd=1, relief='solid',
                         highlightthickness=1, highlightbackground=BORDER)
        frame.grid(row=row, column=0, sticky='ew', pady=(0, 12), ipadx=0)
        frame.columnconfigure(0, weight=1)
        # label strip
        strip = tk.Frame(frame, bg=PRIMARY, height=3)
        strip.grid(row=0, column=0, sticky='ew')
        lbl = tk.Label(frame, text=label, bg=CARD_BG, fg=FG2,
                       font=('Segoe UI', 8, 'bold'), anchor='w', padx=14, pady=6)
        lbl.grid(row=1, column=0, sticky='ew')
        ttk.Separator(frame, orient='horizontal').grid(
            row=2, column=0, sticky='ew')
        inner = tk.Frame(frame, bg=CARD_BG, padx=14, pady=10)
        inner.grid(row=3, column=0, sticky='ew')
        inner.columnconfigure(0, weight=1)
        return inner

    # ── Callbacks ───────────────────────────────────────────
    def refresh_analysis():
        width_counts = analyze_pdf_pages(file_paths)
        current_width_counts.clear()
        current_width_counts.update(width_counts)
        for rbtn, opt in zip(rbtns, options):
            lbl = "A3" if opt == -1 else f"{opt} mm"
            cnt = width_counts.get(opt, 0)
            rbtn.config(text=f"{lbl}  —  {cnt} {stranica_deklinacija(cnt)}")

    def add_files():
        selected = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for f in selected:
            if f not in file_paths:
                file_paths.append(f)
                file_listbox.insert(tk.END, os.path.basename(f))
        if file_paths:
            update_placeholder()
            refresh_analysis()

    def remove_file():
        for i in reversed(file_listbox.curselection()):
            del file_paths[i]
            file_listbox.delete(i)
        update_placeholder()
        if file_paths:
            refresh_analysis()
        else:
            if log_text is not None:
                ranges = log_text.tag_ranges("file_info")
                if ranges:
                    log_text.delete(ranges[0], ranges[1])
            for rbtn, opt in zip(rbtns, options):
                lbl = "A3" if opt == -1 else f"{opt} mm"
                rbtn.config(text=lbl)

    def on_select():
        width = chosen_width.get()
        n = current_width_counts.get(width, 0)
        if width == -1:
            lines = [
                f"Odabrana rola: A3 (297 mm) — {n} {stranica_deklinacija(n)}",
                f"Provjerite da li se ta širina papira nalazi u ploteru.",
            ]
        else:
            lines = [
                f"Odabrana rola: {width} mm — {n} {stranica_deklinacija(n)}",
                f"Provjerite da li se ta širina papira nalazi u ploteru.",
            ]
        plotter_info.config(state='normal')
        plotter_info.delete('1.0', tk.END)
        plotter_info.insert('1.0', "\n".join(lines))
        plotter_info.config(state='disabled')

    def plot_and_mark():
        printer    = printer_choice_large.get()
        copies     = number_of_copies.get()
        paths      = list(file_paths)
        roll_width = chosen_width.get()

        def _run():
            sent = print_large_pages_auto(printer, paths, copies, roll_width)
            if sent:
                def _mark():
                    for rbtn, opt in zip(rbtns, options):
                        if opt == roll_width:
                            txt = rbtn.cget("text")
                            if not txt.endswith("  ✓"):
                                rbtn.config(text=txt + "  ✓")
                root.after(0, _mark)

        threading.Thread(target=_run, daemon=True).start()

    # ── FILES CARD ──────────────────────────────────────────
    fc = make_card(outer, row=2, label="DATOTEKE")
    fc.columnconfigure(0, weight=1)

    file_listbox = tk.Listbox(
        fc, height=4, selectmode=tk.EXTENDED,
        bg='#F8F9FA', fg=FG, font=('Segoe UI', 10),
        selectbackground=SEL_BG, selectforeground='white',
        relief='flat', borderwidth=0,
        highlightthickness=1, highlightbackground=BORDER,
        activestyle='none')
    file_scrollbar = ttk.Scrollbar(fc, orient=tk.HORIZONTAL,
                                   command=file_listbox.xview)
    file_listbox.configure(xscrollcommand=file_scrollbar.set)
    file_listbox.grid(row=0, column=0, sticky='ew', padx=(0, 8))
    file_scrollbar.grid(row=1, column=0, sticky='ew', padx=(0, 8))

    def on_drop(event):
        import re
        raw = event.data
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        paths = [a or b for a, b in paths]
        added = False
        for p in paths:
            if p.lower().endswith('.pdf') and p not in file_paths:
                file_paths.append(p)
                file_listbox.insert(tk.END, os.path.basename(p))
                added = True
        if added:
            update_placeholder()
            refresh_analysis()

    # Placeholder label preko listboxa
    placeholder = tk.Label(
        fc, text="Povucite PDF datoteke ovdje  ili kliknite  + Dodaj",
        bg='#F8F9FA', fg='#BDBDBD', font=('Segoe UI', 9, 'italic'),
        cursor='arrow')
    placeholder.place(in_=file_listbox, relx=0.5, rely=0.5, anchor='center')

    def update_placeholder():
        if file_paths:
            placeholder.place_forget()
        else:
            placeholder.place(in_=file_listbox, relx=0.5, rely=0.5, anchor='center')

    file_listbox.drop_target_register(DND_FILES)
    file_listbox.dnd_bind('<<Drop>>', on_drop)
    placeholder.drop_target_register(DND_FILES)
    placeholder.dnd_bind('<<Drop>>', on_drop)

    btn_fc = tk.Frame(fc, bg=CARD_BG)
    btn_fc.grid(row=0, column=1, rowspan=2, sticky='ns')
    ttk.Button(btn_fc, text="+ Dodaj",  bootstyle='primary-outline',
               width=9, command=add_files).pack(fill='x', pady=(0, 6))
    ttk.Button(btn_fc, text="− Ukloni", bootstyle='danger-outline',
               width=9, command=remove_file).pack(fill='x')

    # ── PRINTERS CARD ───────────────────────────────────────
    pc = make_card(outer, row=3, label="POSTAVKE TISKANJA")
    pc.columnconfigure(1, weight=1)

    tk.Label(pc, text="Broj kopija:", bg=CARD_BG, fg=FG,
             font=('Segoe UI', 10)).grid(row=0, column=0, sticky='w',
                                         padx=(0, 12), pady=(0, 8))
    ttk.Spinbox(pc, from_=1, to=999, textvariable=number_of_copies,
                width=7).grid(row=0, column=1, sticky='w', pady=(0, 8))

    ttk.Separator(pc, orient='horizontal').grid(
        row=1, column=0, columnspan=4, sticky='ew', pady=(0, 8))

    for row_i, (lbl_txt, var, cmd, btn_txt, bstyle) in enumerate([
        ("Printer A4:", printer_choice_a4,
         lambda: print_filtered_document(printer_choice_a4.get(), None, None,
                                         number_of_copies.get(), file_paths),
         "Print A4", "primary"),
        ("Printer A3:", printer_choice_a3,
         lambda: print_filtered_document(None, printer_choice_a3.get(), None,
                                         number_of_copies.get(), file_paths),
         "Print A3", "primary"),
        ("Ploter:", printer_choice_large,
         plot_and_mark, "Plot", "warning"),
    ], start=2):
        tk.Label(pc, text=lbl_txt, bg=CARD_BG, fg=FG,
                 font=('Segoe UI', 10)).grid(
            row=row_i, column=0, sticky='w', padx=(0, 12), pady=5)
        cb = ttk.Combobox(pc, textvariable=var, state='readonly')
        cb.grid(row=row_i, column=1, sticky='ew', pady=5)
        if row_i == 2:
            printer_dropdown_a4 = cb
        elif row_i == 3:
            printer_dropdown_a3 = cb
        else:
            printer_dropdown_large = cb
        _v = var  # capture for lambda
        ttk.Button(pc, text="⚙", width=3, bootstyle='secondary-outline',
                   command=lambda v=_v: open_printer_properties(
                       v.get())).grid(row=row_i, column=2, padx=8)
        ttk.Button(pc, text=btn_txt, bootstyle=bstyle, width=9,
                   command=cmd).grid(row=row_i, column=3)

    # ── PLOTTER CARD ────────────────────────────────────────
    plc = make_card(outer, row=4, label="PLOTER — ŠIRINA PAPIRA")
    plc.columnconfigure(1, weight=1)

    radio_frame = tk.Frame(plc, bg=CARD_BG)
    radio_frame.grid(row=0, column=0, sticky='nsw', padx=(0, 16))

    for index, option in enumerate(options):
        rbtn_lbl = "A3" if option == -1 else f"{option} mm"
        rbtn = tk.Radiobutton(
            radio_frame, text=rbtn_lbl,
            variable=chosen_width, value=option, command=on_select,
            bg=RBTN_BG, fg=FG, selectcolor='#E3F2FD',
            activebackground=RBTN_BG, activeforeground=PRIMARY,
            font=('Segoe UI', 10), relief='flat', bd=0,
            padx=4, pady=4, cursor='hand2')
        rbtn.grid(row=index, column=0, sticky='w')
        rbtns.append(rbtn)

    plotter_info = tk.Text(
        plc, width=42, height=4, wrap=tk.WORD, state='disabled',
        bg='#F8F9FA', fg=FG2, relief='flat', borderwidth=0,
        font=('Segoe UI', 9),
        highlightthickness=1, highlightbackground=BORDER,
        padx=10, pady=8)
    plotter_info.grid(row=0, column=1, sticky='nsew')

    # ── LOG CARD ────────────────────────────────────────────
    lc = make_card(outer, row=5, label="ZAPIS RADNJI")
    lc.columnconfigure(0, weight=1)

    log_text = scrolledtext.ScrolledText(
        lc, wrap=tk.WORD, width=80, height=15,
        bg=LOG_BG, fg='#37474F', insertbackground=FG,
        font=('Consolas', 9), relief='flat', borderwidth=0,
        highlightthickness=1, highlightbackground=BORDER,
        padx=10, pady=8)
    log_text.grid(column=0, row=0, sticky='ew')

    root.protocol("WM_DELETE_WINDOW", delete_temp_pdf_files_and_exit)

    # Load printers
    printers = [printer[2] for printer in win32print.EnumPrinters(2)]
    printer_dropdown_a4['values']    = printers
    printer_dropdown_a3['values']    = printers
    printer_dropdown_large['values'] = printers

    if printers:
        printer_dropdown_a4.current(0)
        printer_dropdown_a3.current(0)
        printer_dropdown_large.current(0)

    printer_a4, printer_a3, printer_large = load_printer_choices()
    if printer_a4 and printer_a3 and printer_large:
        printer_choice_a4.set(printer_a4)
        printer_choice_a3.set(printer_a3)
        printer_choice_large.set(printer_large)

    root.update()
    center_window(root)
    root.deiconify()
    root.mainloop()

if __name__ == "__main__":
 #   atexit.register(delete_temp_pdf_files_and_exit)  # Register the delete_temp_pdf_files_and_exit function to run on program exit
    main()