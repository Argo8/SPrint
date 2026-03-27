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
    """Find resource — works from both .py and PyInstaller .exe"""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)

log_text = None
temp_pdf_files = []
current_lang = 'en'

# ── Translations ─────────────────────────────────────────────────────────────
LANG = {
    'en': {
        'subtitle':       'automatic print router',
        'card_files':     'FILES',
        'card_print':     'PRINT SETTINGS',
        'card_plotter':   'PLOTTER — PAPER WIDTH',
        'card_log':       'ACTION LOG',
        'placeholder':    'Drag PDF files here  or click  + Add',
        'btn_add':        '+ Add',
        'btn_remove':     '− Remove',
        'copies_lbl':     'Copies:',
        'lbl_a4':         'Printer A4:',
        'lbl_a3':         'Printer A3:',
        'lbl_plotter':    'Plotter:',
        'btn_print_a4':   'Print A4',
        'btn_print_a3':   'Print A3',
        'btn_plot':       'Plot',
        'author':         'Author: Filip Kozina',
        # plurals
        'page_1':   'page',    'page_n':   'pages',
        'group_1':  'group',   'group_n':  'groups',
        'copy_1':   'copy',    'copy_n':   'copies',
        'file_1':   'file',    'file_n':   'files',
        # log messages
        'no_doc':           'No document selected.',
        'print_start':      'Starting print of {copies} {copies_word} ({n} {file_word})...',
        'processing':       'Processing: {name}',
        'print_a4_on':      'Printing A4 pages on {printer}...',
        'print_a4_done':    'Finished printing A4 pages.',
        'print_a3_on':      'Printing A3 pages on {printer}...',
        'print_a3_done':    'Finished printing A3 pages.',
        'plot_on':          'Plotting on {printer}...',
        'plot_done':        'Plotting finished.',
        'print_done':       'Print complete.',
        'save_ok':          'Printer choices saved.',
        'save_fail':        'Failed to save printer choices. Error: {e}',
        'temp_del_fail':    'Temporary files not deleted: {e}',
        'dim_fail':         'Auto paper size change failed: {e}',
        'no_roll_sel':      'Select roll width (radio button) before plotting.',
        'no_pages_for_roll':'No pages for roll {label}.',
        'plot_groups':      'Plot on roll {label}: {n} {group_word} by length.',
        'temp_fail':        'Error creating temp file: {e}',
        'group_item':       '  [{i}/{total}] length {length} mm — {n} {page_word}',
        'paper_set':        '    Paper: {w}×{h} mm',
        'dim_warn':         '    Warning: paper dimensions not set automatically.',
        'plot_send_fail':   '    Error sending to plotter: {e}',
        'wait_queue':       '    Waiting for print queue confirmation...',
        'read_fail':        'Error reading {name}: {e}',
        'group_fail':       'Error grouping pages from {name}: {e}',
        'filter_fail':      'Page filtering failed. Error: {e}',
        'adobe_fail':       'Adobe Reader error: {e}',
        'print_fail':       'PDF print failed. Error: {e}',
        'save_printer_fail':'Error: Saving printer; {e}',
        # analyze
        'file_single':  'Selected file: {name} ({n} {page_word})',
        'files_multi':  'Selected {n} files ({total} {page_word} total)',
        'size_other':   'Other',
        'warn_pages':   'WARNING: {n} {page_word} could not be read:',
        'page_num':     'Page {n}: {reason}',
        'all_ok':       'All pages are assigned for printing.',
        # on_select
        'roll_a3':    'Selected roll: A3 (297 mm) — {n} {page_word}',
        'roll_mm':    'Selected roll: {w} mm — {n} {page_word}',
        'roll_check': 'Check that this paper width is loaded in the plotter.',
    },
    'hr': {
        'subtitle':       'automatski usmjerivač tiskanja',
        'card_files':     'DATOTEKE',
        'card_print':     'POSTAVKE TISKANJA',
        'card_plotter':   'PLOTER — ŠIRINA PAPIRA',
        'card_log':       'ZAPIS RADNJI',
        'placeholder':    'Povucite PDF datoteke ovdje  ili kliknite  + Dodaj',
        'btn_add':        '+ Dodaj',
        'btn_remove':     '− Ukloni',
        'copies_lbl':     'Broj kopija:',
        'lbl_a4':         'Printer A4:',
        'lbl_a3':         'Printer A3:',
        'lbl_plotter':    'Ploter:',
        'btn_print_a4':   'Print A4',
        'btn_print_a3':   'Print A3',
        'btn_plot':       'Plot',
        'author':         'Autor: Filip Kozina',
        # plurals (Croatian declension handled in _pages())
        'page_1':   'stranica',  'page_n':   'stranice',
        'group_1':  'grupa',     'group_n':  'grupe',
        'copy_1':   'kopija',    'copy_n':   'kopija',
        'file_1':   'file',      'file_n':   'fileova',
        # log messages
        'no_doc':           'Nije odabran dokument.',
        'print_start':      'Započet print {copies} {copies_word} ({n} {file_word})...',
        'processing':       'Obrađujem: {name}',
        'print_a4_on':      'Print A4 stranica na {printer}...',
        'print_a4_done':    'Završen print A4 stranica.',
        'print_a3_on':      'Print A3 stranica na {printer}...',
        'print_a3_done':    'Završen print A3 stranica.',
        'plot_on':          'Plot na {printer}...',
        'plot_done':        'Plotanje završeno.',
        'print_done':       'Print gotov.',
        'save_ok':          'Printer choices saved.',
        'save_fail':        'Failed to save printer choices. Error: {e}',
        'temp_del_fail':    'Privremeni dokumenti nisu obrisani: {e}',
        'dim_fail':         'Automatska promjena dimenzija papira nije uspjela: {e}',
        'no_roll_sel':      'Odaberite širinu role (radio button) prije plotanja.',
        'no_pages_for_roll':'Nema stranica za rolu {label}.',
        'plot_groups':      'Plot na roli {label}: {n} {group_word} po dužini.',
        'temp_fail':        'Greška pri kreiranju temp filea: {e}',
        'group_item':       '  [{i}/{total}] dužina {length} mm — {n} {page_word}',
        'paper_set':        '    Papir: {w}×{h} mm',
        'dim_warn':         '    Upozorenje: dimenzije nisu automatski postavljene.',
        'plot_send_fail':   '    Greška pri slanju na ploter: {e}',
        'wait_queue':       '    Čekam potvrdu slanja u red...',
        'read_fail':        'Greška pri čitanju {name}: {e}',
        'group_fail':       'Greška pri grupiranju stranica iz {name}: {e}',
        'filter_fail':      'Filtriranje stranica nije uspjelo. Error: {e}',
        'adobe_fail':       'Adobe Reader greška: {e}',
        'print_fail':       'Printanje PDF-a nije uspjelo. Error: {e}',
        'save_printer_fail':'Error: Spremanje printera; {e}',
        # analyze
        'file_single':  'Odabran file: {name} ({n} {page_word})',
        'files_multi':  'Odabrano {n} fileova ({total} {page_word} ukupno)',
        'size_other':   'Ostalo',
        'warn_pages':   'UPOZORENJE: {n} {page_word} nije moglo biti pročitano:',
        'page_num':     'Stranica {n}: {reason}',
        'all_ok':       'Sve stranice su raspoređene za printanje.',
        # on_select
        'roll_a3':    'Odabrana rola: A3 (297 mm) — {n} {page_word}',
        'roll_mm':    'Odabrana rola: {w} mm — {n} {page_word}',
        'roll_check': 'Provjerite da li se ta širina papira nalazi u ploteru.',
    }
}

def T(key, **kw):
    s = LANG[current_lang].get(key, key)
    return s.format(**kw) if kw else s

def _pages(n):
    if current_lang == 'hr':
        if n % 10 == 1 and n % 100 != 11:
            return 'stranica'
        elif n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
            return 'stranice'
        return 'stranica'
    return T('page_1') if n == 1 else T('page_n')

def _groups(n):
    return T('group_1') if n == 1 else T('group_n')

def _copies_word(n):
    return T('copy_1') if n == 1 else T('copy_n')

def _files_word(n):
    return T('file_1') if n == 1 else T('file_n')

# ── Helpers ───────────────────────────────────────────────────────────────────

def center_window(window):
    window_width  = window.winfo_reqwidth()
    window_height = window.winfo_reqheight()
    position_x = int(window.winfo_screenwidth()  / 2 - window_width  / 2)
    position_y = int(window.winfo_screenheight() / 2 - window_height / 2)
    window.geometry(f'{window_width}x{window_height}+{position_x}+{position_y}')

def open_printer_properties(printer_name):
    import subprocess
    subprocess.run(
        ['rundll32', 'printui.dll,PrintUIEntry', '/e', '/n', printer_name],
        creationflags=0x08000000)

def set_printer_paper_dimensions(printer_name, width_mm, height_mm):
    DM_OUT_BUFFER = 2
    DM_IN_BUFFER  = 8
    DM_UPDATE     = 1
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            devmode = win32print.DocumentProperties(
                None, hprinter, printer_name, None, None, DM_OUT_BUFFER)
            devmode.Fields  = devmode.Fields | 0x0002 | 0x0004 | 0x0008
            devmode.PaperSize   = 256
            devmode.PaperWidth  = round(width_mm  * 10)
            devmode.PaperLength = round(height_mm * 10)
            win32print.DocumentProperties(
                None, hprinter, printer_name, None, devmode, DM_IN_BUFFER | DM_UPDATE)
            return True
        finally:
            win32print.ClosePrinter(hprinter)
    except Exception as e:
        redirect_output(T('dim_fail', e=e))
        return False

def _wait_for_new_print_job(printer_name, timeout=12):
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
        time.sleep(3)
        return False

ROLL_OPTIONS = [297, 420, 610, 914]

def _assigned_roll(w_mm, h_mm):
    needed = min(w_mm, h_mm)
    for opt in sorted(ROLL_OPTIONS):
        if opt >= needed:
            return opt
    return max(ROLL_OPTIONS)

_A_SIZES_MM     = [(210, 297), (297, 210), (297, 420), (420, 297)]
_STD_THRESHOLD_PT = 5

def _is_standard_page(page):
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
    groups  = {}
    a3_mode = (roll_width_mm == -1)
    for path in input_pdf_paths:
        try:
            reader = PdfReader(path)
            for page in reader.pages:
                w_mm = round(float(page["/MediaBox"][2]) / 2.83465)
                h_mm = round(float(page["/MediaBox"][3]) / 2.83465)
                if a3_mode:
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
            redirect_output(T('group_fail', name=os.path.basename(path), e=e))
    return groups

def print_large_pages_auto(printer_name, input_pdf_paths, copies, roll_width_mm):
    if not roll_width_mm:
        redirect_output(T('no_roll_sel'))
        return

    a3_mode    = (roll_width_mm == -1)
    paper_width = 297 if a3_mode else roll_width_mm
    roll_label  = "A3 (297 mm)" if a3_mode else f"{roll_width_mm} mm"

    groups = group_large_pages_by_print_length(input_pdf_paths, roll_width_mm)
    if not groups:
        redirect_output(T('no_pages_for_roll', label=roll_label))
        return

    total_groups = len(groups)
    redirect_output(T('plot_groups', label=roll_label, n=total_groups,
                       group_word=_groups(total_groups)))

    sent = False
    for i, (print_length_mm, writer) in enumerate(
            sorted(groups.items(), reverse=True), start=1):
        n        = len(writer.pages)
        tmp_path = f"temp_large_{paper_width}x{print_length_mm}.pdf"
        try:
            with open(tmp_path, 'wb') as f:
                writer.write(f)
            temp_pdf_files.append(tmp_path)
        except Exception as e:
            redirect_output(T('temp_fail', e=e))
            continue

        redirect_output(T('group_item', i=i, total=total_groups,
                           length=print_length_mm, n=n, page_word=_pages(n)))

        ok = set_printer_paper_dimensions(printer_name, paper_width, print_length_mm)
        if ok:
            redirect_output(T('paper_set', w=paper_width, h=print_length_mm))
        else:
            redirect_output(T('dim_warn'))

        try:
            print_pdf(printer_name, tmp_path, copies)
            sent = True
        except Exception as e:
            redirect_output(T('plot_send_fail', e=e))

        if i < total_groups:
            redirect_output(T('wait_queue'))
            _wait_for_new_print_job(printer_name)

    if sent:
        redirect_output(T('plot_done'))
    return sent

def delete_temp_pdf_files_and_exit():
    global temp_pdf_files
    for temp_file in temp_pdf_files:
        try:
            os.remove(temp_file)
        except Exception as e:
            redirect_output(T('temp_del_fail', e=e))
    try:
        save_printer_choices(printer_choice_a4.get(), printer_choice_a3.get(),
                             printer_choice_large.get())
        redirect_output(T('save_ok'))
    except Exception as e:
        redirect_output(T('save_fail', e=e))
    root.quit()

def redirect_output(message):
    global log_text
    if log_text is not None:
        log_text.insert(tk.END, message + '\n')
        log_text.see(tk.END)
        print(message)

def load_printer_choices():
    try:
        config = configparser.ConfigParser()
        config.read('printer_choices.ini')
        choices = config['PRINTERS']
        return choices['A4'], choices['A3'], choices['Large']
    except Exception:
        return None, None, None

def save_printer_choices(printer_a4, printer_a3, printer_large):
    try:
        config = configparser.ConfigParser()
        config['PRINTERS'] = {'A4': printer_a4, 'A3': printer_a3, 'Large': printer_large}
        with open('printer_choices.ini', 'w') as configfile:
            config.write(configfile)
    except Exception as e:
        redirect_output(T('save_printer_fail', e=e))

def print_filtered_document(printer_a4, printer_a3, printer_large, copies, file_paths):
    if not file_paths:
        redirect_output(T('no_doc'))
        return

    n = len(file_paths)
    redirect_output(T('print_start', copies=copies, copies_word=_copies_word(copies),
                       n=n, file_word=_files_word(n)))

    for file_path in file_paths:
        redirect_output(T('processing', name=os.path.basename(file_path)))
        a4_pdf_path    = create_pdf_with_filtered_pages(file_path, "A4")
        a3_pdf_path    = create_pdf_with_filtered_pages(file_path, "A3")
        large_pdf_path = create_pdf_with_filtered_pages(file_path, "large")

        if a4_pdf_path and printer_a4:
            redirect_output(T('print_a4_on', printer=printer_a4))
            print_pdf(printer_a4, a4_pdf_path, copies)
            redirect_output(T('print_a4_done'))

        if a3_pdf_path and printer_a3:
            redirect_output(T('print_a3_on', printer=printer_a3))
            print_pdf(printer_a3, a3_pdf_path, copies)
            redirect_output(T('print_a3_done'))

        if large_pdf_path and printer_large:
            redirect_output(T('plot_on', printer=printer_large))
            print_pdf(printer_large, large_pdf_path, copies)
            redirect_output(T('plot_done'))

    redirect_output(T('print_done'))

def analyze_pdf_pages(input_pdf_paths, threshold=5):
    if isinstance(input_pdf_paths, str):
        input_pdf_paths = [input_pdf_paths]

    sizes_in_mm = {
        "A4": [(210, 297), (297, 210)],
        "A3": [(297, 420), (420, 297)],
    }

    def classify_page(page):
        page_width  = float(page["/MediaBox"][2])
        page_height = float(page["/MediaBox"][3])
        for size_name, variants in sizes_in_mm.items():
            for w_mm, h_mm in variants:
                w_pt = w_mm * 2.83465
                h_pt = h_mm * 2.83465
                if (abs(page_width - w_pt) <= threshold and abs(page_height - h_pt) <= threshold) or \
                   (abs(page_width - h_pt) <= threshold and abs(page_height - w_pt) <= threshold):
                    return size_name
        return "OTHER"

    counts       = {}
    ostalo_dims  = {}
    width_counts = {-1: 0, **{opt: 0 for opt in ROLL_OPTIONS}}
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
                    elif label == "OTHER":
                        w_mm = round(float(page["/MediaBox"][2]) / 2.83465)
                        h_mm = round(float(page["/MediaBox"][3]) / 2.83465)
                        dim  = f"{w_mm}x{h_mm} mm"
                        ostalo_dims[dim] = ostalo_dims.get(dim, 0) + 1
                        width_counts[_assigned_roll(w_mm, h_mm)] += 1
                except Exception as e:
                    failed_pages.append((page_num, str(e)))
        except Exception as e:
            redirect_output(T('read_fail', name=os.path.basename(input_pdf_path), e=e))

    total = sum(counts.values())
    if len(input_pdf_paths) == 1:
        header = T('file_single', name=os.path.basename(input_pdf_paths[0]),
                   n=total, page_word=_pages(total))
    else:
        header = T('files_multi', n=len(input_pdf_paths), total=total,
                   page_word=_pages(total))

    lines = [header]
    for size_key, display in [("A4", "A4"), ("A3", "A3"), ("OTHER", T('size_other'))]:
        if size_key in counts:
            cnt = counts[size_key]
            lines.append(f"  {display}: {cnt} {_pages(cnt)}")
            if size_key == "OTHER":
                for dim, c in sorted(ostalo_dims.items()):
                    lines.append(f"    {dim}: {c} {_pages(c)}")

    if failed_pages:
        nf = len(failed_pages)
        lines.append(T('warn_pages', n=nf, page_word=_pages(nf)))
        for pnum, reason in failed_pages:
            lines.append(T('page_num', n=pnum, reason=reason))
    elif total > 0:
        lines.append(T('all_ok'))

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
        input_pdf         = PdfReader(input_pdf_path)
        output_pdf_writer = PdfWriter()

        def is_acceptable_size(page, target_size, threshold):
            sizes_in_mm = {
                "A4": [(210, 297), (297, 210)],
                "A3": [(297, 420), (420, 297)],
            }
            page_width  = float(page["/MediaBox"][2])
            page_height = float(page["/MediaBox"][3])
            for standard_size in sizes_in_mm[target_size]:
                standard_width_mm, standard_height_mm = standard_size
                standard_width  = standard_width_mm  * 2.83465
                standard_height = standard_height_mm * 2.83465
                if (
                    (abs(page_width  - standard_width)  <= threshold and
                     abs(page_height - standard_height) <= threshold) or
                    (abs(page_width  - standard_height) <= threshold and
                     abs(page_height - standard_width)  <= threshold)
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

        if pages_added:
            output_pdf_path = f"temp_{page_size}_{os.path.basename(input_pdf_path)}"
            with open(output_pdf_path, "wb") as output_pdf_file:
                output_pdf_writer.write(output_pdf_file)
            temp_pdf_files.append(output_pdf_path)
            return output_pdf_path
    except Exception as e:
        redirect_output(T('filter_fail', e=e))
    return None

def _find_adobe():
    candidates = [
        r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
        r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
    ]
    return next((p for p in candidates if os.path.exists(p)), None)

def print_pdf(printer_name, file_name, copies):
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
            redirect_output(T('adobe_fail', e=e))
    try:
        for _ in range(copies):
            win32api.ShellExecute(0, "printto", file_name, f'"{printer_name}"', ".", 0)
    except Exception as e:
        redirect_output(T('print_fail', e=e))

# ── Main UI ───────────────────────────────────────────────────────────────────

def main():
    global log_text, root, current_lang

    root = TkinterDnD.Tk()
    root.withdraw()
    root.title("SPrint")
    root.resizable(False, False)
    try:
        root.iconbitmap(resource_path('icon.ico'))
    except Exception:
        pass

    style = Style(theme='litera')

    PAGE_BG = '#F5F7FA'
    CARD_BG = '#FFFFFF'
    FG      = '#212121'
    FG2     = '#757575'
    PRIMARY = '#1565C0'
    BORDER  = '#E0E0E0'
    SEL_BG  = '#1565C0'
    LOG_BG  = '#FAFAFA'
    RBTN_BG = CARD_BG

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
    style.configure('Author.TLabel',
                    font=('Segoe UI', 9), foreground=FG2)

    root.configure(bg=PAGE_BG)

    # ── Variables ───────────────────────────────────────────
    global printer_choice_a4, printer_choice_a3, printer_choice_large
    printer_choice_a4    = tk.StringVar()
    printer_choice_a3    = tk.StringVar()
    printer_choice_large = tk.StringVar()
    number_of_copies     = tk.IntVar(value=1)
    chosen_width         = tk.IntVar()
    file_paths           = []
    current_width_counts = {}
    options              = [-1, 297, 420, 610, 914]
    rbtns                = []
    w                    = {}   # widget references for language updates

    # ── Language switch ─────────────────────────────────────
    def apply_lang(lang):
        global current_lang
        current_lang = lang
        w['subtitle'].config(text=T('subtitle'))
        w['card_files_lbl'].config(text=T('card_files'))
        w['card_print_lbl'].config(text=T('card_print'))
        w['card_plotter_lbl'].config(text=T('card_plotter'))
        w['card_log_lbl'].config(text=T('card_log'))
        w['placeholder'].config(text=T('placeholder'))
        w['btn_add'].config(text=T('btn_add'))
        w['btn_remove'].config(text=T('btn_remove'))
        w['lbl_copies'].config(text=T('copies_lbl'))
        w['lbl_a4'].config(text=T('lbl_a4'))
        w['lbl_a3'].config(text=T('lbl_a3'))
        w['lbl_plotter'].config(text=T('lbl_plotter'))
        w['btn_print_a4'].config(text=T('btn_print_a4'))
        w['btn_print_a3'].config(text=T('btn_print_a3'))
        w['btn_plot'].config(text=T('btn_plot'))
        w['author_lbl'].config(text=T('author'))
        # toggle button styles
        w['btn_lang_en'].config(
            bootstyle='primary' if lang == 'en' else 'secondary-outline')
        w['btn_lang_hr'].config(
            bootstyle='primary' if lang == 'hr' else 'secondary-outline')
        # refresh radio buttons
        for rbtn, opt in zip(rbtns, options):
            lbl     = "A3" if opt == -1 else f"{opt} mm"
            cnt     = current_width_counts.get(opt, 0)
            checked = "  ✓" if rbtn.cget("text").endswith("  ✓") else ""
            if cnt > 0:
                rbtn.config(text=f"{lbl}  —  {cnt} {_pages(cnt)}{checked}")
            else:
                rbtn.config(text=f"{lbl}{checked}")

    # ── Outer frame ─────────────────────────────────────────
    outer = ttk.Frame(root, padding=(24, 18, 24, 18))
    outer.grid(column=0, row=0, sticky='nsew')
    outer.configure(style='TFrame')
    root.columnconfigure(0, weight=1)
    outer.columnconfigure(0, weight=1)

    # ── Header ──────────────────────────────────────────────
    hdr = ttk.Frame(outer)
    hdr.grid(row=0, column=0, sticky='ew', pady=(0, 6))
    hdr.columnconfigure(1, weight=1)

    ttk.Label(hdr, text="SPrint", style='Title.TLabel').grid(
        row=0, column=0, sticky='w')
    w['subtitle'] = ttk.Label(hdr, text=T('subtitle'), style='Subtitle.TLabel')
    w['subtitle'].grid(row=0, column=1, sticky='w', pady=(9, 0), padx=(8, 0))

    # Language toggle buttons
    lang_frame = tk.Frame(hdr, bg=PAGE_BG)
    lang_frame.grid(row=0, column=2, sticky='e')
    w['btn_lang_en'] = ttk.Button(
        lang_frame, text='EN', width=3, bootstyle='primary',
        command=lambda: apply_lang('en'))
    w['btn_lang_en'].pack(side=tk.LEFT, padx=(0, 3))
    w['btn_lang_hr'] = ttk.Button(
        lang_frame, text='HR', width=3, bootstyle='secondary-outline',
        command=lambda: apply_lang('hr'))
    w['btn_lang_hr'].pack(side=tk.LEFT)

    ttk.Separator(outer, orient='horizontal').grid(
        row=1, column=0, sticky='ew', pady=(2, 14))

    # ── Card helper ─────────────────────────────────────────
    def make_card(parent, row, label):
        frame = tk.Frame(parent, bg=CARD_BG, bd=1, relief='solid',
                         highlightthickness=1, highlightbackground=BORDER)
        frame.grid(row=row, column=0, sticky='ew', pady=(0, 12), ipadx=0)
        frame.columnconfigure(0, weight=1)
        strip = tk.Frame(frame, bg=PRIMARY, height=3)
        strip.grid(row=0, column=0, sticky='ew')
        lbl = tk.Label(frame, text=label, bg=CARD_BG, fg=FG2,
                       font=('Segoe UI', 8, 'bold'), anchor='w', padx=14, pady=6)
        lbl.grid(row=1, column=0, sticky='ew')
        ttk.Separator(frame, orient='horizontal').grid(row=2, column=0, sticky='ew')
        inner = tk.Frame(frame, bg=CARD_BG, padx=14, pady=10)
        inner.grid(row=3, column=0, sticky='ew')
        inner.columnconfigure(0, weight=1)
        return inner, lbl

    # ── Callbacks ───────────────────────────────────────────
    def refresh_analysis():
        width_counts = analyze_pdf_pages(file_paths)
        current_width_counts.clear()
        current_width_counts.update(width_counts)
        for rbtn, opt in zip(rbtns, options):
            lbl = "A3" if opt == -1 else f"{opt} mm"
            cnt = width_counts.get(opt, 0)
            rbtn.config(text=f"{lbl}  —  {cnt} {_pages(cnt)}")

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
        n     = current_width_counts.get(width, 0)
        if width == -1:
            lines = [T('roll_a3', n=n, page_word=_pages(n)), T('roll_check')]
        else:
            lines = [T('roll_mm', w=width, n=n, page_word=_pages(n)), T('roll_check')]
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
    fc, w['card_files_lbl'] = make_card(outer, row=2, label=T('card_files'))
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
        raw   = event.data
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

    w['placeholder'] = tk.Label(
        fc, text=T('placeholder'),
        bg='#F8F9FA', fg='#BDBDBD', font=('Segoe UI', 9, 'italic'),
        cursor='arrow')
    w['placeholder'].place(in_=file_listbox, relx=0.5, rely=0.5, anchor='center')

    def update_placeholder():
        if file_paths:
            w['placeholder'].place_forget()
        else:
            w['placeholder'].place(in_=file_listbox, relx=0.5, rely=0.5, anchor='center')

    file_listbox.drop_target_register(DND_FILES)
    file_listbox.dnd_bind('<<Drop>>', on_drop)
    w['placeholder'].drop_target_register(DND_FILES)
    w['placeholder'].dnd_bind('<<Drop>>', on_drop)

    btn_fc = tk.Frame(fc, bg=CARD_BG)
    btn_fc.grid(row=0, column=1, rowspan=2, sticky='ns')
    w['btn_add'] = ttk.Button(btn_fc, text=T('btn_add'), bootstyle='primary-outline',
                               width=9, command=add_files)
    w['btn_add'].pack(fill='x', pady=(0, 6))
    w['btn_remove'] = ttk.Button(btn_fc, text=T('btn_remove'), bootstyle='danger-outline',
                                  width=9, command=remove_file)
    w['btn_remove'].pack(fill='x')

    # ── PRINTERS CARD ───────────────────────────────────────
    pc, w['card_print_lbl'] = make_card(outer, row=3, label=T('card_print'))
    pc.columnconfigure(1, weight=1)

    w['lbl_copies'] = tk.Label(pc, text=T('copies_lbl'), bg=CARD_BG, fg=FG,
                                font=('Segoe UI', 10))
    w['lbl_copies'].grid(row=0, column=0, sticky='w', padx=(0, 12), pady=(0, 8))
    ttk.Spinbox(pc, from_=1, to=999, textvariable=number_of_copies,
                width=7).grid(row=0, column=1, sticky='w', pady=(0, 8))

    ttk.Separator(pc, orient='horizontal').grid(
        row=1, column=0, columnspan=4, sticky='ew', pady=(0, 8))

    printer_row_defs = [
        ('lbl_a4',      printer_choice_a4,
         lambda: print_filtered_document(printer_choice_a4.get(), None, None,
                                         number_of_copies.get(), file_paths),
         'btn_print_a4', 'primary'),
        ('lbl_a3',      printer_choice_a3,
         lambda: print_filtered_document(None, printer_choice_a3.get(), None,
                                         number_of_copies.get(), file_paths),
         'btn_print_a3', 'primary'),
        ('lbl_plotter', printer_choice_large,
         plot_and_mark, 'btn_plot', 'warning'),
    ]

    for row_i, (lbl_key, var, cmd, btn_key, bstyle) in enumerate(
            printer_row_defs, start=2):
        lbl = tk.Label(pc, text=T(lbl_key), bg=CARD_BG, fg=FG, font=('Segoe UI', 10))
        lbl.grid(row=row_i, column=0, sticky='w', padx=(0, 12), pady=5)
        w[lbl_key] = lbl

        cb = ttk.Combobox(pc, textvariable=var, state='readonly')
        cb.grid(row=row_i, column=1, sticky='ew', pady=5)
        if row_i == 2:
            printer_dropdown_a4 = cb
        elif row_i == 3:
            printer_dropdown_a3 = cb
        else:
            printer_dropdown_large = cb

        _v = var
        ttk.Button(pc, text="⚙", width=3, bootstyle='secondary-outline',
                   command=lambda v=_v: open_printer_properties(
                       v.get())).grid(row=row_i, column=2, padx=8)

        btn = ttk.Button(pc, text=T(btn_key), bootstyle=bstyle, width=9, command=cmd)
        btn.grid(row=row_i, column=3)
        w[btn_key] = btn

    # ── PLOTTER CARD ────────────────────────────────────────
    plc, w['card_plotter_lbl'] = make_card(outer, row=4, label=T('card_plotter'))
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
    lc, w['card_log_lbl'] = make_card(outer, row=5, label=T('card_log'))
    lc.columnconfigure(0, weight=1)

    log_text = scrolledtext.ScrolledText(
        lc, wrap=tk.WORD, width=80, height=15,
        bg=LOG_BG, fg='#37474F', insertbackground=FG,
        font=('Consolas', 9), relief='flat', borderwidth=0,
        highlightthickness=1, highlightbackground=BORDER,
        padx=10, pady=8)
    log_text.grid(column=0, row=0, sticky='ew')

    # ── Author footer ────────────────────────────────────────
    footer = tk.Frame(outer, bg=PAGE_BG)
    footer.grid(row=6, column=0, sticky='ew', pady=(4, 0))
    ttk.Separator(outer, orient='horizontal').grid(
        row=6, column=0, sticky='ew', pady=(8, 0))
    footer2 = tk.Frame(outer, bg=PAGE_BG)
    footer2.grid(row=7, column=0, sticky='ew', pady=(4, 0))
    w['author_lbl'] = ttk.Label(footer2, text=T('author'), style='Author.TLabel')
    w['author_lbl'].pack(side=tk.RIGHT)

    root.protocol("WM_DELETE_WINDOW", delete_temp_pdf_files_and_exit)

    # ── Load printers ────────────────────────────────────────
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
    main()
