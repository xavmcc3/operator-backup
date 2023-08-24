from multiprocessing import Process, freeze_support
from openpyxl.formula.translate import Translator
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter import Tk
import pandas as pd
import datetime
import ctypes
import PIL
import re

from clrprint import clrprint
from pathlib import Path
from io import BytesIO
import subprocess
import openpyxl
import shutil
import sys
import os

# Compile with `python -m PyInstaller --clean operator-tool.spec`
# From `python -m PyInstaller --clean --onefile --icon=operator-tool.ico --name=operator-tool main.py`

hwnd = ctypes.windll.user32.GetForegroundWindow()

def restore_focus():
    ctypes.windll.user32.BringWindowToTop(hwnd)

def file_dialog(filetypes):
    return askopenfilename(filetypes=filetypes)

def folder_dialog():
    return askdirectory()

def empty_dir(path):
    print(f"Emptying ", end="")
    clrprint(path, end="", clr='m')
    print("...", end="\r")

    for filename in os.scandir(path):
        file_path = os.path.join(path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))

def file_is_valid(file: str):
    path = Path(file)
    if "Template" in path.stem:
        return False
    if "~$" in path.stem:
        return False
    if path.suffix in ['.xlsx']:
        return True
    return False

def copy_row(row, row_index, sheet):
        for cell in row:
            sheet.cell(row=row_index, column=cell.column).value = cell.value

def get_name_from_file(src):
    pattern = r'[0-9]'
    path = Path(src)

    rename = re.sub(pattern, '', path.stem)
    return rename.strip()

def load_path_bytes(path):
    with open(path, 'rb') as f:
        template_bytes = BytesIO(f.read())
    return template_bytes

def wb_from_bytes(path_bytes):
    return openpyxl.load_workbook(path_bytes)

def wb_data_from_bytes(path_bytes):
    return openpyxl.load_workbook(path_bytes, data_only=True)


def create_from_template(template_bytes, name, append, dst_folder, use_name=True):
    dst_path = f"{name}{append}.xlsx"
    final_path = os.path.join(os.path.abspath(dst_folder), dst_path)

    print(f"Loading ", end="")
    clrprint(name, end="", clr='y')
    print("...")

    template = wb_from_bytes(template_bytes)

    print(f"Creating ", end="")
    clrprint(dst_path, end="", clr='m')
    print("...")

    ws = template[template.sheetnames[0]]
    for col in ws.iter_cols():
        if str(col[0].value).lower() == "name":
            col[0].value = name if use_name else "Name"
            break
    
    #ws.column_dimensions.group(start='H', end='K', hidden=True)

    template.save(final_path)
    template.close()

    print("Created ", end="")
    clrprint(final_path, clr='g')

def create_all(template_path, src, append, dst):
    print("")
    start_time = datetime.datetime.now()
    processes = []
    template_bytes = load_path_bytes(template_path)
    p = Process(target=create_from_template, args=(template_bytes, "AATemplate", append, dst, False,))
    processes.append(p)
    p.start()
    for file in os.listdir(src):
        if file_is_valid(file):
            name = get_name_from_file(file)
            p = Process(target=create_from_template, args=(template_bytes, name, append, dst,))
            processes.append(p)
            p.start()

    for p in processes:
        p.join()

    clrprint("Finished in ", f"{str(datetime.datetime.now() - start_time)}", clr="w,b")


def archive_before(file_bytes, name, year, dest):
    clrprint("Archiving ", name, " before ", year, "...", sep="", clr="w,y,w,m,w")
    year_date = datetime.datetime(year, 1, 1)

    archive_wb = wb_data_from_bytes(file_bytes)
    archive_ws = archive_wb.get_sheet_by_name(archive_wb.sheetnames[0])
    clrprint("Getting rows in ", name, "...", sep="", clr="w,y,w")

    blank = 0
    row_count = 1
    temp_sheet = archive_wb.create_sheet(f'', index=0)
    for row in archive_ws.iter_rows(min_col=1, max_col=11, min_row=1):
        try:
            date = pd.to_datetime(row[5].value)
        except (IndexError, pd._libs.tslibs.parsing.DateParseError):
            copy_row(row, row_count, temp_sheet)
            row_count += 1
            continue
        if date is None:
            blank += 1
            if blank > 9:
                break
            continue
        if date is None or date > year_date:
            continue
        copy_row(row, row_count, temp_sheet)
        row_count += 1
    
    archive_wb.remove_sheet(archive_ws)
    temp_sheet.title = archive_ws.title

    temp_sheet.column_dimensions = archive_ws.column_dimensions
    temp_sheet.row_dimensions = archive_ws.row_dimensions

    clrprint("Saving ", dest, "...", sep="", clr="w,m,w")
    archive_wb.save(os.path.abspath(dest))
    archive_wb.close()

def remove_before(file_bytes, name, year, dest):
    clrprint("Cleaning ", name, " before ", year, "...", sep="", clr="w,y,w,m,w")
    year_date = datetime.datetime(year, 1, 1)

    formulas = {}
    formula_wb = wb_from_bytes(file_bytes)
    formula_ws = formula_wb.get_sheet_by_name(formula_wb.sheetnames[0])

    temp_sheet = formula_wb.create_sheet(f'', index=0)
    clrprint("Getting formulas in ", name, "...", sep="", clr="w,y,w")
    for cell in formula_ws[2]:
        if cell.column > 11 or not str(cell.value).startswith("="):
            continue
        formulas[cell.column] = (cell.value, cell.coordinate)

    print("Add these:", formulas)
    # TODO add these: formulas

    current_wb = wb_data_from_bytes(file_bytes)
    current_ws = current_wb.get_sheet_by_name(current_wb.sheetnames[0])
    clrprint("Getting rows in ", name, "...", sep="", clr="w,y,w")

    blank = 0
    row_count = 1
    for row in current_ws.iter_rows(min_col=1, max_col=11, min_row=1):
        try:
            date = pd.to_datetime(row[5].value)
        except (IndexError, pd._libs.tslibs.parsing.DateParseError):
            copy_row(row, row_count, temp_sheet)
            row_count += 1
            continue
        if date is None:
            blank += 1
            if blank > 9:
                break
            continue
        if date < year_date:
            continue
        copy_row(row, row_count, temp_sheet)
        row_count += 1
    
    formula_wb.remove_sheet(formula_ws)
    temp_sheet.title = formula_ws.title

    temp_sheet.column_dimensions = formula_ws.column_dimensions
    temp_sheet.row_dimensions = formula_ws.row_dimensions

    clrprint("Saving ", dest, "...", sep="", clr="w,m,w")
    current_wb.save(os.path.abspath(dest))
    current_wb.close()

def extract_years(path, year, dst_folder, use_name=True):
    name = get_name_from_file(path)
    clrprint("Loading ", name, "...", sep="", clr='w,y,w')

    file_bytes = load_path_bytes(path)
    #archive_before(file_bytes, name, year, f"{dst_folder}/{name}{year - 1}.xlsx")
    remove_before(file_bytes, name, year, f"{dst_folder}/{name}{year}.xlsx")

    return
    clrprint("Updating ", name, " to ", year, "...", sep="", clr="w,y,w,m,w")

    current_wb = wb_data_from_bytes(file_bytes)
    current_ws = current_wb.get_sheet_by_name(current_wb.sheetnames[0])
    for row in current_ws.iter_rows(min_col=1, max_col=6, min_row=2):
        try:
            date = pd.to_datetime(row[5].value)
        except (IndexError, pd._libs.tslibs.parsing.DateParseError):
            continue
        if date is None:
            continue
        if date < year_date:
            current_ws.delete_rows(row[0].row, 1)
            clrprint("Deleted", row[0].row, "from", name, clr="w,r,w,y")
    clrprint("Saving ", path.replace('\\', '/'), "...", sep="", clr='w,b,w')
    # current_wb.save(path)
    current_wb.close()

    # old_wb = wb_from_bytes(file_bytes)
    # old_wb.close()
    return
    dst_path = f"{name}{append}.xlsx"
    final_path = os.path.join(os.path.abspath(dst_folder), dst_path)

    print(f"Loading ", end="")
    clrprint(name, end="", clr='y')
    print("...")

    template = wb_from_bytes(template_bytes)

    print(f"Creating ", end="")
    clrprint(dst_path, end="", clr='m')
    print("...")

    ws = template[template.sheetnames[0]]
    for col in ws.iter_cols():
        if str(col[0].value).lower() == "name":
            col[0].value = name if use_name else "Name"
            break
    
    #ws.column_dimensions.group(start='H', end='K', hidden=True)

    template.save(final_path)
    template.close()

    print("Created ", end="")
    clrprint(final_path, clr='g')

def extract_all(src, year):
    start_time = datetime.datetime.now()
    dst_folder = f'Archive'#{year}'
    if not os.path.exists(dst_folder):
        os.makedirs(dst_folder)
    
    processes = []
    for file in os.listdir(src):
        if file_is_valid(file):
            p = Process(target=extract_years, args=(os.path.join(src, file), year, dst_folder))
            processes.append(p)
            p.start()

    for p in processes:
        p.join()

    clrprint("Finished in ", f"{str(datetime.datetime.now() - start_time)}", clr="w,b")

if __name__ == "__main__":
    Tk().withdraw()
    freeze_support()
    print("Welcome to ", end="")
    clrprint("Operator Tool", clr='m')

    print("Choose a starting folder")
    src_dir = './data'#folder_dialog()
    print(src_dir)

    if src_dir == "":
        sys.exit()

    # print("Choose a target folder")
    # dst_dir = folder_dialog()
    # print(dst_dir)

    # if dst_dir == "":
    #     sys.exit()

    # if template == "":
    #     sys.exit()

    try:
        restore_focus()
        year = int('balls')#input("Input a year "))
    except ValueError as e:
        clrprint("Error, defaulting to current year", clr="r")
        year = datetime.date.today().year
    print(year)

    # if True:
    #     empty_dir(dst_dir)

    try:
        result = extract_all(src_dir, year)
    except KeyboardInterrupt:
        print("Program terminated, new files might be corrupted")
    #os.startfile(result)
