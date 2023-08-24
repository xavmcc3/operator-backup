from multiprocessing import Process, freeze_support
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter import Tk
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

def get_name_from_file(src):
    pattern = r'[0-9]'
    path = Path(src)

    rename = re.sub(pattern, '', path.stem)
    return rename.strip()

def load_template_bytes(path):
    print("Loading template bytes...")
    with open(path, 'rb') as f:
        template_bytes = BytesIO(f.read())
    return template_bytes

def new_template(template_bytes):
    return openpyxl.load_workbook(template_bytes)

def create_from_template(template_bytes, name, append, dst_folder, use_name=True):
    dst_path = f"{name}{append}.xlsx"
    final_path = os.path.join(os.path.abspath(dst_folder), dst_path)

    print(f"Loading ", end="")
    clrprint(name, end="", clr='y')
    print("...")

    template = new_template(template_bytes)

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
    template_bytes = load_template_bytes(template_path)
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

if __name__ == "__main__":
    Tk().withdraw()
    freeze_support()
    print("Welcome to ", end="")
    clrprint("Operator Tool", clr='m')

    print("Choose a starting folder")
    src_dir = folder_dialog()
    print(src_dir)

    print("Choose the template file")
    template = file_dialog([('Excel Files', ('.xls', '.xlsx'))])
    print(template)

    if src_dir == "":
        sys.exit()

    print("Choose a target folder")
    dst_dir = folder_dialog()
    print(dst_dir)

    if dst_dir == "":
        sys.exit()

    if template == "":
        sys.exit()

    try:
        restore_focus()
        year = int(input("Input a year "))
    except ValueError as e:
        clrprint("Error, defaulting to current year", clr="r")
        year = datetime.date.today().year
    print(year)

    # if True:
    #     empty_dir(dst_dir)

    try:
        create_all(template, src_dir, year, dst_dir)
    except KeyboardInterrupt:
        print("Program terminated, new files might be corrupted")
    os.startfile(dst_dir)
