from argparse import ArgumentParser
from json import load, dump
from os import path
from shutil import copy
from sys import exit
from pyperclip import paste
from openpyxl import load_workbook

DEFAULT_ABS_PATH = path.dirname(path.abspath(__file__)) + "/"
DEFAULT_FILE_NAME = path.splitext(path.basename(path.abspath(__file__)))[0]
DEFAULT_CONFIG = {
    "path": DEFAULT_ABS_PATH + DEFAULT_FILE_NAME,
    "content": paste().strip(),
    "start": "2:5",
    "end": ":5",
    "copy": False,
}

def create_config_template(output_path):
    with open(output_path, "w", encoding='utf-8') as f:
        dump(DEFAULT_CONFIG, f, indent=4, ensure_ascii=False)

def read_config_file(config_file_path):
    try:
        with open(config_file_path, "r", encoding='utf-8') as f:
            config = load(f)
            return config
    except FileNotFoundError:
        return DEFAULT_CONFIG
    except Exception:
        exit()

def get_index_rc(ws, pos):
    if pos == "":
        exit()

    if pos.startswith(":"):
        if pos.endswith(":"):
            return ws.max_row, ws.max_column
        else:
            col = int(pos[1:])
            return ws.max_row, col if col > 0 else ws.max_column

    if pos.endswith(":"):
        row = int(pos[:-1])
        return row, ws.max_column if row > 0 else 1

    if ":" in pos:
        row, col = map(int, pos.split(":"))
        return row, col if row > 0 and col > 0 else 1

    try:
        num = int(pos)
        return num, num if num > 0 else 1
    except ValueError:
        exit()

def copy_file(excel_file_path):
    base_name, ext = path.splitext(excel_file_path)
    new_file_path = base_name + "_Copy" + ext
    copy(excel_file_path, new_file_path)
    return new_file_path

def main(args):

    if args.cc and not path.exists(args.config):
        create_config_template(args.config)
    config = read_config_file(args.config)

    excel_file = (args.p if args.p else config["path"])

    if (args.copy if args.copy else config["copy"]):
        excel_file = copy_file(excel_file)

    try:
        wb = load_workbook(excel_file)
        ws = wb.active
    except FileNotFoundError:
        exit()
    except Exception:
        exit()

    content = args.c if args.c else config["content"]
    if not content:
        exit()
    
    start_cell = args.s if args.s else config["start"]
    end_cell = args.e if args.e else config["end"]
    start_row, start_col = get_index_rc(ws, start_cell)
    end_row, end_col = get_index_rc(ws, end_cell)

    for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
        for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value is not None and cell.value != "":
                cell.value = content

    wb.save(excel_file)

if __name__ == "__main__":
    parser = ArgumentParser(description="Batch modification of Excel (*.xlsx) files")

    parser.add_argument("-p", help="excel (.xlsx) file path, default: " + DEFAULT_FILE_NAME + ".xlsx", default=DEFAULT_CONFIG["path"] + ".xlsx")
    parser.add_argument("-c", help="content to modify, default: " + DEFAULT_CONFIG["content"], default=DEFAULT_CONFIG["content"])
    parser.add_argument("-s", help="starting cell, format Row:Column. Default: " + DEFAULT_CONFIG["start"], default=DEFAULT_CONFIG["start"])
    parser.add_argument("-e", help="ending cell, format Row:Column. Default: " + DEFAULT_CONFIG["end"], default=DEFAULT_CONFIG["end"])
    parser.add_argument("-copy", action="store_true", help="create a copy of the Excel (.xlsx) file")
    parser.add_argument("-cc", action="store_true", help="create a JSON configuration file template (skipped if already exists), default: " + DEFAULT_FILE_NAME + ".json")
    parser.add_argument("-config", help="JSON configuration file, default: " + DEFAULT_FILE_NAME + ".json", default=DEFAULT_ABS_PATH + DEFAULT_FILE_NAME + ".json")

    args = parser.parse_args()

    main(args)