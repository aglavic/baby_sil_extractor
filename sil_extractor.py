"""
Main script to run the program.
"""

import os
import zipfile

from dataclasses import dataclass
from datetime import datetime
from tkinter import Tk, filedialog
from typing import List, Self

import xlsxwriter

from xlsxwriter.worksheet import Worksheet


@dataclass
class Document:
    name: str
    size: int
    time: datetime

    @property
    def ftype(self):
        return self.name.rsplit(".", 1)[1].upper()

    def __eq__(self, other):
        return self.name == other.name

    def __lt__(self, other):
        return self.name < other.name


@dataclass
class Folder:
    name: str
    files: List[Document]
    sub_folders: List[Self]

    @property
    def depth(self):
        if len(self.sub_folders) > 0:
            return 1 + max([sf.depth for sf in self.sub_folders])
        else:
            return 1

    @property
    def num_files(self):
        return len(self.files) + sum([sf.num_files for sf in self.sub_folders])

    @property
    def sub_folder_names(self):
        return [sf.name for sf in self.sub_folders]

    def add_file(self, sub_paths: list, doc: Document):
        if len(sub_paths) == 0:
            self.files.append(doc)
        else:
            next_sp = sub_paths[0]
            if next_sp in self.sub_folder_names:
                sf = self.sub_folders[self.sub_folder_names.index(next_sp)]
            else:
                sf = Folder(next_sp, [], [])
                self.sub_folders.append(sf)
            sf.add_file(sub_paths[1:], doc)

    def sort(self):
        self.files.sort()
        for sf in self.sub_folders:
            sf.sort()


def analyze_zip(file_name: str):
    ffolder, fname = os.path.split(file_name)
    root = Folder(fname, [], [])
    if not file_name.lower().endswith(".zip"):
        raise ValueError("File has to be of type zip")

    fzi = zipfile.ZipFile(file_name, "r")
    for info in fzi.infolist():
        if info.is_dir():
            continue
        path_splitted = info.filename.split("/")
        doc = Document(path_splitted[-1], info.file_size, datetime(*info.date_time))
        root.add_file(path_splitted[:-1], doc)
    root.sort()
    return root


def write_xls_level(ws: Worksheet, folder: Folder, start_row: int, start_level: int, tree_depth: int):
    row = (
        ["" for i in range(start_level)]
        + [folder.name]
        + ["" for i in range(tree_depth - start_level)]
        + [folder.num_files, "folder"]
    )
    ws.write_row(start_row, 0, row)
    ws.set_row(start_row, None, None, {"level": start_level})
    start_row += 1
    for sf in folder.sub_folders:
        start_row = write_xls_level(ws, sf, start_row, start_level + 1, tree_depth)
    for fi in folder.files:
        row = ["" for i in range(tree_depth)] + [fi.name, fi.size, fi.ftype, fi.time]
        ws.write_row(start_row, 0, row)
        ws.set_row(start_row, None, None, {"level": start_level + 1})
        ws.write(start_row, tree_depth + 1, fi.size, size_format)
        ws.write(start_row, tree_depth + 3, fi.time, date_format)
        start_row += 1
    return start_row


date_format = None
size_format = None


def write_xls(data: List[Folder], out_file):
    global date_format, size_format
    wb = xlsxwriter.Workbook(out_file)
    date_format = wb.add_format({"num_format": "yyyy-m-d hh:mm"})
    size_format = wb.add_format({"num_format": "#,##0"})

    for root in data:
        ws = wb.add_worksheet(root.name[:20])
        ws.outline_settings(symbols_below=False)
        tree_depth = root.depth
        ws.write_row(0, 0, ["Folder"] + ["" for i in range(tree_depth - 1)] + ["Document", "Size", "Type", "Date"])
        total_rows = write_xls_level(ws, root, 1, 0, tree_depth)
        ws.autofilter(0, 0, total_rows, tree_depth + 3)

    wb.close()


def main():
    fls = filedialog.askopenfilenames(
        title="Sil Extractor - Select Source File",
        filetypes=[
            ("Zipfile", "*.zip"),
        ],
    )
    if len(fls) == 0:
        return
    output = []
    for fi in fls:
        root = analyze_zip(fi)
        output.append(root)
    out_file = filedialog.asksaveasfilename(
        title="Sil Extractor - Choose Destination", initialfile="output.xlsx", filetypes=[("Excel", "*.xlsx")]
    )
    if out_file is None:
        return
    write_xls(output, out_file)


if __name__ == "__main__":
    tk_root = Tk()
    tk_root.withdraw()
    main()
    tk_root.destroy()
