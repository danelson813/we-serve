###################################
# we-serve/helpers/style_cells.py #
# Author: dan.nelson1@gmail.com   #
###################################
from openpyxl import load_workbook
from pathlib import Path

# from finalize import file_path
from openpyxl.styles import Alignment

file_path = Path.cwd() / "data" / "final_report.xlsx"


def styling_wb(filename: file_path) -> None:
    wb = load_workbook("data/final_report.xlsx")
    ws = wb.active
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 13
    for row in range(1, 101):
        ws.row_dimensions[row].height = 20
    wb.save("data/final_report.xlsx")
