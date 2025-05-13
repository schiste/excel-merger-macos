import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog
import sys

def merge_excel_files():
    print("Launching file selector for file 1…")
    root = tk.Tk()
    root.withdraw()

    file1_path = filedialog.askopenfilename(title="Select first Excel file", filetypes=[("Excel files", "*.xlsx")])
    if not file1_path:
        print("No file selected for file 1. Exiting.")
        sys.exit()
    print(f"Selected file 1: {file1_path}")

    print("Launching file selector for file 2…")
    file2_path = filedialog.askopenfilename(title="Select second Excel file", filetypes=[("Excel files", "*.xlsx")])
    if not file2_path:
        print("No file selected for file 2. Exiting.")
        sys.exit()
    print(f"Selected file 2: {file2_path}")

    print("Loading file 1 workbook for editing…")
    wb1 = openpyxl.load_workbook(file1_path)
    writer = pd.ExcelWriter("merged.xlsx", engine="openpyxl", mode="w")
    writer._book = wb1
    writer._sheets = {ws.title: ws for ws in wb1.worksheets}

    print("Copying sheets from file 2…")
    xls2 = pd.ExcelFile(file2_path)
    for sheet_name in xls2.sheet_names:
        df = pd.read_excel(xls2, sheet_name=sheet_name)
        new_name = sheet_name
        while new_name in writer._sheets:
            new_name += "_copy"
        df.to_excel(writer, sheet_name=new_name, index=False)
        print(f"Added sheet from file2: {new_name}")

    writer.close()
    print("Merged file saved as: merged.xlsx")

if __name__ == "__main__":
    merge_excel_files()
