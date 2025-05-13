# Excel Sheet Merger

A simple Python tool with a file dialog interface to merge all sheets from two `.xlsx` files into a single workbook.

## Features

- Select two Excel files via UI
- Copies all sheets from both files into one output
- Automatically renames sheets to avoid name collisions
- Outputs a new `merged.xlsx` file
- Console logs for progress visibility

## Requirements

- Python 3.7+
- Packages:
  - `pandas`
  - `openpyxl`
  - `tkinter` (pre-installed on macOS)

Install required packages:

```bash
pip install pandas openpyxl
