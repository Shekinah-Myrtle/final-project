import pandas as pd

# Path to your Excel file
file_path = 'CTPROJECTDB.xlsx'

# Read all sheets; returns a dict: { sheet_name: DataFrame, … }
all_sheets = pd.read_excel(file_path, sheet_name=None)

# Example: loop through and show basic info
for sheet_name, df in all_sheets.items():
    print(f"Sheet: {sheet_name!r} — {df.shape[0]} rows x {df.shape[1]} columns")