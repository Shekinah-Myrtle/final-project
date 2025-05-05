import pandas as pd
import os
from openpyxl import load_workbook

# Path to your Excel file
file_path = 'CTPROJECTDB.xlsx'

# Read all sheets; returns a dict: { sheet_name: DataFrame, … }
all_sheets = pd.read_excel(file_path, sheet_name=None)

for sheet_name, df in all_sheets.items():
    print(f"Sheet: {sheet_name!r} — {df.shape[0]} rows x {df.shape[1]} columns")
input_file = 'CTPROJECTDB.xlsx'

# Load all sheets into a dictionary
all_sheets = pd.read_excel(input_file, sheet_name=None)

# Define material sheets manually or through pattern matching
material_sheets = [
    'Aggregates_Sand', 'Aluminium', 'Asphalt', 'Bitumen', 'Cement_and_Mortar',
    'Ceramic', 'Clay_Bricks', 'Concrete', 'Glass', 'Insulation', 'Paint',
    'Plaster', 'Rubber', 'Steel', 'Timber', 'Vinyl'
]

# Create a new Excel file to save filtered material data
output_file = 'Material_Data_Only.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name in material_sheets:
        # Check if the sheet exists in the original Excel file
        if sheet_name in all_sheets:
            # Extract the material data for the current sheet
            df = all_sheets[sheet_name]

            # Optionally, clean up the data (e.g., drop empty rows/columns)
            df_cleaned = df.dropna(how='all').dropna(axis=1, how='all')

            # Write the cleaned data to the new Excel file
            df_cleaned.to_excel(writer, sheet_name=sheet_name[:31], index=False)

print(f" Data from material sheets has been written to {output_file}")

# Path to your input Excel file
input_file = 'Material_Data_Only.xlsx'

# Define rows to drop for each material
rows_to_drop = {
    'Aggregates_Sand': list(range(0, 28)) + list(range(39, 49)) + list(range(85, 104)),
    'Aluminium': list(range(0, 30)) + list(range(59, 69)) + list(range(105, 124)) + list(range(132, 142)) + list(range(147, 299)) + list(range(384, 1931)),
    'Asphalt': list(range(0, 30)) + list(range(52, 62)) + list(range(66, 80)) + list(range(89, 117)) + list(range(122, 135)) + list(range(137, 138)) + list(range(144, 291)) + list(range(314, 1923)),
    'Bitumen': list(range(0, 28)) + list(range(32, 42)) + list(range(48, 61)) + list(range(78, 97)) + list(range(104, 118)) + list(range(121, 270)) + list(range(288, 1904)),
    'Cement_and_Mortar': list(range(0, 28)) + list(range(87, 97)) + list(range(105, 116)) + list(range(125, 152)) + list(range(161, 172)) + list(range(186, 325)) + list(range(398, 1957)),
    'Ceramic': list(range(0, 28)) + list(range(32, 42)) + list(range(47, 61)) + list(range(70, 99)) + list(range(104, 116)) + list(range(122, 134)) + list(range(144, 287)) + list(range(441, 1919)),
    'Clay_Bricks': list(range(0, 28)) + list(range(0, 60)) + list(range(66, 78)) + list(range(88, 115)) + list(range(122, 135)) + list(range(149, 288)) + list(range(307, 1920)),
    'Concrete': list(range(0, 30)) + list(range(310, 319)) + list(range(326, 338)) + list(range(345, 375)) + list(range(382, 395)) + list(range(625, 2180)),
    'Glass': list(range(0, 28)) + list(range(95, 104)) + list(range(113, 122)) + list(range(133, 160)) + list(range(169, 178)) + list(range(199, 333)) + list(range(525, 1965)),
    'Insulation': list(range(0, 28)) + list(range(34, 43)) + list(range(51, 62)) + list(range(72, 100)) + list(range(108, 118)) + list(range(126, 138)) + list(range(181, 292)) + list(range(1259, 1923)),
    'Paint': list(range(0, 28)) + list(range(30, 39)) + list(range(43, 57)) + list(range(68, 97)) + list(range(100, 114)) + list(range(118, 134)) + list(range(874, 1919)),
    'Plaster': list(range(0, 28)) + list(range(32, 41)) + list(range(47, 60)) + list(range(70, 98)) + list(range(104, 116)) + list(range(122, 136)) + list(range(161, 289)) + list(range(860, 1922)),
    'Rubber': list(range(0, 28)) + list(range(30, 39)) + list(range(43, 58)) + list(range(68, 96)) + list(range(100, 114)) + list(range(118, 134)) + list(range(140, 287)) + list(range(351, 1919)),
    'Steel': list(range(0, 28)) + list(range(48, 76)) + list(range(134, 143)) + list(range(151, 162)) + list(range(171, 199)) + list(range(2067, 219)) + list(range(224, 372)) + list(range(525, 2003)),
    'Timber': list(range(0, 28)) + list(range(74, 83)) + list(range(112, 139)) + list(range(153, 159)) + list(range(205, 312)) + list(range(526, 1944)),
    'Vinyl': list(range(0, 28)) + list(range(31, 40)) + list(range(45, 58)) + list(range(69, 97)) + list(range(102, 115)) + list(range(120, 135)) + list(range(138, 288)) + list(range(568, 1920))
}

# Load all sheets into a dictionary
all_sheets = pd.read_excel(input_file, sheet_name=None)

# Create a new Excel file to save updated material data
output_file = 'Updated_Material_Data.xlsx'

# Delete the file if it exists
if os.path.exists(output_file):
    os.remove(output_file)

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, df in all_sheets.items():
        if sheet_name in rows_to_drop:
            rows = [i for i in rows_to_drop[sheet_name] if i < len(df)]
            df = df.drop(df.index[rows]).reset_index(drop=True)
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

print(f" Updated material data has been written to {output_file}")

# Load the workbook
wb = load_workbook('Updated_Material_Data.xlsx')

# Iterate over each sheet
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    print(f"Processing sheet: {sheet_name}")
    
    # Delete columns A to D (1 to 4)
    for col in range(4, 0, -1):
        ws.delete_cols(col)
        print(f"Deleted column {col}")
    
    # Find the maximum number of columns after deleting A to D
    max_col = ws.max_column
    
    # Delete columns O to Z (15 to max_col)
    for col in range(max_col, 14, -1):
        ws.delete_cols(col)
        print(f"Deleted column {col}")

# Save the workbook as a new file in the current working directory
current_dir = os.getcwd()
new_file_name = 'Updated_Material_Data_New.xlsx'
wb.save(os.path.join(current_dir, new_file_name))

print(f"File saved as {new_file_name} in {current_dir}")