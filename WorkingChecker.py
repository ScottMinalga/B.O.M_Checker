import pandas as pd
import string
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Hide the root Tkinter window
Tk().withdraw()

# Ask user to select both Excel files
file1 = askopenfilename(title='Select the first Excel file', filetypes=[('Excel Files', '*.xlsx')])
file2 = askopenfilename(title='Select the second Excel file', filetypes=[('Excel Files', '*.xlsx')])

# Read Excel files, skipping header rows
sheet1 = pd.read_excel(file1, skiprows=5)
sheet2 = pd.read_excel(file2, skiprows=5)

# Clean and extract 'Stock code' and 'Quantity per' columns
col1 = sheet1['Stock code'].apply(lambda x: str(x).translate(str.maketrans('', '', string.punctuation)).lower())
col2 = sheet2['Stock code'].apply(lambda x: str(x).translate(str.maketrans('', '', string.punctuation)).lower())

# Compare part numbers
diff1 = set(col1) - set(col2)
diff2 = set(col2) - set(col1)

# Find parts with mismatched quantities
qty_mismatches = []
common_parts = set(col1) & set(col2)

for part in common_parts:
    qty1_vals = sheet1.loc[col1 == part, 'Quantity per'].values
    qty2_vals = sheet2.loc[col2 == part, 'Quantity per'].values
    if len(qty1_vals) > 0 and len(qty2_vals) > 0 and qty1_vals[0] != qty2_vals[0]:
        qty_mismatches.append((part, qty1_vals[0], qty2_vals[0]))

# Loop to show results until user chooses to exit
while True:
    print("\nHere is what is missing from BOM 1:")
    if diff1:
        for part_num in sorted(diff1):
            print(f"  - {part_num}")
    else:
        print("  Everything on BOM 1 is on BOM 2.")

    print("\nHere is what is missing from BOM 2:")
    if diff2:
        for part_num in sorted(diff2):
            print(f"  - {part_num}")
    else:
        print("  Everything on BOM 2 is on BOM 1.")

    print("\nHere are parts with mismatched quantities:")
    if qty_mismatches:
        for part, q1, q2 in qty_mismatches:
            print(f"  - {part}: BOM1 = {q1}, BOM2 = {q2}")
    else:
        print("  All common parts have matching quantities.")

    leave = input("\nPress 1 to exit or any other key to view the results again: ")
    if leave == "1":
        break
