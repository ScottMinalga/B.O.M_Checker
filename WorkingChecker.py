import pandas as pd
import string
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Ask the user to select the first Excel file
Tk().withdraw()  # hide the Tkinter root window
file1 = askopenfilename(title='Select the first Excel file',
                        filetypes=[('Excel Files', '*.xlsx')])

# Ask the user to select the second Excel file
file2 = askopenfilename(title='Select the second Excel file',
                        filetypes=[('Excel Files', '*.xlsx')])

# Read in the two Excel sheets
sheet1 = pd.read_excel(file1, skiprows=5)
sheet2 = pd.read_excel(file2)

# Extract the column containing the part numbers from each sheet
col1 = sheet1['Part Number'].apply(lambda x: x.translate(str.maketrans('', '', string.punctuation)))
col2 = sheet2['Part Number'].apply(lambda x: x.translate(str.maketrans('', '', string.punctuation)))

# Compare the two columns to identify any differences
diff = set(col1) - set(col2)

# Output the results
if diff:
    print('The following part numbers are in Syspro but not in the B.O.M:')
    for part_num in diff:
        print(part_num)
if diff:
    print('The part numbers in Syspro match those in B.O.M.')

    # Print the list of missing parts
    missing_parts = set(col2) - set(col1)
    if missing_parts:
        print('\n\nThe following part numbers are in the B.O.M but not in Syspro:')
        for part_num in missing_parts:
            print(part_num)
