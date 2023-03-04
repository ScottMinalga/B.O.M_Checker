import pandas as pd
import tkinter as tk
from tkinter import filedialog


# Function to open a file dialog and get the path to an Excel file
def get_file_path():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path


# Function to compare two Excel files and find missing part numbers
def compare_excel_files():
    # Get the paths to the two Excel files
    file_path1 = get_file_path()
    file_path2 = get_file_path()

    # Read the Excel files into DataFrames
    df1 = pd.read_excel(file_path1, skiprows=5)
    df2 = pd.read_excel(file_path2, skiprows=5)

    # Extract the part number columns from each DataFrame
    part_numbers1 = df1['part_number']
    part_numbers2 = df2['part_number']

    # Find the set difference between the two part number sets
    missing_part_numbers = set(part_numbers1) - set(part_numbers2)

    # Print any missing part numbers
    if len(missing_part_numbers) == 0:
        output_text.set('No missing part numbers.')
    else:
        output_text.set('Missing part numbers:\n' + '\n'.join(str(part_number) for part_number in missing_part_numbers))


# Create the GUI
root = tk.Tk()
root.title('Excel Comparison Tool')

# Create a button to start the comparison
compare_button = tk.Button(root, text='Compare B.O.M', command=compare_excel_files)
compare_button.pack(padx=20, pady=20)

# Create a label to display the output
output_text = tk.StringVar()
output_label = tk.Label(root, textvariable=output_text, justify='left')
output_label.pack(padx=20, pady=20)

# Start the GUI event loop
root.mainloop()
