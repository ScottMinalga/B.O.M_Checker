import pandas as pd

# Read the two Excel sheets into DataFrames
df1 = pd.read_excel('sheet1.xlsx')
df2 = pd.read_excel('sheet2.xlsx')

# Extract the part number columns from each DataFrame
part_numbers1 = df1['part_number']
part_numbers2 = df2['part_number']

# Find the set difference between the two part number sets
missing_part_numbers = set(part_numbers1) - set(part_numbers2)

# Print any missing part numbers
if len(missing_part_numbers) == 0:
    print('No missing part numbers.')
else:
    print('Missing part numbers:')
    for part_number in missing_part_numbers:
        print(part_number)

