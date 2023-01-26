import os
import re
import xlsxwriter

# Create a new Excel workbook and add a worksheet
workbook = xlsxwriter.Workbook('config_info.xlsx')
worksheet = workbook.add_worksheet()

# Add the column headers to the worksheet
worksheet.write('A1', 'NE name')
worksheet.write('B1', 'NE type')
worksheet.write('C1', 'NE version')
worksheet.write('D1', 'Patch version')

# Set the starting row for the data
row = 1

# Iterate through the list of files in the current directory
for filename in os.listdir():
    # Check if the file is a .txt file
    if filename.endswith('.txt'):
        # Open the config file and read the contents
        with open(filename, 'r') as f:
            config_contents = f.read()

        # Use regular expressions to extract the relevant information from the config file
        ne_name_match = re.search(r'sysname (\w+)', config_contents)
        ne_type_match = re.search(r'(ATN \w+) V\d+R\d+C\d+SPC\d+', config_contents)
        ne_version_match = re.search(r'V\d+R\d+C\d+SPC\d+', config_contents)
        patch_version_match = re.search(r'Patch Version: (V\d+R\d+SPH\d+)', config_contents)

        # Extract the matched strings from the regex matches
        ne_name = ne_name_match.group(1) if ne_name_match else ''
        ne_type = ne_type_match.group(1) if ne_type_match else ''
        ne_version = ne_version_match.group(0) if ne_version_match else ''
        patch_version = patch_version_match.group(1) if patch_version_match else ''

        # Add the extracted information to the worksheet
        worksheet.write(row, 0, ne_name)
        worksheet.write(row, 1, ne_type)
        worksheet.write(row, 2, ne_version)
        worksheet.write(row, 3, patch_version)

        # Increment the row counter
        row += 1

# Close the workbook
workbook.close()
