import os
import pandas as pd

# Initialize an empty DataFrame to hold combined data
combined_data = pd.DataFrame()

# Get all Excel files in the current directory
excel_files = [file for file in os.listdir('.') if file.endswith('.xlsx') or file.endswith('.xls')]

for file in excel_files:
    try:
        # Load the Excel file
        excel_data = pd.ExcelFile(file)
        
        # Check if the sheet "Services" exists in the file
        if "Services" in excel_data.sheet_names:
            # Read the "Services" sheet
            sheet_data = pd.read_excel(file, sheet_name="Services")
            sheet_data['Source File'] = file  # Add a column to track the source file
            combined_data = pd.concat([combined_data, sheet_data], ignore_index=True)
    except Exception as e:
        print(f"Error processing file {file}: {e}")

# Save the combined data to a new Excel file
output_file = 'Combined_Services.xlsx'
combined_data.to_excel(output_file, index=False)

print(f"Combined data saved to {output_file}")
