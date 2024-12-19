import os
import pandas as pd

# Initialize empty DataFrames to hold combined data
combined_services = pd.DataFrame()
combined_statistics = pd.DataFrame()

# Get all Excel files in the current directory
excel_files = [file for file in os.listdir('.') if file.endswith('.xlsx') or file.endswith('.xls')]

# First pass: collect all unique columns from Statistics sheets
all_stat_columns = set()
for file in excel_files:
    try:
        excel_data = pd.ExcelFile(file)
        if "Statistics" in excel_data.sheet_names:
            stat_data = pd.read_excel(file, sheet_name="Statistics")
            all_stat_columns.update(stat_data.columns)
    except Exception as e:
        print(f"Error processing file {file} during column collection: {e}")

# Convert to list and ensure "Device Name" is the first column
all_stat_columns = list(all_stat_columns)
if "Device Name" in all_stat_columns:
    all_stat_columns.remove("Device Name")
all_stat_columns = ["Device Name"] + sorted(all_stat_columns)

# Second pass: read and combine data
for file in excel_files:
    try:
        excel_data = pd.ExcelFile(file)
        
        # Process Services sheet
        if "Services" in excel_data.sheet_names:
            services_data = pd.read_excel(file, sheet_name="Services")
            services_data['Source File'] = file
            combined_services = pd.concat([combined_services, services_data], ignore_index=True)
        
        # Process Statistics sheet
        if "Statistics" in excel_data.sheet_names:
            stat_data = pd.read_excel(file, sheet_name="Statistics")
            # Fill missing columns with 0
            for col in all_stat_columns:
                if col not in stat_data.columns:
                    stat_data[col] = 0
            # Reorder columns to match all_stat_columns
            stat_data = stat_data.reindex(columns=all_stat_columns)
            stat_data['Source File'] = file
            combined_statistics = pd.concat([combined_statistics, stat_data], ignore_index=True)
            
    except Exception as e:
        print(f"Error processing file {file}: {e}")

# Save both sheets to a single Excel file
output_file = 'Combined_Data.xlsx'
with pd.ExcelWriter(output_file) as writer:
    combined_services.to_excel(writer, sheet_name='Services', index=False)
    combined_statistics.to_excel(writer, sheet_name='Statistics', index=False)

print(f"Combined data saved to {output_file}")

# Print summary of statistics columns
print("\nStatistics columns found across all files:")
for col in all_stat_columns:
    print(f"- {col}")
