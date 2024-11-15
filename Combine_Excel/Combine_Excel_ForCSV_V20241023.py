import pandas as pd
import glob
import os

# List of the specific 32 columns to detect
expected_columns = [
    "Serial No.", "Optical/Electrical Type", "NE Name", "Port Name", "Port Description", "Port Type",
    "Receive Optical Power(dBm)", "Reference Receive Optical Power (dBm)", "Reference Receive Time", "Receive Status",
    "Upper Threshold for Receive Optical Power(dBm)", "Lower Threshold for Receive Optical Power(dBm)", "Transmit Optical Power(dBm)",
    "Reference Transmit Optical Power (dBm)", "Reference Transmit Time", "Transmit Status", "Upper Threshold for Transmit Optical Power(dBm)",
    "Lower Threshold for Transmit Optical Power(dBm)", "SingleMode/MultiMode", "Speed(Mb/s)", "Wave Length(nm)",
    "Transmission Distance(m)", "Fiber Type", "Manufacturer", "Optical Mode Authentication", "Port Remark", "Port Custom Column",
    "OpticalDirectionType", "Vendor PN", "Model", "Rev(Issue Number)", "PN(BOM Code/Item)"
]

# Get the current working directory
cwd = os.getcwd()

# Create an empty dataframe to store the combined data
all_data = pd.DataFrame()

# Loop through all the CSV files in the current working directory that start with "Router_and_Switch_Optical&Electrical_Module_Report_"
for f in glob.glob(cwd + "/Router_and_Switch_Optical&Electrical_Module_Report_*.csv"):
    try:
        # Read the entire CSV as raw text to ensure no rows are skipped
        with open(f, 'r', encoding='utf-8') as file:
            raw_data = file.readlines()

        # Identify the row where the header starts by checking for the expected columns
        header_row_index = None
        for i, row in enumerate(raw_data):
            if all(col in row for col in expected_columns):
                header_row_index = i
                break
        
        if header_row_index is not None:
            # Re-read the data starting from the detected header row
            df = pd.read_csv(f, skiprows=header_row_index)
        else:
            # If the header is not detected, assume the first row is the header
            df = pd.read_csv(f)

        # If any columns are missing, add them with NaN values to maintain consistent structure
        for col in expected_columns:
            if col not in df.columns:
                df[col] = pd.NA
        
        # Append the dataframe to the all_data dataframe
        all_data = pd.concat([all_data, df], ignore_index=True)

    except Exception as e:
        print(f"An error occurred while reading {f}: {e}")

# Ensure all columns are in the expected order and fill any missing values
all_data = all_data.reindex(columns=expected_columns, fill_value=pd.NA)

# Save the combined data as a new CSV file
all_data.to_csv("Combined.csv", index=False)