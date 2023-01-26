import pandas as pd

# Load the raw data into a pandas DataFrame, skipping the first 3 rows
raw_data = pd.read_excel("clean_data.xlsx")

# Extract the columns you need
ne_data = raw_data[["NE Name", "NE Type (MPU Type)", "Software Version", "Patch Version List", "Subnet Path"]]

# Rename the columns
ne_data = ne_data.rename(columns={"NE Name": "NE Name", "NE Type (MPU Type)": "NE Type", "Software Version": "Current Version", "Patch Version List": "Current patch", "Subnet Path": "Subnet Path"})

# Group the data by "NE Type"
grouped_data = ne_data.groupby("NE Type")

# Create one sheet for each NE Type
with pd.ExcelWriter("output.xlsx") as writer:  
    for group, data in grouped_data:
        data.to_excel(writer, sheet_name=group, index=False)
