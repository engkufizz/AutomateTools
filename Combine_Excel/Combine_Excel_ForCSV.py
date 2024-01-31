import pandas as pd
import glob
import os

# Get the current working directory
cwd = os.getcwd()

# Create an empty dataframe to store the combined data
all_data = pd.DataFrame()

# Loop through all the CSV files in the current working directory that start with "Card Report_"
for f in glob.glob(cwd + "/Card Report_*.csv"):
    # Read each file as a dataframe, skipping the first 3 rows
    df = pd.read_csv(f, skiprows=3)
    # Append the dataframe to the all_data dataframe
    all_data = all_data.append(df, ignore_index=True)

# Save the combined data as a new CSV file
all_data.to_csv("Combined.csv", index=False)