# Import pandas and glob
import pandas as pd
import glob

# Create an empty dataframe to store the combined data
all_data = pd.DataFrame()

# Loop through all the Excel files in the Input folder that start with "CommonCollectResult"
for f in glob.glob("Input/CommonCollectResult*.xlsx"):
    # Read each file as a dataframe
    df = pd.read_excel(f)
    # Append the dataframe to the all_data dataframe
    all_data = all_data.append(df, ignore_index=True)

# Save the combined data as a new Excel file
all_data.to_excel("Combined.xlsx", index=False)