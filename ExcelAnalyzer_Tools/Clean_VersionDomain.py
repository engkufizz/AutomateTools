import pandas as pd

# read in the Excel file
df = pd.read_excel('raw_data.xlsx', skiprows=3)

# remove the text within the brackets
df['Software Version'] = df['Software Version'].str.replace(r'\(.*\)', '', regex=True)

# Extract the version string in the format "V\d+R\d+C\d+SPC\d+" or "V\d+R\d+C\d+"
df["Software Version"] = df["Software Version"].str.extract("(V\d+R\d+C\d+(SPC\d+)?)")

# Replace 'ROOT/' and everything after '/' in column 'E' with an empty string
df['Subnet Path'] = df['Subnet Path'].str.replace('ROOT/', '').str.split('/').str[0]

# write the final modified data to a new Excel file
df.to_excel('clean_data.xlsx', index=False)
