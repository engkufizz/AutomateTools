import pandas as pd

# read in the CSV file
df = pd.read_csv('raw_data.csv', skiprows=3)

# remove the text within the brackets
df['Software Version'] = df['Software Version'].str.replace(r'\(.*\)', '', regex=True)

# extract the required string using regular expression
df['Software Version'] = df['Software Version'].str.extract(r'(V\d+R\d+C\d+(?:SPC\d+)?)')

# Replace 'ROOT/' and everything after '/' in column 'E' with an empty string
df['Subnet Path'] = df['Subnet Path'].str.replace('ROOT/', '').str.split('/').str[0]

# write the final modified data to a new CSV file
df.to_csv('clean_data.csv', index=False)
