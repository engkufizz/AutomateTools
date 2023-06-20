from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd

# Setup webdriver
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=chrome_options)

# List of URLs to scrape and corresponding NE Types
urls_ne_types = [('utl1', 'ne1'), 
        ('url2', 'ne2'), 
        ('url3', 'ne3')]

# Create an empty DataFrame
df = pd.DataFrame(columns=['Patch Name', 'NE Type'])

# Iterate over each URL
for url, ne_type in urls_ne_types:
    # Navigate to the website
    driver.get(url)

    # Wait for 5 seconds
    sleep(5)

    # Locate elements by xpath
    elements = driver.find_elements(By.XPATH, '//td[@name="patchName"]')

    # Append the text attribute of each element and NE Type to the DataFrame
    for element in elements:
        df = df.append({'Patch Name': element.text, 'NE Type': ne_type}, ignore_index=True)

# Close the driver
driver.quit()

# Export the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)
