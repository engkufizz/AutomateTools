from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

# Setup webdriver
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=chrome_options)

# List of URLs to scrape
urls = ['url1', 
        'url2', 
        'url3']

# Iterate over each URL
for url in urls:
    # Navigate to the website
    driver.get(url)

    # Wait for 5 seconds
    sleep(5)

    # Locate elements by xpath
    elements = driver.find_elements(By.XPATH, '//td[@name="patchName"]')

    # Print the text attribute of each element
    for element in elements:
        print(element.text)

# Close the driver
driver.quit()
