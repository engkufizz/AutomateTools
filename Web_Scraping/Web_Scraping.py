from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from time import sleep

# Setup webdriver
webdriver_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=webdriver_service)

# Navigate to the website
driver.get('url')

# Wait for 5 seconds
sleep(5)

# Locate elements by xpath
elements = driver.find_elements(By.XPATH, '//td[@name="patchName"]')

# Print the text attribute of each element
for element in elements:
    print(element.text)

# Close the driver
driver.quit()
