from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# Setup webdriver
webdriver_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=webdriver_service)

# Navigate to the website
driver.get('url')  # replace with your target website

# Wait for 15 seconds
time.sleep(15)

# Scrape the HTML
html = driver.page_source

# Print the HTML
print(html)

# Close the browser
driver.quit()
