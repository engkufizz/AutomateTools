import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# Replace "url_here" with the URL of the webpage you want to refresh
url = "url_here"

# Replace "path_to_chromedriver" with the path to your ChromeDriver executable
service = Service("path_to_chromedriver")
driver = webdriver.Chrome(service=service)

# Set the initial URL and refresh time
driver.get(url)
refresh_time = 60  # in seconds

# Continuously refresh the page every refresh_time seconds
while True:
    time.sleep(refresh_time)
    driver.refresh()
