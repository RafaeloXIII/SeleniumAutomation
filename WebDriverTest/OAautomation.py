from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json
import csv
import pyperclip

# Le arquivo config.json
with open('WebDriverTest\config.json', 'r') as config_file:
    config = json.load(config_file)

# Define your download directory
download_directory = config.get('default_dowload_directory')

# Create a WebDriverManager instance
driver_manager = WebDriverManager(download_directory)

# Create a WebDriver instance for your automation
navegador = driver_manager.create_driver()

# Now you can use the 'navegador' WebDriver instance in your automation script
navegador.get("https://www.example.com")
time.sleep(2)
# Open a new tab
navegador.execute_script("window.open('');")
# Switch to the new tab
navegador.switch_to.window(navegador.window_handles[1])
time.sleep(5)
# Load a different page in the new tab
navegador.get('https://www.google.com')
time.sleep(5)
# Switch back to the first tab
navegador.switch_to.window(navegador.window_handles[0])
time.sleep(2)

# Continue working with the first tab
# ...

# Close the browser
navegador.quit()
