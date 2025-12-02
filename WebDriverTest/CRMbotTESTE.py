from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json
import threading

# Le arquivo config.json

with open('WebDriverTest\config.json', 'r') as config_file:
    config = json.load(config_file)

# Define your download directory
download_directory = config.get('default_dowload_directory')

def crm_automation(url,download_directory):
        
    # Create a WebDriverManager instance
    driver_manager = WebDriverManager(download_directory)

    # Create a WebDriver instance for your automation
    navegador = driver_manager.create_driver()


    # Now you can use the 'navegador' WebDriver instance in your automation script
    navegador.get(url)
    time.sleep(30)

    data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    navegador.quit()

def powerapps_automation(url,download_directory):
        
    # Create a WebDriverManager instance    
    driver_manager = WebDriverManager(download_directory)

    # Create a WebDriver instance for your automation
    navegador = driver_manager.create_driver()
    
    # Now you can use the 'navegador' WebDriver instance in your automation script
    navegador.get(url)
    time.sleep(30)

    data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    navegador.quit()    
    
crm_url = "http://177.220.131.194:3050/#Opportunity"
powerapps_url = "https://apps.powerapps.com/play/e/default-8699d538-fe90-46fa-9de1-9e1369caa465/a/c694c53e-1dd5-4849-a052-d773f198c6b6?tenantId=8699d538-fe90-46fa-9de1-9e1369caa465"

thread1 = threading.Thread(target=crm_automation, args=(crm_url,download_directory))
thread2 = threading.Thread(target=powerapps_automation, args=(powerapps_url,download_directory))

thread1.start()
thread1.join()
thread2.start()
thread2.join()