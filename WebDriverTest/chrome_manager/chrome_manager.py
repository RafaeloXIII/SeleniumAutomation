
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService  
from webdriver_manager.chrome import ChromeDriverManager
import os
import json

with open(r'WebDriverTest\config.json', 'r') as config_file:
    config = json.load(config_file)

class WebDriverManager:
    def __init__(self, download_directory):
        self.download_directory = download_directory

    def create_driver(self):
        chrome_install = ChromeDriverManager().install()

        folder = os.path.dirname(chrome_install)
        chromedriver_path = os.path.join(folder, "chromedriver.exe")

        chrome_service = ChromeService(chromedriver_path)
        options = webdriver.ChromeOptions()
        options.add_argument(config.get("chrome_driver_path"))
        # #Opcional
        # options.add_argument("--headless")

        options.add_experimental_option("prefs", {
            "download.default_directory": self.download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        return webdriver.Chrome(service=chrome_service, options=options)
    
