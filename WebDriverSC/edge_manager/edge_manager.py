
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager 
import json

with open('WebDriverSC/config.json', 'r') as config_file:
    config = json.load(config_file)

class WebDriverManager:
    def __init__(self, download_directory):
        self.download_directory = download_directory

    def create_driver(self):
        servico = Service(EdgeChromiumDriverManager().install())
        options = webdriver.EdgeOptions()
        options.add_argument(r"C:\Users\Rafael Braga\AppData\Local\edgedriver_win64\msedgedriver.exe")
        options.add_argument("--user-data-dir=C:\\Users\\Rafael Braga\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default")
        # #Opcional
        # options.add_argument("--headless")

        options.set_capability("ms:edgeChromium", True)

        options.add_experimental_option("prefs", {
            "download.default_directory": self.download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        return webdriver.Edge(service=servico, options=options)
    
