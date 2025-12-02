from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json
import pandas as pd

# Le arquivo config.json
with open('WebDriverTest\config.json', 'r') as config_file:
    config = json.load(config_file)

# Define your download directory
download_directory = config.get('default_dowload_directory')

# Load the Excel file, skip the header, and select columns by index
df = pd.read_excel('Leads_Sul - ganho(1).xlsx', header=None)

# Select columns by index (assuming index 0-based, so A=0, B=1, ..., AP=41)
df = df.iloc[97:, 0:42] # Skip the header, and select columns from A to AP

# # Display the DataFrame (optional)
# print(df.head())

# Create a WebDriverManager instance
driver_manager = WebDriverManager(download_directory)

# Create a WebDriver instance for your automation
navegador = driver_manager.create_driver()


# Now you can use the 'navegador' WebDriver instance in your automation script
navegador.get("http://177.220.131.194:3050/#Admin/portalUsers")
time.sleep(5)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Loop through the DataFrame rows and fill in the fields using column index
for index, row in df.iterrows():
    print(index)
    
    #CRIAR
    criar_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[1]/div/div[2]/div/a')
    criar_button.click()
    time.sleep(2)

    # #MORE
    # more_element = navegador.find_element(By.XPATH, '/html/body/div[4]/div/div/div/div/div[1]/div[1]/div[1]/div/div[3]/button[1]')
    # more_element.click()
    # time.sleep(0.1)
    
    #CONTA
    conta_element = navegador.find_element(By.XPATH, '/html/body/div[4]/div/div/div/div/div[1]/div[1]/div[1]/div/input')
    conta_element.click()
    conta_element.send_keys(row[6])
    time.sleep(0.1)
    
    # #INFO
    # info_element = navegador.find_element(By.XPATH, '/html/body/div[4]/div/div/div/div/div[1]/div[2]/div/div/div/div[2]/div[2]/input')
    # info_element.click()
    # info_element.send_keys(row[2])
    # time.sleep(0.5)
    # info_element.send_keys(Keys.ARROW_DOWN)
    # info_element.send_keys(Keys.ENTER)
    # time.sleep(0.1)
    
    #SEARCH
    search_element = navegador.find_element(By.XPATH, '/html/body/div[4]/div/div/div/div/div[1]/div[1]/div[1]/div/div[2]/button')
    search_element.click()
    time.sleep(2)
    try:
        #VALUE
        value_element = navegador.find_element(By.XPATH,'/html/body/div[4]/div/div/div/div/div[2]/div[2]/table/tbody/tr/td[1]')
        value = value_element.text
        print(value)
        time.sleep(0.1)
        #OPTION
        option_element = navegador.find_element(By.XPATH, '/html/body/div[4]/div/div/div/div/div[2]/div[2]/table/tbody/tr/td[2]/a')
        option_element.click()
        time.sleep(0.5)
    except:
        time.sleep(0.5)
        close_element = navegador.find_element(By.XPATH, '/html/body/div[4]/div/div/div/header/a/span')
        close_element.click()
        print("RETURN")
        time.sleep(0.5)
        continue
    
    # #USER
    # user_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div/div[1]/div[1]/div/input')
    # user_element.click()
    # user_element.send_keys(value) #UNICA COISA A DEFINIR
    # time.sleep(0.1)
    
    # #NAME
    # name_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div/div[2]/div[1]/div/div/div[1]')
    # name_element.click()
    # time.sleep(0.5)
    # info_element.send_keys(Keys.ARROW_DOWN)
    # info_element.send_keys(Keys.ENTER)
    # time.sleep(1000)
    
    # #EMAIL
    # email_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div/div[3]/div[1]/div/div/div/input')
    # email_element.click()
    # email_element.send_keys(row[6]) 
    # time.sleep(0.1)
    
    #PORTAL
    portal_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[3]/div[2]/div[1]/div[1]/div/div[2]/span/button[2]')
    portal_element.click()
    time.sleep(0.3)
    client_element = navegador.find_element(By.XPATH, '/html/body/div[11]/div/div/div/div/div[2]/div[2]/table/tbody/tr[2]/td[2]/a')
    client_element.click()
    time.sleep(0.1)
    
    #RULES
    rules_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[3]/div[2]/div[2]/div[1]/div/div[2]/span/button[2]')
    rules_element.click()
    time.sleep(0.3)
    report_element = navegador.find_element(By.XPATH, '/html/body/div[11]/div/div/div/div/div[2]/div[2]/table/tbody/tr[2]/td[2]/a')
    report_element.click()
    time.sleep(0.1)
    
    #GERAR
    gerar_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[5]/div[2]/div[1]/div[2]/div/button')
    gerar_element.click()
    time.sleep(0.5)   
    
    #SAVE
    save_button = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[1]/div/button[1]')
    save_button.click()
    time.sleep(10)
    
    navegador.quit()
    time.sleep(2)
    
    # Create a WebDriverManager instance
    driver_manager = WebDriverManager(download_directory)

    # Create a WebDriver instance for your automation
    navegador = driver_manager.create_driver()


    # Now you can use the 'navegador' WebDriver instance in your automation script
    navegador.get("http://177.220.131.194:3050/#Admin/portalUsers")
    time.sleep(5)


# Close the browser when done
navegador.quit()
