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
df = df.iloc[129:, 0:42]  # Skip the header, and select columns from A to AP

# # Display the DataFrame (optional)
# print(df.head())

# Create a WebDriverManager instance
driver_manager = WebDriverManager(download_directory)

# Create a WebDriver instance for your automation
navegador = driver_manager.create_driver()


# Now you can use the 'navegador' WebDriver instance in your automation script
navegador.get("http://177.220.131.194:3050/#Contact")
time.sleep(5)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Loop through the DataFrame rows and fill in the fields using column index
for index, row in df.iterrows():
    print(index)
    criar_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[1]/div/div[2]/div/a')
    criar_button.click()
    time.sleep(1)
    
    #Pipe line
    # Fill fields by using column index (e.g., first column, second     column, etc.)
    #NOME
    nome_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div/div[1]/div[1]/div/div/div[3]/input')
    nome_element.click()
    nome_element.send_keys(row[4] + str(index))
    time.sleep(1)
    
    #EMAIL
    email_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div/div[2]/div[1]/div/div/div/input')
    email_element.click()
    email_element.send_keys(row[6])
    email_element.send_keys(Keys.ENTER)
    time.sleep(0.2)
    
    #CONTATO
    contato_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div/div[1]/div[2]/div/div[2]/span/button')
    contato_element.click()
    time.sleep(0.1)
    
    #MORE
    more_element = navegador.find_element(By.XPATH, '/html/body/div[10]/div/div/div/div/div[1]/div[1]/div[1]/div/div[3]/button[1]')
    more_element.click()
    time.sleep(0.1)
    
    #CNPJ
    cnpj_element = navegador.find_element(By.XPATH, '/html/body/div[10]/div/div/div/div/div[1]/div[1]/div[1]/div/div[3]/ul/li[22]/a')
    cnpj_element.click()
    time.sleep(0.1)
    cnpj_input = navegador.find_element(By.XPATH, '/html/body/div[10]/div/div/div/div/div[1]/div[2]/div/div/div/input')
    cnpj_input.click()
    cnpj_input.send_keys(row[0])
    time.sleep(0.1)
    
    #SEARCH
    search_element = navegador.find_element(By.XPATH,'/html/body/div[10]/div/div/div/div/div[1]/div[1]/div[1]/div/div[2]/button')
    search_element.click()
    time.sleep(0.5)
    
    #ALL
    all_element = navegador.find_element(By.XPATH, '/html/body/div[10]/div/div/div/div/div[2]/div[2]/table/thead/tr/th[1]/span/input')
    all_element.click()                             
    time.sleep(0.1)
    
    #SELECT
    select_element = navegador.find_element(By.XPATH,'/html/body/div[10]/div/div/div/footer/div[2]/button[1]')
    select_element.click()
    time.sleep(0.5)
    
    #SAVE
    save_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[2]/div/div[1]/div/button[1]')
    save_element.click()
    time.sleep(1)

    navegador.quit()
     
    # Create a WebDriverManager instance
    driver_manager = WebDriverManager(download_directory)

    # Create a WebDriver instance for your automation
    navegador = driver_manager.create_driver()


    # Now you can use the 'navegador' WebDriver instance in your automation script
    navegador.get("http://177.220.131.194:3050/#Contact")
    time.sleep(2)

# Close the browser when done
navegador.quit()
