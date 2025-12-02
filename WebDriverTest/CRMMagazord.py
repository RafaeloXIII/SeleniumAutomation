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
df = pd.read_excel('magazord - clientes geral.xlsx', header=None)

# Select columns by index (assuming index 0-based, so A=0, B=1, ..., AP=41)
df = df.iloc[1:, 0:42] # Skip the header, and select columns from A to AP

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

#INPUT
input_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[1]/div/input')
input_element.click()
input_element.send_keys("1609")
input_element.send_keys(Keys.ENTER)
time.sleep(2)

#OPTION
option_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[3]/div[2]/table/tbody/tr/td[2]/a')
option_element.click()
time.sleep(0.5)

#EDIT
edit_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[2]/div/div[1]/div/button[1]')
edit_element.click()
time.sleep(2)

# Loop through the DataFrame rows and fill in the fields using column index
for index, row in df.iterrows():
    print(index)
    #SELECT
    select_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[3]/div[2]/div[2]/div[2]/div/div[2]/span/button')
    select_element.click()
    time.sleep(0.5)

    #MORE
    more_element = navegador.find_element(By.XPATH,'/html/body/div[11]/div/div/div/div/div[1]/div[1]/div[1]/div/div[3]/button[1]')
    more_element.click()
    time.sleep(0.5)

    #COD JMS
    code_element = navegador.find_element(By.XPATH,'/html/body/div[11]/div/div/div/div/div[1]/div[1]/div[1]/div/div[3]/ul/li[4]/a')
    code_element.click()
    time.sleep(0.5)

    #VALOR
    value_element = navegador.find_element(By.XPATH,'/html/body/div[11]/div/div/div/div/div[1]/div[2]/div/div/div/input')
    value_element.click()
    value_element.send_keys(row[2])
    time.sleep(0.5)
    
    #SEARCH
    search_element = navegador.find_element(By.XPATH,'/html/body/div[11]/div/div/div/div/div[1]/div[1]/div[1]/div/div[2]/button')
    search_element.click()
    time.sleep(1)
    
    try:
        #CHOOSE
        choose_element = navegador.find_element(By.XPATH,'/html/body/div[11]/div/div/div/div/div[2]/div[2]/table/thead/tr/th[1]/span/input')
        choose_element.click()
        time.sleep(0.5)
    except:
        print('FAIL')
        #CANCEL
        cancel_element = navegador.find_element(By.XPATH,'/html/body/div[11]/div/div/div/footer/div[2]/button[2]')
        cancel_element.click()
        time.sleep(0.5)
        continue
    
    #CONFIRM
    confirm_element = navegador.find_element(By.XPATH,'/html/body/div[11]/div/div/div/footer/div[2]/button[1]')
    confirm_element.click()
    time.sleep(0.5)
    

#SAVE
save_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[2]/div/div[1]/div/button[1]')
save_button.click()

# Close the browser when done
time.sleep(5000)
navegador.quit()
