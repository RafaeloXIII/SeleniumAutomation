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
df = pd.read_excel('Leads_Sul_Ate 18-09-24- em andamento.xlsx', header=None)

# Select columns by index (assuming index 0-based, so A=0, B=1, ..., AP=41)
df = df.iloc[1:, 0:42]  # Skip the header, and select columns from A to AP


# # Display the DataFrame (optional)
# print(df.head())

# Create a WebDriverManager instance
driver_manager = WebDriverManager(download_directory)

# Create a WebDriver instance for your automation
navegador = driver_manager.create_driver()


# Now you can use the 'navegador' WebDriver instance in your automation script
navegador.get("http://177.220.131.194:3050/#Lead")
time.sleep(5)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Loop through the DataFrame rows and fill in the fields using column index
for index, row in df.iterrows():
    criar_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[1]/div/div[2]/div/a')
    criar_button.click()
    time.sleep(1)
    
    #Pipe line
    # Fill fields by using column index (e.g., first column, second     column, etc.)
    #NOME
    nome_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div/input')
    nome_element.click()
    nome_element.send_keys(row[2])
    time.sleep(0.1)
    #CNPJ
    cnpj_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div[2]/div[3]/div[2]/div/input')
    cnpj_element.click()
    cnpj_element.send_keys(row[0])
    time.sleep(0.1)
    
    #EMAIL
    email_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div/div/div/input')
    email_element.click()
    email_element.send_keys(row[6])
    time.sleep(0.1)
    
    #EMAIL
    nome_contato = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div/div/div[2]/input')
    nome_contato.click()
    nome_contato.send_keys(row[4])
    time.sleep(0.1)
    # #CIDADE
    # cidade_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div[2]/div[5]/div/div/input')
    # cidade_element.click()
    # cidade_element.send_keys(row[9])
    # time.sleep(0.1)
    # #UF
    # if row[8] == 'PR':
    #     estado = 'Paraná'
    # elif row[8] == 'SC':
    #     estado = 'Santa Catarina'
    # elif row[8] == 'RS':
    #     estado = 'Rio Grande do Sul'
    # else:
    #     estado = row[8]  # Keep the original value if it doesn't match PR, SC, or RS
        
    # uf_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[1]/div[2]/div[5]/div/div/div/div[2]/input')
    # uf_element.click()                             
    # uf_element.send_keys(estado)
    # time.sleep(0.1)
    
    #STATUS 
    status_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div/div/div[1]')
    status_element.click()                             
    time.sleep(0.5)
    
    if row[3] == 'Declínio':
        option_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div/div/div[2]/div/div[8]')
        option_element.click()
    elif row[3] == 'Em Andamento':
        option_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div/div/div[2]/div/div[4]')
        option_element.click()                            
    else:
        option_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div/div/div[2]/div/div[7]')
        option_element.click()
    # for _ in range(clicks):
    #     dropdown_element.send_keys(Keys.ARROW_DOWN)
    #     time.sleep(0.1)
    # time.sleep(0.5)
    
    #VOLUME
    if isinstance(row[14], int):
        
            volume_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div/input')
            volume_element.click()
            volume_element.send_keys(row[14])
            time.sleep(0.1)

    #RAMO ATIVIDADE
    ramo_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[3]/div[1]/div/div/div[1]/input')
    ramo_element.click()
    time.sleep(1)
    if pd.notna(row[12]):
        navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[3]/div[1]/div/div/div[1]/input').send_keys(row[12])
        navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[3]/div[1]/div/div/div[1]/input').send_keys(Keys.ENTER)
    time.sleep(0.1)
    #DESC.
    if pd.notna(row[38]):
        desc_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[2]/div[4]/div/div/textarea')
        desc_element.click()
        desc_element.send_keys(str(row[35]) + ";" + str(row[36]) + ";" + str(row[38]))
    time.sleep(0.5)
    #SAVE
    save_button = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[1]/div/button[1]')
    save_button.click()
    time.sleep(1)
    
    
    #LEADS
    leads_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div[1]/div/div[1]/h3/div/div[1]/span/a')
    leads_element.click()
    time.sleep(1)
    print(index)

# Close the browser when done
navegador.quit()
print("END")
