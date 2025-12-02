from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json
import pandas as pd


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def add_error(list,cnpj,name,type):
        list.append({
        'CNPJ':cnpj,
        'Name':name,
        'Error_type': type
    })
        
def format_data_list(data_list):
    formatted_string = ""
    for item in data_list:
        formatted_string += f"CNPJ: {item['CNPJ']}, Nome: {item['Name']},  TIPO DO ERRO: {item['Error_type']}\n"
    return formatted_string

    
# Le arquivo config.json
with open('WebDriverTest\config.json', 'r') as config_file:
    config = json.load(config_file)

# Define your download directory
download_directory = config.get('default_dowload_directory')

# Load the Excel file, skip the header, and select columns by index
df = pd.read_excel('PIPELINE 08-10.xlsx', header=None)

# Select columns by index (assuming index 0-based, so A=0, B=1, ..., AP=41)
df = df.iloc[1:, 0:42]  # Skip the header, and select columns from A to AP

# # Display the DataFrame (optional)
# print(df.head())

# Create a WebDriverManager instance
driver_manager = WebDriverManager(download_directory)

# Create a WebDriver instance for your automation
navegador = driver_manager.create_driver()


# Now you can use the 'navegador' WebDriver instance in your automation script
navegador.get("https://apps.powerapps.com/play/e/default-8699d538-fe90-46fa-9de1-9e1369caa465/a/c694c53e-1dd5-4849-a052-d773f198c6b6?tenantId=8699d538-fe90-46fa-9de1-9e1369caa465")
time.sleep(5)

iframe = navegador.find_element(By.TAG_NAME, "iframe")
navegador.switch_to.frame(iframe)

error_list = []

# data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
# # Loop through the DataFrame rows and fill in the fields using column index
for index, row in df.iterrows():
    
    if len(str(row[0])) != 14:
        add_error(error_list,row[0],row[1],"CNPJ INVALIDO")
        continue
    
    print(row[1])
  
    # CNPJ
    cnpj_input = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[2]/div/div/div[18]/div/div/div/div/input')
    cnpj_input.click()
    cnpj_input.send_keys(row[0])
    time.sleep(10)
    
    #EDIT
    edit_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[2]/div/div/div[47]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/div/div')
    edit_button.click()
    time.sleep(5)
    
    #Envio proposta
    proposal_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[11]/div/div/div/div/button/div')
    proposal_button.click()
    time.sleep(2)
    
    # Ativar
    activate_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[24]/div/div/div/div/div[2]/div[2]')
    activate_button.click()
    time.sleep(1)
    
    # Ticket Medio
    ticket_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[26]/div/div/div/div/input')
    ticket_element.click()
    time.sleep(1)
    
    # Peso Medio
    weight_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[28]/div/div/div/div/input')
    weight_element.click()
    time.sleep(1)
    
    #Proximo 
    next_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[34]')
    next_button.click()
    time.sleep(1)
    
    # Ativar
    activate_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[24]/div/div/div/div/div[2]/div[2]')
    activate_button.click()
    time.sleep(1)
    
    # Shares 
    shares_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[29]/div/div/div/div/input')
    shares_element.click()
    time.sleep(1)
    
    #Proximo 
    next_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[43]')
    next_button.click()
    time.sleep(1)
    
    # Ativar
    activate_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[7]/div/div/div[24]/div/div/div/div/div[2]/div[2]')
    activate_button.click()
    time.sleep(1)
    
    # ID JMS
    id_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[7]/div/div/div[28]/div/div/div/div/input')
    id_element.click()
    time.sleep(1)
    
    # Contrato
    contract_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[7]/div/div/div[30]/div/div/div/div/div/div[1]')
    contract_element.click()
    time.sleep(10000)
    
    # Salvar
    save_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[7]/div/div/div/div/button')
    save_element.click()
    time.sleep(5)
    
    try:
        # Salvar 2
        save2_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[3]/div/div/div/div/div/div/div/div[5]/div/div/div/div/button')
        save2_element.click()
        time.sleep(10)
    except:
        add_error(error_list,row[0],row[1],"ERRO NA EDICAO DE CAMPOS")
        continue
    
time.sleep(5)
navegador.quit()


if not error_list:
    print("No error")
    
else:
    print(error_list)
    # SETUP
    smtp_server = 'smtp.hostinger.com'  # Adjusted to Hostinger's SMTP server
    port = 465  # SSL port for Hostinger
    sender_email = 'ti@portaljtdriverspr.com.br'
    password = 'J&Tex2024'  # Ensure this is secure and consider using environment variables or secure vaults

    # PARA
    receiver_emails = ['rafael.bueno@jtexpress.com.br']#'luana.silva@jtexpress.com.br'

    # Messages
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = ', '.join(receiver_emails)
    message['Subject'] = 'LOG BOT PIPELINE GANHO'

    default_body = "Bom dia\nExecução do Bot finalizada, problemas encontrados nos seguintes Leads:\n\n"
    data_body = format_data_list(error_list)
    body = default_body + data_body
    message.attach(MIMEText(body, 'plain'))

    # Connect to the SMTP server and send the email
    try:
        with smtplib.SMTP_SSL(smtp_server, port) as smtp:
            smtp.login(sender_email, password)  # Login
            smtp.sendmail(sender_email, receiver_emails, message.as_string())  # Send email
        print("Email sent successfully!")
    except smtplib.SMTPException as e:
        print("Unable to send email:", e)

    quit()