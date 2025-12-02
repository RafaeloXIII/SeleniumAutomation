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
    
    # #ADD
    add_button = navegador.find_element(By.CSS_SELECTOR, "#publishedCanvas > div > div.screen-animation.animated > div > div > div:nth-child(32) > div > div > div > div")
    add_button.click() 
    time.sleep(5)

    #CNPJ
    cnpj_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[22]/div/div/div/div/input')
    cnpj_button.click()
    cnpj_button.send_keys(row[0])
    time.sleep(0.5)

    # Razao Social
    razao_social_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[29]/div/div/div/div/input')
    razao_social_button.click()
    razao_social_button.send_keys(row[1])
    time.sleep(0.5)

    # Nome Lead
    nome_lead_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[35]/div/div/div/div/input')
    nome_lead_button.click()
    nome_lead_button.send_keys(row[2] + ' ' + row[3])
    time.sleep(0.5)

    # Email Lead
    email_lead_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[40]/div/div/div/div/input')
    email_lead_button.click()
    email_lead_button.send_keys(row[4])
    time.sleep(0.5)

    # UF
    uf_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[44]/div/div/div/div/div/div[1]')
    uf_button.click()
    print(row[6])
    if str(row[6]) == 'Paraná':
        clicks = 18
    elif str(row[6]) == 'Rio Grande do Sul':
        clicks = 23
    elif str(row[6]) == 'Santa Catarina':
        clicks = 24
    for _ in range(clicks):
        uf_button.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.1)
    time.sleep(0.5)

    # Origem
    origem_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[46]/div/div/div/div/div/div[1]')
    origem_element.click()
    origem_element.send_keys(Keys.ARROW_DOWN)
    time.sleep(0.5)

    # Nomes Fantasia
    nomes_fantasia_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[31]/div/div/div/div/input')
    nomes_fantasia_button.click()
    nomes_fantasia_button.send_keys(row[1])
    time.sleep(0.5)
    
    # Cidade
    cidade_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[48]/div/div/div/div/input')
    cidade_button.click()
    cidade_button.send_keys(row[7])
    time.sleep(0.5)
    
    # Funcionario
    funcionario_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[55]/div/div/div/div/input')
    funcionario_button.click()
    funcionario_button.send_keys(row[8])
    time.sleep(0.5)
    
    # Proximo
    next_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[57]/div/div/div/div')
    next_button.click()
    time.sleep(1)
    
    # Ativar
    activate_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[24]/div/div/div/div/div[2]/div[2]')
    activate_button.click()
    time.sleep(1)
    
    # Segmento
    segmento_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[26]/div/div/div/div[1]')
    segmento_element.click()                            
    time.sleep(1)
    segmento_input = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div/div/div/div/div/input')
    segmento_string = str(row[11]).lstrip(", ")
    segmento_input.send_keys(segmento_string)                      
    time.sleep(1)
    
    #Class
    class_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[28]/div/div/div/div/div/div[1]')
    class_element.click()
    time.sleep(1)
    
    #Volume
    volume_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[32]/div/div/div/div/input')
    volume_element.click()
    volume_element.send_keys(row[13])
    time.sleep(1)
    
    #Proximo 
    next_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/div/div[42]')
    next_button.click()
    time.sleep(1)
    
    # Ativar
    activate_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[24]')
    activate_button.click()
    time.sleep(1)
    
    # Produto
    produto_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[26]/div/div/div/div/div/div[1]')
    produto_element.click()
    time.sleep(1)
    
    # # Tier
    # if row[12] == 'A':
    #     clicks = 0
    # elif row[12] == 'AAA':
    #     clicks = 1
    # elif row[12] == 'B':
    #     clicks = 2
    # elif row[12] == 'C':
    #     clicks = 3
    # elif row[12] == 'D':
    #     clicks = 4
    # elif row[12] == 'E':
    #     clicks = 5
    # tier_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[28]/div/div/div/div/div/div[1]')
    # tier_element.click() 
    # for _ in range(clicks):
    #     tier_element.send_keys(Keys.ARROW_DOWN)
    #     time.sleep(0.1)
    # time.sleep(1000000)
    
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
        # Cancel
        cancel_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[5]/div/div/div[6]/div/div/div/div/button')
        cancel_element.click()
        time.sleep(5)
         
        # CNPJ
        cnpj_input = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[2]/div/div/div[18]/div/div/div/div/input')
        cnpj_input.click()
        cnpj_input.send_keys(row[0])
        time.sleep(10)
        
        #EDIT
        edit_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[2]/div/div/div[47]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/div/div/div/div')
        edit_button.click()
        time.sleep(5)
        
        #Funcionario
        funcionario_button = navegador.find_element(By.XPATH,'/html/body/div[1]/div/div/div/div[3]/div/div/div[55]/div/div/div/div/input')
        funcionario_button.click()
        funcionario_button.send_keys(Keys.CONTROL + "a")
        funcionario_button.send_keys(row[8])
        time.sleep(5)
        
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
    message['Subject'] = 'LOG BOT PIPELINE'

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