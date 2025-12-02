from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json
import csv
import openpyxl
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
navegador.get("https://jmsbr.jtjms-br.com/index")
time.sleep(5)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# InfoBasicas
indicador_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(1)')
indicador_element.click()
time.sleep(3)

# Estrutura org.
estrutura_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(1)')
estrutura_element.click()
time.sleep(3)

#Info do Colaborador
monitoramento_movimentacao = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(7)')
monitoramento_movimentacao.click()              
time.sleep(3)

# xlsx_file = r"Z:\Recursos\Teste\example.xlsx"  # Provide your xlsx file path here
xlsx_file = r"C:\Users\Rafael Braga\Downloads\thais.xlsx"

workbook = openpyxl.load_workbook(xlsx_file)
sheet = workbook.active

logins = []
names = []
cpf = []
phones = []
emails = []
code = []
dep = []
role = []

# Extract info
for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming data starts from the second row
    logins.append(row[0])  # Assuming the codes are in the first column
    names.append(row[1])  
    cpf.append(row[6])
    code.append(row[10])
    dep.append(row[11])
    role.append(row[12])  
    phones.append(row[13])
    emails.append(row[16])
collumn_counter = 0

# Interation
for login in logins:
    #Adicionar
    adicionar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(1)')
    adicionar_element.click()              
    time.sleep(3)
    # Input element
    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[1]/div[2]/div/div[1]/div/div[1]/input")
    # print(len(input_elements))
    pyperclip.copy(login)
    input_elements.click()
    input_elements.send_keys(Keys.CONTROL, 'v')
    time.sleep(0.5)
    
    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[1]/div[2]/div/div[2]/div/div[1]/input")
    pyperclip.copy(names[collumn_counter])
    input_elements.click()
    input_elements.send_keys(Keys.CONTROL, 'v')
    time.sleep(0.5)

    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[1]/div[2]/div/div[7]/div/div[1]/input")
    pyperclip.copy(cpf[collumn_counter])
    input_elements.click()
    input_elements.send_keys(Keys.CONTROL, 'v')
    input_elements.send_keys(Keys.BACKSPACE)
    if collumn_counter > 9: input_elements.send_keys(Keys.BACKSPACE)
    input_elements.send_keys(collumn_counter)
    time.sleep(0.5)

    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[1]/div[2]/div/div[8]/div/div/div[1]/input")
    input_elements.click()
    input_elements.send_keys(Keys.ARROW_DOWN)
    input_elements.send_keys(Keys.ENTER)
    time.sleep(0.5)

    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[2]/div[2]/div/div[1]/div/div/div[1]/input")
    # print(len(input_elements))
    pyperclip.copy(code[collumn_counter])
    input_elements.click()
    input_elements.send_keys(Keys.CONTROL, 'v')
    time.sleep(2)
    
    autofilll_element = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[2]/div[2]/div/div[1]/div/div/div[2]/ul/li[1]/div")
    autofilll_element.click()

    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[2]/div[2]/div/div[2]/div/div/div[1]/input")
    pyperclip.copy(dep[collumn_counter])
    input_elements.click()
    time.sleep(0.2)
    input_elements.send_keys(Keys.CONTROL, 'v')
    time.sleep(0.2)
    input_elements.send_keys(Keys.ARROW_DOWN)
    input_elements.send_keys(Keys.ENTER)
    time.sleep(0.5)

    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[2]/div[2]/div/div[3]/div/div/div[1]/input")
    pyperclip.copy(dep[collumn_counter])
    input_elements.click()
    time.sleep(0.2)
    input_elements.send_keys(Keys.CONTROL, 'v')
    time.sleep(0.2)
    input_elements.send_keys(Keys.ARROW_DOWN)
    input_elements.send_keys(Keys.ENTER)
    time.sleep(0.5)

    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[3]/div[2]/div/div[1]/div/div[1]/input")
    pyperclip.copy(phones[collumn_counter])
    input_elements.click()
    input_elements.send_keys(Keys.CONTROL, 'v')
    time.sleep(0.5)


    input_elements = navegador.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/div/form/div[1]/div[3]/div[2]/div/div[4]/div/div[1]/input")
    pyperclip.copy(emails[collumn_counter])
    input_elements.click()
    input_elements.send_keys(Keys.CONTROL, 'v')
    time.sleep(0.5)

    # Salvar
    cancel_element = navegador.find_element(By.CSS_SELECTOR, "#basicData > div > div > form > div.form-footer > div > button.el-button.el-button--primary.el-button--small")
    cancel_element.click()
    time.sleep(3)

    # # Confirmar
    # confirm_element = navegador.find_element(By.CSS_SELECTOR, "body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary.comfirm-btn")
    # confirm_element.click()
    # time.sleep(2)

    collumn_counter +=1
    

# print("Saindo do loop")
# time.sleep(60)

navegador.quit()