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
time.sleep(10)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# InfoBasicas
indicador_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(1)')
indicador_element.click()
time.sleep(3)

# Estrutura org.
estrutura_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(5)')
estrutura_element.click()
time.sleep(3)

#Info do Colaborador
monitoramento_movimentacao = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(1)')
monitoramento_movimentacao.click()              
time.sleep(3)



# # xlsx_file = r"Z:\Recursos\Teste\example.xlsx"  # Provide your xlsx file path here
xlsx_file = r"C:\Users\Rafael Braga\Downloads\paPEIS.xlsx"

workbook = openpyxl.load_workbook(xlsx_file)
sheet = workbook.active

codes = []
for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming data starts from the second row
    codes.append(row[0])  # Assuming the codes are in the first column

code_counter = 35
for code in codes[code_counter:]:
    #Numero Input
    numero_input = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(1) > div > div > div > input')
    numero_input.click()

    time.sleep(3)
    pyperclip.copy(code)
    numero_input.click()
    numero_input.send_keys(Keys.CONTROL, 'v')
    # numero_input.send_keys('botRS')
    numero_input.send_keys(Keys.ENTER)
    code_counter+=1
    time.sleep(8)

    # #Pesquisar
    # pequisar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
    # pequisar_element.click()              
    # time.sleep(5)
    try:
        #Editar
        editar_element = navegador.find_element(By.CSS_SELECTOR,'#basicData > div > div.user-list.avue-crud > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.el-table_1_column_15 > div > button:nth-child(2)    ')
        editar_element.click()
        time.sleep(2)
    except:
        # Remover
        remover_element = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(1) > div > div > div > div > span > a')
        remover_element.click()
        time.sleep(3)
        print(f"Pulando linha {code_counter}, numero {code}")
        time.sleep(1)
        continue
    

    # List
    list_element = navegador.find_element(By.CSS_SELECTOR,'#pane-first > div > div.br-right > div.brl-list')
    
    # Buttons
    buttons_element = list_element.find_elements(By.CLASS_NAME,'brrli-name')
    for button in buttons_element[1:]:
        print(button.text)
        button.click()
        time.sleep(1)
    
    # <
    leftarrow_element = navegador.find_element(By.CSS_SELECTOR,'#pane-first > div > div.br-center > i.brc-icon.el-icon-arrow-left') 
    leftarrow_element.click()
    time.sleep(0.5)

  # Input
    input_element = navegador.find_element(By.CSS_SELECTOR,'#pane-first > div > div.br-left > div.br-search > div > input')
    input_element.click()
    input_element.send_keys('OperacionalSCPR')
    time.sleep(1)
    
    # List 2
    list2_element = navegador.find_element(By.CSS_SELECTOR,'#pane-first > div > div.br-left > div.brl-list')

    # Button 2
    button2_element = list2_element.find_element(By.CLASS_NAME,'brrli-name')
    button2_element.click()
    time.sleep(1.5)

    # >
    rightarrow_element = navegador.find_element(By.CSS_SELECTOR,'#pane-first > div > div.br-center > i.brc-icon.el-icon-arrow-right') 
    rightarrow_element.click()
    time.sleep(1)

    # Salvar
    salvar_element = navegador.find_element(By.CSS_SELECTOR,'#basicData > div > form > div.form-footer.formauth-footer > button.el-button.el-button--primary.el-button--small')
    salvar_element.click()
    time.sleep(0.5)

    # Confirmar
    salvar_element = navegador.find_element(By.CSS_SELECTOR,'body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary.comfirm-btn')
    salvar_element.click()
    time.sleep(5)

    # Remover
    remover_element = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(1) > div > div > div > div > span > a')
    remover_element.click()
    time.sleep(3)

    print(code_counter)

# code_counter = 0

print("Saindo do loop")
time.sleep(30)

navegador.quit()