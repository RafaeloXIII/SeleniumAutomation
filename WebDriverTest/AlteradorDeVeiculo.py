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

#TRANSPORTE
transporte_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(7)')
transporte_element.click()
time.sleep(5)

#INFORMAÇÃO BASICA
info_basica = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(1) > div > div')
info_basica.click()
time.sleep(5)

#GESTÃO DE MOTORISTAS E VEICULOS 车辆司机管理
motoristas_veiculos = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(3) > div > div > div') 
motoristas_veiculos.click()                                     
time.sleep(5)

#LASTMILE CADASTRO DE MOTORISTA E VEICULOS
lastmile_cadastro = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li.el-submenu.is-opened > ul > li:nth-child(4) > div > div')
lastmile_cadastro.click()
time.sleep(5)

# # xlsx_file = r"Z:\Recursos\Teste\example.xlsx"  # Provide your xlsx file path here
xlsx_file = r"C:\Users\Rafael Braga\Downloads\motoristas.xlsx"

workbook = openpyxl.load_workbook(xlsx_file)
sheet = workbook.active

codes = []
for row in sheet.iter_rows(min_row=1, values_only=True):  # Assuming data starts from the second row
    codes.append(row[0])  # Assuming the codes are in the first column

code_counter = 224
for code in codes[code_counter:]:
    #Placa Input
    numero_input = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(2) > div > div > div > input')
    numero_input.click()

    time.sleep(3)
    pyperclip.copy(code)
    numero_input.click()
    numero_input.send_keys(Keys.CONTROL, 'a')
    numero_input.send_keys(Keys.CONTROL, 'v')

    numero_input.send_keys(Keys.ENTER)
    code_counter+=1
    time.sleep(5)

    # #Pesquisar
    # pequisar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
    # pequisar_element.click()              
    # time.sleep(5)
    try:
        #Editar
        editar_element = navegador.find_element(By.CSS_SELECTOR,'#__qiankun_subapp_wrapper_for_vue_transport_platform_index__ > div > div > div.box-list.avue-crud > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.el-table_1_column_26 > div > button:nth-child(2) > i ')
        editar_element.click()
        time.sleep(2)
    except:
        print(f"Pulando linha {code_counter}, numero {code}")
        time.sleep(1)
        continue

    #TIPO DE VEICULO
    tipo_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[1]/div[1]/div[1]/div[2]/div/div[2]/div/div/div/input')
    tipo_element.click()
    time.sleep(1)

    #Assinatua de Encomenda
    num_down = 16
    for _ in range(num_down):
        tipo_element.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.5)
    tipo_element.send_keys(Keys.ENTER)
    time.sleep(1)

    #TIPO
    tipo_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[1]/div[1]/div[1]/div[2]/div/div[6]/div/div/div/input')
    tipo_element.click()
    time.sleep(1)

    #Assinatua de Encomenda
    num_down = 10
    for _ in range(num_down):
        tipo_element.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.5)
    tipo_element.send_keys(Keys.ENTER)
    time.sleep(1)

    #TIPO DE VEICULO 2
    tipo_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[4]/div/div/div/input')
    tipo_element.click()
    time.sleep(1)

    #Assinatua de Encomenda
    num_down = 6
    for _ in range(num_down):
        tipo_element.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.5)
    tipo_element.send_keys(Keys.ENTER)
    time.sleep(1)

    # Proximo Passo
    proximo_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[1]/div[2]/button[2]')
    proximo_element.click()
    time.sleep(0.5)
    try:
        # Nivel
        nivel_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[31]/div/div/div/input     ')
        nivel_element.click()
        nivel_element.send_keys('Senior áreas')
        time.sleep(0.5)
        nivel_element.send_keys(Keys.ARROW_DOWN)
        nivel_element.send_keys(Keys.ENTER)
        time.sleep(0.5)
    except:
        print(f"Pulando linha {code_counter}, numero {code}")
        #LASTMILE CADASTRO DE MOTORISTA E VEICULOS
        lastmile_cadastro = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li.el-submenu.is-active.is-opened > ul > li.el-menu-item.el-menu-second-item.is-active')
        lastmile_cadastro.click()
        time.sleep(2)
        continue

    # Salvar
    salvar_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[2]/div[2]/button[3]')
    salvar_element.click()
    time.sleep(1)

    # Confirmar
    confirm_element = navegador.find_element(By.CSS_SELECTOR,'body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary.comfirm-btn')
    confirm_element.click()
    time.sleep(5)
    
    print(code_counter)

# code_counter = 0

print("Saindo do loop")
time.sleep(20)

navegador.quit()