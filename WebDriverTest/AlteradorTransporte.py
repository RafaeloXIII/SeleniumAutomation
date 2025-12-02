from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
from selenium.webdriver.support.ui import Select
import json
import csv
import openpyxl
import pyperclip
from unidecode import unidecode

def select_option(navegador, xpath, row, target = 'li'):
    dropdown_element = navegador.find_element(By.XPATH, xpath)
    options = dropdown_element.find_elements(By.TAG_NAME, target)

    text = unidecode(str(row).lower())

    for option in options:
        if  unidecode(option.text.lower()) == text:
            option.click()
            break

    time.sleep(2)


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

#TRANSPORTE
try:
    transporte_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(7)')
except:
    time.sleep(20)
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

xlsx_file = r"C:\Users\Rafael Braga\Downloads\Cadastro JMS.xlsx"

workbook = openpyxl.load_workbook(xlsx_file)
sheet = workbook.active

code_counter = 0
for row in sheet.iter_rows(min_row=3, values_only=True):  # Assuming data starts from the second row

    #Adicionar
    add_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(1)')
    add_element.click()


    #Tipo Veiculo
    tipo_veiculo = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[1]/div[2]/div/div[2]/div/div/div/input')
    tipo_veiculo.click()

    #Dropdown tipo
    select_option(navegador,'/html/body/div[6]',row[0])


    #Placa Input
    numero_input = navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[1]/div[2]/div/div[3]/div/div[1]/input')
    numero_input.click()
    numero_input.send_keys(row[1])
    time.sleep(1.5)


    #TIPO
    tipo_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[1]/div[2]/div/div[6]/div/div/div/input')
    tipo_element.click()
    time.sleep(1.5)
    #Dropdown tipo
    select_option(navegador,'/html/body/div[6]',row[2])


    #Tipo VINCULO
    tipo_vinculo = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[1]/div[2]/div/div[7]/div/div/div/input')
    tipo_vinculo.click()
    

    #Dropdown tipo
    select_option(navegador,'/html/body/div[7]',row[3])

    #Tipo Veiculo2
    tipo_veiculo2 = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[4]/div/div/div[1]/input')
    tipo_veiculo2.click()
    time.sleep(1.5)

    #Dropdown tipo
    select_option(navegador,'/html/body/div[8]',row[4])
    time.sleep(200000)

    #Padrao de Emissao
    padrao_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[6]/div/div/div[1]/input')
    padrao_element.click()
    time.sleep(1.5)
    padrao_element.send_keys(Keys.ARROW_DOWN)
    padrao_element.send_keys(Keys.ENTER)

    

    #Tara
    tara_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[13]/div/div[1]/input')
    tara_element.click()
    time.sleep(1.5)
    tara_element.send_keys(row[6])

    #Chassi
    chassi_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[18]/div/div[1]/input')
    chassi_element.click()
    time.sleep(1.5)
    chassi_element.send_keys(row[7])

    #Marca
    marca_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[21]/div/div[1]/input')
    marca_element.click()
    time.sleep(1.5)
    marca_element.send_keys(row[8])
    
    #Ano fabricacao
    ano_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[28]/div/div/input')
    ano_element.click()
    time.sleep(1.5)
    ano_element.send_keys(str(row[9]))
    ano_element.send_keys(Keys.ENTER)
    time.sleep(1.5)

    #EMPLACAMENTO
    emplacamento_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[31]/div/div/div[1]/div/input')
    emplacamento_element.click()
    time.sleep(3)
    words = str(row[10]).split("/")
    pais, estado, cidade = words
    select_option(navegador,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[31]/div/div/div[2]/div[2]/div/div[1]',pais,target='div')
    time.sleep(1.5)
    select_option(navegador,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[31]/div/div/div[2]/div[2]/div/div[1]',estado,target='div')
    time.sleep(1.5)
    select_option(navegador,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[31]/div/div/div[2]/div[2]/div/div[1]',cidade,target='div')
    time.sleep(1.5)

    #Tipo proprietario
    proprietario_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[3]/div[2]/div/div[2]/div/div/div[1]/input')
    proprietario_element.click()

    tipo_proprietario = str(row[11]).lower()
    if tipo_proprietario == "pessoa física":
        tipo_proprietario = "pessoal física"

    select_option(navegador,'/html/body/div[11]/div[1]/div[1]/ul',tipo_proprietario,target='li')
    # print(str(row[11]).lower())
    time.sleep(1.5)

    #chines
    chines_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[1]/div[3]/div[2]/div/div[3]/div/div[1]/input')
    chines_element.click()
    chines_element.send_keys(row[12])
    time.sleep(1.5)

    #Proximo
    proximo_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[1]/div[2]/button[2]')
    proximo_element.click()
    time.sleep(5)

    #CPF
    cpf_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[1]/div/div[1]/input')
    cpf_element.click()
    cpf_element.send_keys(row[13])
    time.sleep(1.5)

    #Nome
    nome_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[2]/div/div[1]/input')
    nome_element.click()
    nome_element.send_keys(str(row[14]))
    time.sleep(1.5)

    #PIS
    pis_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[4]/div/div[1]/input')
    pis_element.click()
    pis_element.send_keys(row[15])
    time.sleep(1.5)
    
    #Id
    id_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[5]/div/div[1]/input')
    id_element.click()
    id_element.send_keys(row[16])
    time.sleep(1.5)

    #DDD
    ddd_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[15]/div/div/input')
    ddd_element.click()
    ddd_element.send_keys(row[17])
    time.sleep(1.5)

    #TEL
    tel_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[16]/div/div/input')
    tel_element.click()
    tel_element.send_keys(row[18])
    time.sleep(1.5)
   
    #Data de nascimento
    birth_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[17]/div/div[1]/input')
    birth_element.click()
    time.sleep(1.5)
    birth_element.send_keys(str(row[19]))
    birth_element.send_keys(Keys.ENTER)
    time.sleep(1.5)

    #Acesso App
    access_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[28]/div/div[1]/input')
    access_element.click()
    time.sleep(1.5)
    access_element.send_keys(row[20])
    time.sleep(1.5)

    #Nivel
    level_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[31]/div/div/div/input')
    level_element.click()
    time.sleep(1.5)
    level_element.send_keys(row[21])
    time.sleep(1.5)
    level_element.send_keys(Keys.ARROW_DOWN)
    level_element.send_keys(Keys.ENTER)
    time.sleep(1.5)


    time.sleep(2000000)

#     numero_input.send_keys(Keys.ENTER)
#     code_counter+=1
#     time.sleep(5)

#     # #Pesquisar
#     # pequisar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
#     # pequisar_element.click()              
#     # time.sleep(5)
#     try:
#         #Editar
#         editar_element = navegador.find_element(By.CSS_SELECTOR,'#__qiankun_subapp_wrapper_for_vue_transport_platform_index__ > div > div > div.box-list.avue-crud > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.el-table_1_column_26 > div > button:nth-child(2) > i ')
#         editar_element.click()
#         time.sleep(2)
#     except:
#         print(f"Pulando linha {code_counter}, numero {code}")
#         time.sleep(1)
#         continue

#     #TIPO DE VEICULO
#     tipo_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[1]/div[1]/div[1]/div[2]/div/div[2]/div/div/div/input')
#     tipo_element.click()
#     time.sleep(1)

#     #Assinatua de Encomenda
#     num_down = 16
#     for _ in range(num_down):
#         tipo_element.send_keys(Keys.ARROW_DOWN)
#         time.sleep(0.5)
#     tipo_element.send_keys(Keys.ENTER)
#     time.sleep(1)

#     #TIPO
#     tipo_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[1]/div[1]/div[1]/div[2]/div/div[6]/div/div/div/input')
#     tipo_element.click()
#     time.sleep(1)

#     #Assinatua de Encomenda
#     num_down = 10
#     for _ in range(num_down):
#         tipo_element.send_keys(Keys.ARROW_DOWN)
#         time.sleep(0.5)
#     tipo_element.send_keys(Keys.ENTER)
#     time.sleep(1)

#     #TIPO DE VEICULO 2
#     tipo_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[1]/div[1]/div[2]/div[2]/div/div[4]/div/div/div/input')
#     tipo_element.click()
#     time.sleep(1)

#     #Assinatua de Encomenda
#     num_down = 6
#     for _ in range(num_down):
#         tipo_element.send_keys(Keys.ARROW_DOWN)
#         time.sleep(0.5)
#     tipo_element.send_keys(Keys.ENTER)
#     time.sleep(1)

#     # Proximo Passo
#     proximo_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[1]/div[2]/button[2]')
#     proximo_element.click()
#     time.sleep(0.5)
#     try:
#         # Nivel
#         nivel_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[2]/div[1]/div[1]/div[2]/div/div[31]/div/div/div/input     ')
#         nivel_element.click()
#         nivel_element.send_keys('Senior áreas')
#         time.sleep(0.5)
#         nivel_element.send_keys(Keys.ARROW_DOWN)
#         nivel_element.send_keys(Keys.ENTER)
#         time.sleep(0.5)
#     except:
#         print(f"Pulando linha {code_counter}, numero {code}")
#         #LASTMILE CADASTRO DE MOTORISTA E VEICULOS
#         lastmile_cadastro = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li.el-submenu.is-active.is-opened > ul > li.el-menu-item.el-menu-second-item.is-active')
#         lastmile_cadastro.click()
#         time.sleep(2)
#         continue

#     # Salvar
#     salvar_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[12]/div/div/div/form[2]/div[2]/button[3]')
#     salvar_element.click()
#     time.sleep(1)

#     # Confirmar
#     confirm_element = navegador.find_element(By.CSS_SELECTOR,'body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary.comfirm-btn')
#     confirm_element.click()
#     time.sleep(5)
    
#     print(code_counter)

# # code_counter = 0

# print("Saindo do loop")
# time.sleep(20)

navegador.quit()