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

#Numero Input
numero_input = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(1) > div > div > div > input')
numero_input.click()
time.sleep(5)

# xlsx_file = r"Z:\Recursos\Teste\example.xlsx"  # Provide your xlsx file path here
xlsx_file = r"C:\Users\Rafael Braga\Downloads\Senhas Reset(2).xlsx"

workbook = openpyxl.load_workbook(xlsx_file)
sheet = workbook.active

codes = []
for row in sheet.iter_rows(min_row=1, values_only=True):  # Assuming data starts from the second row
    codes.append(row[0])  # Assuming the codes are in the first column

code_counter = 0
numero_input.click()
for code in codes:
    print(type(code))
    if code != None:
        pyperclip.copy(code)
        numero_input.click()
        numero_input.send_keys(Keys.CONTROL, 'v')
        numero_input.send_keys(' ')
        code_counter+=1
        time.sleep(0.5)

        #Pesquisar
        pequisar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
        pequisar_element.click()              
        time.sleep(3)

        try:
            # Editar
            editar_element = navegador.find_element(By.CSS_SELECTOR,'#basicData > div > div.user-list.avue-crud > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.el-table_1_column_15 > div > button:nth-child(2)')
            editar_element.click()
            time.sleep(4)

            # Permitir
            estado_element = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/main/div[2]/div[2]/div/div/div/form/div[1]/div[1]/div[2]/div/div[11]/div/div/label[2]/span[1]')
            estado_element.click()
            time.sleep(3)
                                            
            # Salvar
            salvar_element = navegador.find_element(By.CSS_SELECTOR,'#basicData > div > form > div.form-footer.formauth-footer > button.el-button.el-button--primary.el-button--small')
            salvar_element.click()
            time.sleep(3)

            # Confirmar
            salvar_element = navegador.find_element(By.CSS_SELECTOR,'body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary.comfirm-btn')
            salvar_element.click()
            time.sleep(4)
        except:
            print("Teste")
            numero_input.click()
            time.sleep(3)
            delete_element = navegador.find_element(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(1) > div > div > div > span')
            delete_element.click()
            time.sleep(5)
            break
    code_counter = 0

# print("Saindo do loop")
# time.sleep(60)

navegador.quit()