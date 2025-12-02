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

# Gestao permi.
estrutura_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(5)')
estrutura_element.click()
time.sleep(3)

#Geren. usuario
monitoramento_movimentacao = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(1)')
monitoramento_movimentacao.click()              
time.sleep(20)

#Pesquisar
pesquisar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
pesquisar_element.click()              
time.sleep(3)


code_counter = 0
failed_page = []
for num in range(771):
    code_counter +=1
    #Select all
    all_element = navegador.find_element(By.CSS_SELECTOR, '#basicData > div > div.user-list.avue-crud > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--scrollable-y.el-table--small > div.el-table__fixed > div.el-table__fixed-header-wrapper > table > thead > tr > th.el-table_1_column_1.is-center.el-table-column--selection.is-leaf > div > label > span')
    all_element.click()
    time.sleep(3)

    # Conceder Papel
    conceder_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(4)')
    conceder_element.click()
    time.sleep(3)

    try:
    # Input
        input_element = navegador.find_element(By.CSS_SELECTOR,'#pane-1 > div > div.br-left > div.br-search > div > input')
        input_element.click()
        input_element.send_keys('PRQUALI')
        time.sleep(0.5)

    except:
        failed_page = failed_page.append(code_counter)
        # Proximo
        proximo_element = navegador.find_element(By.CSS_SELECTOR, '#jmsPagination > div > span.el-pagination__jump > div > input')
        proximo_element.click()
        proximo_element.send_keys(Keys.CONTROL,'a')
        proximo_element.send_keys(code_counter + 1)
        proximo_element.send_keys(Keys.ENTER)
        time.sleep(5)
        continue
    
    # List 2
    list2_element = navegador.find_element(By.CSS_SELECTOR,'#pane-1 > div > div.br-left')

    # Button 2
    button2_element = list2_element.find_elements(By.CLASS_NAME,'brrli-name')
    button2_element[1].click()
    time.sleep(1)

    # >
    rightarrow_element = navegador.find_element(By.CSS_SELECTOR,'#pane-1 > div > div.br-center > i.brc-icon.el-icon-arrow-right') 
    rightarrow_element.click()
    time.sleep(2)

    # Salvar
    salvar_element = navegador.find_element(By.CSS_SELECTOR,'#basicData > div > div.form-footer > button.el-button.el-button--primary.el-button--small')
    salvar_element.click()
    time.sleep(10)

    # Proximo
    proximo_element = navegador.find_element(By.CSS_SELECTOR, '#jmsPagination > div > span.el-pagination__jump > div > input')
    proximo_element.click()
    proximo_element.send_keys(Keys.CONTROL,'a')
    proximo_element.send_keys(code_counter + 1)
    proximo_element.send_keys(Keys.ENTER)
    time.sleep(5)

    print(code_counter)

# code_counter = 0

print(f"failed{failed_page}")
time.sleep(30)

navegador.quit()