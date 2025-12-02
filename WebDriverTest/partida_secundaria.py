from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from selenium.webdriver.common.action_chains import ActionChains
from chrome_manager.chrome_manager import WebDriverManager
import json

# Le arquivo config.json

with open(r'C:\Users\Rafael Braga\Bots\WebDriverTest\config.json', 'r') as config_file:
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
transporte_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(7)')
transporte_element.click()
time.sleep(5)

#Linha secundaria
linha_secundaria = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(5) > div > div')
linha_secundaria.click()
time.sleep(5)

#Precisao de partida linha secundaria
precisao_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(6)') 
precisao_element.click()                                     
time.sleep(5)

#Input regional
input_regional = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(4) > div > div > div.el-input.el-input--small.el-input--suffix > input') 
input_regional.click()                                     
time.sleep(1)
input_regional.send_keys("PR")
time.sleep(5)

#CLICK PR
pr_element = navegador.find_element(By.CSS_SELECTOR,"#jmsSearch > form > div:nth-child(4) > div > div > div.select-content > ul > li:nth-child(1)")
pr_element.click()
time.sleep(5)

# consultar
consultar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element.click()
time.sleep(5)

#Total
total_element = navegador.find_element(By.CSS_SELECTOR,"#pane-first > div > div.avue-crud > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--group.el-table--scrollable-x.el-table--enable-row-transition.el-table--small > div.el-table__body-wrapper.is-scrolling-left > table > tbody > tr > td.el-table_1_column_6.cell-blue > div")
total_element.click()
time.sleep(5)

# clicar em exportar
exportar_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(3)')
exportar_element[1].click()
time.sleep(5)

item = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary.comfirm-btn')
item.click()
time.sleep(5)

# clicar em por data de envio na central de downloads
por_data_downloads_elements = navegador.find_elements(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(2) > div > div > div > input')
por_data_downloads_element = por_data_downloads_elements[0]
por_data_downloads_element.click()                                   
time.sleep(5)

por_data_downloads_element.send_keys(Keys.CONTROL, 'a')
time.sleep(3)

# Calcula a data atual menos 5 minutos
data_atual_menos_5 = datetime.now() - timedelta(minutes=5)
data_formatada_menos_5 = data_atual_menos_5.strftime("%Y-%m-%d %H:%M:%S")
time.sleep(5)

# Escrever a data atual menos 5 minutos no campo selecionado
por_data_downloads_element.send_keys(data_formatada_menos_5)
time.sleep(3)

# Confirma o download
por_data_downloads_element.send_keys(Keys.ENTER)
time.sleep(5)

#consultar download
consultar_download = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_download[2].click()
time.sleep(40)

#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, '#pane-second > div > div.download-list2.download-center2 > div > div > div.el-dialog__body > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr:nth-child(1) > td.el-table_3_column_111 > div > button')
baixar_excel.click()
time.sleep(20)
navegador.quit()




