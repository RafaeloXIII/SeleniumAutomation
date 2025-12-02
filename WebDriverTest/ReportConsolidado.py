from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json

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
time.sleep(30)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Indicadores de negocios
indicador_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(8)')
indicador_element.click()
time.sleep(5)

#OPERAR 
operar_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(3)')
operar_element.click()
time.sleep(2)

#Report consolidado
pacotesretidos_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(7)')
pacotesretidos_element.click()
time.sleep(25)

# clicar em por data 
por_data_elements = navegador.find_elements(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(2) > div > div > div.el-date-editor.time-start.el-input.el-input--small.el-input--prefix.el-input--suffix.el-date-editor--datetime > input')
por_data_element = por_data_elements[0]
por_data_element.click()
time.sleep(5)

# Selecionar todo o campo com CTRL+A
por_data_element.send_keys(Keys.LEFT)
time.sleep(3)

# Confirma hora
por_data_element.send_keys(Keys.ENTER)
time.sleep(5)

#CONSULTA 
consulta_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consulta_element[0].click()
time.sleep(1)

#exportar consulta 
exportar_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(2)')
exportar_element[0].click()
time.sleep(30)

# clicar em central de downloads
downloads_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(6)')
downloads_element.click()
time.sleep(120)

# clicar em por data de envio na central de downloads
por_data_downloads_elements = navegador.find_elements(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(1) > div > div > div > input')
por_data_downloads_element = por_data_downloads_elements[1]
por_data_downloads_element.click()
time.sleep(5)

# Selecionar todo o campo com CTRL+A
por_data_downloads_element.send_keys(Keys.CONTROL, 'a')
time.sleep(3)

# Calcula a data atual menos 5 minutos
data_atual_menos_5 = datetime.now() - timedelta(minutes=5)
data_formatada_menos_5 = data_atual_menos_5.strftime("%Y-%m-%d %H:%M:%S")
time.sleep(5)

# Escreve a data atual menos 5 minutos no campo selecionado
por_data_downloads_element.send_keys(data_formatada_menos_5)
time.sleep(3)

# Confirma hora
por_data_downloads_element.send_keys(Keys.ENTER)
time.sleep(5)

#CONSULTA 2
consulta2_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consulta2_element[1].click()
time.sleep(120)

#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-dialog__wrapper > div > div.el-dialog__body > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--small > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr:nth-child(1) > td.el-table_2_column_68.is-center > div > span > button')
baixar_excel.click()                                    
time.sleep(10)

navegador.quit()