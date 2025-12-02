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
time.sleep(5)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# Indicadores de negocios
indicador_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(8) ')
indicador_element.click()
time.sleep(5)

#OPERAR 
operar_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(3)')
operar_element.click()
time.sleep(2)

#RELATORIO DE FALTA DE BIPAGEM
falta_bipagem = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li.el-submenu')
falta_bipagem.click()
time.sleep(2)

#FALTA DE BIPE LM
falta_bipagemLM = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li.el-submenu.is-opened > ul > li:nth-child(2)')
falta_bipagemLM.click()
time.sleep(5)

#Detalhes
detalhe_element = navegador.find_element(By.CSS_SELECTOR, '#tab-second')
detalhe_element.click()
time.sleep(5)

#DATA E HORA
click_data = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(1) > div > div > div:nth-child(1) > div.el-date-editor.el-input.el-input--small.el-input--prefix.el-input--suffix.el-date-editor--date > input')
click_data.click()
time.sleep(3)

click_data.send_keys(Keys.RIGHT)
time.sleep(3)
click_data.send_keys(Keys.ENTER)
time.sleep(3)

#click_data.send_keys(Keys.CONTROL, 'a')
#click_data.send_keys(data_atual)
#time.sleep(5)

# consultar
consultar_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element[1].click()
time.sleep(35)

# clicar em exportar
exportar_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(2)')
exportar_element[1].click()
time.sleep(20)

# clicar em central de downloads
downloads_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(4)')
downloads_element[1].click()
time.sleep(30)

# clicar em por data de envio na central de downloads
por_data_downloads_element = navegador.find_element(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(1) > div > div > div > input')
por_data_downloads_element.click()
time.sleep(15)

# Calcula a data atual menos 5 minutos
data_atual_menos_5 = datetime.now() - timedelta(minutes=5)
data_formatada_menos_5 = data_atual_menos_5.strftime("%Y-%m-%d %H:%M:%S")
time.sleep(5)

# Escrever a data atual menos 5 minutos no campo selecionado
por_data_downloads_element.send_keys(data_formatada_menos_5)
time.sleep(3)

# Confirma o download
por_data_downloads_element.send_keys(Keys.ENTER)
time.sleep(180)

#consultar download
consultar_download = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_download[2].click()
time.sleep(10)

#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-dialog__wrapper > div > div.el-dialog__body > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--small > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr > td.el-table_3_column_37.is-center > div > span > button')
baixar_excel.click()
time.sleep(10)

navegador.quit()