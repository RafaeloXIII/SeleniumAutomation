from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json

# Le arquivo config.json

with open('WebDriverRS\config.json', 'r') as config_file:
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

#GESTÃO DE BASES
gestao_bases = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(5)')
gestao_bases.click()
time.sleep(15)                                  

#GESTÃO DE PEDIDOS JMS
gestao_pedidos = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(2) > div')
gestao_pedidos.click()
time.sleep(15)

#GESTÃO DE PEDIDOS JMS DE REMESSA
pedidos_remessa = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(1)')
pedidos_remessa.click()
time.sleep(15)

# Por hora de envio
por_hora_envio_element = navegador.find_element(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(2) > div > div > div.el-radio-group > label:nth-child(2) > span.el-radio__input')
por_hora_envio_element.click()
time.sleep(5)

# clicar em por data 
por_data_downloads_element = navegador.find_element(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(2) > div > div > div.el-date-editor.time-item.time-start.el-input.el-input--small.el-input--prefix.el-input--suffix.el-date-editor--datetime > input')
por_data_downloads_element.click()
time.sleep(15)

por_data_downloads_element.send_keys(Keys.CONTROL, 'a')
time.sleep(3)

# Calcula a data atual menos 1 dia
data_atual_menos_1 = datetime.now() - timedelta(days=1)
data_formatada_menos_1 = data_atual_menos_1.strftime("%Y-%m-%d 00:00:00")
time.sleep(5)

# Escrever a data atual menos 5 minutos no campo selecionado
por_data_downloads_element.send_keys(data_formatada_menos_1)
time.sleep(3)

# consultar
consultar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element.click()
time.sleep(30)

# clicar em exportar
exportar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(2)')
exportar_element.click()
time.sleep(20)


# clicar em central de downloads
downloads_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(8)')
downloads_element.click()
time.sleep(30)

# clicar em por data de envio na central de downloads
por_data_downloads_element = navegador.find_element(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(2) > div > div > input:nth-child(2)')
por_data_downloads_element.click()
time.sleep(15)

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
consultar_download[1].click()
time.sleep(60)

#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, '#__qiankun_subapp_wrapper_for_vue_network_management_index__ > div > div > div.el-dialog__wrapper.network-download-dialog > div > div.el-dialog__body > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.el-table_2_column_70 > div > button')
baixar_excel.click()
time.sleep(10)
navegador.quit()