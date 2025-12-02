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

with open(r'.\WebDriverTest\config.json', 'r') as config_file:
    config = json.load(config_file)

# Define your download directory
download_directory = config.get('default_dowload_directory')

# Create a WebDriverManager instance
driver_manager = WebDriverManager(download_directory)

# Create a WebDriver instance for your automation
navegador = driver_manager.create_driver()

# Now you can use the 'navegador' WebDriver instance in your automation script
navegador.get("https://jmsbr.jtjms-br.com/index")
time.sleep(20)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


#QUALIDADE DE SERVIÇOS
qualidade_servicos = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(9)')
qualidade_servicos.click()
time.sleep(5)              

#GESTÃO DE ANORMALIDADES
gestao_anormalidades = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(5) > div > div')
gestao_anormalidades.click()
time.sleep(5)

#GESTÃO DE PACOTE PROBLEMÁTICO
gestao_pacoteproblematico = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(1) > div > div > div')
gestao_pacoteproblematico.click()
time.sleep(5) 

#MONITORAMENTO DE PACOTE PROBLEMÁTICO
monitoramento_pacoteproblematico = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li.el-submenu.is-opened > ul > li:nth-child(3)')
monitoramento_pacoteproblematico.click()
time.sleep(20)

# CLICAR EM POR DATA
por_data_element = navegador.find_element(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(2) > div > div > div.el-date-editor.time-item.time-start.el-input.el-input--small.el-input--prefix.el-input--suffix.el-date-editor--datetime > input')
por_data_element.click()
time.sleep(5)

# Selecionar todo o campo com CTRL+A
por_data_element.send_keys(Keys.CONTROL, 'a')
time.sleep(3)

# Calcula a data atual menos 1 dia
data_atual_menos_1 = datetime.now() - timedelta(days=1)
data_formatada_menos_1 = data_atual_menos_1.strftime("%Y-%m-%d 00:00:00")
time.sleep(5)

# Escreve a data atual menos 1di no campo selecionado
por_data_element.send_keys(data_formatada_menos_1)
time.sleep(10)

# Confirma hora
por_data_element.send_keys(Keys.ENTER)
time.sleep(5)

# consultar
consultar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element.click()
time.sleep(5)

# clicar em exportar
exportar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(2)')
exportar_element.click()
time.sleep(5)

# clicar em central de downloads
downloads_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(4)')
downloads_element.click()
time.sleep(5)

# clicar em por data de envio na central de downloads
por_data_downloads_element = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-dialog__wrapper.network-download-dialog > div > div.el-dialog__body > div > div.search > form > div:nth-child(2) > div > div:nth-child(1) > input')
por_data_downloads_element.click()
time.sleep(5)

# por_data_downloads_element.send_keys(Keys.CONTROL, 'a')
# time.sleep(3)

# Calcula a data atual menos 5 minutos
data_atual_menos_5 = datetime.now() - timedelta(minutes=5)
hora_formatada_menos_5 = data_atual_menos_5.strftime("%H:%M:%S")
time.sleep(20)

# # Escrever a data atual menos 5 minutos no campo selecionado
# por_data_downloads_element.send_keys(data_formatada_menos_5)
# time.sleep(3)

#clicar na caixa de hora
caixa_hora_elements = navegador.find_elements(By.CSS_SELECTOR, 'body > div.el-picker-panel.el-date-picker.el-popper.has-time > div.el-picker-panel__body-wrapper > div > div.el-date-picker__time-header > span:nth-child(2) > div.el-input.el-input--small > input')
caixa_hora_element = caixa_hora_elements[1]
caixa_hora_element.click()
time.sleep(5)

# Se lecionar todo o campo com CTRL+A
caixa_hora_element.send_keys(Keys.CONTROL, 'a')
time.sleep(3)

# Escrever a data atual menos 5 minutos no campo selecionado
caixa_hora_element.send_keys(hora_formatada_menos_5)
time.sleep(3)

# Confirma o download
por_data_downloads_element.send_keys(Keys.ENTER)
time.sleep(10)

#consultar download
consultar_download = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-dialog__wrapper.network-download-dialog > div > div.el-dialog__body > div > div.search > div > span > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_download.click()
time.sleep(20)

#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-dialog__wrapper.network-download-dialog > div > div.el-dialog__body > div > div.jms-table-wrap > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.el-table_2_column_37 > div > button')
baixar_excel.click()
time.sleep(10)

navegador.quit()