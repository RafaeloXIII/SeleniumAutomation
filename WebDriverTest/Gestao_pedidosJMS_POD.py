from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from selenium.common.exceptions import NoSuchElementException
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
time.sleep(10)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

#GESTÃO DE BASES
gestao_bases = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(5)')
gestao_bases.click()
time.sleep(15)                                  

#GESTÃO DE PEDIDOS JMS
gestao_pedidos = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(2) > div')
gestao_pedidos.click()
time.sleep(15)

#GESTÃO DE PEDIDOS JMS DE ENTREGA
pedidos_entrega = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(2)')
pedidos_entrega.click()
time.sleep(15)

#HORARIO DE ENTREGA
tempo_entrega = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(2) > div > div > div.el-radio-group > label:nth-child(3) > span.el-radio__input')
tempo_entrega.click()
time.sleep(5)

# consultar
consultar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element.click()
time.sleep(35)


# #Clicar no filtro
# filtro_element = navegador.find_element(By.CSS_SELECTOR,'#__qiankun_subapp_wrapper_for_vue_network_management_index__ > div > div > div.avue-crud > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--small > div.el-table__fixed-right > div.el-table__fixed-header-wrapper > table > thead > tr > th.el-table_1_column_34.is-center.is-leaf > div > button')
# filtro_element.click()                                   
# time.sleep(5)

# #Restaura padrao
# padrao_element = navegador.find_element(By.CSS_SELECTOR,'#__qiankun_subapp_wrapper_for_vue_network_management_index__ > div > div > div.avue-crud > div.drag-column > div > div.dc-bottom > div.dct-top > div.dctt-right > div.dctt-default')
# padrao_element.click()
# time.sleep(5)

# #Aplicar tudo
# tudo_element = navegador.find_element(By.CSS_SELECTOR,' #__qiankun_subapp_wrapper_for_vue_network_management_index__ > div > div > div.avue-crud > div.drag-column > div > div.dc-bottom > div.dct-top > div.dctt-right > div:nth-child(2)')
# tudo_element.click()
# time.sleep(5)

# #Salvar
# salvar_element = navegador.find_element(By.CSS_SELECTOR,'#__qiankun_subapp_wrapper_for_vue_network_management_index__ > div > div > div.avue-crud > div.drag-column > div > div.dc-top > div.dct-top > div.dctt-btn > button.el-button.dcttb-btn.el-button--primary.el-button--mini')
# salvar_element.click()
# time.sleep(5) 



# clicar em exportar
exportar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(2)')
exportar_element.click()
time.sleep(20)


# clicar em central de downloads
downloads_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(7)')
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
baixar_excel = navegador.find_element(By.CSS_SELECTOR, '#__qiankun_subapp_wrapper_for_vue_network_management_index__ > div > div > div.el-dialog__wrapper.network-download-dialog > div > div.el-dialog__body > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.el-table_2_column_99 > div > button')
baixar_excel.click()                                                     
time.sleep(10)
navegador.quit()