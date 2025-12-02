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

#flag> Out of service 
click_flag = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(10) > div > div > div > div > label:nth-child(1) > span.el-checkbox__input > span')
click_flag.click()
time.sleep(5)

# consultar
consultar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element.click()
time.sleep(5)

# clicar em exportar
exportar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(4)')
exportar_element.click()
time.sleep(5)

item = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary.comfirm-btn')

item.click()
time.sleep(5)


# #teste
# ActionChains(navegador).move_to_element(item).perform()



# clicar em por data de envio na central de downloads
por_data_downloads_elements = navegador.find_elements(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(2) > div > div > div > input')
por_data_downloads_element = por_data_downloads_elements[1]
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
consultar_download[1].click()
time.sleep(40)

#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, '#__qiankun_subapp_wrapper_for_vue_transport_platform_index__ > div > div > div.download-list2.download-center2 > div > div > div.el-dialog__body > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr:nth-child(1) > td.el-table_2_column_34 > div > button')
baixar_excel.click()
time.sleep(20)
navegador.quit()




