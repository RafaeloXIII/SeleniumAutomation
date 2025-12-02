from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from selenium.common.exceptions import NoSuchElementException
import sys
sys.path.append(r'c:\Users\Rafael Braga\Bots')
from WebDriverTest.chrome_manager.chrome_manager import WebDriverManager
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

#Financeiro
financeiro_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(6)')
financeiro_element.click()
time.sleep(15)                                  

#Faturamento
faturamento_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(2)')
faturamento_element.click()
time.sleep(15)

#Gestão de abatimento
gestao_faturamento = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(4)')
gestao_faturamento.click()
time.sleep(15)

# Detalhe
detalhe_element = navegador.find_element(By.CSS_SELECTOR,'#tab-first')
detalhe_element.click()
time.sleep(5)

# Consulta de detalhes da taxa de entrega
gestao_conciliacao = navegador.find_elements(By.CSS_SELECTOR,'#pane-first > div > div.zd-tab-submenu > div > div > label.el-radio-button.el-radio-button')
gestao_conciliacao[0].click()                                                          
time.sleep(15)


# consultar
consultar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element.click()
time.sleep(20)

# clicar em exportar
exportar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(3)')
exportar_element.click()
time.sleep(20)

# enviar
enviar_element = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary')
enviar_element.click()                                    
time.sleep(15)


# clicar em central de downloads
downloads_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(4)')
downloads_element[0].click()
time.sleep(20)

# clicar em por data de envio na central de downloads
por_data_downloads_element = navegador.find_element(By.CSS_SELECTOR,'body > div.el-dialog__wrapper.finacl-download-dialog > div > div.el-dialog__body > div > div.search > form > div:nth-child(2) > div > div > input:nth-child(2)')
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
consultar_download = navegador.find_elements(By.CSS_SELECTOR, 'body > div.el-dialog__wrapper.finacl-download-dialog > div > div.el-dialog__body > div > div.search > div > span > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_download[0].click()
time.sleep(30)

#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-dialog__wrapper.finacl-download-dialog > div > div.el-dialog__body > div > div.avue-crud > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.el-table_2_column_46 > div > button')
baixar_excel.click()
time.sleep(10)
navegador.quit()