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
time.sleep(20)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")



# Indicadores de negocios
indicador_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(8) ')
indicador_element.click()
time.sleep(5)

# prazo
prazo_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(1) > div > div')
prazo_element.click()
time.sleep(5)

# taxa de entrega realizada D-1
taxa_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(13) > div > div')
taxa_element.click()
time.sleep(10)
   
# (T-1)
resumo_element = navegador.find_element(By.CSS_SELECTOR, '#tab-summary')
resumo_element.click()
time.sleep(5)

# # teste PR
# remetente_element = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > form:nth-child(4) > div:nth-child(1) > div > div > div > input')
# remetente_element.click()                                    
# time.sleep(5)

# remetente_element.send_keys('PR')
# time.sleep(3)

# remetente_element.send_keys(Keys.ARROW_DOWN)
# remetente_element.send_keys(Keys.ENTER)
# time.sleep(60)

# consultar
consultar_element = navegador.find_elements(By.CSS_SELECTOR, ' #jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element[0].click()
time.sleep(5)

# clicar em exportar
exportar_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(2)')
exportar_element[0].click()
time.sleep(10)

# clicar em central de downloads
downloads_element = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > div > div:nth-child(1) > button')
downloads_element[0].click()
time.sleep(10)

# clicar em por data de envio na central de downloads
por_data_downloads_element = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(1) > div > div > div > input')
por_data_downloads_element.click()
time.sleep(5)

# Se lecionar todo o campo com CTRL+A
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
time.sleep(10)

#consultar download
consultar_download = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_download[1].click()
# navegador.execute_script("arguments[0].click();", consultar_download)
time.sleep(120)


#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-dialog__wrapper > div > div.el-dialog__body > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--small > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr > td.el-table_2_column_37.is-center > div > span > button')
baixar_excel.click()                                    
time.sleep(10)                             
            
navegador.quit()
