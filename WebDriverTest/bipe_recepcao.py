from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json

# Le arquivo config.json

with open(r'C:\Users\Bot\AutomationTests\WebDriverTest\config.json', 'r') as config_file:
    config = json.load(config_file)

# Define your download directory
download_directory = config.get('default_assinados_path')

# Create a WebDriverManager instance
driver_manager = WebDriverManager(download_directory)

# Create a WebDriver instance for your automation
navegador = driver_manager.create_driver()

# Now you can use the 'navegador' WebDriver instance in your automation script
navegador.get("https://jmsbr.jtjms-br.com/index")
time.sleep(30)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


#OPERAÇÃO
operacao_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(2)')
operacao_element.click()
time.sleep(5)                                  

#CONSULTA DO PACOTE
consulta_pacote = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(1) > div > div > span')
consulta_pacote.click()
time.sleep(5)

#CONSULTA DAS BIPAGENS
consulta_bipagens = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(4) > div > div > span')
consulta_bipagens.click()
time.sleep(5)

#Consultar estação de nível inferior
flag_nivelinferior = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(17) > div > label > span.el-checkbox__input > span')
flag_nivelinferior.click()
time.sleep(5)

#CLICAR LISTA SUSPENSA TIPO DE BIPAGEM
lista_tipobipagem = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(7) > div > div > div > div > input')
lista_tipobipagem.click()
time.sleep(5)


# assinatura_element = navegador.find_element(By.CSS_SELECTOR, 'body > div:nth-child(33) > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap > ul > li.el-select-dropdown__item.selected > span')

#Bipe recepcao
num_down_arrows_assinatura = 3
for _ in range(num_down_arrows_assinatura):
  lista_tipobipagem.send_keys(Keys.ARROW_DOWN)
  time.sleep(0.5)
lista_tipobipagem.send_keys(Keys.ENTER)
time.sleep(5)
                                                           

consultar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_element.click()
time.sleep(5)

# clicar em exportar
exportar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(2)')
exportar_element.click()
time.sleep(20)

item = navegador.find_element(By.CSS_SELECTOR, 'body > div.el-message-box__wrapper > div > div.el-message-box__btns > button.el-button.el-button--default.el-button--small.el-button--primary.comfirm-btn')

item.click()
time.sleep(7)

# # #teste
# # ActionChains(navegador).move_to_element(item).perform()
 

# clicar em central de downloads
downloads_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(4)')
downloads_element.click()
time.sleep(5)

# clicar em por data de envio na central de downloads
por_data_downloads_element = navegador.find_element(By.CSS_SELECTOR,'#jmsSearch > form > div:nth-child(1) > div > div > div > input')
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
time.sleep(10)

#consultar download
consultar_download = navegador.find_elements(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
consultar_download[1].click()
# navegador.execute_script("arguments[0].click();", consultar_download)
time.sleep(40)


#baixar excel
baixar_excel = navegador.find_element(By.CSS_SELECTOR, '#__qiankun_subapp_wrapper_for_vue_operating_platform_index__ > div > div > div > div.download-list > div > div > div.el-dialog__body > div > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--enable-row-transition.el-table--small > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr:nth-child(1) > td.el-table_2_column_53 > div > button:nth-child(1) > span > span')
baixar_excel.click()
time.sleep(10)


# print("sucesso")
navegador.quit()

