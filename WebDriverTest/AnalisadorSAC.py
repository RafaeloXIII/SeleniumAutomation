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

data_atual = datetime.now()
um_dia = timedelta(days=1)

data_menos_1 = data_atual - um_dia

data_menos_1_formatada = data_menos_1.strftime("%Y-%m-%d")

#QUALIDADE DE SERVIÇOS
qualidade_servicos = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(9)')
qualidade_servicos.click()
time.sleep(5) 

#ATENDIMENTO
atendimento_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(3)')
atendimento_element.click()
time.sleep(5)

#ORDEM DE SERVICO
servico_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(2) > div > div > div')
servico_element.click()
time.sleep(5)

#TICKET COMUM
ticket_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li.el-submenu.is-opened > ul > li:nth-child(3) > div > div')
ticket_element.click()
time.sleep(5)

#STATUS
status_element = navegador.find_element(By.CSS_SELECTOR, '#jmsSearch > form > div:nth-child(8) > div > div > div > div.el-input.el-input--small.el-input--suffix > input')
status_element.click()
num_down_arrows_assinatura = 2
for _ in range(num_down_arrows_assinatura):
  status_element.send_keys(Keys.ARROW_DOWN)
  time.sleep(0.5)
status_element.send_keys(Keys.ENTER)
time.sleep(5)

#PESQUISAR
pesquisar_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
pesquisar_element.click()
time.sleep(5)

try:
    #ALL
    ticket_element = navegador.find_element(By.CSS_SELECTOR, '#pane-first > div > div.avue-crud > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--scrollable-x.el-table--scrollable-y.el-table--enable-row-transition.el-table--small > div.el-table__fixed > div.el-table__fixed-header-wrapper > table > thead > tr > th.el-table_1_column_1.is-center.el-table-column--selection.is-leaf > div')
    ticket_element.click()
    time.sleep(5)
except:
    print("NO DATA")
    time.sleep(10)
    navegador.quit()
    quit()

#DISTRIBUIR TICKETS
distribuirT_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button:nth-child(4)')
distribuirT_element.click()
time.sleep(5)

#NOME
nome_element = navegador.find_element(By.CSS_SELECTOR, '#pane-first > div > div.distribution-user-list > div.avue-drawer > div.avue-drawer__wrapper > div > div.avue-drawer__body > form > div:nth-child(1) > div > div > input')
nome_element.click()
time.sleep(0.5)
nome_element.send_keys("douglas almeida")
time.sleep(5)

#PESQUISAR 2
pesquisar2_element = navegador.find_element(By.CSS_SELECTOR, '#pane-first > div > div.distribution-user-list > div.avue-drawer > div.avue-drawer__wrapper > div > div.avue-drawer__body > form > div.el-form-item.drawer-search.el-form-item--small > div > button')
pesquisar2_element.click()
time.sleep(5)

#DISTRIBUIR TICKETS
distribuir_element = navegador.find_element(By.CSS_SELECTOR, '#pane-first > div > div.distribution-user-list > div.avue-drawer > div.avue-drawer__wrapper > div > div.avue-drawer__body > div.user-list.el-scrollbar > div.el-scrollbar__wrap > div > div > button')
distribuir_element.click()
time.sleep(5)

#CONFIRMAR
confirmar_element = navegador.find_element(By.CSS_SELECTOR, 'body > div:nth-child(29) > div > div.el-dialog__footer > div > button.el-button.el-button--primary.el-button--small')
confirmar_element.click()
time.sleep(5)

#PROCESSAMENTO TICKET
processamrnto_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li.el-submenu.is-opened > ul > li:nth-child(4)')
processamrnto_element.click()
time.sleep(5)

#PESQUISAR3
pesquisar3_element = navegador.find_element(By.CSS_SELECTOR, '#jmsTopMenu > div.avue-crud__left > button.el-button.btn-query.el-button--primary.el-button--small')
pesquisar3_element.click()
time.sleep(5)

while(1):
    try:
        #EDITAR
        edit_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/main/div[2]/div[10]/div/div/div/div/div/div[2]/div/div/div[1]/div[3]/div[5]/div[2]/table/tbody/tr[1]/td[38]/div/span')
        edit_element.click()                                       
        time.sleep(5)
    except:
        print("NO MORE ELEMENTS")
        navegador.quit()
        quit()

    #OPINIAO TEXTO
    text_element = navegador.find_element(By.CSS_SELECTOR, '#__qiankun_subapp_wrapper_for_vue_service_quality_index__ > div > div > div > div:nth-child(1) > div > div > div:nth-child(2) > div > div.group-content > div:nth-child(3) > span.input-box > div.el-textarea.el-input--small.el-input--suffix > textarea')
    text_element.click()                                                         
    time.sleep(1)
    text_element.send_keys("Ticket recebido, estamos atuando.")
    time.sleep(5)

    #SUBMETER
    submit_element = navegador.find_element(By.CSS_SELECTOR, '#__qiankun_subapp_wrapper_for_vue_service_quality_index__ > div > div > div > div:nth-child(1) > div > div > div.btn-group > button.el-button.el-button--primary.el-button--small')
    submit_element.click()
    time.sleep(10)
