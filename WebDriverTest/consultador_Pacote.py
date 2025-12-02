from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from selenium.webdriver.common.action_chains import ActionChains
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
time.sleep(20)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

#OPERAÇÃO
operacao_element = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.index.not-activity.dark > div.wrap-main > div > div.box-wrap > div.box-left > div.box-flex > div:nth-child(2)')
operacao_element.click()
time.sleep(5)                                  

#CONSULTA DO PACOTE
consulta_pacote = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar.mini > div.el-scrollbar__wrap > div > div.sidebar.siderbar-mini > div.sidebar-body > ul.el-menu-vertical-left.el-menu > li:nth-child(1) > div > div > span')
consulta_pacote.click()
time.sleep(5)

#RASTREAMENTO DO PACOTE
rastreamento_pacote = navegador.find_element(By.CSS_SELECTOR, '#gateway > div.home > div.homeLeft.el-scrollbar > div.el-scrollbar__wrap > div > div.sidebar > div.sidebar-body > ul.el-menu-vertical-right.el-menu > li:nth-child(1)')
rastreamento_pacote.click()
time.sleep(5)

#PACOTE INPUT
input_element = navegador.find_element(By.CSS_SELECTOR,'#__qiankun_subapp_wrapper_for_vue_operating_platform_index__ > div > div > div > div > div.tk_left_1 > div.tk_search > div.input-tag-container.cms-input-tag.trackingExpress > div > input')
input_element.click()
input_element.send_keys('888000382623601')
input_element.send_keys(Keys.ENTER)
time.sleep(5)

#PESQUISAR
rastreamento_pacote = navegador.find_element(By.CSS_SELECTOR, '#__qiankun_subapp_wrapper_for_vue_operating_platform_index__ > div > div > div > div > div.tk_left_1 > div.tk_search-btn > button.el-button.first-btn.el-button--default.el-button--small')
rastreamento_pacote.click()
time.sleep(5)

# DELAY
time.sleep(300)
navegador.quit()