from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
from chrome_manager.chrome_manager import WebDriverManager
import json
import csv
import openpyxl
import pyperclip

# xlsx_file = r"Z:\Recursos\Teste\example.xlsx"  # Provide your xlsx file path here
xlsx_file = r"C:\Users\Rafael Braga\Downloads\COLABORADORES - 2.1 1.xlsx"

workbook = openpyxl.load_workbook(xlsx_file)
sheet = workbook.active

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
navegador.get("https://connectpr.com.br/orangehrm-5.5/web/index.php/pim/viewEmployeeList")
time.sleep(10)

data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

usuario_element = navegador.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[1]/div/div[2]/input')
usuario_element.click()
usuario_element.send_keys('admin')
time.sleep(0.3)

senha_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div/div[1]/div/div[2]/div[2]/form/div[2]/div/div[2]/input')
senha_element.click()
senha_element.send_keys('J&Tex2022')
time.sleep(0.3)

button_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div/div[1]/div/div[2]/div[2]/form/div[3]/button')
button_element.click()
time.sleep(1)

while True:
    row_number = input("Enter the row number to extract information (or 'exit' to quit): ")
    if row_number.lower() == 'exit':
        break
    
    try:
        row_number = int(row_number) + 1
        for row in sheet.iter_rows(min_row=row_number,max_row=row_number, values_only=True):  # Assuming data starts from the second row
            names = row[1]
            situ = row[3]
            sexo = row[9]
            cargo = row[10]
            dep = row[13]
            boss = row[14]
            adDate = row[15].strftime("%d-%Y-%m") if isinstance(row[15], datetime) else row[15]
            birth = row[16].strftime("%d-%Y-%m") if isinstance(row[16], datetime) else row[16]
            state = row[22]
            adress = row[23]
            number = row[24]
            bairro = row[25]
            city = row[26]
            code = row[27]
        # 26 NACIONALIDADE
        nac_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[1]/div[1]/div/div[2]/div/div/div[1]')
        nac_element.click()
        num_down_arrows = 26
        for _ in range(num_down_arrows):
            nac_element.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.2)
        nac_element.send_keys(Keys.ENTER)
        time.sleep(2)
        # SINGLE 1 MARRIES 2
        state_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[1]/div[2]/div/div[2]/div/div/div[1]')
        state_element.click()
        if state == 'SOLTEIRO 未婚':
            num_down_arrows_state = 1
            print("solteiro")
        else:
            num_down_arrows_state = 2
            print('casado')
        for _ in range(num_down_arrows_state):
            state_element.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.5)
        state_element.send_keys(Keys.ENTER)
        time.sleep(2)
        # BIRTH DAY
        birth_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[2]/div[1]/div/div[2]/div/div/input')
        birth_element.click()
        birth_element.send_keys(birth)
        time.sleep(1)
        # SEX CHOOSE
        if sexo == 'Feminino 女':
            female_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div/label')
            female_element.click()
        else:
            male_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/label')
            male_element.click()
        # SALVAR
        save_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[4]/button')
        save_element.click()
        time.sleep(2)

        # DETALHES DE CONTATO
        contact_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[1]/div[2]/div[2]')
        contact_element.click()
        time.sleep(1)
        # RUA1 = ADRESS,NUMBER,BAIRRO
        adress_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div[1]/div/div[2]/input')
        adress_element.click()
        adress_element.send_keys(f'{adress} {number} {bairro}')
        time.sleep(0.2)
        # CITY
        city_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div[3]/div/div[2]/input')
        city_element.click()
        city_element.send_keys(Keys.CONTROL,'a')
        city_element.send_keys(city)
        time.sleep(0.2)
        # CEP
        cep_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div[5]/div/div[2]/input')
        cep_element.click()
        cep_element.send_keys(Keys.CONTROL,'a')
        cep_element.send_keys(code)
        time.sleep(0.2)
        # COUNTRY 30
        count_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div[6]/div/div[2]/div/div/div[1]')
        count_element.click()
        num_down_arrows = 30
        for _ in range(num_down_arrows):
            count_element.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.2)
        count_element.send_keys(Keys.ENTER)
        time.sleep(2)
        # SALVAR
        save_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[4]/button')
        save_element.click()
        time.sleep(2)
            

        # TRABALHO
        work_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[1]/div[2]/div[6]')
        work_element.click()
        time.sleep(1)
        # AD DATE
        adDate_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div[1]/div/div[2]/div/div/input')
        adDate_element.click()
        adDate_element.send_keys(adDate)
        time.sleep(0.2)
        # SUB = 2
        sub_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div[5]/div/div[2]/div/div/div[1]')
        sub_element.click()
        num_down_arrows = 2
        for _ in range(num_down_arrows):
            sub_element.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.2)
        sub_element.send_keys(Keys.ENTER)
        time.sleep(2)
        # LOC = 1
        sub_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div[6]/div/div[2]/div/div/div[1]')
        sub_element.click()
        num_down_arrows = 1
        for _ in range(num_down_arrows):
            sub_element.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.2)
        sub_element.send_keys(Keys.ENTER)
        time.sleep(2)
        # SITU = 2
        situ_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div[7]/div/div[2]/div/div/div[1]')
        situ_element.click()
        num_down_arrows = 2
        for _ in range(num_down_arrows):
            situ_element.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.5)
        situ_element.send_keys(Keys.ENTER)
        time.sleep(2)
        # SALVAR
        save_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/button')
        save_element.click()
        time.sleep(2)
            
        # REPORT PARA
        report_element = navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[1]/div[2]/div[8]')
        report_element.click()
        time.sleep(1)
    except ValueError:
        print("Please enter a valid row number or 'exit' to quit.")


navegador.quit()