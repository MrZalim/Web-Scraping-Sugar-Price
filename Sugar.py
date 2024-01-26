from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import os, shutil
import os
import pandas as pd
import jpype
import asposecells

#Enova@18

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless=new')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

prefs = {"profile.default_content_settings.popups": 0,
"download.default_directory": r"E:\Fonte_Dados\BI\Automações_Python\Açucar\\", #Automatizar diretório para o diretório do Arquivo
"directory_upgrade": True}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome('E:/Fonte_Dados/BI/Automações_Python/chromedriver_win32/chromedriver.exe',chrome_options=chrome_options)

driver.get("https://www.cepea.esalq.usp.br/br/indicador/acucar.aspx")

time.sleep(2)
driver.implicitly_wait(10)
page_title = driver.title
content = driver.page_source
soup = BeautifulSoup(content,"html.parser")
results = soup.find(id="imagenet-wrap-content")
print(results.prettify())
print("Site Armazenado")
time.sleep(2)

dir = 'E:\Fonte_Dados\BI\Automações_Python\Açucar\\'
for files in os.listdir(dir):
    path = os.path.join(dir, files)
    try:
        shutil.rmtree(path)
    except OSError:
        os.remove(path)
print("Arquivos Antigos Removidos")

botao = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='imagenet-content']/div[2]/div[2]/div[1]/div[3]/a[4]")))
botao.click()

WebDriverWait(driver=driver, timeout=10).until(
    lambda x: x.execute_script("return document.readyState === 'complete'")
)

time.sleep(10)
print("Novo Arquivo Baixado")

lista_arquivos = []
for file in os.listdir(dir):
    if file.endswith('.xls'):
        print('Loading file {0}...'.format(file))
if jpype.isJVMStarted():
    from asposecells.api import Workbook
    workbook = Workbook(dir+'\\'+file)
    workbook.save(dir+'\Açucar.xlsx')
else:
    jpype.startJVM()

    from asposecells.api import Workbook
    workbook = Workbook(dir+'\\'+file)
    workbook.save(dir+'\Açucar.xlsx')
time.sleep(2)
print("Arquivo Modificado Criado na Pasta")

os.remove(dir+'\\'+file)
time.sleep(2)
print("Arquivo Original Excluído")

time.sleep(3)
print("Script Fechado")
driver.close()