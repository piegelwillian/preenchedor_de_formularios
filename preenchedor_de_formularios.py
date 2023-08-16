import pyautogui as py
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook

#Escape
py.PAUSE = 0.5
FAILSAFE = True

#Baixar última versão do driver do navegador
service = Service(ChromeDriverManager().install())

caminho_arquivo = '/home/will/Documentos/GitHub/preenchedor_de_formularios/candidatos_vaga_emprego.xlsx'
planilha = load_workbook(filename = caminho_arquivo)

#Seleciona a aba da planilha que contém as informações
sheet_selecionada = planilha['Dados']

#contador
cont = 0

while True:
    
    #Percorre todas as linhas da planilha
    for linha in range(2, len(sheet_selecionada['A']) + 1):
        
        nome = sheet_selecionada['A%s' % linha].value
        email = sheet_selecionada['B%s' % linha].value
        telefone = sheet_selecionada['C%s' % linha].value
        sexo = sheet_selecionada['D%s' % linha].value
        sobre = sheet_selecionada['E%s' % linha].value

        #Entrar no navegador
        driver = webdriver.Chrome(service=service)
        driver.get('https://pt.surveymonkey.com/r/Q7CDDRT')
        
        #Tempo de espera
        wait = WebDriverWait(driver, 20)

        wait.until(EC.visibility_of_element_located(('name', '146482313'))).send_keys(nome)
        wait.until(EC.visibility_of_element_located(('name', '146482337'))).send_keys(email)
        wait.until(EC.visibility_of_element_located(('name', '146482357'))).send_keys(telefone)

        if sexo == 'Masculino':
            wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="question-field-146482383"]/fieldset/div/div/div[1]/div/label'))).click()
        else:
            wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="question-field-146482383"]/fieldset/div/div/div[2]/div/label/span[2]'))).click()

        wait.until(EC.visibility_of_element_located(('name', '146482419'))).send_keys(sobre)

        enviar = wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="patas"]/main/article/section/form/div[2]/button'))).click()

        driver.quit()
        
        #Atualizar o contador. Somar mais um a cada vez que passa pelo FOR.
        cont += 1
    
    break

print('O loop passou por', cont, 'iterações.')

py.alert('O BOT foi executado com exito!')
