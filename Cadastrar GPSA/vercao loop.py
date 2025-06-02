from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
import smartsheet
import time
import re
import schedule


from dotenv import load_dotenv
import os

load_dotenv()

ACCESS_TOKEN = os.getenv('ACCESS_TOKEN')
SHEET_ID = int(os.getenv('SHEET_ID'))
edge_driver_path = os.getenv('EDGE_DRIVER_PATH')
COL_O_QUE_DESEJA_FAZER = int(os.getenv('COL_O_QUE_DESEJA_FAZER'))
COL_SOLICITACAO = int(os.getenv('COL_SOLICITACAO'))
COL_CODIGO_ATIVACAO = int(os.getenv('COL_CODIGO_ATIVACAO'))
COL_IMEI = int(os.getenv('COL_IMEI'))
COL_CRS = int(os.getenv('COL_CRS'))
COL_STATUS = int(os.getenv('COL_STATUS'))
COL_PROBLEMA = int(os.getenv('COL_PROBLEMA'))
LOGIN_USER = os.getenv('LOGIN_USER')
LOGIN_PASS = os.getenv('LOGIN_PASS')



# Configuração do Selenium
edge_driver_path = r"C:\Users\joao.escaim\Downloads\edgedriver_win64\msedgedriver.exe"
driver_service = Service(edge_driver_path)
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")

def fazer_login():
    print("Realizando login...")
    driver.get("https://portal.gpssa.com.br/GPS/Login.aspx")
    time.sleep(3)
    driver.find_element(By.ID, "txtUsername-inputEl").send_keys(LOGIN_USER)
    driver.find_element(By.ID, "txtPassword-inputEl").send_keys(LOGIN_PASS)
    time.sleep(1)
    driver.find_element(By.ID, "btnLogin-btnInnerEl").click()
    time.sleep(3)
    print("Login concluído.")

def processar_dispositivo(codigo_ativacao, imei, crs, row):
    print(f"\nProcessando dispositivo - Código: {codigo_ativacao}, IMEI: {imei}")

    imei = ''.join(filter(str.isdigit, str(imei)))
    if len(imei) < 15:
        client.Sheets.update_rows(SHEET_ID, [
            smartsheet.models.Row({
                "id": row.id,
                "cells": [
                    {"column_id": COL_STATUS, "value": "Solicitação Recusada"},
                    {"column_id": COL_PROBLEMA, "value": "Imei informado esta incorreto."}
                ]
            })
        ])
        return
    elif len(imei) > 15:
        imei = imei[:15]

    print("Navegando para a página de Código de ativação...")
    driver.get("https://portal.gpssa.com.br/SAD/CodigoAtivacao?IDPAGINA=1085&GRUPOABA=SAD%20-%20C%C3%B3digo%20de%20Ativa%C3%A7%C3%A3o&_dc=1743450910564")
    wait.until(EC.presence_of_element_located((By.ID, "gridcolumn-1016-textInnerEl")))
    print("Página de Código de ativação carregada.")

    time.sleep(5)

    # Preenche o código de ativação
    driver.find_element(By.ID, "textfield-1018-inputEl").send_keys(codigo_ativacao)
    time.sleep(5)

    # Clica no editar
    try:
        driver.find_element(By.CSS_SELECTOR, "div[cmd='Editar'].icon-pencil").click()
        time.sleep(3)
    except NoSuchElementException:
        client.Sheets.update_rows(SHEET_ID, [
            smartsheet.models.Row({
                "id": row.id,
                "cells": [
                    {"column_id": COL_STATUS, "value": "Solicitação Recusada"},
                    {"column_id": COL_PROBLEMA, "value": "Código de ativação não localizado"}
                ]
            })
        ])
        return 

    # Preenche o IMEI
    campo_imei = driver.find_element(By.ID, "txtIMEI-inputEl")
    campo_imei.clear()
    campo_imei.send_keys(imei)
    time.sleep(3)

    # Validação de código
    driver.find_element(By.ID, "btnValidarCodigo-btnIconEl").click()
    time.sleep(3)
    driver.find_element(By.ID, "button-1006-btnInnerEl").click()
    time.sleep(6)
    
    # Configurar dispositivo
    print("Configurando dispositivo...")

    print("Navegando para a página de cadastro de Dispositivos...")
    driver.get("https://portal.gpssa.com.br/SAD/Dispositivos?IDPAGINA=562&GRUPOABA=SAD%20-%20Cadastro%20de%20Dispositivos&_dc=1743508635081")
    wait.until(EC.presence_of_element_located((By.ID, "textfield-1024-inputEl")))
    print("Página de cadastro de Dispositivos carregada.")

    time.sleep(10)
    # Preenche o código de ativação
    driver.find_element(By.ID, "textfield-1024-inputEl").send_keys(codigo_ativacao)
    time.sleep(5)

    # Tenta clicar no editar - SEÇÃO MODIFICADA
    try:
        driver.find_element(By.CSS_SELECTOR, "span.icon-pencil").click()
        time.sleep(3)
    except NoSuchElementException:
        print("Botão Editar não encontrado. Verificando dispositivos duplicados...")
        
        # Volta para a página de Código de Ativação para limpar IMEIs duplicados
        print("Navegando para a página de Código de ativação...")
        driver.get("https://portal.gpssa.com.br/SAD/CodigoAtivacao?IDPAGINA=1085&GRUPOABA=SAD%20-%20Código%20de%20Ativação&_dc=1743450910564")
        wait.until(EC.presence_of_element_located((By.ID, "gridcolumn-1016-textInnerEl")))
        
        # Busca por todos os dispositivos com o mesmo IMEI
        imei_input = driver.find_element(By.ID, "textfield-1021-inputEl")
        imei_input.clear()
        imei_input.send_keys(imei)
        time.sleep(2)
        
        # Processa todos os registros encontrados
        while True:
            try:
                edit_buttons = driver.find_elements(By.CSS_SELECTOR, "div[cmd='Editar'].icon-pencil")
                if not edit_buttons:
                    break
                    
                for button in edit_buttons:
                    try:
                        button.click()
                        time.sleep(3)
                        
                        # Altera IMEI para 1
                        imei_field = driver.find_element(By.ID, "txtIMEI-inputEl")
                        imei_field.clear()
                        imei_field.send_keys("1")
                        time.sleep(2)
                        
                        # Valida código
                        driver.find_element(By.ID, "btnValidarCodigo-btnInnerEl").click()
                        time.sleep(2)
                        driver.find_element(By.ID, "button-1006-btnInnerEl").click()
                        time.sleep(8)
                        
                    except StaleElementReferenceException:
                        continue
                        
            except Exception as e:
                print(f"Erro durante limpeza de IMEIs: {str(e)}")
                break
        
        print("IMEIs duplicados alterados. Reiniciando processo...")
        # Chama a função novamente com os mesmos parâmetros
        return processar_dispositivo(codigo_ativacao, imei, crs, row)
    #configurar dispositivo
    print("configurar dispositivo...")

    #clica em status
    driver.find_element(By.ID, "cbbStatus-inputEl").click()
    time.sleep(3)

    #clicar em ativo
    driver.find_element(By.XPATH, "//li[contains(text(), 'Ativo')]").click()
    time.sleep(3)
    
    # Implantação: Produção
    driver.find_element(By.ID, "cbbImplantacao-inputEl").click()
    time.sleep(2)
    driver.find_element(By.XPATH, "//li[contains(text(), 'Produção')]").click()
    time.sleep(2)
    
    # Fuso horário: Brasília
    driver.find_element(By.ID, "cbbFusohorario-inputEl").click()
    time.sleep(2)
    driver.find_element(By.XPATH, "//li[contains(., '(UTC -3) BRASILIA')]").click()
    time.sleep(2)

    # Adicionar CRs
    print("Adicionando CRs...")
    element = driver.find_element(By.ID, "cbbQueryCr-inputEl")
    cr_list = [cr[:5] for cr in crs.split()]

    for cr in cr_list:
        element.clear()
        element.send_keys(cr)
        time.sleep(4)
        element.send_keys(Keys.ENTER)
        time.sleep(4)

    time.sleep(10)

    # Selecionar checkbox (se necessário)
    try:
        checkbox = driver.find_element(By.XPATH, "//span[contains(@class, 'x-column-header-checkbox')]")
        checkbox.click()
    except:
        pass
    
    botoes = driver.find_elements(By.XPATH, "//span[contains(@class, 'icon-pencil')]")
    
    if len(botoes) >= 2:
        botoes[1].click()
    else:
        print("Menos de dois botões encontrados.")

    try:
        checkbox_label = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//label[@for='chkSelf-inputEl']"))
        )

        # Verifica se está marcado pelo input hidden
        try:
            driver.find_element(By.XPATH, "//input[@type='hidden' and @name='chkSelf']")
            # Se marcado, clica duas vezes
            checkbox_label.click()
            checkbox_label.click()
        except:
            # Se desmarcado, clica uma vez
            checkbox_label.click()

    except Exception as e:
        print("Erro ao clicar no checkbox via label:", e)

    # Salvar alterações
    print("Salvando alterações...")
    save_button = driver.find_element(By.XPATH, "//span[@data-ref='btnInnerEl' and contains(@class, 'x-btn-inner') and text()='Salvar']")
    save_button.click()
    time.sleep(10)

    # Localiza o container da tabela onde os CRs estão listados
    grid_container = driver.find_element(By.CLASS_NAME, "x-grid-item-container")

    # Localiza o elemento que contém os números de CR (ajustado para o novo XPath)
    cr_element = grid_container.find_element(
        By.XPATH,
        ".//td[@data-columnid='gridcolumn-1046']/div[@class='x-grid-cell-inner ']"
    )

    # Obtém o texto completo e separa os CRs
    cr_texto_completo = cr_element.text.strip()
    cr_numeros_tela = [cr.strip() for cr in cr_texto_completo.split(';') if cr.strip()]

    print("CRs encontrados na página:", cr_numeros_tela)

    # Compara com a lista original cr_list (supondo que cr_list já está definida)
    crs_nao_cadastrados = [cr for cr in cr_list if cr not in cr_numeros_tela]

    # Exibe os CRs que não foram cadastrados corretamente
    if crs_nao_cadastrados:
        print("CRs NÃO encontrados na tela:", crs_nao_cadastrados)
    else:
        print("Todos os CRs foram cadastrados com sucesso.")

    # Prepara os dados para atualização no Smartsheet
    smartsheet_data = []

    if crs_nao_cadastrados and cr_numeros_tela:
        smartsheet_data = [
            {"column_id": COL_STATUS, "value": "Solicitação Recusada"},
            {"column_id": COL_PROBLEMA, "value": f"Aparelho cadastrado, porém CR {', '.join(crs_nao_cadastrados)} não foram cadastrados."}
        ]
    elif not cr_numeros_tela:
        smartsheet_data = [
            {"column_id": COL_STATUS, "value": "Solicitação Recusada"},
            {"column_id": COL_PROBLEMA, "value": "CR não foi localizado"}
        ]
    else:
        smartsheet_data = [
            {"column_id": COL_STATUS, "value": "Cadastro Realizado"},
            {"column_id": COL_PROBLEMA, "value": ""}
        ]

    # Atualiza o Smartsheet
    client.Sheets.update_rows(SHEET_ID, [
        smartsheet.models.Row({
            "id": row.id,
            "cells": smartsheet_data
        })
    ])
    
    print(f"Processamento concluído para o código {codigo_ativacao}")

def executar_tudo():
    global driver, wait, client, sheet
    # Inicializar Smartsheet API e Selenium a cada execução
    client = smartsheet.Smartsheet(ACCESS_TOKEN)
    sheet = client.Sheets.get_sheet(SHEET_ID)
    driver = webdriver.Edge(service=driver_service, options=options)
    wait = WebDriverWait(driver, 30)
    try:
        fazer_login()
        for row in sheet.rows:
            o_que_deseja_fazer = next((c.value for c in row.cells if c.column_id == COL_O_QUE_DESEJA_FAZER), None)
            if o_que_deseja_fazer == "Cadastro de Aparelho (GPSA)":
                codigo_ativacao = next((c.value for c in row.cells if c.column_id == COL_CODIGO_ATIVACAO), None)
                imei = next((c.value for c in row.cells if c.column_id == COL_IMEI), None)
                crs = next((c.value for c in row.cells if c.column_id == COL_CRS), None)
                if codigo_ativacao and imei and crs:
                    print(f"\nIniciando processamento para: Código {codigo_ativacao}")
                    processar_dispositivo(codigo_ativacao, imei, crs, row)
    finally:
        driver.quit()
        print("Processo concluído. Navegador fechado.")

# Agendamento nos horários desejados
schedule.every().day.at("07:00").do(executar_tudo)
schedule.every().day.at("10:00").do(executar_tudo)
schedule.every().day.at("13:00").do(executar_tudo)
schedule.every().day.at("15:00").do(executar_tudo)
schedule.every().day.at("18:00").do(executar_tudo)

print("Agendador iniciado. Aguardando horário programado...")

while True:
    schedule.run_pending()
    time.sleep(1)