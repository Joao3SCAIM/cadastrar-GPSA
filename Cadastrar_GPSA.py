from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
import smartsheet
import time
import re
import schedule
from dotenv import load_dotenv
import os

import requests # (NOVO)
import traceback # (NOVO)
import urllib3 # (NOVO)

# (NOVO) Desativa o aviso de requisição HTTPS não verificada no terminal
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

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
COL_INTERVALO = int(os.getenv('COL_INTERVALO'))

dotenv_path_login = r"C:\Users\Administrator\gpssa.com.br\DashboardBriansilva - Automacoes - Documentos\TIP\.env"
load_dotenv(dotenv_path_login)

LOGIN_USER = os.getenv('USER_PORTAL_FERNANDO')
LOGIN_PASS = os.getenv('SENHA_PORTAL_FERNANDO')


# (NOVO) URL DO HEALTHCHECK
HC_URL = "https://healthcheck.timeinovacaoeprojeto.shop/ping/c9177bb2-8ee4-4e48-b389-0448dbbefb9f"

# Configuração do Selenium
edge_driver_path = r"C:\WebDrivers\msedgedriver.exe"
driver_service = Service(edge_driver_path)
options = webdriver.EdgeOptions()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--window-size=1920,1080")


# (NOVO) FUNÇÃO DO HEALTHCHECK
def ping_healthcheck(endpoint="", erro_log=None):
    """Envia o ping de forma silenciosa para não travar o robô se a rede oscilar"""
    try:
        url = f"{HC_URL}{endpoint}"
        if erro_log:
            requests.post(url, data=erro_log.encode('utf-8'), timeout=10, verify=False)
        else:
            requests.get(url, timeout=10, verify=False)
    except Exception:
        pass


def fazer_login():
    driver.get("https://portal.gpssa.com.br/GPS/Login.aspx")
    time.sleep(10)
    driver.find_element(By.ID, "txtUsername-inputEl").send_keys(LOGIN_USER)
    driver.find_element(By.ID, "txtPassword-inputEl").send_keys(LOGIN_PASS)
    time.sleep(1)
    driver.find_element(By.ID, "btnLogin-btnInnerEl").click()
    time.sleep(3)


# --- ALTERADO: Adicionado 'intervalo' nos parâmetros ---
def processar_dispositivo(codigo_ativacao, imei, crs, intervalo, row):

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

    driver.get("https://portal.gpssa.com.br/SAD/CodigoAtivacao?IDPAGINA=1085&GRUPOABA=SAD%20-%20C%C3%B3digo%20de%20Ativa%C3%A7%C3%A3o&_dc=1743450910564")
    wait.until(EC.presence_of_element_located((By.ID, "gridcolumn-1016-textInnerEl")))
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
    time.sleep(3)
    campo_imei.send_keys(imei)
    time.sleep(3)

    # Validação de código
    driver.find_element(By.ID, "btnValidarCodigo-btnIconEl").click()
    time.sleep(3)
    driver.find_element(By.ID, "button-1006-btnInnerEl").click()
    time.sleep(6)
    
    # Configurar dispositivo
    driver.get("https://portal.gpssa.com.br/SAD/Dispositivos?IDPAGINA=562&GRUPOABA=SAD%20-%20Cadastro%20de%20Dispositivos&_dc=1743508635081")
    wait.until(EC.presence_of_element_located((By.ID, "textfield-1024-inputEl")))
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
        driver.get("https://portal.gpssa.com.br/SAD/CodigoAtivacao?IDPAGINA=1085&GRUPOABA=SAD%20-%20Código%20de%20Ativação&_dc=1743450910564")
        wait.until(EC.presence_of_element_located((By.ID, "gridcolumn-1016-textInnerEl")))
        
        # Busca por todos os dispositivos com o mesmo IMEI
        time.sleep(3)
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
        # --- ALTERADO: Passando a variável intervalo na chamada recursiva ---
        return processar_dispositivo(codigo_ativacao, imei, crs, intervalo, row)

    # clicar em status
    driver.find_element(By.ID, "cbbStatus-inputEl").click()
    time.sleep(3)

    # clicar em ativo
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

    # =======================================================
    # NOVO: Regra do Intervalo / Refeição
    # =======================================================
    if str(intervalo).strip().lower() == "sim":
        try:
            # Clica na caixa de seleção para abrir as opções
            driver.find_element(By.ID, "cbbTemRefeicao-inputEl").click()
            time.sleep(2)
            
            # Clica na opção "Sim" (usando text() exato para evitar erros com palavras parecidas)
            driver.find_element(By.XPATH, "//li[text()='Sim']").click()
            time.sleep(2)
        except Exception as e:
            print(f"Erro ao selecionar 'Sim' no campo de refeição/intervalo: {e}")
    # =======================================================

    # Salvar alterações
    save_button = driver.find_element(By.XPATH, "//span[@data-ref='btnInnerEl' and contains(@class, 'x-btn-inner') and text()='Salvar']")
    save_button.click()
    time.sleep(10)

    # Localiza o container da tabela onde os CRs estão listados
    grid_container = driver.find_element(By.CLASS_NAME, "x-grid-item-container")

    # Localiza o elemento que contém os números de CR
    cr_element = grid_container.find_element(
        By.XPATH,
        ".//td[@data-columnid='gridcolumn-1046']/div[@class='x-grid-cell-inner ']"
    )

    # Obtém o texto completo e separa os CRs
    cr_texto_completo = cr_element.text.strip()
    cr_numeros_tela = [cr.strip() for cr in cr_texto_completo.split(';') if cr.strip()]

    # Compara com a lista original
    crs_nao_cadastrados = [cr for cr in cr_list if cr not in cr_numeros_tela]

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

def executar_tudo():
    global driver, wait, client, sheet
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
                
                # --- NOVO: Captura da coluna de Intervalo ---
                intervalo = next((c.value for c in row.cells if c.column_id == COL_INTERVALO), None)
                
                if codigo_ativacao and imei and crs:
                    # --- ALTERADO: Passando 'intervalo' para a função ---
                    processar_dispositivo(codigo_ativacao, imei, crs, intervalo, row)
    finally:
        driver.quit()

def linhas_pendentes():
    client = smartsheet.Smartsheet(ACCESS_TOKEN)
    sheet = client.Sheets.get_sheet(SHEET_ID)
    for row in sheet.rows:
        o_que_deseja_fazer = next((c.value for c in row.cells if c.column_id == COL_O_QUE_DESEJA_FAZER), None)
        if o_que_deseja_fazer == "Cadastro de Aparelho (GPSA)":
            return True
    return False


if __name__ == "__main__":
    try:
        ping_healthcheck("/start") # (NOVO) 1. Avisa que começou
        
        while True:
            executar_tudo()
            if not linhas_pendentes():
                break
            print("Ainda existem linhas pendentes. Reiniciando o processo em 30 segundos...")
            time.sleep(10)
            
        ping_healthcheck() # (NOVO) 2. Avisa sucesso caso conclua tudo sem problemas
        print("Fluxo finalizado com sucesso.")
        
    except Exception as e:
        # (NOVO) 3. Avisa ao Healthcheck capturando o log real de qualquer erro/queda
        erro_log = traceback.format_exc()
        ping_healthcheck("/fail", erro_log)
        print(f"Erro Crítico: {e}")
