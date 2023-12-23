import csv
import io
import json
import os
from datetime import datetime, timedelta

import general as apr
import pandas as pd
import pyautogui as p
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time 

def cria_pasta(*args):
    for pasta in args:
        if not os.path.exists(pasta):  # Verifica se a pasta já existe
            os.makedirs(pasta)  # Cria a pasta se ela não existir


def localizar_elemento(driver: WebDriver, locator: tuple, tempo_espera=10, codicao=EC.presence_of_element_located) -> WebElement:
    try:
        elemento = WebDriverWait(driver, tempo_espera).until(codicao(locator))
        return elemento
    except:
        apr.display_error()
        return


def clicar_elemento(driver: WebDriver, locator: tuple, tempo_espera=10):
    elemento = localizar_elemento(driver, locator, tempo_espera, codicao=EC.element_to_be_clickable)
    if isinstance(elemento, WebElement):
        elemento.click()


def input_evento(driver: WebDriver, inp: WebElement, item:str ="", tentativas=10, sleep=1.3):
    rodar = True
    while rodar and tentativas > -1:
        driver.execute_script(f"""
                            let event = new Event('input', {{ bubbles: true }});
                            arguments[0].value = arguments[1];
                            arguments[0].dispatchEvent(event);""",inp, "")
        driver.execute_script(f"""
                            let event = new Event('input', {{ bubbles: true }});
                            arguments[0].value = arguments[1];
                            arguments[0].dispatchEvent(event);""",inp, str(item))
        valor = driver.execute_script("return arguments[0].value;", inp)
        if valor == str(item):
            rodar = False
        if rodar:
            p.sleep(sleep)
        tentativas -= 1


def aguardar_download_concluir(diretorio_download, comando, nome_arquivo: str, tempo_espera=30, vezes=3):
    # Inicia o tempo de espera
    tempo_inicio = time.time()
    nome_arquivo = nome_arquivo.lower()
    while True:
        # Obtém uma lista de todos os arquivos no diretório de download
        # Obtém uma lista de todos os arquivos no diretório de download que começam com o nome do arquivo
        arquivos = [entry.path for entry in os.scandir(diretorio_download) if entry.name.lower().startswith(nome_arquivo) and entry.name.lower().endswith('.csv')]
        # Se a lista de arquivos não estiver vazia, um arquivo foi baixado
        if arquivos:
            print("O arquivo foi baixado com sucesso.")
            # Retorna o caminho do arquivo mais recente
            return max(arquivos, key=os.path.getctime)
        
        # Verifica se o tempo de espera foi excedido
        if time.time() - tempo_inicio > tempo_espera:
            print("O arquivo não foi baixado.")
            comando()
            
            # Diminui o valor de vezes
            vezes -= 1
            
            # Verifica se ainda há tentativas restantes
            if vezes > 0:
                return aguardar_download_concluir(diretorio_download, comando,nome_arquivo, tempo_espera, vezes)
            else:
                return None

        # Aguarda antes de verificar novamente
        time.sleep(1)


def remover_arquivos(diretorio: str, nome: str):
    # Obtém uma lista de todos os arquivos no diretório que começam com NOME
    arquivos = os.scandir(diretorio)
    # Exclui cada arquivo
    for arquivo in arquivos:
        if arquivo.name.lower().startswith(nome.lower()):
            os.remove(arquivo)


def confere_valor_diferenca(df, index, valor_soul_mv, valor_getnet, valores_iguais):
    if abs(0.05) >= abs(valor_soul_mv - valor_getnet) and not valores_iguais:
        df.at[index, 'DIFERENCA'] = abs(valor_soul_mv - valor_getnet)
        if (valor_getnet - valor_soul_mv) > 0:
            df = colocando_msg_df(df, index, 'OBS', 'AJUSTE DE RECEBIMENTO')
        else:
            df = colocando_msg_df(df, index, 'OBS', 'AJUSTE DE TARIFA')
        df = colocando_msg_df(df, index, 'PROCESSAR', 'S')
    elif abs(0.05) < abs(valor_soul_mv - valor_getnet):
        df.at[index, 'DIFERENCA'] = abs(valor_soul_mv - valor_getnet)
        if (valor_getnet - valor_soul_mv) > 0:
            df = colocando_msg_df(df, index, 'OBS', 'AJUSTE DE RECEBIMENTO')
        else:
            df = colocando_msg_df(df, index, 'OBS', 'AJUSTE DE TARIFA')
    elif valores_iguais:
        df = colocando_msg_df(df, index, 'DIFERENCA', '0')
        df = colocando_msg_df(df, index, 'PROCESSAR', 'OK', replace=True)
    return df


def obtem_bool_venvimento_valores(df, index, data_soul_mv, valor_soul_mv, vencimento_getnet, valor_getnet):
    vencimentos_iguais = False
    valores_iguais = False
    if data_soul_mv == vencimento_getnet:
        vencimentos_iguais = True
    if valor_soul_mv == valor_getnet:
        valores_iguais = True
    if vencimentos_iguais and valores_iguais:
        df.at[index, 'PROCESSAR'] = 'OK'
        df = colocando_msg_df(df, index, 'OBS', 'TUDO CORRETO')
    return vencimentos_iguais,valores_iguais, df


def colocando_msg_df(df, index, colunm, msg, **kwargs):
    if kwargs.get('replace'):
        df.at[index, colunm] = msg
    elif pd.isna(df.at[index, colunm]):
        df.at[index, colunm] = msg
    elif not pd.isna(df.at[index, colunm]):
        df.at[index, colunm] += f'\n{msg}'
    
    return df
