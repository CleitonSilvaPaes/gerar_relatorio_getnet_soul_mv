"""
    Conciliacao cartao
"""

import json
import os
from datetime import datetime

import general as apr
import pandas as pd
import pyautogui as p

from selenium.webdriver.chrome.webdriver import WebDriver
from decimal import Decimal
import re

from processo.src.getnet import Gatnet
from processo.src.soulmv import SoulMv
from processo.src.util import *
from dotenv import dotenv_values

def transformando_novo_df(df_getnet: pd.DataFrame, df_soul_mv: pd.DataFrame):
    # Removendo os espacos das colunas
    df_getnet["AUTORIZACAO"] = df_getnet["AUTORIZACAO"].str.strip()                  
    
    colunas_para_excluir = ["CODIGO CENTRALIZADOR", "CODIGO", "VENCIMENTO", 
                            "PRODUTO", "LANCAMENTO", "PLANO DE VENDA",
                            "PARCELA", "TOTAL DE PARCELAS", "CARTAO", "NUMERO CV", 
                            "DATA DA VENDA", "VALOR ORIGINAL", "VALOR BRUTO", "DESCONTOS"]

    df_getnet_selecionado = df_getnet[df_getnet.columns.difference(colunas_para_excluir)]

    df_processar = df_soul_mv.merge(df_getnet_selecionado, left_on='NR DOCUMENTO', right_on='AUTORIZACAO', how='outer')
    
    # Lista atual de colunas
    colunas = df_processar.columns.tolist()
    
    # Remova as colunas que você deseja mover para o final
    for coluna in ['LIBERAR DATA', 'DIFERENCA', 'PROCESSAR', 'OBS']:
        colunas.remove(coluna)
        
    colunas += ['LIBERAR DATA', 'DIFERENCA', 'PROCESSAR', 'OBS']

    # Reordene as colunas do DataFrame
    df_processar = df_processar[colunas]
    return df_processar


def processo_retirar_relatorio_getnet(msg_json, credendias, driver, periodo_registro, PATH_DOWNLOAD_GETNET, PATH_IMA):
    gatnet = Gatnet(driver, msg_json, credendias, periodo_registro, PATH_DOWNLOAD_GETNET)
    if not os.path.exists(gatnet.nome_relatorio):
        gatnet.fazer_login(credendias["usu_getnet"], credendias["senha_getnet"])
        gatnet.acessar_pagina_recebimento()
        gatnet.exportar_recebimentos(msg_json["getnet_estabelecimento"], PATH_IMA)
    return gatnet


def processo_retirar_relatorio_soul(msg_json, credendias, driver, ontem, PATH_IMA, PATH_DOWNLOAD_SOUL_MV):
    soul_mv = SoulMv(driver, ontem, PATH_DOWNLOAD_SOUL_MV)
    if not os.path.exists(soul_mv.nome_relatorio):
        soul_mv.fazer_login_soul(msg_json, credendias)
        if soul_mv.acessar_page_retirar_relatorio(msg_json["rel_ext_bnc"]):
            print("Acessou pagina retirar relatorio")
            caminho_relatorio = soul_mv.retirar_relatorio_extrato_bancario(msg_json, ontem, PATH_IMA)
            soul_mv.ler_relatorio_e_normalizar(caminho_relatorio)
    return soul_mv


def processar_df_verificando_valores_e_data(df_processar):
    for index in df_processar.index:
        # E feita uma conferencia para verifica se autorizacao esta
        # Vazia para assim continuar o lancamento
        if pd.isna(df_processar.at[index, 'OBS']) and not pd.isna(df_processar.at[index, 'NR DOCUMENTO']) and not pd.isna(df_processar.at[index, 'AUTORIZACAO']):
            data_soul_mv = df_processar.at[index, "DATA"]
            valor_soul_mv = round(Decimal(df_processar.at[index, "VALOR"]), 2)
                            
            vencimento_getnet = df_processar.at[index, "VENCIMENTO ORIGINAL"]
            valor_getnet = round(Decimal(df_processar.at[index, "LIQUIDO"]), 2)
                            
            # CONFERENCIA DOS DADOS
            vencimentos_iguais, valores_iguais, df_processar = obtem_bool_venvimento_valores(df_processar, index, data_soul_mv, valor_soul_mv, vencimento_getnet, valor_getnet)
                                
            # VERIFICA DATA VENCIEMNTO 
            if not vencimentos_iguais:
                df_processar = colocando_msg_df(df_processar, index, 'OBS', f'DATA INCORRETA {data_soul_mv}')
                df_processar = colocando_msg_df(df_processar, index, 'LIBERAR DATA', 'S')
            else:
                df_processar = colocando_msg_df(df_processar, index, 'LIBERAR DATA', 'OK')
                                
            # CONFERENCIA VALORES
            df_processar = confere_valor_diferenca(df_processar, index, valor_soul_mv, valor_getnet, valores_iguais)
    return df_processar


def processa_verificando_se_contem_getnet_soul(df_processar):
    for index in df_processar.index:
        if not pd.isna(df_processar.at[index, 'NR DOCUMENTO']) and pd.isna(df_processar.at[index, 'AUTORIZACAO']) \
                            and pd.isna(df_processar.at[index, 'OBS']):
            df_processar.at[index, 'OBS'] = 'NAO ESTA PRESENTE AUTORIZACAO NO RELATORIO GETNET'
            df_processar.at[index, 'LIBERAR DATA'] = 'N'
            df_processar.at[index, 'PROCESSAR'] = 'N'
            df_processar.at[index, 'DIFERENCA'] = '0'
        elif pd.isna(df_processar.at[index, 'NR DOCUMENTO']) and not pd.isna(df_processar.at[index, 'AUTORIZACAO']) \
                            and pd.isna(df_processar.at[index, 'OBS']):
                df_processar.at[index, 'OBS'] = 'NAO ESTA PRESENTE AUTORIZACAO NO RELATORIO SOUL MV'
                df_processar.at[index, 'LIBERAR DATA'] = 'N'
                df_processar.at[index, 'PROCESSAR'] = 'N'
                df_processar.at[index, 'DIFERENCA'] = '0'


def processar_relatorios_soulmv_e_getnet(msg_json: dict, periodo_registro_match: re.Match, gatnet_instancia: Gatnet, 
                                        soul_mv_instancia: SoulMv, arquivo_instancia: os.DirEntry, PATH_IMA: str):
    periodo_registro = periodo_registro_match.group()
    caminho_arquivo = os.path.dirname(arquivo_instancia.path)

    # Monta o caminho do arquivo da GETNET
    arquivo_getnet = os.path.join(caminho_arquivo, f'relatorio_getnet_{periodo_registro}.csv')

    # Lê e normaliza o relatório do SoulMV
    dataframe_soul_mv = soul_mv_instancia.ler_relatorio_e_normalizar(arquivo_instancia.path)

    # Verifica se o arquivo da GETNET existe; se não, realiza o download e lê pelo pandas
    if not os.path.exists(arquivo_getnet):
        gatnet_instancia.periodo_registro = periodo_registro
        gatnet_instancia.fazer_login(msg_json["usu_getnet"], msg_json["senha_getnet"])
        gatnet_instancia.acessar_pagina_recebimento()
        dataframe_getnet = gatnet_instancia.exportar_recebimentos(msg_json["getnet_estabelecimento"], PATH_IMA)
    else:
        dataframe_getnet = pd.read_csv(arquivo_getnet, sep=';', dtype={"AUTORIZACAO": str})

    return periodo_registro, dataframe_soul_mv, dataframe_getnet


def main(p_ativ):
    p.FAILSAFE = False

    # PASTA PADRAO
    PATH_DOC = os.path.join(os.path.dirname(__file__), p_ativ, 'doc')
    PATH_IMA = os.path.join(os.path.dirname(__file__), p_ativ, 'ima')
    PATH_IMAOUT = os.path.join(os.path.dirname(__file__), p_ativ, 'imaout')
    PATH_OUT = os.path.join(os.path.dirname(__file__), p_ativ, 'out')
    PATH_PRO = os.path.join(os.path.dirname(__file__), p_ativ, 'pro')
    PATH_DOWNLOAD_GETNET = os.path.join(PATH_PRO, "download", "getnet")
    PATH_DOWNLOAD_SOUL_MV = os.path.join(PATH_PRO, "download", "soul_mv")

    # Verifica se existe a pasta para o downloadd dos arquivos
    cria_pasta(PATH_DOWNLOAD_GETNET, PATH_DOWNLOAD_SOUL_MV)

    msg_json = json.load(open(os.path.join(PATH_DOC, "DGLMsg.json"), "r", encoding="utf-8"))
    credendias = dotenv_values('.env')

    driver = apr.get_driver()
    # driver = None
    ontem = datetime.now() - timedelta(days=1)
    # ontem = datetime(day=28, month=11, year=2023)
    periodo_registro = ontem.strftime('%d%m%Y')

    try:
        if isinstance(driver, WebDriver) or driver is None:
            
            # Caso nao tenha relatorio fazer o download do relatorio
            # Processo Relatorio retirada relatorio
            # --------------------------------------------------------------------------------------
            # GETNET
            gatnet = processo_retirar_relatorio_getnet(msg_json, credendias, driver, periodo_registro, PATH_DOWNLOAD_GETNET, PATH_IMA)

            # SOUL MV
            soul_mv = processo_retirar_relatorio_soul(msg_json, credendias, driver, ontem, PATH_IMA, PATH_DOWNLOAD_SOUL_MV)
            # --------------------------------------------------------------------------------------
            
            # Fazer um copia dos relatorio e mantar para pasta pro
            # --------------------------------------------------------------------------------------
            soul_mv.copiar_para_pasta_pro(PATH_PRO)
            gatnet.copiar_para_pasta_pro(PATH_PRO)
            # --------------------------------------------------------------------------------------
            
            # Procesando os relatorio 
            # --------------------------------------------------------------------------------------
            padrao_data_arq = r'\d{8}'
            for arq in os.scandir(PATH_PRO):
                # Obten a data escrita no arquivo
                periodo_registro = re.search(padrao_data_arq, arq.name)
                if arq.name.lower().endswith('.csv') and 'extrato_banco_soul_mv' in arq.name.lower():
                    periodo_registro, df_soul_mv, df_getnet = processar_relatorios_soulmv_e_getnet(msg_json, periodo_registro,
                                                                                                    gatnet, soul_mv, arq, PATH_IMA)
                    
                    df_processar = transformando_novo_df(df_getnet, df_soul_mv)

                    
                    # Colocando informacao quando nao encontrado a autorizacao
                    processa_verificando_se_contem_getnet_soul(df_processar)
                    
                    # df_processar.to_csv(os.path.join(PATH_PRO, f'relatorio_getnet_soulmv_{periodo_registro}.csv'), sep=';', index=False)
                    df_processar = processar_df_verificando_valores_e_data(df_processar)
                    
                    df_processar.to_csv(os.path.join(PATH_PRO, f'relatorio_getnet_soulmv_{periodo_registro}.csv'), sep=';', index=False)
        
        else:
            raise "Nao foi possivel obter o webdriver"
        
    except Exception as e:
        print(e.__class__(e))
    finally:
        if isinstance(driver, WebDriver):
            driver.quit() 


if __name__ == '__main__':
    main("processo")