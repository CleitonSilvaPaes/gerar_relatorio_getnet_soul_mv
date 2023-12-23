import csv
import io
import json
import os
from datetime import datetime, timedelta

import general as apr
import pandas as pd
import pyautogui as p
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from decimal import Decimal
import re
import shutil

from .util import *

class SoulMv():
    def __init__(self, driver: WebDriver, periodo_registro: datetime, PATH_DOWNLOAD_SOUL_MV:str):
        self.driver = driver
        self.PATH_DOWNLOAD_SOUL_MV = PATH_DOWNLOAD_SOUL_MV
        self.XPATH_IFRAME = (By.XPATH, '//*[@id="child_FINAN.HTML"]')
        self.nome_relatorio = os.path.join(PATH_DOWNLOAD_SOUL_MV, f'extrato_banco_soul_mv_{periodo_registro.strftime("%d%m%Y")}.csv')

    def fazer_login_soul(self, msg_json: dict, credendias: dict) -> tuple[bool, str]:
        self.driver.get(credendias["site_soul_teste"])
        XPATH_USUARIO = (By.XPATH, '//*[@id="username"]')
        XPATH_SENHA = (By.XPATH, '//*[@id="password"]')
        XPATH_BTN_LOGIN = (By.XPATH, '//*[@id="context_login"]/section[4]/input[7]')
        XPATH_OPCAO_EMPRESA = (By.XPATH, '//*[@id="companies"]')

        try:
            input_usuario = localizar_elemento(self.driver, XPATH_USUARIO, codicao=EC.element_to_be_clickable)
            input_senha = localizar_elemento(self.driver, XPATH_SENHA, codicao=EC.element_to_be_clickable)
            opcao_empresa = localizar_elemento(self.driver, XPATH_OPCAO_EMPRESA, codicao=EC.element_to_be_clickable)
            btn_login = localizar_elemento(self.driver, XPATH_BTN_LOGIN, codicao=EC.element_to_be_clickable)

            input_usuario.send_keys(credendias["usu_soul"])
            input_senha.send_keys(credendias["senha_soul"])

            if isinstance(opcao_empresa, WebElement):
                selecao_empresa = Select(opcao_empresa)
                selecao_empresa.select_by_value('1')
            btn_login.click()
            return True, 'Sucesso'
        except Exception:
            return False, apr.display_error()

    def acessar_page_retirar_relatorio(self, nome_page):
        XPATH_BTN_CONTROLADORIA = (By.CSS_SELECTOR, '#workspace-menubar > ul > li.menu-node > a')
        XPATH_CONTROLE_FINANCEIRO = (By.CSS_SELECTOR, '#workspace-menubar > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node > a')
        XPATH_BTN_RELATORIOS = (By.CSS_SELECTOR, '#workspace-menubar > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node.menu-parent-open > div > ul > li:nth-child(3) > a')
        XPATH_CONTROLE_BANCARIO = (By.CSS_SELECTOR, '#workspace-menubar > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node > a')
        XPATH_EXTRATO_BANCARIO = (By.CSS_SELECTOR, '#workspace-menubar > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node.menu-parent-open > div > ul > li.menu-node > a')
            
        if not self.selecionar_page_correta(nome_page, 10, 1):
            
            clicar_elemento(self.driver, XPATH_BTN_CONTROLADORIA, 20)
            clicar_elemento(self.driver, XPATH_CONTROLE_FINANCEIRO, 20)
            clicar_elemento(self.driver, XPATH_BTN_RELATORIOS, 20)
            clicar_elemento(self.driver, XPATH_CONTROLE_BANCARIO, 20)
            clicar_elemento(self.driver, XPATH_EXTRATO_BANCARIO, 20)

            return True
        return True

    def retirar_relatorio_extrato_bancario(self, msg_json:dict, periodo:datetime, PATH_IMA:str): 
        page_correta = self.selecionar_page_correta(msg_json["rel_ext_bnc"], 60, 1)
        
        # Remove o arquivo antigo da pasta download_directory
        remover_arquivos(self.PATH_DOWNLOAD_SOUL_MV, os.path.basename(self.nome_relatorio))
        
        XPATH_INPUT_TIPO_IMPRESSAO = (By.CSS_SELECTOR, '#frames10_ac')
        XAPTH_CONTA_CORRENTE = (By.CSS_SELECTOR, '#inp\:cdConCor')
        XPATH_DIA_INICIAL = (By.CSS_SELECTOR, '#inp\:dtInicial')
        XPATH_DIA_FINAL = (By.CSS_SELECTOR, '#inp\:dtFinal')
        XPATH_BOTAO_IMPRIMIR = (By.CSS_SELECTOR, '#frames20')

        IMAGEM_TELA_SALVAR_COMO = os.path.join(PATH_IMA, 'salvar_como.png')
        
        periodo_inicial = periodo
        periodo_final = periodo + timedelta(days=20)
        frame = localizar_elemento(self.driver, self.XPATH_IFRAME, 20, EC.frame_to_be_available_and_switch_to_it)
        if frame and page_correta:
            input_tipo_impressao = localizar_elemento(self.driver, XPATH_INPUT_TIPO_IMPRESSAO, 40)
            input_conta_corrente = localizar_elemento(self.driver, XAPTH_CONTA_CORRENTE, 20)
            input_dia_inicial = localizar_elemento(self.driver, XPATH_DIA_INICIAL, 20)
            input_dia_final = localizar_elemento(self.driver, XPATH_DIA_FINAL, 20)
            botao_imprimir = localizar_elemento(self.driver, XPATH_BOTAO_IMPRIMIR)

            input_evento(self.driver, input_conta_corrente, msg_json["banco_relatorio_bancario"])
            input_evento(self.driver, input_dia_inicial, periodo_inicial.strftime('%d%m%Y'))
            input_evento(self.driver, input_dia_final, periodo_final.strftime('%d%m%Y'))
            input_evento(self.driver, input_tipo_impressao, 'CSV')
            comando = self.driver.execute_script("arguments[0].click();", botao_imprimir)
            try:
                apr.validate_to_proceed(IMAGEM_TELA_SALVAR_COMO, 40, "SALVAR COMO", confidence=0.8)
                p.write(self.nome_relatorio)
                p.press('enter')
            except p.ImageNotFoundException:
                pass
            caminho_arquivo_csv = aguardar_download_concluir(self.PATH_DOWNLOAD_SOUL_MV, comando, os.path.basename(self.nome_relatorio))
            return caminho_arquivo_csv

    def selecionar_page_correta(self, nome_page, tentativa=10, intervalo=2):
        self.driver.switch_to.default_content()
        
        XPATH_UL_PAGE = (By.XPATH, '//*[@id="react"]/div/div[1]/ul')
        elemento = localizar_elemento(self.driver, XPATH_UL_PAGE, 20)
        if elemento:
            JS_CODE = f"""
                var ul = document.evaluate('{XPATH_UL_PAGE[1]}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    var lis = ul.querySelectorAll('li');
                    for (var i = 0; i < lis.length; i++) {{
                        var li = lis[i];
                        if (li.textContent.includes('{nome_page}')) {{
                            return li;
                        }}
                    }}
                    return null;
            """
            
            for _ in range(tentativa):
                elemento = self.driver.execute_script(JS_CODE)
                if isinstance(elemento, WebElement):
                    elemento.click()
                    return True
                print(f'PROCURANDO A PAGINA {nome_page}')
                p.sleep(intervalo)
            return False

    def ler_arquivo(self, caminho):
        with open(caminho, 'r') as file:
            content = file.read()
            content = apr.RemoverAcentos(content)
        return content

    def extrair_dados(self, reader):
        data_list = []
        historio_list = []
        nr_documento_list = []
        valor_list = []
        patter_date = r'\d{2}\/\d{2}\/\d{4}'
        patter_hist = r'TESOURARIA|DEPOSITO|GERAL'
        patter_valor = r'\d+\,\d{2}'
        try:
            for index, row in enumerate(reader):
                if index >= 7:
                    if len(row) < 5:
                        row = row[0]
                        row = re.split(r',(?=(?:[^"]*"[^"]*")*[^"]*$)', row)
                        row = [item.replace('"', '') for item in row]
                    for item in row:
                        match_hist = re.search(patter_hist, item)
                        if match_hist:
                            historio_list.append(item)
                            match_date = re.search(patter_date, str(row))
                            if match_date:
                                data_list.append(match_date.group())
                            else:
                                data_list.append(data_list[-1])
                            encontrou = False
                            for i, element in enumerate(row):
                                element = str(element).replace('.', '')
                                macth_valor = re.search(patter_valor, element)
                                if macth_valor:
                                    encontrou = True
                                    valor_list.append(macth_valor.group())
                                    nr_documento_list.append(row[i-1])
                            if not encontrou:
                                data_list.pop()
                                historio_list.pop()
        except:
            apr.display_error()
        return data_list, historio_list, nr_documento_list, valor_list

    def criar_dataframe(self, data_list, historio_list, nr_documento_list, valor_list):
        df = pd.DataFrame({
                "DATA": data_list,
                "HISTORIO": historio_list,
                "NR DOCUMENTO": nr_documento_list,
                "VALOR": valor_list,
                "LIBERAR DATA": None,
                "DIFERENCA": None,
                "PROCESSAR": None,
                "OBS": None
            })
        return df

    def ler_relatorio_e_normalizar(self, caminho: str):
        content = self.ler_arquivo(caminho)
        f = io.StringIO(content)
        reader = csv.reader(f, delimiter=',')
        data_list, historio_list, nr_documento_list, valor_list = self.extrair_dados(reader)
        df = self.criar_dataframe(data_list, historio_list, nr_documento_list, valor_list)
        if df.empty:
            df = pd.read_csv(caminho, sep=';')
        df = self.agrupar_dados(df)
        df.to_csv(caminho, sep=';', index=False)
        return df

    def agrupar_dados(self, df: pd.DataFrame):
        df["VALOR"] = df["VALOR"].astype(str)
        df["NR DOCUMENTO"] = df["NR DOCUMENTO"].str.strip()   
        df_agrupado = df.groupby('NR DOCUMENTO').agg({'VALOR': self._somar_valores , 'DATA': 'first', 'HISTORIO': 'first', 'LIBERAR DATA': 'first', 'DIFERENCA': 'first' , 'PROCESSAR': 'first', 'OBS': 'first'}).reset_index()
        return df_agrupado
    
    def _somar_valores(self, x, *args, **kwargs):
        total = Decimal(0)
        for i in x:
            if ',' in i:
                i = i.replace('.', '').replace(',', '.')
            total += Decimal(i)
        return str(total)


    def copiar_para_pasta_pro(self, PATH_PRO: str):
        list_arq_pro = os.listdir(PATH_PRO)
        for rel in os.scandir(self.PATH_DOWNLOAD_SOUL_MV):
            if os.path.basename(rel.path) not in list_arq_pro:
                shutil.copy(rel.path, PATH_PRO)
