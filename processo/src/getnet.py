import csv
import io
import os

import general as apr
import pandas as pd
import pyautogui as p
from selenium.common.exceptions import (JavascriptException, TimeoutException)
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
import shutil

from .util import *

class Gatnet():
    def __init__(self, driver: WebDriver, msg_json: dict, credendias, periodo_registro, PATH_DOWNLOAD_GETNET:str):
        self.driver = driver
        self.PATH_DOWNLOAD_GETNET = PATH_DOWNLOAD_GETNET
        self.page_inicial = credendias['site_gatnet']
        self.periodo_registro = periodo_registro
        self.nome_relatorio = os.path.join(PATH_DOWNLOAD_GETNET, f'relatorio_getnet_{self.periodo_registro}.csv')
        
    def verifica_popup_tela(self):
        XPATH_POPUP = (By.XPATH, '//ngb-modal-window//*[@aria-label="Fechar"]')
        XPATH_BTN_ACESSAR_CONTA = (By.XPATH, '//*[contains(text(), "ACESSAR MINHA CONTA")]')
        el = localizar_elemento(self.driver, XPATH_POPUP, 15)
        if el:
            el.click()
        el = localizar_elemento(self.driver, XPATH_BTN_ACESSAR_CONTA, 15)
        if el:
            el.click()
    def fazer_login(self, username, password):
        self.driver.get(self.page_inicial)
        XPATH_USUARIO = (By.XPATH, '//*[@id="username"]')
        XPATH_SENHA = (By.XPATH, '//*[@id="password"]')
        XPATH_BTN_ENTRAR = (By.XPATH, '//*[contains(text(), " ACESSAR ")]')
        
        self.verifica_popup_tela()
        
        input_usuario = localizar_elemento(self.driver, XPATH_USUARIO, 15, EC.element_to_be_clickable)
        input_senha = localizar_elemento(self.driver, XPATH_SENHA, 15, EC.element_to_be_clickable)
            
        input_usuario.send_keys(username)
        input_senha.send_keys(password)
            
        btn_entrar = localizar_elemento(self.driver, XPATH_BTN_ENTRAR, 10, EC.element_to_be_clickable)
        btn_entrar.click()

    def acessar_pagina_recebimento(self):
        XPATH_O_QUE_RECEBI = (By.XPATH, '//*[@id="body-index"]/app-root/main/app-home-financeiro-v2/div/div/div[2]/div/div/div[1]/app-card-o-que-recebi/div/div/div[2]/div[1]/a')
        a_link = localizar_elemento(self.driver, XPATH_O_QUE_RECEBI, 20, EC.element_to_be_clickable)
        a_link.click()

    def exportar_recebimentos(self, estabelecimento, PATH_IMA):
        XPATH_BTN_EXPORTAR = (By.CSS_SELECTOR, '#body-index > app-root > main > app-que-recebi > app-extrato-v2 > div > div:nth-child(1) > div.col-12.pt-3.container-filtros-topo > div > button.btn.btn-sm.btn-outline-secondary.btn-exportar.float-right.mr-2')
        XPATH_BTN_APRESENTACAO = (By.CSS_SELECTOR, '#visaoExtrato')
        XPATH_BTN_FORMATO = (By.CSS_SELECTOR, '#formatoExtrato')
        XPATH_BTN_HOJE = (By.CSS_SELECTOR, '#modal-getnet > div > div.modal-body.px-0.pb-4.pt-1 > div.col-12.text-secondary > div.row.mb-3.mt-3 > div > app-date-picker-line > div > button')
        XPATH_INPUT_DIA_DE = (By.CSS_SELECTOR, '#modal-getnet > div > div.modal-body.px-0.pb-4.pt-1 > div.col-12.text-secondary > div.row.mb-3.mt-3 > div > app-date-picker-line > div.dropdown-hover.text-primary.date-picker-line.cursor-pointer.dropdown-header.pb-0.show.dropdown > div > div:nth-child(8) > div:nth-child(1) > div > #dataDe')
        XPATH_INPUT_DIA_ATE = (By.CSS_SELECTOR, '#modal-getnet > div > div.modal-body.px-0.pb-4.pt-1 > div.col-12.text-secondary > div.row.mb-3.mt-3 > div > app-date-picker-line > div.dropdown-hover.text-primary.date-picker-line.cursor-pointer.dropdown-header.pb-0.show.dropdown > div > div:nth-child(8) > div:nth-child(2) > div > #dataAte')
        XPATH_BTN_APLICAR = (By.CSS_SELECTOR, '#modal-getnet > div > div.modal-body.px-0.pb-4.pt-1 > div.col-12.text-secondary > div.row.mb-3.mt-3 > div > app-date-picker-line > div.dropdown-hover.text-primary.date-picker-line.cursor-pointer.dropdown-header.pb-0.show.dropdown > div > div:nth-child(8) > div.col-12.text-center.mt-2 > #btnFiltrarPersonalizado')
        XPATH_BTN_EXPORTAR_ARQUIVO = (By.CSS_SELECTOR, '#btnExportar')
        
        btn_apresentacao, btn_formato, btn_hoje = self.obtem_botao_exportar(XPATH_BTN_EXPORTAR, XPATH_BTN_APRESENTACAO, XPATH_BTN_FORMATO, XPATH_BTN_HOJE)
        if btn_apresentacao is not None and btn_formato is not None and btn_hoje is not None:
            btn_apresentacao, is_transformer = self.transformar_elemento_select(btn_apresentacao)
            if is_transformer:
                btn_apresentacao.select_by_value('A')
            
            btn_formato, is_transformer = self.transformar_elemento_select(btn_formato)
            if is_transformer:
                btn_formato.select_by_value("C")
                
            btn_hoje.click()
            
            input_dia_de = localizar_elemento(self.driver, XPATH_INPUT_DIA_DE, 15, EC.element_to_be_clickable)
            input_dia_ate = localizar_elemento(self.driver, XPATH_INPUT_DIA_ATE, 15, EC.element_to_be_clickable)
            btn_aplicar = localizar_elemento(self.driver, XPATH_BTN_APLICAR, 15, EC.element_to_be_clickable)
            input_dia_de.send_keys(self.periodo_registro)
            input_dia_ate.send_keys(self.periodo_registro)
            btn_aplicar.click()
            
            self.selecionar_estabelecimento(estabelecimento)
            
            btn_exportar_arquivo = localizar_elemento(self.driver, XPATH_BTN_EXPORTAR_ARQUIVO, 15, EC.element_to_be_clickable)
            btn_exportar_arquivo.click()
            
            if self.verifica_se_efetuou_download(self.download_csv_recebimentos(PATH_IMA)):
                return self.ler_relatorio_normalizar()
            else:
                return pd.DataFrame()
        else:
            return pd.DataFrame()

    def obtem_botao_exportar(self, XPATH_BTN_EXPORTAR, XPATH_BTN_APRESENTACAO, XPATH_BTN_FORMATO, XPATH_BTN_HOJE, tempo=15):
        btn_apresentacao = None
        btn_formato = None
        btn_hoje = None
        while True and tempo > -1:
            try:
                btn_exportar = localizar_elemento(self.driver, XPATH_BTN_EXPORTAR, 20, EC.element_to_be_clickable)
                btn_exportar.click()
                
                btn_apresentacao = localizar_elemento(self.driver, XPATH_BTN_APRESENTACAO, 15, EC.element_to_be_clickable)
                btn_formato = localizar_elemento(self.driver, XPATH_BTN_FORMATO, 15, EC.element_to_be_clickable)
                btn_hoje = localizar_elemento(self.driver, XPATH_BTN_HOJE, 15, EC.element_to_be_clickable)
                break
            except TimeoutException:
                self.driver.refresh()
            tempo -= 1
            p.sleep(1)
        return btn_apresentacao,btn_formato,btn_hoje

    def download_csv_recebimentos(self, PATH_IMA, tempo = 10):
        XPATH_BTN_ATUALIZAR = (By.CSS_SELECTOR, '#body-index > app-root > main > app-downloads > div > div > div.col-6.text-right > button')
        IMAGEM_TELA_SALVAR_COMO = os.path.join(PATH_IMA, 'salvar_como.png')
                
        btn_atualizar = localizar_elemento(self.driver, XPATH_BTN_ATUALIZAR, 15, EC.element_to_be_clickable)
        
        js_code = "return document.querySelector('#body-index > app-root > main > app-downloads > div > div > div.col-12.mb-3 > table > tbody > tr').childNodes[6].querySelector('a')"
        efetuado_download = False
        while True and tempo > -1:
            try:
                # Vefica se o elemento do js esta com o atributo aria-disabled
                elemento:WebElement = self.driver.execute_script(js_code)
                value_atributes = elemento.get_attribute('aria-disabled')
                if value_atributes == "true" and not efetuado_download:
                    btn_atualizar.click()
                else:
                    if not efetuado_download:
                        elemento.click()
                    try:
                        apr.validate_to_proceed(IMAGEM_TELA_SALVAR_COMO, 20, "SALVAR COMO", confidence=0.8)
                        p.write(self.nome_relatorio)
                        p.press('enter')
                        efetuado_download = True
                        break
                    except p.ImageNotFoundException:
                        pass
                    except:
                        apr.display_error()
            except JavascriptException:
                apr.display_error()
            except:
                apr.display_error()
            tempo -= 1
            p.sleep(1)
        return efetuado_download

    def verifica_se_efetuou_download(self, is_download=False, tempo_procura=15):
        if is_download:
            while True and tempo_procura > -1:
                p.press('esc')
                if os.path.exists(self.nome_relatorio):
                    return True
                tempo_procura -= 1
                p.sleep(1)
        return False

    def transformar_elemento_select(self, elemento):
        if isinstance(elemento, WebElement):
            elemento = Select(elemento)
            return elemento, True
        return elemento, False

    def selecionar_estabelecimento(self, estabelecimento):
        # localizar o botao e clicar
        XPATH_BOTAO_SEL_ECS = (By.CSS_SELECTOR, '#modal-getnet > div > div.modal-body.px-0.pb-4.pt-1 > div.col-12.text-secondary > div.row.mb-3.mt-3 > div.col-6.ng-star-inserted > app-dropdown-matriz > button')
        XPATH_BOTAO_SEL = (By.CSS_SELECTOR, '#body-index > ngb-modal-window > div > div > div > div:nth-child(1) > div > div > div > div.col-md-2.col-sm-12.h-100.text-center.my-auto > div.row.mb-4 > div.col-12.my-2 > button')
        XPATH_BOTAO_OK = (By.CSS_SELECTOR, '#body-index > ngb-modal-window > div > div > div > div.col-12.text-center.my-3 > button.btn.btn-sm.btn-primary.button-matriz-submit.mx-2')
        XPATH_INPUT_ESTABELECIMENTO = (By.CSS_SELECTOR, '#body-index > ngb-modal-window > div > div > div > div:nth-child(1) > div > div > div > div:nth-child(1) > div > div > div > div > div:nth-child(2) > div > input')
        
        botao_sel_ecs = localizar_elemento(self.driver, XPATH_BOTAO_SEL_ECS, 20, EC.element_to_be_clickable)
        botao_sel_ecs.click()
        
        input_estabelecimento = localizar_elemento(self.driver, XPATH_INPUT_ESTABELECIMENTO, 20, EC.element_to_be_clickable)
        botao_selecionar = localizar_elemento(self.driver, XPATH_BOTAO_SEL, 20, EC.element_to_be_clickable)
        
        input_estabelecimento.send_keys(estabelecimento)
        estabelecimentos = self.obter_estabelecimentos()
        encontrou_elemento = False
        for item in estabelecimentos:
            if item[0].text == estabelecimento:
                item[0].click()
                encontrou_elemento = True
                break
        if encontrou_elemento:
            botao_selecionar.click()
            botao_ok = localizar_elemento(self.driver, XPATH_BOTAO_OK, 15, EC.element_to_be_clickable)
            botao_ok.click()

    def obter_estabelecimentos(self) -> list[list[WebElement, WebElement]]:
        js_code = """
            let necessario = [];
            let novaLista = [];
            let divs = document.querySelectorAll("#body-index > ngb-modal-window > div > div > div > div:nth-child(1) > div > div > div > div:nth-child(4) > div.card.bg-salt-grey.h-100 > div > div")

            divs.forEach((div, index) => {
                let children = Array.from(div.children);
                necessario.push(children.slice(0, 2));
                // Se este é o último elemento visível, rola para ele
                if (index === divs.length - 1) {
                    div.scrollIntoView();
                }
            });
            necessario.forEach((item) => {
                let codigo = item[0]
                let banco = item[1]
                novaLista.push([codigo, banco]);
            });
            return novaLista
        """

        lista_dados = []
        while True:
            total = len(lista_dados)
            resp = self.driver.execute_script(js_code)
            if resp is not None:
                for item in resp:
                    if item not in lista_dados:
                        lista_dados.append(item)
            else:
                break
            if total == len(lista_dados):
                break
            p.sleep(0.6)
        return lista_dados

    def ler_relatorio_normalizar(self):
        with open(self.nome_relatorio, 'r') as file:
            content = file.read()
            
            content = apr.RemoverAcentos(content)
        f = io.StringIO(content)
        
        # Le o arquivo
        reader = csv.reader(f, delimiter=';')
        rows = []
        for row in reader:
            rows.append(row)
        f = io.StringIO('\n'.join([';'.join(rows) for rows in rows]))
        df = pd.read_csv(f, delimiter=';', header=0, dtype={'AUTORIZACAO': str})
        df = df[df["AUTORIZACAO"].notna()]
        df.to_csv(self.nome_relatorio, sep=';', index=False)
        return df

    def copiar_para_pasta_pro(self, PATH_PRO: str):
        list_arq_pro = os.listdir(PATH_PRO)
        for rel in os.scandir(self.PATH_DOWNLOAD_GETNET):
            if os.path.basename(rel.path) not in list_arq_pro:
                shutil.copy(rel.path, PATH_PRO)
