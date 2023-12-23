# -*- coding: utf-8 -*-
import os
import os.path
import shutil
import sys
import time
from datetime import datetime
from unicodedata import normalize

import pandas as pd
import pyautogui as p
import pyperclip
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver

from termcolor import colored


def display_error():
    """ 
    display error msg and return it as text
    'import sys' is required for this function
    """
    exctp, exc, exctb = sys.exc_info()
    fname = os.path.split(exctb.tb_frame.f_code.co_filename)[1]
    error_msg = f'\nerror: {exc}'
    error_msg += f'\nclass: {exctp}'
    error_msg += f'\nline: {exctb.tb_lineno}'
    error_msg += f'\nfunction: {exctb.tb_frame.f_code.co_name}'
    print(colored('\nTRACEBACK:', 'yellow'))
    print(colored(error_msg, 'red'))
    return '\nTRACEBACK:\n' + error_msg


def RemoverAcentos(txt):
    return normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')


def RemoveArquivo(pArqPro, raise_exception=False):
    try:
        if (os.path.isfile(pArqPro)):
            os.remove(pArqPro)
    except Exception as e:
        if raise_exception:
            raise OSError(f'Erro {e} em RemoveArquivo')
        else:
            print(f'Erro {e} em RemoveArquivo')


def ValidaTelaParaSeguir(pAtiv, pImagem, pTempo, pChamada):
    found_image = False
    max_wait = pTempo
    nErro = 0

    while (found_image == False) and nErro < max_wait:
        if ((p.locateOnScreen(pImagem, confidence=0.9) != None)):
           found_image = True
           print('espera ok ', pChamada)
        else:   
           found_image = False 
           nErro = nErro + 1
           time.sleep(0.6)
           print('espera nok ',pChamada, ' Imagem',pImagem, nErro)
    if (nErro >= max_wait):
        print('erro ValidaTelaParaSeguir', pChamada)
        raise ValueError(pChamada)


def validate_to_proceed(ima_path, secs, ima_name, confidence=0.9, region=None):
    found_image = False
    max_wait = secs
    err_num = 0
    while (found_image == False) and (err_num < max_wait):
        if ((p.locateOnScreen(ima_path, confidence=confidence, region=region) != None)):
            found_image = True
            print(f'found = {ima_name}')
        else:
            found_image = False
            found_image = False
            found_image = False
            err_num = err_num + 1
            time.sleep(0.6)
            print(f'searching = {ima_name} {ima_path} {err_num}')
    if (err_num >= max_wait):
        print(f'error = {ima_name}')
        raise ValueError(ima_name)


#Move o arquivo da pasta de origem onde o usuário libera para início do processamento.
#e na aplicação é feito o controle da linha executada e linha pendente
def MoveParaProcessamento(pAtiv, pArqPro):
    if (os.path.isfile(pArqPro)):
        dir_path= os.path.dirname(os.path.realpath(__file__))
        now = datetime.now()
        now_string = 'PRO_'+now.strftime("%Y%m%d%H%M%S%f")+'_'
        pArqDest = dir_path+'\\'+pAtiv+'\\pro\\'+now_string+os.path.basename(pArqPro)
        print(pArqDest)
        # os.rename(pArqPro, pArqDest)
        shutil.move(pArqPro, pArqDest)
    else: 
        print('Nada para Mover!')


def takeScreenshot(vaAtiv):
    myScreenshot = p.screenshot()
    now = datetime.now()
    ArquivoSaida = ".\\{0}\\imaout\\{1}.png".format(vaAtiv, now.strftime("%d%m%Y%H%M%S"))
    myScreenshot.save(ArquivoSaida)

#Elimina todas as aplicações do windows
def FinalizarAplicacao(pApl):
    comando = 'taskkill /f /im '+pApl
    os.system(comando)


def find_image(path, secs, ima_name, exc=True, clicks=0, click_wait=0.5, right_btn=False, gray=False, cfn=0.9, x_right=0, x_left=0, y_down=0, y_up=0, move=False, sendkeys='', press_enter=False, region_search=(0, 0, 1920, 1080)):
    """
    Funcao para tratar imagens:

    path : Path
        diretorio da imagem (orbrigatorio)
    secs : int
        tempo limite para buscar a imagem (orbrigatorio)
    ima_name : str
        nome da imagem no terminal (orbrigatorio)
    exc : bool
        determina se aplica o erro (raise error)
    clicks : int
        determina quantidade de clicks (1 = click; 2 = double click; 3 + = loop executanto clicks)
    click_wait : int
        determina intervalo entre clicks quando sao mais de 2 clicks
    right_btn : bool
        clica com o botao direito do mouse quando clicks for 1
    gray : bool
        determina se aplica grayscale na busca da imagem
    cfn : float
        determina o confidence para a imagem
    x_right : int
        move ou clica em imagem + quantiade x de pixels para a direita  
    x_left : int
        move ou clica em imagem + quantiade x de pixels para a esquerda  
    y_down : int
        move ou clica em imagem + quantiade x de pixels para baixo  
    y_up : int
        move ou clica em imagem + quantiade x de pixels para cima
    move : bool
        move mouse se parametro clicks for 0
    sendkeys : str
        caso seja passado algum text no parametro, o mesmo sera copiado e colado apos clicks
    press_enter : bool
        precionar tecla enter apos clicks e apos sendkeys  
    region_search : tuple
        determina a area da tela em forma de retangulo onde a funcao ira procurar por imagem
    """

    x_position = None
    y_position = None
    loop_count = 0
    found_image = False

    try:
        while found_image == False and loop_count < secs:
            var_ima = p.locateCenterOnScreen(
                image=path,
                confidence=cfn,
                grayscale=gray,
                region=region_search
            )

            if var_ima:
                print(f'\nFound "{ima_name}"') if loop_count > 0 else print(f'Found "{ima_name}"')

                x_position, y_position = var_ima
                found_image = True
                
                x_position += x_right
                x_position -= x_left
                y_position += y_down
                y_position -= y_up

                if clicks == 0:
                    if move:
                        p.moveTo(x=x_position, y=y_position)

                    return x_position, y_position

                elif clicks == 1:
                    p.click(x=x_position, y=y_position) if right_btn is False else p.click(x=x_position, y=y_position, button='right')
                    print(f'Clicked "{ima_name}"')

                    if sendkeys:
                        p.sleep(0.3)
                        pyperclip.copy(sendkeys)
                        p.hotkey('ctrl', 'v')

                    if press_enter:
                        p.sleep(0.3)
                        p.press('enter')

                    return x_position, y_position

                elif clicks == 2:
                    p.doubleClick(x=x_position, y=y_position)
                    print(f'Double-clicked "{ima_name}"')

                    if sendkeys:
                        p.sleep(0.3)
                        pyperclip.copy(sendkeys)
                        p.hotkey('ctrl', 'v')

                    if press_enter:
                        p.sleep(0.3)
                        p.press('enter')

                    return x_position, y_position

                elif clicks > 2:
                    for _ in range(clicks):
                        p.click(x=x_position, y=y_position)
                        print(f'Clicked "{ima_name}" - {c+1}', end='\r')
                        p.sleep(click_wait)

                    print()
                    return x_position, y_position

            found_image = False 
            loop_count += 1
            print(f'Searching "{ima_name}" = {loop_count} Seconds', end='\r')
            p.sleep(1)
            if secs == 1:
                print('\n')

    except:
        if loop_count > 0:
            print('\n')
        display_error()
        return

    if (loop_count >= secs
        and found_image is False
        and exc is True):

        raise ValueError(f'\nFailed to find image "{ima_name}"') 


def SalvarExcelComFiltros(p_ativ, msg_json, get_file_path, save_file_path, current_file_name, new_file_name, name_sheet, headers_list):
    try:
        # Ler arquivo
        now = datetime.now()
        current_file_name = f'{current_file_name}.xlsx' if '.xlsx' not in current_file_name else current_file_name
        lst_df = pd.read_excel(f'{get_file_path}{current_file_name}', sheet_name=name_sheet, dtype=str)
        lst_df.columns = headers_list
        # Gerar novo arquivo
        new_file_name = f'{new_file_name}.xlsx' if '.xlsx' not in new_file_name else new_file_name
        writer = pd.ExcelWriter(f'{save_file_path}{new_file_name}', engine='xlsxwriter')
        lst_df.to_excel(writer, index=False, sheet_name=name_sheet)
        work_book = writer.book
        work_sheet = writer.sheets[name_sheet]
        # Ajustar headers
        headerformat = work_book.add_format({ 'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
        for col_num, value in enumerate(lst_df.columns.values):
            work_sheet.write(0, col_num, value, headerformat)
        # Ajustar tamanho
        for column in lst_df:  
            column_width = max(lst_df[column].astype(str).map(len).max(), len(column))
            col_idx = lst_df.columns.get_loc(column)
            writer.sheets[name_sheet].set_column(col_idx, col_idx, column_width)
        work_sheet.autofilter(0, 0, lst_df.shape[0], lst_df.shape[1] -1) 
        writer.save()
        print(colored(f'{new_file_name}.txt salvo em {save_file_path}', 'green'))
    except:
        exc_type, error, exc_tb = sys.exc_info()
        error_msg = f'ERROR: {error}\nCLASS: {exc_type}\nFUNC: {sys._getframe().f_code.co_name}\nLINE: {exc_tb.tb_lineno}\n'
        print(colored(error_msg, 'red'))


def get_driver():
    """Retornar driver sem utilizar chromedriver.exe"""

    try:
        option = webdriver.ChromeOptions()
        option.add_argument("--disable-notifications")
        option.add_argument("--disable-infobars")
        option.add_argument("--start-maximized")
        option.add_argument("--disable-popup-blocking")
        # buster = os.path.join(os.path.dirname(__file__), "ATIV005", "doc", "Buster-Captcha-Solver-for-Humans.crx")
        # option.add_extension(buster)

        prefs = {
            # "download.default_directory": "download_directory",
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "download.prompt_for_download": True,
            "plugins.always_open_pdf_externally": True,
        }

        option.add_experimental_option("prefs", prefs)
        option.add_experimental_option("excludeSwitches", ["enable-logging"])

        # web_driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager(version='latest').install()), options=option)
        web_driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=option)

        web_driver.execute_script("window.alert = null;")
        web_driver.execute_script("Window.prototype.alert = null;")
        web_driver.execute_script("window.onbeforeunload = null;")
        web_driver.maximize_window()
        return web_driver
    except:
        display_error()
        return


# def get_driver(download_directory: str = ""):
#     try:
#         options = Options()
#         options.add_argument("--disable-notifications")
#         options.add_argument('--disable-infobars')
#         options.add_argument('--start-maximized')
#         options.add_argument('--disable-popup-blocking')   

#         prefs = {
#             # "download.default_directory": download_directory,
#             "credentials_enable_service": False,
#             "profile.password_manager_enabled": False,
#             "download.prompt_for_download": True,
#             "plugins.always_open_pdf_externally": True,
#             }

#         options.add_experimental_option("excludeSwitches", ["enable-logging"])
#         options.add_experimental_option("prefs", prefs)
#         executable_path = os.path.dirname(__file__)
#         executable_path = os.path.join(executable_path, "..", "DriverWeb", "chromedriver.exe")

#         driver = webdriver.Chrome(options=options,  executable_path=executable_path)

#         driver.execute_script("window.alert = null;")
#         driver.execute_script("Window.prototype.alert = null;")
#         driver.execute_script("window.onbeforeunload = null;")
#         driver.maximize_window()
#         return driver
#     except:
#         display_error()
#         return
    