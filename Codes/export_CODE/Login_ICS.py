import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from bs4 import BeautifulSoup
import requests
import pyautogui
import time
from pathlib import Path
from detectResolution import detectResolution
import pyperclip
from general_CODE import CreatedTools
import pandas as pd
import datetime
import openpyxl
import xlwings as xw


pyautogui.PAUSE = 0.5
resolution = detectResolution()
dir =  Path(__file__).resolve().parent
#imagesDir = dir + "/image"
dirIMG = f'{dir}\\images\\{resolution}'

ics_link = "https://ics.totalexpress.com.br/index.php"
operacoesCAF_link = "https://ics.totalexpress.com.br/agentes/caf.php"
buscaPorLote_link = "https://ics.totalexpress.com.br/oper/relat_ultimostatus.php"
cafsIcsDrive_dir = "G:\Meu Drive\DRIVE MISAEL\REPOSITORIO EASY\RELATORIOS\ICS_cafs.xlsm"

def LoginICS(baseICS):
    pyautogui.press('win')
    pyautogui.write('chrome')
    pyautogui.press('enter')
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage('contaGoogleLoginNavegador.png',0,0,"click",dirIMG):
        return    
    pyperclip.copy(ics_link)
    pyautogui.hotkey('win','up')
    pyautogui.hotkey('alt','d')
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage('selectUsers_ICS.png',0,0,"click",dirIMG,10):
        return
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage(f'loginName_ICS_{baseICS}.png',0,0,"click",dirIMG):
        if not CreatedTools.FindImage('rolagemListaLogins_ICS.png',0,0,"moveTo",dirIMG):
            return
        pyautogui.scroll(-200)
        if not CreatedTools.FindImage(f'loginName_ICS_{baseICS}.png',0,0,"click",dirIMG):
            return
    pyautogui.press('enter')
    #--------------------------------------------------------------------------------------------


def SelectCheckBox():
    #---------------------------------------------------------------------------------------------------------
    pyautogui.press('pgup')
    maxLoops = 6
    loops = 0
    desmarcarTudo_LOTE_position = pyautogui.locateCenterOnScreen(f'{dir}/images/{resolution}/desmarcar_BuscaPorLote.png', confidence=0.9)
    #---------------------------------------------------------------------------------------------------------
    if desmarcarTudo_LOTE_position == None:
        while desmarcarTudo_LOTE_position == None:
            loops = loops +1
            pyautogui.scroll(-200)
            if pyautogui.locateCenterOnScreen(f'{dir}/images/{resolution}/desmarcar_BuscaPorLote.png', confidence=0.9) != None:
                desmarcarTudo_LOTE_position = pyautogui.locateCenterOnScreen(f'{dir}/images/{resolution}/desmarcar_BuscaPorLote.png', confidence=0.9)
                pyautogui.click(desmarcarTudo_LOTE_position)
                break
            if loops >= maxLoops:
                print("----------> Error on localize desmarcar_BuscaPorLote!")
                exit
    else:
        pyautogui.click(desmarcarTudo_LOTE_position)
    #----------------------------------------------------------------------------------------------------------
    images_Dir = f"{dir}/select_box/{resolution}"
    images_Path = Path(images_Dir)
    images_list = list(images_Path.glob("*"))
    #----------------------------------------------------------------------------------------------------------
    pyautogui.press('pgup')
    pyautogui.scroll(-1200)
    upPage = False
    #---------------------------------------------------------------------------------------------------------
    for image_Dir in images_list:
        maxLoops = 10
        loops = 0       
        image_Rdir = f'{dir}/select_box/{resolution}/{image_Dir.name}'
        if pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9) == None:
            if  upPage == False:
                pyautogui.press('pgup')
                upPage = True
            while pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9) == None:
                loops = loops +1
                pyautogui.scroll(-400)
                if pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9) != None:
                    selectBox_position = pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9)
                    pyautogui.click(selectBox_position)
                    break
                if loops >= maxLoops:
                    print("----------> Error on localize checkBox!")
                    exit
        else:
            selectBox_position = pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9)
            pyautogui.click(selectBox_position)
            
       

    
def TabelarCafs(INICIO=1,PAGINAS=3, REPOSITORIO = cafsIcsDrive_dir ,FILTROS=['selecionarEncerrados.png','selecioanarAndamento.png']):    
    if os.path.isfile(REPOSITORIO):
        df_cafs = pd.read_excel(REPOSITORIO)
    else:
        df_cafs = pd.DataFrame()
    #--------------------------------------------------------------------------------------------------
    for filtro in FILTROS:
        if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"click",dirIMG,10):
            return
        pyautogui.press('pageup')
        #--------------------------------------------------------------------------------------------------
        if not CreatedTools.FindImage('operacaoesGuiaButton_ICS.png',0,0,"click",dirIMG):
            return
        #--------------------------------------------------------------------------------------------------
        if not CreatedTools.FindImage('subMenuCaf_ICS.png',0,0,"click",dirIMG):
            return
        if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"aguardar",dirIMG,10):
            return
        #--------------------------------------------------------------------------------------------------
        if not CreatedTools.FindImage('operacoesCAF\\filtroEncerradosAndamento.png',0,0,"click",dirIMG):
            return
        if not CreatedTools.FindImage(f'operacoesCAF\\{filtro}',0,0,"click",dirIMG):
            return
        #--------------------------------------------------------------------------------------------------
        if not CreatedTools.FindImage('operacoesCAF\\campoCAF.png',0,0,"click",dirIMG):
            return
        pyautogui.press('enter')
        if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"aguardar",dirIMG,10):
            return
        #--------------------------------------------------------------------------------------------------
        if filtro == 'selecioanarAndamento.png':
            INICIO = 1
            PAGINAS = 20
            
            
        for pagina in range(INICIO,PAGINAS):
            if pagina <= PAGINAS:
                pyautogui.press('pagedown')
                if not CreatedTools.FindImage('operacoesCAF\\irParaPagina.png',0,0,"click",dirIMG):
                    return
                pyperclip.copy(pagina)
                pyautogui.hotkey('ctrl','v')
                pyautogui.press('enter')
                if CreatedTools.FindImage('alertaOK_ICS.png',0,0,"click",dirIMG):
                    break 
                if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"aguardar",dirIMG,10):
                    return   
            #--------------------------------------------------------------------------------------------------
            if not CreatedTools.FindImage('operacoesCAF\\pagina50linhas.png',0,0,"click",dirIMG):
                return
            if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"aguardar",dirIMG,10):
                return
            #--------------------------------------------------------------------------------------------------
            pyautogui.hotkey('ctrl','u')
            if not CreatedTools.FindImage('operacoesCAF\\viewCarregada.png',0,0,"aguardar",dirIMG,10):
                return   
            #--------------------------------------------------------------------------------------------------
            pyautogui.hotkey('ctrl','a')
            pyautogui.hotkey('ctrl','c')
            pyautogui.hotkey('ctrl','w')
            html_caf = pyperclip.paste()
            soup = BeautifulSoup(html_caf, 'html.parser')
            table = soup.find('table', {'id': 'tabela'})
            if table is not None:
                df_temp = pd.read_html(str(table))[0]
                #--------------------------------------------------------------
                if filtro == 'selecioanarAndamento.png':
                    df_temp['STATUS'] = 'EmAndamento'
                elif filtro == 'selecionarEncerrados.png':
                    df_temp['STATUS'] = 'Encerado'
                #--------------------------------------------------------------              
                df_cafs = pd.concat([df_cafs, df_temp], ignore_index=True)
                df_cafs = df_cafs.drop_duplicates(subset='C.A.F.', keep='last')
                #--------------------------------------------------------------
                df_cafs.to_excel(REPOSITORIO, index=False)
            #--------------------------------------------------------------------------------------------------




def BaixarLote(DATA_INICIO='2001-01-01',DATA_FINAL='2001-01-01',QTD_CAFS=10,IMG_CAMPO="",REPOSITORIO = cafsIcsDrive_dir):
    if os.path.isfile(REPOSITORIO):
        df_cafs = pd.read_excel(REPOSITORIO)
    else:
        df_cafs = pd.DataFrame()
    #-------------------------------------------------------------------------
    CreatedTools.funcionVBA('CorrigirDatas',REPOSITORIO)
    df_cafs['Data de Abertura'] = pd.to_datetime(df_cafs['Data de Abertura'])
    #-------------------------------------------------------------------------
    tipo_de_dados = df_cafs[(df_cafs['Data de Abertura'] >= DATA_INICIO) & (df_cafs['Data de Abertura'] <= DATA_FINAL)]
    print(tipo_de_dados)
    
    
 
 
 
    

          
#SelectCheckBox()
#TabelarCafs(1,30)
BaixarLote()