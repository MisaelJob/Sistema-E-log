import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from bs4 import BeautifulSoup
import pyautogui
import time
from pathlib import Path
from detectResolution import detectResolution
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pyperclip
from general_CODE import CreatedTools

pyautogui.PAUSE = 1
resolution = detectResolution()
dir =  Path(__file__).resolve().parent
#imagesDir = dir + "/image"
dirIMG = f'{dir}\\images\\{resolution}'

ics_link = "https://ics.totalexpress.com.br/index.php"
operacoesCAF_link = "https://ics.totalexpress.com.br/agentes/caf.php"
buscaPorLote_link = "https://ics.totalexpress.com.br/oper/relat_ultimostatus.php"


def loginICS(baseICS):
    pyautogui.hotkey('win','m')
    pyautogui.press('win')
    pyautogui.write('chrome')
    pyautogui.press('enter')
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage('contaGoogleLoginNavegador.png',0,0,"click",dirIMG):
        return    
    pyautogui.hotkey('win','up')
    pyautogui.write(ics_link)
    pyautogui.press('enter')
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage('selectUsers_ICS.png',0,0,"click",dirIMG):
        return
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage(f'loginName_ICS_{baseICS}.png',0,0,"click",dirIMG):
        return
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage('loginEnter_ICS.png',0,0,"click",dirIMG):
        return
    
  
loginICS("Gerencial")   
        
        
    


def guiaBuscaPorLote():
    guiaRelatorios_position = pyautogui.locateCenterOnScreen(f'images/{resolution}/guiaRelatoriosButton_ICS.png', confidence=0.8)
    if guiaRelatorios_position == None:
        print("----------> Error on to localize Relatorios tab!")
        exit
    else:
        pyautogui.moveTo(guiaRelatorios_position)
        #-------------------------------------------------------------------------------------------------------------------------------
        guiaBuscaPorLote_position = pyautogui.locateCenterOnScreen(f'images/{resolution}/buscaPorLoteGuiaButton_ICS.png', confidence=0.8)
        if guiaBuscaPorLote_position == None:
            print("----------> Error on to localize BuscaPorLote tab!")
            exit
        else:   
            pyautogui.click(guiaBuscaPorLote_position)


def selectCheckBox():
    #pyautogui.moveTo(x=1,y=1)
    #pyautogui.press('down', presses=4)
    #pyautogui.press('enter')
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
            
       
def extractCafs():
    browserTooBar_position = pyautogui.locateCenterOnScreen(f'{dir}/images/{resolution}/navegador_ToolBar.png', confidence=0.8)
    pyautogui.click(browserTooBar_position)
    #-------------------------------------------------------------------------------------------------------------------------------
    pyautogui.hotkey('alt','d')
    pyautogui.write(operacoesCAF_link)
    pyautogui.press('enter')
    #-------------------------------------------------------------------------------------------------------------------------------
    pyautogui.hotkey('ctrl','u')
    time.sleep(3)
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','c')
    pyautogui.hotkey('ctrl','w')
    html = pyperclip.paste()
    table_soup = BeautifulSoup(html,'html.parser').find_all('table',{'id':'tabela'})
    print(len(table_soup))

    
#selectCheckBox()

