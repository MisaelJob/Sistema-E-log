from tkinter import Tk
import pyautogui
import pandas as pd
from pathlib import Path
import time
import win32com.client
import re
import pyperclip

def DetectResolution():
    root = Tk()
    Xwidht = root.winfo_screenwidth()
    Yhight = root.winfo_screenheight()
    nowResolution = f"{Xwidht}x{Yhight}"
    print(f'----------> Resulução de tela {nowResolution},localizada.')
    return nowResolution
resolution = DetectResolution()


def RootFolder():
    thisArchive_dir = Path().absolute()
    caracters_dir = str(thisArchive_dir).find("Sistema-E-log") + 13
    nowRootFolder_dir = str(thisArchive_dir)[0:caracters_dir]
    print(f'----------> Pasta raiz definida: {nowRootFolder_dir}.')
    return nowRootFolder_dir
rootFolder_dir = RootFolder()


def FindImage(imageName, action="click",attempts=4,imageFolder="Codes/EnvioEspelho_CODE/images"):
    returnValue = False
    for attempt in range(1, attempts+1, 1):
        time.sleep(1)
        try:
            pesquisa_wtt_pos = pyautogui.locateCenterOnScreen(f"{rootFolder_dir}/{imageFolder}/{resolution}/{imageName}", confidence=0.9)
        except: 
            break
        #---------------------------------------------------------------------------------
        if pesquisa_wtt_pos != None:
            if action == "click":
                pyautogui.click(pesquisa_wtt_pos)
                returnValue = True
                break
            elif action == "moveTo":
                pyautogui.moveTo(pesquisa_wtt_pos)
                returnValue = True
                break
            else:
                break   
        elif attempt == attempts:
            break
        #--------------------------------------------------------------------------------- 
    print(f'----------> Imagem {imageName}, localizada.')
    return returnValue


def funcionVBA(*args):
   excelApp = win32com.client.Dispatch("Excel.Application")
   excelWorkbook = excelApp.ActiveWorkbook
   result = excelWorkbook.Application.Run(*args)
   print(f'----------> Função VBA {args[0]} chamada.')
   return result
  

def cttName(name):
    correctName = str(name)
    correctName = correctName.strip()
    correctName = correctName.upper()
    sobrenomes = correctName.split()
    novo_nome = ' '.join(sobrenomes[:-1])
    print(f'----------> Valor: {name}, tratado para: {novo_nome}')
    return novo_nome


def toTelephoneNum(text):
    tel = text
    tellDD = ""
    tel8 =  ""
    telRe = re.findall(r'[0-9]*', tel)
    telRe = ''.join(telRe)
    #----------------------------------------------------
    digTel = len(telRe)
    if digTel >= 12:
        tellDD = str(telRe[2])+str(telRe[3])
    elif digTel >= 10:
        tellDD = str(telRe[0])+str(telRe[1])
    tel8 = telRe[-8:]
    
    formateTel = f"55{tellDD}{tel8}"
    print(f'----------> Valor: {text}, tratado para: {formateTel}')
    return formateTel
    
    
def ProcurarContato_wtt(pesquisa):
    FindImage('fecharPerfil_wtt.png')
    FindImage('limparPesquisa_wtt.png')
    pyautogui.press('esc')
    #----------------------------------------------------
    if not FindImage('pesquisaVazia_wtt.png'):
        if not FindImage('pesquisaLimpa_wtt.png'):
            return False
    #----------------------------------------------------
    pyautogui.write(pesquisa)
    pyautogui.press('enter')
    FindImage('limparPesquisa_wtt.png')
    #----------------------------------------------------
    if not FindImage('opcoesPerfil_wtt.png'):
        return False
    if not FindImage('dadosDoContato_wtt.png'):
        if not FindImage('dadosDoContato_2_wtt.png'):
            if not FindImage('dadosDoContato_3_wtt.png'):
                return False
    #----------------------------------------------------
    pyautogui.moveRel(-50, 230, duration=0.25)
    pyautogui.click(clicks=3)
    pyautogui.hotkey('ctrl','c')
    pesqContato_wtt = pyperclip.paste()
    #----------------------------------------------------
    if not FindImage('fecharPerfil_wtt.png'):
        return False
    #----------------------------------------------------
    if pesqContato_wtt == pesquisa:
        contatoLocalizado = True
    elif cttName(pesqContato_wtt) == pesquisa:
        contatoLocalizado = True
    elif toTelephoneNum(pesqContato_wtt) == toTelephoneNum(pesquisa):    
        contatoLocalizado = True
    else:
        contatoLocalizado = False
    #----------------------------------------------------
    print(f'----------> Localização do cotato, {pesquisa}: {contatoLocalizado}')
    return contatoLocalizado
    #----------------------------------------------------

  
    
    
    


