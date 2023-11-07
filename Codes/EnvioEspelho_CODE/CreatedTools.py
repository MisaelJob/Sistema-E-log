from tkinter import Tk
import pyautogui
import pandas as pd
from pathlib import Path
import time
import win32com.client
import re
import pyperclip
import requests

def DetectResolution():
    root = Tk()
    Xwidht = root.winfo_screenwidth()
    Yhight = root.winfo_screenheight()
    nowResolution = f"{Xwidht}x{Yhight}"
    #print(f'----------> Resulução de tela {nowResolution},localizada.')
    return nowResolution
resolution = DetectResolution()


def RootFolder():
    thisArchive_dir = Path().absolute()
    caracters_dir = str(thisArchive_dir).find("Sistema-E-log") + 13
    nowRootFolder_dir = str(thisArchive_dir)[0:caracters_dir]
    #print(f'----------> Pasta raiz definida: {nowRootFolder_dir}.')
    return nowRootFolder_dir
rootFolder_dir = RootFolder()


def FindImage(imageName,posX = 0,posY = 0,action="click",attempts=4,imageFolder="Codes/EnvioEspelho_CODE/images"):
    returnValue = False
    for attempt in range(1, attempts+1, 1):
        time.sleep(1)
        try:
            pesquisa_wtt_posX, pesquisa_wtt_posY = pyautogui.locateCenterOnScreen(f"{rootFolder_dir}/{imageFolder}/{resolution}/{imageName}", confidence=0.9)
        except: 
            continue
        #---------------------------------------------------------------------------------
        pesquisa_wtt_posX = pesquisa_wtt_posX + posX
        pesquisa_wtt_posY = pesquisa_wtt_posY + posY
        
        if pesquisa_wtt_posX != None:
            if action == "click":
                pyautogui.click(pesquisa_wtt_posX,pesquisa_wtt_posY)
                returnValue = True
                break
            elif action == "moveTo":
                pyautogui.moveTo(pesquisa_wtt_posX,pesquisa_wtt_posY)
                returnValue = True
                break
            else:
                break   
        elif attempt == attempts:
            break
        #--------------------------------------------------------------------------------- 
    if returnValue:
        #print(f'----------> Imagem {imageName}, encontrada com sucesso.')
        pass
    else:
        print(f'----------> Imagem {imageName}, não encontrada.')
        pass
    return returnValue


def funcionVBA(*args):
   excelApp = win32com.client.Dispatch("Excel.Application")
   excelWorkbook = excelApp.ActiveWorkbook
   resultado = excelWorkbook.Application.Run(*args)
   #print(f'----------> Função VBA {args[0]} chamada.')
   return resultado
  

def cttName(name):
    correctName = str(name)
    correctName = correctName.strip()
    correctName = correctName.upper()
    sobrenomes = correctName.split()
    novo_nome = ' '.join(sobrenomes[:-1])
    novo_nome = novo_nome.strip()
    #print(f'----------> Valor: {name}, tratado para: {novo_nome}')
    return novo_nome


def toTelephoneNum(text):
    numeros = re.findall(r'\d', text)
    
    if len(numeros) >= 12:
        formateTel = int(''.join(numeros[:4] + numeros[-8:]))
    elif len(numeros) >= 10:
        formateTel = int('55' + ''.join(numeros[:2] + numeros[-8:]))
    else:
        formateTel = 0
    return formateTel

    

def ProcurarContato_wtt(pesquisa,telefone=0):
    returnContatoEncontrado = False    
    FindImage('fecharPerfil_wtt.png')
    pyautogui.press('esc')
    #----------------------------------------------------
    if not valido(pesquisa):
        return False
    pyperclip.copy(pesquisa)
    #----------------------------------------------------
    metodosDeBusca = ["pesquisaWtt","linkDireto"]
    for metodoAtual in metodosDeBusca:
        if returnContatoEncontrado == True:
                break
        #------------------------------------------------
        if metodoAtual == "pesquisaWtt":
            #pyautogui.hotkey("ctrl","alt","/")
            pyautogui.click(824,169)
            pyautogui.hotkey("ctrl","a")
            pyautogui.hotkey('ctrl','v')
            pyautogui.press('enter')
            #--------------------------------------------
            #pyautogui.hotkey("ctrl","alt","/")
            pyautogui.click(824,169)
            pyautogui.hotkey("ctrl","a")
            pyautogui.press('backspace')
            #---------------------------------------------------------------------------------------------
        elif metodoAtual == "linkDireto":
            pyautogui.hotkey('alt','d')
            pyperclip.copy("https://web.whatsapp.com/send/?phone=" + str(toTelephoneNum(telefone)))
            pyautogui.hotkey('ctrl','v')
            pyautogui.press('enter')
            time.sleep(2)
            if not FindImage('inicioPagina_wtt.png',attempts=200):
                continue
            time.sleep(2)
        #---------------------------------------------------------------------------------------------
        if not FindImage('opcoesPerfil_wtt.png',20):
            if not FindImage('opcoesPerfil2_wtt.png'):
                pass
        #----------------------------------------------------
        if not FindImage('dadosDoContato_wtt.png'):
            if not FindImage('dadosDoContato_2_wtt.png'):
                if not FindImage('dadosDoContato_3_wtt.png'):
                    continue
        #----------------------------------------------------
        tentativaValidadarNomeContato = range(0,3,1)
        for tentantivaContador in tentativaValidadarNomeContato:
            if returnContatoEncontrado == True:
                break
            #----------------------------------------------------
            if tentantivaContador == 0:
                pyautogui.moveRel(-50, 213, duration=0.5)
            elif tentantivaContador == 1:
                pyautogui.moveRel(0, 20, duration=0.5)
            elif tentantivaContador == 2:
                pyautogui.moveRel(0, 30, duration=0.5)
            #----------------------------------------------------
            pyperclip.copy("")
            pesqContato_wtt = ""
            #-----------------------------------
            pyautogui.click(clicks=3)
            pyautogui.hotkey('ctrl','c')
            pesqContato_wtt = pyperclip.paste()
            #-----------------------------------
            if pesqContato_wtt == pesquisa:
                returnContatoEncontrado = True    
            elif cttName(pesqContato_wtt) == pesquisa:
                returnContatoEncontrado = True 
            #-------------------------------------------------------------------------------
            if valido(toTelephoneNum(pesqContato_wtt)) and valido(toTelephoneNum(telefone)):
                if toTelephoneNum(pesqContato_wtt) == toTelephoneNum(telefone):
                    returnContatoEncontrado = True                                     
            elif valido(toTelephoneNum(pesqContato_wtt)) and valido(toTelephoneNum(pesquisa)):   
                if toTelephoneNum(pesqContato_wtt) == toTelephoneNum(pesquisa):
                    returnContatoEncontrado = True                 
    #----------------------------------------------------------------------------------------
    if not FindImage('fecharPerfil_wtt.png'):
        pyautogui.press("esc") 
    #----------------------------------------
    return returnContatoEncontrado
    #----------------------------------------
    
    
 
    
    
def ArchiveType(arquivo):
    arquivo = str(arquivo)
    arquivo = arquivo.lower()
    resposta = False
    if  arquivo.find(".png") != -1:
        resposta = "image"
    elif arquivo.find(".jpeg") != -1:
        resposta = "image"
    elif arquivo.find(".jpg") != -1:
        resposta = "image"
    elif arquivo.find(".gif") != -1:
        resposta = "image" 
    elif arquivo.find(".tiff") != -1:
        resposta = "image"
    elif arquivo.find(".svg") != -1:
        resposta = "image"
    elif arquivo.find(".webp") != -1:
        resposta = "image"
    elif arquivo.find(".") != -1:
        resposta = "archive"
    elif arquivo.find("espelho") != -1 or arquivo.find("fechamento") != -1:
        resposta = "espelho"
    #----------------------------------------------------
    if resposta:
        #print(f'----------> Arquivo {arquivo}, encontrada com sucesso.')
        pass
    else:
        #print(f'----------> Arquivo {arquivo}, não encontrada.')
        pass
    #----------------------------------------------------   
    return resposta
        

def MousePosition_X_Y():
    time.sleep(2)
    print(pyautogui.position())
MousePosition_X_Y()




def valido(variavel):
    if type(variavel) == int:
        if variavel > 0:
            return True
    elif type(variavel) == str:
        if len(variavel) > 0:
            return True
    #print(type(variavel))
    return False
    
