from tkinter import Tk
import pyautogui
import pandas as pd
from pathlib import Path
import time
import win32com.client
import re
import pyperclip
import requests
import os



def RootFolder():
    thisArchive_dir = os.path.abspath(os.path.dirname(__file__))
    caracters_dir = thisArchive_dir.find("Sistema-E-log") + 13
    nowRootFolder_dir = thisArchive_dir[0:caracters_dir]
    return nowRootFolder_dir
rootFolder_dir = RootFolder()
print(f'Diretório raiz definido: {rootFolder_dir}.')


def DetectResolution():
    root = Tk()
    Xwidht = root.winfo_screenwidth()
    Yhight = root.winfo_screenheight()
    nowResolution = f"{Xwidht}x{Yhight}"
    #print(f'----------> Resulução de tela {nowResolution},localizada.')
    return nowResolution
resolution = DetectResolution()




def FindImage(imageName,posX = 0,posY = 0,action="click",imageFolder=f"{rootFolder_dir}\\Codes\\EnvioEspelho_CODE\\images\\{resolution}",aguardar = 4):
    #-------------------------------------------------------
    particaoDoTexto = imageName.split(".")
    dirIMG = particaoDoTexto[0] + "*." + particaoDoTexto[1]
    #-------------------------------------------------------
    path_IMG = Path(imageFolder)
    list_variacoesIMG = list(path_IMG.glob(dirIMG))
    #-------------------------------------------------------
    returnValue = False
    pesquisa_wtt_posX = None
    pesquisa_wtt_posY = None
    #-------------------------------------------------------
    for tentativa in range(0,aguardar):
        time.sleep(1)
        if pesquisa_wtt_posX is not None:
            break
        
        for variacao in list_variacoesIMG:
            dirVariacao = variacao.as_posix()
            print(tentativa,dirVariacao)
            #----------------------------------     
            try:
                pesquisa_wtt_posX, pesquisa_wtt_posY = pyautogui.locateCenterOnScreen(dirVariacao, confidence=0.9)
            except: 
               continue
            #------------------------------------------------
            if pesquisa_wtt_posX != None:
                pesquisa_wtt_posX = pesquisa_wtt_posX + posX
                pesquisa_wtt_posY = pesquisa_wtt_posY + posY
                break
        
    #---------------------------------------------------------------------------------
    if pesquisa_wtt_posX != None:
        if action == "click":
            pyautogui.click(pesquisa_wtt_posX,pesquisa_wtt_posY)
            returnValue = True     
        elif action == "moveTo":
            pyautogui.moveTo(pesquisa_wtt_posX,pesquisa_wtt_posY)
            returnValue = True
        elif action == "position":
            returnValue = [pesquisa_wtt_posX, pesquisa_wtt_posY]
        elif action == "aguardar":
            returnValue = True
    #---------------------------------------------------------------------------------
    if not returnValue:
        print(f'----------> Imagem {imageName}, não encontrada.')
    #---------------------------------------------------------------------------------
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
    tel = str(text)
    tellDD = ""
    tel8 =  ""
    telRe = re.findall(r'[0-9]*', tel)
    telRe = ''.join(telRe)
    #----------------------------------------------------
    tel8 = telRe[-8:]
    
    formateTel = f"{tel8}"
    try:
        formateTel = int(formateTel)
    except:
        formateTel = 0
    #print(f'----------> Valor: {text}, tratado para: {formateTel}')
    return formateTel
    
    
def ProcurarContato_wtt(pesquisa,telefone=0):
    #----------------------------------------------------
    FindImage('fecharPerfil_wtt.png')
    pyautogui.press('esc')
    #----------------------------------------------------
    if not valido(pesquisa):
        return False
    #----------------------------------------------------
    pyperclip.copy(pesquisa)
    pyautogui.hotkey("ctrl","alt","/")
    pyautogui.hotkey("ctrl","a")
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    #----------------------------------------------------
    pyautogui.hotkey("ctrl","alt","/")
    pyautogui.hotkey("ctrl","a")
    pyautogui.press('backspace')
    #----------------------------------------------------
    
    
    
    if not FindImage('opcoesPerfil_wtt.png',20):
        pass
    #----------------------------------------------------
    if not FindImage('dadosDoContato_wtt.png'):
        if not FindImage('dadosDoContato_2_wtt.png'):
            if not FindImage('dadosDoContato_3_wtt.png'):
                return False
    #----------------------------------------------------
    
    tentativaNomeContato = range(0,1,2)
    for tentantivaContador in tentativaNomeContato:
        print(tentantivaContador)
        #-----------------------------------------------------------------------------------
        if FindImage('contaOficialPerfil.png'):
            pyautogui.scroll(-2000)
            time.sleep(1)
            #-----------------------------------------------------------------------------------
            if tentantivaContador == 0:
                pyautogui.moveRel(-50, 100, duration=0.5)
            #-----------------------------------------------------------------------------------
            elif tentantivaContador == 1:
                pyautogui.moveRel(-50, 225, duration=0.5)
            else:
                return False 
        #-----------------------------------------------------------------------------------
        else:
            if tentantivaContador == 0:
                pyautogui.moveRel(-50, 233, duration=0.5)
            elif tentantivaContador == 1:
                pyautogui.moveRel(-50, 260, duration=0.5)
            elif tentantivaContador == 1:
                pyautogui.moveRel(-50, 282, duration=0.5)
        #-----------------------------------------------------------------------------------
        pyperclip.copy("")
        pesqContato_wtt = ""
        #-----------------------------------------------------------------------------------
        pyautogui.click(clicks=3)
        pyautogui.hotkey('ctrl','c')
        pesqContato_wtt = pyperclip.paste()
        #-----------------------------------------------------------------------------------
        if not FindImage('fecharPerfil_wtt.png'):
                pyautogui.press("esc") 
        #-----------------------------------------------------------------------------------
        

        if pesqContato_wtt == pesquisa:
            return True
        elif cttName(pesqContato_wtt) == pesquisa:
            return True
        else:
            #print(f"nome:{pesquisa} wtt:{pesqContato_wtt}")
            pass

    
        if valido(toTelephoneNum(pesqContato_wtt)) and valido(toTelephoneNum(telefone)):
            if toTelephoneNum(pesqContato_wtt) == toTelephoneNum(telefone):
                return True
        elif valido(toTelephoneNum(pesqContato_wtt)) and valido(toTelephoneNum(pesquisa)):   
            if toTelephoneNum(pesqContato_wtt) == toTelephoneNum(pesquisa):
                return True 
        else:
            #print(f"nome:{toTelephoneNum(pesquisa)} tel:{toTelephoneNum(telefone)} wtt:{toTelephoneNum(pesqContato_wtt)}")
            pass


    return False
    #-----------------------------------------------------------------------------------
   
    
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
#MousePosition_X_Y()


def valido(variavel):
    if type(variavel) == int:
        if variavel > 0:
            return True
    elif type(variavel) == str:
        if len(variavel) > 0:
            return True
    #print(type(variavel))
    return False

