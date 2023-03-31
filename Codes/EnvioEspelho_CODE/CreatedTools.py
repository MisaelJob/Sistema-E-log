from tkinter import Tk
import pyautogui
import pandas as pd
from pathlib import Path
import time
import win32com.client
import re

def DetectResolution():
    root = Tk()
    Xwidht = root.winfo_screenwidth()
    Yhight = root.winfo_screenheight()
    nowResolution = f"{Xwidht}x{Yhight}"
    return nowResolution
resolution = DetectResolution()


def RootFolder():
    thisArchive_dir = Path().absolute()
    caracters_dir = str(thisArchive_dir).find("Sistema-E-log") + 13
    nowRootFolder_dir = str(thisArchive_dir)[0:caracters_dir]
    return nowRootFolder_dir
rootFolder_dir = RootFolder()


def FindImage(imageName, action="click",attempts=4,imageFolder="Codes/EnvioEspelho_CODE/images"):
    returnValue = False
    for attempt in range(1, attempts+1, 1):
        time.sleep(1)
        try:
            pesquisa_wtt_pos = pyautogui.locateCenterOnScreen(f"{rootFolder_dir}/{imageFolder}/{resolution}/{imageName}", confidence=0.9)
        except: 
            print("---------->FindImage() Erro ao localizar arquivo da imagem!")
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
                print("---------->FindImage() Parâmetro action inválido")
                break   
        elif attempt == attempts:
            print("---------->FindImage() Erro ao localizar imagem na tela!")
            break
        #--------------------------------------------------------------------------------- 
    return returnValue


def funcionVBA(*args):
   excelApp = win32com.client.Dispatch("Excel.Application")
   excelWorkbook = excelApp.ActiveWorkbook
   result = excelWorkbook.Application.Run(*args)
   return result
   

def cttName(name):
    correctName = str(name)
    correctName = correctName.strip()
    correctName = correctName.upper()
    sobrenomes = correctName.split()
    novo_nome = ' '.join(sobrenomes[:-1])
    return novo_nome

def toTelephoneNum(text):
    tel = str(text)
    tel = re.findall(r'\d+', tel)
    tel = ''.join(tel)
    
    digTel = len(tel)
    if digTel >= 12:
        tellDD = tel[3:4]
    elif digTel >= 10:
        tellDD = tel[0:2]
    tel8 = tel[-8:]
    
    formateTel = '55'+tellDD+tel8
    return formateTel
    
    
    
    
    


