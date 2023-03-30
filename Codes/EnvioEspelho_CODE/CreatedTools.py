from tkinter import Tk
import pyautogui
import pandas as pd
from pathlib import Path
import time

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


def FindImage(imageName, action="click",attempts=10,imageFolder="Codes/EnvioEspelho_CODE/images"):
    for attempt in range(1, attempts+1, 1):
        time.sleep(1)
        #---------------------------------------------------------------------------------
        try:
            pesquisa_wtt_pos = pyautogui.locateCenterOnScreen(f"{rootFolder_dir}/{imageFolder}/{resolution}/{imageName}")
        except: 
            print("---------->FindImage() Erro ao localizar arquivo da imagem!")
            exit()
        #---------------------------------------------------------------------------------
        if pesquisa_wtt_pos != None:
            if action == "click":
                pyautogui.click(pesquisa_wtt_pos)
            elif action == "moveTo":
                pyautogui.moveTo(pesquisa_wtt_pos)
            else:
                print("---------->FindImage() Parâmetro action inválido")
                exit()    
        elif attempt == attempts:
            print("---------->FindImage() Erro ao localizar imagem na tela!")
            exit()
        #--------------------------------------------------------------------------------- 
   
