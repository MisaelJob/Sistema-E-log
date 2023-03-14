from bs4 import BeautifulSoup
import pyautogui
import time
from pathlib import Path
from tkinter import Tk

def teste():
    loginEnterICS_position = pyautogui.locateCenterOnScreen('images/1920X1080/loginEnter_ICS.png', confidence=0.8)
    #pyautogui.moveTo(loginEnterICS_position)

    novaLista = []
    novaLista.append(loginEnterICS_position)

    #for i in novaLista:
    #    if i == None:
    #        print()
     #       exit()
     #   print("Noneeeeeeeeeeeee!")

    root = Tk()
    testeReso = root.winfo_screenwidth()
    print(testeReso)

teste()