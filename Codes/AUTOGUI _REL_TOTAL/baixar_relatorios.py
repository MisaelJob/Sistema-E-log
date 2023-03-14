import pyautogui
import time
from pathlib import Path

images_Dir = r"C:\Users\Misael\Documents\Estudos\AUTOGUI _REL_TOTAL\select_box\1440x900"
images_Path = Path(images_Dir)
images_list = list(images_Path.glob("*"))

def baixar_TotalExpress():
    pyautogui.click(pyautogui.locateCenterOnScreen(r"images\1440x900\desmarcar_BuscaPorLote.png"))
    for image_Dir in images_list:
        pyautogui.press("PageUp")
        meetBox = False
        while meetBox==False:
            try:
                x = pyautogui.locateOnScreen(str(image_Dir), confidence=0.8).left -10
                y = pyautogui.locateOnScreen(str(image_Dir), confidence=0.8)
                pyautogui.click(x,y)
                break
            except Exception:
                pyautogui.scroll(-600)

baixar_TotalExpress()