import pyautogui
import pyperclip
import win32com.client
import pandas as pd
import time
from pathlib import Path
import CreatedTools
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By


#PERMITIR QUE CODIGO TENHA POSSIBILIDADE DE ESCOLHER ENVIO DE MSG ARQUIVO E ESPELHO
def DefinirInfos_EnvioEspelho(msg):
    #----------------------------------------------------
    tabelaDeEnvio_dir = r"C:\Users\Misael\Documents\Estudos\Sistema-E-log\Relatorios\ENVIOS 1Q0323.xlsx"
    tabelaDeEnvio_dt = pd.read_excel(tabelaDeEnvio_dir)
    #----------------------------------------------------
    nomeEspelho_tabEnvio = ""
    nomeContato_tabEnvio = ""
    telefone_tabEnvio = ""
    base_tabelaDeEnvio = ""
    tipoPagamento_tabEnvio = ""
    status_tabEnvio = ""
    #----------------------------------------------------
    nomeContato_wtt = ""
    #----------------------------------------------------
    for index in range(0,len(tabelaDeEnvio_dt)):
        nomeEspelho_tabEnvio = tabelaDeEnvio_dt.loc[index,'ESPELHO']
        nomeContato_tabEnvio = tabelaDeEnvio_dt.loc[index,'CONTATO']
        telefone_tabEnvio = tabelaDeEnvio_dt.loc[index,'TEL']
        base_tabEnvio = tabelaDeEnvio_dt.loc[index,'BASE']
        tipoPagamento_tabEnvio = tabelaDeEnvio_dt.loc[index,'TIPO']
        status_tabEnvio = tabelaDeEnvio_dt.loc[index,'STATUS']
        #----------------------------------------------------
        if status_tabEnvio.find("#") == -1:
            CreatedTools.FindImage('limparPesquisa_wtt.png')
            pyautogui.press('esc')
            if not CreatedTools.FindImage('pesquisaVazia_wtt.png'):
                if not CreatedTools.FindImage('pesquisaLimpa_wtt.png'):
                    print("---------->DefinirInfos_EnvioEspelho() Erro ao localizar barra de pesquisa do WhatsApp!")
                    continue
            #----------------------------------------------------
            pyautogui.write(nomeContato_tabEnvio)
            pyautogui.press('enter')
            CreatedTools.FindImage('limparPesquisa_wtt.png')
            #----------------------------------------------------
            if not CreatedTools.FindImage('opcoesPerfil_wtt.png'):
                continue
            if not CreatedTools.FindImage('dadosDoContato_wtt.png'):
                if not CreatedTools.FindImage('dadosDoContato_2_wtt.png'):
                    continue
            #----------------------------------------------------
            pyautogui.moveRel(-50, 230, duration=0.25)
            pyautogui.click(clicks=3)
            pyautogui.hotkey('ctrl','c')
            nomeContato_wtt = pyperclip.paste()
            #----------------------------------------------------
            if nomeContato_wtt == nomeContato_tabEnvio or CreatedTools.cttName(nomeContato_wtt) == nomeContato_tabEnvio or nomeContato_wtt == telefone_tabEnvio:
                if not CreatedTools.FindImage('fecharPerfil_wtt.png'):
                    continue
                #----------------------------------------------------
                if not CreatedTools.funcionVBA('selecionarEspelho', nomeEspelho_tabEnvio, tipoPagamento_tabEnvio):
                    continue
                time.sleep(4)
                #----------------------------------------------------
                if not CreatedTools.FindImage('iconesChat_wtt.png'):
                    continue
                pyautogui.moveRel(100, 0, duration=0.25)
                pyautogui.click()
                pyautogui.hotkey('ctrl','v')
                #----------------------------------------------------
                mesagemDeEnvio = f"Olá {nomeContato_tabEnvio}.\n" + msg
                pyautogui.write(mesagemDeEnvio)
                #pyautogui.press('enter')
                #----------------------------------------------------
                tabelaDeEnvio_dt.loc[index,'STATUS'] = "#ENVIADO"
                #----------------------------------------------------  
            else:
                continue
                #se não voltar ao inicio e tentar pesquisar telefone
                        #se não
                            #listar como não enviado
        else:
            tabelaDeEnvio_dt.loc[index,'STATUS'] = "NÃO ENVIADO"
        #dar respota sobre o envio no grupo de fechamento




#time.sleep(3)
#print(pyautogui.position())










def backup_envios(msg):
    #----------------------------------------------------
    tabelaDeEnvio_dir = r"C:\Users\Misael\Documents\Estudos\Sistema-E-log\Relatorios\ENVIOS 1Q0323.xlsx"
    tabelaDeEnvio_dt = pd.read_excel(tabelaDeEnvio_dir)
    #----------------------------------------------------
    nomeEspelho_tabEnvio = ""
    nomeContato_tabEnvio = ""
    telefone_tabEnvio = ""
    base_tabelaDeEnvio = ""
    tipoPagamento_tabEnvio = ""
    status_tabEnvio = ""
    #----------------------------------------------------
    nomeContato_wtt = ""
    #----------------------------------------------------
    for index in range(0,len(tabelaDeEnvio_dt)):
        nomeEspelho_tabEnvio = tabelaDeEnvio_dt.loc[index,'ESPELHO']
        nomeContato_tabEnvio = tabelaDeEnvio_dt.loc[index,'CONTATO']
        telefone_tabEnvio = tabelaDeEnvio_dt.loc[index,'TEL']
        base_tabEnvio = tabelaDeEnvio_dt.loc[index,'BASE']
        tipoPagamento_tabEnvio = tabelaDeEnvio_dt.loc[index,'TIPO']
        status_tabEnvio = tabelaDeEnvio_dt.loc[index,'STATUS']
        #----------------------------------------------------
        if status_tabEnvio.find("#") == -1:
            CreatedTools.FindImage('limparPesquisa_wtt.png')
            pyautogui.press('esc')
            if not CreatedTools.FindImage('pesquisaVazia_wtt.png'):
                if not CreatedTools.FindImage('pesquisaLimpa_wtt.png'):
                    print("---------->DefinirInfos_EnvioEspelho() Erro ao localizar barra de pesquisa do WhatsApp!")
                    continue
            #----------------------------------------------------
            pyautogui.write(nomeContato_tabEnvio)
            pyautogui.press('enter')
            CreatedTools.FindImage('limparPesquisa_wtt.png')
            #----------------------------------------------------
            if not CreatedTools.FindImage('opcoesPerfil_wtt.png'):
                continue
            if not CreatedTools.FindImage('dadosDoContato_wtt.png'):
                if not CreatedTools.FindImage('dadosDoContato_2_wtt.png'):
                    continue
            #----------------------------------------------------
            pyautogui.moveRel(-50, 230, duration=0.25)
            pyautogui.click(clicks=3)
            pyautogui.hotkey('ctrl','c')
            nomeContato_wtt = pyperclip.paste()
            #----------------------------------------------------
            if nomeContato_wtt == nomeContato_tabEnvio or CreatedTools.cttName(nomeContato_wtt) == nomeContato_tabEnvio or nomeContato_wtt == telefone_tabEnvio:
                if not CreatedTools.FindImage('fecharPerfil_wtt.png'):
                    continue
                #----------------------------------------------------
                if not CreatedTools.funcionVBA('selecionarEspelho', nomeEspelho_tabEnvio, tipoPagamento_tabEnvio):
                    continue
                time.sleep(4)
                #----------------------------------------------------
                if not CreatedTools.FindImage('iconesChat_wtt.png'):
                    continue
                pyautogui.moveRel(100, 0, duration=0.25)
                pyautogui.click()
                pyautogui.hotkey('ctrl','v')
                #----------------------------------------------------
                mesagemDeEnvio = f"Olá {nomeContato_tabEnvio}.\n" + msg
                pyautogui.write(mesagemDeEnvio)
                #pyautogui.press('enter')
                #----------------------------------------------------
                tabelaDeEnvio_dt.loc[index,'STATUS'] = "#ENVIADO"
                #----------------------------------------------------  
            else:
                continue
                #se não voltar ao inicio e tentar pesquisar telefone
                        #se não
                            #listar como não enviado
        else:
            tabelaDeEnvio_dt.loc[index,'STATUS'] = "NÃO ENVIADO"
        #dar respota sobre o envio no grupo de fechamento
