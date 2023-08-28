import pyautogui
import pyperclip
import pandas as pd
import time
from pathlib import Path
import CreatedTools

def EnvioMensagem_wtt(msg,arquivo="ESPELHO"):
    #----------------------------------------------------
    rt = CreatedTools.rootFolder_dir
    tabelaDeEnvio_dir = f"{rt}\Relatorios\ENVIO.xlsx"
    tabelaDeEnvio_dt = pd.read_excel(tabelaDeEnvio_dir)
    pyautogui.PAUSE = 1
    maxErrors = 10
    errorsCount = 0
    #----------------------------------------------------
    nomeEspelho_tabEnvio = ""
    nomeContato_tabEnvio = ""
    telefone_tabEnvio = ""
    base_tabelaDeEnvio = ""
    tipoPagamento_tabEnvio = ""
    status_tabEnvio = ""
    #----------------------------------------------------
    for index in range(0,len(tabelaDeEnvio_dt)):
        nomeEspelho_tabEnvio = tabelaDeEnvio_dt.loc[index,'ESPELHO']
        nomeContato_tabEnvio = tabelaDeEnvio_dt.loc[index,'ENVIO PARA']
        telefone_tabEnvio = tabelaDeEnvio_dt.loc[index,'TEL']
        base_tabEnvio = tabelaDeEnvio_dt.loc[index,'BASE']
        tipoPagamento_tabEnvio = tabelaDeEnvio_dt.loc[index,'TIPO']
        status_tabEnvio = tabelaDeEnvio_dt.loc[index,'STATUS']
        #----------------------------------------------------
        try:
            if status_tabEnvio.find("#") != -1:
                continue
        except:
            status_tabEnvio = ""
        #----------------------------------------------------  
        if not CreatedTools.ProcurarContato_wtt(nomeContato_tabEnvio):
            if not CreatedTools.ProcurarContato_wtt(telefone_tabEnvio):
                tabelaDeEnvio_dt.loc[index,'STATUS'] = "NÃO ENCONTRADO"
                tabelaDeEnvio_dt.to_excel(tabelaDeEnvio_dir,index=False)
                continue
        #----------------------------------------------------
        if arquivo == "ESPELHO":
            if not CreatedTools.funcionVBA('selecionarEspelho', nomeEspelho_tabEnvio, tipoPagamento_tabEnvio):
                #-----ERROR-------------------
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                continue
                #-----------------------------
            time.sleep(4)
            if not CreatedTools.FindImage('iconesChat_wtt.png'):
                #-----ERROR-------------------
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                continue
                #-----------------------------
            pyautogui.moveRel(100, 0, duration=0.25)
            pyautogui.click()
            pyautogui.hotkey('ctrl','v')
            #----------------------------------------------------
        else:
            if not CreatedTools.FindImage('chatAnexar_wtt.png'):
               #-----ERROR-------------------
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                continue
                #-----------------------------
            if not CreatedTools.FindImage('anexarArquivo_wtt.png'):
                #-----ERROR-------------------
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                continue
                #-----------------------------
            pyperclip.copy(arquivo)
            pyautogui.hotkey("ctrl","v")
            pyautogui.press('enter')
        #----------------------------------------------------
        mesagemDeEnvio = f"Olá {nomeContato_tabEnvio}.\n" + msg
        pyperclip.copy(mesagemDeEnvio)
        pyautogui.hotkey('ctrl','v')
        pyautogui.press('enter')
        #----------------------------------------------------
        tabelaDeEnvio_dt.loc[index,'STATUS'] = "#ENVIADO"
        tabelaDeEnvio_dt.to_excel(tabelaDeEnvio_dir,index=False)
    #----------------------------------------------------
    resultadoEnvios = tabelaDeEnvio_dt['STATUS'].value_counts()
    resultadoMensagem = f"🤖*MISATRON*\n\nEnvios do {arquivo} foi finalizado, resultado:\n\nErros       {errorsCount}\n{resultadoEnvios}"
    #----------------------------------------------------
    print(resultadoMensagem)
    CreatedTools.ProcurarContato_wtt("Equipe Financeiro")
    pyautogui.press('enter')
    pyperclip.copy(resultadoMensagem)
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    #----------------------------------------------------



EnvioMensagem_wtt("Segue ESPELHO da segunda quinzena de MARÇO","ESPELHO")
<<<<<<< Updated upstream
#teste de comit
=======






>>>>>>> Stashed changes















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
