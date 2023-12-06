import pyautogui
import pyperclip
import pandas as pd
import time
from pathlib import Path
import CreatedTools
import datetime

rt = CreatedTools.rootFolder_dir
maxErrors = 10
errorsCount = 0


def EnvioMensagem_wtt(mensagem,arquivo=""):
    #----------------------------------------------------
    rt = CreatedTools.rootFolder_dir
    tabelaDeEnvio_dir = f"{rt}\Relatorios\ENVIOS.xlsx"
    tabelaDeEnvio_dt = pd.read_excel(tabelaDeEnvio_dir)
    pyautogui.PAUSE = 0.6
    maxErrors = 10
    errorsCount = 0
    #----------------------/------------------------------
    nomeEspelho_tabEnvio = ""
    nomeContato_tabEnvio = ""
    telefone_tabEnvio = ""
    base_tabelaDeEnvio = ""
    tipoPagamento_tabEnvio = ""
    status_tabEnvio = ""
    #----------------------------------------------------
    if not CreatedTools.FindImage(imageName='inicioPagina_wtt.png',action="moveTo"):
        print("----------> Pagina da Web não encontrada!")
        return
    #----------------------------------------------------
    for index in range(0,len(tabelaDeEnvio_dt)):
        nomeEspelho_tabEnvio = tabelaDeEnvio_dt.loc[index,'ENTREGADOR EASY']
        nomeContato_tabEnvio = tabelaDeEnvio_dt.loc[index,'NOME REPRESENTANTE REAL']
        telefone_tabEnvio = tabelaDeEnvio_dt.loc[index,'TELEFONE']
        base_tabEnvio = tabelaDeEnvio_dt.loc[index,'BASE']
        tipoPagamento_tabEnvio = tabelaDeEnvio_dt.loc[index,'MODALIDADE']
        variavel_tabEnvio = tabelaDeEnvio_dt.loc[index,'VARIAVEL']
        somaDescontos_tabEnvio = tabelaDeEnvio_dt.loc[index,'Soma de Descontos']
        bonusNF_tabEnvio = tabelaDeEnvio_dt.loc[index,'BÔNUS NF']
        totalEspelho_tabEnvio = tabelaDeEnvio_dt.loc[index,'valor final REAL']
        chave_tabEnvio = tabelaDeEnvio_dt.loc[index,'#CHAVE']
        status_tabEnvio = tabelaDeEnvio_dt.loc[index,'STATUS']
        comentarios = True
        #----------------------------------------------------
        try:
            if status_tabEnvio.find("#") != -1:
                continue
        except:
            status_tabEnvio = ""
        #----------------------------------------------------
        if not CreatedTools.ProcurarContato_wtt(nomeContato_tabEnvio,telefone_tabEnvio):  
            tabelaDeEnvio_dt.loc[index,'STATUS'] = "NÃO ENCONTRADO"
            tabelaDeEnvio_dt.to_excel(tabelaDeEnvio_dir,index=False)
            if comentarios:
                print("----------> Contato não encontrado!")
            continue
        #----------------------------------------------------
        if arquivo == "":
            pass
        #******************************************************************************************************************
        elif CreatedTools.ArchiveType(arquivo) == "espelho":
            try:
                CreatedTools.funcionVBA('selecionarEspelho', nomeEspelho_tabEnvio, tipoPagamento_tabEnvio)
            except:
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                if comentarios:
                    print("----------> Erro ao selecionar espelho!")
                continue 
                #-----------------------------
            else:
                time.sleep(4)
                if not CreatedTools.FindImage('iconesChat_wtt.png'):
                    errorsCount = errorsCount + 1
                    if errorsCount >= maxErrors:
                        break
                    if comentarios:
                        print("----------> Botões do chat não encontrados!")
                    continue
                    #-----------------------------
                else:
                    pyautogui.moveRel(100, 0, duration=0.25)
                    pyautogui.click()
                    pyautogui.hotkey('ctrl','v')
                    pyautogui.press('enter')
            #------------------------------------------------------------------------------------------------
            if somaDescontos_tabEnvio > 0:
                try:    
                    CreatedTools.funcionVBA('FiltroOutrosRelatorios')
                except:
                    errorsCount = errorsCount + 1
                    if errorsCount >= maxErrors:
                        break
                    if comentarios:
                        print("----------> Erro no filtro de relatorios extras/extravios!")
                    continue 
                    #-----------------------------
                else:
                    time.sleep(4)
                    if not CreatedTools.FindImage('iconesChat_wtt.png'):
                        errorsCount = errorsCount + 1
                        if errorsCount >= maxErrors:
                            break
                        if comentarios:
                            print("----------> Botões do chat não encontrados!")
                        continue
                        #-----------------------------
                    else:
                        pyautogui.moveRel(100, 0, duration=0.25)
                        pyautogui.click()
                        pyautogui.hotkey('ctrl','v')
                        #------------------------------------------------
                        pyperclip.copy("*Relatório de descontos aplicados:*\n Questionamento quanto à aplicação de extravios, multas ou divergência de descontos você deve contatar o GRIS, através do Supervisor *RAFAEL* no fone: wa.me/555197242536.")
                        pyautogui.hotkey('ctrl','v')
                        time.sleep(0.5)
                        #-------------------------------------------------
                        pyautogui.press('enter')
                if totalEspelho_tabEnvio < 0:
                    pyperclip.copy(f"⚠️Atenção seus pagamentos podem estar bloqueados por razão de seu espelho estar com valor negativo de {totalEspelho_tabEnvio}!")
                    pyautogui.hotkey('ctrl','v')
                    time.sleep(0.5)
                    #-------------------------------------------------
                    pyautogui.press('enter')          
        #************************************************************************************************************
        elif CreatedTools.ArchiveType(arquivo) == "arquivo":
            if not CreatedTools.FindImage('chatAnexar_wtt.png'):
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                if comentarios:
                    print("----------> Botão anexar do chat não encontrados!")
                continue
                #-----------------------------
            if not CreatedTools.FindImage('anexarArquivo_wtt.png'):
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                if comentarios:
                    print("----------> Botões de anexar arquivo não encontrados!")
                continue
                #-----------------------------
            pyperclip.copy(arquivo)
            pyautogui.hotkey("ctrl","v")
            if not CreatedTools.FindImage('abrirArquivo_wtt.png'):
                pyautogui.press('esc')
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                if comentarios:
                    print("----------> Botão abrir arquivo não encontrado!")
                continue
        #*************************************************************************************************************
        elif CreatedTools.ArchiveType(arquivo) == "image":
            if not CreatedTools.FindImage('chatAnexar_wtt.png'):
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                if comentarios:
                    print("----------> Botão anexar do chat não encontrado!")
                continue
                #-----------------------------
            if not CreatedTools.FindImage('anexarImagem_wtt.png'):
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                if comentarios:
                    print("----------> Botão anexar imagem não encontrado!")
                continue
                #-----------------------------     
            pyperclip.copy(arquivo)
            pyautogui.hotkey("ctrl","v")
            if not CreatedTools.FindImage('abrirArquivo_wtt.png'):
                pyautogui.press('esc')
                errorsCount = errorsCount + 1
                if errorsCount >= maxErrors:
                    break
                if comentarios:
                    print("----------> Botão abrir arquivo não encontrado!")
                continue
        #----------------------------------------------------
        if mensagem == "TESTE":
            print(nomeEspelho_tabEnvio)
            pass
        else:
            CreatedTools.FindImage('iconesChat_wtt.png',100)
            #----------------------------------------------------
            #mensagemDeEnvio = f"Olá {nomeContato_tabEnvio}?\n" + mensagem
            mensagemDeEnvio = mensagem
            pyperclip.copy(mensagemDeEnvio)
            pyautogui.hotkey('ctrl','v')
            pyautogui.press('enter')
            pyautogui.press('esc')
            pyautogui.press('esc')
            #----------------------------------------------------
            tabelaDeEnvio_dt.loc[index,'STATUS'] = "#ENVIADO"
            tabelaDeEnvio_dt.to_excel(tabelaDeEnvio_dir,index=False)
    #----------------------------------------------------



def publicarResultados(arquivo,contato):
    if not CreatedTools.FindImage(imageName='inicioPagina_wtt.png',action='moveTo'):
        print("----------> Pagina da Web não encontrada!")
        return
    #----------------------------------------------------
    
    tabelaDeEnvio_dir = f"{rt}\Relatorios\ENVIOS.xlsx"
    tabelaDeEnvio_dt = pd.read_excel(tabelaDeEnvio_dir)
    
    resultadoEnvios = tabelaDeEnvio_dt['STATUS'].value_counts()
    resultadoMensagem = f"🤖 *MISATRON 2.2* \n\nOlá {contato} o envido de *{arquivo}* foi finalizado, resultado:\n\nErros       {errorsCount}\n{resultadoEnvios} \n\n Tomem cuidado, estou sendo atualizado! 😎"
    #-----------------------------------------------------------------
    if CreatedTools.ProcurarContato_wtt(contato):
        pyperclip.copy(resultadoMensagem)
        pyautogui.press('enter') 
        pyautogui.hotkey('ctrl','v')
        time.sleep(0.5)
        pyautogui.press('enter') 
        #-----------------------------------------------------------------
        if not CreatedTools.FindImage('chatAnexar_wtt.png'):
            return
        if not CreatedTools.FindImage('anexarArquivo_wtt.png'):
            return
        pyperclip.copy(tabelaDeEnvio_dir)
        pyautogui.hotkey("ctrl","a")
        pyautogui.hotkey("ctrl","v")
        #if not CreatedTools.FindImage('abrirArquivo_wtt.png'):
        pyautogui.press('enter')
        pyautogui.press('enter') 
        
        


def executarEm_hora_minuto(hora, minuto):
    horarioDeExecutar = datetime.time(hora,minuto)
    
    while True:
        horaAtual = datetime.datetime.now().hour
        minutoAtual = datetime.datetime.now().minute
        horarioAtual = datetime.time(horaAtual,minutoAtual)
        
        if horarioDeExecutar == horarioAtual:
            #print(horarioAtual)
            return True
        time.sleep(20)
    #-----------------------------------------------------------------


def chamarFuncoesEnvio():
    arquivo = "ESPELHO"
    #mensagemPronta = "🚨🚀Descubra a revolução nas entregas! 🚚 O app agileGo já está disponível na Play Store! 📲✨\n\nFaça o download agora e aproveite:\n\n🌟 Valores de bônus semanais que vão te surpreender!\n📣 Bônus incríveis por indicação de amigos.\n💰 Valores agressivos por entrega (70%), garantindo o seu bolso cheio!\n🎉 Brindes para os primeiros a se cadastrar e carregar, e muito mais.\n\nNão perca tempo, junte-se à equipe agileGo e ganhe mais a cada entrega.\n\nBaixe agora em https://play.google.com/store/apps/details?id=br.com.agilego e comece a lucrar! 💵💼\n\nSaiba mais: www.agilego.com.br"

    mensagemPronta = "Segue *primeira* quinzena de *Novembro.*\n\nAlgumas CAFs não estavam aparecendo após manutenção do sistema da TOTAL EXPRESS, favor considerar apenas o ultimo envio."
    
    #if executarEm_hora_minuto(18,15):      
    
    #EnvioMensagem_wtt(mensagemPronta,"ESPELHO")
    EnvioMensagem_wtt(mensagemPronta,arquivo)
    publicarResultados(arquivo,"Equipe Financeiro")
 
    
chamarFuncoesEnvio()



