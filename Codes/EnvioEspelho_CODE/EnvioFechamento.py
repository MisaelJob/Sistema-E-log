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
    if not CreatedTools.FindImage('inicioPagina_wtt.png'):
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
            if not CreatedTools.ProcurarContato_wtt(telefone_tabEnvio,nomeContato_tabEnvio):
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
            mensagemDeEnvio = f"Olá {nomeContato_tabEnvio}?\n" + mensagem
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
    if not CreatedTools.FindImage('inicioPagina_wtt.png'):
        print("----------> Pagina da Web não encontrada!")
        return
    #----------------------------------------------------
    
    tabelaDeEnvio_dir = f"{rt}\Relatorios\ENVIOS.xlsx"
    tabelaDeEnvio_dt = pd.read_excel(tabelaDeEnvio_dir)
    
    resultadoEnvios = tabelaDeEnvio_dt['STATUS'].value_counts()
    resultadoMensagem = f"🤖*MISATRON*2.1\n\nOlá {contato} o envido de *{arquivo}* foi finalizado, resultado:\n\nErros       {errorsCount}\n{resultadoEnvios}"
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
    pyautogui.hotkey("ctrl","v")
    if not CreatedTools.FindImage('abrirArquivo_wtt.png'):
        pyautogui.press('esc')
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
    #mensagemPronta = "Estamos chegando!!!\n\nVocê quer aumentar sua renda?\n\nO agileGo, o novo app de entregas que vai proporcionar mais oportunidades de ganho para você, veja o diferencial:\n\n- Receber pedidos de estabelecimentos de forma rápida e fácil;\n- Mais entregas por rota;\n- Maior ganho financeiro;\n- Ter mais flexibilidade para escolher seus horários e regiões de entrega.\n\nSe você está procurando uma oportunidade de ganho que te dê mais autonomia e renda, cadastre-se no agileGo.\n\nEm breve faça o seu cadastro e seja um dos nossos parceiros.\n\nConfira nosso site: https://www.agilego.com.br/\n\nNos siga nas redes:\n\nInstagram: https://abreai.link/3v8tl\nFacebook: https://abreai.link/nyo8k\nLinkedin: https://abreai.link/yliup"
    mensagemPronta = "Teste de envio"
    
    EnvioMensagem_wtt(mensagemPronta)
    #EnvioMensagem_wtt(mensagemPronta,"ESPELHO")
    
    #if executarEm_hora_minuto(18,15):      
    #    EnvioMensagem_wtt("Acesse este link para visualizar nosso catálogo no WhatsApp: https://wa.me/c/555191086827")    
    publicarResultados("Fechamento","Equipe Financeiro")
 
    
chamarFuncoesEnvio()



