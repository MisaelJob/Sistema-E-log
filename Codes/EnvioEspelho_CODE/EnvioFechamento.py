import pyautogui
import pandas as pd
from pathlib import Path
import CreatedTools
#---------------------------------------------------------------------------------
tabelaDeEnvio_dir = r"C:\Users\Misael\Documents\Estudos\Sistema-E-log\Relatorios\ENVIOS 1Q0323.xlsx"
tabelaDeEnvio_dt = pd.read_excel(tabelaDeEnvio_dir)
resolution = CreatedTools.DetectResolution()
#---------------------------------------------------------------------------------
nomeEspelho_tabelaDeEnvio = ""
nomeContato_tabelaDeEnvio = ""
telefone_tabelaDeEnvio = ""
base_tabelaDeEnvio = ""
tipoPagamento_tabelaDeEnvio = ""
#---------------------------------------------------------------------------------
nomeContato_wtt = ""
status_tabelaDeEnvio = "SEM ANALISE"
#---------------------------------------------------------------------------------
pesquisa_wtt_pos = ""
limparPesquisa_wtt_pos = ""
perfil_wtt_pos = ""
nomePerfil_wtt_pos = ""
chat_wtt_pos = ""
#---------------------------------------------------------------------------------
def DefinirInfos_EnvioEspelho():
    for index in range(0,len(tabelaDeEnvio_dt)):
        nomeEspelho_tabelaDeEnvio = tabelaDeEnvio_dt.loc[index,'ESPELHO']
        nomeContato_tabelaDeEnvio = tabelaDeEnvio_dt.loc[index,'CONTATO']
        telefone_tabelaDeEnvio = tabelaDeEnvio_dt.loc[index,'TEL']
        base_tabelaDeEnvio = tabelaDeEnvio_dt.loc[index,'BASE']
        tipoPagamento_tabelaDeEnvio = tabelaDeEnvio_dt.loc[index,'TIPO']

        #se nome for valido para envio
            #limpar barra de pesquisa
            #selecionar pesquisa
            #colar o nome contato
            #limpar barra de pesquisa
            #selecionar conversa
            #abrir perfil
            #copiar o nome dop perfil
            #verificar se o nome de perfil é o mesmo do nome de contato
                #verificar se o nome de perfil é o telefone
                #se não voltar ao inicio e tentar pesquisar telefone
                    #se não
                        #listar como não enviado
            #selecionar conversa
            #colar espelho
            #colar mensagem
            #enviar
            #listar como enviado
        #se não
            #listar como não enviado






