import sys
import os
sys.path.append('.')
from bs4 import BeautifulSoup
import pyautogui
import time
from pathlib import Path
import pyperclip
from Codes import CreatedTools
import pandas as pd
import openpyxl
import xlwings as xw
import datetime
import zipfile
import rarfile
import shutil
import psutil

pyautogui.PAUSE = 0.5
resolution = CreatedTools.DetectResolution()
dir =  Path(__file__).resolve().parent
dirIMG = f'{dir}\\images\\{resolution}'
download_dir = Path.home() / "Downloads"


ics_link = "https://ics.totalexpress.com.br/index.php"
operacoesCAF_link = "https://ics.totalexpress.com.br/agentes/caf.php"
buscaPorLote_link = "https://ics.totalexpress.com.br/oper/relat_ultimostatus.php"
cafsIcsDrive_dir = "G:\\Meu Drive\\DRIVE MISAEL\\REPOSITORIO EASY\\RELATORIOS\\ICS_cafs.xlsm"
totalCafsQuinzena_dir = "G:\Meu Drive\DRIVE MISAEL\REPOSITORIO EASY\RELATORIOS\TotalCafsQuinzena.xlsx"



def LoginICS(baseICS):
    pyautogui.press('win')
    pyautogui.write('chrome')
    pyautogui.press('enter')
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage('contaGoogleLoginNavegador.png',0,0,"click",3,dirIMG):
        return    
    pyperclip.copy(ics_link)
    pyautogui.hotkey('win','up')
    pyautogui.hotkey('alt','d')
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage('selectUsers_ICS.png',0,0,"click",3,dirIMG):
        return
    #--------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage(f'loginName_ICS_{baseICS}.png',0,0,"click",3,dirIMG):
        if not CreatedTools.FindImage('rolagemListaLogins_ICS.png',0,0,"moveTo",3,dirIMG):
            return
        pyautogui.scroll(-200)
        if not CreatedTools.FindImage(f'loginName_ICS_{baseICS}.png',0,0,"click",3,dirIMG):
            return
    pyautogui.press('enter')
    #--------------------------------------------------------------------------------------------



def SelectCheckBox():
    #---------------------------------------------------------------------------------------------------------
    pyautogui.press('pgup')
    maxLoops = 6
    loops = 0
    desmarcarTudo_LOTE_position = pyautogui.locateCenterOnScreen(f'{dir}/images/{resolution}/desmarcar_BuscaPorLote.png', confidence=0.9)
    #---------------------------------------------------------------------------------------------------------
    if desmarcarTudo_LOTE_position == None:
        while desmarcarTudo_LOTE_position == None:
            loops = loops +1
            pyautogui.scroll(-200)
            if pyautogui.locateCenterOnScreen(f'{dir}/images/{resolution}/desmarcar_BuscaPorLote.png', confidence=0.9) != None:
                desmarcarTudo_LOTE_position = pyautogui.locateCenterOnScreen(f'{dir}/images/{resolution}/desmarcar_BuscaPorLote.png', confidence=0.9)
                pyautogui.click(desmarcarTudo_LOTE_position)
                break
            if loops >= maxLoops:
                print("----------> Error on localize desmarcar_BuscaPorLote!")
                exit
    else:
        pyautogui.click(desmarcarTudo_LOTE_position)
    #----------------------------------------------------------------------------------------------------------
    images_Dir = f"{dir}/select_box/{resolution}"
    images_Path = Path(images_Dir)
    images_list = list(images_Path.glob("*"))
    #----------------------------------------------------------------------------------------------------------
    pyautogui.press('pgup')
    pyautogui.scroll(-1200)
    upPage = False
    #---------------------------------------------------------------------------------------------------------
    for image_Dir in images_list:
        maxLoops = 10
        loops = 0       
        image_Rdir = f'{dir}/select_box/{resolution}/{image_Dir.name}'
        if pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9) == None:
            if  upPage == False:
                pyautogui.press('pgup')
                upPage = True
            while pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9) == None:
                loops = loops +1
                pyautogui.scroll(-400)
                if pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9) != None:
                    selectBox_position = pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9)
                    pyautogui.click(selectBox_position)
                    break
                if loops >= maxLoops:
                    print("----------> Error on localize checkBox!")
                    exit
        else:
            selectBox_position = pyautogui.locateCenterOnScreen(image_Rdir, confidence=0.9)
            pyautogui.click(selectBox_position)        
       

    
def ExportarOperacaoCafs(INICIO=1,PAGINAS=3, REPOSITORIO = cafsIcsDrive_dir ,FECHAR_NAVEGADOR=True,FILTROS=['selecionarEncerrados.png','selecioanarAndamento.png']):    
    if os.path.isfile(REPOSITORIO):
        df_cafs = pd.read_excel(REPOSITORIO, engine='openpyxl')
    else:
        df_cafs = pd.DataFrame()
    #----------------------------------------------------------------------------------------------------------------------------------------
    for filtro in FILTROS:
        if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"click",20,dirIMG):
            return
        pyautogui.press('pageup')
        #-----------------------------------------------------------------------------------
        if not CreatedTools.FindImage('operacaoesGuiaButton_ICS.png',0,0,"click",3,dirIMG):
            return
        #-----------------------------------------------------------------------------------
        if not CreatedTools.FindImage('subMenuCaf_ICS.png',0,0,"click",3,dirIMG):
            return
        if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"click",20,dirIMG):
            return
        #-------------------------------------------------------------------------------------------------
        if not CreatedTools.FindImage('operacoesCAF\\filtroEncerradosAndamento.png',50,0,"click",3,dirIMG):
            return
        if not CreatedTools.FindImage(f'operacoesCAF\\{filtro}',50,0,"click",3,dirIMG):
            return
        #--------------------------------------------------------------------------------
        if not CreatedTools.FindImage('operacoesCAF\\campoCAF.png',0,0,"click",3,dirIMG):
            return
        pyautogui.press('enter')
        if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"click",20,dirIMG):
            return
        #-----------------------------------------------------------------------------
        if filtro == 'selecioanarAndamento.png':
            INICIO = 1
            PAGINAS = 20
        #----------------------------------------------------------------------------------------------------------------------------------------
        for pagina in range(INICIO,PAGINAS):
            if pagina <= PAGINAS:
                pyautogui.press('pagedown',presses=3)
                if not CreatedTools.FindImage('operacoesCAF\\irParaPagina.png',-70,0,"click",3,dirIMG):
                    return
                #--------------------------------------------------------------------------------------
                if pagina > 1:
                    pyperclip.copy(pagina)
                    pyautogui.hotkey('ctrl','v')
                    pyautogui.press('enter')
                #------------------------------------------------------------------
                if CreatedTools.FindImage('alertaOK_ICS.png',0,0,"click",3,dirIMG):
                    break 
                if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"click",20,dirIMG):
                    return   
            #--------------------------------------------------------------------------------------
            if not CreatedTools.FindImage('operacoesCAF\\pagina50linhas.png',0,0,"click",3,dirIMG):
                return
            if not CreatedTools.FindImage('guiaCarregada_ICS.png',0,0,"click",20,dirIMG):
                return
            #-------------------------------------------------------------------------------------
            pyautogui.hotkey('ctrl','u')
            if not CreatedTools.FindImage('operacoesCAF\\viewCarregada.png',0,0,"click",20,dirIMG):
                return   
            #--------------------------------------------------------------------------------------
            pyautogui.hotkey('ctrl','a')
            pyautogui.hotkey('ctrl','c')
            pyautogui.hotkey('ctrl','w')
            html_caf = pyperclip.paste()
            soup = BeautifulSoup(html_caf, 'html.parser')
            table = soup.find('table', {'id': 'tabela'})
            if table is not None:
                df_temp = pd.read_html(str(table))[0]
                #---------------------------------------
                if filtro == 'selecioanarAndamento.png':
                    df_temp['STATUS'] = 'EmAndamento'
                elif filtro == 'selecionarEncerrados.png':
                    df_temp['STATUS'] = 'Encerado'
                #-------------------------------------------------------------- 
                df_cafs = pd.concat([df_cafs, df_temp], ignore_index=True)
                df_cafs = df_cafs.drop_duplicates(subset='C.A.F.', keep='last')
    #----------------------------------------------------------------------------------------------------------------------------------------
    CreatedTools.SavarDataFrameEmExcel(df_cafs,REPOSITORIO)
    if FECHAR_NAVEGADOR:
        pyautogui.hotkey('alt','f4')
    #----------------------------------------------------------------------------------------------------------------------------------------

          
            

def BaixarLote(CAFS="",LOTE="loteCaf_BuscaPorLote.png"):
    if not CreatedTools.FindImage(imageName='guiaCarregada_ICS.png',action="click",imageFolder=dirIMG):
        return
    #-----------------------------------------------------------------------------------------------------------
    pyperclip.copy("https://ics.totalexpress.com.br/oper/relat_ultimostatus.php")
    pyautogui.hotkey('alt','d')
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    #-------------------------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage(imageName='guiaCarregada_ICS.png',action="click",imageFolder=dirIMG,attempts=20):
        return
    #-------------------------------------------------------------------------------------------------------------
    if not CreatedTools.FindImage(imageName=LOTE,action="click",imageFolder=dirIMG,posY=60):
        return
    #-------------------------------------------------------------------------------------------------------------
    pyperclip.copy(CAFS)
    pyautogui.hotkey('ctrl','v')
    SelectCheckBox()
    #-------------------------------------------------------------------------------------------------------------
    pyautogui.press('pagedown',presses=4)
    if not CreatedTools.FindImage(imageName='prosseguir_BuscaPorLote.png',action="click",imageFolder=dirIMG):
        return
    #-------------------------------------------------------------------------------------------------------------
    pyautogui.press('pageup',presses=4)
    if not CreatedTools.FindImage(imageName='dowloadButton_BuscaPorLote.png',action="click",imageFolder=dirIMG,attempts=120):
        return
    #-------------------------------------------------------------------------------------------------------------



def ListarArquivosRecentes(TEMPO_MODIFICACAO="00:00:00",DIAS_MODIFICACAO=0,DIRETORIO=download_dir):
    arquivosRecentes_list = []
    data_atual = datetime.datetime.now()
    horasMod, minutosMod, segundosMod = map(float, TEMPO_MODIFICACAO.split(':'))
    data_limiteModificacao = data_atual - datetime.timedelta(days=DIAS_MODIFICACAO,hours=int(horasMod),minutes=int(minutosMod),seconds=int(segundosMod))
    #------------------------------------------------------------------------------------------------------------------------------------   
    for nomeArquivo in os.listdir(DIRETORIO):
        caminho_completo = os.path.join(DIRETORIO, nomeArquivo)    
        if os.path.isfile(caminho_completo):
            data_modificacao = datetime.datetime.fromtimestamp(os.path.getmtime(caminho_completo))
            #---------------------------------------------------------------------------------------
            if data_modificacao >= data_limiteModificacao:
                arquivosRecentes_list.append(str(caminho_completo))
    #-------------------------------------------------------------------------------------------------
    return arquivosRecentes_list       
      
      
              
def DescompactarArquivos(DIRETORIOS_LISTA=[],LISTAR_NAO_DESCOMPACTADOS=False, DELATAR_COMPACTADADOS=True):
    arquivosDescompactados_list = []
    arquivosCompactados_list = []
    arquivosParaDeletar_list = []
    #----------------------------------------------------
    arquivosCompactados_list.extend(DIRETORIOS_LISTA)
    #----------------------------------------------------
    for caminho_arquivo in DIRETORIOS_LISTA:
        pasta_destino = os.path.dirname(caminho_arquivo)
        #------------------------------------------------
        if caminho_arquivo.lower().endswith('.zip'):
            with zipfile.ZipFile(caminho_arquivo, 'r') as zip_ref:
                zip_ref.extractall(pasta_destino)
                arquivosInternos = zip_ref.namelist()
                #------------------------------------------------------
                for i_dir in arquivosInternos:
                    arquivoDescompactado_name = os.path.basename(i_dir)
                    arquivoDescompactado_dir = os.path.join(pasta_destino, arquivoDescompactado_name)
                    #--------------------------------------------------------------------------------
                    arquivosDescompactados_list.append(arquivoDescompactado_dir)
                    #-----------------------------------------------------------
                    arquivosCompactados_list.remove(caminho_arquivo)
                    arquivosParaDeletar_list.append(caminho_arquivo)
        #----------------------------------------------------------------------------------------------
        elif caminho_arquivo.lower().endswith('.rar'):
            with rarfile.RarFile(caminho_arquivo, 'r') as rar_ref:
                rar_ref.extractall(pasta_destino)
                arquivosInternos = rar_ref.namelist()
                #------------------------------------
                for i_dir in arquivosInternos:
                    arquivoDescompactado_name = os.path.basename(i_dir)
                    arquivoDescompactado_dir = os.path.join(pasta_destino, arquivoDescompactado_name)
                    #--------------------------------------------------------------------------------
                    arquivosDescompactados_list.append(arquivoDescompactado_dir)
                    #-----------------------------------------------------------
                    arquivosCompactados_list.remove(caminho_arquivo)
                    arquivosParaDeletar_list.append(caminho_arquivo)
    #------------------------------------------------------------------------------------------------
    if LISTAR_NAO_DESCOMPACTADOS:
        arquivosDescompactados_list.extend(arquivosCompactados_list)
    #------------------------------------------------------------------------------------------------
    if DELATAR_COMPACTADADOS:
        for arqDel in arquivosParaDeletar_list:
            CreatedTools.DeletarArquivo(arqDel)
    #------------------------------------------------------------------------------------------------
    return arquivosDescompactados_list
    
    
    
def ConcatenarArquivosParaDF(DIRETORIOS_LISTA=[]):
    dataFrameConcatenado = pd.DataFrame()
    #-----------------------------------------------------------------
    excelTipo_list = ['.xls','.xlsx','.xlsm','.xlt','.xltx','.xlsb']
    #---------------------------------------------------------
    for arquivo in DIRETORIOS_LISTA:
        tipoArquivo = (os.path.splitext(arquivo))[1]
        #-----------------------------------------------------
        try:
            if tipoArquivo in excelTipo_list:
                df = pd.read_excel(arquivo)
            #-----------------------------------------------------
            elif tipoArquivo == ".csv":
                df = pd.read_csv(arquivo)
            #-----------------------------------------------------  
        except:
            df = pd.read_html(arquivo,header=0)
            df = df[0]
        #---------------------------------------------------------------------------------
        if len(dataFrameConcatenado) > 0:
            dataFrameConcatenado = pd.concat([dataFrameConcatenado,df], ignore_index=True)
        else:
            dataFrameConcatenado = df
    #----------------------------------------------------------------------------------     
    dataFrameConcatenado = dataFrameConcatenado.drop_duplicates()
    #----------------------------------------------------------------------------------
    return dataFrameConcatenado
         
       

def RelatorioTotalExpress(DATA_INICIO='2001-01-01',DATA_FINAL='2031-01-01',QTD_LOTE_CAFS=300,PAGINAS_CAF=30,DIRETORIO=cafsIcsDrive_dir):
    def TempoDeExecucao():
        if not hasattr(TempoDeExecucao, 'inicioDaExecucao'):
            TempoDeExecucao.inicioDaExecucao = datetime.datetime.now()
        #-------------------------------------------------------------
        diferencaDeTempo = datetime.datetime.now() - TempoDeExecucao.inicioDaExecucao
        tempoDeExecucao_time = str(diferencaDeTempo)
        #-------------------------------------------
        return tempoDeExecucao_time
    TempoDeExecucao()
    #-----------------------------------------------------------------------------------------------------------------------
    paginas = PAGINAS_CAF
    for baseOp in ['PFD','CSX']:
        LoginICS(baseOp) 
        #------------------------------
        ExportarOperacaoCafs(1,paginas)  
    #-----------------------------------------------------------------------------------------------------------------------  
    CreatedTools.funcionVBA('TratarColunasDeNumeros',DIRETORIO)
    if os.path.isfile(DIRETORIO):
        df_cafs = pd.read_excel(DIRETORIO)
    else:
        df_cafs = pd.DataFrame() 
    #-----------------------------------------------------------------------------------------------------------------------
    def tentar_formatos(data):
        formatos = ['%d/%m/%Y', '%Y-%m-%d %H:%M:%S']  
        for formato in formatos:
            try:
                return pd.to_datetime(data, format=formato)
            except ValueError:
                pass
        
    #----------------------------------------------------------------------------------
    df_cafs['Data de Abertura'] = df_cafs['Data de Abertura'].apply(tentar_formatos)
    #-----------------------------------------------------------------------------------------------------------------------
    cafsFiltradas = df_cafs[(df_cafs['Data de Abertura'] >= DATA_INICIO) & (df_cafs['Data de Abertura'] <= DATA_FINAL)]
    print(f'---->CAFs para baixar: {len(cafsFiltradas)}')
    #-----------------------------------------------------------------------------------------------------------------------
    LoginICS('Gerencial')
    cafsBaixadas = 0
    for i in range(0,len(cafsFiltradas),QTD_LOTE_CAFS):
        selecaoLoteCafs = cafsFiltradas[i:i+QTD_LOTE_CAFS]
        selecaoLoteCafs = selecaoLoteCafs['C.A.F.']
        texto_da_coluna = selecaoLoteCafs.to_string(index=False)
        pyperclip.copy(texto_da_coluna)
        BaixarLote(texto_da_coluna)
        cafsBaixadas = cafsBaixadas + len(selecaoLoteCafs)
        print(f'CAFs Baixadas: {cafsBaixadas}')
    #-----------------------------------------------------------------------------------------------------------------------
    arquivosbaixados_list = ListarArquivosRecentes(TEMPO_MODIFICACAO=TempoDeExecucao(),DIRETORIO=download_dir)
    #-----------------------------------------------------------------------------------------------------------------------
    arquivosDescompactados_list = DescompactarArquivos(arquivosbaixados_list)
    #-----------------------------------------------------------------------------------------------------------------------
    relatorioTotalBaixado_df = ConcatenarArquivosParaDF(arquivosDescompactados_list)
    arquivosbaixados_list = ListarArquivosRecentes(TEMPO_MODIFICACAO=TempoDeExecucao(),DIRETORIO=download_dir)
    relatorioTotalBaixado_df.to_excel(f'{download_dir}\\relatorioTotalFinal.xlsx')
    #-----------------------------------------------------------------------------------------------------------------------
    #Filtrar Status
    #Filtrar Data
    #Indentificar Multiplas
    





#SelectCheckBox()
RelatorioTotalExpress(DATA_INICIO='2023-10-24',DATA_FINAL='2023-11-15',PAGINAS_CAF=2)
#ExportarOperacaoCafs(1,30)

