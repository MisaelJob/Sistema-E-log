from tkinter import Tk
import pyautogui
import pandas as pd
from pathlib import Path
import time
import re
import pyperclip
import xlwings as xw
import openpyxl
import os
import datetime
import psutil
import xlrd

def DetectResolution():
    root = Tk()
    Xwidht = root.winfo_screenwidth()
    Yhight = root.winfo_screenheight()
    nowResolution = f"{Xwidht}x{Yhight}"
    #print(f'----------> Resulução de tela {nowResolution},localizada.')
    return nowResolution
resolution = DetectResolution()



def RootFolder():
    thisArchive_dir = Path().absolute()
    caracters_dir = str(thisArchive_dir).find("Sistema-E-log") + 13
    nowRootFolder_dir = str(thisArchive_dir)[0:caracters_dir]
    #print(f'----------> Pasta raiz definida: {nowRootFolder_dir}.')
    return nowRootFolder_dir
rootFolder_dir = RootFolder()

 

def FindImage(imageName,posX = 0,posY = 0,action="click",attempts=4,imageFolder=f"{rootFolder_dir}/Codes/EnvioEspelho_CODE/images/{resolution}"):
    ultimoPontoImageName = str(imageName).rfind('.')
    nomeImgSemTipo = imageName[0:ultimoPontoImageName]
    enderecoImagem_list = [f'{imageFolder}\\{imageName}']
    #-------------------------------------------------------------------------
    #enderecoImagem_list.append(f'{imageFolder}\\{imageName}')
    arqSecundarios_list = list(Path(imageFolder).glob(f"{nomeImgSemTipo}(*"))
    enderecoImagem_list += arqSecundarios_list   
    #-------------------------------------------------------------------------
    for attempt in range(1, int(attempts)+1, 1):
        for img in enderecoImagem_list:
            returnValue = False
            #-----------------------
            pesquisa_wtt_posX = None 
            pesquisa_wtt_posY = None
            #----------------------
            time.sleep(1)
            #--------------------------------------------------------------------------------    
            if not os.path.exists(img):
                
                continue
            #----------------------------------------------------------
            try:
                pesquisa_wtt_posX, pesquisa_wtt_posY = pyautogui.locateCenterOnScreen(str(img), confidence=0.9)
                if pesquisa_wtt_posX != None:
                    break
            except:
                pass
        #-------------------------------------------------
        if pesquisa_wtt_posX != None:
            pesquisa_wtt_posX = pesquisa_wtt_posX + posX
            pesquisa_wtt_posY = pesquisa_wtt_posY + posY
            #----------------------------------------------
            if action == "aguardar":
                returnValue = True
                continue
            #--------------------------------------------------------
            elif action == "click":
                pyautogui.moveTo(x=pesquisa_wtt_posX,y=pesquisa_wtt_posY,duration=0.3)
                pyautogui.click(x=pesquisa_wtt_posX,y=pesquisa_wtt_posY)
                returnValue = True
                break
            #--------------------------------------------------------
            elif action == "moveTo":
                pyautogui.moveTo(x=pesquisa_wtt_posX,y=pesquisa_wtt_posY)
                returnValue = True
                break
            #--------------------------------------------------------
            else:
                continue   
        #-------------------------------------------------------------
    if returnValue == False:
        print(f'----------> Imagem {imageName}, não encontrada.')
        pass
    return returnValue



import xlwings as xw

def funcionVBA(FUNCTION_NAME, DIRETORIO, *args):
    try:
        app = xw.apps.active
        #----------------------------------------
        if not DIRETORIO == "":
            arquivoExcel = xw.Book(DIRETORIO)
        else:
            arquivoExcel = app.books.active
        #----------------------------------------
        arquivoExcel.macro(FUNCTION_NAME).run(*args)
        #----------------------------------------
        if not DIRETORIO == "":
            try:
                arquivoExcel.save()
                arquivoExcel.app.quit()
            except Exception as e:
                print(f'Erro ao salvar e fechar o arquivo: {e}')
    except Exception as e:
        print(f'Erro ao executar função {FUNCTION_NAME}: {e}')





def cttName(name):
    correctName = str(name)
    correctName = correctName.strip()
    correctName = correctName.upper()
    sobrenomes = correctName.split()
    novo_nome = ' '.join(sobrenomes[:-1])
    novo_nome = novo_nome.strip()
    #print(f'----------> Valor: {name}, tratado para: {novo_nome}')
    return novo_nome



def toTelephoneNum(text):
    text=str(text)
    numeros = re.findall(r'\d', text)
    #-------------------------------------------------------
    if len(numeros) >= 12:
        formateTel = int(''.join(numeros[:4] + numeros[-8:]))
    elif len(numeros) >= 10:
        formateTel = int('55' + ''.join(numeros[:2] + numeros[-8:]))
    elif len(numeros) >= 8:
        formateTel = int('5551' + str(numeros[-8:]))
    else:
        formateTel = 0
    return formateTel

    

def ProcurarContato_wtt(pesquisa,telefone=0):
    returnContatoEncontrado = False    
    FindImage('fecharPerfil_wtt.png')
    pyautogui.press('esc')
    #----------------------------------------------------
    if not valido(pesquisa):
        return False
    pyperclip.copy(pesquisa)
    #----------------------------------------------------
    metodosDeBusca = ["pesquisaWtt","linkDireto"]
    for metodoAtual in metodosDeBusca:
        if returnContatoEncontrado == True:
                break
        #------------------------------------------------
        if metodoAtual == "pesquisaWtt":
            #pyautogui.hotkey("ctrl","alt","/")
            pyautogui.click(824,169)
            pyautogui.hotkey("ctrl","a")
            pyautogui.hotkey('ctrl','v')
            pyautogui.press('enter')
            #--------------------------------------------
            #pyautogui.hotkey("ctrl","alt","/")
            pyautogui.click(824,169)
            pyautogui.hotkey("ctrl","a")
            pyautogui.press('backspace')
            #---------------------------------------------------------------------------------------------
        elif metodoAtual == "linkDireto":
            pyautogui.hotkey('alt','d')
            pyperclip.copy("https://web.whatsapp.com/send/?phone=" + str(toTelephoneNum(telefone)))
            pyautogui.hotkey('ctrl','v')
            pyautogui.press('enter')
            time.sleep(2)
            if not FindImage(imageName='inicioPagina_wtt.png',attempts=200,action='moveTo'):
                continue
            time.sleep(2)
        #---------------------------------------------------------------------------------------------
        if not FindImage('opcoesPerfil_wtt.png',20):
            if not FindImage('opcoesPerfil2_wtt.png'):
                pass
        #----------------------------------------------------
        if not FindImage('dadosDoContato_wtt.png'):
            if not FindImage('dadosDoContato_2_wtt.png'):
                if not FindImage('dadosDoContato_3_wtt.png'):
                    continue
        #----------------------------------------------------
        tentativaValidadarNomeContato = range(0,3,1)
        for tentantivaContador in tentativaValidadarNomeContato:
            if returnContatoEncontrado == True:
                break
            #----------------------------------------------------
            if tentantivaContador == 0:
                pyautogui.moveRel(-50, 213, duration=0.5)
            elif tentantivaContador == 1:
                pyautogui.moveRel(0, 20, duration=0.5)
            elif tentantivaContador == 2:
                pyautogui.moveRel(0, 30, duration=0.5)
            #----------------------------------------------------
            pyperclip.copy("")
            pesqContato_wtt = ""
            #-----------------------------------
            pyautogui.click(clicks=3)
            pyautogui.hotkey('ctrl','c')
            pesqContato_wtt = pyperclip.paste()
            #-----------------------------------
            if pesqContato_wtt == pesquisa:
                returnContatoEncontrado = True    
            elif cttName(pesqContato_wtt) == pesquisa:
                returnContatoEncontrado = True 
            #-------------------------------------------------------------------------------
            if valido(toTelephoneNum(pesqContato_wtt)) and valido(toTelephoneNum(telefone)):
                if toTelephoneNum(pesqContato_wtt) == toTelephoneNum(telefone):
                    returnContatoEncontrado = True                                     
            elif valido(toTelephoneNum(pesqContato_wtt)) and valido(toTelephoneNum(pesquisa)):   
                if toTelephoneNum(pesqContato_wtt) == toTelephoneNum(pesquisa):
                    returnContatoEncontrado = True                 
    #----------------------------------------------------------------------------------------
    if not FindImage('fecharPerfil_wtt.png'):
        pyautogui.press("esc") 
    #----------------------------------------
    return returnContatoEncontrado
    #----------------------------------------
    
    
  
def ArchiveType(arquivo):
    arquivo = str(arquivo)
    arquivo = arquivo.lower()
    resposta = False
    if  arquivo.find(".png") != -1:
        resposta = "image"
    elif arquivo.find(".jpeg") != -1:
        resposta = "image"
    elif arquivo.find(".jpg") != -1:
        resposta = "image"
    elif arquivo.find(".gif") != -1:
        resposta = "image" 
    elif arquivo.find(".tiff") != -1:
        resposta = "image"
    elif arquivo.find(".jfif") != -1:
        resposta = "image"
    elif arquivo.find(".svg") != -1:
        resposta = "image"
    elif arquivo.find(".webp") != -1:
        resposta = "image"
    elif arquivo.find(".") != -1:
        resposta = "archive"
    elif arquivo.find("espelho") != -1 or arquivo.find("fechamento") != -1:
        resposta = "espelho"
    #----------------------------------------------------
    if resposta:
        #print(f'----------> Arquivo {arquivo}, encontrada com sucesso.')
        pass
    else:
        #print(f'----------> Arquivo {arquivo}, não encontrada.')
        pass
    #----------------------------------------------------   
    return resposta
        
        

def MousePosition_X_Y():
    time.sleep(2)
    print(pyautogui.position())
#MousePosition_X_Y()



def valido(variavel):
    if type(variavel) == int:
        if variavel > 0:
            return True
    elif type(variavel) == str:
        if len(variavel) > 0:
            return True
    #print(type(variavel))
    return False
    


def SavarDataFrameEmExcel(DATA_FRAME, DIRETORIO):
    print("Save")
    try:
        funcionVBA('TratarColunasDeNumeros',DIRETORIO)
    except:
        print('Erro ao executar o codigo VBA: TratarColunasDeNumeros')
    #----------------------------------------------------------------------------------------------------------------
    # Remove coluna de índices
    df = DATA_FRAME.reset_index(drop=True)

    # Remove colunas com nomes vazios ou começando com "Unnamed"
    df = df.loc[:, ~df.columns.str.match('Unnamed')]
    colunas_para_remover = [coluna for coluna in df.columns if coluna.startswith('Unnamed') or pd.isna(coluna)]
    df = df.drop(columns=colunas_para_remover)
    #----------------------------------------------------------------------------------------------------------------
    app = xw.App()
    try:
        wb = app.books.open(DIRETORIO)
    except:
        wb = app.books.add()
    sheet = wb.sheets[0]
    #----------------------------------------------------------------------------------------------------------------
    '''
    for coluna in df.columns:
        
        if coluna == 'Data de Abertura':
            df[coluna] = df[coluna].str.replace(r'\s+', '', regex=True)
            #-----------------------------------------------------------
            if pd.api.types.is_numeric_dtype(df[coluna]):
                try:
                    df[coluna] = df[coluna].apply(lambda x: xlrd.xldate.xldate_as_datetime(x, 0) if not pd.isna(x) else x)
                except ValueError:
                    pass 
        elif coluna == 'Hora':
            df[coluna] = df[coluna].str.replace(' hs', '')
        #-----------------------------------------------------------------------------------------------------------------
        if pd.api.types.is_datetime64_any_dtype(df[coluna]):
            try:
                df[coluna] = df[coluna].dt.strftime('%d-%m-%Y %H:%M:%S')
            except AttributeError:
                pass
            df[coluna] = df[coluna].astype(str)
        elif pd.api.types.is_timedelta64_dtype(df[coluna]):
            df[coluna] = df[coluna].astype(str)
        elif pd.api.types.is_object_dtype(df[coluna]) and isinstance(df.iloc[0][coluna], datetime.time):
            df[coluna] = df[coluna].apply(lambda x: x.strftime('%H:%M:%S') if isinstance(x, datetime.time) else x)
        #----------------------------------------------------------------------------------------------------------------- 
        if pd.api.types.is_datetime64_any_dtype(df[coluna]):
            try:
                df[coluna] = df[coluna].dt.strftime('%d-%m-%Y %H:%M:%S')
            except AttributeError:
                pass
            df[coluna] = df[coluna].astype(str)
        elif pd.api.types.is_timedelta64_dtype(df[coluna]):
            df[coluna] = df[coluna].astype(str)
        elif pd.api.types.is_object_dtype(df[coluna]) and isinstance(df.iloc[0][coluna], datetime.time):
            df[coluna] = df[coluna].apply(lambda x: x.strftime('%H:%M:%S') if isinstance(x, datetime.time) else x)
        '''
        #-----------------------------------------------------------------------------------------------------------------
      
    #----------------------------------------------------------------------------------------------------------------
    if not df.empty:
        #-------------------------------------------
        sheet.range('A1:BZ999888').clear_contents()
        sheet.range('a1').value = df
        #-------------------------------------------
       
        # Obtém os nomes das colunas da planilha
        nomes_colunas = sheet.range((1, 1), (1, sheet.api.UsedRange.Columns.Count)).value

        # Verifica se o primeiro nome de coluna está vazio ou começa com "Unnamed"
        if not nomes_colunas[0] or nomes_colunas[0].startswith("Unnamed"):
            # Remove a primeira coluna da planilha
            sheet.api.Columns(1).Delete()
        
        #-------------------------------------------
        wb.save()
        wb.close()
        app.quit()
        #-------------------------------------------
        try:
            funcionVBA('TratarColunasDeNumeros',DIRETORIO)
        except:
            print('Erro ao executar o codigo VBA: SimplificarDados')
        #------------------------------------------------------------
    else:
        print("O DataFrame está vazio.")
    #----------------------------------------------------------------------------------------------------------------   



def DeletarArquivo(FILE_PATH,TEMPO_MAXIMO=300):
    arquivoEmUso = False
    #-------------------------------------------
    for tempoEspera in range(1,TEMPO_MAXIMO,1):
        time.sleep(1)
        #-----------------------------------------------------------------
        for process in psutil.process_iter(['pid', 'name', 'open_files']):
            try:
                open_files = process.info.get('open_files', [])
                if open_files and any(isinstance(f, str) and FILE_PATH.lower() in f.lower() for f in open_files):
                    arquivoEmUso = True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                arquivoEmUso = False
        #------------------------------------------------------------------------------
        if not arquivoEmUso:
            break
    #------------------------------------------------------------------------------
    try:
        if not arquivoEmUso:
            os.remove(FILE_PATH)
            return True
    except FileNotFoundError:
        print(f"Arquivo não encontrado: {FILE_PATH}")
    except:
        print(f"Erro ao tentar Deletar: {FILE_PATH}")
    #----------------------------------------------------------------------------------




    
    