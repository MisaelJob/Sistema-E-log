from pathlib import Path
import pandas as pd

def ConcatArchives(tagName, folder_Dir=None):
    fileCategoriName = "_concatenado"
    suffixType = ".xlsx"
    newSheetFolder = r"C:\Users\Misael\Documents\Estudos\Sistema-E-log\Relatorios"
    
    newSheetName = tagName + fileCategoriName + suffixType
    
    if folder_Dir == None:
        folder_Dir = r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Downloads"
    
   
    if not Path.exists(Path(folder_Dir)):
        Path.mkdir(folder_Dir)
    
    
    importArchivesList = list(Path(folder_Dir).glob("*" + tagName + "*"))
    
    dadosConcatArquive_df = pd.DataFrame()
          
    for archive in importArchivesList:
        dadosImport_df = pd.read_excel(archive)
        dadosConcatArquive_df = pd.concat([dadosConcatArquive_df, dadosImport_df], ignore_index=True)
    
    
    dadosConcatArquive_df = dadosConcatArquive_df.reset_index()
    
    print("---- ConcatArchives" ,importArchivesList ,"finalizado com sucesso!")
    return dadosConcatArquive_df
   

