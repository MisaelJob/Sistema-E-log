import pandas as pd
from pathlib import Path

def StandardFormat(Cliente,sheet_Dir,newSheetFolder_Dir = None):
    
    fileCategoriName = "_formatado"
    suffixType = ".xlsx"

    #Obrigado a passar porque é um arquivo porra
    #if sheet_Dir == None:
    #    sheet_Dir = r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Downloads"

    if newSheetFolder_Dir == None:
        newSheetFolder_Dir = r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Relatorios" + r"\\" + fileCategoriName
    
    newSheetName = str(Path(sheet_Dir).stem) + fileCategoriName + suffixType

    newSheet_Dir = newSheetFolder_Dir + "\\" + newSheetName
    #-------------------------------------------------------------------------------------------------------------
    if not Path.exists(Path(newSheetFolder_Dir)):
        Path(newSheetFolder_Dir).mkdir()

    if Path.exists(Path(newSheet_Dir)):
       Path(newSheet_Dir).unlink()
    #-------------------------------------------------------------------------------------------------------------
    configSheet = pd.read_excel(r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Config\config.xlsx", sheet_name='Columns')
    clientLine = configSheet.loc[configSheet['CLIENTE'] == Cliente].values.tolist()
    standardLine = configSheet.columns.to_list()
    #-------------------------------------------------------------------------------------------------------------
    sheet_df = pd.read_excel(sheet_Dir)
    columns_list = sheet_df.columns.to_list()
    columns_list = list(set(clientLine[0]).intersection(columns_list))
    sheet_df = sheet_df.loc[:,columns_list]
    
    outColumns = list(set(clientLine[0]).difference(columns_list))

    for col in outColumns:       
        sheet_df = sheet_df.assign(newColumns = col)
        sheet_df = sheet_df.rename(columns={'newColumns': col})
    
    sheet_df = sheet_df[clientLine[0]]
    sheet_df = sheet_df.rename(columns=dict(zip(sheet_df.columns,standardLine)))
    sheet_df.to_excel(newSheet_Dir)
    
    print("---- StandardFormat finalizado com sucesso!")
    return newSheetFolder_Dir
    
