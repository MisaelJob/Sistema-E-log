import pandas as pd
from pathlib import Path

def StandardFormat(sheet_Dir, nomeCliente):
    fileCategoriName = "_formatado"
    suffixType = ".xlsx"
    newSheetFolder_Dir = str(Path(sheet_Dir).parent) + r"\Formated"
    
    newSheetName = str(Path(sheet_Dir).stem) + fileCategoriName + suffixType

    newSheet_Dir = newSheetFolder_Dir + "\\" + newSheetName

    #-------------------------------------------------------------------------------------------------------------
    if not Path.exists(Path(newSheetFolder_Dir)):
        Path(newSheetFolder_Dir).mkdir()

    if Path.exists(Path(newSheet_Dir)):
       Path(newSheet_Dir).unlink()
    #-------------------------------------------------------------------------------------------------------------
    configSheet = pd.read_excel(r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Config\config.xlsx", sheet_name='Columns')
    clientLine = configSheet.loc[configSheet['CLIENTE'] == nomeCliente].values.tolist()
    standardLine = configSheet.columns.tolist()
    #-------------------------------------------------------------------------------------------------------------
    new_sheet = pd.DataFrame(columns=clientLine)
    new_sheet.to_excel(newSheet_Dir)
    new_list = clientLine
    
    old_sheet = pd.read_excel(sheet_Dir)
    old_list = old_sheet.columns.to_list()

    for oldCol in old_list:
        for newCol in new_list:
            if newCol==oldCol:
                new_sheet[newCol] =  old_sheet[oldCol]
            elif newCol==(oldCol+"#2"):
                new_sheet[newCol] =  old_sheet[oldCol]
            
    for col in new_list:
        if "#Auto" in str(col):
            new_sheet[col] = str(col).replace("#Auto", "")     
    
    new_sheet[nomeCliente] = nomeCliente
    new_sheet = new_sheet.fillna(0)

    try:
        new_sheet = new_sheet.drop("Unnamed: 0", axis=1)
    except:
        try:
            new_sheet = new_sheet.drop("Unnamed: 0", axis=0, errors='ignore')
        except:
            pass
    
    new_sheet.columns = standardLine
    new_sheet.to_excel(newSheet_Dir)
    
    print("---- StandardFormat finalizado com sucesso!")

    
