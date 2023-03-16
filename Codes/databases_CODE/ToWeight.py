import pandas as pd

def ToWeight(dataFrame,columnsName):
    dataFrame =  dataFrame.reset_index(drop=True)
    
    i = 0
    for row in dataFrame[columnsName]:
        rowValue = str(dataFrame.loc[i, columnsName])

        if not rowValue.find(".")!= -1 or rowValue.find(",")!= -1:         
            if len(rowValue) > 4:
                rowValue = rowValue[:-3] + "." + rowValue[-3:]
            
        try:
            rowValue = float(rowValue)
        except:
            rowValue = 0.0
    
        dataFrame[columnsName] = dataFrame[columnsName].fillna(0.0)
    
        dataFrame.loc[i, columnsName] = rowValue
        i = i + 1 
    
    dataFrame = dataFrame.drop("Unnamed: 0", axis=1, errors='ignore')
    print("---- ToWeight finalizado com sucesso!")
    return dataFrame

arquivoPD = pd.read_excel(r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Relatorios\ARQUIVO DE TEXTO BASE 1ªQ0123.xlsx", sheet_name='BD')

ToWeight(arquivoPD,"PESO A")


