import pandas as pd
from datetime import datetime
from pandas import Timestamp
import numpy

def DuplicateFilter(dataFrame):

    dataFrame = dataFrame.drop_duplicates()
    dataFrame =  dataFrame.reset_index(drop=True)
    
    configSheet = pd.read_excel(r"C:\Users\Misael\Documents\Estudos\Sistema-E-log\Config\config.xlsx", sheet_name='Multiplas')
    colList = configSheet['MULTIPLAS'].tolist()

    concatLinhas1 = ""
    concatLinhas2 = ""
    
    for i in range(0, len(dataFrame['Nº PEÇA']),1):
        
        concatLinhas1 = ""
        for col in colList:
            concatLinhas1 = concatLinhas1 + str(dataFrame.loc[i,col]) 
        
        awbA = str(dataFrame.loc[i,'Nº PEÇA'])
        cellMultipla = ""
        
        for n in range(0, len(dataFrame['Nº PEÇA']),1):
            awbB = str(dataFrame.loc[n,'Nº PEÇA'])
            
            concatLinhas2 = ""
            for col in colList:
                concatLinhas2 = concatLinhas2 + str(dataFrame.loc[n,col])
        
            if concatLinhas1 == concatLinhas2 and awbA != awbB:
                if cellMultipla == "":
                    cellMultipla = awbB
                else:
                    cellMultipla = cellMultipla + ", " + awbB

            dataFrame.loc[i,'MULTIPLAS'] = cellMultipla

    try:
        dataFrame = dataFrame.drop("Unnamed: 0", axis=1, errors='ignore')
    except:
        pass
    try:
        dataFrame = dataFrame.drop("index", axis=1, errors='ignore')
    except:
        pass
    print("---- DuplicateFilter finalizado com sucesso!")
    return dataFrame