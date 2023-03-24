import pandas as pd
from datetime import datetime
from pandas import Timestamp
from Codes.relatorios.ToDate import ToDate

def DateFilter(dataFrame, columnsName, start = None, end = None):
  
    dataFrame = ToDate(dataFrame,columnsName)
    dataFrame = dataFrame.reset_index()

    configSheet = pd.read_excel(r"C:\Users\Misael\Documents\Estudos\Sistema-E-log\Config\config.xlsx", sheet_name='Date')

    if start == None and end == None:
        start = ToDate(configSheet.loc[0, 'start'])
        end = ToDate(configSheet.loc[0, 'end'])
        
        start = start.replace(hour=0, minute=0)
        dataFrame = dataFrame.loc[dataFrame[columnsName] >= start]
        
        end = end.replace(hour=23, minute=59)
        dataFrame = dataFrame.loc[dataFrame[columnsName] <= end]
    else:  
        if not start == None:
            start = ToDate(start)
            start = start.replace(hour=0, minute=0)
            dataFrame = dataFrame.loc[dataFrame[columnsName] >= start]
        if not end == None:
            end = ToDate(end)
            end = end.replace(hour=23, minute=59)
            dataFrame = dataFrame.loc[dataFrame[columnsName] <= end]
    
    #dataFrame = dataFrame.drop("Unnamed: 0", axis=1, errors='ignore')
    print("---- DateFilter finalizado com sucesso!")
    return dataFrame