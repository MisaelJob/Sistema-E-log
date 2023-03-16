import pandas as pd
from datetime import datetime
from pandas import Timestamp
import numpy

def StatusFilter(clientList, dataFrame, columnsName):
    
    
    configSheet= pd.read_excel(r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Config\config.xlsx", sheet_name='Status')
    
    if type(clientList) == list:
        listStatus = clientList
    else:
        listStatus = configSheet[clientList].tolist()
    
    dataFrame = dataFrame[dataFrame[columnsName].isin(listStatus)]
    
    dataFrame =  dataFrame.reset_index(drop=True)
    dataFrame = dataFrame.drop("Unnamed: 0", axis=1, errors='ignore')
    print("---- StatusFilter finalizado com sucesso!")
    return dataFrame
