import pandas as pd
from datetime import datetime
from pandas import Timestamp
import numpy

def ToDate(dataFrame, columnsName = None):
    
    #---------------------------------------------------------------------------------------------
    if columnsName == None:
        
        cell = dataFrame
       
        if type(cell) == float or type(cell) == int:
                if cell > 36526:
                    try:
                        cell = pd.to_datetime(cell, unit='D', origin='1899-12-30')
                    except:
                        cell =  datetime(1899, 12, 30, 0, 0, 0)
        elif type(cell) == pd._libs.tslibs.timestamps.Timestamp:
                cell = pd.to_datetime(cell)
        elif type(cell) == str:
            try:
                cell = datetime.strptime(cell, "%Y-%m-%d")
            except:
                try:
                    cell = datetime.strptime(cell, "%Y/%m/%d")
                except:
                    try:
                        cell = datetime.strptime(cell, "%d-%m-%Y")
                    except:
                        try:
                            cell = datetime.strptime(cell, "%d/%m/%Y")
                        except:
                            try:
                                cell = datetime.strptime(cell, "%m-%d-%Y")
                            except:
                                try:
                                    cell = datetime.strptime(cell, "%m/%d/%Y")
                                except:
                                    cell = datetime(1899, 12, 30, 0, 0, 0)
            else:
                try:
                    cell = cell.astype(float)
                    cell = pd.to_datetime(cell, unit='D', origin='1899-12-30')
                except:
                    cell = datetime(1899, 12, 30, 0, 0, 0)
        
        dataFrame = cell
    #---------------------------------------------------------------------------------------------
    else:
        i = 0
        for cell in dataFrame[columnsName]:
            
            cell = dataFrame.loc[i, columnsName]
        
            if type(cell) == float or type(cell) == int:
                if cell > 36526:
                    try:
                        cell = pd.to_datetime(cell, errors='coerce', unit='D', origin='1899-12-30')
                    except:
                        cell =  datetime(1899, 12, 30, 0, 0, 0)
            elif type(cell) == pd._libs.tslibs.timestamps.Timestamp:
                cell = pd.to_datetime(cell)
            elif type(cell) == str:
                try:
                    cell = datetime.strptime(cell, "%Y-%m-%d")
                except:
                    try:
                        cell = datetime.strptime(cell, "%Y/%m/%d")
                    except:
                        try:
                            cell = datetime.strptime(cell, "%d-%m-%Y")
                        except:
                            try:
                                cell = datetime.strptime(cell, "%d/%m/%Y")
                            except:
                                try:
                                    cell = datetime.strptime(cell, "%m-%d-%Y")
                                except:
                                    try:
                                        cell = datetime.strptime(cell, "%m/%d/%Y")
                                    except:
                                        cell = datetime(1899, 12, 30, 0, 0, 0)
            else:
                try:
                    cell = cell.astype(float)
                    cell = pd.to_datetime(cell, unit='D', origin='1899-12-30')
                except:
                    cell = datetime(1899, 12, 30, 0, 0, 0)
            dataFrame.loc[i, columnsName] = cell
            i = i + 1
    #---------------------------------------------------------------------------------------------
    
    try:
        dataFrame = dataFrame.drop("Unnamed: 0", axis=1, errors='ignore')
    except:
        pass
    print("---- ToDate finalizado com sucesso!")
    return dataFrame
