import pandas as pd

def ColumnsFromOther(dataFrame,refColumns,anotherDataFrame,importColumnList):
    dataFrame =  dataFrame.reset_index(drop=True)

    r = 0
    for referenceRef in dataFrame[refColumns]:
        dataFrame.loc[r,importColumnList] = ""

        i = 0
        for importRef in anotherDataFrame[refColumns]:
            if str(referenceRef) == str(importRef):
                
                for col in importColumnList:
                    
                    dataFrame.loc[r,col] = str(anotherDataFrame.loc[i, col])
            
            i = i + 1
        
        r = r + 1
    
    try:
        dataFrame = dataFrame.drop("Unnamed: 0", axis=1, errors='ignore')
    except:
        pass
    
    
    print("---- ColumnsFromOther finalizado com sucesso!")
    return dataFrame
    


