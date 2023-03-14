import pandas as pd
from datetime import datetime
from pandas import Timestamp
from pathlib import Path
from Codes.relatorios.ConcatArchives import ConcatArchives
from Codes.relatorios.StandardFormat import StandardFormat
from Codes.relatorios.SegmentArchives import SegmentArchives


#extrair informações dos relatorios secundarios
#colocar em formato padrão
#juntar relatorios baixados
#juntar os cinco relatorios


#@@CRIANDO RELATORIO GERAL (TRATADO)--------------------------------------
def UnitedReport():
    downloads_Dir = r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Downloads"
    relatorios_Dir = r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Relatorios"
    config_Dir = r"C:\Users\Misael\Documents\Estudos\Assistente_Transportadora\Config\config.xlsx"

    configColums = pd.read_excel(config_Dir, sheet_name="Columns")
    configAuxiliares = pd.read_excel(config_Dir, sheet_name="Auxiliares")
    toComplemented_list = configAuxiliares.columns.to_list()
    cliente_list = configColums['CLIENTE'].tolist()
    #-----------------------------------------------------------------------------------------
    for cliente in cliente_list:
        Segment_Dir = SegmentArchives(cliente)
        
        clintArchives_list = list(Path(Segment_Dir).glob("*" + cliente + "*"))
        for clientArchive_Dir in clintArchives_list:
            formated_Dir = StandardFormat(cliente,clientArchive_Dir)

        clintArchives_list = list(Path(formated_Dir).glob("*" + cliente + "*"))
        print(clintArchives_list)
        for clientArchive_Dir in clintArchives_list:
            clientArchive_DF = pd.read_excel(clientArchive_Dir)
        
            clientArchive_DF.to_excel(clientArchive_Dir)
        
        clientCocatened_df = ConcatArchives(cliente,formated_Dir)
        clientCocatened_df.to_excel(relatorios_Dir + cliente + "_concatened.xlsx")       
        #---------------------------------------------------------------------------------------------
    
UnitedReport()


