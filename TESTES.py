import pandas as pd
from datetime import datetime
from pandas import Timestamp
from pathlib import Path


thisArchive_dir = Path().absolute()
caracters_dir = str(thisArchive_dir).find("Sistema-E-log") + 13
rootFolder_dir = str(thisArchive_dir)[0:caracters_dir]
    
#extrair informações dos relatorios secundarios
#colocar em formato padrão
#juntar relatorios baixados
#juntar os cinco relatorios


#@@CRIANDO RELATORIO GERAL (TRATADO)--------------------------------------
def UnitedReport():
    import sys
    print(sys.version)
UnitedReport()


