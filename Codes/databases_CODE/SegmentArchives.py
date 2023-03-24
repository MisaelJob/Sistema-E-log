
from pathlib import Path
import pandas as pd

def SegmentArchives(tagName, folder_Dir=None, newFolder_Dir = None, segmentNumber = 5000):
    categoriName = "_segmt"

    #--------------------------------------------------------------------------------------------
    if folder_Dir == None:
        folder_Dir = r"C:\Users\Misael\Documents\Estudos\Sistema-E-log\Downloads"
    if not Path(folder_Dir).exists():
        Path.mkdir(folder_Path)

    folder_Path = Path(folder_Dir)
    #--------------------------------------------------------------------------------------------
    if newFolder_Dir == None:
        newFolder_Dir = r"C:\Users\Misael\Documents\Estudos\Sistema-E-log\Relatorios" + "\\" + categoriName
    if not Path.exists(Path(newFolder_Dir)):
        Path.mkdir(newFolder_Dir)
    #--------------------------------------------------------------------------------------------
    archivesInFolder_list = list(folder_Path.glob("*" + tagName + "*"))
    archiveIndex = 1
    
    for archive_Dir in archivesInFolder_list:
        archive_df = pd.read_excel(archive_Dir)
        archive_df =  archive_df.reset_index(drop=True)

        rowsInArchive = archive_df.shape[0]
        num_sheets = rowsInArchive // segmentNumber + 1

        for i in range(0,num_sheets, 1):
            startRow = i * segmentNumber+i
            endRow = (i+1) * segmentNumber
            
            segmentSheet = archive_df.iloc[startRow:endRow]

            segmentSheet.to_excel(newFolder_Dir + "\\" + tagName + categoriName + "_" + str(archiveIndex) + ".xlsx")
            
            archiveIndex = archiveIndex + 1
    
    return newFolder_Dir











 