from pathlib import Path

centralFolder = Path(__file__).resolve().parents[2]
def Directorys(dir=None):
    if dir == None:
        dir = str(centralFolder)
    elif dir == "dowloads":
        dir = str(centralFolder) + r"\Downloads"
    elif dir == "relatorios":
        dir = str(centralFolder) + r"\Relatorios"
    elif dir == "config":
        dir = str(centralFolder) + r"\Config\config.xlsx"

    return dir