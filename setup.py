#código mencionado no vídeo:
import sys
from cx_Freeze import setup, Executable

#Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"], "includes": ["tkinter","openpyxl"]}

#GUI applications require a different base on Windows (the default is for a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Reajuste - Wsac",
    version="0.1",
    description="Aplicativo para analise de dados feito em python por Rafael Brandolin.",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base, icon='box.ico'),]
)

#----------------------------------------------------
#--------- Comando para gerar o arquivo.exe ---------
#----------------------------------------------------
#---- No Terminal digite: python .\setup.py build ---
#----------------------------------------------------
#-- É Preciso ter a biblioteca cx-freeze instalada --
#----------------------------------------------------
#---- No Terminal digite: pip install cx-freeze -----
#----------------------------------------------------