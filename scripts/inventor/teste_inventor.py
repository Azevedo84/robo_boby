with open(r"C:\Temp\teste.txt", "w") as f:
    f.write("Começou")

import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

with open(r"C:\Temp\teste.txt", "a") as f:
    f.write("\nAntes do import core.erros")

try:
    from core.erros import trata_excecao

    with open(r"C:\Temp\teste.txt", "a") as f:
        f.write("\nImport core.erros OK")

except Exception as e:
    with open(r"C:\Temp\teste.txt", "a") as f:
        f.write(f"\nERRO IMPORT CORE:\n{repr(e)}")
    raise

with open(r"C:\Temp\teste.txt", "a") as f:
    f.write("\nDepois do import core.erros")

try:

    with open(r"C:\Temp\teste.txt", "a") as f:
        f.write("\nAntes do import win32com")

    import win32com.client

    with open(r"C:\Temp\teste.txt", "a") as f:
        f.write("\nDepois do import win32com")

    with open(r"C:\Temp\teste.txt", "a") as f:
        f.write("\nAntes do Dispatch")

    import pythoncom

    pythoncom.CoInitialize()

    import subprocess

    subprocess.Popen(
        [r"C:\Program Files\Autodesk\Inventor 2016\Bin\Inventor.exe", "/Automation"]
    )

    with open(r"C:\Temp\teste.txt", "a") as f:
        f.write("\nInventor iniciado por Popen")

    import time

    time.sleep(10)

    inventor = win32com.client.Dispatch("Inventor.Application")

    with open(r"C:\Temp\teste.txt", "a") as f:
        f.write("\nDepois do Dispatch")

    inventor.Quit()

    with open(r"C:\Temp\teste.txt", "a") as f:
        f.write("\nInventor fechado")

    raise Exception("SUCESSO! O Inventor abriu normalmente.")

except Exception as e:
    trata_excecao(e)