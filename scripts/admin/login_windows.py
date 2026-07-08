from datetime import datetime
import getpass
import socket

agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

texto = f"""
LOGIN DETECTADO

Usuário: {getpass.getuser()}
Computador: {socket.gethostname()}
Data: {agora}

"""

with open(r"C:\Logs\login_windows.txt", "a", encoding="utf-8") as arq:
    arq.write(texto)
    arq.write("-" * 50 + "\n")