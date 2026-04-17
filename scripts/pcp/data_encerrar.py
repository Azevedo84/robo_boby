import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
from core.erros import trata_excecao
from datetime import datetime


class DataEncerramento:
    def __init__(self):

        self.manipula_data()

    def manipula_data(self):
        try:
            cursor = conecta.cursor()
            cursor.execute("SELECT data FROM DATALIMITE;")
            sel_estrutura = cursor.fetchone()

            primeiro_dia_mes_atual = datetime.today().replace(day=1).date()

            if sel_estrutura and sel_estrutura[0] != primeiro_dia_mes_atual:
                cursor.execute("UPDATE DATALIMITE SET data = ?", (primeiro_dia_mes_atual,))
                conecta.commit()

                print("DATA ATUALIZADA!")

            cursor.close()
            conecta.close()

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = DataEncerramento()