import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

import win32com.client
from core.banco import conecta_engenharia, conecta
from core.erros import trata_excecao
from core.email_service import dados_email
from core.inventor import definir_classificacao
from core.inventor import padronizar_caminho, corrigir_caminho_inventor
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

class ConsultarCotas:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

        self.manipula_comeco()

    def manipula_comeco(self):
        try:
            cotas = {38, 31.80, 25.30}

            placeholders = ", ".join("?" for _ in cotas)

            cursor_eng = conecta_engenharia.cursor()
            cursor_eng.execute(f"""
                SELECT cota.ID_ARQUIVO, arq.caminho, arq2.caminho, idw.ID_ARQUIVO_REFERENCIA, cota.VALOR_COTA
                FROM COTAS_IDW as cota
                INNER JOIN PROPRIEDADES_IDW AS idw ON cota.ID_ARQUIVO = idw.ID_ARQUIVO
                INNER JOIN ARQUIVOS AS arq ON idw.ID_ARQUIVO_REFERENCIA = arq.id
                INNER JOIN ARQUIVOS AS arq2 ON idw.ID_ARQUIVO = arq2.id
                WHERE cota.VALOR_COTA IN ({placeholders})
            """, tuple(cotas))

            cotas_desenho = cursor_eng.fetchall()

            arquivos = {}

            for id_arquivo, caminho, caminho2, id_arq, valor_cota in cotas_desenho:
                if id_arquivo not in arquivos:
                    arquivos[id_arquivo] = {
                        "caminho": caminho,
                        "caminho2": caminho2,
                        "id_arq": id_arq,
                        "cotas": set()
                    }

                arquivos[id_arquivo]["cotas"].add(float(valor_cota))

            arquivos_completos = []

            for id_arquivo, dados in arquivos.items():
                if cotas.issubset(dados["cotas"]):
                    arquivos_completos.append(
                        (id_arquivo, dados["caminho"], dados["caminho2"], dados["id_arq"])
                    )

            for id_arquivo, caminho, caminho2, id_arq in arquivos_completos:
                print("ID IDW:", id_arquivo, "ID ARQUIVO:", id_arq)
                print("CAMINHO:", caminho)
                print("CAMINHO2:", caminho2)

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    ConsultarCotas()