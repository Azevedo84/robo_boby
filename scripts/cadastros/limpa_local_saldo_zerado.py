import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
from core.erros import trata_excecao


class LimpaLocalSaldoZerado:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

        self.manipula_comeco()

    def manipula_comeco(self):
        try:
            cursor = conecta.cursor()
            cursor.execute("""
                SELECT DISTINCT prod.id, prod.codigo, prod.localizacao
                FROM movimentacao AS mov
                INNER JOIN produto AS prod ON mov.produto = prod.id
                WHERE prod.quantidade = 0
                AND prod.localizacao IS NOT NULL 
                  AND prod.localizacao NOT LIKE 'A-%';
            """)
            dados_mov = cursor.fetchall()

            if dados_mov:
                for i in dados_mov:
                    id_prod, cod_prod, local = i
                    print(i)

                    cursor = conecta.cursor()
                    cursor.execute("UPDATE produto SET LOCALIZACAO = NULL WHERE id = ?", (id_prod,))

                    conecta.commit()

                    print(f"Cadastro do produto {cod_prod} atualizado com Sucesso {local}!")

        except Exception as e:
            trata_excecao(e)
            raise

chama_classe = LimpaLocalSaldoZerado()
chama_classe.manipula_comeco()
