import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta, conecta_robo
from core.erros import trata_excecao
from reportlab.lib.styles import getSampleStyleSheet


class RelatorioPIPendentes:
    def __init__(self):
        self.styles = getSampleStyleSheet()

        self.destinatario = ['<maquinas@unisold.com.br>']

        self.iniciar_processo()

    def iniciar_processo(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prod.codigo, prod.DESCRICAO, prod.obs, prod.unidade, prodint.qtde, "
                           f"prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A' "
                           f"order by ped.emissao;")
            dados_pi = cursor.fetchall()

            if dados_pi:
                for i_pi in dados_pi:
                    emissao_pi, num_pi, clie_pi, cod, descr, ref, um, qtde_pi, entrega_pi = i_pi
                    print(i_pi)

        except Exception as e:
            trata_excecao(e)
            raise


    def gerar_pdf(self):
        try:
            arquivo = "PI Pendentes.pdf"

            print("PDF gerado:",arquivo)

        except Exception as e:
            trata_excecao(e)
            raise

    def pode_enviar_oc_fim_mes(self):
        try:
            esta_tudo_certo = True

            if not esta_tudo_certo:
                return False

            return True

        except Exception as e:
            trata_excecao(e)
            raise

    def enviar_email(self):
        try:
            print("Email enviado com sucesso.")

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    try:
        rel = RelatorioPIPendentes()

        if rel.pode_enviar_oc_fim_mes():
            rel.gerar_pdf()
            rel.enviar_email()

        else:
            print("Houve problemas com Pedidos")

    except Exception as e:
        trata_excecao(e)
        raise