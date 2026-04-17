import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta, conecta_robo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date
from core.erros import trata_excecao
from core.email_service import dados_email


class EnviaPIProntas:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

    def envia_email(self, lista_produtos):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'OV - Produtos Prontos - Gerar Ordem de Venda!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos com saldo no estoque para gerar OV:\n\n" \

            body += f"{lista_produtos}\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado")

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_dados_pi(self):
        try:
            lista_string = ""
            lista_lista = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.id, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodint.qtde, prodint.data_previsao, prod.quantidade "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()

            if dados_interno:
                for i in dados_interno:
                    num_ped, id_cliente, cod, descr, ref, um, qtde, entrega, saldo = i

                    if saldo >= qtde:
                        dados = (num_ped, id_cliente, cod, descr, ref, um, qtde, entrega, saldo)
                        lista_lista.append(dados)

                        lista_string += f"{num_ped}, " \
                                      f"{id_cliente}, " \
                                      f"{cod}, " \
                                      f"{descr}, " \
                                      f"{ref}, " \
                                      f"{um}, " \
                                      f"{qtde}, " \
                                      f"{entrega}, " \
                                      f"{saldo}\n\n"

            return lista_string, lista_lista

        except Exception as e:
            trata_excecao(e)
            raise

    def inserir_no_banco(self, lista):
        try:
            data_hoje = date.today()

            for dados_ex in lista:
                num_ped, id_cliente, cod, descr, ref, um, qtde, entrega, saldo = dados_ex

                cursor = conecta_robo.cursor()
                cursor.execute(f"Insert into ENVIA_PI_PRONTAS (ID, NUM_PI, COD_PRODUTO, DATA_ENTREGA) "
                               f"values (GEN_ID(GEN_ENVIA_PI_PRONTAS_ID,1), {num_ped}, {cod}, '{data_hoje}');")
                print(f"PI {num_ped} enviado com sucesso!")
            conecta_robo.commit()

        except Exception as e:
            trata_excecao(e)
            raise

    def verifica_banco(self):
        try:
            texto, lista = self.manipula_dados_pi()

            nova_tabela = []

            if lista:
                for dados_ex in lista:
                    num_ped, id_cliente, cod, descr, ref, um, qtde, entrega, saldo = dados_ex

                    cursor = conecta_robo.cursor()
                    cursor.execute(f"SELECT * FROM envia_pi_prontas where num_pi = {num_ped} and cod_produto = {cod};")
                    select_envio = cursor.fetchall()

                    if not select_envio:
                        nova_tabela.append(dados_ex)

            if nova_tabela:
                self.envia_email(texto)
                self.inserir_no_banco(nova_tabela)

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = EnviaPIProntas()
chama_classe.verifica_banco()
