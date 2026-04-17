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
from core.erros import trata_excecao
from core.email_service import dados_email


class EnviaPICriadas:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

    def manipula_comeco(self):
        try:
            tabela_final = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prodint.data_criacao, ped.emissao, prodint.id_pedidointerno, prod.codigo, "
                           f"prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodint.qtde, prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    criacao, emisao, num_pi, cod, descr, ref, um, qtde, previsao = i

                    cursor = conecta_robo.cursor()
                    cursor.execute(f"SELECT NUM_PI, DATA_EMISSAO, COD_PRODUTO, QTDE, "
                                   f"DATA_PREVISAO "
                                   f"FROM ENVIA_PI_CRIADAS "
                                   f"where num_pi = {num_pi} "
                                   f"and data_emissao = '{emisao}' "
                                   f"and cod_produto = {cod} "
                                   f"and data_previsao = '{previsao}'"
                                   f"and qtde = {qtde};")
                    dados_criados = cursor.fetchall()

                    if not dados_criados:
                        cursor = conecta_robo.cursor()
                        cursor.execute(f"Insert into ENVIA_PI_CRIADAS (ID, NUM_PI, DATA_EMISSAO, COD_PRODUTO, QTDE, "
                                       f"DATA_PREVISAO) "
                                       f"values (GEN_ID(GEN_ENVIA_PI_CRIADAS_ID,1), {num_pi}, '{emisao}', {cod}, "
                                       f"'{qtde}', '{previsao}');")
    
                        conecta_robo.commit()

                        dados = (criacao, emisao, num_pi, cod, descr, ref, um, qtde, previsao)
                        tabela_final.append(dados)

                if tabela_final:
                    self.envia_email(tabela_final)

        except Exception as e:
            trata_excecao(e)
            raise

    def envia_email(self, dados):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'PI - Pedidos Internos Criados!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos criados nos Pedidos Internos:\n\n"

            for i in dados:
                criacao, emisao, num_pi, cod, descr, ref, um, qtde, previsao = i

                data_criacao = criacao.strftime("%d/%m/%Y")
                data_emisao = emisao.strftime("%d/%m/%Y")
                data_previsao = previsao.strftime("%d/%m/%Y")

                body += f"- Nª PI: {num_pi}\n" \
                        f"- Data Criação: {data_criacao}\n" \
                        f"- Data Emissão: {data_emisao}\n" \
                        f"- Produto: {cod} - {descr} - {ref} - {qtde}{um}\n" \
                        f"- Previsão: {data_previsao}\n\n"

            body += f"\n{msg_final}"

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


chama_classe = EnviaPICriadas()
chama_classe.manipula_comeco()
