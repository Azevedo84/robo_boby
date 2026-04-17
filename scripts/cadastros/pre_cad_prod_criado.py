import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from core.erros import trata_excecao
from core.email_service import dados_email


class EnviaPreCadastroProduto:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

    def envia_email(self, lista_produtos):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'PRE - Cadastro de Produtos!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos para gerar cadastro:\n\n" \

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

    def manipula_dados_prod(self):
        try:
            lista_string = ""
            tabela = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, obs, descricao, descr_compl, referencia, um, ncm, "
                           f"kg_mt, fornecedor, data_criacao FROM PRODUTOPRELIMINAR "
                           f"WHERE codigo IS NULL;")
            dados_banco = cursor.fetchall()

            if dados_banco:
                for i in dados_banco:
                    id_pre, obs, descr, compl, ref, um, ncm, kg_mt, forn, emissao = i

                    datis = emissao.strftime("%d/%m/%Y")

                    dados = (datis, obs, descr, compl, ref, um, ncm, kg_mt, forn)
                    tabela.append(dados)

                    lista_string += f"Emissão: {datis}, " \
                                  f"Registro: {id_pre}, " \
                                  f"Observação: {obs}\n\n" \
                                    f"Descrição: {descr}\n" \
                                    f"Descrição Compl.: {compl}\n" \
                                    f"Referencia: {ref}\n" \
                                    f"UM: {um}\n" \
                                    f"NCM: {ncm}\n" \
                                    f"KG/MT: {kg_mt}\n" \
                                    f"Fornecedor: {forn}\n\n"

            if lista_string:
                self.envia_email(lista_string)

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = EnviaPreCadastroProduto()
chama_classe.manipula_dados_prod()
