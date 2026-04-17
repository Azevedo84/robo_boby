import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
from core.erros import trata_excecao
from core.email_service import dados_email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta


class EnviaRequisicaoPendentes:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

    def envia_email(self, numero, data, cod, descr, ref, um, qtde):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'Req - Requisição Nº {numero} - Não foi gerado Ordem de Compra!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}<br><br>" \
                   f"A Requisição Nº <b>{numero}</b> não foi gerado OC.<br><br>" \

            body += f"- <b>Data Emissão:</b> {data}<br> " \
                    f"- <b>Código:</b> {cod}<br> " \
                    f"- <b>Descrição:</b> {descr}<br> " \
                    f"- <b>Referência:</b> {ref}<br> " \
                    f"- <b>Quantidade:</b> {qtde} {um}<br>"

            body += "<br><br>"

            body += f"Att,<br><br>" \
                    f"Suzuki Máquinas Ltda<br>" \
                    f"Fone (51) 3561.2583/(51) 3170.0965<br>" \
                    f"Mensagem enviada automaticamente, por favor não responda.<br>" \
                    f"Se houver algum problema com o recebimento de emails favor entrar em contato " \
                    f"pelo email maquinas@unisold.com.br."

            msg.attach(MIMEText(body, 'html'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)

            server.quit()

            data_obj = datetime.strptime(data, '%d-%m-%Y')
            data_formatada = data_obj.strftime('%Y-%m-%d')

            cursor = conecta.cursor()
            cursor.execute(f"Insert into envia_req_falta (ID, num_req, data_emissao) "
                           f"values (GEN_ID(GEN_ENVIA_REQ_FALTA_ID,1), {numero}, '{data_formatada}');")
            conecta.commit()

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_dados(self):
        try:
            data_atual = datetime.now()
            data_semana_anterior = data_atual - timedelta(days=14)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prodreq.produto, prodreq.quantidade, "
                           f"(extract(day from req.data)||'-'||"
                           f"extract(month from req.data)||'-'||"
                           f"extract(year from req.data)) AS DATA, prodreq.numero, "
                           f"prodreq.destino, prodreq.id_prod_sol "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                           f"where prodreq.status = 'A';")
            select_req = cursor.fetchall()

            tabela = []
            for dados_req in select_req:
                produto, qtde, data, numero, destino, id_prod_sol = dados_req

                cur = conecta.cursor()
                cur.execute(f"SELECT codigo, descricao, COALESCE(obs, ' ') as obs, unidade "
                            f"FROM produto where id = {produto};")
                detalhes_produtos = cur.fetchall()
                cod, descr, ref, um = detalhes_produtos[0]

                if id_prod_sol is None:
                    num_sol = "X"
                else:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, mestre "
                                   f"FROM produtoordemsolicitacao "
                                   f"WHERE id = {id_prod_sol};")
                    extrair_sol = cursor.fetchall()
                    id_sol, num_so = extrair_sol[0]

                    if not extrair_sol:
                        num_sol = "X"
                    else:
                        num_sol = num_so

                data1 = datetime.strptime(data, '%d-%m-%Y')
                if data1 < data_semana_anterior:
                    dados = (num_sol, numero, data, cod, descr, ref, um, qtde, destino)
                    tabela.append(dados)

            if tabela:
                num_requisicao = ""
                for tti in tabela:
                    num_sol_f, num_f, data_f, cod_f, descr_f, ref_f, um_f, qtde_f, destino_f = tti

                    if not str(num_f) in num_requisicao:
                        num_requisicao += str(num_f)

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT * FROM ENVIA_REQ_FALTA WHERE num_req = {num_f};")
                        extrai = cursor.fetchall()

                        if not extrai:
                            self.envia_email(num_f, data_f, cod_f, descr_f, ref_f, um_f, qtde_f)
                            print("EMAIL ENVIADO COM SUCESSO!!")

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = EnviaRequisicaoPendentes()
chama_classe.manipula_dados()
