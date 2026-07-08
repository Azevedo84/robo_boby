import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta, conecta_robo
from core.erros import trata_excecao
from core.email_service import dados_email
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class NFComprasPendentes:
    def __init__(self):

        self.destinatario = ['<maquinas@unisold.com.br>']

        self.manipula_comeco()

    def manipula_comeco(self):
        try:
            hoje = date.today()

            primeiro_dia = hoje.replace(day=1)

            if hoje.month == 12:
                proximo_mes = date(hoje.year + 1, 1, 1)
            else:
                proximo_mes = date(hoje.year, hoje.month + 1, 1)

            # Volta um mês no primeiro_dia
            if primeiro_dia.month == 1:
                primeiro_dia = date(primeiro_dia.year - 1, 12, 1)
            else:
                primeiro_dia = date(primeiro_dia.year, primeiro_dia.month - 1, 1)

            cursor = conecta.cursor()
            cursor.execute("""
                SELECT DISTINCT
                    mov.data,
                    oc.numero,
                    ent.nota,
                    forn.id, 
                    forn.razao
                FROM entradaprod ent
                INNER JOIN movimentacao mov ON ent.movimentacao = mov.id
                INNER JOIN fornecedores forn ON ent.fornecedor = forn.id
                INNER JOIN ordemcompra oc ON ent.ordemcompra = oc.id
                WHERE mov.data >= ?
                  AND mov.data < ?
                  AND oc.entradasaida = 'E'
                ORDER BY mov.data, oc.numero
            """, (primeiro_dia, proximo_mes))
            dados_compras = cursor.fetchall()

            print(primeiro_dia, proximo_mes)

            if dados_compras:
                for i in dados_compras:
                    data_mov, num_oc, num_nf, id_forn, fornecedor = i

                    cursor = conecta_robo.cursor()
                    cursor.execute("""
                                SELECT ID
                                FROM ENVIA_NF_PRE_SEM_VINCULO
                                WHERE NUM_NF = ? 
                                and ID_FORNECEDOR = ?
                            """, (num_nf, id_forn))
                    registro = cursor.fetchone()

                    if not registro:
                        cursor_itens = conecta.cursor()
                        cursor_itens.execute("""
                                            SELECT NF.NUMERO_NF, NF.DATA_EMISSAO, FORN.RAZAO, FORN.ESTADO
                                            FROM PRE_NF_ENTRADA NF
                                            INNER JOIN FORNECEDORES as forn ON nf.ID_FORNECEDOR = forn.id
                                            WHERE NF.NUMERO_NF = ?
                                        """, (num_nf,))
                        nfs_pre = cursor_itens.fetchall()

                        if not nfs_pre:
                            print(data_mov, num_oc, num_nf, fornecedor)

                            self.enviar_email(num_nf, id_forn, num_oc, fornecedor)


        except Exception as e:
            trata_excecao(e)
            raise

    def gravar_envio(self, num_nf, id_forn):
        try:
            cursor = conecta_robo.cursor()
            cursor.execute("""
                INSERT INTO ENVIA_NF_PRE_SEM_VINCULO
                (NUM_NF, ID_FORNECEDOR)
                VALUES
                (?, ?)
            """, (num_nf, id_forn))
            conecta_robo.commit()

            print("Envio registrado com sucesso.")

        except Exception as e:
            trata_excecao(e)
            raise

    def enviar_email(self, num_nf, id_forn, num_oc, fornecedor):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            assunto = f"A NF {num_nf} - {fornecedor} não foi enviada automaticamente"

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['To'] = ", ".join(self.destinatario)
            msg['Subject'] = assunto

            body = (
                f"{saudacao}\n\n"
                f"A NF {num_nf} da OC {num_oc} - {fornecedor} não foi enviada automaticamente.\n\n"
                f"{msg_final}"
            )

            msg.attach(MIMEText(body, "plain"))

            # ==========================
            # Envia email
            # ==========================
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(
                email_user,
                self.destinatario,
                msg.as_string()
            )

            server.quit()

            # ==========================
            # Grava no banco
            # ==========================
            self.gravar_envio(num_nf, id_forn)

            print("Relatório enviado com sucesso.")

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = NFComprasPendentes()