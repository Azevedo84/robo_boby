from core.banco import conecta, conecta_engenharia
from core.erros import grava_erro_banco
import re
import sys
import os
import traceback
import inspect
from dados_email import email_user, password
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import datetime


class GerarFilaValidacaoERP:
    def __init__(self):
        print("🚀 Gerador fila validação ERP iniciado")
        self.nome_arquivo = os.path.basename(__file__)

        self.processar()

    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

        except Exception as e:
            nome_funcao_trat = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            tb = traceback.extract_tb(exc_traceback)
            num_linha_erro = tb[-1][1]
            print(f'Houve um problema no arquivo: {self.nome_arquivo} na função: "{nome_funcao_trat}"\n'
                  f'{e} {num_linha_erro}')
            grava_erro_banco(nome_funcao_trat, e, self.nome_arquivo, num_linha_erro)

    def dados_email(self):
        try:
            to = ['<maquinas@unisold.com.br>']

            current_time = (datetime.now())
            horario = current_time.strftime('%H')
            hora_int = int(horario)
            saudacao = ""
            if 4 < hora_int < 13:
                saudacao = "Bom Dia!"
            elif 12 < hora_int < 19:
                saudacao = "Boa Tarde!"
            elif hora_int > 18:
                saudacao = "Boa Noite!"
            elif hora_int < 5:
                saudacao = "Boa Noite!"

            msg_final = ""

            msg_final += f"Att,\n"
            msg_final += f"Suzuki Máquinas Ltda\n"
            msg_final += f"Fone (51) 3561.2583/(51) 3170.0965\n\n"
            msg_final += f"🟦 Mensagem gerada automaticamente pelo sistema de Planejamento e Controle da Produção (PCP) do ERP Suzuki.\n"
            msg_final += "🔸Por favor, não responda este e-mail diretamente."


            return saudacao, msg_final, to

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def envia_email_desenho_duplicado(self, texto, desenho):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'ENGENHARIA PI - DESENHO DUPLICADO {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO desenho {desenho} está duplicado!\n\n"

            for i in texto:
                body += f"{i}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado DUPLICADO",desenho)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def inserir_fila_validacao_erp(self, cursor, id_arquivo):
        try:
            # 🔹 evita duplicado
            cursor.execute("""
                SELECT 1 
                FROM FILA_VALIDACAO_ERP
                WHERE ID_ARQUIVO = ?
                  AND STATUS IN ('PENDENTE', 'ERRO')
            """, (id_arquivo,))

            if cursor.fetchone():
                print(f"⚠️ Já na fila ERP: {id_arquivo}")
                return

            cursor.execute("""
                INSERT INTO FILA_VALIDACAO_ERP (ID_ARQUIVO, STATUS)
                VALUES (?, 'PENDENTE')
            """, (id_arquivo,))

            print(f"📥 Inserido na fila ERP: {id_arquivo}")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def processar(self):
        try:
            cursor_erp = conecta.cursor()
            cursor_eng = conecta_engenharia.cursor()

            cursor_erp.execute("""
                SELECT prod.codigo, prod.obs
                FROM PRODUTOPEDIDOINTERNO prodint
                JOIN produto prod ON prodint.id_produto = prod.id
                WHERE prodint.status = 'A' AND prod.descricao NOT LIKE '%KIT%'
            """)
            registros = cursor_erp.fetchall()

            print(f"📦 Total pedidos ativos: {len(registros)}")

            for codigo, obs in registros:

                ref = self.extrair_referencia(obs)

                if not ref:
                    print(f"⚠️ Produto {codigo} sem referência válida")
                    continue

                cursor_eng.execute("""
                    SELECT ID, TIPO_ARQUIVO
                    FROM ARQUIVOS
                    WHERE NOME_BASE = ?
                      AND TIPO_ARQUIVO IN ('IPT', 'IAM')
                """, (ref,))

                resultados = cursor_eng.fetchall()

                if not resultados:
                    print(f"❌ Sem desenho: {ref}")
                    continue

                if len(resultados) > 1:
                    print(f"⚠️ Duplicado: {ref}")
                    continue

                id_arquivo, tipo = resultados[0]

                # 🔹 INSERE NA FILA ERP
                self.inserir_fila_validacao_erp(cursor_eng, id_arquivo)

            # ✅ 👉 AQUI FICA O COMMIT
            conecta_engenharia.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def extrair_referencia(self, obs):
        try:
            if not obs:
                return None

            s = re.sub(r"[^\d.]", "", obs)
            s = re.sub(r"\.+$", "", s)

            return s if s else None

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

if __name__ == "__main__":
    GerarFilaValidacaoERP()