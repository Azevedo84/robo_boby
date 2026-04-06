from core.banco import conecta_engenharia
import sys
import os
from core.erros import grava_erro_banco
import traceback
import inspect
from dados_email import email_user, password
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import datetime


class ValidarEstruturaERP:
    def __init__(self):
        print("🚀 Worker validação ERP iniciado")

        self.nome_arquivo = os.path.basename(__file__)

        self.processar_fila()

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

    def processar_fila(self):
        try:
            cursor = conecta_engenharia.cursor()

            cursor.execute("""
                SELECT ID_ARQUIVO
                FROM FILA_VALIDACAO_ERP
                WHERE STATUS = 'PENDENTE'
                ORDER BY ID
            """)

            fila = [row[0] for row in cursor.fetchall()]

            print(f"📦 Total na fila ERP: {len(fila)}")

            for id_arquivo in fila:
                print(f"\n🔍 Validando ID: {id_arquivo}")

                try:
                    self.validar_arquivo(cursor, id_arquivo)

                    cursor.execute("""
                        UPDATE FILA_VALIDACAO_ERP
                        SET STATUS = 'OK',
                            DATA_PROCESSAMENTO = CURRENT_TIMESTAMP
                        WHERE ID_ARQUIVO = ?
                    """, (id_arquivo,))

                    conecta_engenharia.commit()

                    print(f"✅ OK: {id_arquivo}")

                except Exception as e:
                    print(f"❌ ERRO: {id_arquivo} → {e}")

                    cursor.execute("""
                        UPDATE FILA_VALIDACAO_ERP
                        SET STATUS = 'ERRO',
                            DATA_PROCESSAMENTO = CURRENT_TIMESTAMP,
                            OBSERVACAO = ?
                        WHERE ID_ARQUIVO = ?
                    """, (str(e)[:200], id_arquivo))

                    conecta_engenharia.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def validar_arquivo(self, cursor, id_arquivo):
        # 🔹 pega árvore completa
        ids = self.buscar_toda_estrutura(cursor, id_arquivo)

        # 🔹 limpa erros de TODA a estrutura
        for id_item in ids:
            cursor.execute("""
                DELETE FROM ERROS_VALIDACAO_ERP
                WHERE ID_ARQUIVO = ?
            """, (id_item,))

        print(f"🌳 Total itens na estrutura: {len(ids)}")

        tem_erro = False

        for id_item in ids:
            erros = self.validar_item(cursor, id_item)

            if erros:
                tem_erro = True

                for erro in erros:
                    self.registrar_erro(cursor, id_item, erro)

        # 🔥 SE TEM ERRO → LEVANTA EXCEÇÃO
        if tem_erro:
            raise Exception("Estrutura com inconsistências")

    def buscar_toda_estrutura(self, cursor, id_pai):
        try:
            visitados = set()
            fila = [id_pai]

            while fila:
                atual = fila.pop()

                if atual in visitados:
                    continue

                visitados.add(atual)

                cursor.execute("""
                    SELECT ID_FILHO
                    FROM ESTRUTURA
                    WHERE ID_PAI = ?
                """, (atual,))

                filhos = [row[0] for row in cursor.fetchall()]

                fila.extend(filhos)

            return list(visitados)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def validar_item(self, cursor, id_arquivo):
        try:
            cursor.execute("""
                SELECT TIPO_ARQUIVO
                FROM ARQUIVOS
                WHERE ID = ?
            """, (id_arquivo,))

            tipo = cursor.fetchone()[0]

            if tipo == "IPT":
                return self.validar_ipt(cursor, id_arquivo)

            elif tipo == "IAM":
                return self.validar_iam(cursor, id_arquivo)

            return []

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def validar_ipt(self, cursor, id_arquivo):
        try:
            cursor.execute("""
                SELECT DESCRIPTION, AUTHORITY, COST_CENTER, REVISION_NUMBER
                FROM PROPRIEDADES_IPT
                WHERE ID_ARQUIVO = ?
            """, (id_arquivo,))

            row = cursor.fetchone()

            erros = []

            if not row:
                erros.append(("SEM_PROPRIEDADES", None, None))
                return erros

            descricao, authority, cost_center, revision = row

            if not descricao:
                erros.append(("FALTA_DESCRICAO", "DESCRIPTION", None))

            if not authority:
                erros.append(("FALTA_AUTHORITY", "AUTHORITY", None))

            if not cost_center:
                erros.append(("FALTA_COST_CENTER", "COST_CENTER", None))

            if not revision:
                erros.append(("FALTA_REVISION", "REVISION_NUMBER", None))

            return erros

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def validar_iam(self, cursor, id_arquivo):
        try:
            cursor.execute("""
                SELECT DESCRIPTION, AUTHORITY
                FROM PROPRIEDADES_IAM
                WHERE ID_ARQUIVO = ?
            """, (id_arquivo,))

            row = cursor.fetchone()

            erros = []

            if not row:
                erros.append(("SEM_PROPRIEDADES", None, None))
                return erros

            descricao, authority = row

            if not descricao:
                erros.append(("FALTA_DESCRICAO", "DESCRIPTION", None))

            if not authority:
                erros.append(("FALTA_AUTHORITY", "AUTHORITY", None))

            return erros

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def registrar_erro(self, cursor, id_arquivo, erro):

        tipo, campo, valor = erro

        try:
            cursor.execute("""
                INSERT INTO ERROS_VALIDACAO_ERP (
                    ID_ARQUIVO, TIPO, CAMPO, VALOR_ENCONTRADO
                )
                VALUES (?, ?, ?, ?)
            """, (id_arquivo, tipo, campo, valor))

        except Exception:
            # 🔥 ignora QUALQUER erro de insert (duplicado)
            # porque só pode acontecer duplicidade aqui
            pass

if __name__ == "__main__":
    ValidarEstruturaERP()