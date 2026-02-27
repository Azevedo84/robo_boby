import sys
from banco_dados.conexao import conecta, conecta_robo
from banco_dados.controle_erros import grava_erro_banco
import os
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import traceback
from dados_email import email_user, password


class EnviaErrosERP:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

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

    def manipula_comeco(self):
        try:
            tabela_final = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ID, NOME_PC, ARQUIVO, FUNCAO, MENSAGEM, ENTREGUE, CRIACAO "
                           f"FROM ZZZ_ERROS "
                           f"WHERE (entregue IS NULL OR entregue = '');")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    id_erro, nome_pc, arquivo, funcao, mensagem, entregue, criacao = i

                    cursor = conecta_robo.cursor()
                    cursor.execute(f"SELECT ID, ID_ERRO FROM ENVIA_ERROS_ERP "
                                   f"where ID_ERRO = {id_erro};")
                    dados_criados = cursor.fetchall()

                    if not dados_criados:
                        cursor = conecta_robo.cursor()
                        cursor.execute(f"Insert into ENVIA_ERROS_ERP (ID, ID_ERRO) "
                                       f"values (GEN_ID(GEN_ENVIA_ERROS_ERP_ID,1), {id_erro});")

                        cursor = conecta.cursor()
                        cursor.execute(f"UPDATE ZZZ_ERROS SET ENTREGUE = 'S' "
                                       f"WHERE id= {id_erro};")

                        conecta_robo.commit()

                        dados = (id_erro, nome_pc, arquivo, funcao, mensagem, entregue, criacao)
                        tabela_final.append(dados)

                if tabela_final:
                    self.envia_email(tabela_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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

            msg_final = f"Att,\n" \
                        f"Suzuki Máquinas Ltda\n" \
                        f"Fone (51) 3561.2583/(51) 3170.0965\n\n" \
                        f"Mensagem enviada automaticamente, por favor não responda.\n\n" \
                        f"Se houver algum problema com o recebimento de emails ou conflitos com o arquivo excel, " \
                        f"favor entrar em contato pelo email maquinas@unisold.com.br.\n\n"

            return saudacao, msg_final, to

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email(self, dados):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'ERROS - PROBLEMAS COM O ERP - SUZUKI!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de erros do ERP - Suzuki:\n\n"

            for i in dados:
                id_erro, nome_pc, arquivo, funcao, mensagem, entregue, criacao = i

                data_criacao = criacao.strftime('%d/%m/%Y %H:%M:%S')

                body += f"- Nome Computador: {nome_pc}\n" \
                        f"- Data Criação: {data_criacao}\n" \
                        f"- Nome do Arquivo: {arquivo}\n" \
                        f"- Nome da Função: {funcao}\n" \
                        f"- Mensagem de Erro: {mensagem}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaErrosERP()
chama_classe.manipula_comeco()
