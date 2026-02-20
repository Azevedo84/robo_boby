import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect
from datetime import datetime, date, timedelta

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from dados_email import email_user, password


class VencimentoPI:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

        dados_vence, dados_vencidos = self.inicio_de_tudo()

        if dados_vence:
            tab_ordenada = sorted(dados_vence, key=lambda x: x[6])
            self.envia_email_vencendo(tab_ordenada)

        if dados_vencidos:
            tab_ord_vencido = sorted(dados_vencidos, key=lambda x: x[6])
            self.envia_email_vencido(tab_ord_vencido)

    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

            # 'Log' em arquivo local apenas se houver erro
            with open(self.arquivo_log, "a", encoding="utf-8") as f:
                f.write(f"Erro na função {nome_funcao} do arquivo {arquivo}: {mensagem} (linha {num_linha_erro})\n")

        except Exception as e:
            nome_funcao_trat = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            tb = traceback.extract_tb(exc_traceback)
            num_linha_erro = tb[-1][1]
            print(f'Houve um problema no arquivo: {self.nome_arquivo} na função: "{nome_funcao_trat}"\n'
                  f'{e} {num_linha_erro}')
            grava_erro_banco(nome_funcao_trat, e, self.nome_arquivo, num_linha_erro)

            with open(self.arquivo_log, "a", encoding="utf-8") as f:
                f.write(
                    f"Erro na função {nome_funcao_trat} do arquivo {self.nome_arquivo}: {e} (linha {num_linha_erro})\n")

    def dados_email(self):
        try:
            to = ['<maquinas@unisold.com.br>', '<ahcmaquinas@gmail.com>']

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

    def envia_email_vencendo(self, lista_produtos):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'PI - Pedido Interno com Prazo Vencendo!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de Pedidos Vencendo:\n\n"

            for i in lista_produtos:
                num_ped, cod, descr, ref, um, qtde, entrega, data_necessidade = i

                entrega_br = entrega.strftime('%d/%m/%Y')

                body += f"- Nº Pedido Interno: {num_ped}\n" \
                        f"- Código: {cod} - Descrição: {descr} - {um} - {ref} - Qtde: {qtde}\n" \
                        f"- Data Entrega: {entrega_br}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print(f'EMAIL PEDIDOS VENCENDO ENVIADO COM SUCESSO!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_vencido(self, lista_produtos):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'PI - Pedido Interno com Prazo Vencido!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de Pedidos Vencidos:\n\n"

            for i in lista_produtos:
                num_ped, cod, descr, ref, um, qtde, entrega, data_necessidade = i

                entrega_br = entrega.strftime('%d/%m/%Y')

                body += f"- Nº Pedido Interno: {num_ped}\n" \
                        f"- Código: {cod} - Descrição: {descr} - {um} - {ref} - Qtde: {qtde}\n" \
                        f"- Data Entrega: {entrega_br}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print(f'EMAIL PEDIDOS VENCIDOS ENVIADO COM SUCESSO!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def inicio_de_tudo(self):
        try:
            lista_vencendo = []
            lista_vencido = []

            data_hoje = date.today()
            data_necessidade = data_hoje + timedelta(days=30)

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

                    if entrega < data_hoje:
                        dados = (num_ped, cod, descr, ref, um, qtde, entrega, data_necessidade)
                        lista_vencido.append(dados)

                    elif entrega < data_necessidade:
                        dados = (num_ped, cod, descr, ref, um, qtde, entrega, data_necessidade)
                        lista_vencendo.append(dados)

            return lista_vencendo, lista_vencido

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = VencimentoPI()