import sys
from banco_dados.conexao import conecta, conecta_robo
from banco_dados.controle_erros import grava_erro_banco
import os
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, date
import traceback
from dados_email import email_user, password


class EnviaPIProntas:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

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

    def mensagem_email(self):
        try:
            current_time = (datetime.now())
            horario = current_time.strftime('%H')
            hora_int = int(horario)
            saudacao = "teste"
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

            to = ['<maquinas@unisold.com.br>']

            return saudacao, msg_final, to

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email(self, lista_produtos):
        try:
            saudacao, msg_final, to = self.mensagem_email()

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

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaPIProntas()
chama_classe.verifica_banco()
