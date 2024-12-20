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


class EnviaPICriadas:
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

            email_user = 'ti.ahcmaq@gmail.com'
            password = 'poswxhqkeaacblku'

            return saudacao, msg_final, email_user, to, password

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email(self, dados):
        try:
            saudacao, msg_final, email_user, to, password = self.dados_email()

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

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaPICriadas()
chama_classe.manipula_comeco()
