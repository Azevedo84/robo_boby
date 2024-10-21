import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import traceback


class EnviaRequisicaoPendentes:
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

    def envia_email(self, numero, data, cod, descr, ref, um, qtde):
        try:
            to = ['<maquinas@unisold.com.br>']

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

            email_user = 'ti.ahcmaq@gmail.com'
            password = 'poswxhqkeaacblku'

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

            server.sendmail(email_user, to, text)

            server.quit()

            data_obj = datetime.strptime(data, '%d-%m-%Y')
            data_formatada = data_obj.strftime('%Y-%m-%d')

            cursor = conecta.cursor()
            cursor.execute(f"Insert into envia_req_falta (ID, num_req, data_emissao) "
                           f"values (GEN_ID(GEN_ENVIA_REQ_FALTA_ID,1), {numero}, '{data_formatada}');")
            conecta.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaRequisicaoPendentes()
chama_classe.manipula_dados()
