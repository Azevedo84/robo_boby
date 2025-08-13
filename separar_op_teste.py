import sys
from banco_dados.conexao import conecta, conecta_robo
from banco_dados.controle_erros import grava_erro_banco
from comandos.conversores import valores_para_float
import os
import traceback
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from datetime import datetime

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from dados_email import email_user, password


class EnviaOrdensProducao:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.caminho_original = ""
        self.arq_original = ""
        self.num_desenho_arq = ""
        self.qtde_produto = 0
        self.id_estrut = ""
        self.cod_prod = ""
        self.descr_prod = ""
        self.ref_prod = ""
        self.um_prod = ""
        self.num_op = ""
        self.tipo = ""

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

    def excluir_arquivo(self, caminho_arquivo):
        try:
            if os.path.exists(caminho_arquivo):
                os.remove(caminho_arquivo)
            else:
                print("O arquivo não existe no caminho especificado.")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def listagem_material_estrutura(self, cod_prod, qtde):
        try:
            lista_local = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {cod_prod};")
            select_prod = cursor.fetchall()

            idez, cod, id_estrut = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"conj.conjunto, prod.unidade, prod.localizacao, "
                           f"(estprod.quantidade * {qtde}) as qtde, "
                           f"prod.quantidade "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"where estprod.id_estrutura = {id_estrut} "
                           f"order by conj.conjunto DESC, prod.descricao ASC;")
            tabela_estrutura = cursor.fetchall()

            if tabela_estrutura:
                for i in tabela_estrutura:
                    cod, descr, ref, conj, um, local, qtde, saldo = i
                    dados = (cod, descr, ref, um, qtde, local, saldo)
                    lista_local.append(dados)

            return lista_local

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def estrutura_prod_qtde_op(self, num_op, id_estrut):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                           f"(estprod.quantidade)) AS Qtde, prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} ORDER BY prod.descricao;")
            sel_estrutura = cursor.fetchall()

            return sel_estrutura

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def teste(self, id_mat_e, lista_substitutos, num_op):
        try:
            saldo_substituto = 0

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id = {id_mat_e};")
            sel_estrutura = cursor.fetchall()

            cod_original = sel_estrutura[0][1]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT subs.cod_subs, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"conj.conjunto, prod.unidade, prod.localizacao, prod.quantidade "
                           f"FROM SUBSTITUTO_MATERIAPRIMA as subs "
                           f"INNER JOIN produto as prod ON subs.cod_subs = prod.codigo "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"WHERE subs.id_mat = {id_mat_e} "
                           f"and prod.quantidade > 0 "
                           f"AND subs.num_op = {num_op};")
            dados_mat_com_op = cursor.fetchall()
            if dados_mat_com_op:
                cod_subs, descr, ref, conj, um, local, saldo = dados_mat_com_op[0]

                dados = (cod_subs, descr, ref, um, local, saldo, cod_original)
                lista_substitutos.append(dados)

                saldo_substituto = float(saldo)

            else:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT subs.cod_subs, prod.descricao, COALESCE(prod.obs, '') as obs, "
                               f"conj.conjunto, prod.unidade, prod.localizacao, prod.quantidade "
                               f"FROM SUBSTITUTO_MATERIAPRIMA as subs "
                               f"INNER JOIN produto as prod ON subs.cod_subs = prod.codigo "
                               f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                               f"WHERE subs.id_mat = {id_mat_e} "
                               f"and prod.quantidade > 0 "
                               f"AND subs.num_op is NULL;")
                dados_mat_sem_op = cursor.fetchall()
                if dados_mat_sem_op:
                    cod_subs, descr, ref, conj, um, local, saldo = dados_mat_sem_op[0]

                    dados = (cod_subs, descr, ref, um, local, saldo, cod_original)
                    lista_substitutos.append(dados)

                    saldo_substituto = float(saldo)

            return lista_substitutos, saldo_substituto

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consumo_op_por_id(self, num_op, id_materia_prima):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"select prodser.id_estrut_prod, prodser.data, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"prodser.quantidade, prodser.QTDE_ESTRUT_PROD "
                           f"from produtoos as prodser "
                           f"INNER JOIN produto as prod ON prodser.produto = prod.id "
                           f"where prodser.numero = {num_op} and prodser.id_estrut_prod = {id_materia_prima};")
            consumo_os = cursor.fetchall()

            return consumo_os

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_ops_concluidas(self):
        try:
            lista_substitutos = []

            select_estrut = self.estrutura_prod_qtde_op(self.num_op, self.id_estrut)

            todos_materiais_consumidos = True

            for dados_estrut in select_estrut:
                saldo_final = 0

                id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_e, saldo_prod_e = dados_estrut

                lista_substitutos, saldo_subs = self.teste(id_mat_e, lista_substitutos, self.num_op)

                qtde_e_float = valores_para_float(qtde_e)
                saldo_final += saldo_subs + float(saldo_prod_e)

                select_os = self.consumo_op_por_id(self.num_op, id_mat_e)
                if select_os:
                    qtde_itens = len(select_os)

                    if qtde_itens == 1:
                        for dados_os in select_os:
                            id_mat_os, data_os, cod_os, descr_os, ref_os, um_os, qtde_os, qtde_mat_os = dados_os

                            qtde_mat_os_float = valores_para_float(qtde_mat_os)

                            if qtde_mat_os_float < qtde_e_float:
                                sobras = qtde_e_float - qtde_mat_os_float
                                if sobras > saldo_final:
                                    todos_materiais_consumidos = False
                                    break
                    else:
                        qtde_consumo_op_final = 0

                        for dados_os in select_os:
                            id_mat_os, data_os, cod_os, descr_os, ref_os, um_os, qtde_os, qtde_mat_os = dados_os

                            qtde_mat_os_float = valores_para_float(qtde_mat_os)
                            qtde_consumo_op_final += qtde_mat_os_float

                        if qtde_consumo_op_final < qtde_e_float:
                            sobras = qtde_e_float - qtde_consumo_op_final
                            if sobras > saldo_final:
                                todos_materiais_consumidos = False
                                break
                else:
                    if saldo_final < qtde_e_float:
                        todos_materiais_consumidos = False
                        break

            return todos_materiais_consumidos, lista_substitutos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_usinagem(self):
        try:
            saudacao, msg_final, to = self.dados_email()

            to = ['<maquinas@unisold.com.br>']

            subject = f'OP - Separar material OP {self.num_op}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"A Ordem de Produção Nº {self.num_op} está com todo material em estoque para ser separado.\n\n" \
                   f"{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            diretorio_atual = os.path.dirname(os.path.abspath(__file__))

            caminho_arquivo = os.path.join(diretorio_atual, self.caminho_original)

            attachment = open(caminho_arquivo, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=Header(self.arq_original, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'OP {self.num_op} enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_conjunto(self, caminho_listagem, arquivo_listagem):
        try:
            saudacao, msg_final, to = self.dados_email()

            to = ['<maquinas@unisold.com.br>']

            subject = f'Separar material {self.num_op}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n"

            msg.attach(MIMEText(body, 'plain'))

            attachment2 = open(caminho_listagem, 'rb')
            part2 = MIMEBase('application', "octet-stream")
            part2.set_payload(attachment2.read())
            encoders.encode_base64(part2)
            part2.add_header('Content-Disposition', 'attachment', filename=Header(arquivo_listagem, 'utf-8').encode())
            msg.attach(part2)
            attachment2.close()

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)

            server.quit()

            print(f'OP {self.num_op} enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_nao_acha_op(self):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'OP - Não foi encontrado o número da OP'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f'{saudacao}\n\n' \
                   f'Houve algum problema com o arquivo "{self.arq_original}".\n\n' \
                   f'{msg_final}'

            msg.attach(MIMEText(body, 'plain'))

            diretorio_atual = os.path.dirname(os.path.abspath(__file__))

            caminho_arquivo = os.path.join(diretorio_atual, self.caminho_original)

            attachment = open(caminho_arquivo, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=Header(self.arq_original, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'Email do erro da OP foi enviado com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_op_nao_existe(self):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'OP - A OP {self.num_op} não existe!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f'{saudacao}\n\n' \
                   f'Houve algum problema com o arquivo "{self.arq_original}", pois esta OP não existe!\n\n' \
                   f'{msg_final}'

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(self.caminho_original, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment',
                            filename=Header(self.arq_original, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'Email do erro da OP foi enviado com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_op_encerrada(self):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'OP - A OP {self.num_op} já foi encerrada!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f'{saudacao}\n\n' \
                   f'Houve algum problema com o arquivo "{self.arq_original}", pois esta OP já foi encerrada!\n\n' \
                   f'{msg_final}'

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(self.caminho_original, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment',
                            filename=Header(self.arq_original, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'Email do erro da OP foi enviado com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_numeros_duplicados(self):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'OP - Arquivos Duplicados na pasta "OP"'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f'{saudacao}\n\n' \
                   f'Favor verificar os arquivos da pasta "OP" no servidor (PUBLICO). Pois foi encontrado ' \
                   f'itens semelhantes.\n\n' \
                   f'{msg_final}'

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)

            server.quit()

            print(f'Email do erro da OP foi enviado com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def adicionar_tabelas_listagem(self, dados, cabecalho):
        try:
            elements = []

            style_lista = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                                      ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                      ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                      ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                      ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                                      ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                      ('FONTSIZE', (0, 0), (-1, 0), 10),
                                      ('FONTSIZE', (0, 1), (-1, -1), 8)])

            table = Table([cabecalho] + dados)
            table.setStyle(style_lista)
            elements.append(table)

            return elements

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def gerar_pdf_listagem_separar(self, caminho_listagem, lista_subs):
        try:
            lista_estrutura = self.listagem_material_estrutura(self.cod_prod, self.qtde_produto)

            dados_op = [(self.num_op, self.cod_prod, self.descr_prod, self.ref_prod,
                         self.um_prod, self.qtde_produto)]

            margem_esquerda = 0
            margem_direita = 5
            margem_superior = 25
            margem_inferior = 5

            doc = SimpleDocTemplate(caminho_listagem, pagesize=A4,
                                    leftMargin=margem_esquerda,
                                    rightMargin=margem_direita,
                                    topMargin=margem_superior,
                                    bottomMargin=margem_inferior)

            cabecalho_op = ['Nº OP', 'CÓDIGO', 'DESCRIÇÃO', 'REFERÊNCIA', 'UM', 'QTDE']
            elementos_op = self.adicionar_tabelas_listagem(dados_op, cabecalho_op)

            cabecalho_lista = ['CÓDIGO', 'DESCRIÇÃO', 'REFERÊNCIA', 'UM', 'QTDE', 'LOCALIZAÇÃO', 'SALDO']
            elementos_lista = self.adicionar_tabelas_listagem(lista_estrutura, cabecalho_lista)

            espaco_em_branco = Table([[None]], style=[('SIZE', (0, 0), (0, 0), 20)])

            if lista_subs:
                styles = getSampleStyleSheet()
                estilo_centralizado = ParagraphStyle(name='Centralizado', parent=styles['Heading2'], alignment=1)
                texto_substituicoes = Paragraph("Lista de Opções para Substituição", estilo_centralizado)

                cabecalho_subs = ['CÓDIGO', 'DESCRIÇÃO', 'REFERÊNCIA', 'UM', 'LOCALIZAÇÃO', 'SALDO', 'CÓD. ORIGINAL']
                elementos_subs = self.adicionar_tabelas_listagem(lista_subs, cabecalho_subs)
                elementos = elementos_op + [espaco_em_branco] + elementos_lista + \
                            [espaco_em_branco, texto_substituicoes, espaco_em_branco] + elementos_subs
            else:
                elementos = elementos_op + [espaco_em_branco] + elementos_lista

            doc.build(elementos)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_banco(self):
        try:
            cursor = conecta_robo.cursor()
            cursor.execute(f"SELECT * FROM envia_ops_prontas where num_op = {self.num_op};")
            select_envio = cursor.fetchall()

            return select_envio

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def procura_arquivos(self):
        try:
            caminho = r'\\publico\C\OP\Aguardando Material/'
            extensao = ".pdf"

            arquivos_pdfs = []

            for arquivo in os.listdir(caminho):
                if arquivo.endswith(extensao):
                    caminho_arquivo = os.path.join(caminho, arquivo)
                    arquivos_pdfs.append(caminho_arquivo)

            return arquivos_pdfs

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_comeco(self):
        try:
            lista = ["8095", "8096",]

            for i in lista:
                self.num_op = i
                print("self.num_op", self.num_op)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT numero, datainicial, status, produto, quantidade "
                               f"FROM ordemservico where numero = {self.num_op};")
                extrair_dados = cursor.fetchall()
                if not extrair_dados:
                    self.envia_email_op_nao_existe()
                else:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT op.numero, op.codigo, op.id_estrutura, prod.descricao, "
                                   f"COALESCE(prod.obs, ''), "
                                   f"prod.unidade, COALESCE(prod.tipomaterial, ''), op.quantidade "
                                   f"FROM ordemservico as op "
                                   f"INNER JOIN produto as prod ON op.produto = prod.id "
                                   f"where op.numero = {self.num_op} "
                                   f"AND op.status = 'A';")
                    select_status = cursor.fetchall()

                    if not select_status:
                        self.envia_email_op_encerrada()
                    else:
                        self.id_estrut = select_status[0][2]
                        self.cod_prod = select_status[0][1]
                        self.descr_prod = select_status[0][3]
                        self.ref_prod = select_status[0][4]
                        self.um_prod = select_status[0][5]
                        self.qtde_produto = select_status[0][7]
                        self.tipo = select_status[0][6]

                        todo_material_consumido, lista_subs = self.verifica_ops_concluidas()

                        diretorio_destino = r'\\publico\C\OP\Aguardando Material/'
                        arquivo_listagem = f'Listagem - OP {self.num_op}.pdf'
                        caminho_listagem = os.path.join(diretorio_destino, arquivo_listagem)

                        print("CONJUNTO")
                        self.gerar_pdf_listagem_separar(caminho_listagem, lista_subs)
                        self.envia_email_conjunto(caminho_listagem, arquivo_listagem)
                        self.excluir_arquivo(caminho_listagem)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaOrdensProducao()
chama_classe.manipula_comeco()
