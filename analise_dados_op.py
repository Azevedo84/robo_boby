import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import inspect
import os
import traceback

from comandos.excel import edita_fonte, criar_workbook, edita_preenchimento, edita_alinhamento
from comandos.excel import edita_bordas


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from datetime import datetime
from dados_email import email_user, password


class DadosOrdensDeProducao:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

        media_peca, media_conj_15, media_conj_mais = self.calculo_1_dados_ops_encerradas()
        self.calculo_2_dados_ops_abertas(media_peca, media_conj_15, media_conj_mais)

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

    def calculo_1_dados_ops_encerradas(self):
        try:
            ano_atual = datetime.now().year
            anos_filtrar = (ano_atual - 2, ano_atual - 1)

            ops_encerradas_peca = 0
            dias_ops_peca = 0

            ops_encerrada_conj_15 = 0
            dias_ops_conj_15 = 0

            ops_encerrada_conj_mais = 0
            dias_ops_conj_mais = 0

            cursor = conecta.cursor()
            cursor.execute(f"select op.datainicial, op.numero, op.codigo, prod.descricao, "
                           f"op.datafinal, tip.tipomaterial, prod.id_versao, op.status "
                           f"from ordemservico as op "
                           f"INNER JOIN produto as prod ON op.produto = prod.id "
                           f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                           f"WHERE EXTRACT(YEAR FROM op.datainicial) "
                           f"BETWEEN {anos_filtrar[0]} "
                           f"AND {anos_filtrar[1]} and op.status = 'B';")
            dados_op = cursor.fetchall()

            if dados_op:
                for i in dados_op:
                    data_ini, num_op, cod_prod, descr, data_fim, tipo, id_estrut, status = i

                    cursor.execute(
                        f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade "
                        f"FROM estrutura_produto as estprod "
                        f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                        f"WHERE estprod.id_estrutura = {id_estrut};")
                    dados_estrutura = cursor.fetchall()

                    qtde_itens_estrut = len(dados_estrutura)

                    if data_fim:
                        tempo_demorado_aberta = (data_fim - data_ini).days

                        if tipo == "USINAGEM":
                            ops_encerradas_peca += 1
                            dias_ops_peca += tempo_demorado_aberta
                        if tipo == "CONJUNTO":
                            if qtde_itens_estrut < 16:
                                ops_encerrada_conj_15 += 1
                                dias_ops_conj_15 += tempo_demorado_aberta
                            else:
                                ops_encerrada_conj_mais += 1
                                dias_ops_conj_mais += tempo_demorado_aberta

            media_peca = dias_ops_peca / ops_encerradas_peca
            media_conj_15 = dias_ops_conj_15 / ops_encerrada_conj_15
            media_conj_mais = dias_ops_conj_mais / ops_encerrada_conj_mais

            media_peca = round(media_peca, 0)
            media_conj_15 = round(media_conj_15, 0)
            media_conj_mais = round(media_conj_mais, 0)

            return media_peca, media_conj_15, media_conj_mais


        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_2_dados_ops_abertas(self, media_peca, media_conj_15, media_conj_mais):
        try:
            lista_temp = []
            lista_final = []

            data_atual = datetime.now().date()
            ano_atual = datetime.now().year
            anos_filtrar = (ano_atual - 2, ano_atual - 1)

            cursor = conecta.cursor()
            cursor.execute(f"select op.datainicial, op.numero, op.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                           f"op.datafinal, tip.tipomaterial, prod.id_versao, op.status "
                           f"from ordemservico as op "
                           f"INNER JOIN produto as prod ON op.produto = prod.id "
                           f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                           f"WHERE EXTRACT(YEAR FROM op.datainicial) "
                           f"BETWEEN {anos_filtrar[0]} "
                           f"AND {anos_filtrar[1]} and op.status = 'A';")
            dados_op = cursor.fetchall()

            if dados_op:
                for i in dados_op:
                    data_ini, num_op, cod_prod, descr, ref, data_fim, tipo, id_estrut, status = i

                    cursor.execute(
                        f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade "
                        f"FROM estrutura_produto as estprod "
                        f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                        f"WHERE estprod.id_estrutura = {id_estrut};")
                    dados_estrutura = cursor.fetchall()

                    qtde_itens_estrut = len(dados_estrutura)

                    if not data_fim:
                        tempo_demora = (data_atual - data_ini).days

                        if tipo == "USINAGEM":
                            if tempo_demora > media_peca:
                                dados = ("USINAGEM MAIOR", data_ini, num_op, cod_prod, descr, ref, tempo_demora)
                                lista_temp.append(dados)

                        if tipo == "CONJUNTO":
                            if qtde_itens_estrut < 16:
                                if tempo_demora > media_conj_15:
                                    dados = ("CONJUNTO 15-", data_ini, num_op, cod_prod, descr, ref, tempo_demora)
                                    lista_temp.append(dados)
                            else:
                                if tempo_demora > media_conj_mais:
                                    dados = ("CONJUNTO 15+", data_ini, num_op, cod_prod, descr, ref, tempo_demora)
                                    lista_temp.append(dados)

            if lista_temp:
                lista_final_ordenada = sorted(lista_temp, key=lambda x: (x[0], x[1]))

                for ii in lista_final_ordenada:
                    tipo, data_ini, num_op, cod_prod, descr, ref, demora = ii

                    data_formatada = data_ini.strftime("%d/%m/%Y")

                    dadus = (tipo, data_formatada, num_op, cod_prod, descr, ref, demora)
                    lista_final.append(dadus)

            self.excel(lista_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel(self, tabela_final):
        try:
            if tabela_final:
                caminho = fr'C:\Users\Anderson\PycharmProjects\robo_boby\Demora_OP.xlsx'
                arquivo = 'Demora_OP.xlsx'

                if tabela_final:
                    workbook = criar_workbook()
                    sheet = workbook.active
                    sheet.title = "Posição"

                    headers = ["Tipo", "Emissão", "Nº OP", "Código", "Descrição", "Referência", "Dias"]
                    sheet.append(headers)

                    header_row = sheet[1]
                    for cell in header_row:
                        edita_fonte(cell, negrito=True)
                        edita_preenchimento(cell)
                        edita_alinhamento(cell)

                    for d_ex in tabela_final:
                        tipo, data_ini, num_op, cod_prod, descr, ref, demora = d_ex

                        codigu = int(cod_prod)
                        op_int = int(num_op)

                        sheet.append([tipo, data_ini, op_int, codigu, descr, ref, demora])

                    for row in sheet.iter_rows(min_row=1,
                                               max_row=sheet.max_row,
                                               min_col=1,
                                               max_col=sheet.max_column):
                        for cell in row:
                            edita_bordas(cell)
                            edita_alinhamento(cell)

                    workbook.save(caminho)

                    print("Excel Salvo!")
                    self.envia_email(caminho, arquivo)
                    self.excluir_arquivo(caminho)

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

    def envia_email(self, caminho, arquivo):
        try:
            saudacao, msg_final, to = self.dados_email()

            to = ['<maquinas@unisold.com.br>']

            subject = f'Relatório Demora OPs'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Segue relatório diário com a posição da produção.\n\n" \
                   f"{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            diretorio_atual = os.path.dirname(os.path.abspath(__file__))

            caminho_arquivo = os.path.join(diretorio_atual, caminho)

            attachment = open(caminho_arquivo, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment',
                            filename=Header(arquivo, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'Email com relátorio enviado com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = DadosOrdensDeProducao()