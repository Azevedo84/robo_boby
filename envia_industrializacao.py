import sys
from banco_dados.conexao import conecta, conecta_robo
from banco_dados.controle_erros import grava_erro_banco
from comandos.conversores import valores_para_float
import os
import inspect
import smtplib
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from dados_email import email_user, password


class EnviaIndustrializacao:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.manipula_comeco()

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

    def gerar_pdf_listagem_separar(self, caminho_listagem, lista_final):
        try:
            margem_esquerda = 0
            margem_direita = 5
            margem_superior = 25
            margem_inferior = 5

            doc = SimpleDocTemplate(caminho_listagem, pagesize=A4,
                                    leftMargin=margem_esquerda,
                                    rightMargin=margem_direita,
                                    topMargin=margem_superior,
                                    bottomMargin=margem_inferior)

            titulo = ['INDÚSTRIALIZAÇÃO']
            style_lista = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                                      ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                      ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                      ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                      ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                                      ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                      ('FONTSIZE', (0, 0), (-1, 0), 10),
                                      ('FONTSIZE', (0, 1), (-1, -1), 8)])

            table = Table([titulo])
            table.setStyle(style_lista)
            elements = [table]

            cabecalho_lista = ['CÓDIGO', 'DESCRIÇÃO', 'REFERÊNCIA', 'UM', 'LOCALIZAÇÃO', 'SALDO']
            elem_lista = self.adicionar_tabelas_listagem(lista_final, cabecalho_lista)

            cabecalho_transp = ['', 'TRANSPORTE']
            dados_transp = [('PESO LÍQUIDO', ''), ('PESO BRUTO', ''), ('VOLUME', '')]
            elem_transp = self.adicionar_tabelas_listagem(dados_transp, cabecalho_transp)

            cabecalho_medida = ['MEDIDAS', '            ']
            dados_medida = [('ALTURA (MM)', ''), ('LARGURA (MM)', ''), ('COMPRIMENTO (MM)', '')]
            elem_medida = self.adicionar_tabelas_listagem(dados_medida, cabecalho_medida)

            cabecalho_motorista = ['DAMDFE', 'MOTORISTA']
            dados_motorista = [('PLACA', ''), ('NOME', ''), ('CPF', '')]
            elem_motorista = self.adicionar_tabelas_listagem(dados_motorista, cabecalho_motorista)

            espaco_em_branco = Table([[None]], style=[('SIZE', (0, 0), (0, 0), 20)])

            # Criar tabela para colocar medidas e motorista lado a lado
            tabela_medida_motorista = Table([[elem_transp, elem_medida, elem_motorista]],
                                            colWidths=[170, 170])  # Ajuste as larguras conforme necessário

            elementos = (elements + [espaco_em_branco] + elem_lista + [espaco_em_branco] +
                         [tabela_medida_motorista])  # Adiciona a tabela com medidas e motorista lado a lado

            doc.build(elementos)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def inserir_no_banco(self, lista_banco):
        try:
            for i in lista_banco:
                id_mov, cod = i

                cursor = conecta_robo.cursor()
                cursor.execute(f"Insert into ENVIA_INDUSTRIALIZACAO (ID, id_envia_mov, cod_prod) "
                               f"values (GEN_ID(GEN_ENVIA_INDUSTRIALIZACAO_ID,1), {id_mov}, {cod});")

                conecta_robo.commit()

                print(f"Salvo no banco com sucesso!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email(self, caminho, arquivo):
        try:
            saudacao, msg_final, to = self.dados_email()

            to = ['<maquinas@unisold.com.br>']

            subject = f'IND - Separar Produtos para industrialização'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Alguns produtos estão prontos para enviar para industrialização.\n\n" \
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

            print(f'Email enviado com sucesso!')

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

    def manipula_comeco(self):
        try:
            lista_produtos = []

            lista_final = []
            lista_banco = []

            cursor = conecta.cursor()
            cursor.execute("""
                SELECT id, data_mov 
                FROM envia_mov 
                WHERE data_mov >= DATEADD(-1 MONTH TO CURRENT_DATE)
            """)
            dados_mov = cursor.fetchall()

            if dados_mov:
                for i in dados_mov:
                    id_mov, data_mov = i

                    print(i)

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                                   f"prod.unidade, prod.localizacao, prod.quantidade "
                                   f"FROM movimentacao AS mov "
                                   f"INNER JOIN produto prod ON mov.produto = prod.id "
                                   f"WHERE mov.data = '{data_mov}' and mov.tipo < 200;")
                    dados_mov = cursor.fetchall()

                    if dados_mov:
                        for ii in dados_mov:
                            cod, descr, ref, um, local, saldo = ii

                            prod_saldo_encontrado = False
                            for cod_sal_e, descr_e in lista_produtos:
                                if cod_sal_e == cod:
                                    prod_saldo_encontrado = True
                                    break

                            if not prod_saldo_encontrado:
                                saldo_float = valores_para_float(saldo)

                                if saldo_float > 0:
                                    dados_colhidos = self.manipula_dados_onde_usa(cod)
                                    if dados_colhidos:
                                        dados = (cod, descr)

                                        lista_produtos.append(dados)

                                        cur = conecta_robo.cursor()
                                        cur.execute(f"SELECT * from ENVIA_INDUSTRIALIZACAO "
                                                    f"where id_envia_mov = {id_mov} and cod_prod = {cod};")
                                        dados_salvos = cur.fetchall()

                                        if cod == "17814":
                                            print("bbbbb", dados_salvos)

                                        if not dados_salvos:
                                            dadoss = (cod, descr, ref, um, local, saldo)
                                            lista_final.append(dadoss)

                                            dadosss = (id_mov, cod)
                                            lista_banco.append(dadosss)

            if lista_final:
                arquivo = 'Listagem - Ind.pdf'
                caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), arquivo)

                self.gerar_pdf_listagem_separar(caminho, lista_final)
                self.inserir_no_banco(lista_banco)
                self.envia_email(caminho, arquivo)
                self.excluir_arquivo(arquivo)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_onde_usa(self, cod_prod):
        try:
            planilha_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, estprod.id_estrutura, estprod.quantidade "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f" where prod.codigo = {cod_prod};")
            tabela_estrutura = cursor.fetchall()
            for i in tabela_estrutura:
                ides_mat, id_estrutura, qtde = i

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo "
                               f"from produto "
                               f" where id_versao = {id_estrutura};")
                produto_pai = cursor.fetchall()
                if produto_pai:


                    cod_produto = produto_pai[0][1]

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, ''), prod.unidade, "
                                   f"COALESCE(prod.obs2, '') "
                                   f"from estrutura as est "
                                   f"INNER JOIN produto prod ON est.id_produto = prod.id "
                                   f"where prod.codigo = {cod_produto} and prod.tipomaterial = 119;")
                    select_prod = cursor.fetchall()

                    if select_prod:
                        cod, descr, ref, um, obs = select_prod[0]
                        dados = (cod, descr, ref, um, qtde)
                        planilha_nova.append(dados)

            if planilha_nova:
                planilha_nova_ordenada = sorted(planilha_nova, key=lambda x: x[1])

                return planilha_nova_ordenada

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaIndustrializacao()