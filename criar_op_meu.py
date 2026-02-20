import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
from pdf2image import convert_from_path
from PIL import ImageFont, ImageDraw
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import time
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from datetime import datetime, date, timedelta
import traceback
from dados_email import email_user, password


class EnviaOrdensProducao:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

        self.caminho_original = ""
        self.arq_original = ""
        self.num_desenho_arq = ""
        self.qtde_produto = 0
        self.cod_prod = ""
        self.descr_prod = ""
        self.ref_prod = ""
        self.um_prod = ""
        self.num_op = ""
        self.tipo = ""

        self.data_emissao = ""
        self.data_entrega = ""

        self.caminho_poppler = r'C:\Program Files\poppler-24.08.0\Library\bin'

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
            return None

    def converte_png_para_pdf(self, input_png, output_pdf):
        try:
            c = canvas.Canvas(output_pdf, pagesize=landscape(A4))
            img = ImageReader(input_png)
            width, height = img.getSize()

            aspect_ratio = width / height
            target_width = 800
            target_height = target_width / aspect_ratio

            x = (A4[1] - target_width) / 2
            y = (A4[0] - target_height) / 2

            c.drawImage(img, x, y, width=target_width, height=target_height)
            c.showPage()
            c.save()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def cria_imagem_do_desenho_usinagem(self):
        try:
            num_op_str = f"OP {self.num_op}"
            if self.qtde_produto > 1:
                qtde_ordem = f"{self.qtde_produto} PÇS"
            else:
                qtde_ordem = f"{self.qtde_produto} PÇ"

            data_emissao = f"Emissão: {self.data_emissao}"

            images = convert_from_path(self.caminho_original, 500, poppler_path=self.caminho_poppler)

            imgs = images[0]

            draw = ImageDraw.Draw(imgs)
            font = ImageFont.truetype("tahoma.ttf", 150)
            font1 = ImageFont.truetype("tahoma.ttf", 70)

            def criar_texto(pos_horizontal, pos_vertical, texto, cor, fonte, largura_tra):
                draw.text((pos_horizontal, pos_vertical), texto, fill=cor, font=fonte, stroke_width=largura_tra)

            criar_texto(4500, 2900, num_op_str, (0, 0, 0), font, 4)
            criar_texto(5160, 2900, qtde_ordem, (0, 0, 0), font, 4)
            criar_texto(5000, 3150, data_emissao, (0, 0, 0), font1, 0)

            arquivo_final = f"{self.num_desenho_arq}.png"
            imgs.save(arquivo_final)

            time.sleep(1)

            return arquivo_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def cria_imagem_do_desenho_conjunto(self):
        try:
            num_op_str = f"OP {self.num_op}"
            if self.qtde_produto > 1:
                qtde_ordem = f"{self.qtde_produto} PÇS"
            else:
                qtde_ordem = f"{self.qtde_produto} PÇ"

            data_emissao = f"Emissão: {self.data_emissao}"

            images = convert_from_path(self.caminho_original, 500, poppler_path=self.caminho_poppler)

            imgs = images[0]

            draw = ImageDraw.Draw(imgs)
            font = ImageFont.truetype("tahoma.ttf", 150)
            font1 = ImageFont.truetype("tahoma.ttf", 70)

            def criar_texto(pos_horizontal, pos_vertical, texto, cor, fonte, largura_tra):
                draw.text((pos_horizontal, pos_vertical), texto, fill=cor, font=fonte, stroke_width=largura_tra)

            criar_texto(4500, 2830, num_op_str, (0, 0, 0), font, 4)
            criar_texto(5160, 2830, qtde_ordem, (0, 0, 0), font, 4)
            criar_texto(5000, 3080, data_emissao, (0, 0, 0), font1, 0)

            arquivo_final = f"{self.num_desenho_arq}.png"
            imgs.save(arquivo_final)

            time.sleep(1)

            return arquivo_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            return None

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

    def cria_pdf_envia_email_usinagem(self):
        try:
            num_op = str(self.num_op)

            arquivo_imagem = self.cria_imagem_do_desenho_usinagem()

            diretorio_destino = r'\\publico\C\OP\Aguardando Material/'
            arquivo_pdf_final = f'OP {num_op} - {self.num_desenho_arq}.pdf'
            caminho_completo = os.path.join(diretorio_destino, arquivo_pdf_final)

            self.converte_png_para_pdf(arquivo_imagem, caminho_completo)

            self.excluir_arquivo(arquivo_imagem)
            self.excluir_arquivo(self.caminho_original)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def cria_pdf_envia_email_conjunto(self):
        try:
            arquivo_imagem = self.cria_imagem_do_desenho_conjunto()

            diretorio_destino = r'\\publico\C\OP\Aguardando Material/'
            arquivo_pdf_final = f'OP {self.num_op} - {self.num_desenho_arq}.pdf'
            caminho_completo = os.path.join(diretorio_destino, arquivo_pdf_final)

            self.converte_png_para_pdf(arquivo_imagem, caminho_completo)

            self.excluir_arquivo(arquivo_imagem)
            print("este é o arquivo pra excluir", self.caminho_original)
            self.excluir_arquivo(self.caminho_original)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def estrutura_prod_qtde_op(self, id_produto):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where id = {id_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod, id_versao = select_prod[0]

            if id_versao:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                               f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                               f"((SELECT quantidade FROM ordemservico where numero = {self.num_op}) * "
                               f"(mat.quantidade)) AS Qtde, prod.quantidade "
                               f"FROM estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"where estprod.id_estrutura = {id_versao} ORDER BY prod.descricao;")
                sel_estrutura = cursor.fetchall()

                return sel_estrutura

            return []

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def envia_email_nao_acha_desenho(self):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'OP - Não foi encontrado o desenho {self.num_desenho_arq}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"O desenho {self.num_desenho_arq} não foi encontrado no cadastro dos produtos.\n\n" \
                   f"{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(self.caminho_original, 'rb')

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

            print(f'Desenho {self.num_desenho_arq} enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_tipo_nao_cadastrado(self):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'OP - O produto {self.cod_prod} não tem o "Tipo de Material" definido no cadsatro'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Favor definir o Tipo de Material no cadastro do produto: {self.cod_prod}.\n\n" \
                   f"{msg_final}"

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

            print(f'produto sem tipo {self.cod_prod} enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_desenho_duplicado(self, lista_produtos):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'OP - Foi encontrado produtos com desenho duplicado no cadastro!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Segue lista dos itens encontrados:\n\n"

            for i in lista_produtos:
                body += f"{i}.\n" \

            body += f"\n\n{msg_final}"

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

            print(f'produto desenho duplicado enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_sem_estrutura(self, lista_produtos):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'OP - Produto não possui estrutura!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Segue lista dos itens encontrados:\n\n"

            for i in lista_produtos:
                body += f"{i}.\n" \

            body += f"\n\n{msg_final}"

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

            print(f'produto sem estrutura enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_produto_ops_envio(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"select id, numero from ordemservico "
                           f"where numero = (select max(numero) from ordemservico);")
            select_numero = cursor.fetchall()
            idez, num = select_numero[0]
            self.num_op = int(num) + 1

            situacao = self.criar_op()

            if situacao:
                if self.tipo == "87":
                    print("CONJUNTO")
                    self.cria_pdf_envia_email_conjunto()
                else:
                    print("USINAGEM")

                    self.cria_pdf_envia_email_usinagem()

                print("FINALIZADO")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def criar_op(self):
        try:
            situacao = False

            emissao_certo = date.today()
            self.data_emissao = emissao_certo.strftime('%d/%m/%Y')

            previsao = emissao_certo + timedelta(weeks=4)
            self.data_entrega = previsao.strftime('%d/%m/%Y')

            cod_barras = "SUZ000" + str(self.num_op)

            cur = conecta.cursor()
            cur.execute(f"SELECT id, descricao, COALESCE(obs, ' ') as obs, unidade, id_versao "
                        f"FROM produto where codigo = {self.cod_prod};")
            detalhes_produto = cur.fetchall()
            id_prod, descricao_id, referencia_id, unidade_id, id_versao = detalhes_produto[0]

            id_prod_int = int(id_prod)

            if id_versao:
                obs_certo = "OP CRIADA PELO SERVIDOR"

                cursor = conecta.cursor()
                cursor.execute(f"Insert into ordemservico "
                               f"(id, produto, numero, quantidade, datainicial, obs, codbarras, status, codigo, "
                               f"dataprevisao, id_estrutura, etapa) "
                               f"values (GEN_ID(GEN_ORDEMSERVICO_ID,1), {id_prod_int}, {self.num_op}, "
                               f"'{self.qtde_produto}', '{emissao_certo}', '{obs_certo}', '{cod_barras}', 'A', "
                               f"'{self.cod_prod}', '{previsao}', {id_versao}, 'ABERTA');")

                conecta.commit()

                situacao = True

                print(f'A Ordem de Produção Nº {self.num_op} foi criado com sucesso!')

            return situacao

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def procura_produto_pelo_desenho(self):
        try:
            num_des_com_d = f"D {self.num_desenho_arq}"

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT DISTINCT codigo, descricao, COALESCE(obs, ''), unidade, COALESCE(tipomaterial, ''), "
                f"COALESCE(localizacao, ''), id_versao "
                f"FROM produto "
                f"WHERE obs = '{num_des_com_d}';")
            detalhes_produto = cursor.fetchall()

            if detalhes_produto:
                qtde_itens = len(detalhes_produto)
                if qtde_itens == 1:
                    for i in detalhes_produto:
                        self.cod_prod = i[0]
                        self.tipo = i[4]
                        id_versao = i[6]

                        if id_versao:
                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT * "
                                           f"FROM estrutura_produto "
                                           f"where id_estrutura = {id_versao};")
                            select_estrut = cursor.fetchall()
                            if select_estrut:
                                if self.tipo:
                                    self.descr_prod = i[1]
                                    self.ref_prod = i[2]
                                    self.um_prod = i[3]

                                    self.manipula_produto_ops_envio()
                                else:
                                    self.envia_email_tipo_nao_cadastrado()
                            else:
                                self.envia_email_sem_estrutura(detalhes_produto)
                        else:
                            self.envia_email_sem_estrutura(detalhes_produto)
                else:
                    print("tem mais itens", qtde_itens)
                    self.envia_email_desenho_duplicado(detalhes_produto)
            else:
                print("desenho não encontrado")
                self.envia_email_nao_acha_desenho()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def procura_arquivos(self):
        try:
            caminho = r'\\publico\C\OP\\'
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
            return None

    def manipula_comeco(self):
        try:
            arquivos_pdfs = self.procura_arquivos()

            for arqsv in arquivos_pdfs:
                self.caminho_original = arqsv
                print("self.caminho_original", self.caminho_original)

                self.arq_original = self.caminho_original[16:]

                if " - " in self.arq_original:
                    inicio = self.arq_original.find(" - ")
                    dadinhos1 = self.arq_original[inicio + 3:]
                elif "-" in self.arq_original:
                    inicio = self.arq_original.find("-")
                    dadinhos1 = self.arq_original[inicio + 1:]
                else:
                    inicio = ""
                    dadinhos1 = ""

                if inicio and dadinhos1:
                    qtde = self.arq_original[:inicio]
                    try:
                        self.qtde_produto = int(qtde)

                        self.num_desenho_arq = dadinhos1[:-4]

                        print("Arquivo Original:", self.caminho_original)
                        print("N Desenho:", self.num_desenho_arq)
                        print("Quantidade:", self.qtde_produto)

                        self.procura_produto_pelo_desenho()

                    except ValueError:
                        print("A QUANTIDADE NÃO ESTÁ CERTA")

                else:
                    print('O ARQUIVO NÃO ESTÁ NO PADRÃO CERTO: "QTDE - NUM. DESENHO"')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaOrdensProducao()
chama_classe.manipula_comeco()
