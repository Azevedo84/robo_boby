from core.banco import conecta, conecta_robo
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from core.erros import trata_excecao
from core.email_service import dados_email


class SepararOVS:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

        self.manipula_comeco()

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
            trata_excecao(e)
            raise

    def gerar_pdf_listagem_separar(self, caminho_listagem, dados_cliente, produtos_ov):
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

            cabecalho_op = ['DATA', 'Nº OV', 'CLIENTE']
            elem_op = self.adicionar_tabelas_listagem(dados_cliente, cabecalho_op)

            cabecalho_lista = ['CÓDIGO', 'DESCRIÇÃO', 'REFERÊNCIA', 'UM', 'QTDE', 'LOCALIZAÇÃO', 'SALDO']
            elem_lista = self.adicionar_tabelas_listagem(produtos_ov, cabecalho_lista)

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

            elementos = (elem_op + [espaco_em_branco] +
                         elem_lista + [espaco_em_branco] +
                         [tabela_medida_motorista])  # Adiciona a tabela com medidas e motorista lado a lado

            doc.build(elementos)

        except Exception as e:
            trata_excecao(e)
            raise

    def inserir_no_banco(self, dados_cliente, num_ov):
        try:
            id_cliente = dados_cliente[0][3]

            cursor = conecta_robo.cursor()
            cursor.execute(f"Insert into SEPARAR_OVS (ID, NUM_OV, CLIENTE_ID) "
                           f"values (GEN_ID(GEN_SEPARAR_OVS_ID,1), {num_ov}, {id_cliente});")

            conecta_robo.commit()

            print(f"OV {num_ov} salvo no banco com sucesso!")

        except Exception as e:
            trata_excecao(e)
            raise

    def envia_email(self, num_ov, caminho, arquivo):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            to = ['<maquinas@unisold.com.br>']

            subject = f'OV - Separar material OV {num_ov}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"A Ordem de Venda Nº {num_ov} está com todo material em estoque para ser separado.\n\n" \
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

            server.sendmail(email_user, self.destinatario, text)
            attachment.close()

            server.quit()

            print(f'OV {num_ov} enviada com sucesso!')

        except Exception as e:
            trata_excecao(e)
            raise

    def excluir_arquivo(self, caminho_arquivo):
        try:
            if os.path.exists(caminho_arquivo):
                os.remove(caminho_arquivo)
            else:
                print("O arquivo não existe no caminho especificado.")

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_comeco(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.id, oc.data, oc.numero, cli.razao, oc.cliente "
                           f"FROM ordemcompra as oc "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"where oc.entradasaida = 'S' "
                           f"AND oc.STATUS = 'A';")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for dados in dados_oc:
                    lista_email = []
                    dados_cliente = []

                    id_oc, data, numero, cliente, id_cliente = dados

                    data_final = data.strftime('%d/%m/%Y')

                    cursor = conecta_robo.cursor()
                    cursor.execute(f"SELECT * FROM SEPARAR_OVS where NUM_OV = {numero} and CLIENTE_ID = {id_cliente};")
                    dados_lancados = cursor.fetchall()

                    if not dados_lancados:
                        coco = (data_final, numero, cliente, id_cliente)
                        dados_cliente.append(coco)

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT prodoc.codigo, prod.descricao, "
                                       f"COALESCE(prod.obs, ''), prod.unidade, prodoc.quantidade, "
                                       f"prodoc.produzido, prod.localizacao, prod.quantidade "
                                       f"FROM produtoordemcompra as prodoc "
                                       f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                                       f"where prodoc.mestre = {id_oc} "
                                       f"AND prodoc.produzido < prodoc.quantidade "
                                       f"ORDER BY prodoc.dataentrega;")
                        dados_prod_oc = cursor.fetchall()

                        for dados1 in dados_prod_oc:
                            codigo, descricao, ref, um, qtde, produzido, local, saldo = dados1
                            falta = "%.3f" % (float(qtde) - float(produzido))

                            dadus = (codigo, descricao, ref, um, falta, local, saldo)
                            lista_email.append(dadus)

                        if lista_email:
                            num_ov = dados_cliente[0][1]

                            caminho = fr'C:\Users\Anderson\PycharmProjects\robo_boby\Listagem - OV {num_ov}.pdf'
                            arquivo = f'Listagem - OV {num_ov}.pdf'

                            self.gerar_pdf_listagem_separar(arquivo, dados_cliente, lista_email)
                            self.inserir_no_banco(dados_cliente, num_ov)
                            self.envia_email(num_ov, caminho, arquivo)
                            self.excluir_arquivo(arquivo)

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = SepararOVS()