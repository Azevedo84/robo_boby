import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from comandos.conversores import valores_para_float
import os
import traceback
import inspect
from dados_email import email_user, password


import imaplib
import email
from email.header import decode_header
from email.utils import parseaddr
from io import BytesIO
from PyPDF2 import PdfReader
import re

import subprocess
import tempfile
import time
from datetime import datetime

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


class ManipularEmailOC:
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

            msg_final = ""

            msg_final += f"Att,\n"
            msg_final += f"Suzuki Máquinas Ltda\n"
            msg_final += f"Fone (51) 3561.2583/(51) 3170.0965\n\n"
            msg_final += f"🟦 Mensagem gerada automaticamente pelo sistema de Planejamento e Controle da Produção (PCP) do ERP Suzuki.\n"
            msg_final += "🔸Por favor, não responda este e-mail diretamente."


            return saudacao, msg_final, to

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def envia_email_sem_requisicao(self, texto, cod_produto):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'SEM REQUISICAO DA OC - CÓDIGO DO PRODUTO NÃO TEM REQUISIÇÃO!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO cód do produto Nª {cod_produto}, não está no texto da OC!\n\n"
            body += f"{texto}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado sem REQUISIÇÃO")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_sem_codigo(self, texto, num_oc):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'SEM CÓDIGO DA OC - OC SEM IDENTIFICAR PRODUTOS!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nA OC Nª {num_oc}, não está no texto da OC!\n\n"
            body += f"{texto}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado sem REQUISIÇÃO")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_sem_fornecedor(self, cod_forn, num_oc):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'SEM FORNECEDOR OC - CÓDIGO DO FORNECEDOR NÃO EXISTE!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO cód do fornecedor Nª {cod_forn}, da OC {num_oc}não existe!\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado sem fornecedor")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_valor_total_diferente(self, texto, valor_pdf, valor_somado, num_oc):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'VALOR DIVERGENTE OC - O VALOR TOTAL DO PDF DIFERENTE!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO valro total do PDF da OC {num_oc} é de {valor_pdf} e a soma foi {valor_somado}!\n\n"
            body += f"{texto}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado total diferente")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def imprimir_pdf_background(self, pdf_bytes, nome_base="OC_SUZUKI"):
        with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".pdf",
                prefix=nome_base + "_"
        ) as temp_pdf:
            temp_pdf.write(pdf_bytes)
            caminho_pdf = temp_pdf.name

        print(f"🖨️ Imprimindo silenciosamente: {caminho_pdf}")

        try:
            caminho_sumatra = r"C:\Users\Anderson\AppData\Local\SumatraPDF\SumatraPDF.exe"

            subprocess.run(
                [
                    caminho_sumatra,
                    "-print-to-default",
                    "-silent",
                    caminho_pdf
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # tempo mínimo para o spooler
            time.sleep(2)

        except Exception as e:
            print("❌ Erro ao imprimir:", e)

            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

        finally:
            # apaga o PDF temporário
            try:
                os.remove(caminho_pdf)
            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
                return None

    def verificando_emails_caixa_entrada(self):
        try:
            # Conecta ao Gmail
            imap = imaplib.IMAP4_SSL("imap.gmail.com")
            imap.login(email_user, password)

            # Seleciona a caixa de entrada
            status, _ = imap.select("INBOX")
            print("SELECT STATUS:", status)

            # Busca todos os e-mails
            status, data = imap.search(None, "ALL")
            ids = data[0].split()
            print(f"Encontrados {len(ids)} emails na caixa de entrada\n")

            return ids, imap

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def percorer_email(self, ids, imap):
        try:
            lista_final_ocs = []
            pdf_bytes = None

            # Percorre todos os e-mails
            for num in ids:
                status, msg_data = imap.fetch(num, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])

                        # Remetente e assunto
                        from_ = msg.get("From")
                        nome, email_remetente = parseaddr(from_)
                        subject, encoding = decode_header(msg.get("Subject"))[0]
                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding if encoding else "utf-8")

                        print(f"\n")
                        print(f"De: {email_remetente}")
                        print(f"Assunto: {subject}")

                        # Processa e-mails multipart (com anexos)
                        if msg.is_multipart():
                            for part in msg.walk():
                                filename = part.get_filename()
                                if filename:
                                    # Decodifica corretamente o nome do arquivo
                                    decoded_filename, charset = decode_header(filename)[0]
                                    if isinstance(decoded_filename, bytes):
                                        decoded_filename = decoded_filename.decode(charset or "utf-8")

                                    texto = ""

                                    # Processa apenas PDFs
                                    if decoded_filename.lower().endswith(".pdf"):
                                        # Lê o PDF em memória
                                        pdf_bytes = part.get_payload(decode=True)
                                        pdf_file = BytesIO(pdf_bytes)
                                        reader = PdfReader(pdf_file)

                                        for page in reader.pages:
                                            page_content = page.extract_text()
                                            if page_content:
                                                texto += page_content

                                    # filtro OC SUZUKI
                                    if ("O.C.:" in texto
                                        and "Emissão:" in texto
                                        and "SUZUKI RECICLAD IND.MAQ" in texto
                                        and "93.183.853/0001-97" in texto
                                        and pdf_bytes):
                                        dados = self.filtrar_email_com_ordem_pdf(num, pdf_bytes, texto)
                                        if dados:
                                            lista_final_ocs.append(dados)

                                            print(f"EMAIL DE ORDEM DE COMPRA")
                                            print(f"\n")
                                    else:
                                        print("NÃO É ORDEM DE COMPRA")

                        else:
                            print("EMAIL SEM ANEXO")
                            # Caso não seja multipart (sem anexos)
                            body = msg.get_payload(decode=True).decode()
                            print("Corpo:", body[:200], "...")

            # ✅ Só retorna se os dois existirem
            if lista_final_ocs:
                return lista_final_ocs

            return None

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_num_ordem(self, texto):
        try:
            oc_match = re.search(r'O\.C\.: \s*(\d+)', texto)
            num_oc = oc_match.group(1) if oc_match else None

            return num_oc

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_emissao(self, texto):
        try:
            emissao = re.search(r'Emissão:\s*(\d{2}/\d{2}/\d{4})', texto)
            data_emissao = emissao.group(1) if emissao else None

            return data_emissao

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_cod_fornecedor(self, texto):
        try:
            codigo_fornecedor = re.search(r'C[oó]digo:\s*(\d+)', texto)
            codigo_forn = codigo_fornecedor.group(1) if codigo_fornecedor else None

            return codigo_forn

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_total_geral(self, texto):
        try:
            total_geral_match = re.search(r'TOTAL GERAL:\s*([\d.,]+)', texto)
            total_geral = total_geral_match.group(1) if total_geral_match else None

            return total_geral

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_frete(self, texto):
        try:
            frete_match = re.search(r'Valor\s+frete:\s*([\d.,]+)', texto)
            frete = frete_match.group(1) if frete_match else None

            return frete

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_outras_despesas(self, texto):
        try:
            outras_despesas_match = re.search(r'Outras despesas:\s*([\d.,]+)', texto)
            outras_despesas = outras_despesas_match.group(1) if outras_despesas_match else None

            return outras_despesas

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_desconto(self, texto):
        try:
            descontos_match = re.search(r'Descontos:\s*([\d.,]+)', texto)
            descontos = descontos_match.group(1) if descontos_match else None

            return descontos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def extrair_codigo_produto(self, texto_item, debug=False):
        try:
            cursor = conecta.cursor()

            # -------------------------------
            # LIMPEZA FORTE DE NCM
            # -------------------------------
            texto_limpo = re.sub(
                r'NCM\s*:\s*\d{6,8}',
                '',
                texto_item,
                flags=re.IGNORECASE
            )

            # -------------------------------
            # ESTRATÉGIAS (nome, peso, regex)
            # -------------------------------
            estrategias = [

                # 1️⃣ Código após referência técnica (ex: 82.01.123.45 550312)
                (
                    "ref_tecnica",
                    100,
                    r'\b\d{2}\.\d{2}\.\d{3}\.\d{2}\s+(\d{4,6})\b'
                ),

                # 2️⃣ Linha do ITEM (layout padrão SIGER)
                (
                    "linha_item",
                    95,
                    r'^\s*\d+\s+\S+\s+\S+\s+(\d{4,6})\b',
                ),

                # 3️⃣ Palavra CÓDIGO explícita
                (
                    "palavra_codigo",
                    90,
                    r'C[ÓO]DIGO\s+(\d{4,6})'
                ),

                # 4️⃣ Fallback controlado (ainda necessário)
                (
                    "fallback_controlado",
                    40,
                    r'^\s*\d+\s+.*?\b(\d{4,6})\b'
                ),
            ]

            candidatos = []

            # -------------------------------
            # EXECUTA TODAS AS ESTRATÉGIAS
            # -------------------------------
            for nome, peso, padrao in estrategias:
                encontrados = re.findall(
                    padrao,
                    texto_limpo,
                    re.IGNORECASE | re.MULTILINE
                )

                for codigo in encontrados:
                    candidatos.append({
                        "codigo": codigo,
                        "peso": peso,
                        "estrategia": nome
                    })

            if debug:
                print("\n--- CANDIDATOS ENCONTRADOS ---")
                for c in candidatos:
                    print(c)

            # -------------------------------
            # ESCOLHA DO MELHOR CÓDIGO
            # -------------------------------
            for c in sorted(candidatos, key=lambda x: x["peso"], reverse=True):
                cursor.execute(
                    "SELECT 1 FROM produto WHERE codigo = ?",
                    (c["codigo"],)
                )

                if cursor.fetchone():
                    if debug:
                        print(f"✔ Código escolhido: {c['codigo']} ({c['estrategia']})")
                    return c["codigo"]

            if debug:
                print("✖ Nenhum código válido encontrado")

            return None

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(
                nome_funcao,
                str(e),
                self.nome_arquivo,
                exc_traceback
            )
            return None

    def manipular_produtos_oc(self, texto):
        try:
            # --- BLOCO DE ITENS ---
            bloco_itens = re.search(
                r'ITEM REFERÊNCIA.*?DT\. ENTREGA(.*?)OBSERVAÇÕES:',
                texto,
                re.DOTALL
            )

            itens = []

            if bloco_itens:
                bloco = bloco_itens.group(1)

                itens_brutos = re.findall(
                    r'''
                                            \n\s*\d+\s+        # número do item
                                            .*?                # tudo do item
                                            (?=\n\s*\d+\s+|\Z) # até o próximo item ou fim
                                            ''',
                    bloco,
                    re.DOTALL | re.VERBOSE
                )

                codigo_produto = ""
                qtde = ""
                unit = ""

                for item_txt in itens_brutos:
                    codigo_produto = self.extrair_codigo_produto(item_txt, debug=True)
                    print("primeiro codigo extraido", codigo_produto)

                    if codigo_produto:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT id, codigo, descricao FROM produto where codigo = {codigo_produto};")
                        dados_produto = cursor.fetchall()

                        if dados_produto:
                            codigo_produto = dados_produto[0][1]

                            dados_match = re.search(
                                r'''
                                                        \n\s*([\d.,]+)\s+       # quantidade
                                                        ([A-Z]{1,3})\s+         # unidade
                                                        ([\d.,]+)\s+            # valor unitário
                                                        ([\d.,]+)\s+            # valor total
                                                        ([\d.,]+)\s+            # IPI
                                                        (\d{2}/\d{2}/\d{4})     # data entrega
                                                        ''',
                                item_txt,
                                re.VERBOSE
                            )

                            if dados_match:
                                qtde = dados_match.group(1)
                                unit = dados_match.group(3)


                                itens.append({
                                    "codigo_produto": codigo_produto,
                                    "quantidade": dados_match.group(1),
                                    "unidade": dados_match.group(2),
                                    "vl_unitario": dados_match.group(3),
                                    "vl_total": dados_match.group(4),
                                    "ipi": dados_match.group(5),
                                    "data_entrega": dados_match.group(6)
                                })

                if not itens or not codigo_produto or not qtde or not unit:
                    itens = []

                    for item_txt in itens_brutos:
                        texto_item = item_txt.replace('\n', ' ')

                        dados_match = re.search(
                            r'''
                                                    \s([\d.,]+)\s+       # quantidade
                                                    ([A-Z]{1,3})\s+      # unidade
                                                    ([\d.,]+)\s+         # valor unitário
                                                    ([\d.,]+)\s+         # valor total
                                                    ([\d.,]+)\s+         # IPI
                                                    (\d{2}/\d{2}/\d{4})  # data entrega
                                                    ''',
                            item_txt,
                            re.VERBOSE
                        )

                        cod_match = re.search(r'(\d{5,6})', texto_item)
                        codigo_produto = cod_match.group(1) if cod_match else None
                        print("segundo codigo extraido", codigo_produto)

                        if codigo_produto:
                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT id, codigo, descricao FROM produto where codigo = {codigo_produto};")
                            dados_produto = cursor.fetchall()

                            if dados_produto:
                                codigo_produto = dados_produto[0][1]

                                itens.append({
                                    "codigo_produto": codigo_produto,
                                    "quantidade": dados_match.group(1),
                                    "unidade": dados_match.group(2),
                                    "vl_unitario": dados_match.group(3),
                                    "vl_total": dados_match.group(4),
                                    "ipi": dados_match.group(5),
                                    "data_entrega": dados_match.group(6)
                                })

            return itens

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def filtrar_email_com_ordem_pdf(self, num, pdf_bytes, texto):
        try:
            dados = ()

            num_oc = self.mascara_num_ordem(texto)
            data_emissao = self.mascara_emissao(texto)
            codigo_forn = self.mascara_cod_fornecedor(texto)
            total_geral = self.mascara_total_geral(texto)
            frete = self.mascara_frete(texto)
            outras_despesas = self.mascara_outras_despesas(texto)
            descontos = self.mascara_desconto(texto)

            print("Num ordem:", num_oc)

            itens = self.manipular_produtos_oc(texto)

            if num_oc and itens and codigo_forn:
                dados = (texto, num, pdf_bytes, num_oc, data_emissao, codigo_forn, total_geral, frete,
                         outras_despesas, descontos, itens)
            else:
                if num_oc:
                    self.envia_email_sem_codigo(texto, num_oc)

            return dados

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_comeco(self):
        try:
            ids, imap = self.verificando_emails_caixa_entrada()

            lista_final_ocs = self.percorer_email(ids, imap)

            if lista_final_ocs:
                self.gravar_dados_ordens(imap, lista_final_ocs)

            imap.expunge()
            imap.logout()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def gravar_dados_ordens(self, imap, lista_final_ocs):
        try:
            obs_m = "OC LANÇADA PELO BOBY DE AZEVEDO"
            for i in lista_final_ocs:
                texto, num, pdf_bytes, num_oc, emissao, cod_forn, total_geral, frete, outras_despesas, descontos, itens = i

                valor_total_oc = 0

                frete_float = valores_para_float(frete)
                descontos_float = valores_para_float(descontos)

                valor_total_oc += frete_float
                valor_total_oc -= descontos_float

                total_geral_float = valores_para_float(total_geral)

                emissao_fire = datetime.strptime(emissao, '%d/%m/%Y').date()

                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT oc.numero, oc.data, oc.status FROM ordemcompra as oc "
                    f"where oc.entradasaida = 'E' and oc.numero = {num_oc};")
                dados_oc = cursor.fetchall()
                if not dados_oc:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_forn};")
                    dados_fornecedor = cursor.fetchall()

                    if not dados_fornecedor:
                        print(f'O Fornecedor {cod_forn} não está cadastrado!')
                        self.envia_email_sem_fornecedor(cod_forn, num_oc)
                    else:
                        testar_erros = 0

                        for ii in itens:
                            cod_produto = ii.get("codigo_produto")
                            qtde = ii.get("quantidade")
                            unit = ii.get("vl_unitario")

                            qtde_float = valores_para_float(qtde)
                            unit_float = valores_para_float(unit)

                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT id, descricao FROM produto where codigo = {cod_produto};")
                            dados_produto = cursor.fetchall()

                            dados_req = self.manipula_dados_req(cod_produto)

                            if not dados_produto:
                                print(f'O produto {cod_produto} não está cadastrado')
                                testar_erros += 1
                            elif not qtde_float:
                                print(f'A quantidade divergente do produto {cod_produto}')
                                testar_erros += 1
                            elif not unit_float:
                                print(f'O valor unitário divergente do produto {cod_produto}')
                                testar_erros += 1
                            elif not dados_req:
                                print(f'O produto {cod_produto} não tem requisição')
                                self.envia_email_sem_requisicao(texto, cod_produto)
                                testar_erros += 1
                            elif len(dados_req) > 1:
                                print(f'O produto {cod_produto} tem multiplas requisições')
                                testar_erros += 1

                        if not testar_erros:
                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_forn};")
                            dados_fornecedor = cursor.fetchall()
                            id_fornecedor, razao = dados_fornecedor[0]

                            cursor = conecta.cursor()
                            cursor.execute("select GEN_ID(GEN_ORDEMCOMPRA_ID,0) from rdb$database;")
                            ultimo_oc0 = cursor.fetchall()
                            ultimo_oc1 = ultimo_oc0[0]
                            ultimo_oc = int(ultimo_oc1[0]) + 1

                            cursor = conecta.cursor()
                            cursor.execute(f"Insert into ordemcompra "
                                           f"(ID, ENTRADASAIDA, NUMERO, DATA, STATUS, FORNECEDOR, "
                                           f"LOCALESTOQUE, FRETE, DESCONTOS, OBS) "
                                           f"values (GEN_ID(GEN_ORDEMCOMPRA_ID,1), "
                                           f"'E', {int(num_oc)}, '{emissao_fire}', 'A', {id_fornecedor}, '1', "
                                           f"{frete_float}, "
                                           f"{descontos_float}, '{obs_m}');")

                            for indice, ii in enumerate(itens, start=1):
                                cod_produto = ii.get("codigo_produto")
                                qtde = ii.get("quantidade")
                                unit = ii.get("vl_unitario")
                                ipi = ii.get("ipi")
                                entrega = ii.get("data_entrega")

                                total_prod = ii.get("vl_total")

                                dados_req = self.manipula_dados_req(cod_produto)

                                num_req, item_req = dados_req[0]

                                codigo_int = int(cod_produto)

                                entrega_fire = datetime.strptime(entrega, '%d/%m/%Y').date()
                                qtde_float = valores_para_float(qtde)
                                unit_float = valores_para_float(unit)
                                ipi_float = valores_para_float(ipi)

                                total_prod_float = valores_para_float(total_prod)

                                valor_total_oc += total_prod_float

                                if ipi_float:
                                    valor_total_oc += total_prod_float * (ipi_float / 100)

                                cursor = conecta.cursor()
                                cursor.execute(f"SELECT id, descricao FROM produto where codigo = {codigo_int};")
                                dados_produto = cursor.fetchall()

                                id_produto, descricao = dados_produto[0]

                                cursor = conecta.cursor()
                                cursor.execute(f"SELECT prodreq.id, prodreq.numero, prodreq.item,  "
                                               f"prod.codigo, prod.descricao as DESCRICAO, "
                                               f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                                               f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                                               f"prod.unidade, prodreq.quantidade "
                                               f"FROM produtoordemrequisicao as prodreq "
                                               f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                                               f"WHERE prodreq.numero = {num_req} "
                                               f"and prodreq.item = {item_req} "
                                               f"ORDER BY prodreq.numero;")
                                extrair_req = cursor.fetchall()

                                id_req = extrair_req[0][0]

                                cursor = conecta.cursor()
                                cursor.execute(
                                    f"Insert into produtoordemcompra (ID, MESTRE, ITEM, PRODUTO, QUANTIDADE, UNITARIO, "
                                    f"IPI, DATAENTREGA, NUMERO, CODIGO, PRODUZIDO, ID_PROD_REQ) "
                                    f"values (GEN_ID(GEN_PRODUTOORDEMCOMPRA_ID,1), {ultimo_oc}, "
                                    f"{indice}, {id_produto}, {qtde_float}, {unit_float}, {ipi_float}, "
                                    f"'{entrega_fire}', {int(num_oc)}, '{codigo_int}', 0.0, {id_req});")

                                cursor = conecta.cursor()
                                cursor.execute(f"UPDATE produtoordemrequisicao SET STATUS = 'B', "
                                               f"PRODUZIDO = {qtde_float} WHERE id = {id_req};")

                            tolerancia = 2.00  # diferença máxima aceitável em reais

                            if abs(valor_total_oc - total_geral_float) <= tolerancia:
                                conecta.commit()

                                print(f'Ordem de Compra Nº {num_oc} do {razao} foi lançada com sucesso!')

                                self.imprimir_pdf_background(pdf_bytes)
                                imap.store(num, '+FLAGS', '\\Deleted')
                            else:
                                self.envia_email_valor_total_diferente(texto, total_geral_float, valor_total_oc, num_oc)

                else:
                    print(f"Ordem de Compra Nº `{num_oc} já está lançada!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_req(self, cod_prod):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prodreq.numero, prodreq.item "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"WHERE prodreq.status = 'A' "
                           f"and prod.codigo = {cod_prod} "
                           f"ORDER BY prodreq.numero;")
            extrair_req = cursor.fetchall()

            return extrair_req

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

chama_classe = ManipularEmailOC()