import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from dados_email import email_user, password
from comandos.conversores import valores_para_float
import os
import traceback
import inspect
import xml.etree.ElementTree as ET # noqa
from typing import Any
import re
from datetime import datetime

import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from email.header import decode_header
from email.utils import parseaddr
import smtplib
import imaplib


class ConferenciaXmlNf:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.caminho_poppler = r'C:\Program Files\poppler-24.08.0\Library\bin'

        self.pasta_xml = r"C:/pasta_nf"

        self.cnpj_maquinas = "93183853000197"

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

    def manipula_comeco(self):
        try:
            ids, imap = self.verificando_emails_caixa_entrada()

            lista_xmls = self.percorer_email(ids, imap)

            if not lista_xmls:
                print("Nenhum XML encontrado nos e-mails.")
            else:
                self.processar_xmls(lista_xmls)

            imap.expunge()
            imap.logout()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
            lista_emails_processar = []

            for num in ids:
                status, msg_data = imap.fetch(num, "(RFC822)")

                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])

                        from_ = msg.get("From")
                        nome, email_remetente = parseaddr(from_)

                        subject, encoding = decode_header(msg.get("Subject"))[0]
                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding or "utf-8")

                        print(f"\nDe: {email_remetente}")
                        print(f"Assunto: {subject}")

                        anexos_email = []
                        xml_encontrado = None

                        if msg.is_multipart():
                            for part in msg.walk():

                                content_disposition = part.get("Content-Disposition")

                                if content_disposition and "attachment" in content_disposition:

                                    filename = part.get_filename()

                                    if filename:
                                        decoded_filename, charset = decode_header(filename)[0]

                                        if isinstance(decoded_filename, bytes):
                                            decoded_filename = decoded_filename.decode(charset or "utf-8")

                                        payload = part.get_payload(decode=True)

                                        anexo_dict = {
                                            "nome": decoded_filename,
                                            "conteudo": payload
                                        }

                                        anexos_email.append(anexo_dict)

                                        # 🔥 verifica se é XML
                                        if decoded_filename.lower().endswith(".xml"):
                                            xml_encontrado = anexo_dict

                        # Se encontrou XML nesse email
                        if xml_encontrado:
                            lista_emails_processar.append({
                                "xml": xml_encontrado,
                                "anexos": anexos_email
                            })

            return lista_emails_processar

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return []

    def processar_xmls(self, lista_xmls):
        try:
            for email_data in lista_xmls:

                xml_nome = email_data["xml"]["nome"]
                xml_bytes = email_data["xml"]["conteudo"]
                anexos_email = email_data["anexos"]

                print(f"\n🔎 Processando: {xml_nome}")

                if not self.verifica_se_e_nfe(xml_bytes):
                    print("Arquivo não é NF-e válida.")
                    continue

                erros = []

                dados_nf = self.ler_nfe_xml(xml_bytes)

                if not dados_nf:
                    print("Erro ao ler NF.")
                    continue

                dados_fornecedor, erros = self.conferir_dados_gerais(dados_nf, erros)

                if not dados_fornecedor:
                    msg = "Fornecedor inválido."
                    print(msg)
                    self.envia_email_erros_nf(erros, msg, anexos_email)
                    continue

                dados_pre_nota, erros = self.conferir_dados_produtos(
                    dados_fornecedor, dados_nf, erros
                )

                if erros:
                    msg = "⚠ NF COM DIVERGÊNCIAS"
                    print(msg)
                    self.envia_email_erros_nf(erros, msg, anexos_email)
                else:
                    print("✅ NF VALIDADA COM SUCESSO!")

                    dados_nf_bc = self.verifica_pre_ja_lancado(dados_pre_nota)

                    if not dados_nf_bc:
                        self.salvar_pre_nota(dados_pre_nota)
                    else:
                        erro = []
                        msg = "⚠ NF JÁ FOI SALVA NA TABELA PRÉ LANÇAMENTO"
                        erro.append(msg)
                        print(msg)
                        self.envia_email_erros_nf(erro, msg, anexos_email)

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

    def envia_email_erros_nf(self, erros, msg_status, anexos_email):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f"PRÉ LANÇAMENTO NF – {msg_status}"

            msg_email = MIMEMultipart()
            msg_email['From'] = email_user
            msg_email['To'] = ", ".join(to) if isinstance(to, list) else to
            msg_email['Subject'] = subject

            body = f"{saudacao}\n\n"

            if erros:
                for erro in erros:
                    body += f"- {erro}\n"

            body += "\n" + msg_final

            msg_email.attach(MIMEText(body, 'plain'))

            # 🔥 Anexando TODOS os anexos originais
            for anexo in anexos_email:
                part = MIMEBase('application', "octet-stream")
                part.set_payload(anexo["conteudo"])
                encoders.encode_base64(part)

                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=Header(anexo["nome"], 'utf-8').encode()
                )

                msg_email.attach(part)

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, msg_email.as_string())
            server.quit()

            print('EMAIL COM PROBLEMAS ENVIADO COM SUCESSO!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def listar_xmls_pasta(self):
        try:
            arquivos_xml = []

            for nome_arquivo in os.listdir(self.pasta_xml):

                if nome_arquivo.lower().endswith(".xml"):

                    caminho_completo = os.path.join(self.pasta_xml, nome_arquivo)

                    if os.path.isfile(caminho_completo):
                        arquivos_xml.append(caminho_completo)

            return arquivos_xml

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_se_e_nfe(self, xml_bytes):
        try:
            root = ET.fromstring(xml_bytes)

            if root.find('.//{http://www.portalfiscal.inf.br/nfe}infNFe') is not None:
                return True

            return False

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return False

    def remove_espacos_e_especiais(self, string):
        try:
            if not string:
                return None
            return re.sub(r'\D', '', str(string)).strip()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def conferir_dados_gerais(self, dados_nf, erros):
        try:
            dados_fornecedor = []

            cnpj = dados_nf['emitente']['cnpj']

            cnpj_destinatario = dados_nf['destinatario']['cnpj']

            if cnpj_destinatario == self.cnpj_maquinas:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, data_criacao, registro, razao, cnpj "
                               f"FROM fornecedores "
                               f"WHERE cnpj = '{cnpj}';")
                dados_fornecedor = cursor.fetchall()
            else:
                erros.append(f"CNPJ DO DESTINATÁRIO NÃO É DESTINADO A SUZUKI MÁQUINAS: {cnpj_destinatario}")

            return dados_fornecedor, erros

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def conferir_dados_produtos(self, dados_fornecedor, dados_nf, erros):
        try:
            id_fornecedor = dados_fornecedor[0][0]
            cnpj = dados_fornecedor[0][4]

            dados_pre_nota = {
                "fornecedor_id": id_fornecedor,
                "numero_nf": dados_nf["numero"],
                "serie": dados_nf["serie"],
                "data_emissao": dados_nf["data_emissao"],
                "valor_produtos": valores_para_float(dados_nf["totais"]["valor_produtos"]),
                "valor_nf": valores_para_float(dados_nf["totais"]["valor_nf"]),
                "valor_icms": valores_para_float(dados_nf["totais"]["valor_icms"]),
                "valor_frete": valores_para_float(dados_nf["totais"]["frete_total"]),
                "valor_desconto": valores_para_float(dados_nf["totais"]["desconto_total"]),
                "peso_bruto": valores_para_float(dados_nf["peso_bruto"]),
                "peso_liquido": valores_para_float(dados_nf["peso_liquido"]),
                "faturas": dados_nf["faturas"],
                "itens": []
            }

            ocs_encontradas = {}

            for prod_nf in dados_nf['produtos']:

                cod_prod_f = prod_nf['codigo']

                vinculo = self.conferir_vinculo_produtos(id_fornecedor, cod_prod_f)

                if not vinculo:
                    erros.append(f"Produto não vinculado: {cod_prod_f}")
                    continue

                id_produto = vinculo[0][0]
                cod_siger = vinculo[0][1]

                dados_oc = self.consulta_oc_pendente(cnpj, cod_siger)

                if not dados_oc:
                    erros.append(f"O produto {cod_siger} não foi encontrado nas OCs pendentes!")
                    continue

                (id_oc, emissao_oc, num_oc, nome_forn, frete, descont, obs_oc,
                 cod, descr, ref, um, ncm, qtde, unit, ipi, dt_entrega) = dados_oc[0]

                # =============================
                # GUARDA OC AGRUPADA
                # =============================

                if id_oc not in ocs_encontradas:
                    ocs_encontradas[id_oc] = {
                        "emissao": emissao_oc,
                        "fornecedor": nome_forn,
                        "frete": valores_para_float(frete),
                        "desconto": valores_para_float(descont),
                        "obs": obs_oc,
                    }

                # =============================
                # VALIDA ITEM
                # =============================

                ncm_nf = self.remove_espacos_e_especiais(prod_nf['ncm'])
                qtde_nf = valores_para_float(prod_nf['quantidade'])
                unit_nf = valores_para_float(prod_nf['valor_unitario'])
                ipi_nf = valores_para_float(prod_nf['ipi'])

                cfop_nf = self.remove_espacos_e_especiais(prod_nf['cfop'])

                ncm_oc = self.remove_espacos_e_especiais(ncm)
                qtde_oc = valores_para_float(qtde)
                unit_oc = valores_para_float(unit)
                ipi_oc = valores_para_float(ipi)

                if ncm_nf != ncm_oc:
                    erros.append(f"NCM diferente no produto {cod_siger}")

                if round(qtde_nf, 4) != round(qtde_oc, 4):
                    erros.append(f"Quantidade diferente no produto {cod_siger}")

                if round(unit_nf, 4) != round(unit_oc, 4):
                    erros.append(f"Valor unitário diferente no produto {cod_siger}")

                if round(ipi_nf, 4) != round(ipi_oc, 4):
                    erros.append(f"IPI diferente no produto {cod_siger}")

                # =============================
                # MONTA ITEM INTERNO
                # =============================

                item_interno = {
                    "id_oc": id_oc,
                    "ncm_nf": ncm_nf,
                    "cfop_nf": cfop_nf,
                    "id_produto": id_produto,
                    "descricao": descr,
                    "quantidade_nf": qtde_nf,
                    "quantidade_oc": qtde_oc,
                    "valor_unitario_nf": unit_nf,
                    "valor_unitario_oc": unit_oc,
                    "valor_total_nf": qtde_nf * unit_nf,
                    "ipi_nf": ipi_nf
                }

                dados_pre_nota["itens"].append(item_interno)

            # =============================
            # VALIDA TOTAIS (FRETE / DESCONTO)
            # =============================

            frete_total_oc = sum(oc["frete"] for oc in ocs_encontradas.values())
            desconto_total_oc = sum(oc["desconto"] for oc in ocs_encontradas.values())

            frete_nf = dados_pre_nota["valor_frete"]
            desconto_nf = dados_pre_nota["valor_desconto"]

            if round(frete_nf, 2) != round(frete_total_oc, 2):
                erros.append("Frete da NF diferente do total das OCs")

            if round(desconto_nf, 2) != round(desconto_total_oc, 2):
                erros.append("Desconto da NF diferente do total das OCs")

            return dados_pre_nota, erros

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def conferir_vinculo_produtos(self, id_fornecedor, cod_prod_f):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod_f.ID_PRODUTO, prod.codigo "
                           f"FROM PRODUTO_FORNECEDOR as prod_f "
                           f"INNER JOIN produto prod ON prod_f.ID_PRODUTO = prod.id "
                           f"WHERE prod_f.ID_FORNECEDOR = {id_fornecedor} "
                           f"AND prod_f.COD_PRODUTO_F = {cod_prod_f};")

            vinculo = cursor.fetchall()

            return vinculo

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_oc_pendente(self, cnpj, cod_prod):
        try:
            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.id, oc.data, oc.numero, forn.razao, oc.frete, oc.descontos, oc.obs, "
                f"prodoc.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, prod.ncm, prodoc.quantidade, prodoc.unitario, prodoc.ipi, "
                f"prodoc.dataentrega "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where prodoc.codigo = '{cod_prod}' "
                f"and forn.cnpj = '{cnpj}' "
                f"and oc.entradasaida = 'E' "
                f"AND oc.STATUS = 'A' "
                f"AND prodoc.produzido < prodoc.quantidade;")
            dados_oc = cursor.fetchall()

            return dados_oc

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def ler_nfe_xml(self, xml_bytes: bytes) -> dict[str, Any] | None:
        try:
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

            root = ET.fromstring(xml_bytes)

            infnfe = root.find('.//nfe:infNFe', ns)
            if infnfe is None:
                return None

            # ==========================
            # IDE
            # ==========================
            numero = infnfe.findtext('nfe:ide/nfe:nNF', default=None, namespaces=ns)
            serie = infnfe.findtext('nfe:ide/nfe:serie', default=None, namespaces=ns)
            data_emissao = infnfe.findtext('nfe:ide/nfe:dhEmi', default=None, namespaces=ns)

            # ==========================
            # EMITENTE
            # ==========================
            emit = infnfe.find('nfe:emit', ns)

            emitente = {
                "cnpj": emit.findtext('nfe:CNPJ', default=None, namespaces=ns) if emit is not None else None,
                "nome": emit.findtext('nfe:xNome', default=None, namespaces=ns) if emit is not None else None
            }

            # ==========================
            # DESTINATÁRIO
            # ==========================
            dest = infnfe.find('nfe:dest', ns)

            destinatario = {
                "cnpj": dest.findtext('nfe:CNPJ', default=None, namespaces=ns) if dest is not None else None,
                "nome": dest.findtext('nfe:xNome', default=None, namespaces=ns) if dest is not None else None
            }

            # ==========================
            # TOTAIS
            # ==========================
            total = infnfe.find('nfe:total/nfe:ICMSTot', ns)

            valor_produtos = None
            valor_nf = None
            valor_icms = None
            frete_total = None
            desconto_total = None

            if total is not None:
                valor_produtos = total.findtext('nfe:vProd', default=None, namespaces=ns)
                valor_nf = total.findtext('nfe:vNF', default=None, namespaces=ns)
                valor_icms = total.findtext('nfe:vICMS', default=None, namespaces=ns)
                frete_total = total.findtext('nfe:vFrete', default=None, namespaces=ns)
                desconto_total = total.findtext('nfe:vDesc', default=None, namespaces=ns)

            totais = {
                "valor_produtos": valor_produtos,
                "valor_nf": valor_nf,
                "valor_icms": valor_icms,
                "frete_total": frete_total,
                "desconto_total": desconto_total
            }

            # ==========================
            # PESOS
            # ==========================
            peso_bruto = 0.0
            peso_liquido = 0.0

            for vol in infnfe.findall('nfe:transp/nfe:vol', ns):
                pL = vol.findtext('nfe:pesoL', default='0', namespaces=ns)
                pB = vol.findtext('nfe:pesoB', default='0', namespaces=ns)

                try:
                    peso_liquido += float(pL)
                except:
                    pass

                try:
                    peso_bruto += float(pB)
                except:
                    pass

            # ==========================
            # FATURAS / DUPLICATAS
            # ==========================
            faturas = []

            for dup in infnfe.findall('nfe:cobr/nfe:dup', ns):
                numero_dup = dup.findtext('nfe:nDup', default=None, namespaces=ns)
                data_venc = dup.findtext('nfe:dVenc', default=None, namespaces=ns)
                valor_dup = dup.findtext('nfe:vDup', default=None, namespaces=ns)

                faturas.append({
                    "numero": numero_dup,
                    "vencimento": data_venc,
                    "valor": valor_dup
                })

            # ==========================
            # PRODUTOS
            # ==========================
            produtos = []

            for det in infnfe.findall('nfe:det', ns):

                prod = det.find('nfe:prod', ns)
                if prod is None:
                    continue

                imposto = det.find('nfe:imposto', ns)

                ipi_valor = None

                if imposto is not None:
                    ipi = imposto.find('nfe:IPI', ns)
                    if ipi is not None:
                        ipi_trib = ipi.find('nfe:IPITrib', ns)
                        if ipi_trib is not None:
                            ipi_valor = ipi_trib.findtext('nfe:vIPI', default=None, namespaces=ns)

                produtos.append({
                    "codigo": prod.findtext('nfe:cProd', default=None, namespaces=ns),
                    "descricao": prod.findtext('nfe:xProd', default=None, namespaces=ns),
                    "ncm": prod.findtext('nfe:NCM', default=None, namespaces=ns),
                    "cfop": prod.findtext('nfe:CFOP', default=None, namespaces=ns),
                    "quantidade": prod.findtext('nfe:qCom', default=None, namespaces=ns),
                    "valor_unitario": prod.findtext('nfe:vUnCom', default=None, namespaces=ns),
                    "valor_total": prod.findtext('nfe:vProd', default=None, namespaces=ns),
                    "ipi": ipi_valor
                })

            return {
                "numero": numero,
                "serie": serie,
                "data_emissao": data_emissao,
                "emitente": emitente,
                "destinatario": destinatario,
                "totais": totais,
                "peso_bruto": peso_bruto,
                "peso_liquido": peso_liquido,
                "faturas": faturas,
                "produtos": produtos
            }

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def verifica_pre_ja_lancado(self, dados_pre_nota):
        try:
            id_fornecedor = dados_pre_nota["fornecedor_id"]
            num_nf = dados_pre_nota["numero_nf"]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT * "
                           f"FROM PRE_NF_COMPRA "
                           f"WHERE ID_FORNECEDOR = {id_fornecedor} "
                           f"and NUMERO_NF = {num_nf};")
            dados_nf = cursor.fetchall()

            return dados_nf

        except Exception as e:
            conecta.rollback()
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_pre_nota(self, dados_pre_nota):
        try:
            # ==========================
            # INSERE CABEÇALHO
            # ==========================
            data_emissao_str = dados_pre_nota["data_emissao"]
            data_emissao = datetime.fromisoformat(data_emissao_str).date()

            valores_pre = (
                dados_pre_nota["fornecedor_id"],
                dados_pre_nota["numero_nf"],
                data_emissao,
                dados_pre_nota["valor_produtos"],
                dados_pre_nota["valor_nf"],
                dados_pre_nota["valor_frete"],
                dados_pre_nota["valor_desconto"],
                dados_pre_nota["peso_bruto"],
                dados_pre_nota["peso_liquido"],
            )
            print("valores_pre", valores_pre)

            sql_pre = """
                INSERT INTO PRE_NF_COMPRA (
                ID, ID_FORNECEDOR, NUMERO_NF, DATA_EMISSAO, VALOR_PRODUTOS, VALOR_TOTAL, 
                FRETE, DESCONTOS, PESO_BRUTO, PESO_LIQUIDO
                )
                VALUES (GEN_ID(GEN_PRE_NF_COMPRA_ID, 1), ?, ?, ?, ?, ?, ?, ?, ?, ?)
                RETURNING ID
                """

            cursor = conecta.cursor()
            cursor.execute(sql_pre, valores_pre)
            id_pre_nota = cursor.fetchone()[0]

            for indice, item in enumerate(dados_pre_nota["itens"], start=1):
                valores_item = (
                    id_pre_nota,
                    indice,
                    item["id_produto"],
                    item["id_oc"],
                    item["ncm_nf"],
                    item["cfop_nf"],
                    item["quantidade_nf"],
                    item["valor_unitario_nf"],
                    item["ipi_nf"],
                )
                print("valores_item", valores_item)

                sql_pre_prod = """
                    INSERT INTO PRE_NF_COMPRA_PRODUTOS (
                    ID, ID_NF_PRE, ITEM, ID_PRODUTO, ID_OC, NCM, CFOP, QTDE, UNIT, IPI
                    )
                    VALUES (GEN_ID(GEN_PRE_NF_COMPRA_PRODUTOS_ID, 1), ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """
                cursor.execute(sql_pre_prod, valores_item)

            faturas = dados_pre_nota["faturas"]
            if faturas:
                for fatura in faturas:
                    data_venc_str = fatura["vencimento"]
                    data_venc = datetime.fromisoformat(data_venc_str).date()

                    valores_fatura = (
                        id_pre_nota,
                        data_venc,
                        fatura["valor"],
                    )

                    print("valores_fatura", valores_fatura)

                    sql_pre_prod = """
                                        INSERT INTO PRE_NF_COMPRA_FATURAS (
                                        ID, ID_NF_PRE, VENCIMENTO, VALOR)
                                        VALUES (GEN_ID(GEN_PRE_NF_COMPRA_FATURAS_ID, 1), ?, ?, ?)
                                        """
                    cursor.execute(sql_pre_prod, valores_fatura)


            conecta.commit()
            print("NF LANÇADA!")

        except Exception as e:
            conecta.rollback()
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = ConferenciaXmlNf()
