import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from dados_email import email_user, password
from comandos.conversores import valores_para_float
import os
import traceback
import inspect
import xml.etree.ElementTree as ET # noqa
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

        # noinspection HttpUrlsUsage
        self.nfe_ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

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
                self.processar_xmls(imap, lista_xmls)

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

            # Busca todos os emails
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
                                "num": num,
                                "xml": xml_encontrado,
                                "anexos": anexos_email
                            })

            return lista_emails_processar

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return []

    def obter_root_e_inf(self, xml_bytes):
        try:
            root = ET.fromstring(xml_bytes)
            infnfe = root.find('.//nfe:infNFe', self.nfe_ns)

            return root, infnfe

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def processar_xmls(self, imap, lista_xmls):
        try:
            for email_data in lista_xmls:
                ja_foi_lancado = False

                num = email_data["num"]
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

                cnpj_destinatario = dados_nf['destinatario']['cnpj']
                chave_existe = self.conferir_chave_nf(dados_nf)
                fornecedor_existe = self.conferir_fornecedor(dados_nf)
                transportador = dados_nf["transportadora"]

                if not cnpj_destinatario == self.cnpj_maquinas:
                    erros.append(f"CNPJ DO DESTINATÁRIO NÃO É DESTINADO A SUZUKI MÁQUINAS: {cnpj_destinatario}")
                elif chave_existe:
                    num_nf = dados_nf["numero"]
                    nome_fornecedor = dados_nf['emitente']['fantasia']
                    erros.append(f"NF JÁ FOI LANÇADA NO SISTEMA: Nº NF: {num_nf} - {nome_fornecedor}")

                    ja_foi_lancado = True

                elif not fornecedor_existe:
                    nome_fornecedor = dados_nf['emitente']['fantasia']
                    cnpj_fornecedor = dados_nf['emitente']['cnpj']
                    erros.append(f"O CNPJ DO FORNECEDOR NÃO ESTÁ CADASTRADO: {nome_fornecedor} - {cnpj_fornecedor}")
                elif transportador:
                    cnpj_transportador = dados_nf['transportadora']['cnpj']

                    if cnpj_transportador:
                        transportador_existe = self.conferir_transportador(dados_nf)

                        if not transportador_existe:
                            nome_transportador = dados_nf['transportadora']['nome']
                            erros.append(f"O CNPJ DO TRANSPORTADOR NÃO ESTÁ CADASTRADO EM "
                                         f"FORNECEDORES: {nome_transportador} - {cnpj_transportador}")

                if erros:
                    msg = "⚠ NF COM DIVERGÊNCIAS"
                    self.envia_email_erros_nf(erros, msg, anexos_email)

                    if ja_foi_lancado:
                        imap.store(num, '+FLAGS', '\\Deleted')
                else:
                    print("✅ NF VALIDADA COM SUCESSO!")

                    produtos = self.cadastrar_produtos(fornecedor_existe, dados_nf)

                    self.salvar_pre_nota(dados_nf, produtos, fornecedor_existe)

                    imap.store(num, '+FLAGS', '\\Deleted')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_se_e_nfe(self, xml_bytes):
        try:
            root, infnfe = self.obter_root_e_inf(xml_bytes)
            return infnfe is not None

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return False

    def ler_nfe_xml(self, xml_bytes: bytes) -> dict | None:
        try:
            root, infnfe = self.obter_root_e_inf(xml_bytes)

            if infnfe is None:
                return None

            ns = self.nfe_ns

            # ==========================
            # CHAVE DA NF
            # ==========================
            chave = infnfe.attrib.get("Id", "").replace("NFe", "")

            # ==========================
            # IDE
            # ==========================
            numero = infnfe.findtext('nfe:ide/nfe:nNF', None, ns)
            serie = infnfe.findtext('nfe:ide/nfe:serie', None, ns)
            data_emissao = infnfe.findtext('nfe:ide/nfe:dhEmi', None, ns)
            nat_op = infnfe.findtext('nfe:ide/nfe:natOp', None, ns)
            tp_nf = infnfe.findtext('nfe:ide/nfe:tpNF', None, ns)
            cuf = infnfe.findtext('nfe:ide/nfe:cUF', None, ns)

            # ==========================
            # EMITENTE
            # ==========================
            emit = infnfe.find('nfe:emit', ns)

            emitente = {
                "cnpj": emit.findtext('nfe:CNPJ', None, ns) if emit is not None else None,
                "nome": emit.findtext('nfe:xNome', None, ns) if emit is not None else None,
                "fantasia": emit.findtext('nfe:xFant', None, ns) if emit is not None else None,
                "uf": emit.findtext('nfe:enderEmit/nfe:UF', None, ns) if emit is not None else None,
                "pais": emit.findtext('nfe:enderEmit/nfe:xPais', None, ns) if emit is not None else None,
            }

            # ==========================
            # DESTINATÁRIO
            # ==========================
            dest = infnfe.find('nfe:dest', ns)

            destinatario = {
                "cnpj": dest.findtext('nfe:CNPJ', None, ns) if dest is not None else None,
                "cpf": dest.findtext('nfe:CPF', None, ns) if dest is not None else None,
                "nome": dest.findtext('nfe:xNome', None, ns) if dest is not None else None,
                "uf": dest.findtext('nfe:enderDest/nfe:UF', None, ns) if dest is not None else None,
                "pais": dest.findtext('nfe:enderDest/nfe:xPais', None, ns) if dest is not None else None,
            }

            # ==========================
            # TOTAIS (TRADUZIDO)
            # ==========================
            total = infnfe.find('nfe:total/nfe:ICMSTot', ns)

            totais = {
                "valor_produtos": None,
                "valor_nf": None,
                "valor_icms": None,
                "valor_icms_st": None,
                "valor_ipi": None,
                "valor_frete": None,
                "valor_seguro": None,
                "valor_desconto": None,
                "valor_outros": None
            }

            if total is not None:
                totais["valor_produtos"] = total.findtext('nfe:vProd', None, ns)
                totais["valor_nf"] = total.findtext('nfe:vNF', None, ns)
                totais["valor_icms"] = total.findtext('nfe:vICMS', None, ns)
                totais["valor_icms_st"] = total.findtext('nfe:vICMSST', None, ns)
                totais["valor_ipi"] = total.findtext('nfe:vIPI', None, ns)
                totais["valor_frete"] = total.findtext('nfe:vFrete', None, ns)
                totais["valor_seguro"] = total.findtext('nfe:vSeg', None, ns)
                totais["valor_desconto"] = total.findtext('nfe:vDesc', None, ns)
                totais["valor_outros"] = total.findtext('nfe:vOutro', None, ns)

            # ==========================
            # TRANSPORTE
            # ==========================
            transp = infnfe.find('nfe:transp', ns)

            transportadora = {
                "cnpj": None,
                "nome": None,
                "qVol": 0,
                "esp": [],
                "peso_bruto": 0.0,
                "peso_liquido": 0.0
            }

            if transp is not None:
                transporta = transp.find('nfe:transporta', ns)

                if transporta is not None:
                    transportadora["cnpj"] = transporta.findtext('nfe:CNPJ', None, ns)
                    transportadora["nome"] = transporta.findtext('nfe:xNome', None, ns)

                for vol in transp.findall('nfe:vol', ns):
                    qvol = vol.findtext('nfe:qVol', '0', ns)
                    esp = vol.findtext('nfe:esp', None, ns)
                    pl = vol.findtext('nfe:pesoL', '0', ns)
                    pb = vol.findtext('nfe:pesoB', '0', ns)

                    try:
                        transportadora["qVol"] += int(qvol)
                    except:
                        pass

                    if esp:
                        transportadora["esp"].append(esp)

                    try:
                        transportadora["peso_liquido"] += float(pl)
                    except:
                        pass

                    try:
                        transportadora["peso_bruto"] += float(pb)
                    except:
                        pass

            # ==========================
            # FATURAS
            # ==========================
            faturas = []
            for dup in infnfe.findall('nfe:cobr/nfe:dup', ns):
                faturas.append({
                    "numero": dup.findtext('nfe:nDup', None, ns),
                    "vencimento": dup.findtext('nfe:dVenc', None, ns),
                    "valor": dup.findtext('nfe:vDup', None, ns)
                })

            # ==========================
            # PRODUTOS
            # ==========================
            produtos = []

            for det in infnfe.findall('nfe:det', ns):
                prod = det.find('nfe:prod', ns)
                if prod is None:
                    continue

                inf_ad_prod = det.findtext('nfe:infAdProd', None, ns)

                ipi = det.find('nfe:imposto/nfe:IPI', ns)

                p_ipi = None

                if ipi is not None:
                    p_ipi = ipi.findtext('.//nfe:pIPI', None, ns)

                pedido = prod.findtext('nfe:VPed', None, ns)
                item_pedido = prod.findtext('nfe:vItemPed', None, ns)

                info_completa = " | ".join(
                    filtro for filtro in [
                        f"Pedido: {pedido}" if pedido else None,
                        f"Item: {item_pedido}" if item_pedido else None,
                        inf_ad_prod
                    ] if filtro
                )

                produtos.append({
                    "codigo": prod.findtext('nfe:cProd', None, ns),
                    "descricao": prod.findtext('nfe:xProd', None, ns),
                    "um": prod.findtext('nfe:uCom', None, ns),
                    "ncm": prod.findtext('nfe:NCM', None, ns),
                    "cfop": prod.findtext('nfe:CFOP', None, ns),
                    "quantidade": prod.findtext('nfe:qCom', None, ns),
                    "valor_unitario": prod.findtext('nfe:vUnCom', None, ns),
                    "valor_total": prod.findtext('nfe:vProd', None, ns),
                    "ipi_prod": p_ipi,
                    "informacoes_adicionais": info_completa
                })

            # ==========================
            # PROTOCOLO
            # ==========================
            protocolo = root.find('.//nfe:protNFe/nfe:infProt/nfe:nProt', ns)
            protocolo = protocolo.text if protocolo is not None else None

            return {
                "chave": chave,
                "numero": numero,
                "serie": serie,
                "data_emissao": data_emissao,
                "natOp": nat_op,
                "tpNF": tp_nf,
                "cUF": cuf,
                "emitente": emitente,
                "destinatario": destinatario,
                "totais": totais,
                "transportadora": transportadora,
                "faturas": faturas,
                "produtos": produtos,
                "protocolo": protocolo
            }

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

    def remove_espacos_e_especiais(self, string):
        try:
            if not string:
                return None
            return re.sub(r'\D', '', str(string)).strip()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def conferir_chave_nf(self, dados_nf):
        try:
            chave_nf = dados_nf["chave"]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT * "
                           f"FROM PRE_NF_COMPRA "
                           f"WHERE CHAVE_NFE = '{chave_nf}';")
            dados_chave = cursor.fetchall()

            return dados_chave

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def conferir_fornecedor(self, dados_nf):
        try:
            cnpj_fornecedor = dados_nf['emitente']['cnpj']

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, data_criacao, registro, razao, cnpj "
                           f"FROM fornecedores "
                           f"WHERE cnpj = '{cnpj_fornecedor}';")
            dados_fornecedor = cursor.fetchall()


            return dados_fornecedor

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def conferir_transportador(self, dados_nf):
        try:
            cnpj_transportador = dados_nf['transportadora']['cnpj']

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, data_criacao, registro, razao, cnpj "
                           f"FROM fornecedores "
                           f"WHERE cnpj = '{cnpj_transportador}';")
            dados_transportador = cursor.fetchall()

            return dados_transportador

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def cadastrar_produtos(self, dados_fornecedor, dados_nf):
        try:
            produtos = []

            id_fornecedor = dados_fornecedor[0][0]

            for prod_nf in dados_nf['produtos']:
                cod_prod = prod_nf['codigo']
                descricao = prod_nf['descricao']
                um = prod_nf['um']
                ncm = self.remove_espacos_e_especiais(prod_nf['ncm'])
                cfop = self.remove_espacos_e_especiais(prod_nf['cfop'])
                qtde = valores_para_float(prod_nf['quantidade'])
                unit = valores_para_float(prod_nf['valor_unitario'])
                ipi = valores_para_float(prod_nf['ipi_prod'])
                inf_produto = prod_nf['informacoes_adicionais']

                cursor = conecta.cursor()
                sql = """
                SELECT ID, CODIGO_FORNECEDOR
                FROM PRE_PRODUTO_FORNECEDOR
                WHERE CODIGO_FORNECEDOR = ?
                AND ID_FORNECEDOR = ?
                """

                cursor.execute(sql, (cod_prod, id_fornecedor))
                dados_produto = cursor.fetchall()

                if not dados_produto:
                    print("PRECISA CADASTRAR:", cod_prod, descricao, um, ncm)

                    inf_produto = (
                        cod_prod,
                        descricao,
                        um,
                        id_fornecedor,
                        ncm
                    )

                    sql_pre_prod = """
                    INSERT INTO PRE_PRODUTO_FORNECEDOR (
                    ID, CODIGO_FORNECEDOR, DESCRICAO, UM, ID_FORNECEDOR, NCM)
                    VALUES (GEN_ID(GEN_PRE_PRODUTO_FORNECEDOR_ID, 1), ?, ?, ?, ?, ?) 
                    RETURNING ID
                    """
                    cursor.execute(sql_pre_prod, inf_produto)
                    id_produto = cursor.fetchone()[0]

                    produtos.append({
                        "id": id_produto,
                        "cfop": cfop,
                        "qtde": qtde,
                        "unit": unit,
                        "ipi": ipi,
                        "obs": inf_produto
                    })
                else:
                    id_produto = dados_produto[0][0]

                    produtos.append({
                        "id": id_produto,
                        "cfop": cfop,
                        "qtde": qtde,
                        "unit": unit,
                        "ipi": ipi,
                        "obs": inf_produto
                    })

            conecta.commit()

            return produtos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_pre_nota(self, dados_nf, produtos, dados_fornecedor):
        try:
            id_fornecedor = dados_fornecedor[0][0]

            data_emissao_str = dados_nf["data_emissao"]
            data_emissao = datetime.fromisoformat(data_emissao_str).date()

            id_transportadora = None

            transportador = dados_nf["transportadora"]
            if transportador:
                cnpj_transportador = dados_nf['transportadora']['cnpj']

                if cnpj_transportador:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, data_criacao, registro, razao, cnpj "
                                   f"FROM fornecedores "
                                   f"WHERE cnpj = '{cnpj_transportador}';")
                    dados_transportador = cursor.fetchall()

                    id_transportadora = dados_transportador[0][0]

            volume = int(dados_nf['transportadora']['qVol'])

            esp_volume = dados_nf['transportadora']['esp']

            if isinstance(esp_volume, list):
                esp_volume = esp_volume[0]

            peso_bruto = valores_para_float(dados_nf['transportadora']['peso_bruto'])
            peso_liq = valores_para_float(dados_nf['transportadora']['peso_liquido'])

            valores_pre = (
                dados_nf["chave"],
                dados_nf["numero"],
                dados_nf["serie"],
                data_emissao,
                dados_nf["natOp"],
                dados_nf["tpNF"],
                id_fornecedor,
                valores_para_float(dados_nf['totais']['valor_produtos']),
                valores_para_float(dados_nf['totais']['valor_nf']),
                valores_para_float(dados_nf['totais']['valor_ipi']),
                valores_para_float(dados_nf['totais']['valor_frete']),
                valores_para_float(dados_nf['totais']['valor_desconto']),
                valores_para_float(dados_nf['totais']['valor_outros']),
                id_transportadora,
                volume,
                esp_volume,
                peso_bruto,
                peso_liq
            )

            sql_pre = """
                INSERT INTO PRE_NF_COMPRA (ID, 
                CHAVE_NFE, NUMERO_NF, SERIE, DATA_EMISSAO, NAT_OP, TP_NF, ID_FORNECEDOR, 
                VALOR_PRODUTOS, VALOR_TOTAL, VALOR_IPI, VALOR_FRETE, VALOR_DESCONTO, VALOR_OUTRO, 
                ID_TRANSPORTADOR, VOLUME, ESP_VOLUME, PESO_BRUTO, PESO_LIQUIDO)
                VALUES (GEN_ID(GEN_PRE_NF_COMPRA_ID, 1), 
                ?, ?, ?, ?, ?, ?, ?, 
                ?, ?, ?, ?, ?, ?, 
                ?, ?, ?, ?, ?)
                RETURNING ID
                """

            cursor = conecta.cursor()
            cursor.execute(sql_pre, valores_pre)
            id_pre_nota = cursor.fetchone()[0]

            if not produtos:
                raise Exception("Lista de produtos está vazia ou None")

            for indice, item in enumerate(produtos, start=1):
                valores_item = (
                    id_pre_nota,
                    indice,
                    item["id"],
                    item["cfop"],
                    item["qtde"],
                    item["unit"],
                    item["ipi"],
                    item["obs"],
                )

                sql_pre_prod = """
                    INSERT INTO PRE_NF_COMPRA_PRODUTOS (ID, 
                    ID_NF_PRE, ITEM, ID_PRODUTO_FORN, CFOP, QTDE, UNIT, IPI, OBS
                    )
                    VALUES (GEN_ID(GEN_PRE_NF_COMPRA_PRODUTOS_ID, 1), 
                    ?, ?, ?, ?, ?, ?, ?, ?)
                    """
                cursor.execute(sql_pre_prod, valores_item)

            faturas = dados_nf["faturas"]
            if faturas:
                for fatura in faturas:
                    data_venc_str = fatura["vencimento"]
                    data_venc = datetime.fromisoformat(data_venc_str).date()

                    valores_fatura = (
                        id_pre_nota,
                        data_venc,
                        fatura["valor"],
                    )

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
