import os
import win32com.client
from core.banco import conecta_engenharia
from core.erros import trata_excecao
from core.email_service import dados_email
from core.inventor import definir_classificacao
from core.inventor import padronizar_caminho, corrigir_caminho_inventor
from datetime import datetime
from multiprocessing import freeze_support
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

class WorkerFilaConferencia:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

        self.manipula_comeco()

    def envia_email_desenho_sem_vinculo(self, texto, desenho):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA FILA - DESENHO SEM VÍNCULO {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO desenho {desenho} tem problemas com vínculos dos arquivos!\n\n"

            for i in texto:
                body += f"{i}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado SEM VINCULO")

        except Exception as e:
            trata_excecao(e)
            raise

    def envia_email_muitos_caracteres(self, desenho, campos):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA FILA - CAMPOS COM MAIS DE 100 CARACTERES {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO desenho {desenho} tem campos com muitos caracteres!\n\n"

            for i in campos:
                body += f"{i}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado MUITOS CARACTERES")

        except Exception as e:
            trata_excecao(e)
            raise

    def enviar_email_erro_unicode(self, caminho, props, erro):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ERRO UTF-8 FILA - {os.path.basename(caminho)}'

            body = f"""{saudacao}

            Foi detectado um erro de encoding ao processar um arquivo do Inventor.

            Arquivo:
            {caminho}

            Erro:
            {erro}

            Propriedades capturadas:
            {props}

            {msg_final}
            """

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)
            server.sendmail(email_user, self.destinatario, msg.as_string())
            server.quit()

            print("📧 Email enviado (erro UTF-8)")

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_estrutura_inventor(self, doc, cursor, origem, id_arquivo):
        try:
            comp = doc.ComponentDefinition
            bom = comp.BOM

            try:
                bom.StructuredViewEnabled = True
                bom.StructuredViewFirstLevelOnly = False
            except Exception as e:
                msg1 = f"Sem BOM estruturada: {doc.FullFileName}"
                msg2 = f"{e}"

                cursor.execute("""
                                    SELECT ID, caminho
                                    FROM ARQUIVOS
                                    WHERE ID = ?
                                """, (id_arquivo,))

                dados_arq = cursor.fetchall()

                id_arq, caminho = dados_arq[0]

                if not "\\inventor\\biblioteca" in caminho:
                    self.enviar_email_erro_unicode(caminho, msg1, msg2)
                else:
                    print(f"❌ Sem BOM estruturada (biblioteca): {doc.FullFileName}")
                    print(f"Produto da Biblioteca e vai ser exclúido da fila!")

                    cursor.execute("""
                            DELETE FROM FILA_CONFERENCIA
                            WHERE ID_ARQUIVO=?
                        """, (id_arquivo,))

                    return {}

            view = None
            for v in bom.BOMViews:
                if v.ViewType == 62465:
                    view = v
                    break

            if not view:
                print(f"❌ BOM view não encontrada: {doc.FullFileName}")
                raise Exception(f"BOM view não encontrada: {doc.FullFileName}")

            estrutura = {}

            for row in view.BOMRows:

                # 🔴 VALIDA COMPONENT DEFINITIONS (erro COM comum)
                try:
                    comp_defs = row.ComponentDefinitions
                    if comp_defs.Count == 0:
                        continue
                except Exception as e:
                    print(f"❌ BOMRow inválida (erro COM): {e}")
                    nome = os.path.basename(doc.FullFileName)
                    raise Exception(f"BOM inválida (ComponentDefinitions): {nome}")

                # 🔴 PEGA COMPONENTE
                try:
                    comp_def = comp_defs.Item(1)
                except Exception as e:
                    print(f"❌ Erro ao acessar ComponentDefinition: {e}")
                    raise Exception(f"Erro na BOM (Item): {doc.FullFileName}")

                # 🔴 PEGA CAMINHO
                try:
                    caminho = comp_def.Document.FullFileName
                    caminho_novo = corrigir_caminho_inventor(caminho)
                except Exception as e:
                    print(f"❌ Erro ao acessar Document: {e}")
                    raise Exception(f"Erro na BOM (Document): {doc.FullFileName}")

                if not caminho_novo:
                    print("❌ Caminho vazio")
                    raise Exception(f"Caminho vazio na BOM: {doc.FullFileName}")

                caminho_strip = caminho_novo.strip()
                caminho_padrao = padronizar_caminho(caminho_strip)

                if not caminho_padrao:
                    print("❌ CAMINHO SUSPEITO após padronização", caminho_novo)
                    raise Exception(f"Caminho inválido na BOM: {doc.FullFileName}")

                # 🔴 BUSCA/CRIA ID
                id_filho = self.consulta_e_cria_id_arquivo(cursor, caminho_padrao)

                if not id_filho:
                    print(f"❌ Falha ao obter ID do arquivo: {caminho_padrao}")
                    raise Exception(f"Erro ao criar/buscar arquivo: {caminho_padrao}")

                # 🔴 QUANTIDADE
                try:
                    qtde = float(row.ItemQuantity)
                except Exception as e:
                    print(f"❌ Erro ao ler quantidade: {e}")
                    raise Exception(f"Erro na quantidade da BOM: {doc.FullFileName}")

                estrutura[id_filho] = qtde

                # 🔴 FILA (só depois que passou tudo acima)
                if origem != "ALTERADOS":
                    self.inserir_fila_conferencia(cursor, id_filho)

            return estrutura

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_e_cria_id_arquivo(self, cursor, caminho):
        try:
            if not caminho:
                return False

            caminho = caminho.strip()

            caminho = padronizar_caminho(caminho)

            if not caminho:
                print("🚨 CAMINHO SUSPEITO - consulta_e_cria_id_arquivo", caminho)
                return False

            cursor.execute("""
                SELECT ID FROM ARQUIVOS WHERE LOWER(CAMINHO)=?
            """, (caminho,))

            row = cursor.fetchone()

            if row:
                return row[0]

            nome: str = os.path.basename(caminho)
            nome_sem_ext = os.path.splitext(nome)[0]

            classificacao = definir_classificacao(caminho, nome_sem_ext)

            stat = os.stat(caminho)

            tamanho = stat.st_size
            data_mod = datetime.fromtimestamp(stat.st_mtime)

            extensaos = os.path.splitext(nome)[1].lower()
            tipo_arquivo = extensaos.replace(".", "").upper()

            cursor.execute("""
                INSERT INTO ARQUIVOS (
                    ARQUIVO, NOME_BASE, TIPO_ARQUIVO,
                    CLASSIFICACAO, CAMINHO, TAMANHO, DATA_MOD
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                RETURNING ID
            """, (
                nome,
                nome_sem_ext,
                tipo_arquivo,
                classificacao,
                caminho,
                tamanho,
                data_mod
            ))

            return cursor.fetchone()[0]

        except Exception as e:
            trata_excecao(e)
            raise

    def inserir_fila_conferencia(self, cursor, id_arquivo):
        try:
            cursor.execute("""
                SELECT 1 FROM FILA_CONFERENCIA WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            if cursor.fetchone():
                return

            cursor.execute("""
                INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO, ORIGEM)
                VALUES (?, ?)
            """, (id_arquivo, "ALTERADOS"))

        except Exception as e:
            trata_excecao(e)
            raise

    def sincronizar_estrutura(self, cursor, id_pai, estrutura_nova):
        try:
            estrutura_atual = self.consulta_estrutura_atual(cursor, id_pai)

            inseridos = 0
            atualizados = 0
            deletados = 0

            # 🔹 INSERT ou UPDATE
            for id_filho, qtde_nova in estrutura_nova.items():

                if id_filho not in estrutura_atual:
                    cursor.execute("""
                        INSERT INTO ESTRUTURA (ID_PAI, ID_FILHO, QTDE)
                        VALUES (?, ?, ?)
                    """, (id_pai, id_filho, qtde_nova))
                    inseridos += 1

                else:
                    qtde_atual = estrutura_atual[id_filho]

                    if round(qtde_atual, 4) != round(qtde_nova, 4):
                        cursor.execute("""
                            UPDATE ESTRUTURA
                            SET QTDE=?
                            WHERE ID_PAI=? AND ID_FILHO=?
                        """, (qtde_nova, id_pai, id_filho))
                        atualizados += 1

            # 🔻 DELETE só do que não existe mais
            for id_filho in estrutura_atual:

                if id_filho not in estrutura_nova:
                    cursor.execute("""
                        DELETE FROM ESTRUTURA
                        WHERE ID_PAI=? AND ID_FILHO=?
                    """, (id_pai, id_filho))
                    deletados += 1

            print(f"🟡 Estrutura → +{inseridos} | ~{atualizados} | -{deletados}")

            return inseridos, atualizados, deletados

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_estrutura_atual(self, cursor, id_pai):
        try:
            cursor.execute("""
                SELECT ID_FILHO, QTDE
                FROM ESTRUTURA
                WHERE ID_PAI=?
            """, (id_pai,))

            estrutura = {}

            for row in cursor.fetchall():
                id_filho, qtde = row
                estrutura[id_filho] = float(qtde)

            return estrutura

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_idw_colocar_na_fila(self, cursor, id_arquivo, caminho):
        try:
            if "\\inventor\\biblioteca" not in caminho:
                cursor.execute("""
                    SELECT ID_ARQUIVO, ID_ARQUIVO_REFERENCIA
                    FROM PROPRIEDADES_IDW
                    WHERE ID_ARQUIVO_REFERENCIA = ?
                """, (id_arquivo,))

                dados_idw = cursor.fetchall()

                if dados_idw:
                    for i in dados_idw:
                        id_arq, id_arq_ref = i

                        self.inserir_fila_conferencia(cursor, id_arq)

        except Exception as e:
            trata_excecao(e)
            raise

    def conectar_inventor(self):
        try:
            print("Conectando ao Inventor...")

            try:
                inventor = win32com.client.GetActiveObject("Inventor.Application")
                print("♻️ Reutilizando Inventor já aberto")
            except:
                inventor = win32com.client.Dispatch("Inventor.Application")
                print("🆕 Abrindo novo Inventor")

            inventor.Visible = False
            inventor.SilentOperation = True

            return inventor

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_comeco(self):
        try:
            freeze_support()
            print("🚀 Worker de fila iniciado")
            self.worker_fila()

        except Exception as e:
            trata_excecao(e)
            raise

    def worker_fila(self):
        inventor = None
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                    SELECT fila.ID_ARQUIVO, fila.ORIGEM, arq.caminho, arq.NOME_BASE
                    FROM FILA_CONFERENCIA as fila
                    INNER JOIN ARQUIVOS AS arq ON fila.ID_ARQUIVO = arq.id
                    ORDER BY fila.ID
                """)
            fila = cursor.fetchall()

            print(f"\n📦 Total na fila: {len(fila)}")

            if fila:
                inventor = self.conectar_inventor()

                for id_arquivo, origem, caminho, nome_base in fila:
                    print("\n")
                    print(f"🔍 Abrindo: {caminho}")

                    doc = None

                    try:
                        doc = inventor.Documents.Open(caminho, False)

                        sucesso = self.processar_arquivo(cursor, id_arquivo, caminho, doc, origem, nome_base)

                        if sucesso:
                            cursor.execute("""
                                            DELETE FROM FILA_CONFERENCIA
                                            WHERE ID_ARQUIVO=?
                                        """, (id_arquivo,))
                            conecta_engenharia.commit()

                            print(f"PRODUTO ATUALIZADO COM SUCESSO - SAI DA FILA: {caminho}")
                        else:
                            conecta_engenharia.rollback()
                            print(f"PRODUTO COM PROBLEMAS - FICA NA FILA: {caminho}")

                    except Exception as e:
                        conecta_engenharia.rollback()
                        trata_excecao(e)
                        continue

                    finally:
                        if doc is not None:
                            try:
                                doc.Close(True) # type: ignore
                            except Exception as e:
                                trata_excecao(e)

        except Exception as e:
            trata_excecao(e)
            raise

        finally:
            if inventor is not None:
                try:
                    inventor.Quit()
                except Exception as e:
                    trata_excecao(e)
                    raise

    def processar_arquivo(self, cursor, id_arquivo, caminho, doc, origem, nome_base):
        try:
            print(f"\n📂 Processando: {caminho}")

            props = self.consulta_propriedades_inventor(doc)

            cursor.execute("""
                SELECT TIPO_ARQUIVO FROM ARQUIVOS WHERE ID=?
            """, (id_arquivo,))
            tipo = cursor.fetchone()[0]

            if tipo == "IAM":
                self.consulta_idw_colocar_na_fila(cursor, id_arquivo, caminho)

                dados = {
                    "revision_number": props.get("Revision Number"),
                    "part_number": props.get("Part Number"),
                    "cost_center": props.get("Cost Center"),
                    "description": props.get("Description"),
                    "material": props.get("Material"),
                    "authority": props.get("Authority"),
                    "engineer": props.get("Engineer"),
                }

                erros = []

                for campo, valor in dados.items():
                    if valor is not None and len(str(valor)) > 100:
                        erros.append((campo, valor))

                if erros:
                    msg = []
                    for campo, valor in erros:
                        dadinhos = f" - {campo}: {len(str(valor))} caracteres"
                        msg.append(dadinhos)

                    nome_arquivo = os.path.basename(caminho)

                    self.envia_email_muitos_caracteres(nome_arquivo, msg)

                    return False

                return self.processar_iam(cursor, doc, id_arquivo, dados, origem)
            elif tipo == "IPT":
                self.consulta_idw_colocar_na_fila(cursor, id_arquivo, caminho)

                dados = {
                    "revision_number": props.get("Revision Number"),
                    "part_number": props.get("Part Number"),
                    "cost_center": props.get("Cost Center"),
                    "description": props.get("Description"),
                    "material": props.get("Material"),
                    "vendor": props.get("Vendor"),
                    "authority": props.get("Authority"),
                    "engineer": props.get("Engineer"),
                    "comprimento": props.get("Comprimento")
                }

                erros = []

                for campo, valor in dados.items():
                    if valor is not None and len(str(valor)) > 100:
                        erros.append((campo, valor))

                if erros:
                    msg = []
                    for campo, valor in erros:
                        dadinhos = f" - {campo}: {len(str(valor))} caracteres"
                        msg.append(dadinhos)

                    nome_arquivo = os.path.basename(caminho)

                    self.envia_email_muitos_caracteres(nome_arquivo, msg)

                    return False

                return self.salvar_propriedades_ipt(cursor, id_arquivo, dados)
            elif tipo == "IDW":
                print("arquivo IDW")
                return self.processar_idw(cursor, doc, id_arquivo, caminho)

        except Exception as e:
            trata_excecao(e)

            cursor.execute("""
                            SELECT ID, NOME_BASE, CAMINHO
                            FROM ARQUIVOS
                            WHERE ID = ?
                        """, (id_arquivo,))

            idw_result = cursor.fetchall()
            desenho = idw_result[0][1]
            self.envia_email_desenho_sem_vinculo(idw_result, desenho)

            return False

    def consulta_propriedades_inventor(self, doc):
        try:
            props = {}

            for prop_set in doc.PropertySets:
                for prop in prop_set:

                    if prop.Name == "Thumbnail":
                        continue

                    nome = prop.Name.strip()

                    try:
                        valor = prop.Value
                        props[nome] = str(valor) if valor is not None else None

                    except Exception as e:
                        if "Falha catastrófica" in str(e):
                            continue

                        props[nome] = None

            return props

        except Exception as e:
            trata_excecao(e)
            raise

    def processar_iam(self, cursor, doc, id_arquivo, props_iam, origem):
        try:
            estrutura_nova = self.consulta_estrutura_inventor(doc, cursor, origem, id_arquivo)

            total_itens = len(estrutura_nova)
            props_iam["total_itens"] = total_itens

            self.salvar_propriedades_iam(cursor, id_arquivo, props_iam)

            self.sincronizar_estrutura(cursor, id_arquivo, estrutura_nova)

            return True

        except Exception as e:
            trata_excecao(e)
            return False

    def salvar_propriedades_iam(self, cursor, id_arquivo, props_iam):
        try:
            cursor.execute("""
                SELECT ID FROM PROPRIEDADES_IAM WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            row = cursor.fetchone()

            if row:
                cursor.execute("""
                    UPDATE PROPRIEDADES_IAM
                    SET REVISION_NUMBER=?, 
                    PART_NUMBER=?, 
                    COST_CENTER=?, 
                    DESCRIPTION=?, 
                    MATERIAL=?, 
                    AUTHORITY=?, 
                    ENGINEER=?, 
                    TOTAL_ITENS=?
                    WHERE ID_ARQUIVO=?
                """, (
                        props_iam["revision_number"],
                        props_iam["part_number"],
                        props_iam["cost_center"],
                        props_iam["description"],
                        props_iam["material"],
                        props_iam["authority"],
                        props_iam["engineer"],
                        props_iam["total_itens"],
                        id_arquivo
                    ))
            else:
                print("PRECISA INSERIR!!     ", props_iam)

                cursor.execute("""
                    INSERT INTO PROPRIEDADES_IAM
                    (ID_ARQUIVO, REVISION_NUMBER, PART_NUMBER, COST_CENTER, 
                    DESCRIPTION, MATERIAL, AUTHORITY, ENGINEER, TOTAL_ITENS)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, ( id_arquivo,
                        props_iam["revision_number"],
                        props_iam["part_number"],
                        props_iam["cost_center"],
                        props_iam["description"],
                        props_iam["material"],
                        props_iam["authority"],
                       props_iam["engineer"],
                        props_iam["total_itens"],
                    ))

        except Exception as e:
            trata_excecao(e)
            raise

    def salvar_propriedades_ipt(self, cursor, id_arquivo, props_ipt):
        try:
            cursor.execute("""
                SELECT ID FROM PROPRIEDADES_IPT WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            row = cursor.fetchone()

            if row:
                cursor.execute("""
                    UPDATE PROPRIEDADES_IPT
                    SET REVISION_NUMBER=?, 
                    PART_NUMBER=?, 
                    COST_CENTER=?, 
                    DESCRIPTION=?, 
                    MATERIAL=?, 
                    VENDOR=?, 
                    AUTHORITY=?,
                    ENGINEER=?,
                    COMPRIMENTO=?
                    WHERE ID_ARQUIVO=?
                """, (
                        props_ipt["revision_number"],
                        props_ipt["part_number"],
                        props_ipt["cost_center"],
                        props_ipt["description"],
                        props_ipt["material"],
                        props_ipt["vendor"],
                        props_ipt["authority"],
                        props_ipt["engineer"],
                        props_ipt["comprimento"],
                        id_arquivo
                    ))
            else:
                cursor.execute("""
                    INSERT INTO PROPRIEDADES_IPT
                    (ID_ARQUIVO, REVISION_NUMBER, PART_NUMBER, COST_CENTER, 
                    DESCRIPTION, MATERIAL, VENDOR, AUTHORITY, ENGINEER, COMPRIMENTO)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (id_arquivo,
                      props_ipt["revision_number"],
                      props_ipt["part_number"],
                      props_ipt["cost_center"],
                      props_ipt["description"],
                      props_ipt["material"],
                      props_ipt["vendor"],
                      props_ipt["authority"],
                      props_ipt["engineer"],
                      props_ipt["comprimento"]
                    ))
            return True

        except Exception as e:
            trata_excecao(e)
            return False

    def processar_idw(self, cursor, doc, id_arquivo, caminho):
        try:
            for ref in doc.ReferencedDocuments:
                caminho_ref = ref.FullFileName

                id_ref = self.consulta_e_cria_id_arquivo(cursor, caminho_ref)

                cursor.execute("""
                    DELETE FROM PROPRIEDADES_IDW WHERE ID_ARQUIVO=?
                """, (id_arquivo,))

                cursor.execute("""
                    INSERT INTO PROPRIEDADES_IDW (ID_ARQUIVO, ID_ARQUIVO_REFERENCIA)
                    VALUES (?, ?)
                """, (id_arquivo, id_ref))

                try:
                    # 🔥 limpar cotas
                    cursor.execute("""
                        DELETE FROM COTAS_IDW WHERE ID_ARQUIVO=?
                    """, (id_arquivo,))

                    cotas_unicas = set()  # 🔥 evita repetição

                    for sheet in doc.Sheets:
                        for dim in sheet.DrawingDimensions:

                            try:
                                valor = dim.ModelValue
                                valor_mm = round((valor * 10), 2)
                            except:
                                continue  # 🔥 ignora erro

                            if valor_mm is None:
                                continue  # 🔥 garantia extra

                            cotas_unicas.add(valor_mm)

                    # 🔥 salva só valores únicos
                    for valor_mm in cotas_unicas:
                        cursor.execute("""
                            INSERT INTO COTAS_IDW (ID, ID_ARQUIVO, VALOR_COTA)
                            VALUES (GEN_ID(GEN_COTAS_IDW_ID,1), ?, ?)
                        """, (id_arquivo, valor_mm))

                except Exception as e:
                    print("Erro ao pegar cotas:", e)

                self.inserir_pdf_fila(id_arquivo, caminho)

                break

            return True

        except Exception as e:
            trata_excecao(e)

            cursor.execute("""
                            SELECT ID, NOME_BASE, CAMINHO
                            FROM ARQUIVOS
                            WHERE ID = ?
                        """, (id_arquivo,))

            idw_result = cursor.fetchall()
            desenho = idw_result[0][1]
            self.envia_email_desenho_sem_vinculo(idw_result, desenho)

            return False

    def inserir_pdf_fila(self, id_arquivo, caminho_arquivo):
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                SELECT ID, ID_ARQUIVO
                FROM FILA_GERAR_PDF
                WHERE ID_ARQUIVO = ?
            """, (id_arquivo,))

            tem_na_fila = cursor.fetchall()

            if not tem_na_fila:
                sql = """
                    INSERT INTO FILA_GERAR_PDF (ID, ID_ARQUIVO) 
                    VALUES (GEN_ID(GEN_FILA_GERAR_PDF_ID,1), ?);
                """

                cursor.execute(sql, (id_arquivo,)) # ✅ AQUI TAMBÉM
                conecta_engenharia.commit()

                print("Produto inserido na fila de PDF!", caminho_arquivo)

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    freeze_support()
    WorkerFilaConferencia()