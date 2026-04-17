from core.erros import trata_excecao
from core.email_service import envia_email_desenho_duplicado, envia_email_desenho_sem_vinculo
from core.email_service import envia_email_sem_idw, envia_email_arquivo_nao_encontrado
from core.email_service import dados_email
import os
import win32com.client
from core.banco import conecta, conecta_engenharia
from core.inventor import definir_classificacao
from core.inventor import padronizar_caminho
from datetime import datetime
import re
from multiprocessing import Process, freeze_support
import time
from typing import cast, Any
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


class LerPedidosInternos:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

        self.manipula_comeco()

    def buscar_idw_por_referencia(self, cursor, ref):
        try:
            # 🔹 1. tentativa - nome exato
            cursor.execute("""
                SELECT ID, TIPO_ARQUIVO, CAMINHO
                FROM ARQUIVOS
                WHERE NOME_BASE = ?
                  AND TIPO_ARQUIVO = 'IDW'
            """, (ref,))

            resultados = cursor.fetchall()

            if resultados:
                return resultados

            # 🔹 2. tentativa - com prefixo "XX - "
            cursor.execute("""
                SELECT ID, TIPO_ARQUIVO, CAMINHO
                FROM ARQUIVOS
                WHERE NOME_BASE LIKE ?
                  AND TIPO_ARQUIVO = 'IDW'
            """, (f"__ - {ref}",))  # exatamente 2 caracteres

            resultados = cursor.fetchall()

            if resultados:
                return resultados

            return []

        except Exception as e:
            trata_excecao(e)
            raise

    def envia_email_muitos_caracteres(self, desenho, campos):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA (FILA) - CAMPOS COM MAIS DE 100 CARACTERES {desenho}'

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

    def extrair_referencia(self, obs):
        try:
            if not obs:
                return None

            s = re.sub(r"[^\d.]", "", obs)
            s = re.sub(r"\.+$", "", s)

            return s if s else None

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_caminho_arquivo(self, cursor, id_arquivo):
        try:
            cursor.execute("""
                SELECT CAMINHO FROM ARQUIVOS WHERE ID=?
            """, (id_arquivo,))
            return cursor.fetchone()[0]

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_propriedades_inventor(self, doc):
        try:
            props = {}

            for prop_set in doc.PropertySets:
                for prop in prop_set:
                    try:
                        props[prop.Name] = str(prop.Value)

                    except Exception as e:
                        trata_excecao(e)

            return props

        except Exception as e:
            trata_excecao(e)
            raise

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
                cursor.execute("""
                                INSERT INTO PROPRIEDADES_IAM
                                (ID_ARQUIVO, REVISION_NUMBER, PART_NUMBER, COST_CENTER, 
                                DESCRIPTION, MATERIAL, AUTHORITY, ENGINEER, TOTAL_ITENS)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (id_arquivo,
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

    def consulta_e_cria_id_arquivo(self, cursor, caminho):
        try:
            cursor.execute("""
                SELECT ID FROM ARQUIVOS WHERE LOWER(CAMINHO)=?
            """, (caminho,))

            row = cursor.fetchone()

            if row:
                return row[0]

            # 🔹 dados do arquivo (teu padrão)
            nome: str = os.path.basename(caminho)
            nome_sem_ext = os.path.splitext(nome)[0]

            classificacao = definir_classificacao(caminho, nome_sem_ext)

            try:
                stat = os.stat(caminho)
            except FileNotFoundError:
                raise Exception(f"Arquivo não existe fisicamente: {caminho}")

            tamanho = stat.st_size
            data_mod = datetime.fromtimestamp(stat.st_mtime).replace(second=0, microsecond=0)

            extensao = os.path.splitext(nome)[1].lower()
            tipo_arquivo = extensao.replace(".", "").upper()

            # 🔹 INSERT completo (sem invenção)
            cursor.execute("""
                INSERT INTO ARQUIVOS (
                    ARQUIVO, 
                    NOME_BASE, 
                    TIPO_ARQUIVO,
                    CLASSIFICACAO,
                    CAMINHO,
                    TAMANHO,
                    DATA_MOD
                )
                VALUES (
                    ?, ?, ?, ?, ?, ?, ? 
                )
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

            print(f"🆕 Arquivo cadastrado: {caminho}")

            return cursor.fetchone()[0]

        except Exception as e:
            trata_excecao(e)
            raise

    def inserir_fila_conferencia(self, cursor, id_arquivo):
        try:
            # 🔹 busca somente IPT/IAM
            cursor.execute("""
                        SELECT ID, TIPO_ARQUIVO
                        FROM ARQUIVOS
                        WHERE ID = ?
                    """, (id_arquivo,))

            resultados = cursor.fetchall()

            id_arquivo, tipo = resultados[0]

            cursor.execute("""
                SELECT 1 FROM FILA_CONFERENCIA
                WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            if cursor.fetchone():
                print("⚠️ Já está na fila")
                return

            # 🔹 verifica se já processado
            ja_processado = False

            if tipo == "IPT":
                cursor.execute("""
                        SELECT 1 FROM PROPRIEDADES_IPT WHERE ID_ARQUIVO=?
                    """, (id_arquivo,))
                ja_processado = cursor.fetchone() is not None

            elif tipo == "IAM":
                cursor.execute("""
                        SELECT 1 FROM PROPRIEDADES_IAM WHERE ID_ARQUIVO=?
                    """, (id_arquivo,))
                tem_props = cursor.fetchone() is not None

                cursor.execute("""
                        SELECT 1 FROM ESTRUTURA WHERE ID_PAI=?
                    """, (id_arquivo,))
                tem_estrutura = cursor.fetchone() is not None

                ja_processado = tem_props and tem_estrutura

            if not ja_processado:
                cursor.execute("""
                    INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO, ORIGEM)
                    VALUES (?. ?)
                """, (id_arquivo, "PEDIDOS"))

                print("✅ Inserido na fila de conferência", id_arquivo, tipo)

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

        except Exception as e:
            trata_excecao(e)
            raise

    def inicio_de_tudo(self, cursor_engenharia):
        try:
            cursor_erp = conecta.cursor()

            cursor_erp.execute("""
                SELECT prod.codigo, prod.obs
                FROM PRODUTOPEDIDOINTERNO prodint
                JOIN produto prod ON prodint.id_produto = prod.id
                WHERE prodint.status = 'A'
            """)

            fila_execucao = []

            for codigo, obs in cursor_erp.fetchall():
                ja_processado = False

                ref = self.extrair_referencia(obs)

                if not ref:
                    print(f"⚠️ Produto {codigo} sem referência válida")
                    continue

                # 🔹 busca IPT/IAM
                cursor_engenharia.execute("""
                    SELECT ID, TIPO_ARQUIVO, CAMINHO
                    FROM ARQUIVOS
                    WHERE NOME_BASE = ?
                      AND TIPO_ARQUIVO IN ('IPT', 'IAM')
                """, (ref,))

                resultados = cursor_engenharia.fetchall()

                if not resultados:
                    print(f"❌ Produto {codigo}: nenhum IPT/IAM para ({ref})")
                    continue

                if len(resultados) > 1:
                    envia_email_desenho_duplicado(self.destinatario, resultados, ref)
                    continue

                id_arquivo, tipo, caminho = resultados[0]

                caminho = padronizar_caminho(caminho)

                if not caminho:
                    print("🚨 CAMINHO INVÁLIDO (PEDIDOS)")
                    continue

                # 🔹 verifica se IPT/IAM já processado
                if tipo == "IPT":
                    cursor_engenharia.execute("""
                        SELECT 1 FROM PROPRIEDADES_IPT WHERE ID_ARQUIVO=?
                    """, (id_arquivo,))
                    ja_processado = cursor_engenharia.fetchone() is not None

                elif tipo == "IAM":
                    cursor_engenharia.execute("""
                        SELECT 1 FROM PROPRIEDADES_IAM WHERE ID_ARQUIVO=?
                    """, (id_arquivo,))
                    tem_props = cursor_engenharia.fetchone() is not None

                    cursor_engenharia.execute("""
                        SELECT 1 FROM ESTRUTURA WHERE ID_PAI=?
                    """, (id_arquivo,))
                    tem_estrutura = cursor_engenharia.fetchone() is not None

                    ja_processado = tem_props and tem_estrutura

                # 🔹 PRIMEIRO adiciona IPT/IAM
                if not ja_processado:
                    fila_execucao.append(id_arquivo)

                # 🔹 DEPOIS adiciona IDW
                idw_result = self.buscar_idw_por_referencia(cursor_engenharia, ref)

                if not idw_result:
                    print(f"⚠️ Produto {codigo}: não possui IDW ({ref})")
                    envia_email_sem_idw(self.destinatario, ref)

                elif len(idw_result) > 1:
                    envia_email_desenho_duplicado(self.destinatario, idw_result, ref)

                else:
                    id_idw = idw_result[0][0]

                    cursor_engenharia.execute("""
                                            SELECT 1 FROM PROPRIEDADES_IDW WHERE ID_ARQUIVO=?
                                        """, (id_idw,))
                    tem_props_idw = cursor_engenharia.fetchone()
                    if not tem_props_idw:
                        print("NÃO ESTA SALVO!! ", id_idw)
                        fila_execucao.append(id_idw)

            # 🔹 remove duplicados mantendo ordem
            fila_execucao = list(dict.fromkeys(fila_execucao))

            print(f"\n📦 Total na fila: {len(fila_execucao)}")

            # 🔹 EXECUÇÃO SEQUENCIAL COM TIMEOUT
            for id_arquivo in fila_execucao:

                print(f"\n🚀 inicio_de_tudo - Iniciando ID: {id_arquivo}")

                p = Process(target=self.processar_arquivo_unitario, args=(id_arquivo,))
                p.start()

                tempo_max = 60
                inicio = time.time()

                while p.is_alive():
                    tempo = time.time() - inicio

                    if tempo > tempo_max:
                        print(f"⏱️ TIMEOUT! Arquivo {id_arquivo} travou ({int(tempo)}s)")
                        print("💀 Matando processo...")
                        # 🔹 DEPOIS adiciona IDW
                        cursor_engenharia.execute("""
                                            SELECT ID, NOME_BASE, CAMINHO
                                            FROM ARQUIVOS
                                            WHERE ID = ?
                                        """, (id_arquivo,))

                        idw_result = cursor_engenharia.fetchall()
                        desenho = idw_result[0][1]
                        envia_email_desenho_sem_vinculo(self.destinatario, idw_result, desenho)

                        try:
                            p.terminate()
                            p.join(5)

                        except Exception as e:
                            trata_excecao(e)

                        os.system("taskkill /f /im Inventor.exe")

                        break

                    time.sleep(1)

                else:
                    print(f"✅ Finalizado: {id_arquivo}")

        except Exception as e:
            trata_excecao(e)
            raise

    def consultar_idw_existe(self, cursor_engenharia, codigo, ref):
        try:
            idw_result = self.buscar_idw_por_referencia(cursor_engenharia, ref)

            if not idw_result:
                print(f"⚠️ Produto {codigo}: não possui IDW ({ref})")
                return None

            elif len(idw_result) > 1:
                envia_email_desenho_duplicado(self.destinatario, idw_result, ref)
                return None

            else:
                id_idw, caminho_idw = idw_result[0]
                return id_idw

        except Exception as e:
            trata_excecao(e)
            raise

    def processar(self, id_arquivo):
        try:
            print(f"\n🚀 processar Iniciando ID: {id_arquivo}")

            p = Process(target=self.processar_arquivo_unitario, args=(id_arquivo,))
            p.start()

            tempo_max = 60  # ⏱️ ajusta aqui se quiser

            inicio = time.time()

            while p.is_alive():
                tempo = time.time() - inicio

                if tempo > tempo_max:
                    print(f"⏱️ TIMEOUT! Arquivo {id_arquivo} travou ({int(tempo)}s)")
                    print("💀 Matando processo e seguindo...")

                    p.terminate()
                    p.join()

                    break

                time.sleep(1)

            else:
                print(f"✅ Finalizado: {id_arquivo}")

        except Exception as e:
            trata_excecao(e)
            raise

    def processar_arquivo_unitario(self, id_arquivo):
        try:
            cursor = conecta_engenharia.cursor()

            inventor = None
            doc: Any = None
            caminho = None

            try:
                inventor = win32com.client.Dispatch("Inventor.Application")
                inventor.Visible = False
                inventor.SilentOperation = True

                caminho = self.consulta_caminho_arquivo(cursor, id_arquivo)

                print(f"🔧 Abrindo: {caminho}")

                doc = inventor.Documents.Open(caminho, False)

                self.processar_arquivo(cursor, id_arquivo, caminho, doc)

            except Exception as e:
                print(f"❌ ERRO INTERNO")
                print(f"🆔 ID: {id_arquivo}")
                print(f"📁 Caminho: {caminho}")
                print(f"💥 Erro: {str(e)}")

                try:
                    conecta_engenharia.rollback()
                except Exception as e:
                    trata_excecao(e)

            finally:
                if doc:
                    try:
                        doc.Close(True)

                    except Exception as e:
                        trata_excecao(e)

                if inventor:
                    try:
                        inventor.Quit()

                    except Exception as e:
                        trata_excecao(e)

                try:
                    conecta_engenharia.commit()

                except Exception as e:
                    trata_excecao(e)

        except Exception as e:
            trata_excecao(e)
            raise

    def processar_arquivo(self, cursor, id_arquivo, caminho, doc):
        try:
            print(f"\n📂 Processando: {caminho}")  # 👈 AQUI

            props = self.consulta_propriedades_inventor(doc)

            cursor.execute("""
                SELECT TIPO_ARQUIVO FROM ARQUIVOS WHERE ID=?
            """, (id_arquivo,))
            tipo = cursor.fetchone()[0]

            if tipo == "IAM":
                print("👉 Tipo: IAM")
                dados_iam = {
                    "revision_number": props.get("Revision Number"),
                    "part_number": props.get("Part Number"),
                    "cost_center": props.get("Cost Center"),
                    "description": props.get("Description"),
                    "material": props.get("Material"),
                    "authority": props.get("Authority"),
                    "engineer": props.get("Engineer")
                }

                erros = []

                for campo, valor in dados_iam.items():
                    if valor is not None and len(str(valor)) > 100:
                        erros.append((campo, valor))

                if erros:
                    msg = []
                    for campo, valor in erros:
                        dadinhos = f" - {campo}: {len(str(valor))} caracteres"
                        msg.append(dadinhos)

                    nome_arquivo = os.path.basename(caminho)

                    self.envia_email_muitos_caracteres(nome_arquivo, msg)

                    return None

                self.processar_iam(cursor, doc, id_arquivo, dados_iam)
            elif tipo == "IPT":
                print("👉 Tipo: IPT")

                dados_ipt = {
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

                for campo, valor in dados_ipt.items():
                    if valor is not None and len(str(valor)) > 100:
                        erros.append((campo, valor))

                if erros:
                    msg = []
                    for campo, valor in erros:
                        dadinhos = f" - {campo}: {len(str(valor))} caracteres"
                        msg.append(dadinhos)

                    nome_arquivo = os.path.basename(caminho)

                    self.envia_email_muitos_caracteres(nome_arquivo, msg)

                    return None
                self.salvar_propriedades_ipt(cursor, id_arquivo, dados_ipt)
            elif tipo == "IDW":
                print("👉 Tipo: IDW")
                self.processar_idw(cursor, doc, id_arquivo, caminho)

        except Exception as e:
            trata_excecao(e)
            raise

    def processar_iam(self, cursor, doc, id_arquivo, props_iam):
        try:
            estrutura_nova = self.consulta_estrutura_inventor(doc, cursor)

            total_itens = len(estrutura_nova)
            props_iam["total_itens"] = total_itens

            self.salvar_propriedades_iam(cursor, id_arquivo, props_iam)

            inseridos, atualizados, deletados = self.sincronizar_estrutura(cursor, id_arquivo, estrutura_nova)

            if inseridos == 0 and atualizados == 0 and deletados == 0:
                print("🟢 Nenhuma alteração na estrutura")
            else:
                print(f"🟡 Alterações → +{inseridos} inseridos | ~{atualizados} atualizados | -{deletados} removidos")

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_estrutura_inventor(self, doc, cursor):
        try:
            comp = doc.ComponentDefinition
            bom = comp.BOM

            bom.StructuredViewEnabled = True
            bom.StructuredViewFirstLevelOnly = False

            estrutura = {}

            view = None
            for v in bom.BOMViews:
                if v.ViewType == 62465:
                    view = v
                    break

            if view is None:
                return {}

            view = cast(Any, view)

            for row in view.BOMRows:
                if row.ComponentDefinitions.Count == 0:
                    continue

                comp_def = row.ComponentDefinitions.Item(1)
                caminho = comp_def.Document.FullFileName

                caminho = padronizar_caminho(caminho)

                if not caminho:
                    print("🚨 CAMINHO INVÁLIDO")
                    continue

                id_filho = self.consulta_e_cria_id_arquivo(cursor, caminho)

                if not id_filho:
                    envia_email_arquivo_nao_encontrado(self.destinatario, caminho)
                    raise Exception(f"Arquivo não cadastrado: {caminho}")

                self.inserir_fila_conferencia(cursor, id_filho)

                try:
                    quantidade = float(row.ItemQuantity)
                except (TypeError, ValueError):
                    quantidade = 0.0

                estrutura[id_filho] = quantidade

            return estrutura

        except Exception as e:
            trata_excecao(e)
            raise

    def sincronizar_estrutura(self, cursor, id_pai, estrutura_nova):
        try:
            estrutura_atual = self.consulta_estrutura_atual(cursor, id_pai)

            inseridos = 0
            atualizados = 0
            deletados = 0

            # 🔹 INSERT / UPDATE
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

            # 🔻 DELETE
            for id_filho in estrutura_atual:

                if id_filho not in estrutura_nova:
                    cursor.execute("""
                        DELETE FROM ESTRUTURA
                        WHERE ID_PAI=? AND ID_FILHO=?
                    """, (id_pai, id_filho))
                    deletados += 1

            return inseridos, atualizados, deletados

        except Exception as e:
            trata_excecao(e)
            raise

    def processar_idw(self, cursor, doc, id_arquivo, caminho):
        try:
            refs = list(doc.ReferencedDocuments)

            if not refs:
                print("⚠️ IDW sem referência")
                envia_email_arquivo_nao_encontrado(self.destinatario, caminho)
                return  # 👈 SEGUE O FLUXO

            ref = refs[0]
            caminho_ref = ref.FullFileName

            if not caminho_ref:
                print("⚠️ Caminho de referência vazio")
                envia_email_arquivo_nao_encontrado(self.destinatario, caminho)
                return  # 👈 SEGUE

            id_referencia = self.consulta_e_cria_id_arquivo(cursor, caminho_ref)

            if not id_referencia:
                print(f"⚠️ Referência não cadastrada: {caminho_ref}")
                envia_email_arquivo_nao_encontrado(self.destinatario, caminho)
                return  # 👈 SEGUE

            cursor.execute("""
                SELECT ID FROM PROPRIEDADES_IDW WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            row = cursor.fetchone()

            if row:
                cursor.execute("""
                    UPDATE PROPRIEDADES_IDW
                    SET ID_ARQUIVO_REFERENCIA=?
                    WHERE ID_ARQUIVO=?
                """, (id_referencia, id_arquivo))
            else:
                cursor.execute("""
                    INSERT INTO PROPRIEDADES_IDW (ID_ARQUIVO, ID_ARQUIVO_REFERENCIA)
                    VALUES (?, ?)
                """, (id_arquivo, id_referencia))

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_comeco(self):
        try:
            freeze_support()  # 👈 ESSA LINHA É O QUE FALTA

            try:
                cursor = conecta_engenharia.cursor()

                try:
                    self.inicio_de_tudo(cursor)

                    conecta_engenharia.commit()

                except Exception as e:
                    print("ERRO:", e)
                    conecta_engenharia.rollback()

            except Exception as e:
                print("ERRO:", e)

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    freeze_support()
    LerPedidosInternos()
