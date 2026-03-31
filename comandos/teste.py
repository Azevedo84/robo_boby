import os
import win32com.client
from banco_dados.conexao import conecta, conecta_engenharia
from comandos.inventor import padrao_desenho
from datetime import datetime
import re
from multiprocessing import Process, freeze_support
import time

def extrair_referencia(obs):
    if not obs:
        return None

    s = re.sub(r"[^\d.]", "", obs)
    s = re.sub(r"\.+$", "", s)

    return s if s else None

def alimentar_fila_por_pedidos(cursor_engenharia):
    cursor_erp = conecta.cursor()

    cursor_erp.execute("""
        SELECT prod.codigo, prod.obs
        FROM PRODUTOPEDIDOINTERNO prodint
        JOIN produto prod ON prodint.id_produto = prod.id
        WHERE prodint.status = 'A'
    """)

    for codigo, obs in cursor_erp.fetchall():

        ref = extrair_referencia(obs)

        if not ref:
            print(f"⚠️ Produto {codigo} sem referência válida")
            continue

        # 🔹 busca somente IPT/IAM
        cursor_engenharia.execute("""
            SELECT ID, TIPO_ARQUIVO
            FROM ARQUIVOS
            WHERE NOME_BASE = ?
              AND TIPO_ARQUIVO IN ('IPT', 'IAM')
        """, (ref,))

        resultados = cursor_engenharia.fetchall()

        if not resultados:
            print(f"❌ Produto {codigo}: nenhum IPT/IAM para ({ref})")
            continue

        if len(resultados) > 1:
            print(f"❌ Produto {codigo}: mais de um IPT/IAM para ({ref})")
            continue

        id_arquivo, tipo = resultados[0]

        # 🔹 verifica se já processado
        ja_processado = False

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

        if ja_processado:
            print(f"⏭️ Já processado: {codigo}")
            continue

        inserir_fila_conferencia(cursor_engenharia, id_arquivo)

def get_fila(cursor):
    cursor.execute("""
        SELECT ID_ARQUIVO FROM FILA_CONFERENCIA
    """)
    return [row[0] for row in cursor.fetchall()]

def consulta_estrutura_atual(cursor, id_pai):

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

def inserir_fila_conferencia(cursor, id_arquivo):

    cursor.execute("""
        SELECT 1 FROM FILA_CONFERENCIA
        WHERE ID_ARQUIVO=?
    """, (id_arquivo,))

    if cursor.fetchone():
        print("⚠️ Já está na fila")
        return

    cursor.execute("""
        INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO)
        VALUES (?)
    """, (id_arquivo,))

    print("✅ Inserido na fila de conferência")

def sincronizar_estrutura(cursor, id_pai, estrutura_nova):

    estrutura_atual = consulta_estrutura_atual(cursor, id_pai)

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

def consulta_estrutura_inventor(doc, cursor):

    comp = doc.ComponentDefinition
    bom = comp.BOM

    bom.StructuredViewEnabled = True
    bom.StructuredViewFirstLevelOnly = False

    view = None
    for v in bom.BOMViews:
        if v.ViewType == 62465:
            view = v

    estrutura = {}

    if not view:
        return {}

    for row in view.BOMRows:

        if row.ComponentDefinitions.Count == 0:
            continue

        comp_def = row.ComponentDefinitions.Item(1)
        caminho = comp_def.Document.FullFileName

        id_filho = consulta_e_cria_id_arquivo(cursor, caminho)

        if not id_filho:
            raise Exception(f"Arquivo não cadastrado: {caminho}")

        estrutura[id_filho] = float(row.ItemQuantity)

    return estrutura

def consulta_e_cria_id_arquivo(cursor, caminho):
    caminho_original = caminho
    caminho = caminho.lower()

    cursor.execute("""
        SELECT ID FROM ARQUIVOS WHERE CAMINHO=?
    """, (caminho,))

    row = cursor.fetchone()

    if row:
        return row[0]

    # 🔹 dados do arquivo (teu padrão)
    nome = os.path.basename(caminho_original)
    nome_sem_ext = os.path.splitext(nome)[0]

    if padrao_desenho.search(nome_sem_ext):
        classificacao = "ACABADO"
    else:
        classificacao = "MATERIA_PRIMA"

    try:
        stat = os.stat(caminho_original)
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
            EXTENSAO,
            TIPO_ARQUIVO,
            CLASSIFICACAO,
            CAMINHO,
            TAMANHO,
            DATA_MOD
        )
        VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?
        )
        RETURNING ID
    """, (
        nome,
        nome_sem_ext,
        extensao,
        tipo_arquivo,
        classificacao,
        caminho,
        tamanho,
        data_mod
    ))

    print(f"🆕 Arquivo cadastrado: {caminho}")

    return cursor.fetchone()[0]

def consulta_propriedades_inventor(doc):

    props = {}

    for prop_set in doc.PropertySets:
        for prop in prop_set:
            try:
                props[prop.Name] = str(prop.Value)
            except:
                pass

    return props

def processar_arquivo(cursor, id_arquivo, caminho, doc):
    print(f"\n📂 Processando: {caminho}")  # 👈 AQUI

    props = consulta_propriedades_inventor(doc)

    ext = caminho.lower()

    if ext.endswith(".iam"):
        print("👉 Tipo: IAM")
        processar_iam(cursor, doc, id_arquivo, props)

    elif ext.endswith(".ipt"):
        print("👉 Tipo: IPT")
        processar_ipt(cursor, id_arquivo, props)

    elif ext.endswith(".idw"):
        print("👉 Tipo: IDW")
        processar_idw(cursor, doc, id_arquivo)

def processar_iam(cursor, doc, id_arquivo, props):

    estrutura_nova = consulta_estrutura_inventor(doc, cursor)

    total_itens = len(estrutura_nova)

    salvar_propriedades_iam(cursor, id_arquivo, props, total_itens)

    inseridos, atualizados, deletados = sincronizar_estrutura(cursor, id_arquivo, estrutura_nova)

    if inseridos == 0 and atualizados == 0 and deletados == 0:
        print("🟢 Nenhuma alteração na estrutura")
    else:
        print(f"🟡 Alterações → +{inseridos} inseridos | ~{atualizados} atualizados | -{deletados} removidos")

def processar_ipt(cursor, id_arquivo, props):

    salvar_propriedades_ipt(cursor, id_arquivo, props)

def processar_idw(cursor, doc, id_arquivo):

    id_referencia = None

    for ref in doc.ReferencedDocuments:
        caminho_ref = ref.FullFileName
        id_referencia = consulta_e_cria_id_arquivo(cursor, caminho_ref)
        break  # pega o principal

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

def salvar_propriedades_iam(cursor, id_arquivo, props, total_itens):

    cursor.execute("""
        SELECT ID FROM PROPRIEDADES_IAM WHERE ID_ARQUIVO=?
    """, (id_arquivo,))

    row = cursor.fetchone()

    dados = (
        props.get("Revision Number"),
        props.get("Part Number"),
        props.get("Cost Center"),
        props.get("Description"),
        props.get("Material"),
        props.get("Authority"),
        total_itens
    )

    if row:
        cursor.execute("""
            UPDATE PROPRIEDADES_IAM
            SET REVISION_NUMBER=?,
                PART_NUMBER=?,
                COST_CENTER=?,
                DESCRIPTION=?,
                MATERIAL=?,
                AUTHORITY=?,
                TOTAL_ITENS=?
            WHERE ID_ARQUIVO=?
        """, (*dados, id_arquivo))
    else:
        cursor.execute("""
            INSERT INTO PROPRIEDADES_IAM
            (ID_ARQUIVO, REVISION_NUMBER, PART_NUMBER, COST_CENTER, DESCRIPTION, MATERIAL, AUTHORITY, TOTAL_ITENS)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (id_arquivo, *dados))

def salvar_propriedades_ipt(cursor, id_arquivo, props):

    cursor.execute("""
        SELECT ID FROM PROPRIEDADES_IPT WHERE ID_ARQUIVO=?
    """, (id_arquivo,))

    row = cursor.fetchone()

    dados = (
        props.get("Revision Number"),
        props.get("Part Number"),
        props.get("Cost Center"),
        props.get("Description"),
        props.get("Material"),
        props.get("Vendor"),
        props.get("Authority"),
        props.get("Comprimento"),  # 👈 nome exato do Inventor
    )

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
                COMPRIMENTO=?
            WHERE ID_ARQUIVO=?
        """, (*dados, id_arquivo))
    else:
        cursor.execute("""
            INSERT INTO PROPRIEDADES_IPT
            (ID_ARQUIVO, REVISION_NUMBER, PART_NUMBER, COST_CENTER, DESCRIPTION, MATERIAL, VENDOR, AUTHORITY, COMPRIMENTO)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (id_arquivo, *dados))

def get_caminho(cursor, id_arquivo):
    cursor.execute("""
        SELECT CAMINHO FROM ARQUIVOS WHERE ID=?
    """, (id_arquivo,))
    return cursor.fetchone()[0]

def processar_arquivo_unitario(id_arquivo):
    cursor = conecta_engenharia.cursor()

    inventor = None
    doc = None
    caminho = None

    try:
        inventor = win32com.client.Dispatch("Inventor.Application")
        inventor.Visible = False
        inventor.SilentOperation = True

        caminho = get_caminho(cursor, id_arquivo)

        print(f"🔧 Abrindo: {caminho}")

        doc = inventor.Documents.Open(caminho, False)

        processar_arquivo(cursor, id_arquivo, caminho, doc)

    except Exception as e:
        print(f"❌ ERRO INTERNO")
        print(f"🆔 ID: {id_arquivo}")
        print(f"📁 Caminho: {caminho}")
        print(f"💥 Erro: {str(e)}")

        try:
            conecta_engenharia.rollback()
        except:
            pass

    finally:
        if doc:
            try:
                doc.Close(True)
            except:
                pass

        if inventor:
            try:
                inventor.Quit()
            except:
                pass

        try:
            conecta_engenharia.commit()  # 👈 ISSO AQUI FALTAVA
        except:
            pass

def processar_lote(cursor):

    fila_ids = get_fila(cursor)

    if not fila_ids:
        print("📭 Fila vazia")
        return

    print(f"📦 Total para processar: {len(fila_ids)}")

    processados = []

    for id_arquivo in fila_ids:

        print(f"\n🚀 Iniciando ID: {id_arquivo}")

        p = Process(target=processar_arquivo_unitario, args=(id_arquivo,))
        p.start()

        TEMPO_MAX = 60  # ⏱️ ajusta aqui se quiser

        inicio = time.time()

        while p.is_alive():
            tempo = time.time() - inicio

            if tempo > TEMPO_MAX:
                print(f"⏱️ TIMEOUT! Arquivo {id_arquivo} travou ({int(tempo)}s)")
                print("💀 Matando processo e seguindo...")

                p.terminate()
                p.join()

                break

            time.sleep(1)

        else:
            print(f"✅ Finalizado: {id_arquivo}")
            processados.append(id_arquivo)

    # 🔹 remove só os que deram certo
    for id_arquivo in processados:
        cursor.execute("""
            DELETE FROM FILA_CONFERENCIA
            WHERE ID_ARQUIVO=?
        """, (id_arquivo,))

    print("🗑️ Limpeza da fila concluída")

inventor = None
doc = None

if __name__ == "__main__":
    freeze_support()  # 👈 ESSA LINHA É O QUE FALTA

    try:
        cursor = conecta_engenharia.cursor()

        try:
            alimentar_fila_por_pedidos(cursor)
            processar_lote(cursor)

            conecta_engenharia.commit()

        except Exception as e:
            print("ERRO:", e)
            conecta_engenharia.rollback()

    except Exception as e:
        print("ERRO:", e)