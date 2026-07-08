import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

import os
from datetime import datetime, timedelta
from core.banco import conecta_engenharia
from core.erros import trata_excecao
from core.inventor import pasta_arq, ignorar_pastas, extensoes, padrao_desenho, definir_classificacao
from core.inventor import padronizar_caminho

cursor = conecta_engenharia.cursor()

print("Carregando arquivos do banco...")

cursor.execute("""
SELECT id, caminho, data_mod
FROM arquivos
""")
registros = cursor.fetchall()

# dicionário para comparação rápida
arquivos_banco = {}

for r in registros:
    caminho = padronizar_caminho(r[1])

    if not caminho:
        print("🚨 CAMINHO INVÁLIDO NO BANCO:", r[1])
        continue

    arquivos_banco[caminho] = (r[0], r[2])

arquivos_encontrados = set()

print("Iniciando varredura...\n")

novos = 0
alterados = 0
deletados = 0

def insert_divergencia(dados):
    try:
        id_divergencia, id_arquivo, obs = dados

        cursor = conecta_engenharia.cursor()
        cursor.execute("""
                        SELECT ID_TIPO_DIVERGENCIA, ID_ARQUIVO, OBS
                        FROM DIVERGENCIAS
                        WHERE ID_TIPO_DIVERGENCIA = ? and ID_ARQUIVO = ?
                    """, (id_divergencia, id_arquivo,))
        tem_divergencia = cursor.fetchall()

        if not tem_divergencia:
            sql = """
                    INSERT INTO DIVERGENCIAS (ID_TIPO_DIVERGENCIA, ID_ARQUIVO, OBS)
                    VALUES (?, ?, ?);
                    """
            cursor.execute(sql, (id_divergencia, id_arquivo, obs))

            conecta_engenharia.commit()

            print("\n\n")
            print("Divergencia Inserida com sucesso!", id_divergencia, id_arquivo, obs)
            print("\n\n")

    except Exception as e:
        trata_excecao(e)
        raise

# -------------------------
# FUNÇÃO FILA
# -------------------------
def inserir_fila_conferencia(cursor, id_arquivo, desenho):
    try:
        if desenho.lower() not in ("a4  assembles.idw", "a4  peças.idw"):
            # 🔹 evita duplicado
            cursor.execute("""
                SELECT 1 FROM FILA_CONFERENCIA
                WHERE ID_ARQUIVO = ?
            """, (id_arquivo,))

            if cursor.fetchone():
                print("⚠️ Já está na fila:", id_arquivo)
                return

            # 🔹 insere o próprio arquivo
            cursor.execute("""
                INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO, ORIGEM)
                VALUES (?, ?)
            """, (id_arquivo, "ALTERADOS"))

            print("📥 Inserido na fila:", id_arquivo, desenho)

            # 🔹 pega dados do arquivo
            cursor.execute("""
                SELECT NOME_BASE, TIPO_ARQUIVO
                FROM ARQUIVOS
                WHERE ID = ?
            """, (id_arquivo,))

            row = cursor.fetchone()

            if not row:
                return

            nome_base, tipo = row

            # 🔥 só IPT e IAM precisam garantir IDW
            if tipo not in ("IPT", "IAM"):
                return

            # 🔹 extrai código do desenho
            match = padrao_desenho.search(nome_base)

            if not match:
                return

            codigo = match.group()

            # 🔹 busca IDW pelo código
            cursor.execute("""
                SELECT ID, TIPO_ARQUIVO, CAMINHO
                FROM ARQUIVOS
                WHERE NOME_BASE = ?
                  AND TIPO_ARQUIVO = 'IDW'
            """, (codigo,))

            resultados = cursor.fetchall()

            if len(resultados) == 1:
                id_idw = resultados[0][0]

                # 🔹 evita duplicado
                cursor.execute("""
                    SELECT 1 FROM FILA_CONFERENCIA
                    WHERE ID_ARQUIVO = ?
                """, (id_idw,))

                if not cursor.fetchone():
                    cursor.execute("""
                        INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO, ORIGEM)
                        VALUES (?, ?)
                    """, (id_idw, "ALTERADOS"))

                    print("📄 IDW inserido na fila:", id_idw, desenho)

            elif len(resultados) > 1:
                ids = [linha[0] for linha in resultados]
                dados = (1, id_arquivo, f"VARREDURA - IDS dos arquivos IPT/IAM: {ids}")
                insert_divergencia(dados)

    except Exception as e:
        print("Erro ao inserir na fila:", e)


# -------------------------
# VARREDURA
# -------------------------
for root, dirs, files in os.walk(pasta_arq):

    dirs[:] = [d for d in dirs if d.lower() not in ignorar_pastas]

    for file in files:

        if not file.lower().endswith(extensoes):
            continue

        nome_sem_ext = os.path.splitext(file)[0]

        caminho_original = os.path.join(root, file)
        caminho_certo = padronizar_caminho(caminho_original)

        if not caminho_certo:
            print("🚨 IGNORADO (SCAN):", caminho_original)
            continue

        classificacao = definir_classificacao(caminho_original, nome_sem_ext)

        if not os.path.exists(caminho_original):
            print(f"❌ Caminho inválido: {caminho_original}")
            continue

        stat = os.stat(caminho_original)

        tamanho = stat.st_size
        data_mod = datetime.fromtimestamp(stat.st_mtime).replace(second=0, microsecond=0)
        extensaos = os.path.splitext(file)[1].lower()
        tipo_arquivo = extensaos.replace(".", "").upper()

        arquivos_encontrados.add(caminho_certo)

        # -------------------------
        # ARQUIVO NOVO
        # -------------------------
        if caminho_certo not in arquivos_banco:

            cursor.execute("""
            INSERT INTO arquivos
            (arquivo, NOME_BASE, caminho, tipo_arquivo, classificacao, tamanho, data_mod)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            RETURNING id
            """, (file, nome_sem_ext, caminho_certo, tipo_arquivo, classificacao, tamanho, data_mod))

            id_arquivo = cursor.fetchone()[0]

            print("🆕 NOVO:", file)

            inserir_fila_conferencia(cursor, id_arquivo, file)

            novos += 1

        else:
            id_arquivo, data_banco = arquivos_banco[caminho_certo]
            data_banco = data_banco.replace(second=0, microsecond=0)

            # -------------------------
            # ARQUIVO ALTERADO
            # -------------------------
            if data_mod > data_banco + timedelta(minutes=1):

                cursor.execute("""
                UPDATE arquivos
                SET data_mod = ?, 
                    tamanho = ?, 
                    classificacao = ?, 
                    tipo_arquivo = ?
                WHERE id = ?
                """, (data_mod, tamanho, classificacao, tipo_arquivo, id_arquivo))

                print("🔄 ALTERADO:", file)

                inserir_fila_conferencia(cursor, id_arquivo, file)

                alterados += 1


# -------------------------
# DETECTAR DELETADOS
# -------------------------
for caminho_banco, (id_arquivo, _) in arquivos_banco.items():

    if caminho_banco not in arquivos_encontrados:

        # 1. remove vínculos
        cursor.execute("DELETE FROM estrutura WHERE id_pai = ?", (id_arquivo,))
        cursor.execute("DELETE FROM estrutura WHERE id_filho = ?", (id_arquivo,))

        # 2. remove propriedades
        cursor.execute("DELETE FROM propriedades_ipt WHERE id_arquivo = ?", (id_arquivo,))
        cursor.execute("DELETE FROM propriedades_iam WHERE id_arquivo = ?", (id_arquivo,))
        cursor.execute("DELETE FROM propriedades_idw WHERE id_arquivo = ?", (id_arquivo,))
        cursor.execute("DELETE FROM propriedades_idw WHERE ID_ARQUIVO_REFERENCIA = ?", (id_arquivo,))
        cursor.execute("DELETE FROM COTAS_IDW WHERE id_arquivo = ?", (id_arquivo,))

        cursor.execute("DELETE FROM FILA_CONFERENCIA WHERE id_arquivo = ?", (id_arquivo,))
        cursor.execute("DELETE FROM FILA_LANCA_PROPRIEDADE WHERE id_arquivo = ?", (id_arquivo,))
        cursor.execute("DELETE FROM FILA_GERAR_PDF WHERE id_arquivo = ?", (id_arquivo,))
        cursor.execute("DELETE FROM DIVERGENCIAS WHERE id_arquivo = ?", (id_arquivo,))

        # 3. remove arquivo
        cursor.execute("DELETE FROM arquivos WHERE id = ?", (id_arquivo,))

        print("🗑️ DELETADO:", caminho_banco)

        deletados += 1


# -------------------------
# FINALIZAÇÃO
# -------------------------
conecta_engenharia.commit()

print("\n--------------------------------")
print("Varredura finalizada")
print("Novos:", novos)
print("Alterados:", alterados)
print("Deletados:", deletados)
print("--------------------------------")