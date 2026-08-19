from core.banco import conecta
import pandas as pd
import re


# ============================================================
# CONFIGURAÇÃO
# ============================================================

ARQUIVO_EXCEL = r"C:\Users\Anderson\Desktop\Cópia de Saldo Almox MAQ.xlsx"

DATA_MOVIMENTO = "2026-08-07"

PRODUTO_SUCATA = 26186
CODIGO_SUCATA = 56391

TIPO_SAIDA = 240
TIPO_ENTRADA = 140

LOCAL_ESTOQUE = 1

DESCRICAO_CI = "TRANSFERENCIA PARA SUCATA"


# ============================================================
# LER EXCEL
# ============================================================

df = pd.read_excel(ARQUIVO_EXCEL)

print()
print("==========================================")
print("ARQUIVO EXCEL")
print("==========================================")
print(f"Linhas encontradas: {len(df)}")


# ============================================================
# VALIDAR COLUNAS
# ============================================================

colunas_necessarias = ["CODIGO", "SALDO", "PESO"]

for coluna in colunas_necessarias:
    if coluna not in df.columns:
        raise Exception(
            f"A coluna '{coluna}' não foi encontrada no Excel."
        )


# ============================================================
# CONEXÃO
# ============================================================

cursor = conecta.cursor()

try:

    # ========================================================
    # 1. DESCOBRIR O CI
    # ========================================================

    sql_ci_dia = """
        SELECT FIRST 1
            ID,
            DATA,
            TIPO,
            PRODUTO,
            QUANTIDADE,
            OBS
        FROM MOVIMENTACAO
        WHERE TIPO IN (140, 240)
          AND DATA = ?
          AND OBS STARTING WITH 'CI '
        ORDER BY ID DESC
    """

    cursor.execute(sql_ci_dia, (DATA_MOVIMENTO,))
    movimento = cursor.fetchone()


    if movimento:

        obs_mov = movimento[5]

        match = re.search(
            r"CI\s+(\d+)",
            obs_mov or "",
            re.IGNORECASE
        )

        if not match:
            raise Exception(
                f"Não foi possível encontrar o CI no OBS:\n{obs_mov}"
            )

        ci = int(match.group(1))

        print()
        print("==========================================")
        print("CI ENCONTRADO NO DIA")
        print("==========================================")
        print(f"CI: {ci}")

    else:

        # ----------------------------------------------------
        # Não existe 140/240 no dia.
        # Procurar o último anterior.
        # ----------------------------------------------------

        sql_ultimo_ci = """
            SELECT FIRST 1
                ID,
                DATA,
                TIPO,
                PRODUTO,
                QUANTIDADE,
                OBS
            FROM MOVIMENTACAO
            WHERE TIPO IN (140, 240)
              AND DATA < ?
              AND OBS STARTING WITH 'CI '
            ORDER BY DATA DESC, ID DESC
        """

        cursor.execute(sql_ultimo_ci, (DATA_MOVIMENTO,))
        movimento = cursor.fetchone()

        if not movimento:
            raise Exception(
                "Não foi encontrada nenhuma movimentação 140/240 "
                "anterior à data do lançamento."
            )

        obs_mov = movimento[5]

        match = re.search(
            r"CI\s+(\d+)",
            obs_mov or "",
            re.IGNORECASE
        )

        if not match:
            raise Exception(
                f"Não foi possível encontrar o CI no OBS:\n{obs_mov}"
            )

        ci_anterior = int(match.group(1))
        ci = ci_anterior + 1

        print()
        print("==========================================")
        print("NOVO CI")
        print("==========================================")
        print(f"CI anterior: {ci_anterior}")
        print(f"Novo CI:     {ci}")


    # ========================================================
    # 2. MONTAR OS LANÇAMENTOS EM MEMÓRIA
    # ========================================================

    lancamentos = []
    nao_encontrados = []

    for indice, linha in df.iterrows():

        codigo_siger = int(linha["CODIGO"])
        saldo = linha["SALDO"]
        peso = linha["PESO"]


        # ----------------------------------------------------
        # Buscar ID interno pelo código Siger
        # ----------------------------------------------------

        sql_produto = """
            SELECT ID
            FROM PRODUTO
            WHERE CODIGO = ?
        """

        cursor.execute(sql_produto, (codigo_siger,))
        produto = cursor.fetchone()

        if not produto:

            nao_encontrados.append(codigo_siger)
            continue

        produto_id = produto[0]

        # ----------------------------------------------------
        # OBS
        # ----------------------------------------------------

        obs = f"CI {ci} - {DESCRICAO_CI} {codigo_siger}/{CODIGO_SUCATA}"


        # ----------------------------------------------------
        # SAÍDA 240
        # ----------------------------------------------------

        lancamentos.append({
            "tipo": TIPO_SAIDA,
            "produto": produto_id,
            "codigo": codigo_siger,
            "quantidade": saldo,
            "obs": obs
        })


        # ----------------------------------------------------
        # ENTRADA 140 - SUCATA
        # ----------------------------------------------------

        lancamentos.append({
            "tipo": TIPO_ENTRADA,
            "produto": PRODUTO_SUCATA,
            "codigo": CODIGO_SUCATA,
            "quantidade": peso,
            "obs": obs
        })


    # ========================================================
    # 3. VALIDAR
    # ========================================================

    print()
    print("==========================================")
    print("RESUMO")
    print("==========================================")

    print(f"Linhas Excel:       {len(df)}")
    print(f"Lançamentos:        {len(lancamentos)}")
    print(f"CI utilizado:       {ci}")
    print(f"Não encontrados:    {len(nao_encontrados)}")


    if nao_encontrados:

        print()
        print("CÓDIGOS NÃO ENCONTRADOS:")

        for codigo in nao_encontrados:
            print(codigo)

        raise Exception(
            "Existem produtos sem cadastro. "
            "Nenhum lançamento deve ser realizado."
        )


    # ========================================================
    # 4. MOSTRAR OS LANÇAMENTOS
    # ========================================================

    print()
    print("==========================================")
    print("LANÇAMENTOS QUE SERIAM FEITOS")
    print("==========================================")


    for numero, lancamento in enumerate(lancamentos, start=1):

        print(
            f"{numero:03d} | "
            f"TIPO {lancamento['tipo']} | "
            f"PRODUTO {lancamento['produto']} | "
            f"CODIGO {lancamento['codigo']} | "
            f"QTD {lancamento['quantidade']} | "
            f"{lancamento['obs']}"
        )

    # ========================================================
    # 5. GRAVAR OS LANÇAMENTOS
    # ========================================================

    print()
    print("==========================================")
    print("GRAVAÇÃO")
    print("==========================================")

    sql_insert = """
                 INSERT INTO MOVIMENTACAO (PRODUTO, \
                                           OBS, \
                                           TIPO, \
                                           QUANTIDADE, \
                                           DATA, \
                                           FUNCIONARIO, \
                                           LOCALESTOQUE, \
                                           CODIGO)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?) \
                 """

    try:

        for numero, lancamento in enumerate(lancamentos, start=1):
            cursor.execute(
                sql_insert,
                (
                    lancamento["produto"],
                    lancamento["obs"],
                    lancamento["tipo"],
                    lancamento["quantidade"],
                    DATA_MOVIMENTO,
                    None,
                    LOCAL_ESTOQUE,
                    lancamento["codigo"]
                )
            )

            print(
                f"{numero:03d} | "
                f"TIPO {lancamento['tipo']} | "
                f"PRODUTO {lancamento['produto']} | "
                f"QTD {lancamento['quantidade']}"
            )

        # ----------------------------------------------------
        # Se chegou aqui, os 188 foram executados
        # ----------------------------------------------------

        conecta.commit()

        print()
        print("==========================================")
        print("SUCESSO")
        print("==========================================")
        print(f"{len(lancamentos)} movimentos gravados.")
        print(f"CI: {ci}")
        print(f"Data: {DATA_MOVIMENTO}")


    except Exception as erro:

        # ----------------------------------------------------
        # Qualquer erro cancela TODOS os movimentos
        # ----------------------------------------------------

        conecta.rollback()

        print()
        print("==========================================")
        print("ERRO - ROLLBACK")
        print("==========================================")
        print("Nenhum movimento foi mantido no banco.")
        print()
        print(f"Erro: {erro}")

        raise


finally:

    cursor.close()
    conecta.close()