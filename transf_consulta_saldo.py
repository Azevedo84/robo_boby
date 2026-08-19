from core.banco import conecta
import pandas as pd


# ============================================================
# CONFIGURAÇÃO
# ============================================================

ARQUIVO_EXCEL = r"C:\Users\Anderson\Desktop\Cópia de Saldo Almox MAQ.xlsx"


# ============================================================
# LER EXCEL
# ============================================================

df = pd.read_excel(ARQUIVO_EXCEL)

cursor = conecta.cursor()

try:

    # ========================================================
    # BUSCAR ID DOS PRODUTOS
    # ========================================================

    resultados = []

    for _, linha in df.iterrows():

        codigo_siger = int(linha["CODIGO"])

        cursor.execute(
            """
            SELECT ID
            FROM PRODUTO
            WHERE CODIGO = ?
            """,
            (codigo_siger,)
        )

        produto = cursor.fetchone()

        if not produto:
            continue

        produto_id = produto[0]


        # ====================================================
        # CONSULTAR SALDO EM TODOS OS LOCAIS
        # ====================================================

        cursor.execute(
            """
            SELECT
                LOCAL_ESTOQUE,
                SALDO
            FROM SALDO_ESTOQUE
            WHERE PRODUTO_ID = ?
              AND SALDO <> 0
            ORDER BY LOCAL_ESTOQUE
            """,
            (produto_id,)
        )

        saldos = cursor.fetchall()


        for local_estoque, saldo in saldos:

            resultados.append({
                "CODIGO": codigo_siger,
                "PRODUTO_ID": produto_id,
                "LOCAL_ESTOQUE": local_estoque,
                "SALDO": saldo
            })


    # ========================================================
    # RESULTADO
    # ========================================================

    print()
    print("==========================================")
    print("SALDOS ENCONTRADOS")
    print("==========================================")

    if not resultados:

        print("Nenhum produto possui saldo em estoque.")

    else:

        for resultado in resultados:

            print(
                f"Código: {resultado['CODIGO']} | "
                f"Produto: {resultado['PRODUTO_ID']} | "
                f"Local: {resultado['LOCAL_ESTOQUE']} | "
                f"Saldo: {resultado['SALDO']}"
            )

        print()
        print(f"Total de registros encontrados: {len(resultados)}")


finally:

    cursor.close()
    conecta.close()