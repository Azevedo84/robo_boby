from core.banco import conecta
from core.erros import trata_excecao

def calcular_percentual_producao(id_produto):
    try:
        cursor = conecta.cursor()

        total_itens = 0
        total_consumidos = 0

        def analisar_produto(numero_op):

            nonlocal total_itens
            nonlocal total_consumidos

            # =====================================================
            # OP DO PRODUTO
            # =====================================================

            cursor.execute("""
                SELECT
                    QUANTIDADE,
                    ID_ESTRUTURA
                FROM ORDEMSERVICO
                WHERE
                    NUMERO = ?
            """, (numero_op,))

            dados_op = cursor.fetchone()

            if not dados_op:
                return

            quantidade_op, id_estrutura = dados_op

            quantidade_op = quantidade_op or 1

            if not id_estrutura:
                return

            # =====================================================
            # ITENS DA ESTRUTURA
            # =====================================================

            cursor.execute("""
                SELECT
                    EP.ID,
                    EP.ID_PROD_FILHO,
                    EP.QUANTIDADE,
                    P.CONJUNTO,
                    P.CODIGO
                FROM ESTRUTURA_PRODUTO EP

                INNER JOIN PRODUTO P
                    ON P.ID = EP.ID_PROD_FILHO

                WHERE
                    EP.ID_ESTRUTURA = ?
            """, (id_estrutura,))

            itens = cursor.fetchall()

            for (
                id_estrutura_produto,
                id_filho,
                quantidade_estrutura,
                conjunto_filho,
                codigo_filho
            ) in itens:

                quantidade_estrutura = quantidade_estrutura or 0

                # =================================================
                # CADA LINHA DA ESTRUTURA = 1 ITEM
                # =================================================

                total_itens += 1
                print(numero_op, itens, total_itens)

                quantidade_necessaria = (
                    quantidade_estrutura * quantidade_op
                )

                # =================================================
                # CONSUMO DESTE ITEM NA OP
                # =================================================

                cursor.execute("""
                    SELECT
                        COALESCE(SUM(QTDE_ESTRUT_PROD), 0)
                    FROM PRODUTOOS
                    WHERE
                        NUMERO = ?
                        AND ID_ESTRUT_PROD = ?
                """, (
                    numero_op,
                    id_estrutura_produto
                ))

                resultado = cursor.fetchone()

                consumido = resultado[0] or 0

                # =================================================
                # ITEM COMPLETO
                # =================================================

                if consumido >= quantidade_necessaria:
                    total_consumidos += 1

                # =================================================
                # SE FOR CONJUNTO, ANALISA A OP DELE
                # =================================================

                if conjunto_filho == 10:

                    cursor.execute("""
                        SELECT
                            NUMERO
                        FROM ORDEMSERVICO
                        WHERE
                            PRODUTO = ?
                            AND STATUS = 'A'
                        ORDER BY ID DESC
                    """, (id_filho,))

                    op_filho = cursor.fetchone()

                    if op_filho:

                        analisar_produto(
                            op_filho[0]
                        )

        # =========================================================
        # ENCONTRA A OP PRINCIPAL
        # =========================================================

        cursor.execute("""
            SELECT
                NUMERO
            FROM ORDEMSERVICO
            WHERE
                PRODUTO = ?
                AND STATUS = 'A'
            ORDER BY ID DESC
        """, (id_produto,))

        op_principal = cursor.fetchone()

        if not op_principal:
            return 0

        analisar_produto(
            op_principal[0]
        )

        # =========================================================
        # RESULTADO
        # =========================================================

        if total_itens:

            percentual = int(
                (total_consumidos / total_itens) * 100
            )

        else:
            percentual = 0

        print(
            f"Total de itens: {total_itens} | "
            f"Consumidos: {total_consumidos} | "
            f"Percentual: {percentual}%"
        )

        return percentual

    except Exception as e:
        trata_excecao(e)
        raise


if __name__ == "__main__":

    try:

        cursor = conecta.cursor()

        resultado = calcular_percentual_producao(
            37132
        )

        print(resultado)

    except Exception as e:
        trata_excecao(e)
        raise