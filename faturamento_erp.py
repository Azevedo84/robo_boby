from core.banco import conecta
import pandas as pd
from pathlib import Path

sql = """
SELECT
    M.DATA,
    C.RAZAO AS CLIENTE,
    SP.NUMERO AS PEDIDO,
    N.DESCRICAO AS NATUREZA,
    OC.NUMERO AS ORDEM_COMPRA,
    P.CODIGO,
    P.DESCRICAO,
    SP.QUANTIDADE,
    COALESCE(
    POC.UNITARIO,
    SP.UNITARIO,
                (
                    CASE
                        WHEN P.CONJUNTO = 10 THEN P.CUSTOESTRUTURA
                        ELSE P.CUSTOUNITARIO
                    END * 1.05 / 0.7663
                )
            ) AS UNITARIO, 
    SP.QUANTIDADE *
                COALESCE(
                    POC.UNITARIO,
                    SP.UNITARIO,
                    (
                        CASE
                            WHEN P.CONJUNTO = 10 THEN P.CUSTOESTRUTURA
                            ELSE P.CUSTOUNITARIO
                        END * 1.05 / 0.7663
                    )
                ) AS TOTAL

FROM SAIDAPROD SP

INNER JOIN MOVIMENTACAO M
    ON M.ID = SP.MOVIMENTACAO

INNER JOIN CLIENTES C
    ON C.ID = SP.CLIENTE

INNER JOIN PRODUTO P
    ON P.ID = SP.PRODUTO

INNER JOIN NATOP N
    ON N.ID = SP.NATUREZA

LEFT JOIN ORDEMCOMPRA OC
    ON OC.ID = SP.ORDEMCOMPRA

LEFT JOIN PRODUTOORDEMCOMPRA POC
    ON POC.MESTRE = OC.ID
   AND POC.PRODUTO = SP.PRODUTO

WHERE
    M.TIPO = 230
    AND M.DATA BETWEEN '2017-01-01' AND '2025-12-31'
    AND N.DESCRICAO LIKE '%VENDA%'

ORDER BY
    M.DATA,
    C.RAZAO,
    SP.NUMERO,
    P.CODIGO;
"""

df = pd.read_sql(sql, conecta)

desktop = Path.home() / "Desktop"
df.to_excel(desktop / "vendas_erp.xlsx", index=False)

print("Concluído.")