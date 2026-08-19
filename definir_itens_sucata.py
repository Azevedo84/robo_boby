from core.banco import conecta
import pandas as pd
import pdfplumber
import re

arquivo = r"C:\Users\Anderson\Desktop\pacifil.xlsx"
pdf = r"C:\Users\Anderson\Desktop\pendencias\MAQ_06.2026_LOCAL 1 - PROPRIO.pdf"

valores_pdf = {}
codigo_pendente = None

with pdfplumber.open(pdf) as arq:

    for pagina in arq.pages:
        texto = pagina.extract_text()

        if not texto:
            continue

        for linha in texto.split("\n"):

            if not linha.startswith("|"):
                continue

            partes = [p.strip() for p in linha.split("|")]

            # Se existe um código pendente, tenta pegar o valor nesta linha
            if codigo_pendente is not None:
                try:
                    valor = float(partes[5].replace(".", "").replace(",", "."))
                    valores_pdf[codigo_pendente] = valor
                    codigo_pendente = None
                    continue
                except:
                    pass

            if len(partes) < 7:
                continue

            dados = partes[2].split()

            if not dados:
                continue

            if not dados[0].isdigit():
                continue

            codigo = int(dados[0])

            # Tenta pegar o valor na própria linha
            try:
                valor = float(partes[5].replace(".", "").replace(",", "."))
                valores_pdf[codigo] = valor
                codigo_pendente = None
            except:
                # Se não conseguiu, guarda o código para procurar o valor na próxima linha
                codigo_pendente = codigo

df = pd.read_excel(arquivo)

resultado = []

cursor = conecta.cursor()

for codigo in df.iloc[:, 0].dropna().astype(int):

    cursor.execute("""
        SELECT
            prod.codigo,
            prod.descricao,
            COALESCE(prod.obs, '') AS obs,
            prod.unidade,
            tip.tipomaterial,
            loc.nome AS local_estoque,
            COALESCE(se.saldo, 0) AS saldo
        FROM produto prod
        LEFT JOIN tipomaterial tip
            ON prod.tipomaterial = tip.id
        LEFT JOIN saldo_estoque se
            ON se.produto_id = prod.id
        LEFT JOIN localestoque loc
            ON loc.id = se.local_estoque
        WHERE prod.codigo = ?
        ORDER BY loc.nome;
    """, (codigo,))
    detalhes_pai = cursor.fetchall()

    if detalhes_pai:
        cod, descri, ref, um, tipo, local_est, saldo = detalhes_pai[0]

        valor_unitario = valores_pdf.get(int(cod))

        saldo_total = 0
        saldo_almox = 0

        for linha in detalhes_pai:
            local = linha[5] if linha[5] else "SEM ESTOQUE"
            saldo = float(linha[6]) if linha[6] is not None else 0

            saldo_total += saldo


            if local.upper() == "ALMOX":
                saldo_almox += saldo

        if saldo_total > 0:
            resultado.append({
                "CODIGO": cod,
                "DESCRICAO": descri,
                "REFERENCIA": ref,
                "UNIDADE": um,
                "SALDO ALMOX": saldo_almox,
                "SALDO TOTAL": saldo_total,
                "VALOR_UNITARIO": valor_unitario
            })
        if saldo_total < 0:
            print("BBBBBBBBBBB", cod, descri, ref, um, local_est, saldo)


    else:
        print(f"==== Código não existe!")

df_saida = pd.DataFrame(resultado)
arquivo_saida = r"C:\Users\Anderson\Desktop\Saldo_Almox.xlsx"
df_saida.to_excel(arquivo_saida, index=False)
print(f"Arquivo gerado: {arquivo_saida}")