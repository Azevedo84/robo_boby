import pandas as pd
import re
import pdfplumber
from pathlib import Path

desktop = Path.home() / "Desktop"

pdf = pdfplumber.open(desktop / "MAQ_FAT_302_L_0000.pdf")

dados = []

cliente = ""
pedido = ""
nota = ""
data = ""
oc = ""

for i, pagina in enumerate(pdf.pages):

    texto = pagina.extract_text().split("\n")

    for linha in texto:

        if "Nome do Cliente" in linha:
            m = re.search(r"Nome do Cliente\.+:\s*(.*?)\s+Codigo:", linha)
            if m:
                cliente = m.group(1).strip()

        elif "Codigo do Pedido" in linha:
            pedido = re.findall(r"\d+\.\d+", linha)[0]

        elif "Dt.Faturamento:" in linha:
            m = re.search(r"(\d{2}/\d{2}/\d{4})", linha)
            if m:
                data = m.group(1)

        elif "Serie/Num.Nota:" in linha:
            m = re.search(r"M\d/\s*(\d+)", linha)
            if m:
                nota = m.group(1)

        elif linha.startswith("OC:"):
            oc = linha.replace("OC:", "").strip()

        # linhas dos produtos
        elif re.match(r"^\s*\d+[,\.]?\d*\s+\w+", linha):

            if "*** TOTAL" in linha:
                continue

            if "*" in linha:
                continue

            partes = linha.split()

            try:
                qtd = partes[0]
                un = partes[1]

                if len(partes) >= 16:
                    referencia = partes[-8]
                    valor_unit = partes[-7]
                    valor_total = partes[-5]
                else:
                    referencia = partes[-6]
                    valor_unit = partes[-5]
                    valor_total = partes[-3]

                descricao = " ".join(partes[3:partes.index("D")])

                dados.append({
                    "Cliente": cliente,
                    "Data": data,
                    "Pedido": pedido,
                    "Nota": nota,
                    "OC": oc,
                    "Quantidade": qtd,
                    "Unidade": un,
                    "Descrição": descricao,
                    "Referência": referencia,
                    "Valor Unitário": valor_unit,
                    "Valor Total": valor_total
                })

            except:
                pass

df = pd.DataFrame(dados)

df["Quantidade"] = (
    pd.to_numeric(
        df["Quantidade"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    )
)

df["Valor Unitário"] = (
    pd.to_numeric(
        df["Valor Unitário"].astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    )
)

df["Valor Total"] = (
    pd.to_numeric(
        df["Valor Total"].astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    )
)

df.to_excel(desktop / "vendas_2025.xlsx", index=False)

print(df.head())
print(f"\n{len(df)} itens exportados.")