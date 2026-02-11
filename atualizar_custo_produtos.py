from banco_dados.conexao import conecta
import pandas as pd
from pathlib import Path
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

arquivo = Path.home() / "Desktop" / "relatorio custo medio.xlsx"

# =========================
# 1. LÊ O EXCEL SEM CABEÇALHO
# =========================
df_raw = pd.read_excel(arquivo, header=None)

# =========================
# 2. ACHA A LINHA CORRETA DO CABEÇALHO
# (usa "Descrição do Produto", não "Produto")
# =========================
linha_inicio = df_raw[df_raw.apply(
    lambda row: row.astype(str).str.contains("Descrição do Produto").any(),
    axis=1
)].index[0]

# =========================
# 3. LÊ A TABELA A PARTIR DO CABEÇALHO REAL
# =========================
df = pd.read_excel(arquivo, skiprows=linha_inicio)

# =========================
# 4. REMOVE COLUNAS VAZIAS
# =========================
df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

# =========================
# 5. RENOMEIA AS COLUNAS (EXATAS DO EXCEL)
# =========================
df = df.rename(columns={
    "Produto": "codigo",
    "Descrição do Produto": "produto",
    "UN": "unidade",
    "Loc": "local",
    "Qtd.física": "quantidade",
    "Custo médio": "custo_medio"
})

# =========================
# 6. MANTÉM SÓ O QUE INTERESSA
# =========================
df = df[
    ["codigo", "produto", "unidade", "local", "quantidade", "custo_medio"]
]

# =========================
# 7. REMOVE LINHAS INVÁLIDAS
# =========================
df = df.dropna(subset=["codigo"])

# =========================
# 8. CONVERTE NÚMEROS COM VÍRGULA
# =========================
def to_float(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    return float(str(valor).replace(".", "").replace(",", "."))

df["quantidade"] = df["quantidade"].apply(to_float)
df["custo_medio"] = df["custo_medio"].apply(to_float)

# =========================
# 9. ATUALIZA O BANCO
# =========================
cursor = conecta.cursor()

for _, row in df.iterrows():
    cod_produto = int(row["codigo"])
    descricao = row["produto"]
    custo_medio = row["custo_medio"]

    cursor.execute(
        "SELECT id, conjunto FROM produto WHERE codigo = ?",
        (cod_produto,)
    )
    dados = cursor.fetchone()

    if not dados:
        continue

    id_prod, conjunto = dados

    if conjunto == 10:
        continue

    cursor.execute(
        "SELECT 1 FROM produto WHERE id = ? AND CUSTOUNITARIO = ?",
        (id_prod, custo_medio)
    )

    if not cursor.fetchone():
        cursor.execute(
            "UPDATE produto SET CUSTOUNITARIO = ? WHERE id = ?",
            (custo_medio, id_prod)
        )
        conecta.commit()

        print(f"Produto {cod_produto} - {descricao} atualizado")

print("✅ Finalizado com sucesso")

cursor.close()
conecta.close()
