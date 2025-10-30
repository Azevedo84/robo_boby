import os
import pandas as pd
from banco_dados.conexao import conecta

# Caminho do arquivo na área de trabalho
usuario = os.getlogin()
caminho = fr"C:\Users\{usuario}\Desktop\estoque_mais_usados.xlsx"

# Lê a planilha
df = pd.read_excel(caminho, sheet_name="CX2")

# Garante que as colunas estejam corretas
df.columns = ["Codigo", "Descricao"]

# Cria variáveis dinamicamente
variaveis = {}
for _, linha in df.iterrows():
    codigo = str(linha["Codigo"])
    descricao = str(linha["Descricao"])
    print(codigo)
    print(descricao)

    local = "G-2-ESQ-CX9#"

    campos_atualizados = [f"LOCALIZACAO = '{local}'"]

    campos_update = ", ".join(campos_atualizados)
    cursor = conecta.cursor()
    cursor.execute(f"UPDATE produto SET {campos_update} "
                   f"WHERE codigo = '{codigo}';")

    conecta.commit()
