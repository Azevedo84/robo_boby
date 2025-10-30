import os
import pandas as pd
from banco_dados.conexao import conecta

# Caminho do arquivo na área de trabalho
usuario = os.getlogin()
caminho = fr"C:\Users\{usuario}\Desktop\ops vinculos pi.xlsx"

# Lê a planilha
df = pd.read_excel(caminho, sheet_name="Planilha1")

# Garante que as colunas estejam corretas
df.columns = ["SOL", "PRODUTO", "PI", "PRODUTO_PI"]

# Cria variáveis dinamicamente
variaveis = {}
for _, linha in df.iterrows():
    num_op = str(linha["SOL"])
    cod_produto_oc = str(linha["PRODUTO"])
    num_pi = str(linha["PI"])
    cod_produto_pi = str(linha["PRODUTO_PI"])
    print(num_op, cod_produto_oc, num_pi, cod_produto_pi)

    cursor = conecta.cursor()
    cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {cod_produto_oc};")
    select_prod_op = cursor.fetchall()
    id_produto_oc, cod, id_versao = select_prod_op[0]

    cursor = conecta.cursor()
    cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {cod_produto_pi};")
    select_pi = cursor.fetchall()
    id_produto_pi, cod_pi, id_versao_pi = select_pi[0]

    print(id_produto_pi, num_pi)

    cursor = conecta.cursor()
    cursor.execute(f"Insert into VINCULO_PRODUTO_PI "
                   f"(id, id_pedidointerno, id_produto_pi, tipo, numero, id_produto) "
                   f"values (GEN_ID(GEN_VINCULO_PRODUTO_PI_ID,1), {num_pi}, {id_produto_pi}, 'SOL', "
                   f"{num_op}, {id_produto_oc});")

    conecta.commit()
