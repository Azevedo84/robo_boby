import pandas as pd
from banco_dados.conexao import conecta

# Caminho do Excel na área de trabalho
caminho_excel = r"C:\Users\Anderson\Desktop\op_almox.xlsx"

# Lê o Excel
df = pd.read_excel(caminho_excel)

# Remove possíveis linhas vazias
df = df.dropna(subset=["OP"])

cursor = conecta.cursor()

for num_op_excel in df["OP"]:
    try:
        num_op_excel = int(num_op_excel)

        cursor = conecta.cursor()
        cursor.execute(f"SELECT op.id, op.numero, op.codigo, op.id_estrutura "
                       f"FROM ordemservico as op "
                       f"where op.numero = {num_op_excel};")
        ops_abertas = cursor.fetchall()

        if ops_abertas:
            id_op, num_op, cod, id_estrut = ops_abertas[0]

            cursor = conecta.cursor()
            cursor.execute(f"UPDATE ordemservico SET etapa = 'ALMOX' "
                           f"WHERE id = {id_op};")

    except Exception as e:
        print(f"Erro na OP {num_op_excel}: {e}")

# Salva alterações
conecta.commit()

cursor.close()
conecta.close()

print("Processo finalizado.")