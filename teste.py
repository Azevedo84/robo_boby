from banco_dados.conexao import conecta

cursor = conecta.cursor()
cursor.execute("""
    SELECT produto_id, local_estoque, COUNT(*)
    FROM SALDO_ESTOQUE
    GROUP BY produto_id, local_estoque
    HAVING COUNT(*) > 1;
""")
duplicatas = cursor.fetchall()

if duplicatas:
    print("Registros duplicados encontrados:")
    for produto_id, local_estoque, quantidade in duplicatas:
        print(f"Produto ID: {produto_id}, Local Estoque: {local_estoque}, Quantidade: {quantidade}")
else:
    print("Nenhum registro duplicado encontrado.")