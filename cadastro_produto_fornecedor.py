from banco_dados.conexao import conecta


cod_fornecedor_siger = "10695"
cod_meu_produto = "56380"
cod_produto_f = "051770"
um_f = "MT"

cursor = conecta.cursor()
cursor.execute(f"SELECT id, registro, cnpj FROM FORNECEDORES where registro = '{cod_fornecedor_siger}'")
dados_fornecedor = cursor.fetchall()

cursor = conecta.cursor()
cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {cod_meu_produto};")
dados_siger_prod = cursor.fetchall()

if dados_fornecedor and dados_siger_prod:
    id_fornecedor = dados_fornecedor[0][0]

    id_siger_prod = dados_siger_prod[0][0]

    cursor = conecta.cursor()
    cursor.execute(f"SELECT ID_PRODUTO, COD_PRODUTO_F, ID_FORNECEDOR, UM_F "
                   f"FROM PRODUTO_FORNECEDOR "
                   f"where ID_PRODUTO = {id_siger_prod} "
                   f"and ID_FORNECEDOR = {id_fornecedor} "
                   f"and COD_PRODUTO_F = {cod_produto_f};")
    dados_prod_fornecedor = cursor.fetchall()

    if not dados_prod_fornecedor:
        cursor = conecta.cursor()
        cursor.execute(f"Insert into PRODUTO_FORNECEDOR (ID_PRODUTO, COD_PRODUTO_F, ID_FORNECEDOR, UM_F) "
                       f"values ({id_siger_prod}, {cod_produto_f}, {id_fornecedor}, '{um_f}');")

        conecta.commit()

        print("LANÇADO")
    else:
        print("PRODUTO JÁ FOI LANÇADO!")