from banco_dados.conexao import conecta

cursor = conecta.cursor()
cursor.execute(f"SELECT op.id, op.numero, op.codigo, op.id_estrutura "
               f"FROM ordemservico as op "
               f"where op.status = 'A';")
ops_abertas = cursor.fetchall()

if ops_abertas:
    for i in ops_abertas:
        id_op, num_op, cod, id_estrut = i

        cursor = conecta.cursor()
        cursor.execute(f"UPDATE ordemservico SET etapa = 'ABERTA' "
                       f"WHERE id = {id_op};")

        conecta.commit()

        print(f"OP {num_op} ALTERADA")