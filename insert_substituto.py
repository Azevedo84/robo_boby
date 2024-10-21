import fdb

conecta = fdb.connect(database=r'C:\HallSys\db\Horus\Suzuki\ESTOQUE.GDB',
                      host='PUBLICO',
                      port=3050,
                      user='sysdba',
                      password='masterkey',
                      charset='ANSI')

cod_pai = "21423"
cod_filho = "21163"
cod_subs = "16778"

num_op = ""

cursor = conecta.cursor()
cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {cod_pai};")
select_prod = cursor.fetchall()

idez, cod, id_estrut = select_prod[0]

cursor = conecta.cursor()
cursor.execute(f"SELECT estprod.id, prod.codigo "
               f"FROM estrutura_produto as estprod "
               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
               f"WHERE estprod.id_estrutura = {id_estrut} and prod.codigo = {cod_filho};")
dados_mat = cursor.fetchall()

if dados_mat:
    id_mat = dados_mat[0][0]
    print(dados_mat)

    if num_op:
        cursor = conecta.cursor()
        cursor.execute(f"SELECT * "
                       f"FROM SUBSTITUTO_MATERIAPRIMA "
                       f"WHERE ID_MAT = {id_mat} "
                       f"and COD_SUBS = {cod_subs} "
                       f"and NUM_OP = {num_op};")
        dados_subs = cursor.fetchall()

        if not dados_subs:
            cursor = conecta.cursor()
            cursor.execute(f"Insert into SUBSTITUTO_MATERIAPRIMA (ID, ID_MAT, COD_SUBS, NUM_OP) "
                           f"values (GEN_ID(GEN_SUBSTITUTO_MATERIAPRIMA_ID,1), {id_mat}, {cod_subs}, {num_op});")

            conecta.commit()

            print("SALVO COM SUCESSO!")

        else:
            print("ESTE ITEM JÁ FOI SALVO NA TABELA SUBSTITUTOS")
    else:
        cursor = conecta.cursor()
        cursor.execute(f"SELECT * "
                       f"FROM SUBSTITUTO_MATERIAPRIMA "
                       f"WHERE ID_MAT = {id_mat} "
                       f"and COD_SUBS = {cod_subs};")
        dados_subs = cursor.fetchall()

        if not dados_subs:
            cursor = conecta.cursor()
            cursor.execute(f"Insert into SUBSTITUTO_MATERIAPRIMA (ID, ID_MAT, COD_SUBS) "
                           f"values (GEN_ID(GEN_SUBSTITUTO_MATERIAPRIMA_ID,1), {id_mat}, {cod_subs});")

            conecta.commit()

            print("SALVO COM SUCESSO!")

        else:
            print("ESTE ITEM JÁ FOI SALVO NA TABELA SUBSTITUTOS")

