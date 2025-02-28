from banco_dados.conexao import conecta

cursor = conecta.cursor()
cursor.execute(f"SELECT ID, codigo, DESCRICAO, DESCRICAOCOMPLEMENTAR, OBS FROM produto;")
select_prod = cursor.fetchall()

for i in select_prod:
    id_prod, cod, descr, compl, ref = i

    if cod.endswith(" "):
        cod_sem_espaco = cod.rstrip()

        cursor = conecta.cursor()
        cursor.execute("UPDATE produto SET CODIGO = ? WHERE id = ?", (cod_sem_espaco, id_prod,))

        conecta.commit()

        print(f"Cadastro do produto {cod} - {descr} atualizado com Sucesso!")

    elif cod.startswith(" "):
        cod_sem_espaco = cod.lstrip()

        cursor = conecta.cursor()
        cursor.execute("UPDATE produto SET CODIGO = ? WHERE id = ?", (cod_sem_espaco, id_prod,))

        conecta.commit()

        print(f"Cadastro do produto {cod} - {descr} atualizado com Sucesso!")

    elif descr.endswith(" "):
        texto_sem_espaco = descr.rstrip()

        cursor = conecta.cursor()
        cursor.execute("UPDATE produto SET DESCRICAO = ? WHERE id = ?", (texto_sem_espaco, id_prod,))

        conecta.commit()

        print(f"Cadastro do produto {cod} - {descr} atualizado com Sucesso!")

    elif descr.startswith(" "):
        descr_sem_espaco = cod.lstrip()

        cursor = conecta.cursor()
        cursor.execute("UPDATE produto SET DESCRICAO = ? WHERE id = ?", (descr_sem_espaco, id_prod,))

        conecta.commit()

        print(f"Cadastro do produto {cod} - {descr} atualizado com Sucesso!")

    elif compl:
        if compl.endswith(" "):
            compl_sem_espaco = compl.rstrip()

            cursor = conecta.cursor()
            cursor.execute("UPDATE produto SET DESCRICAOCOMPLEMENTAR = ? WHERE id = ?", (compl_sem_espaco, id_prod,))

            conecta.commit()

            print(f"Cadastro do produto {cod} - {compl_sem_espaco} atualizado com Sucesso!")

        elif compl.startswith(" "):
            compl_sem_espaco = compl.lstrip()

            cursor = conecta.cursor()
            cursor.execute("UPDATE produto SET DESCRICAOCOMPLEMENTAR = ? WHERE id = ?", (compl_sem_espaco, id_prod,))

            conecta.commit()

            print(f"Cadastro do produto {cod} - {compl_sem_espaco} atualizado com Sucesso!")

    elif ref:
        if ref.endswith(" "):
            ref_sem_espaco = ref.rstrip()

            cursor = conecta.cursor()
            cursor.execute("UPDATE produto SET OBS = ? WHERE id = ?", (ref_sem_espaco, id_prod,))

            conecta.commit()

            print(f"Cadastro do produto {cod} - {ref_sem_espaco} atualizado com Sucesso!")

        elif ref.startswith(" "):
            ref_sem_espaco = ref.lstrip()

            cursor = conecta.cursor()
            cursor.execute("UPDATE produto SET OBS = ? WHERE id = ?", (ref_sem_espaco, id_prod,))

            conecta.commit()

            print(f"Cadastro do produto {cod} - {ref_sem_espaco} atualizado com Sucesso!")