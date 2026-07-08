from core.banco import conecta_engenharia

descricao = "ARQUIVO SEM PROPRIEDADE"

cursor = conecta_engenharia.cursor()
cursor.execute(f"Insert into TIPO_DIVERGENCIA (DESCRICAO) "
               f"values ('{descricao}');")

conecta_engenharia.commit()

print("SALVO COM SUCESSO!")