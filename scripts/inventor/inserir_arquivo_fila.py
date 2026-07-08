from core.banco import conecta_engenharia

descricao = "CRIAR NOVA VERSÃO ESTRUTURA"

cursor = conecta_engenharia.cursor()
cursor.execute("""
                INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO, ORIGEM)
                VALUES (?, ?)
            """, (26927, "ALTERADOS"))

conecta_engenharia.commit()

print("SALVO COM SUCESSO!")