from core.banco import conecta_engenharia
from core.erros import trata_excecao


def limpar_propriedades_idw_orfas():
    try:
        cursor = conecta_engenharia.cursor()

        print("🔎 Buscando registros órfãos...")

        # 1. CONSULTAR registros que serão apagados
        cursor.execute("""
            SELECT p.ID_ARQUIVO, a.TIPO_ARQUIVO
            FROM PROPRIEDADES_IDW p
            INNER JOIN ARQUIVOS a ON a.ID = p.ID_ARQUIVO
        """)

        registros = cursor.fetchall()

        if registros:
            for i in registros:
                if i[1] != "IDW":
                    print(i)

    except Exception as e:
        conecta_engenharia.rollback()
        trata_excecao(e)


if __name__ == "__main__":
    limpar_propriedades_idw_orfas()