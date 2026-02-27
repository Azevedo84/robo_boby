from banco_dados.conexao import conecta
import os
import socket
import traceback
from datetime import datetime


def grava_erro_banco(nome_funcao, e, nome_arquivo, num_linha):
    try:
        nome_computador = socket.gethostname()

        msg_editada = str(e).replace("'", "*")
        msg_editada1 = msg_editada.replace('"', '*')

        msg_final = f"{msg_editada1} {num_linha}"

        cursor = conecta.cursor()
        cursor.execute(f"Insert into ZZZ_ERROS (id, arquivo, funcao, mensagem, nome_pc) "
                       f"values (GEN_ID(GEN_ZZZ_ERROS_ID,1), '{nome_arquivo}', '{nome_funcao}', '{msg_final}', "
                       f"'{nome_computador}');")
        conecta.commit()

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")

        caminho_log = os.path.join(desktop, f"{nome_arquivo}.log")

        with open(caminho_log, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 80)
            f.write(f"\nData: {datetime.now()}")
            f.write(f"\nArquivo: {nome_arquivo}")
            f.write(f"\nFunção: {nome_funcao}")
            f.write(f"\nComputador: {nome_computador}")
            f.write(f"\nErro:\n{traceback.format_exc()}")
            f.write("\n")

    except Exception as e:
        print(e)
