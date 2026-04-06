from core.banco import conecta
import os
import socket
import traceback
from datetime import datetime


def grava_erro_banco(nome_funcao, mensagem, nome_arquivo, num_linha):
    try:
        nome_computador = socket.gethostname()

        msg_final = f"{mensagem} {num_linha}"

        cursor = conecta.cursor()

        # ✅ SQL parametrizado (seguro)
        cursor.execute(
            "INSERT INTO ZZZ_ERROS (id, arquivo, funcao, mensagem, nome_pc) "
            "VALUES (GEN_ID(GEN_ZZZ_ERROS_ID,1), ?, ?, ?, ?)",
            (nome_arquivo, nome_funcao, msg_final, nome_computador)
        )

        conecta.commit()

        # 📄 salva log no desktop
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        caminho_log = os.path.join(desktop, f"{os.path.basename(nome_arquivo)}.log")

        with open(caminho_log, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 80)
            f.write(f"\nData: {datetime.now()}")
            f.write(f"\nArquivo: {nome_arquivo}")
            f.write(f"\nFunção: {nome_funcao}")
            f.write(f"\nComputador: {nome_computador}")
            f.write(f"\nLinha: {num_linha}")
            f.write(f"\nErro:\n{mensagem}")
            f.write("\n")

    except Exception as erro_gravacao:
        # ⚠️ fallback simples (não pode dar loop)
        print("⚠️ ERRO AO GRAVAR NO BANCO")
        print("Erro original:", mensagem)
        print("Erro ao gravar:", erro_gravacao)


def trata_excecao(e):
    try:
        tb = traceback.extract_tb(e.__traceback__)
        ultimo = tb[-1]

        nome_funcao = ultimo.name
        arquivo = ultimo.filename
        linha = ultimo.lineno
        mensagem = str(e)

        # 🔍 mostra no console
        traceback.print_exc()

        print(
            f'Houve um problema no arquivo: {arquivo} '
            f'na função: "{nome_funcao}"\n{mensagem} (linha {linha})'
        )

        grava_erro_banco(nome_funcao, mensagem, arquivo, linha)

    except Exception as erro_trat:
        # ⚠️ fallback final
        print("⚠️ ERRO AO TRATAR EXCEÇÃO")
        print("Erro original:", e)
        print("Erro no tratamento:", erro_trat)