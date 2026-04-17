from core.banco import conecta
import os
import socket
import traceback
from datetime import datetime


def grava_erro_banco(nome_funcao, mensagem, nome_arquivo, num_linha):
    try:
        nome_computador = socket.gethostname()

        # 🔹 LIMITA TAMANHO (evita erro no banco)
        mensagem_limpa = str(mensagem)
        if len(mensagem_limpa) > 500:
            mensagem_limpa = mensagem_limpa[:500]

        arquivo_limpo = str(nome_arquivo)[-80:]
        funcao_limpa = str(nome_funcao)[-80:]

        cursor = conecta.cursor()

        cursor.execute(
            "INSERT INTO ZZZ_ERROS (id, arquivo, funcao, mensagem, nome_pc) "
            "VALUES (GEN_ID(GEN_ZZZ_ERROS_ID,1), ?, ?, ?, ?)",
            (arquivo_limpo, funcao_limpa, mensagem_limpa, nome_computador)
        )

        conecta.commit()

        # ===============================
        # 📄 LOG NA ÁREA DE TRABALHO
        # ===============================
        desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
        nome_log = f"{os.path.basename(nome_arquivo)}.log"
        caminho_log = os.path.join(desktop, nome_log)

        with open(caminho_log, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 80)
            f.write(f"\nData: {datetime.now()}")
            f.write(f"\nArquivo: {nome_arquivo}")
            f.write(f"\nFunção: {nome_funcao}")
            f.write(f"\nComputador: {nome_computador}")
            f.write(f"\nLinha: {num_linha}")
            f.write(f"\nErro completo:\n{mensagem}")
            f.write("\n")

    except Exception as erro_gravacao:
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

        print(
            f'Houve um problema no arquivo: {arquivo} '
            f'na função: "{nome_funcao}"\n{mensagem} (linha {linha})'
        )

        trace_completo = traceback.format_exc()

        grava_erro_banco(
            nome_funcao,
            f"{mensagem}\n{trace_completo}",
            arquivo,
            linha
        )

    except Exception as erro_trat:
        print("⚠️ ERRO AO TRATAR EXCEÇÃO")
        print("Erro original:", e)
        print("Erro no tratamento:", erro_trat)