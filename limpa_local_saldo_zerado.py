import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect


class LimpaLocalSaldoZerado:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

        self.manipula_comeco()

    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

            # 'Log' em arquivo local apenas se houver erro
            with open(self.arquivo_log, "a", encoding="utf-8") as f:
                f.write(f"Erro na função {nome_funcao} do arquivo {arquivo}: {mensagem} (linha {num_linha_erro})\n")

        except Exception as e:
            nome_funcao_trat = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            tb = traceback.extract_tb(exc_traceback)
            num_linha_erro = tb[-1][1]
            print(f'Houve um problema no arquivo: {self.nome_arquivo} na função: "{nome_funcao_trat}"\n'
                  f'{e} {num_linha_erro}')
            grava_erro_banco(nome_funcao_trat, e, self.nome_arquivo, num_linha_erro)

            with open(self.arquivo_log, "a", encoding="utf-8") as f:
                f.write(
                    f"Erro na função {nome_funcao_trat} do arquivo {self.nome_arquivo}: {e} (linha {num_linha_erro})\n")

    def manipula_comeco(self):
        try:
            cursor = conecta.cursor()
            cursor.execute("""
                SELECT DISTINCT prod.id, prod.codigo, prod.localizacao
                FROM movimentacao AS mov
                INNER JOIN produto AS prod ON mov.produto = prod.id
                WHERE prod.quantidade = 0
                AND prod.localizacao IS NOT NULL 
                  AND prod.localizacao NOT LIKE 'A-%';
            """)
            dados_mov = cursor.fetchall()

            if dados_mov:
                for i in dados_mov:
                    id_prod, cod_prod, local = i
                    print(i)

                    cursor = conecta.cursor()
                    cursor.execute("UPDATE produto SET LOCALIZACAO = NULL WHERE id = ?", (id_prod,))

                    conecta.commit()

                    print(f"Cadastro do produto {cod_prod} atualizado com Sucesso {local}!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

chama_classe = LimpaLocalSaldoZerado()
chama_classe.manipula_comeco()
