import sys
from banco_dados.conexao import conecta
from banco_dados.conexao_nuvem import conectar_banco_nuvem
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect

class AtualizaCustoDaNuvem:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.inicia_processo()

    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

        except Exception as e:
            nome_funcao_trat = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            tb = traceback.extract_tb(exc_traceback)
            num_linha_erro = tb[-1][1]
            print(f'Houve um problema no arquivo: {self.nome_arquivo} na função: "{nome_funcao_trat}"\n'
                  f'{e} {num_linha_erro}')
            grava_erro_banco(nome_funcao_trat, e, self.nome_arquivo, num_linha_erro)

    def inicia_processo(self):
        conecta_nuvem = conectar_banco_nuvem()
        try:
            cursor_local = conecta.cursor()
            cursor_nuvem = conecta_nuvem.cursor()

            # 1. Pega dados da nuvem
            cursor_nuvem.execute("SELECT CODIGO_PROD, CUSTO_MEDIO FROM CUSTO_MEDIO_PRODUTO;")
            dados_nuvem = cursor_nuvem.fetchall()

            if not dados_nuvem:
                print("Nenhum dado na nuvem")
            else:
                print(f"{len(dados_nuvem)} produtos na nuvem")

            # 2. Pega produtos locais (somente os que importam)
            cursor_local.execute("SELECT id, codigo, conjunto, CUSTOUNITARIO FROM produto;")
            produtos_locais = cursor_local.fetchall()

            # Cria dicionário: codigo -> (id, conjunto, custo_atual)
            produtos_dict = {p[1]: (p[0], p[2], p[3]) for p in produtos_locais}

            # 3. Prepara lista de updates
            updates = []
            for cod_prod, custo_medio in dados_nuvem:
                produto_local = produtos_dict.get(cod_prod)
                if produto_local:
                    id_prod, conjunto, custo_atual = produto_local
                    if conjunto != 10 and custo_medio != custo_atual:
                        updates.append((custo_medio, id_prod))

            # 4. Executa updates em lote
            if updates:
                cursor_local.executemany(
                    "UPDATE produto SET CUSTOUNITARIO = ? WHERE id = ?",
                    updates
                )
                conecta.commit()
                print(f"{len(updates)} produtos atualizados")
            else:
                print("Nenhum produto precisou ser atualizado")

            cursor_nuvem.execute("SELECT COUNT(*) FROM CUSTO_MEDIO_PRODUTO;")
            total = cursor_nuvem.fetchone()[0]
            print(f"Vai apagar {total} registros da nuvem")

            # Depois de atualizar os produtos locais
            cursor_nuvem.execute("DELETE FROM CUSTO_MEDIO_PRODUTO;")
            conecta_nuvem.commit()
            print("Tabela temporária na nuvem limpa")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

        finally:
            if 'conexao' in locals():
                conecta_nuvem.close()


chama_classe = AtualizaCustoDaNuvem()