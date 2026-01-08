import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect


class ClassificarOps:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.manipula_comeco()

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

    def manipula_comeco(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, op.numero, op.codigo, op.id_estrutura, prod.descricao, "
                           f"COALESCE(prod.obs, ''), "
                           f"prod.unidade, COALESCE(prod.tipomaterial, ''), op.quantidade, "
                           f"COALESCE(prod.id_servico_interno, '') "
                           f"FROM ordemservico as op "
                           f"INNER JOIN produto as prod ON op.produto = prod.id "
                           f"where op.status = 'A' order by op.numero;")
            ops_abertas = cursor.fetchall()

            if ops_abertas:
                for i in ops_abertas:
                    id_produto, num_op, cod, id_estrutura, descr, ref, um, tipo, qtde, id_servico = i

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, id_pedidointerno, tipo, numero, id_produto "
                                   f"FROM VINCULO_PRODUTO_PI "
                                   f"where numero = '{num_op}' and tipo = 'OP' "
                                   f"and id_produto = {id_produto};")
                    consulta_vinculos = cursor.fetchall()
                    print(num_op, cod, consulta_vinculos)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

chama_classe = ClassificarOps()