import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect
import openpyxl



class MaisUsados:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.consulta_mais_usados()

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

    def consulta_mais_usados(self):
        try:
            # Caminho para salvar na área de trabalho
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            arquivo_excel = os.path.join(desktop, "estoque_antigos.xlsx")

            # Cria um novo workbook e planilha
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Armazem 9"

            # Cabeçalho
            ws.append(["Código", "Descrição", "Local", "Tipo", "Nome Saldo", "Saldo"])

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, prod.localizacao, prod.tipomaterial, "
                           f"loc.nome, sald.saldo "
                           f"from produto as prod "
                           f"INNER JOIN SALDO_ESTOQUE as sald ON prod.id = sald.produto_id "
                           f"INNER JOIN LOCALESTOQUE loc ON sald.local_estoque = loc.id "
                           f"where prod.quantidade > 0;")
            dados_prod = cursor.fetchall()

            #f"INNER JOIN SALDO_ESTOQUE as sald ON prod.id = sald.produto_id "

            for i in dados_prod:
                id_prod, cod_prod, descr, local, tipo, nome_saldo, saldo = i

                cursor = conecta.cursor()
                cursor.execute(f"SELECT codigo from MOVIMENTACAO "
                               f"where codigo = '{cod_prod}' AND data >= '2018-01-01';")
                dados_mov = cursor.fetchall()

                if dados_mov:
                    qtde_mov = len(dados_mov)
                else:
                    qtde_mov = 0

                if not qtde_mov:
                    print(cod_prod, descr, local)

                    ws.append([cod_prod, descr, local, tipo, nome_saldo, saldo])

            # Salva o arquivo na área de trabalho
            wb.save(arquivo_excel)
            print(f"Arquivo salvo em: {arquivo_excel}")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = MaisUsados()