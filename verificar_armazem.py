import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect
import openpyxl
from decimal import Decimal



class VerificaArmazem:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

        self.consulta_armazem()

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

    def consulta_armazem(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ''), "
                           f"prod.unidade, saldo.saldo, prod.localizacao "
                           f"from SALDO_ESTOQUE as saldo "
                           f"INNER JOIN produto as prod ON saldo.produto_id = prod.id "
                           f"where saldo.LOCAL_ESTOQUE = 9;")
            dados_armazem = cursor.fetchall()

            if dados_armazem:
                # Caminho para salvar na área de trabalho
                desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                arquivo_excel = os.path.join(desktop, "estoque_armazem.xlsx")

                # Cria um novo workbook e planilha
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Armazem 9"

                # Cabeçalho
                ws.append(["Código", "Descrição", "Referência", "Unidade", "Saldo", "Local", "Qtde Mov"])

                # Adiciona os dados
                for cod, descr, ref, um, saldo, local in dados_armazem:
                    if saldo < 0:
                        # Converte Decimal para float
                        if isinstance(saldo, Decimal):
                            saldo = float(saldo)

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT codigo from MOVIMENTACAO "
                                       f"where codigo = '{cod}' AND data >= '2024-01-01';")
                        dados_mov = cursor.fetchall()

                        if dados_mov:
                            qtdess = len(dados_mov)
                        else:
                            qtdess = 0

                        ws.append([cod, descr, ref, um, saldo, local, qtdess])

                # Salva o arquivo na área de trabalho
                wb.save(arquivo_excel)
                print(f"Arquivo salvo em: {arquivo_excel}")
            else:
                print("Nenhum dado encontrado.")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = VerificaArmazem()