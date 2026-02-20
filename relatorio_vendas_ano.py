import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect

from openpyxl import Workbook


class ExecutaVendas:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

        self.resultado_ops_vinculos = {}
        self.resultado_sol_vinculos = {}
        self.resultado_req_vinculos = {}
        self.resultado_oc_vinculos = {}

        self.iniciar_tudo()

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

    def iniciar_tudo(self):
        try:
            tabela_nova = []

            clientes = [103, 2, 6, 5, 4, 3, 86]
            for i in clientes:
                cod_cliente = i

                cursor = conecta.cursor()
                cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prod.codigo, prod.descricao, "
                               f"COALESCE(prod.obs, '') as ref, prod.unidade, prodoc.quantidade, "
                               f"prodoc.UNITARIO, prod.CONJUNTO, estrut.QTDE_ITENS "
                               f"FROM PRODUTOORDEMCOMPRA as prodoc "
                               f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                               f"INNER JOIN RESUMO_ESTRUTURA as estrut ON prod.id = estrut.produto "
                               f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                               f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                               f"where oc.entradasaida = 'S'"
                               f"and oc.cliente = {cod_cliente} "
                               f"AND oc.data >= '2025-01-01' "
                               f"AND oc.data <= '2025-12-31';")
                dados_interno = cursor.fetchall()
                if dados_interno:
                    for ii in dados_interno:
                        emissao, num_ov, clie, cod, descr, ref, um, qtde, unit, conj, itens = ii

                        if conj == 10:
                            emi = f'{emissao.day}/{emissao.month}/{emissao.year}'

                            dados = (emi, num_ov, clie, cod, descr, ref, um, qtde, unit, conj, itens)
                            tabela_nova.append(dados)
            if tabela_nova:
                lista_ordenada = sorted(tabela_nova, key=lambda x: x[1])

                wb = Workbook()
                ws = wb.active
                ws.title = "Vendas"

                ws.append([
                    "Emissão", "Nº OV", "Cliente", "Código", "Descrição",
                    "Ref", "UM", "Qtd", "Unit", "Conjunto", "Itens"
                ])

                for linha in lista_ordenada:
                    ws.append(linha)

                desktop = os.path.join(os.path.expanduser("~"), "Desktop")

                if not os.path.exists(desktop):
                    desktop = os.path.join(os.path.expanduser("~"), "Área de Trabalho")

                caminho_arquivo = os.path.join(desktop, "relatorio_vendas.xlsx")

                wb.save(caminho_arquivo)

                print(f"Excel criado em: {caminho_arquivo}")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = ExecutaVendas()
