import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect
from datetime import datetime


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
            previsao = datetime.now()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, op.numero, op.codigo, op.id_estrutura, prod.descricao, "
                           f"COALESCE(prod.obs, ''), "
                           f"prod.unidade, COALESCE(prod.tipomaterial, ''), op.quantidade, "
                           f"COALESCE(ser.descricao, '') "
                           f"FROM ordemservico as op "
                           f"INNER JOIN produto as prod ON op.produto = prod.id "
                           f"LEFT JOIN SERVICO_INTERNO as ser ON ser.id = prod.id_servico_interno "
                           f"where op.status = 'A' order by op.numero;")
            ops_abertas = cursor.fetchall()

            ops_por_numero = {}

            if ops_abertas:
                for i in ops_abertas:
                    id_produto, num_op, cod, id_estrutura, descr, ref, um, tipo, qtde, servico_in = i

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, id_pedidointerno, ID_PRODUTO_PI, tipo, numero, id_produto "
                                   f"FROM VINCULO_PRODUTO_PI "
                                   f"where numero = '{num_op}' and tipo = 'OP' "
                                   f"and id_produto = {id_produto};")
                    consulta_vinculos = cursor.fetchall()
                    if consulta_vinculos:
                        for ii in consulta_vinculos:
                            id_vinculo, id_pedido, id_produto_pedido, tipo, numero, id_produto = ii

                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT ped.emissao, prod.codigo, "
                                           f"prod.descricao, "
                                           f"COALESCE(prod.obs, '') as obs, "
                                           f"prod.unidade, prodint.data_previsao "
                                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                                           f"where prodint.status = 'A' "
                                           f"and prodint.id_produto = {id_produto_pedido} "
                                           f"and prodint.id_pedidointerno = {id_pedido} "
                                           f"ORDER BY prodint.data_previsao ASC;")
                            dados_interno = cursor.fetchall()
                            if dados_interno:
                                for iii in dados_interno:
                                    emissao, codi, descri, refi, umi, previsao = iii

                                    dados = (num_op, id_pedido, cod, descr, ref, um, qtde, previsao, servico_in)

                                    # Se ainda não existe essa OP no dicionário
                                    if num_op not in ops_por_numero:
                                        ops_por_numero[num_op] = dados
                                    else:
                                        # Compara a data de previsão (posição 5 da tupla)
                                        previsao_atual = ops_por_numero[num_op][7]
                                        if previsao < previsao_atual:
                                            ops_por_numero[num_op] = dados

                    else:
                        dados = (num_op, "sem vinculo", cod, descr, ref, um, qtde, previsao, servico_in)

                        # Se ainda não existe essa OP no dicionário
                        if num_op not in ops_por_numero:
                            ops_por_numero[num_op] = dados
                        else:
                            # Compara a data de previsão (posição 5 da tupla)
                            previsao_atual = ops_por_numero[num_op][7]
                            if previsao < previsao_atual:
                                ops_por_numero[num_op] = dados

                lista = list(ops_por_numero.values())

                if lista:
                    self.manipula_dados(lista)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados(self, lista):
        try:
            op_ab_editado = []

            for i in lista:
                num_op, pi, cod, descr, ref, um, qtde, previsao, servico = i

                cursor = conecta.cursor()
                cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                               f"prod.descricao, "
                               f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                               f"ordser.quantidade, ordser.id_estrutura, ordser.etapa "
                               f"from ordemservico as ordser "
                               f"INNER JOIN produto prod ON ordser.produto = prod.id "
                               f"where ordser.status = 'A' and ordser.numero = '{num_op}';")
                op_abertas = cursor.fetchall()
                if op_abertas:
                    emissao, previsao_aaa, op, cod, descr, ref, um, qtde, id_estrut, etapa = op_abertas[0]

                    if id_estrut:
                        total_estrut = 0
                        total_consumo = 0

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT estprod.id, "
                                       f"((SELECT quantidade FROM ordemservico where numero = {op}) * "
                                       f"(estprod.quantidade)) AS Qtde "
                                       f"FROM estrutura_produto as estprod "
                                       f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                       f"where estprod.id_estrutura = {id_estrut};")
                        itens_estrutura = cursor.fetchall()

                        for dads in itens_estrutura:
                            ides, quantidade = dads
                            total_estrut += 1

                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT max(prodser.ID_ESTRUT_PROD), "
                                           f"sum(prodser.QTDE_ESTRUT_PROD) as total "
                                           f"FROM estrutura_produto as estprod "
                                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                           f"INNER JOIN produtoos as prodser ON estprod.id = prodser.ID_ESTRUT_PROD "
                                           f"where prodser.numero = {op} and estprod.id = {ides} "
                                           f"group by prodser.ID_ESTRUT_PROD;")
                            itens_consumo = cursor.fetchall()
                            for duds in itens_consumo:
                                id_mats, qtde_mats = duds
                                if ides == id_mats and quantidade == qtde_mats:
                                    total_consumo += 1

                        msg = f"{total_estrut}/{total_consumo}"

                        if total_estrut == total_consumo:
                            dados = (num_op, pi, cod, descr, ref, um, qtde, previsao, servico, msg, etapa)
                            op_ab_editado.append(dados)
                        else:
                            if not servico:
                                dados = (num_op, pi, cod, descr, ref, um, qtde, previsao, "SEM SERVIÇO DEFINIDO", msg, etapa)
                                op_ab_editado.append(dados)
                            elif servico != "MONTAGEM" and servico != "SOLDA" and servico != "ELETRICO":
                                dados = (num_op, pi, cod, descr, ref, um, qtde, previsao, "CORTE", msg, etapa)
                                op_ab_editado.append(dados)
                            else:
                                dados = (num_op, pi, cod, descr, ref, um, qtde, previsao, "AGUARDANDO MATERIAL", msg, etapa)
                                op_ab_editado.append(dados)

            if op_ab_editado:
                self.gerar_excel(op_ab_editado)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def gerar_excel(self, lista):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment
            from openpyxl.utils import get_column_letter
            import os
            from datetime import date, datetime

            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            arquivo = os.path.join(desktop, "ops_abertas.xlsx")

            wb = Workbook()
            ws = wb.active
            ws.title = "OPs Abertas"

            cabecalhos = [
                "Número OP", "Nº PI", "Código Produto", "Descrição", "Observação",
                "Unidade", "Quantidade", "Previsão", "Serviço Interno", "Consumo", "Etapa"
            ]
            ws.append(cabecalhos)

            alinhamento_central = Alignment(horizontal="center", vertical="center")

            for linha in lista:
                num_op, pi, cod, descr, ref, um, qtde, previsao, servico, consumo, etapa = linha

                # Código do produto como número (remove letras se existir)
                try:
                    cod = int(cod)
                except:
                    pass

                ws.append([
                    num_op,
                    pi,
                    cod,
                    descr,
                    ref,
                    um,
                    qtde,
                    previsao,
                    servico,
                    consumo,
                    etapa
                ])

            # Formatação das células
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = alinhamento_central

                    # Data em formato brasileiro
                    if isinstance(cell.value, (date, datetime)):
                        cell.number_format = "DD/MM/YYYY"

            # Centraliza cabeçalhos
            for cell in ws[1]:
                cell.alignment = alinhamento_central

            # Ajusta largura das colunas automaticamente
            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)

                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))

                ws.column_dimensions[col_letter].width = max_length + 2

            wb.save(arquivo)

            print("Excel Gerado!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

chama_classe = ClassificarOps()