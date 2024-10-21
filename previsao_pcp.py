import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from comandos.conversores import valores_para_float
from datetime import timedelta, date, datetime
import inspect
import os
import math
import traceback


from comandos.excel import edita_alinhamento, edita_bordas, edita_preenchimento
from comandos.excel import edita_fonte, criar_workbook, letra_coluna
from pathlib import Path


class PcpPrevisao:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        data_inicio = date.today() + timedelta(1)
        dia_da_semana = data_inicio.weekday()
        if dia_da_semana == 5:
            data_ini = date.today() + timedelta(3)
        elif dia_da_semana == 6:
            data_ini = date.today() + timedelta(2)
        else:
            data_ini = date.today() + timedelta(1)

        self.date_Inicio = data_ini
        self.line_Func = "5"
        self.line_HorasDia = "8"
        self.line_SemanCompra = "1"
        
        self.manipula_dados_pi()

    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')
            print(f'Houve um problema no arquivo:\n\n{arquivo}\n\n'
                                 f'Comunique o desenvolvedor sobre o problema descrito abaixo:\n\n'
                                 f'{nome_funcao}: {mensagem}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

        except Exception as e:
            nome_funcao_trat = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            tb = traceback.extract_tb(exc_traceback)
            num_linha_erro = tb[-1][1]
            print(f'Houve um problema no arquivo: {self.nome_arquivo} na função: "{nome_funcao_trat}"\n'
                  f'{e} {num_linha_erro}')
            grava_erro_banco(nome_funcao_trat, e, self.nome_arquivo, num_linha_erro)

    def retorna_calculo_meses(self, dados_tabela):
        try:
            tab_prev = []
            for datas in dados_tabela:
                cod_prev = datas[1]
                dt_previsao = datas[6]
                data_obj = datetime.strptime(dt_previsao, "%d/%m/%Y").date()
                cc = (cod_prev, data_obj)
                tab_prev.append(cc)

            tab_ordenada = sorted(tab_prev, key=lambda x: x[1])

            data_mais_alta = max(tab_ordenada, key=lambda x: x[1])

            tab_meses = []
            for item in tab_ordenada:
                cod_isso = item[0]
                data = item[1]
                diferenca = (data - data_mais_alta[1]).days

                if diferenca < 0:
                    diferenca1 = diferenca * -1
                else:
                    diferenca1 = 0

                meses = diferenca1 / 30 if diferenca1 != 0 else 0

                meses_arredondados = math.ceil(meses)

                dd = (cod_isso, meses_arredondados)
                tab_meses.append(dd)

            return tab_meses

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_folgas(self, data_inicio, data_fim):
        try:
            contagem_sabados = 0
            contagem_domingos = 0

            data_atual = data_inicio
            while data_atual <= data_fim:
                if data_atual.weekday() == 5:
                    contagem_sabados += 1
                elif data_atual.weekday() == 6:
                    contagem_domingos += 1

                data_atual += timedelta(days=1)

            return contagem_sabados, contagem_domingos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def retorna_data_entrega(self, id_pais):
        try:
            tempos_de_entrega = []
            fornecedor = ''
            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, prodoc.produto, prodoc.quantidade, mov.data, forn.razao "
                           f"FROM produtoordemcompra as prodoc "
                           f"INNER JOIN entradaprod as ent ON prodoc.mestre = ent.ordemcompra "
                           f"INNER JOIN movimentacao as mov ON ent.movimentacao = mov.id "
                           f"INNER JOIN fornecedores as forn ON ent.fornecedor = forn.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"WHERE prodoc.produto = '{id_pais}' and oc.entradasaida = 'E';")
            extrair_prod = cursor.fetchall()

            if extrair_prod:
                for registro in extrair_prod:
                    data_emissao = registro[0]
                    data_entrega = registro[4]
                    fornecedor = registro[5]

                    tempo_entrega_dias = (data_entrega - data_emissao).days
                    tempos_de_entrega.append(tempo_entrega_dias)
            if tempos_de_entrega:
                media_entrega = sum(tempos_de_entrega) / len(tempos_de_entrega)
            else:
                media_entrega = 0

            entrega = int(media_entrega)

            return entrega, fornecedor

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_producao(self, cod_prod):
        try:
            dados_ops = []
            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where ordser.status = 'A' and prod.codigo = {cod_prod} "
                           f"order by ordser.numero;")
            op_abertas = cursor.fetchall()
            if op_abertas:
                for dados_op in op_abertas:
                    emissao, previsao, op, cod, descr, ref, um, qtde = dados_op

                    dados = (op, qtde)
                    dados_ops.append(dados)

            return dados_ops

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_pi(self):
        try:
            dados_p_tabela = []
            tabela_final = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prodint.id_pedidointerno, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodint.qtde, prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    num_pi, cod, descr, ref, um, qtde, entrega = i

                    dados = (num_pi, cod, descr, ref, um, qtde, entrega, "", "")

                    dados_p_tabela.append(dados)

                tab_ordenada = sorted(dados_p_tabela, key=lambda x: x[6])

                for ii in tab_ordenada:
                    num_pis, cods, des, refs, um, qti, pr, niv, calc = ii

                    prev = pr.strftime('%d/%m/%Y')

                    coco = (num_pis, cods, des, refs, um, qti, prev, niv, calc)
                    tabela_final.append(coco)

            return tabela_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_1_dados_previsao(self):
        try:
            tudo_tudo = []

            dados_tabela = self.manipula_dados_pi()

            if dados_tabela:
                saldos = []

                tab_meses = self.retorna_calculo_meses(dados_tabela)

                for i in dados_tabela:
                    num_pi, codigo, descr, ref, um, qtde, previsao, nivi, entreg = i
                    agrega_pt = 0
                    for cod_isso, meses_isso in tab_meses:
                        if cod_isso == codigo:
                            if meses_isso > 0:
                                agrega_pt = meses_isso * 2

                    cod_origem = ""
                    descr_origem = ""

                    pcte_p_estrutura = [agrega_pt, 1, codigo, qtde, cod_origem, descr_origem, saldos, num_pi]

                    estrutura = self.calculo_3_verifica_estrutura(pcte_p_estrutura)

                    if estrutura:
                        for ii in estrutura:
                            tudo_tudo.append(ii)

            if tudo_tudo:
                tudo_tudo_ordenada = sorted(tudo_tudo, key=lambda x: (-x[0], -x[1]))

                self.excel(tudo_tudo_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_3_verifica_estrutura(self, dados_total, ordens_verificadas=None):
        try:
            testador = "10499"

            pontos, nivel, codigo, qtdei, cod_or, descr_or, lista_saldos, num_pi = dados_total

            if ordens_verificadas is None:
                ordens_verificadas = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prod.quantidade, tip.tipomaterial, prod.id_versao "
                           f"FROM produto as prod "
                           f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                           f"WHERE prod.codigo = '{codigo}';")
            detalhes_pai = cursor.fetchall()

            if not detalhes_pai:
                raise Exception(f"Produto com código {codigo} não encontrado.")

            id_pai, cod_pai, descr_pai, ref_pai, um_pai, saldo, tipo, id_estrut = detalhes_pai[0]

            chave_verificacao = (id_pai, cod_or)
            if chave_verificacao in ordens_verificadas:
                return []

            else:
                ordens_verificadas.append(chave_verificacao)

                consumo_anterior = 0

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = '{cod_or}';")
                select_prod = cursor.fetchall()
                if select_prod:
                    id_estrutura = select_prod[0][2]

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT estprod.id, estprod.id_estrutura "
                                   f"from estrutura_produto as estprod "
                                   f"where estprod.id_estrutura = {id_estrutura} "
                                   f"and estprod.id_prod_filho = {id_pai};")
                    tabela_estrutura = cursor.fetchall()

                    if tabela_estrutura:
                        id_prod_filho = tabela_estrutura[0][0]

                        cursor.execute(f"SELECT SUM(produtos.qtde_estrut_prod) "
                                       f"FROM produtoos as produtos "
                                       f"INNER JOIN ordemservico as ord ON produtos.mestre = ord.id "
                                       f"INNER JOIN produto prod ON ord.produto = prod.id "
                                       f"WHERE produtos.ID_ESTRUT_PROD = {id_prod_filho} "
                                       f"AND prod.codigo = '{cod_or}';")
                        consumo_anterior = cursor.fetchone()[0] or 0  # Consumo anterior ou 0 se não houver

                        if cod_pai == testador:
                            print(num_pi, codigo, descr_pai, ref_pai, um_pai, "\n",
                                  " - qtdei:", qtdei, " - consumo_op:", consumo_anterior, cod_or)

                qtdei_float = valores_para_float(qtdei)
                saldo_float = valores_para_float(saldo)
                consumo_anterior_float = valores_para_float(consumo_anterior)

                prod_saldo_encontrado = False
                for cod_sal_e, saldo_e in lista_saldos:
                    if cod_sal_e == cod_pai:
                        prod_saldo_encontrado = True
                        break

                if prod_saldo_encontrado:
                    for i_ee, (cod_ee, saldo_ee) in enumerate(lista_saldos):
                        if cod_ee == cod_pai:
                            novo_saldo_lim = (saldo_ee + consumo_anterior_float) - qtdei_float
                            novo_saldo = round(novo_saldo_lim, 2)

                            lista_saldos[i_ee] = (cod_pai, novo_saldo)
                            break
                else:
                    novo_saldo_lim = (saldo_float + consumo_anterior_float) - qtdei_float
                    novo_saldo = round(novo_saldo_lim, 2)

                    lista_saldos.append((cod_pai, novo_saldo))

                filhos = []

                if cod_pai == testador:
                    print(" - novo_saldo:", novo_saldo, "\n")

                if novo_saldo >= 0 or consumo_anterior_float == qtdei_float:
                    return []  # Não precisa produzir mais, pois o saldo é suficiente
                else:
                    saldo_positivo = novo_saldo * -1
                    if saldo_positivo < qtdei_float:
                        nova_qtde = saldo_positivo
                    else:
                        nova_qtde = qtdei_float
                    if cod_pai == testador:
                        print("       nova_qtde:", nova_qtde, "\n")
                    dadoss = (pontos, nivel, cod_pai, descr_pai, ref_pai, um_pai, nova_qtde, cod_or, descr_or, num_pi)
                    filhos.append(dadoss)

                    nivel_plus = nivel + 1
                    pts_plus = pontos + 1

                    if id_estrut:
                        cursor.execute(
                            f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                            f"(estprod.quantidade * {nova_qtde}) as qtde "
                            f"FROM estrutura_produto as estprod "
                            f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                            f"WHERE estprod.id_estrutura = {id_estrut};")
                        dados_estrutura = cursor.fetchall()

                        if dados_estrutura:
                            for prod in dados_estrutura:
                                cod_f, descr_f, ref_f, um_f, qtde_f = prod

                                pcte_filho = [pts_plus, nivel_plus, cod_f, qtde_f, cod_pai, descr_pai, lista_saldos,
                                              num_pi]
                                filhos.extend(self.calculo_3_verifica_estrutura(pcte_filho, ordens_verificadas))

                return filhos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel(self, dados_extraidos):
        try:
            desktop = Path.home() / "Desktop"
            desk_str = str(desktop)
            nome_req = '\Teste PCP.xlsx'
            caminho = (desk_str + nome_req)

            if not dados_extraidos:
                print(f'A Tabela "Lista de Materiais" está vazia!')
            else:
                self.excel_total(caminho, dados_extraidos)
                self.excel_teste(caminho, dados_extraidos)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel_total(self, caminho, nova_tabela):
        try:
            workbook = criar_workbook()
            sheet = workbook.active
            sheet.title = "Lista Completa"

            headers = ["Pontos", "Nivel", "Código", "Descrição", "Referência", "UM", "Qtde", "Cod. Or",
                       "Descrição Origem", "Nº PI"]
            sheet.append(headers)

            header_row = sheet[1]
            for cell in header_row:
                edita_fonte(cell, negrito=True)
                edita_preenchimento(cell)
                edita_alinhamento(cell)

            for d_ex in nova_tabela:
                pts_ex, nivel_ex, cod_ex, de_ex, ref_ex, um_ex, qtde_ex, cod_pai_ex, descr_pai_ex, num_pi_ex = d_ex

                nivius = int(nivel_ex)
                codigu = int(cod_ex)
                pi_int = int(num_pi_ex)

                if cod_pai_ex:
                    cod_pai_ex_int = int(cod_pai_ex)
                else:
                    cod_pai_ex_int = 0

                if qtde_ex == "":
                    qtde_e = 0.00
                else:
                    qtde_e = float(qtde_ex)

                sheet.append([pts_ex, nivius, codigu, de_ex, ref_ex, um_ex, qtde_e, cod_pai_ex_int,
                              descr_pai_ex, pi_int])

            for row in sheet.iter_rows(min_row=1,
                                       max_row=sheet.max_row,
                                       min_col=1,
                                       max_col=sheet.max_column):
                for cell in row:
                    edita_bordas(cell)
                    edita_alinhamento(cell)

            for column in sheet.columns:
                max_length = 0
                column_letter = letra_coluna(column[0].column)
                for cell in column:
                    if isinstance(cell.value, (int, float)):
                        cell_value_str = "{:.2f}".format(cell.value)
                    else:
                        cell_value_str = str(cell.value)
                    if len(cell_value_str) > max_length:
                        max_length = len(cell_value_str)

                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width

            for row in sheet.iter_rows(min_row=2,
                                       max_row=sheet.max_row,
                                       min_col=7,
                                       max_col=9):
                for cell in row:
                    cell.number_format = '0.00'

            for linha in sheet.iter_rows(min_row=2,
                                         max_row=sheet.max_row,
                                         min_col=9,
                                         max_col=9):
                for cell in linha:
                    cell.number_format = '0'

            workbook.save(caminho)

            print("Excel Salvo!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel_teste(self, caminho, nova_tabela):
        try:
            soma_quantidades = {}

            for i in nova_tabela:
                pontos, nivel, cod_pai, descr_pai, ref_pai, um_pai, nova_qtde, cod_or, descr_or, num_pi = i

                nova_qtde = float(nova_qtde)

                if cod_pai in soma_quantidades:
                    soma_quantidades[cod_pai] += nova_qtde
                else:
                    soma_quantidades[cod_pai] = nova_qtde

            nova_tabela_somada = []
            for cod_pai, total_qtde in soma_quantidades.items():
                nova_tabela_somada.append((cod_pai, total_qtde))

            if nova_tabela_somada:
                for ii in nova_tabela_somada:
                    pass

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = PcpPrevisao()
chama_classe.calculo_1_dados_previsao()
