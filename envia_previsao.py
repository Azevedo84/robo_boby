import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from comandos.conversores import valores_para_float
from comandos.excel import carregar_workbook, edita_alinhamento, edita_bordas, edita_preenchimento
from comandos.excel import edita_fonte, criar_workbook, letra_coluna, ajusta_larg_coluna
from datetime import timedelta, date, datetime
import inspect
import os
import math
from pathlib import Path
import re
import traceback


class TelaPcpPrevisaoV2:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

        self.line_SemanCompra = "1"

        self.dados_pi = []
        self.dados_previsao = []
        self.data_inicio = None
        self.funcionario = "5"
        self.horas_dia = "8"

        self.definir_data_inicio()
        self.manipula_dados_pi()
        self.calculo_1_chamar_funcao()
        self.excel()

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

    def retorna_calculo_meses(self, dados_tabela):
        try:
            print("retorna_calculo_meses")

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
            return None

    def retorna_oc_abertas(self, cod_prod):
        try:
            qtdes_oc = 0

            if cod_prod:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT COALESCE(prodreq.mestre, ''), req.dataemissao, prodreq.quantidade "
                               f"FROM produtoordemsolicitacao as prodreq "
                               f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                               f"INNER JOIN ordemsolicitacao as req ON prodreq.mestre = req.idsolicitacao "
                               f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                               f"WHERE prodreq.status = 'A' "
                               f"and prod.codigo = {cod_prod} "
                               f"AND preq.id_prod_sol IS NULL "
                               f"ORDER BY prodreq.mestre;")
                dados_sol = cursor.fetchall()

                if dados_sol:
                    for i_sol in dados_sol:
                        num_sol, emissao_sol, qtde_sol = i_sol
                        qtdes_oc += float(qtde_sol)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT sol.idsolicitacao, prodreq.quantidade, req.data, prodreq.numero, "
                               f"prodreq.destino, prodreq.id_prod_sol "
                               f"FROM produtoordemrequisicao as prodreq "
                               f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                               f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                               f"INNER JOIN produtoordemsolicitacao as prodsol ON prodreq.id_prod_sol = prodsol.id "
                               f"INNER JOIN ordemsolicitacao as sol ON prodsol.mestre = sol.idsolicitacao "
                               f"where prodreq.status = 'A' "
                               f"and prod.codigo = {cod_prod};")
                dados_req = cursor.fetchall()

                if dados_req:
                    for i_req in dados_req:
                        num_sol_req, qtde_req, emissao_req, num_req, destino, id_prod_sol = i_req
                        qtdes_oc += float(qtde_req)

                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT sol.idsolicitacao, prodreq.numero, oc.data, oc.numero, forn.razao, "
                    f"prodoc.quantidade, prodoc.produzido, prodoc.dataentrega "
                    f"FROM ordemcompra as oc "
                    f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                    f"INNER JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                    f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                    f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                    f"INNER JOIN produtoordemsolicitacao as prodsol ON prodreq.id_prod_sol = prodsol.id "
                    f"INNER JOIN ordemsolicitacao as sol ON prodsol.mestre = sol.idsolicitacao "
                    f"where oc.entradasaida = 'E' "
                    f"AND oc.STATUS = 'A' "
                    f"AND prodoc.produzido < prodoc.quantidade "
                    f"and prod.codigo = {cod_prod}"
                    f"ORDER BY oc.numero;")
                dados_oc = cursor.fetchall()

                if dados_oc:
                    for i_oc in dados_oc:
                        num_sol_oc, id_req_oc, emissao_oc, num_oc, forncec_oc, qtde_oc, prod_oc, entrega_oc = i_oc
                        qtdes_oc += float(qtde_oc)

            return qtdes_oc

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def retorna_ops_saldo_ops_abertas(self, cod_pai, cod_filho):
        try:
            qtde_ops = 0

            if cod_pai:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {cod_pai};")
                select_prod = cursor.fetchall()
                id_pai, cod, id_versao = select_prod[0]

                if id_versao:
                    cursor = conecta.cursor()
                    cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.id, "
                                   f"prod.descricao, "
                                   f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                                   f"ordser.quantidade "
                                   f"from ordemservico as ordser "
                                   f"INNER JOIN produto prod ON ordser.produto = prod.id "
                                   f"where ordser.status = 'A' AND prod.codigo = {cod_pai} "
                                   f"order by ordser.numero;")
                    op_abertas = cursor.fetchall()

                    if op_abertas:
                        for i in op_abertas:
                            num_op = i[2]

                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                                           f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                                           f"(estprod.quantidade)) AS Qtde, "
                                           f"COALESCE(prod.localizacao, ''), prod.quantidade "
                                           f"FROM estrutura_produto as estprod "
                                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                           f"where estprod.id_estrutura = {id_versao} and prod.codigo = {cod_filho} "
                                           f"ORDER BY prod.descricao;")
                            select_estrut = cursor.fetchall()
                            if select_estrut:
                                for dados_estrut in select_estrut:
                                    id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_e, local_e, saldo_e = dados_estrut

                                    cursor = conecta.cursor()
                                    cursor.execute(f"SELECT max(estprod.id), max(prod.codigo), max(prod.descricao), "
                                                   f"sum(prodser.QTDE_ESTRUT_PROD)as total "
                                                   f"FROM estrutura_produto as estprod "
                                                   f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                                   f"INNER JOIN produtoos as prodser ON "
                                                   f"estprod.id = prodser.id_estrut_prod "
                                                   f"where prodser.numero = {num_op} and estprod.id = {id_mat_e} "
                                                   f"group by prodser.id_estrut_prod;")
                                    select_os_resumo = cursor.fetchall()

                                    if select_os_resumo:
                                        for dados_res in select_os_resumo:
                                            id_mat_sum, cod_sum, descr_sum, qtde_sum = dados_res

                                            qtde_sum_float = valores_para_float(qtde_sum)

                                            if cod_filho == "21313":
                                                print("dados op:", dados_res)

                                            qtde_ops += qtde_sum_float

            return qtde_ops

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def manipula_ordens_compra(self, dados_total):
        try:
            lista_ocs, cod_pai, qtdei_float = dados_total

            qtde_ocs = self.retorna_oc_abertas(cod_pai)

            prod_ocs_encontrado = False
            for cod_oc_p, saldo_p in lista_ocs:
                if cod_oc_p == cod_pai:
                    prod_ocs_encontrado = True
                    break

            if not prod_ocs_encontrado:
                if qtde_ocs > 0:
                    lanca_saldo = (cod_pai, qtde_ocs)
                    lista_ocs.append(lanca_saldo)

                    qtde_float_com_oc = qtdei_float - qtde_ocs
                else:
                    qtde_float_com_oc = qtdei_float
            else:
                qtde_float_com_oc = qtdei_float

            return lista_ocs, qtde_float_com_oc

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def manipula_ordens_producao(self, dados_total):
        try:
            lista_ops, cod_origem, cod_pai, qtdei_float = dados_total

            qtde_ops = self.retorna_ops_saldo_ops_abertas(cod_origem, cod_pai)

            prod_ops_encontrado = False
            for cod_op_p, saldo_p in lista_ops:
                if cod_op_p == cod_pai:
                    prod_ops_encontrado = True
                    break

            if not prod_ops_encontrado:
                if qtde_ops > 0:
                    lanca_saldo = (cod_pai, qtde_ops)
                    lista_ops.append(lanca_saldo)

                    qtde_float_op = qtdei_float - qtde_ops
                else:
                    qtde_float_op = qtdei_float
            else:
                if qtde_ops > 0:
                    qtde_float_op = qtdei_float - qtde_ops
                else:
                    qtde_float_op = qtdei_float

            return lista_ops, qtde_float_op

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

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
            return None

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
            return None

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
            return None

    def manipula_dados_pi(self):
        try:
            print("manipula_dados_pi")

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

            if tabela_final:
                self.dados_pi = tabela_final
                print("lança tabela", tabela_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_data_inicio(self):
        try:
            data_inicio = date.today() + timedelta(1)
            dia_da_semana = data_inicio.weekday()
            if dia_da_semana == 5:
                data_ini = date.today() + timedelta(3)
            elif dia_da_semana == 6:
                data_ini = date.today() + timedelta(2)
            else:
                data_ini = date.today() + timedelta(1)

            self.data_inicio = data_ini

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_1_chamar_funcao(self):
        try:
            print("calculo_1_chamar_funcao")

            dados = self.dados_pi

            if not dados:
                print(f'A tabela "Pedidos Internos Pendentes" está vazia!')
            else:
                self.calculo_2_dados_previsao()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_2_dados_previsao(self):
        try:
            print("calculo_2_dados_previsao")

            tudo_tudo = []
            dados_niveis = []

            dados_tabela = self.dados_pi

            if dados_tabela:
                saldos = []
                ops = []
                ocs = []

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
                    pacote = ""

                    pcte_p_estrutura = [agrega_pt, 1, codigo, qtde, cod_origem, descr_origem, saldos, ops, ocs,
                                        codigo, num_pi, pacote]

                    estrutura = self.calculo_3_verifica_estrutura(pcte_p_estrutura)

                    if estrutura:
                        ultimo_nivel = max(item[1] for item in estrutura)
                        cece = (codigo, ultimo_nivel, previsao)
                        dados_niveis.append(cece)

                        for ii in estrutura:
                            tudo_tudo.append(ii)

            if tudo_tudo:
                self.calculo_4_final_lanca_tabelas(tudo_tudo)
            else:
                print(f'Este Plano de produção está concluído!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_3_verifica_estrutura(self, dados_total):
        try:
            print("calculo_3_verifica_estrutura")

            pontos, nivel, codigos, qtdei, cod_or, descr_or, lista_saldos, lista_ops, lista_ocs, cod_fat, \
            num_pi, pacote = dados_total

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prod.quantidade, tip.tipomaterial, prod.id_versao "
                           f"FROM produto as prod "
                           f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                           f"where prod.codigo = {codigos};")
            detalhes_pai = cursor.fetchall()
            id_pai, cod_pai, descr_pai, ref_pai, um_pai, saldo, tipo, id_estrut = detalhes_pai[0]

            qtdei_float = valores_para_float(qtdei)
            saldo_float = valores_para_float(saldo)

            pcte_ops = [lista_ops, cod_or, cod_pai, qtdei_float]
            lista_ops_mex, qtde_float_com_op = self.manipula_ordens_producao(pcte_ops)

            pcte_ocs = [lista_ocs, cod_pai, qtde_float_com_op]
            lista_ocs_mex, qtde_flt_c_oc = self.manipula_ordens_compra(pcte_ocs)

            prod_saldo_encontrado = False
            for cod_sal_e, saldo_e in lista_saldos:
                if cod_sal_e == cod_pai:
                    prod_saldo_encontrado = True
                    break

            if prod_saldo_encontrado:
                for i_ee, (cod_ee, saldo_ee) in enumerate(lista_saldos):
                    if cod_ee == cod_pai:
                        novo_saldo = saldo_ee - qtde_flt_c_oc
                        lista_saldos[i_ee] = (cod_pai, novo_saldo)
                        break

            else:
                novo_saldo = saldo_float - qtde_flt_c_oc
                lanca_saldo = (cod_pai, novo_saldo)
                lista_saldos.append(lanca_saldo)

            if cod_pai == "21578":
                print("num_pi:", num_pi, "cod_or:", cod_or, "descr_or:", descr_or, "qtdei:", qtdei, novo_saldo, qtde_flt_c_oc)

            if novo_saldo < 0 < qtde_flt_c_oc:
                coco = novo_saldo + qtde_flt_c_oc
                if coco > 0:
                    nova_qtde = novo_saldo * -1
                else:
                    nova_qtde = qtde_flt_c_oc

                if cod_pai == "21578":
                    print("entrei", nova_qtde)

                dadoss = (pontos, nivel, cod_pai, descr_pai, ref_pai, um_pai, nova_qtde, cod_or, descr_or,
                          cod_fat, num_pi, pacote)
                nov_msg = f"{cod_pai}({qtde_flt_c_oc}), "
                pacote += nov_msg

                filhos = [dadoss]

                nivel_plus = nivel + 1
                pts_plus = pontos + 1

                if id_estrut:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                                   f"(estprod.quantidade * {nova_qtde}) as qtde "
                                   f"FROM estrutura_produto as estprod "
                                   f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                   f"where estprod.id_estrutura = {id_estrut};")
                    dados_estrutura = cursor.fetchall()

                    if dados_estrutura:
                        for prod in dados_estrutura:
                            cod_f, descr_f, ref_f, um_f, qtde_f = prod

                            pcte_filho = [pts_plus, nivel_plus, cod_f, qtde_f, cod_pai, descr_pai, lista_saldos,
                                          lista_ops_mex, lista_ocs_mex, cod_fat, num_pi, pacote]
                            filhos.extend(self.calculo_3_verifica_estrutura(pcte_filho))

            else:
                filhos = []

            return filhos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def calculo_4_final_lanca_tabelas(self, estrutura_final):
        try:
            print("calculo_4_final_lanca_tabelas")

            lista_de_listas_ordenada = sorted(estrutura_final, key=lambda x: (-x[0], -x[1]))

            lista_com_tempos = []

            for est in lista_de_listas_ordenada:
                pts, nivel, cod, descr, ref, um, qtde, cod_pai, descr_pai, cod_fat, num_pi, pacote = est

                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.id, prod.codigo, COALESCE(prod.tempo, 0) as temps, tip.tipomaterial "
                               f"FROM produto as prod "
                               f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                               f"where prod.codigo = {cod};")
                detalhes_tempo = cursor.fetchall()
                id_pai, cod_bc, tempo, tipo = detalhes_tempo[0]

                entrega, forn = self.retorna_data_entrega(id_pai)

                dedos = (pts, nivel, cod, tempo, forn, entrega, tipo, qtde, cod_pai, descr_pai, cod_fat, num_pi, pacote)
                lista_com_tempos.append(dedos)

            self.calculo_5_recebe_dados(lista_com_tempos)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_5_recebe_dados(self, estrutura_ord):
        try:
            print("calculo_5_recebe_dados")

            temp_acum = 0.00

            data_ini = self.data_inicio
            data_inicio = datetime.combine(data_ini, datetime.min.time())

            qtde_func = 10
            horas_dia = 8
            semana_compra = 2

            horas_por_dia = float(qtde_func) * float(horas_dia)

            tabela_pra_tabela = []
            tabela_pra_datas = []

            nivel_anterior = 0

            ultima_data = None

            for i in estrutura_ord:
                pontos_i, nivel_i, cod_i, temp_i, forn_i, entreg_i, tipo_i, qtde_i, cod_or_i, descr_or_i, \
                cod_fat, num_pi, pacote = i

                if not tipo_i:
                    tipo = ""
                else:
                    tipo = tipo_i

                if temp_i > 0:
                    tempo_i2 = float(temp_i)

                    total_horas = int(tempo_i2)
                    tempo_decimal = (tempo_i2 - total_horas)
                    minutos_convert = (float(tempo_decimal) * 0.5) / 0.3
                    hor_min_convertido = float(total_horas) + minutos_convert
                    soma_horas_com_qtde = float(qtde_i) * hor_min_convertido

                    temp_pc = "%.2f" % hor_min_convertido

                    temp_acum += soma_horas_com_qtde

                    dias_de_producao = temp_acum / horas_por_dia

                    data_producao = data_inicio + timedelta(dias_de_producao)

                else:
                    if pontos_i < nivel_anterior or nivel_anterior == 0:
                        soma_horas_com_qtde = (int(semana_compra) * 7) * horas_por_dia

                        temp_acum += soma_horas_com_qtde
                        dias_de_producao = temp_acum / horas_por_dia
                        data_producao = data_inicio + timedelta(dias_de_producao)

                        nivel_anterior = pontos_i
                    else:
                        data_producao = ultima_data

                    temp_pc = ""

                dia_da_semana = data_producao.weekday()
                if dia_da_semana == 5:
                    temp_acum += (horas_por_dia * 2)
                    dias_de_prod = temp_acum / horas_por_dia
                    data_final_prod = data_inicio + timedelta(dias_de_prod)

                elif dia_da_semana == 6:
                    temp_acum += horas_por_dia
                    dias_de_prod = temp_acum / horas_por_dia
                    data_final_prod = data_inicio + timedelta(dias_de_prod)

                else:
                    dias_de_prod = temp_acum / horas_por_dia
                    data_final_prod = data_inicio + timedelta(dias_de_prod)

                dt_prod = data_final_prod.strftime('%d/%m/%Y')
                ultima_data = data_final_prod

                qtde_c = "%.2f" % qtde_i
                acum_c = "%.2f" % temp_acum

                dedos1 = (pontos_i, nivel_i, cod_i, qtde_c, dt_prod, temp_pc, acum_c, tipo, forn_i,
                          cod_or_i, descr_or_i, cod_fat, num_pi, pacote)

                tabela_pra_tabela.append(dedos1)

                dedos3 = (cod_i, data_final_prod, dt_prod, nivel_i)
                tabela_pra_datas.append(dedos3)

            ultimo_data = max(item[1].date() for item in tabela_pra_datas)

            sabados, domingos = self.definir_folgas(data_inicio.date(), ultimo_data)
            dias_folga = sabados + domingos

            print("Folga", dias_folga)
            print("total itens", (str(len(tabela_pra_tabela))))
            print("data final", ultimo_data)

            self.calculo_6_manipula_final_tab(tabela_pra_tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_6_manipula_final_tab(self, lista_final):
        try:
            print("calculo_6_manipula_final_tab")

            tabela_nova = []
            tabela_p_pi = []

            if lista_final:
                for i in lista_final:
                    pontos, nivel, codigo, qtde, entrega, horas, acumulado, tipo, forn, cod_pai, descr_pai, \
                    cod_fat, num_pi, pacote = i

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                                   f"prod.unidade, prod.conjunto "
                                   f"FROM produto as prod "
                                   f"where prod.codigo = {codigo};")
                    detalhes_tempo = cursor.fetchall()
                    codis, descr, ref, um, conjunto = detalhes_tempo[0]

                    if conjunto == 10:
                        conj = "PRODUTO ACABADO"
                    else:
                        conj = "MATERIA-PRIMA"

                    dados = (pontos, codis, descr, ref, um, qtde, entrega, conj, cod_pai, pacote, num_pi)
                    tabela_nova.append(dados)

                    dados1 = (nivel, codis, entrega, cod_fat, num_pi)
                    tabela_p_pi.append(dados1)

            if tabela_nova:
                self.calculo_7_manipula_previsao_pi(tabela_p_pi)

                print("lança tabela previsão", tabela_nova)
                self.dados_previsao = tabela_nova

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_7_manipula_previsao_pi(self, dados_previsao):
        try:
            print("calculo_7_manipula_previsao_pi")

            nova_lista = []

            print("extrair tabela PI")

            dados_pi = self.dados_pi

            contagem_cod_fat_num_pi_fat = {}

            for nivel_pr, cod_pr, entrega, cod_fat, num_pi_fat in dados_previsao:
                chave = (cod_fat, num_pi_fat)
                if chave in contagem_cod_fat_num_pi_fat:
                    contagem_cod_fat_num_pi_fat[chave] += 1
                else:
                    contagem_cod_fat_num_pi_fat[chave] = 1

            for i in dados_pi:
                num_pi, cod_pi, descr_pi, ref_pi, um_pi, qtde_pi, limite_pi, previsao_pi, niveis_pi = i

                for chave, contagem in contagem_cod_fat_num_pi_fat.items():
                    cod_fat, num_pi_fat = chave
                    entrega = None

                    if num_pi == num_pi_fat and cod_pi == cod_fat:
                        for dados_previsao_item in dados_previsao:
                            if cod_pi == dados_previsao_item[1] and num_pi == dados_previsao_item[4]:
                                entrega = dados_previsao_item[2]
                                break

                        dados = (num_pi, cod_pi, descr_pi, ref_pi, um_pi, qtde_pi, limite_pi, entrega, contagem)
                        nova_lista.append(dados)

            if nova_lista:
                self.calculo_8_manipula_previsao_pi2(dados_pi, nova_lista)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_8_manipula_previsao_pi2(self, dados_pi, nova_lista):
        try:
            print("calculo_8_manipula_previsao_pi2")

            lista_nova_nova = []

            dados_pi_set = {(x[0], x[1]) for x in dados_pi}
            nova_lista_set = {(x[0], x[1]) for x in nova_lista}

            elementos_faltantes = dados_pi_set - nova_lista_set

            for elemento in elementos_faltantes:
                num_pi_enc, cod_enc = elemento

                for i in dados_pi:
                    num_pi, cod, descr_pi, ref_pi, um_pi, qtde_pi, limite_pi, entrega, contagem = i

                    if num_pi == num_pi_enc and cod == cod_enc:
                        dados = (num_pi, cod, descr_pi, ref_pi, um_pi, qtde_pi, limite_pi, "CONCLUÍDO", "0")
                        lista_nova_nova.append(dados)

            for ii in nova_lista:
                lista_nova_nova.append(ii)

            if lista_nova_nova:
                def converter_data(data_str):
                    if data_str.lower() == 'concluído':
                        return datetime.max
                    else:
                        return datetime.strptime(data_str, '%d/%m/%Y')

                if lista_nova_nova:
                    tab_ordenada = sorted(lista_nova_nova, key=lambda x: converter_data(x[7]))

                    print("lança tabela PI", tab_ordenada)
                    self.dados_pi = tab_ordenada

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel(self):
        try:
            lista_acumulada = []

            nova_lista_compras = []
            nova_lista_acabado = []

            desktop = Path.home() / "Desktop"
            desk_str = str(desktop)
            nome_req = '\Material Pendente.xlsx'
            caminho = (desk_str + nome_req)

            dados_extraidos = self.dados_previsao
            if not dados_extraidos:
                print(f'A Tabela "Lista de Materiais" está vazia!')
            else:
                self.excel_total(caminho, dados_extraidos)

                self.excel_pedido_interno(caminho)

                for dados_ex in dados_extraidos:
                    nivel, cod, descr, ref, um, qtde, entrega, conj, cod_pai, pacote, num_pi = dados_ex
                    qtde_float = valores_para_float(qtde)

                    prod_acum_encont = False
                    for nivel_e, cod_e, qtde_e, grade_ops_e in lista_acumulada:
                        if cod_e == cod:
                            prod_acum_encont = True
                            break

                    if prod_acum_encont:
                        for i_ee, (nivel_ee, cod_ee, qtde_ee, grade_ops_ee) in enumerate(lista_acumulada):
                            qtde_ee_float = valores_para_float(qtde_ee)

                            if cod_ee == cod:
                                nova_qtde = qtde_ee_float + qtde_float
                                nova_grade_ops = f"{grade_ops_ee} // {cod_pai}({qtde})"
                                lista_acumulada[i_ee] = (nivel, cod, nova_qtde, nova_grade_ops)
                                break
                    else:
                        novo_saldo = qtde_float
                        nova_grade = f"{cod_pai}({qtde})"
                        lanca_saldo = (nivel, cod, novo_saldo, nova_grade)
                        lista_acumulada.append(lanca_saldo)

            if lista_acumulada:
                for dedos in lista_acumulada:
                    nivel_a, cod_a, qtde_a, grade = dedos

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                                   f"prod.unidade, prod.conjunto, tip.tipomaterial "
                                   f"FROM produto as prod "
                                   f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                                   f"where prod.codigo = {cod_a};")
                    detalhes_produto = cursor.fetchall()
                    id_prod, codis, descr, ref, um, conjunto, tipo = detalhes_produto[0]

                    if conjunto != 10:
                        dads_compr = (nivel_a, cod_a, descr, ref, um, qtde_a, tipo, grade)
                        nova_lista_compras.append(dads_compr)
                    else:
                        dads_acab = (nivel_a, cod_a, descr, ref, um, qtde_a, tipo, grade)
                        nova_lista_acabado.append(dads_acab)

            if nova_lista_compras:
                self.excel_comprado(caminho, nova_lista_compras)

            if nova_lista_acabado:
                self.excel_acabado(caminho, nova_lista_acabado)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel_total(self, caminho, nova_tabela):
        try:
            lista_nova = []

            for i in nova_tabela:
                nivel, cod, descr, ref, um, qtde, entrega, conj, cod_pai, pacote, num_pi = i

                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.id, tip.tipomaterial "
                               f"FROM produto as prod "
                               f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                               f"where prod.codigo = {cod};")
                detalhes_produto = cursor.fetchall()
                tipo = detalhes_produto[0][1]

                dados = (nivel, cod, descr, ref, um, qtde, entrega, conj, tipo, cod_pai, pacote, num_pi)
                lista_nova.append(dados)

            if lista_nova:
                workbook = criar_workbook()
                sheet = workbook.active
                sheet.title = "Lista Completa"

                headers = ["Nivel", "Código", "Descrição", "Referência", "UM", "Qtde", "Entrega", "Conjunto",
                           "Tipo", "Origem", "Estrutura", "Nº PI"]
                sheet.append(headers)

                header_row = sheet[1]
                for cell in header_row:
                    edita_fonte(cell, negrito=True)
                    edita_preenchimento(cell)
                    edita_alinhamento(cell)

                for d_ex in lista_nova:
                    nivel_ex, cod_ex, de_ex, ref_ex, um_ex, qtde_ex, entr_ex, conj_ex, tipo_ex, cod_pai_ex, \
                    pacote_ex, pi = d_ex

                    nivius = int(nivel_ex)
                    codigu = int(cod_ex)
                    pi_int = int(pi)

                    if cod_pai_ex:
                        cod_pai_ex_int = int(cod_pai_ex)
                    else:
                        cod_pai_ex_int = 0

                    if qtde_ex == "":
                        qtde_e = 0.00
                    else:
                        qtde_e = float(qtde_ex)

                    sheet.append([nivius, codigu, de_ex, ref_ex, um_ex, qtde_e, entr_ex, conj_ex, tipo_ex,
                                  cod_pai_ex_int, pacote_ex, pi_int])

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

    def excel_pedido_interno(self, caminho_arquivo):
        try:
            workbook = carregar_workbook(caminho_arquivo)

            planilha = workbook.create_sheet(title="Situação PI")

            headers = ["Emissão", "Nº PI", "Código", "Descrição", "Referência", "UM", "Qtde", "Entrega", "Projeção",
                       "Qtde Total", "Qtde Falta", "%"]
            planilha.append(headers)

            print("extrair_tabela(self.table_PI)")

            dados_tabela = []

            if dados_tabela:
                for i in dados_tabela:
                    num_pi, codigo, descr, ref, um, qtde, limite, projecao, contagem = i

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT ped.emissao, prod.id, prodint.qtde, prodint.data_previsao "
                                   f"FROM PRODUTOPEDIDOINTERNO as prodint "
                                   f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                                   f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                                   f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                                   f"where prodint.status = 'A' "
                                   f"and prodint.id_pedidointerno = {num_pi} "
                                   f"and prod.codigo = {codigo};")
                    dados_interno = cursor.fetchall()
                    if dados_interno:
                        emissao = dados_interno[0][0]
                        emi = emissao.strftime('%d/%m/%Y')

                        entrega = dados_interno[0][3]
                        ent = entrega.strftime('%d/%m/%Y')

                        id_prod = dados_interno[0][1]

                        num_pi_int = int(num_pi)
                        codigo_int = int(codigo)
                        contagem_int = int(contagem)

                        if qtde == "":
                            qtde_float = 0.00
                        else:
                            qtde_float = float(qtde)

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT produto, qtde_itens FROM RESUMO_ESTRUTURA where produto = {id_prod};")
                        detalhes_resumo = cursor.fetchall()
                        if detalhes_resumo:
                            total_itens = int(detalhes_resumo[0][1])
                        else:
                            total_itens = 0

                        if total_itens:
                            porcentagem = ((total_itens - contagem_int) / total_itens) * 100
                        else:
                            porcentagem = 0

                        porc_int = int(porcentagem)

                        planilha.append([emi, num_pi_int, codigo_int, descr, ref, um, qtde_float,
                                         ent, projecao, total_itens, contagem_int, porc_int])

                        for linha in planilha.iter_rows(min_row=1, max_row=1):
                            for cell in linha:
                                edita_fonte(cell, negrito=True)
                                edita_preenchimento(cell)
                                edita_alinhamento(cell)

                        for linha in planilha.iter_rows(min_row=1,
                                                        max_row=planilha.max_row,
                                                        min_col=1,
                                                        max_col=planilha.max_column):
                            for cell in linha:
                                edita_bordas(cell)
                                edita_alinhamento(cell)

                        for coluna in planilha.columns:
                            max_length = 0
                            for cell in coluna:
                                if isinstance(cell.value, (int, float)):
                                    cell_value_str = "{:.2f}".format(cell.value)
                                else:
                                    cell_value_str = str(cell.value)
                                if len(cell_value_str) > max_length:
                                    max_length = len(cell_value_str)

                            adjusted_width = (max_length + 2)
                            ajusta_larg_coluna(planilha, coluna, adjusted_width)

                        for linha in planilha.iter_rows(min_row=2,
                                                        max_row=planilha.max_row,
                                                        min_col=7,
                                                        max_col=9):
                            for cell in linha:
                                cell.number_format = '0.00'

                        for linha in planilha.iter_rows(min_row=2,
                                                        max_row=planilha.max_row,
                                                        min_col=9,
                                                        max_col=9):
                            for cell in linha:
                                cell.number_format = '0'

            workbook.save(caminho_arquivo)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel_comprado(self, caminho_arquivo, nova_tabela):
        try:
            workbook = carregar_workbook(caminho_arquivo)

            planilha = workbook.create_sheet(title="Comprado")

            headers = ["Nivel", "Código", "Descrição", "Referência", "UM", "Qtde", "Tipo", "Grade"]
            planilha.append(headers)

            regex = re.compile(r"(\d+)?\((\d+\.\d+)\)")

            for dados_ex in nova_tabela:
                nivel, cod, descr, ref, um, qtde, tipo, grade = dados_ex

                matches = regex.findall(grade)
                codigos = [match[0] for match in matches if match[0]]

                grade_final = ""

                for codigo in codigos:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') "
                                   f"FROM produto as prod "
                                   f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                                   f"where prod.codigo = {codigo};")
                    detalhes_pai = cursor.fetchall()
                    codi, descri, refi = detalhes_pai[0]

                    grade_final += f"{codi} - {descri} - {refi} "

                nivius = int(nivel)
                codigu = int(cod)

                if qtde == "":
                    qtde_e = 0.00
                else:
                    qtde_e = float(qtde)

                planilha.append([nivius, codigu, descr, ref, um, qtde_e, tipo, grade_final])

            for linha in planilha.iter_rows(min_row=1, max_row=1):
                for cell in linha:
                    edita_fonte(cell, negrito=True)
                    edita_preenchimento(cell)
                    edita_alinhamento(cell)

            for linha in planilha.iter_rows(min_row=1, max_row=planilha.max_row, min_col=1,
                                            max_col=planilha.max_column):
                for cell in linha:
                    edita_bordas(cell)
                    edita_alinhamento(cell)

            for coluna in planilha.columns:
                max_length = 0
                for cell in coluna:
                    if isinstance(cell.value, (int, float)):
                        cell_value_str = "{:.2f}".format(cell.value)
                    else:
                        cell_value_str = str(cell.value)
                    if len(cell_value_str) > max_length:
                        max_length = len(cell_value_str)

                adjusted_width = (max_length + 2)
                ajusta_larg_coluna(planilha, coluna, adjusted_width)

            for linha in planilha.iter_rows(min_row=2, max_row=planilha.max_row, min_col=7,
                                            max_col=9):
                for cell in linha:
                    cell.number_format = '0.00'

            workbook.save(caminho_arquivo)
            print("Excel Salvo!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel_acabado(self, caminho_arquivo, nova_tabela):
        try:
            lista_nova = []

            for i in nova_tabela:
                nivel, cod, descr, ref, um, qtde, tipo, grade = i
                dados_op = self.manipula_dados_tabela_producao(cod)

                if not dados_op:
                    dados = (nivel, cod, descr, ref, um, qtde, tipo, grade)
                    lista_nova.append(dados)

            if lista_nova:
                workbook = carregar_workbook(caminho_arquivo)
                planilha = workbook.create_sheet(title="Acabado_sem_op")

                headers = ["Nivel", "Código", "Descrição", "Referência", "UM", "Qtde", "Tipo", "Grade"]
                planilha.append(headers)

                header_row = planilha[1]
                for cell in header_row:
                    edita_fonte(cell, negrito=True)
                    edita_preenchimento(cell)
                    edita_alinhamento(cell)

                regex = re.compile(r"(\d+)?\((\d+\.\d+)\)")

                for dados_ex in lista_nova:
                    nivel, cod, descr, ref, um, qtde, tipo, grade = dados_ex

                    matches = regex.findall(grade)
                    codigos = [match[0] for match in matches if match[0]]

                    grade_final = ""

                    for codigo in codigos:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') "
                                       f"FROM produto as prod "
                                       f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                                       f"where prod.codigo = {codigo};")
                        detalhes_pai = cursor.fetchall()
                        codi, descri, refi = detalhes_pai[0]

                        grade_final += f"{codi} - {descri} - {refi} "

                    nivius = int(nivel)
                    codigu = int(cod)

                    if qtde == "":
                        qtde_e = 0.00
                    else:
                        qtde_e = float(qtde)

                    planilha.append([nivius, codigu, descr, ref, um, qtde_e, tipo, grade_final])

                for row in planilha.iter_rows(min_row=1, max_row=planilha.max_row, min_col=1,
                                              max_col=planilha.max_column):
                    for cell in row:
                        edita_bordas(cell)
                        edita_alinhamento(cell)

                for column in planilha.columns:
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
                    planilha.column_dimensions[column_letter].width = adjusted_width

                for row in planilha.iter_rows(min_row=2, max_row=planilha.max_row, min_col=7, max_col=9):
                    for cell in row:
                        cell.number_format = '0.00'

                workbook.save(caminho_arquivo)

                print("Excel Salvo!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


tela = TelaPcpPrevisaoV2()
