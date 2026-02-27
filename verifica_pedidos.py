import sys
from banco_dados.conexao import conecta, conecta_robo
from banco_dados.controle_erros import grava_erro_banco
from comandos.conversores import valores_para_float
import os
import traceback
import inspect


class VerificaPedidos:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.iniciar_processo()

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

    def iniciar_processo(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ''), prod.unidade, prodint.qtde, "
                           f"prodint.data_previsao, prodint.status, prod.quantidade, prod.conjunto, prod.tipomaterial "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_pi, clie, cod, descr, ref, um, qtde, entrega, status, saldo, conj, tipo = i

                    saldo_float = valores_para_float(saldo)

                    if saldo < qtde:
                        qtde_ops = self.manipula_op_aberta(cod)
                        qtde_compra = self.manipula_dado_compra(cod)

                        qtde_consumo = self.manipula_dado_consumo(cod)

                        qtde_venda = self.manipula_dado_venda(cod)

                        if (qtde_venda + qtde_consumo) > (qtde_ops + qtde_compra + saldo_float):
                            if tipo == 119:
                                cursor = conecta.cursor()
                                cursor.execute("""
                                                SELECT id, data_mov 
                                                FROM envia_mov 
                                                WHERE data_mov >= DATEADD(-1 MONTH TO CURRENT_DATE)
                                            """)
                                dados_mov = cursor.fetchall()

                                if dados_mov:
                                    for ii in dados_mov:
                                        id_mov, data_mov = ii

                                        cursor = conecta.cursor()
                                        cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                                                       f"prod.unidade, prod.localizacao, prod.quantidade "
                                                       f"FROM movimentacao AS mov "
                                                       f"INNER JOIN produto prod ON mov.produto = prod.id "
                                                       f"WHERE mov.data = '{data_mov}' "
                                                       f"and mov.tipo < 200 and prod.codigo = '{cod}';")
                                        dados_mov = cursor.fetchall()

                                        if dados_mov:
                                            cur = conecta_robo.cursor()
                                            cur.execute(f"SELECT * from ENVIA_INDUSTRIALIZACAO "
                                                        f"where id_envia_mov = {id_mov} and cod_prod = {cod};")
                                            dados_salvos = cur.fetchall()
                                            if not dados_salvos:
                                                print("Cód.:", cod)
                                                print("Descr:", descr)
                                                print("Qtde:", qtde)
                                                print("Saldo: ", saldo)
                                                print("Qtde ops:", qtde_ops)
                                                print("Qtde Ovs:", qtde_venda)
                                                print()

                            else:
                                print("Cód.:", cod)
                                print("Descr:", descr)
                                print("Qtde:", qtde)
                                print("Saldo: ", saldo)
                                print("Qtde ops:", qtde_ops)
                                print("Qtde Ovs:", qtde_venda)
                                print()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


    def manipula_op_aberta(self, cod_prod):
        try:
            qtde_ops = 0

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.id_estrutura "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where ordser.status = 'A' and prod.codigo = {cod_prod} "
                           f"order by ordser.numero;")
            op_abertas = cursor.fetchall()

            if op_abertas:
                for dados_op in op_abertas:
                    emissao, previsao, op, cod, descr, ref, um, qtde, id_estrut = dados_op

                    qtde_float = valores_para_float(qtde)
                    qtde_ops += qtde_float

            return qtde_ops

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dado_venda(self, cod_prod):
        try:
            qtde_venda = 0

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prodint.qtde, "
                           f"prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A' and prod.codigo = {cod_prod};")
            dados_pi = cursor.fetchall()
            if dados_pi:
                for i_pi in dados_pi:
                    emissao_pi, num_pi, clie_pi, qtde_pi, entrega_pi = i_pi

                    qtde_pi_float = valores_para_float(qtde_pi)

                    qtde_venda += qtde_pi_float

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prodoc.quantidade, prodoc.dataentrega, "
                           f"COALESCE(prodoc.id_pedido, ''), COALESCE(prodoc.id_expedicao, '') "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"LEFT JOIN pedidointerno as ped ON prodoc.id_pedido = ped.id "
                           f"where prodoc.quantidade > prodoc.produzido "
                           f"and oc.status = 'A' "
                           f"and oc.entradasaida = 'S' "
                           f"and prod.codigo = {cod_prod};")
            dados_ov = cursor.fetchall()
            if dados_ov:
                for i_ov in dados_ov:
                    emissao_ov, num_ov, clie_ov, qtde_ov, entrega_ov, num_pi_ov, num_exp = i_ov

                    qtde_ov_float = valores_para_float(qtde_ov)

                    qtde_venda += qtde_ov_float

            return qtde_venda

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dado_consumo(self, cod_prod):
        try:
            qtde_consumo = 0

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, estprod.id_estrutura "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where prod.codigo = {cod_prod};")
            dados_estrut = cursor.fetchall()
            for i_estrut in dados_estrut:
                ides_mat, id_estrutura = i_estrut

                cursor = conecta.cursor()
                cursor.execute(f"select id, id_produto, num_versao, data_versao, obs, data_criacao "
                               f"from estrutura where id = {id_estrutura};")
                estrutura = cursor.fetchall()

                id_produto = estrutura[0][1]

                cursor = conecta.cursor()
                cursor.execute(f"select ordser.datainicial, ordser.numero, prod.codigo, prod.descricao "
                               f"from ordemservico as ordser "
                               f"INNER JOIN produto prod ON ordser.produto = prod.id "
                               f"where ordser.status = 'A' "
                               f"and prod.id = {id_produto} and prod.id_versao = {id_estrutura} "
                               f"order by ordser.numero;")
                op_abertas = cursor.fetchall()
                if op_abertas:
                    for ii in op_abertas:
                        emissao, num_op, cod_pai, descr_pai = ii

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT estprod.id, prod.codigo, "
                                       f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                                       f"(estprod.quantidade)) AS Qtde "
                                       f"FROM estrutura_produto as estprod "
                                       f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                       f"where estprod.id = {ides_mat};")
                        select_estrut = cursor.fetchall()
                        if select_estrut:
                            id_mat, cod_estrut, qtde_total = select_estrut[0]

                            total_float = valores_para_float(qtde_total)

                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT max(prod.codigo), max(prod.descricao), "
                                           f"sum(prodser.QTDE_ESTRUT_PROD) as total "
                                           f"FROM estrutura_produto as estprod "
                                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                           f"INNER JOIN produtoos as prodser ON estprod.id = prodser.ID_ESTRUT_PROD "
                                           f"where estprod.id_estrutura = {id_estrutura} "
                                           f"and prodser.numero = {num_op} and estprod.id = {id_mat} "
                                           f"group by prodser.ID_ESTRUT_PROD;")
                            select_os_resumo = cursor.fetchall()
                            if select_os_resumo:
                                for os_cons in select_os_resumo:
                                    cod_cons, descr_cons, qtde_cons_total = os_cons

                                    cons_float = valores_para_float(qtde_cons_total)

                                    qtde_consumo += total_float - cons_float
                            else:
                                qtde_consumo += total_float

            if qtde_consumo:
                qtde_consumo = round(qtde_consumo, 2)

            return qtde_consumo

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dado_compra(self, cod_prod):
        try:
            qtde_compra = 0

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

                    qtde_sol_float = valores_para_float(qtde_sol)
                    qtde_compra += qtde_sol_float

            cursor = conecta.cursor()
            cursor.execute(f"SELECT sol.idsolicitacao, prodreq.quantidade, req.data, prodreq.numero, "
                           f"prodreq.destino, prodreq.id_prod_sol "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                           f"LEFT JOIN produtoordemsolicitacao as prodsol ON prodreq.id_prod_sol = prodsol.id "
                           f"LEFT JOIN ordemsolicitacao as sol ON prodsol.mestre = sol.idsolicitacao "
                           f"where prodreq.status = 'A' "
                           f"and prod.codigo = {cod_prod};")
            dados_req = cursor.fetchall()

            if dados_req:
                for i_req in dados_req:
                    num_sol_req, qtde_req, emissao_req, num_req, destino, id_prod_sol = i_req

                    qtde_req_float = valores_para_float(qtde_req)
                    qtde_compra += qtde_req_float

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT sol.idsolicitacao, prodreq.numero, oc.data, oc.numero, forn.razao, "
                f"prodoc.quantidade, prodoc.produzido, prodoc.dataentrega "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"LEFT JOIN produtoordemsolicitacao as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"LEFT JOIN ordemsolicitacao as sol ON prodsol.mestre = sol.idsolicitacao "
                f"where oc.entradasaida = 'E' "
                f"AND oc.STATUS = 'A' "
                f"AND prodoc.produzido < prodoc.quantidade "
                f"and prod.codigo = {cod_prod}"
                f"ORDER BY oc.numero;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i_oc in dados_oc:
                    num_sol_oc, id_req_oc, emissao_oc, num_oc, forncec_oc, qtde_oc, prod_oc, entrega_oc = i_oc

                    qtde_oc_float = valores_para_float(qtde_oc)
                    qtde_compra += qtde_oc_float

            return qtde_compra

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = VerificaPedidos()