import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import traceback
import inspect
from comandos.conversores import valores_para_float
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Side, Border
from pathlib import Path



class ExecutaPlanoPcp:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.resultado_ops_vinculos = {}
        self.resultado_sol_vinculos = {}
        self.resultado_req_vinculos = {}
        self.resultado_oc_vinculos = {}

        # inicia processo
        self.iniciar_tudo()

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

    def buscar_vendas_e_consumos(self, codigo_produto, lista_final, visitados):
        try:
            # Evitar repetição / ciclos
            if codigo_produto in visitados:
                return
            visitados.add(codigo_produto)

            # 1) Buscar vendas desse produto
            dados_venda = self.manipula_dados_tabela_venda(codigo_produto)
            if dados_venda:
                for item in dados_venda:
                    # evitar duplicatas na lista da OP
                    if item not in lista_final:
                        lista_final.append(item)

            # 2) Buscar consumos -> localizar pais e subir recursivamente
            dados_consumo, tabela_industrializados = self.manipula_dados_tabela_consumo(codigo_produto)
            if dados_consumo:
                for consumo in dados_consumo:
                    # consumo: (emis, num_op, qtde_total, qtde_cons_total, cod_pai, descr_pai)
                    emis, num_op, qtde_total, qtde_cons_total, cod_pai, descr_pai = consumo

                    # recursão para o pai (cod_pai)
                    self.buscar_vendas_e_consumos(cod_pai, lista_final, visitados)

            if tabela_industrializados:
                for indi in tabela_industrializados:
                    cod_t, descr_t, ref_t, tipo_t = indi

                    # recursão para o pai (cod_pai)
                    self.buscar_vendas_e_consumos(cod_t, lista_final, visitados)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def iniciar_tudo(self):
        try:
            lista_final = []
            lista_sem_destino = []

            lista_final, lista_sem_destino = self.executa_pelas_ops_abertas(lista_final, lista_sem_destino)

            lista_final, lista_sem_destino = self.executa_pelas_sol_abertas(lista_final, lista_sem_destino)

            lista_final, lista_sem_destino = self.executa_pelas_req_abertas(lista_final, lista_sem_destino)

            lista_final, lista_sem_destino = self.executa_pelas_oc_abertas(lista_final, lista_sem_destino)

            if lista_final:
                self.insert_vinculos(lista_final)

            if lista_sem_destino:
                self.gerar_excel("\PI sem vinculos.xlsx", lista_sem_destino)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def insert_vinculos(self, lista_final):
        try:
            for i in lista_final:
                tipo_vinculo, num_op, cod, descr, ref, um, num_pi, id_prod_pi = i

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, descricao, obs, tipomaterial FROM produto where codigo = '{cod}';")
                select_prod = cursor.fetchall()
                id_prod_t, descr_t, ref_t, tipo_t = select_prod[0]

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, id_pedidointerno, tipo, numero, id_produto "
                               f"FROM VINCULO_PRODUTO_PI "
                               f"where id_pedidointerno = {num_pi} "
                               f"and tipo = '{tipo_vinculo}' "
                               f"and numero = {num_op} "
                               f"and id_produto = {id_prod_t} "
                               f"and ID_PRODUTO_PI = {id_prod_pi};")
                consulta_vinculos = cursor.fetchall()

                if len(consulta_vinculos) > 1:
                    id_vinculo_del = consulta_vinculos[1][0]

                    cursor = conecta.cursor()
                    cursor.execute(f"DELETE FROM VINCULO_PRODUTO_PI WHERE ID = {id_vinculo_del};")

                    conecta.commit()
                    print(f"ID {id_vinculo_del} CANCELADO")
                if not consulta_vinculos:
                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into VINCULO_PRODUTO_PI "
                                   f"(id, id_pedidointerno, id_produto_pi, tipo, numero, id_produto) "
                                   f"values (GEN_ID(GEN_VINCULO_PRODUTO_PI_ID,1), {num_pi}, {id_prod_pi}, '{tipo_vinculo}', "
                                   f"{num_op}, {id_prod_t});")

                    conecta.commit()
                    print(f"produtos {num_op, cod, descr, ref, um} vinculados a PI com sucesso!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def executa_pelas_ops_abertas(self, lista_final, lista_sem_destino):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"""
                SELECT op.numero, op.codigo, op.id_estrutura, prod.descricao,
                       COALESCE(prod.obs, ''), prod.unidade,
                       COALESCE(prod.tipomaterial, ''), op.quantidade
                FROM ordemservico as op
                INNER JOIN produto as prod ON op.produto = prod.id
                WHERE op.status = 'A'
                ORDER BY op.numero;
            """)
            ops_abertas = cursor.fetchall()

            if ops_abertas:
                for i in ops_abertas:
                    num_op, cod, id_estrut, descr, ref, um, tipo, qtde = i

                    # Lista específica para esta OP
                    vincular_pedido_interno = []

                    # conjunto para evitar ciclos nesta árvore (por OP)
                    visitados = set()

                    # Chama a recursão que vai subir (filho -> pai -> avô...) e
                    # juntar todas as vendas encontradas em vincular_pedido_interno
                    self.buscar_vendas_e_consumos(cod, vincular_pedido_interno, visitados)

                    # Guarda no atributo (dicionário) com o número da OP como chave
                    self.resultado_ops_vinculos[num_op] = vincular_pedido_interno

                    if vincular_pedido_interno:
                        for item in vincular_pedido_interno:
                            num_pi, num_ov, num_exp, emi_ov, clie_ov, id_prod_pi, qtde_ov, entreg_ov = item

                            dados = ("OP", num_op, cod, descr, ref, um, num_pi, id_prod_pi)
                            lista_final.append(dados)
                    else:
                        dados = ("OP", num_op, cod, descr, ref, um, "", "")
                        lista_sem_destino.append(dados)

            return lista_final, lista_sem_destino

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def executa_pelas_sol_abertas(self, lista_final, lista_sem_destino):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE(prodreq.mestre, ''), prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ''), prod.unidade "
                           f"FROM produtoordemsolicitacao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"INNER JOIN ordemsolicitacao as req ON prodreq.mestre = req.idsolicitacao "
                           f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                           f"WHERE prodreq.status = 'A' "
                           f"AND preq.id_prod_sol IS NULL "
                           f"ORDER BY prodreq.mestre;")
            dados_sol = cursor.fetchall()

            if dados_sol:
                for i_sol in dados_sol:
                    num_sol, cod_prod, descr, ref, um = i_sol

                    dedos_sol = (num_sol, cod_prod, descr, ref, um)
                    tabela_nova.append(dedos_sol)

            if tabela_nova:
                for iiii in tabela_nova:
                    num_sol_i, cod_prod_i, descr, ref, um = iiii

                    # Lista específica para esta OP
                    vincular_pedido_interno = []

                    # conjunto para evitar ciclos nesta árvore (por OP)
                    visitados = set()

                    # Chama a recursão que vai subir (filho -> pai -> avô...) e
                    # juntar todas as vendas encontradas em vincular_pedido_interno
                    self.buscar_vendas_e_consumos(cod_prod_i, vincular_pedido_interno, visitados)

                    # Guarda no atributo (dicionário) com o número da OP como chave
                    self.resultado_sol_vinculos[num_sol_i] = vincular_pedido_interno

                    if vincular_pedido_interno:
                        for item in vincular_pedido_interno:
                            num_pi, num_ov, num_exp, emi_ov, clie_ov, id_prod_pi, qtde_ov, entreg_ov = item

                            dados = ("SOL", num_sol_i, cod_prod_i, descr, ref, um, num_pi, id_prod_pi)
                            lista_final.append(dados)
                    else:
                        dados = ("SOL", num_sol_i, cod_prod_i, descr, ref, um, "", "")
                        lista_sem_destino.append(dados)

            return lista_final, lista_sem_destino


        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def executa_pelas_req_abertas(self, lista_final, lista_sem_destino):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT sol.idsolicitacao, prodreq.numero, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ''), prod.unidade "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                           f"LEFT JOIN produtoordemsolicitacao as prodsol ON prodreq.id_prod_sol = prodsol.id "
                           f"LEFT JOIN ordemsolicitacao as sol ON prodsol.mestre = sol.idsolicitacao "
                           f"where prodreq.status = 'A';")
            dados_req = cursor.fetchall()

            if dados_req:
                for i_req in dados_req:
                    num_sol_req, num_req, cod_prod, descr, ref, um = i_req

                    dedos_req = (num_req, cod_prod, descr, ref, um)
                    tabela_nova.append(dedos_req)

            if tabela_nova:
                for iiii in tabela_nova:
                    num_req_i, cod_prod_i, descr, ref, um = iiii

                    # Lista específica para esta OP
                    vincular_pedido_interno = []

                    # conjunto para evitar ciclos nesta árvore (por OP)
                    visitados = set()

                    # Chama a recursão que vai subir (filho -> pai -> avô...) e
                    # juntar todas as vendas encontradas em vincular_pedido_interno
                    self.buscar_vendas_e_consumos(cod_prod_i, vincular_pedido_interno, visitados)

                    # Guarda no atributo (dicionário) com o número da OP como chave
                    self.resultado_req_vinculos[num_req_i] = vincular_pedido_interno

                    if vincular_pedido_interno:
                        for item in vincular_pedido_interno:
                            num_pi, num_ov, num_exp, emi_ov, clie_ov, id_prod_pi, qtde_ov, entreg_ov = item

                            dados = ("REQ", num_req_i, cod_prod_i, descr, ref, um, num_pi, id_prod_pi)
                            lista_final.append(dados)
                    else:
                        dados = ("REQ", num_req_i, cod_prod_i, descr, ref, um,  "", "")
                        lista_sem_destino.append(dados)

            return lista_final, lista_sem_destino


        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def executa_pelas_oc_abertas(self, lista_final, lista_sem_destino):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT sol.idsolicitacao, prodreq.numero, oc.numero, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ''), prod.unidade "
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
                f"ORDER BY oc.numero;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i_oc in dados_oc:
                    num_sol_oc, id_req_oc, num_oc, cod_prod, descr, ref, um = i_oc

                    dedos_oc = (num_oc, cod_prod, descr, ref, um)
                    tabela_nova.append(dedos_oc)

            if tabela_nova:
                for iiii in tabela_nova:
                    num_oc_i, cod_prod_i, descr, ref, um = iiii

                    # Lista específica para esta OP
                    vincular_pedido_interno = []

                    # conjunto para evitar ciclos nesta árvore (por OP)
                    visitados = set()

                    # Chama a recursão que vai subir (filho -> pai -> avô...) e
                    # juntar todas as vendas encontradas em vincular_pedido_interno
                    self.buscar_vendas_e_consumos(cod_prod_i, vincular_pedido_interno, visitados)

                    # Guarda no atributo (dicionário) com o número da OP como chave
                    self.resultado_oc_vinculos[num_oc_i] = vincular_pedido_interno

                    if vincular_pedido_interno:
                        for item in vincular_pedido_interno:
                            num_pi, num_ov, num_exp, emi_ov, clie_ov, id_prod_pi, qtde_ov, entreg_ov = item

                            dados = ("OC", num_oc_i, cod_prod_i, descr, ref, um, num_pi, id_prod_pi)
                            lista_final.append(dados)
                    else:
                        dados = ("OC", num_oc_i, cod_prod_i, descr, ref, um,  "", "")
                        lista_sem_destino.append(dados)

            return lista_final, lista_sem_destino


        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_consumo(self, cod_prod):
        try:
            tabela_nova = []
            tabela_industrializados = []

            qtde_necessidade = 0

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, estprod.id, estprod.id_estrutura "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where prod.codigo = {cod_prod};")
            dados_estrut = cursor.fetchall()
            for i_estrut in dados_estrut:
                prod_id, ides_mat, id_estrutura = i_estrut

                cursor = conecta.cursor()
                cursor.execute(f"select id, id_produto, num_versao, data_versao, obs, data_criacao "
                               f"from estrutura where id = {id_estrutura};")
                estrutura = cursor.fetchall()

                id_produto = estrutura[0][1]

                cursor = conecta.cursor()
                cursor.execute(f"SELECT codigo, descricao, obs, tipomaterial FROM produto where id = {id_produto};")
                select_prod = cursor.fetchall()
                cod_t, descr_t, ref_t, tipo_t = select_prod[0]

                if tipo_t == 119:
                    dadis = (cod_t, descr_t, ref_t, tipo_t)
                    tabela_industrializados.append(dadis)
                else:
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

                            emis = f'{emissao.day}/{emissao.month}/{emissao.year}'

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

                                        qtde_necessidade += total_float - cons_float

                                        if qtde_necessidade > 0:
                                            dados = (emis, num_op, qtde_total, qtde_cons_total, cod_pai, descr_pai)

                                            tabela_nova.append(dados)

                                else:
                                    dados = (emis, num_op, qtde_total, "0", cod_pai, descr_pai)
                                    tabela_nova.append(dados)

                                    qtde_necessidade += total_float

            return tabela_nova, tabela_industrializados

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_venda(self, cod_prod):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prod.id, prodint.qtde, "
                           f"prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A' and prod.codigo = {cod_prod};")
            dados_pi = cursor.fetchall()

            if dados_pi:
                for i_pi in dados_pi:
                    emissao_pi, num_pi, clie_pi, id_prod, qtde_pi, entrega_pi = i_pi

                    emi_pi = f'{emissao_pi.day}/{emissao_pi.month}/{emissao_pi.year}'
                    entreg_pi = f'{entrega_pi.day}/{entrega_pi.month}/{entrega_pi.year}'

                    dados_pi = (num_pi, "", "", emi_pi, clie_pi, id_prod, qtde_pi, entreg_pi)
                    tabela_nova.append(dados_pi)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prod.id, prodoc.quantidade, prodoc.dataentrega, "
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
                    emissao_ov, num_ov, clie_ov, id_prod, qtde_ov, entrega_ov, num_pi_ov, num_exp = i_ov

                    emi_ov = f'{emissao_ov.day}/{emissao_ov.month}/{emissao_ov.year}'
                    entreg_ov = f'{entrega_ov.day}/{entrega_ov.month}/{entrega_ov.year}'

                    dados = (num_pi_ov, num_ov, num_exp, emi_ov, clie_ov, id_prod, qtde_ov, entreg_ov)
                    tabela_nova.append(dados)

            return tabela_nova

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def gerar_excel(self, nome_arquivo, dados_final):
        try:
            # Criação do DataFrame
            df = pd.DataFrame(dados_final, columns=[
                'Tipo Vinculo', 'Nº', 'Código', 'Descrição', 'Referência', 'UM', 'Nº PI', 'Cod Prod PI'])

            # Conversão dos tipos de dados
            df['Nº'] = df['Nº'].astype(int)
            df['Código'] = df['Código'].astype(int)

            desktop = Path.home() / "Desktop"
            caminho = str(desktop) + nome_arquivo

            df.to_excel(caminho, index=False)

            workbook = load_workbook(caminho)
            sheet = workbook.active

            # Ajustar a largura das colunas
            for column in sheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except TypeError:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column[0].column_letter].width = adjusted_width

            thin = Side(border_style="thin", color="000000")
            for cell in sheet[1]:
                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

            workbook.save(caminho)

            print(f'Dados salvos em {caminho}')


        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = ExecutaPlanoPcp()
