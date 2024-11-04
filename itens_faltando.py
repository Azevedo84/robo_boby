import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from comandos.conversores import valores_para_float
import os
import traceback
import inspect
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Border, Side


class EnviaOrdensProducao:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

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

    def estrutura_prod_qtde_op(self, num_op, id_estrut):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                           f"(estprod.quantidade)) AS Qtde, prod.quantidade, tip.tipomaterial "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                           f"where estprod.id_estrutura = {id_estrut} ORDER BY prod.descricao;")
            sel_estrutura = cursor.fetchall()

            return sel_estrutura

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def teste(self, id_mat_e, lista_substitutos, num_op):
        try:
            saldo_substituto = 0

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id = {id_mat_e};")
            sel_estrutura = cursor.fetchall()

            cod_original = sel_estrutura[0][1]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT subs.cod_subs, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"conj.conjunto, prod.unidade, prod.localizacao, prod.quantidade "
                           f"FROM SUBSTITUTO_MATERIAPRIMA as subs "
                           f"INNER JOIN produto as prod ON subs.cod_subs = prod.codigo "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"WHERE subs.id_mat = {id_mat_e} "
                           f"and prod.quantidade > 0 "
                           f"AND subs.num_op = {num_op};")
            dados_mat_com_op = cursor.fetchall()
            if dados_mat_com_op:
                cod_subs, descr, ref, conj, um, local, saldo = dados_mat_com_op[0]

                dados = (cod_subs, descr, ref, um, local, saldo, cod_original)
                lista_substitutos.append(dados)

                saldo_substituto = float(saldo)

            else:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT subs.cod_subs, prod.descricao, COALESCE(prod.obs, '') as obs, "
                               f"conj.conjunto, prod.unidade, prod.localizacao, prod.quantidade "
                               f"FROM SUBSTITUTO_MATERIAPRIMA as subs "
                               f"INNER JOIN produto as prod ON subs.cod_subs = prod.codigo "
                               f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                               f"WHERE subs.id_mat = {id_mat_e} "
                               f"and prod.quantidade > 0 "
                               f"AND subs.num_op is NULL;")
                dados_mat_sem_op = cursor.fetchall()
                if dados_mat_sem_op:
                    cod_subs, descr, ref, conj, um, local, saldo = dados_mat_sem_op[0]

                    dados = (cod_subs, descr, ref, um, local, saldo, cod_original)
                    lista_substitutos.append(dados)

                    saldo_substituto = float(saldo)

            return lista_substitutos, saldo_substituto

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consumo_op_por_id(self, num_op, id_materia_prima):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT SUM(prodser.quantidade) AS total_quantidade "
                           f"FROM produtoos AS prodser "
                           f"WHERE prodser.numero = {num_op} AND prodser.id_estrut_prod = {id_materia_prima};")
            total_quantidade = cursor.fetchone()[0]  # Acessa o primeiro elemento retornado para obter o valor da soma

            # Verifica se a consulta retornou um valor (para evitar erros se não houver registros)
            total_quantidade = total_quantidade if total_quantidade is not None else 0

            return total_quantidade

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_ops_concluidas(self, num_op, id_estrut, cod_pai, descr_pai):
        try:
            lista_substitutos = []

            select_estrut = self.estrutura_prod_qtde_op(num_op, id_estrut)

            falta_material = []

            for dados_estrut in select_estrut:
                saldo_final = 0

                id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_e, saldo_prod_e, tipo_e = dados_estrut

                #print("OP ", self.num_op, "dados estrutura: ", dados_estrut)

                lista_substitutos, saldo_subs = self.teste(id_mat_e, lista_substitutos, num_op)

                qtde_e_float = valores_para_float(qtde_e)
                saldo_final += saldo_subs + valores_para_float(saldo_prod_e)

                qtde_total_item_op = self.consumo_op_por_id(num_op, id_mat_e)

                qtde_total_item_op_float = valores_para_float(qtde_total_item_op)

                if qtde_total_item_op_float < qtde_e_float:
                    #print("qtde estrut: ", qtde_e_float, "estoque: ", saldo_final, "qtde ops: ", qtde_total_item_op_float)
                    sobras = qtde_e_float - saldo_final - qtde_total_item_op_float
                    if sobras > 0:
                        dados = (num_op, cod_pai, descr_pai, cod_e, descr_e, ref_e, um_e, sobras, tipo_e)
                        falta_material.append(dados)
                        #print("falta material: ", cod_e, descr_e, sobras)

                #print("\n\n")

            return falta_material

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def gerar_excel(self, material_faltando):
        try:
            # Criação do DataFrame
            df = pd.DataFrame(material_faltando, columns=[
                'Nº OP', 'Cód. Pai', 'Descrição Pai', 'Cód.', 'Descrição', 'Referência', 'Um', 'Qtde', 'Tipo'])

            # Conversão dos tipos de dados
            df['Nº OP'] = df['Nº OP'].astype(int)
            df['Cód. Pai'] = df['Cód. Pai'].astype(int)
            df['Cód.'] = df['Cód.'].astype(int)
            df['Qtde'] = df['Qtde'].astype(float)

            desktop = Path.home() / "Desktop"
            nome_req = '\material_faltando.xlsx'
            caminho = str(desktop) + nome_req

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
                    except:
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

    def manipula_comeco(self):
        try:
            tudo_tudo = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT op.numero, op.codigo, op.id_estrutura, prod.descricao, "
                           f"COALESCE(prod.obs, ''), "
                           f"prod.unidade, COALESCE(prod.tipomaterial, ''), op.quantidade "
                           f"FROM ordemservico as op "
                           f"INNER JOIN produto as prod ON op.produto = prod.id "
                           f"where op.status = 'A';")
            ops_abertas = cursor.fetchall()

            if ops_abertas:
                for i in ops_abertas:
                    num_op, cod, id_estrut, descr, ref, um, tipo, qtde = i

                    material_faltando = self.verifica_ops_concluidas(num_op, id_estrut, cod, descr)

                    if material_faltando:
                        for ii in material_faltando:
                            print(ii)
                            tudo_tudo.append(ii)

            if tudo_tudo:
                for iii in tudo_tudo:
                    num_op, cod_pai, descr_pai, cod_e, descr_e, ref_e, um_e, sobras, tipo_e = iii
                    criar ordens de compra pendente de cada item
                self.gerar_excel(tudo_tudo)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaOrdensProducao()
chama_classe.manipula_comeco()
