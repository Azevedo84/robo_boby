import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from datetime import timedelta, date, datetime
import inspect
import os
import math
import traceback
from pathlib import Path
import pandas as pd


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
            cod_maquina = "21404"
            qtde = 1
            nivel = 1

            estrutura = self.calculo_3_verifica_estrutura(nivel, cod_maquina, qtde)

            if estrutura:
                self.exportar_para_excel(estrutura)
                for i in estrutura:
                    print(i)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_3_verifica_estrutura(self, nivel, codigo, quantidade):
        try:
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

            dadoss = (nivel, cod_pai, descr_pai, ref_pai, um_pai, quantidade)
            filhos = [dadoss]

            if id_estrut:
                nivel_plus = nivel + 1

                cursor.execute(
                    f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                    f"(estprod.quantidade * {quantidade}) as qtde "
                    f"FROM estrutura_produto as estprod "
                    f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                    f"WHERE estprod.id_estrutura = {id_estrut};")
                dados_estrutura = cursor.fetchall()
                if dados_estrutura:
                    for prod in dados_estrutura:
                        cod_f, descr_f, ref_f, um_f, qtde_f = prod

                        filhos.extend(self.calculo_3_verifica_estrutura(nivel_plus, cod_f, qtde_f))

            return filhos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def exportar_para_excel(self, estrutura):
        try:
            df = pd.DataFrame(estrutura, columns=['Nível', 'Código', 'Descrição', 'Referência', 'Unidade',
                                                  'Quantidade'])

            # Função para deslocar os dados com base no nível
            def deslocar_dados(row):
                nivel = row['Nível']
                deslocamento = (nivel - 1) * 6  # Cada nível desloca 6 colunas
                return pd.Series(row.values, index=pd.Index(range(deslocamento, deslocamento + len(row))),
                                 name=row.name)

            # Aplica a função para deslocar os dados
            novo_df = df.apply(deslocar_dados, axis=1)

            # Preencher células vazias
            novo_df.fillna('', inplace=True)  # Preencher células vazias

            desktop = Path.home() / "Desktop"
            desk_str = str(desktop)
            nome_req = '\ Nivel Estrutura.xlsx'
            caminho = (desk_str + nome_req)

            # Salvar para Excel (apenas exemplo, ajuste conforme necessário)
            novo_df.to_excel(caminho, index=False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = PcpPrevisao()
chama_classe.calculo_1_dados_previsao()
