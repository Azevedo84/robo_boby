import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from conversores import valores_para_float
import os
import inspect
from datetime import datetime
import traceback
import math


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

    def dados_email(self):
        try:
            to = ['<maquinas@unisold.com.br>']

            current_time = (datetime.now())
            horario = current_time.strftime('%H')
            hora_int = int(horario)
            saudacao = ""
            if 4 < hora_int < 13:
                saudacao = "Bom Dia!"
            elif 12 < hora_int < 19:
                saudacao = "Boa Tarde!"
            elif hora_int > 18:
                saudacao = "Boa Noite!"
            elif hora_int < 5:
                saudacao = "Boa Noite!"

            msg_final = f"Att,\n" \
                        f"Suzuki Máquinas Ltda\n" \
                        f"Fone (51) 3561.2583/(51) 3170.0965\n\n" \
                        f"Mensagem enviada automaticamente, por favor não responda.\n\n" \
                        f"Se houver algum problema com o recebimento de emails ou conflitos com o arquivo excel, " \
                        f"favor entrar em contato pelo email maquinas@unisold.com.br.\n\n"

            email_user = 'ti.ahcmaq@gmail.com'
            password = 'poswxhqkeaacblku'

            return saudacao, msg_final, email_user, to, password

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def ordens_producao_abertas(self, cod_pai, cod_filho, lista_ops):
        try:
            qtde_ops = 0

            if cod_pai:
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
                        id_produto = i[3]

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT mat.id, prod.codigo, prod.descricao, "
                                       f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                                       f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                                       f"(mat.quantidade)) AS Qtde, "
                                       f"COALESCE(prod.localizacao, ''), prod.quantidade "
                                       f"FROM materiaprima as mat "
                                       f"INNER JOIN produto as prod ON mat.produto = prod.id "
                                       f"where mat.mestre = {id_produto} and prod.codigo = {cod_filho} "
                                       f"ORDER BY prod.descricao;")
                        select_estrut = cursor.fetchall()
                        if select_estrut:
                            for dados_estrut in select_estrut:
                                id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_e, local_e, saldo_e = dados_estrut

                                cursor = conecta.cursor()
                                cursor.execute(f"SELECT max(mat.id), max(prod.codigo), max(prod.descricao), "
                                               f"sum(prodser.qtde_materia)as total "
                                               f"FROM materiaprima as mat "
                                               f"INNER JOIN produto as prod ON mat.produto = prod.id "
                                               f"INNER JOIN produtoos as prodser ON mat.id = prodser.id_materia "
                                               f"where mat.mestre = {id_produto} "
                                               f"and prodser.numero = {num_op} and mat.id = {id_mat_e} "
                                               f"group by prodser.id_materia;")
                                select_os_resumo = cursor.fetchall()

                                if select_os_resumo:
                                    for dados_res in select_os_resumo:
                                        id_mat_sum, cod_sum, descr_sum, qtde_sum = dados_res

                                        qtde_sum_float = valores_para_float(qtde_sum)

                                        prod_op_encontrado = False
                                        for cod_pai_e, cod_filho_e, num_op_e in lista_ops:
                                            print(cod_pai_e, cod_pai, cod_filho_e, cod_filho, num_op_e, num_op)
                                            if cod_pai_e == cod_pai and cod_filho_e == cod_filho and num_op_e == num_op:
                                                prod_op_encontrado = True
                                                break

                                        if not prod_op_encontrado:
                                            qtde_ops += qtde_sum_float

                                            lanca_saldo = (str(cod_pai), str(cod_filho), str(num_op))
                                            lista_ops.append(lanca_saldo)

            return qtde_ops, lista_ops

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

    def manipula_dados_pi(self):
        try:
            dados_p_tabela = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, prodint.id_pedidointerno, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodint.qtde, prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A' "
                           f"order by ped.emissao;")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emiss, num_pi, cod, descr, ref, um, qtde, prev = i

                    emissao = emiss.strftime('%d/%m/%Y')
                    previsao = prev.strftime('%d/%m/%Y')

                    dados = (emissao, num_pi, cod, descr, ref, um, qtde, previsao, "", "")

                    dados_p_tabela.append(dados)

            return dados_p_tabela

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def retorna_calculo_meses(self, dados_tabela):
        try:
            tab_prev = []
            for datas in dados_tabela:
                cod_prev = datas[2]
                dt_previsao = datas[7]
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

    def manipula_comeco(self):
        try:
            tudo_tudo = []

            dados_tabela = self.manipula_dados_pi()

            if dados_tabela:
                saldos = []
                lista_ops = []

                tab_meses = self.retorna_calculo_meses(dados_tabela)

                for i in dados_tabela:
                    emissao, num_pi, codigo, descr, ref, um, qtde, previsao, nivil, entrega = i
                    agrega_pt = 0
                    for cod_isso, meses_isso in tab_meses:
                        if cod_isso == codigo:
                            if meses_isso > 0:
                                agrega_pt = meses_isso * 5

                    pcte_p_estrutura = [agrega_pt, 1, num_pi, codigo, qtde, codigo, descr, saldos, lista_ops, codigo]

                    estrutura = self.calculo_3_verifica_estrutura(pcte_p_estrutura)

                    if estrutura:
                        for ii in estrutura:
                            tudo_tudo.append(ii)

            if tudo_tudo:
                for titi in tudo_tudo:
                    pass
                    # print(titi)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_3_verifica_estrutura(self, dados_total):
        try:
            pontos, nivel, num_pi, cod, qtde, cod_pai, descr_pai, lista_saldos, lista_ops, cod_orig = dados_total

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prod.quantidade, tip.tipomaterial "
                           f"FROM produto as prod "
                           f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                           f"where prod.codigo = {cod};")
            detalhes_pai = cursor.fetchall()
            id_b, cod_b, descr_b, ref_b, um_b, saldo_b, tipo_b = detalhes_pai[0]

            qtde_float = valores_para_float(qtde)
            saldo_float = valores_para_float(saldo_b)

            produzir = 0

            prod_saldo_encontrado = False
            for cod_sal_e, saldo_e in lista_saldos:
                if cod_sal_e == cod_b:
                    prod_saldo_encontrado = True
                    break

            if not prod_saldo_encontrado:
                if saldo_float < qtde_float:
                    novo_saldo_nao_existe = 0
                else:
                    novo_saldo_nao_existe = saldo_float - qtde_float

                if saldo_float >= qtde_float:
                    produzir = 0
                else:
                    produzir = (saldo_float - qtde_float) * -1

                lanca_saldo = (cod_b, novo_saldo_nao_existe)
                lista_saldos.append(lanca_saldo)
            else:
                for i_ee, (cod_ee, saldo_ee) in enumerate(lista_saldos):
                    if cod_ee == cod_b:
                        if saldo_ee < qtde_float:
                            novo_saldo_existe = 0
                        else:
                            novo_saldo_existe = saldo_ee - qtde_float

                        if saldo_ee >= qtde_float:
                            produzir = 0
                        else:
                            produzir = (saldo_ee - qtde_float) * -1

                        lista_saldos[i_ee] = (cod_b, novo_saldo_existe)
                        break

            produzir_arred = round(produzir, 2)

            if produzir_arred:
                print(num_pi, cod_b, descr_b, ref_b, produzir_arred, cod_pai)

                qtde_ops, lista_op_nova = self.ordens_producao_abertas(cod_pai, cod_b, lista_ops)

                if qtde_ops:
                    produzir_final = produzir_arred + qtde_ops
                else:
                    produzir_final = produzir_arred

                dadoss = (pontos, nivel, num_pi, cod_b, descr_b, ref_b, um_b, produzir_final, cod_orig)

                filhos = [dadoss]

                nivel_plus = nivel + 1
                pts_plus = pontos + 1

                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                               f"(mat.quantidade * {produzir_final}) as qtde "
                               f"FROM materiaprima as mat "
                               f"INNER JOIN produto prod ON mat.produto = prod.id "
                               f"where mestre = {id_b};")
                dados_estrutura = cursor.fetchall()

                if dados_estrutura:
                    for prod in dados_estrutura:
                        cod_f, descr_f, ref_f, um_f, qtde_f = prod

                        pcte_filho = [pts_plus, nivel_plus, num_pi, cod_f, qtde_f, cod_b, descr_b, lista_saldos,
                                      lista_op_nova, cod_orig]
                        filhos.extend(self.calculo_3_verifica_estrutura(pcte_filho))

            else:
                filhos = []

            return filhos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaOrdensProducao()
chama_classe.manipula_comeco()
