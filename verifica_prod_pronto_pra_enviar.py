import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from dados_email import email_user, password
from comandos.conversores import valores_para_float
import os
import traceback
import inspect
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import smtplib


class VerificaProdutosProntos:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.manipula_dados_cliente()

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

            return saudacao, msg_final, to

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_problema(self, cod, problema):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'Custo Produto - Problemas no calculo de Custo do produto {cod}!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f'{saudacao}\n\n' \
                   f'Houve algum problema com o produto {cod}:\n\n' \
                   f'{problema}!\n\n' \
                   f'{msg_final}'

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)

            server.quit()

            print(f'Email do erro da OP foi enviado com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_cliente(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.id, COALESCE(ped.NUM_REQ_CLIENTE, '') as reqs, cli.razao, "
                           f"prod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, COALESCE(prod.NCM, ''), "
                           f"prod.unidade, prodint.qtde, prodint.data_previsao, prod.quantidade, "
                           f"prod.conjunto, prod.terceirizado, prod.custounitario, prod.custoestrutura, "
                           f"COALESCE(ped.obs, '') as obs, ped.solicitante "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    (num_ped, req, cliente, id_cod, cod, descr, ref, ncm, um, qtde, ent, saldo, conj, terc,
                     unit, estrut, obs, solic) = i

                    if conj == 10:
                        self.lanca_dados_acabado(cod)
                    else:
                        self.lanca_dados_compras(cod)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_acabado(self, codigo_produto):
        try:
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.descricao, COALESCE(tip.tipomaterial, '') as tipus, "
                        f"COALESCE(prod.obs, '') as ref, prod.unidade, "
                        f"COALESCE(prod.ncm, '') as ncm, COALESCE(prod.obs2, '') as obs "
                        f"FROM produto as prod "
                        f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                        f"where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descr, tipo, ref, um, ncm, obs = detalhes_produto[0]

            if not tipo:
                msg = 'O campo "Tipo de Material" não pode estar vazio!'
                self.envia_email_problema(codigo_produto, msg)
            else:
                versao = self.lanca_versoes(codigo_produto)

                if codigo_produto and versao:
                    self.produtos_problema(codigo_produto, versao)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_versoes(self, codigo_produto):
        try:
            versao = 0

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod, id_versao = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                           f"from estrutura "
                           f"where id_produto = {id_pai} order by data_versao;")
            tabela_versoes = cursor.fetchall()

            if tabela_versoes:
                for i in tabela_versoes:
                    id_estrut, num_versao, data, obs, criacao = i

                    if id_versao == id_estrut:
                        versao = num_versao

            return versao

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def produtos_problema(self, codigo_produto, versao):
        try:
            tabela_nova = []

            extrai_estrutura = self.lanca_estrutura(codigo_produto, versao)

            if extrai_estrutura:
                tab_estrut = []
                for itens in extrai_estrutura:
                    cod = itens[0]
                    qtde = itens[4]

                    estrutura = self.verifica_estrutura_problema(1, cod, qtde)
                    if estrutura:
                        for i in estrutura:
                            tab_estrut.append(i)

                tab_ordenada = sorted(tab_estrut, key=lambda x: -x[0])

                for i in tab_ordenada:
                    niv, codi, descr, ref, um, qtdi, conj, temp, terc, unit, estrut = i

                    if conj == 10:
                        if temp or terc:
                            pass
                        else:
                            if estrut:
                                estrut_float = float(estrut)
                            else:
                                estrut_float = 0
                            total = float(qtdi) * estrut_float

                            dados = (codi, descr, ref, um, qtdi, estrut_float, total, conj)
                            tabela_nova.append(dados)
                    else:
                        if not unit:
                            unit_float = 0

                            total = float(qtdi) * float(unit_float)

                            dados = (codi, descr, ref, um, qtdi, unit_float, total, conj)
                            tabela_nova.append(dados)

                if tabela_nova:
                    tabela_nova_ordenada = sorted(tabela_nova, key=lambda x: (x[1], x[0]))

                    self.envia_email_problema(codigo_produto, tabela_nova_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_estrutura(self, codigo_produto, num_versao):
        try:
            nova_tabela = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                           f"from estrutura "
                           f"where id_produto = {id_pai} and num_versao = {num_versao};")
            tabela_versoes = cursor.fetchall()
            id_estrutura = tabela_versoes[0][0]

            cursor = conecta.cursor()
            cursor.execute(f"UPDATE produto SET custoestrutura = '{0}' where id = {id_pai};")
            conecta.commit()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"prod.conjunto, prod.unidade, "
                           f"(estprod.quantidade * 1) as qtde, prod.terceirizado, prod.custounitario, "
                           f"prod.custoestrutura "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"where estprod.id_estrutura = {id_estrutura} "
                           f"order by conj.conjunto DESC, prod.descricao ASC;")
            tabela_estrutura = cursor.fetchall()

            if tabela_estrutura:
                for i in tabela_estrutura:
                    cod, descr, ref, conjunto, um, qtde, terc, unit, estrut = i

                    qtde_float = valores_para_float(qtde)
                    unit_float = valores_para_float(unit)
                    estrut_float = valores_para_float(estrut)

                    if conjunto == 10:
                        total = qtde_float * estrut_float

                        dados = (cod, descr, ref, um, qtde, estrut, total, conjunto)
                        nova_tabela.append(dados)
                    else:
                        total = qtde_float * unit_float

                        dados = (cod, descr, ref, um, qtde, unit, total, conjunto)
                        nova_tabela.append(dados)

            return nova_tabela

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_estrutura_problema(self, nivel, codigo, qtde):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, prod.obs, prod.unidade, "
                           f"prod.conjunto, prod.tempo, prod.terceirizado, prod.custounitario, "
                           f"prod.custoestrutura, prod.id_versao "
                           f"FROM produto as prod "
                           f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"where prod.codigo = {codigo};")
            detalhes_pai = cursor.fetchall()
            (id_pai, c_pai, des_pai, ref_pai, um_pai, conj_pai, temp_pai, terc_pai, unit_pai, est_pai,
             id_estrut) = detalhes_pai[0]

            filhos = [(nivel, codigo, des_pai, ref_pai, um_pai, qtde, conj_pai, temp_pai, terc_pai, unit_pai, est_pai)]

            nivel_plus = nivel + 1

            if id_estrut:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                               f"(estprod.quantidade * {qtde}) as qtde "
                               f"FROM estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"WHERE estprod.id_estrutura = {id_estrut};")
                dados_estrutura = cursor.fetchall()

                if dados_estrutura:
                    for prod in dados_estrutura:
                        cod_f, descr_f, ref_f, um_f, qtde_f = prod

                        filhos.extend(self.verifica_estrutura_problema(nivel_plus, cod_f, qtde_f))

            return filhos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_compras(self, codigo_produto):
        try:
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.descricao, COALESCE(tip.tipomaterial, '') as tipus, "
                        f"COALESCE(prod.obs, '') as ref, prod.unidade, "
                        f"COALESCE(prod.ncm, '') as ncm, COALESCE(prod.obs2, '') as obs, prod.conjunto "
                        f"FROM produto as prod "
                        f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                        f"where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descr, tipo, ref, um, ncm, obs, id_conj = detalhes_produto[0]

            if not tipo:
                msg = 'O campo "Tipo de Material" não pode estar vazio!'
                self.envia_email_problema(codigo_produto, msg)
            else:
                self.lanca_descricao_custo_compra(codigo_produto)


        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_descricao_custo_compra(self, codigo):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, custounitario FROM produto WHERE codigo = {codigo};")
            dados_produto = cursor.fetchall()
            if dados_produto:
                for i in dados_produto:
                    id_prod, custo_compra = i

                    if not custo_compra:
                        msg = "FALTA CUSTO DO PRODUTO DE COMPRA:"
                        self.envia_email_problema(codigo, msg)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == "__main__":
    try:
        chama_classe = VerificaProdutosProntos()
    finally:
        try:
            conecta.close()
        except:
            pass

        import sys
        sys.exit(0)
