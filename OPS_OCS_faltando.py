import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from comandos.conversores import valores_para_float
import os
import traceback
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from dados_email import email_user, password


class OPSeOCSFaltando:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.inicio_de_tudo()

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

    def manipula_dados_estrutura(self, cod_prod):
        try:
            nova_tabela = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {cod_prod};")
            select_prod = cursor.fetchall()
            idez, cod, id_estrut = select_prod[0]

            if id_estrut:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                               f"conj.conjunto, prod.unidade, (estprod.quantidade * 1) as qtde, "
                               f"COALESCE(prod.ncm, '') as ncm "
                               f"from estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                               f"where estprod.id_estrutura = {id_estrut} "
                               f"order by conj.conjunto DESC, prod.descricao ASC;")
                tabela_estrutura = cursor.fetchall()

                if tabela_estrutura:
                    for i in tabela_estrutura:
                        cod, descr, ref, conjunto, um, qtde, ncm = i

                        qtde_float = float(qtde)

                        dados = (cod, descr, ref, um, qtde_float)
                        nova_tabela.append(dados)

            return nova_tabela

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def inicio_de_tudo(self):
        try:
            lista_comprado = []
            lista_industrializacao = []
            lista_acabado = []

            cursor = conecta.cursor()
            cursor.execute(
                "SELECT prod.codigo, prod.descricao, "
                "COALESCE(prod.obs, '') as obs, "
                "prod.unidade, "
                "SUM(prodint.qtde) as total_qtde "
                "FROM PRODUTOPEDIDOINTERNO AS prodint "
                "INNER JOIN produto AS prod ON prodint.id_produto = prod.id "
                "INNER JOIN pedidointerno AS ped ON prodint.id_pedidointerno = ped.id "
                "INNER JOIN clientes AS cli ON ped.id_cliente = cli.id "
                "WHERE prodint.status = 'A' "
                "GROUP BY prod.codigo, prod.descricao, prod.obs, prod.unidade "
                "ORDER BY prod.codigo;"
            )
            dados_agrupados = cursor.fetchall()

            if dados_agrupados:
                for i in dados_agrupados:
                    tipo_prod = 0

                    qtde_entradas = 0

                    cod, descr, ref, um, qtde_pi = i

                    qtde_pi_float = valores_para_float(qtde_pi)

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT conjunto, tipomaterial, quantidade FROM produto where codigo = {cod};")
                    select_prod = cursor.fetchall()
                    conjunto, tipo_material, saldo = select_prod[0]

                    if tipo_material:
                        tipo_prod = tipo_material
                    conj_prod = conjunto

                    saldo_float = valores_para_float(saldo)

                    qtde_entradas += saldo_float

                    cursor = conecta.cursor()
                    cursor.execute(
                        "SELECT SUM(ordser.quantidade) "
                        "FROM ordemservico AS ordser "
                        "INNER JOIN produto prod ON ordser.produto = prod.id "
                        "WHERE ordser.status = 'A' "
                        "AND prod.codigo = ? "
                        "GROUP BY prod.codigo;",
                        (cod,)
                    )
                    total_op = cursor.fetchall()

                    if total_op:
                        qtde_op = total_op[0][0]

                        qtde_op_float = valores_para_float(qtde_op)

                        qtde_entradas += qtde_op_float

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT SUM(prodreq.quantidade) "
                                   f"FROM produtoordemsolicitacao as prodreq "
                                   f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                                   f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                                   f"WHERE prodreq.status = 'A' "
                                   f"and prod.codigo = {cod} "
                                   f"AND preq.id_prod_sol IS NULL "
                                   f"GROUP BY prod.codigo;")
                    dados_sol = cursor.fetchall()
                    if dados_sol:
                        qtde_sol = dados_sol[0][0]

                        qtde_sol_float = valores_para_float(qtde_sol)

                        qtde_entradas += qtde_sol_float

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT SUM(prodreq.quantidade) "
                                   f"FROM produtoordemrequisicao as prodreq "
                                   f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                                   f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                                   f"LEFT JOIN produtoordemsolicitacao as prodsol ON prodreq.id_prod_sol = prodsol.id "
                                   f"LEFT JOIN ordemsolicitacao as sol ON prodsol.mestre = sol.idsolicitacao "
                                   f"where prodreq.status = 'A' "
                                   f"and prod.codigo = {cod} "
                                   f"GROUP BY prod.codigo;")
                    dados_req = cursor.fetchall()
                    if dados_req:
                        qtde_req = dados_req[0][0]

                        qtde_req_float = valores_para_float(qtde_req)

                        qtde_entradas += qtde_req_float

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT SUM(prodoc.quantidade) "
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
                                   f"and prod.codigo = {cod} "
                                   f"GROUP BY prod.codigo;")
                    dados_oc = cursor.fetchall()
                    if dados_oc:
                        qtde_oc = dados_oc[0][0]

                        qtde_oc_float = valores_para_float(qtde_oc)

                        qtde_entradas += qtde_oc_float

                    if qtde_pi_float > qtde_entradas:
                        necessidade = qtde_pi_float - qtde_entradas

                        if conj_prod == 10:
                            if tipo_prod == 119:
                                resultado = self.manipula_dados_estrutura(cod)
                                if resultado:
                                    for ii in resultado:
                                        cod_f, descr_f, ref_f, um_f, qtde_f = ii

                                        cursor = conecta.cursor()
                                        cursor.execute(
                                            "SELECT SUM(ordser.quantidade) "
                                            "FROM ordemservico AS ordser "
                                            "INNER JOIN produto prod ON ordser.produto = prod.id "
                                            "WHERE ordser.status = 'A' "
                                            "AND prod.codigo = ? "
                                            "GROUP BY prod.codigo;",
                                            (cod_f,)
                                        )
                                        total_op = cursor.fetchall()

                                        if not total_op:
                                            dados = (cod_f, descr_f, ref_f, um_f, necessidade)
                                            print(cod, dados)
                                            lista_industrializacao.append(dados)

                            else:
                                dados = (cod, descr, ref, um, necessidade)
                                lista_acabado.append(dados)
                        else:
                            dados = (cod, descr, ref, um, necessidade)
                            lista_comprado.append(dados)

            if lista_acabado:
                self.envia_email_acabado(lista_acabado)

            if lista_comprado:
                self.envia_email_comprado(lista_comprado)

            if lista_industrializacao:
                self.envia_email_industrializado(lista_industrializacao)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
            return None

    def envia_email_acabado(self, lista_produtos):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'PCP - Gerar OP dos produtos!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos para Gerar OP:\n\n"

            for i in lista_produtos:
                cod, descr, ref, um, necessidade = i

                body += f"- Código: {cod}\n" \
                        f"- Descrição: {descr}\n UM: {um}" \
                        f"- Referência: {ref}\n" \
                        f"- Qtde Necessidade: {necessidade}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print(f'EMAIL ACABADO ENVIADO COM SUCESSO!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_industrializado(self, lista_produtos):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'PCP - Produtos para Industrializar!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos para Industrializar:\n\n"

            for i in lista_produtos:
                cod, descr, ref, um, necessidade = i

                body += f"- Código: {cod}\n" \
                        f"- Descrição: {descr}\n UM: {um}" \
                        f"- Referência: {ref}\n" \
                        f"- Qtde Necessidade: {necessidade}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print(f'EMAIL INDUSTRIALIZADO ENVIADO COM SUCESSO!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_comprado(self, lista_produtos):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'PCP - Produtos para Comprar!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos para Comprar:\n\n"

            for i in lista_produtos:
                cod, descr, ref, um, necessidade = i

                body += f"- Código: {cod}\n" \
                        f"- Descrição: {descr}\n UM: {um}" \
                        f"- Referência: {ref}\n" \
                        f"- Qtde Necessidade: {necessidade}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print(f'EMAIL COMPRADO ENVIADO COM SUCESSO!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = OPSeOCSFaltando()