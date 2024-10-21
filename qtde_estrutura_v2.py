import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import traceback


class LancaItensEstrutura:
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

    def envia_email(self, dados):
        try:
            saudacao, msg_final, email_user, to, password = self.dados_email()

            subject = f'SEM TIPO - PRODUTOS SEM TIPO DEFINIDOS (PI)!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos sem tipo definido:\n\n"

            for i in dados:
                body += f"- {i}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_pi(self):
        try:
            nova_lista_sem_tipo = []

            timestamp_atual = datetime.now()
            timestamp_formatado = timestamp_atual.strftime('%Y-%m-%d %H:%M:%S')

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prodint.id_pedidointerno, prod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodint.qtde "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    num_pi, id_prod, cod, descr, ref, um, qtde = i
                    print("PRODUTOS DO PI:", i)

                    estrutura = self.calculo_3_verifica_estrutura(cod, qtde)

                    if estrutura:
                        qtde_itens = len(estrutura)

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT * FROM RESUMO_ESTRUTURA where produto = {id_prod};")
                        detalhes_resumo = cursor.fetchall()

                        if detalhes_resumo:
                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT * FROM RESUMO_ESTRUTURA "
                                           f"where produto = {id_prod} "
                                           f"and qtde_itens = '{qtde_itens}';")
                            detalhes_resumo1 = cursor.fetchall()

                            if not detalhes_resumo1:
                                cursor = conecta.cursor()
                                cursor.execute(f"UPDATE RESUMO_ESTRUTURA "
                                               f"SET QTDE_ITENS = '{qtde_itens}', "
                                               f"DATA_ATUALIZACAO = '{timestamp_formatado}' "
                                               f"WHERE produto = {id_prod};")

                                conecta.commit()

                                print("ATUALIZADO", cod, descr, ref, um, qtde_itens)

                        else:
                            cursor = conecta.cursor()
                            cursor.execute(f"Insert into RESUMO_ESTRUTURA (ID, produto, qtde_itens, data_atualizacao) "
                                           f"values (GEN_ID(GEN_RESUMO_ESTRUTURA_ID,1), {id_prod}, '{qtde_itens}', "
                                           f"'{timestamp_formatado}');")

                            conecta.commit()

                            print("NOVO", cod, descr, ref, um, qtde_itens)

                        nova_lista_sem_tipo = self.manipula_estrutura(id_prod, estrutura, nova_lista_sem_tipo)

            if nova_lista_sem_tipo:
                self.envia_email(nova_lista_sem_tipo)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_3_verifica_estrutura(self, codigo, qtde):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prod.quantidade, tip.id, tip.tipomaterial, prod.conjunto, prod.id_versao "
                           f"FROM produto as prod "
                           f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                           f"where prod.codigo = {codigo};")
            detalhes_pai = cursor.fetchall()
            id_pai, cod_pai, descr_pai, ref_pai, um_pai, saldo, id_tipo, tipo, id_conj, id_estrut = detalhes_pai[0]

            dadoss = (cod_pai, descr_pai, ref_pai, um_pai, qtde, id_tipo, tipo, id_conj)

            filhos = [dadoss]

            if id_estrut:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                               f"(estprod.quantidade * {qtde}) as qtde "
                               f"FROM estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"where estprod.id_estrutura = {id_estrut};")
                dados_estrutura = cursor.fetchall()

                if dados_estrutura:
                    for prod in dados_estrutura:
                        cod_f, descr_f, ref_f, um_f, qtde_f = prod

                        filhos.extend(self.calculo_3_verifica_estrutura(cod_f, qtde_f))

            return filhos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_estrutura(self, id_prod, dados_estrutura, nova_lista_sem_tipo):
        try:
            contagem_tipos = {}

            qtde_conjunto = 0
            qtde_materiaprima = 0

            for i in dados_estrutura:
                cod_pai, descr_pai, ref_pai, um_pai, qtde, id_tipo, tipo, id_conj = i

                if not tipo:
                    nova_lista_sem_tipo.append(i)

                if id_tipo in contagem_tipos:
                    contagem_tipos[id_tipo] += 1
                else:
                    contagem_tipos[id_tipo] = 1

                if id_conj == 10:
                    qtde_conjunto += 1
                else:
                    qtde_materiaprima += 1

            timestamp_atual = datetime.now()
            timestamp_formatado = timestamp_atual.strftime('%Y-%m-%d %H:%M:%S')

            self.manipula_tipo(id_prod, contagem_tipos, timestamp_formatado)
            self.manipula_qtde_conj_mat(id_prod, qtde_conjunto, qtde_materiaprima, timestamp_formatado)

            return nova_lista_sem_tipo

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_tipo(self, id_prod, contagem_tipos, timestamp):
        try:
            tipos_na_lista = set(contagem_tipos.keys())

            cursor = conecta.cursor()
            cursor.execute(f"SELECT tipomaterial FROM RESUMO_ESTRUTURA_TIPO WHERE id_produto = {id_prod};")
            tipos_existentes = cursor.fetchall()

            tipos_existentes_set = set(t[0] for t in tipos_existentes)

            tipos_para_deletar = tipos_existentes_set - tipos_na_lista

            if tipos_para_deletar:
                for tipo in tipos_para_deletar:
                    cursor.execute(
                        f"DELETE FROM RESUMO_ESTRUTURA_TIPO WHERE id_produto = {id_prod} AND tipomaterial = '{tipo}';")
                conecta.commit()
                print("DELETADOS TIPOS", id_prod, tipos_para_deletar)

            for did in contagem_tipos.items():
                chave, qtde_tip = did

                if chave is None:
                    novo_tipo = "9999"
                else:
                    novo_tipo = chave

                cursor = conecta.cursor()
                cursor.execute(f"SELECT * FROM RESUMO_ESTRUTURA_TIPO "
                               f"where id_produto = {id_prod} "
                               f"and tipomaterial = '{novo_tipo}';")
                detalhes_resumo = cursor.fetchall()

                if detalhes_resumo:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT * FROM RESUMO_ESTRUTURA_TIPO "
                                   f"where id_produto = {id_prod} "
                                   f"and tipomaterial = '{novo_tipo}' "
                                   f"and qtde_itens = '{qtde_tip}';")
                    detalhes_resumo1 = cursor.fetchall()

                    if not detalhes_resumo1:
                        cursor = conecta.cursor()
                        cursor.execute(f"UPDATE RESUMO_ESTRUTURA_TIPO "
                                       f"SET QTDE_ITENS = '{qtde_tip}', "
                                       f"tipomaterial = '{novo_tipo}', "
                                       f"DATA_ATUALIZACAO = '{timestamp}' "
                                       f"WHERE id_produto = {id_prod} and tipomaterial = '{novo_tipo}';")

                        conecta.commit()

                        print("ATUALIZADO TIPOS", id_prod, novo_tipo, qtde_tip)

                else:
                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into RESUMO_ESTRUTURA_TIPO (ID, id_produto, tipomaterial, qtde_itens, "
                                   f"data_atualizacao) "
                                   f"values (GEN_ID(GEN_RESUMO_ESTRUTURA_TIPO_ID,1), {id_prod}, '{novo_tipo}', "
                                   f"'{qtde_tip}', "
                                   f"'{timestamp}');")

                    conecta.commit()

                    print("NOVO TIPOS", id_prod, novo_tipo, qtde_tip)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_qtde_conj_mat(self, id_prod, qtde_conj, qtde_mat, timestamp):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT * FROM RESUMO_ESTRUTURA where produto = {id_prod};")
            detalhes_resumo = cursor.fetchall()

            if detalhes_resumo:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT * FROM RESUMO_ESTRUTURA "
                               f"where produto = {id_prod} "
                               f"and qtde_acabado = '{qtde_conj}' "
                               f"and qtde_materiaprima = '{qtde_mat}';")
                detalhes_resumo1 = cursor.fetchall()

                if not detalhes_resumo1:
                    cursor = conecta.cursor()
                    cursor.execute(f"UPDATE RESUMO_ESTRUTURA "
                                   f"SET qtde_acabado = '{qtde_conj}', "
                                   f"qtde_materiaprima = '{qtde_mat}', "
                                   f"DATA_ATUALIZACAO = '{timestamp}' "
                                   f"WHERE produto = {id_prod};")

                    conecta.commit()

                    print("ACABADO ATUALIZADO", id_prod)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = LancaItensEstrutura()
chama_classe.manipula_dados_pi()
