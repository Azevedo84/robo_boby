import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from core.erros import trata_excecao
from core.email_service import dados_email


class LancaItensEstrutura:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

    def envia_email_sem_tipo(self, dados):
        try:
            saudacao, msg_final, email_user, password = dados_email()

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

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado")

        except Exception as e:
            trata_excecao(e)
            raise

    def envia_email_sem_desenho_pdf(self, dados):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'SEM DESENHO PDF - PRODUTOS SEM DESENHO PDF NO SERVIDOR!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos sem desenho PDF na pasta OP do servidor:\n\n"

            for i in dados:
                body += f"- {i}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado")

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_dados_pi(self):
        try:
            nova_lista_sem_tipo = []
            lista_sem_desenho_pdf = []

            timestamp_atual = datetime.now()
            timestamp_formatado = timestamp_atual.strftime('%Y-%m-%d %H:%M:%S')

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prodint.id_pedidointerno, prod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodint.qtde, prod.conjunto "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    num_pi, id_prod, cod, descr, ref, um, qtde, conj = i
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
                        lista_sem_desenho_pdf = self.verifica_desenhos_pdf(cod, descr, ref, conj, estrutura, lista_sem_desenho_pdf)

            if nova_lista_sem_tipo:
                self.envia_email_sem_tipo(nova_lista_sem_tipo)
            if lista_sem_desenho_pdf:
                self.envia_email_sem_desenho_pdf(lista_sem_desenho_pdf)

        except Exception as e:
            trata_excecao(e)
            raise

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
            trata_excecao(e)
            raise

    def verifica_desenhos_pdf(self, cod, descr, ref, id_conjunto, estrutura, lista_sem_desenho_pdf):
        try:
            for i in estrutura:
                cod_pai, descr_pai, ref_pai, um_pai, qtde, id_tipo, tipo, id_conj = i

                if id_conj == 10:
                    import re

                    s = re.sub(r"[^\d.]", "", ref_pai)  # remove tudo que não é número ou ponto
                    s = re.sub(r"\.+$", "", s)

                    caminho_pdf = rf"\\Publico\C\OP\Projetos\{s}.pdf"

                    if not os.path.exists(caminho_pdf):
                        dados = (cod_pai, descr_pai, ref_pai)
                        lista_sem_desenho_pdf.append(dados)

            if id_conjunto == 10:
                import re

                s = re.sub(r"[^\d.]", "", ref)  # remove tudo que não é número ou ponto
                s = re.sub(r"\.+$", "", s)

                caminho_pdf = rf"\\Publico\C\OP\Projetos\{s}.pdf"

                if not os.path.exists(caminho_pdf):
                    dados = (cod, descr, ref)
                    print("sem projeto", dados)
                    lista_sem_desenho_pdf.append(dados)

            return lista_sem_desenho_pdf

        except Exception as e:
            trata_excecao(e)
            raise

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
            trata_excecao(e)
            raise

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
            trata_excecao(e)
            raise

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
            trata_excecao(e)
            raise


chama_classe = LancaItensEstrutura()
chama_classe.manipula_dados_pi()
