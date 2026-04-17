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
from datetime import datetime, date, timedelta
from core.erros import trata_excecao
from core.email_service import dados_email


class BancoAnder:
    def __init__(self):
        self.dados = {}

    def inserir(self, tabela, dados):
        try:
            if tabela not in self.dados:
                self.dados[tabela] = []
            self.dados[tabela].append(dados)

        except Exception as e:
            trata_excecao(e)
            raise

    def consultar(self, tabela):
        try:
            return self.dados.get(tabela, [])

        except Exception as e:
            trata_excecao(e)
            raise

    def consultar_1_cond(self, tabela, campo, condicao):
        try:
            return [dado for dado in self.dados.get(tabela, []) if dado.get(campo) == condicao]

        except Exception as e:
            trata_excecao(e)
            raise

    def consultar_2_cond(self, tabela, campo1, condicao1, campo2, condicao2):
        try:
            return [dado for dado in self.dados.get(tabela, []) if
                    dado.get(campo1) == condicao1 and dado.get(campo2) == condicao2]

        except Exception as e:
            trata_excecao(e)
            raise

    def consultar_3_cond(self, tabela, campo1, condicao1, campo2, condicao2, campo3, condicao3):
        try:
            return [dado for dado in self.dados.get(tabela, []) if
                    dado.get(campo1) == condicao1 and
                    dado.get(campo2) == condicao2 and
                    dado.get(campo3) == condicao3]

        except Exception as e:
            trata_excecao(e)
            raise


bc_ander = BancoAnder()


class EnviaOrdensCompraPendentes:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

    def envia_email(self, num_oc, data_entrego, fornecedore, dados_banco):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            insert_tabela = []
            fornec_edi = fornecedore.capitalize()
            palavras = fornec_edi.split()
            primeino_nome = palavras[0]

            data_entrego_str = data_entrego.strftime("%Y-%m-%d")
            datati = datetime.strptime(data_entrego_str, "%Y-%m-%d")
            data_formatada = datati.strftime("%d/%m/%Y")

            subject = f'OC - Ordem de Compra Nº {num_oc} - {primeino_nome} não foi entregue!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}<br><br>" \
                   f"A ordem de compra Nº <b>{num_oc}</b> não foi entregue.<br><br>" \
                   f"<b>Fornecedor:</b> {fornec_edi}<br>" \
                   f"<b>Previsão de Entrega:</b> {data_formatada}<br><br>" \
                   f"<b>Produtos que faltam ser entregues:</b><br><br>"

            todas_produtos_oc = dados_banco.consultar_1_cond('tab_produtos_ordem', 'Número OC', num_oc)
            for i in todas_produtos_oc:
                codigo_produtos = i['Código Produto']

                qtdes_produtos = i['Qtde Falta']
                qtde_virg = qtdes_produtos.replace('.', ',')
                inicio = qtde_virg.find(",") + 1
                final_qtde = qtde_virg[inicio:]
                if final_qtde == "000":
                    fim = qtde_virg.find(",")
                    qtde_ajustado = qtde_virg[:fim]
                else:
                    qtde_ajustado = qtde_virg

                curs = conecta.cursor()
                curs.execute(f"SELECT codigo, descricao, COALESCE(obs, ''), unidade "
                             f"FROM produto where codigo = {codigo_produtos};")
                detalhes_produtos = curs.fetchall()
                codigo_p, descricao_p, referencia_p, um_p = detalhes_produtos[0]

                titi = (num_oc, codigo_p, data_entrego)
                insert_tabela.append(titi)

                body += f"- <b>Código:</b> {codigo_p} - <b>Descrição:</b> {descricao_p} " \
                        f"- <b>Referência:</b> {referencia_p} - <b>Quantidade:</b> {qtde_ajustado} {um_p}<br>"

                todas_ops = dados_banco.consultar_1_cond('tab_ops_abertas', 'Código Produto Filho', codigo_p)
                if todas_ops:
                    for ii in todas_ops:
                        num_opii = ii['Número OP']
                        desc_opii = ii['Descrição Produto']

                        body += f"            A OP Nº <b>{num_opii}</b> - <b>{desc_opii}</b> está " \
                                f"aguardando este material para finalizar o serviço!<br>"

                body += "<br>"

            body += "<br><br>"

            body += f"Att,<br><br>" \
                    f"Suzuki Máquinas Ltda<br>" \
                    f"Fone (51) 3561.2583/(51) 3170.0965<br>" \
                    f"Mensagem enviada automaticamente, por favor não responda.<br>" \
                    f"Se houver algum problema com o recebimento de emails favor entrar em contato " \
                    f"pelo email maquinas@unisold.com.br."

            msg.attach(MIMEText(body, 'html'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)

            server.quit()

            if insert_tabela:
                for tititi in insert_tabela:
                    ins_num_oc, ins_codigo, ins_data_entrega = tititi

                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into envia_oc_entregue (ID, numero_oc, codigo_produto, data_entrega) "
                                   f"values (GEN_ID(GEN_ENVIA_OC_ENTREGUE_ID,1), {ins_num_oc}, {ins_codigo}, "
                                   f"'{ins_data_entrega}');")

            conecta.commit()

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_dados(self):
        try:
            data_hoje = date.today()
            data_necessidade = data_hoje - timedelta(days=3)

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.id, oc.data, oc.numero, forn.razao, prodoc.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where oc.entradasaida = 'E' AND oc.STATUS = 'A' AND prodoc.produzido < prodoc.quantidade "
                f"ORDER BY prodoc.dataentrega;")
            dados_oc = cursor.fetchall()

            for dados in dados_oc:
                id_oc, data, numero, fornecedor, codigo, descricao, ref, um, qtde, produzido, data_entrega = dados
                falta = "%.3f" % (float(qtde) - float(produzido))

                cursor = conecta.cursor()
                cursor.execute(f"SELECT * from envia_oc_entregue "
                               f"where numero_oc = {numero} "
                               f"AND codigo_produto = {codigo} "
                               f"AND data_entrega = '{data_entrega}';")
                ja_enviei = cursor.fetchall()

                if data_entrega < data_necessidade and not ja_enviei:
                    todas_oc = bc_ander.consultar_1_cond('tab_ordens', 'Número OC', numero)

                    if not todas_oc:
                        bc_ander.inserir('tab_ordens', {'Número OC': numero,
                                                        'Data Emissão': data,
                                                        'Data Entrega': data_entrega,
                                                        'Fornecedor': fornecedor})

                    todas_produtos_oc = bc_ander.consultar_2_cond('tab_produtos_ordem',
                                                                  'Número OC', numero,
                                                                  'Código Produto', codigo)
                    if not todas_produtos_oc:
                        bc_ander.inserir('tab_produtos_ordem', {'Número OC': numero,
                                                                'Código Produto': codigo,
                                                                'Qtde Falta': falta})

                    cursor = conecta.cursor()
                    cursor.execute(
                        f"SELECT id, mestre, quantidade, produto, codigo from materiaprima where codigo = {codigo};")
                    estrut = cursor.fetchall()
                    if estrut:
                        for itens in estrut:
                            ides, cod_pai, qtde_filho, id_filho, cod_filho = itens

                            cur = conecta.cursor()
                            cur.execute(
                                f"SELECT codigo, descricao, COALESCE(obs, ''), unidade "
                                f"FROM produto "
                                f"where id = {cod_pai};")
                            detalhes_produtos = cur.fetchall()

                            for detalhe in detalhes_produtos:
                                codigo_pai, descricao_pai, referencia_pai, um_pai = detalhe

                                todos_pai = bc_ander.consultar_2_cond('tab_pais', 'Número OC', numero, 'Código Pai',
                                                                      codigo_pai)
                                if not todos_pai:
                                    bc_ander.inserir('tab_pais', {'Número OC': numero,
                                                                  'Código Produto': codigo,
                                                                  'Código Pai': codigo_pai,
                                                                  'Descrição Pai': descricao_pai,
                                                                  'Referência Pai': referencia_pai,
                                                                  'UM Pai': um_pai})

                                cursor = conecta.cursor()
                                cursor.execute(
                                    f"select ordser.datainicial, ordser.numero, prod.codigo, prod.descricao, "
                                    f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                                    f"ordser.quantidade, COALESCE(ordser.obs, '') "
                                    f"from ordemservico as ordser "
                                    f"INNER JOIN produto prod ON ordser.produto = prod.id "
                                    f"where ordser.status = 'A' AND prod.codigo = {codigo_pai};")
                                op_abertas = cursor.fetchall()
                                if op_abertas:
                                    for ops in op_abertas:
                                        data_op, num_op, codigo_op, descr_op, ref_op, op_um, op_qtde, op_obs = ops

                                        todos_ops = bc_ander.consultar_2_cond('tab_ops_abertas',
                                                                              'Número OC', numero,
                                                                              'Número OP', num_op)
                                        if not todos_ops:
                                            bc_ander.inserir('tab_ops_abertas', {'Número OC': numero,
                                                                                 'Código Pai': codigo_pai,
                                                                                 'Código Produto Filho': codigo,
                                                                                 'Número OP': num_op,
                                                                                 'Código Produto Pai': codigo_op,
                                                                                 'Descrição Produto': descr_op,
                                                                                 'Referência Produto': ref_op,
                                                                                 'Quantidade OP': op_qtde})

            dados_user = bc_ander.consultar('tab_ordens')
            for didos in dados_user:
                print(didos)
                num_orc = didos['Número OC']
                data_entr = didos['Data Entrega']
                forne = didos['Fornecedor']
                self.envia_email(num_orc, data_entr, forne, bc_ander)

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = EnviaOrdensCompraPendentes()
chama_classe.manipula_dados()
