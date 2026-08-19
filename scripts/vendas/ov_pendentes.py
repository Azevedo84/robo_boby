import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
from core.erros import trata_excecao
from core.email_service import dados_email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, date, timedelta


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


class EnviaOrdensVendaPendentes:
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


            subject = f'OV - Ordem de Venda Nº {num_oc} - {primeino_nome} não foi entregue!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}<br><br>" \
                   f"A ordem de Venda Nº <b>{num_oc}</b> não foi entregue.<br><br>" \
                   f"<b>Cliente:</b> {fornec_edi}<br>" \
                   f"<b>Data de Emissão:</b> {data_formatada}<br><br>" \
                   f"<b>Produtos que faltam ser entregues:</b><br><br>"

            todas_produtos_ocs = dados_banco.consultar_1_cond('tab_produtos_ordem', 'Número OC', num_oc)
            for i in todas_produtos_ocs:
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

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_dados(self):
        try:
            data_hoje = date.today()
            data_necessidade = data_hoje - timedelta(days=2)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.id, oc.data, oc.numero, cli.razao, prodoc.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ''), "
                           f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega "
                           f"FROM ordemcompra as oc "
                           f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"where oc.entradasaida = 'S' AND oc.STATUS = 'A' AND prodoc.produzido < prodoc.quantidade "
                           f"ORDER BY prodoc.dataentrega;")
            dados_oc = cursor.fetchall()

            for dados in dados_oc:
                print(dados)
                id_oc, data, numero, fornecedor, codigo, descricao, ref, um, qtde, produzido, data_entrega = dados
                falta = "%.3f" % (float(qtde) - float(produzido))

                if data < data_necessidade:
                    todas_oc = bc_ander.consultar_1_cond('tab_ordens', 'Número OC', numero)

                    if not todas_oc:
                        bc_ander.inserir('tab_ordens', {'Número OC': numero,
                                                        'Data Emissão': data,
                                                        'Fornecedor': fornecedor})

                    todas_produtos_oc = bc_ander.consultar_2_cond('tab_produtos_ordem',
                                                                  'Número OC', numero,
                                                                  'Código Produto', codigo)
                    if not todas_produtos_oc:
                        bc_ander.inserir('tab_produtos_ordem', {'Número OC': numero,
                                                                'Código Produto': codigo,
                                                                'Qtde Falta': falta})

            dados_user = bc_ander.consultar('tab_ordens')
            for didos in dados_user:
                print(didos)
                num_orc = didos['Número OC']
                data_entr = didos['Data Emissão']
                forne = didos['Fornecedor']

                self.envia_email(num_orc, data_entr, forne, bc_ander)

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = EnviaOrdensVendaPendentes()
chama_classe.manipula_dados()
