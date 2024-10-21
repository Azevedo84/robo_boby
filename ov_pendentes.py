import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, date, timedelta
import traceback


class BancoAnder:
    def __init__(self):
        self.dados = {}

    def inserir(self, tabela, dados):
        if tabela not in self.dados:
            self.dados[tabela] = []
        self.dados[tabela].append(dados)

    def consultar(self, tabela):
        return self.dados.get(tabela, [])

    def consultar_1_cond(self, tabela, campo, condicao):
        return [dado for dado in self.dados.get(tabela, []) if dado.get(campo) == condicao]

    def consultar_2_cond(self, tabela, campo1, condicao1, campo2, condicao2):
        return [dado for dado in self.dados.get(tabela, []) if
                dado.get(campo1) == condicao1 and dado.get(campo2) == condicao2]

    def consultar_3_cond(self, tabela, campo1, condicao1, campo2, condicao2, campo3, condicao3):
        return [dado for dado in self.dados.get(tabela, []) if
                dado.get(campo1) == condicao1 and
                dado.get(campo2) == condicao2 and
                dado.get(campo3) == condicao3]


bc_ander = BancoAnder()


class EnviaOrdensVendaPendentes:
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

    def envia_email(self, num_oc, data_entrego, fornecedore, dados_banco):
        try:
            insert_tabela = []
            fornec_edi = fornecedore.capitalize()
            palavras = fornec_edi.split()
            primeino_nome = palavras[0]

            data_entrego_str = data_entrego.strftime("%Y-%m-%d")
            datati = datetime.strptime(data_entrego_str, "%Y-%m-%d")
            data_formatada = datati.strftime("%d/%m/%Y")

            to = ['<maquinas@unisold.com.br>']

            current_time = (datetime.now())
            horario = current_time.strftime('%H')
            hora_int = int(horario)
            saudacao = "teste"
            if 4 < hora_int < 13:
                saudacao = "Bom Dia!"
            elif 12 < hora_int < 19:
                saudacao = "Boa Tarde!"
            elif hora_int > 18:
                saudacao = "Boa Noite!"
            elif hora_int < 5:
                saudacao = "Boa Noite!"

            email_user = 'ti.ahcmaq@gmail.com'
            password = 'poswxhqkeaacblku'

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

            server.sendmail(email_user, to, text)

            server.quit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaOrdensVendaPendentes()
chama_classe.manipula_dados()
