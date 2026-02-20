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
from dados_email import email_user, password


class EnviaCadastroProduto:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        self.diretorio_script = os.path.dirname(nome_arquivo_com_caminho)
        nome_base = os.path.splitext(self.nome_arquivo)[0]
        self.arquivo_log = os.path.join(self.diretorio_script, f"{nome_base}_erros.txt")

    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

            # 'Log' em arquivo local apenas se houver erro
            with open(self.arquivo_log, "a", encoding="utf-8") as f:
                f.write(f"Erro na função {nome_funcao} do arquivo {arquivo}: {mensagem} (linha {num_linha_erro})\n")

        except Exception as e:
            nome_funcao_trat = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            tb = traceback.extract_tb(exc_traceback)
            num_linha_erro = tb[-1][1]
            print(f'Houve um problema no arquivo: {self.nome_arquivo} na função: "{nome_funcao_trat}"\n'
                  f'{e} {num_linha_erro}')
            grava_erro_banco(nome_funcao_trat, e, self.nome_arquivo, num_linha_erro)

            with open(self.arquivo_log, "a", encoding="utf-8") as f:
                f.write(
                    f"Erro na função {nome_funcao_trat} do arquivo {self.nome_arquivo}: {e} (linha {num_linha_erro})\n")

    def mensagem_email(self):
        try:
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

            msg_final = f"Att,\n" \
                        f"Suzuki Máquinas Ltda\n" \
                        f"Fone (51) 3561.2583/(51) 3170.0965\n\n" \
                        f"Mensagem enviada automaticamente, por favor não responda.\n\n" \
                        f"Se houver algum problema com o recebimento de emails ou conflitos com o arquivo excel, " \
                        f"favor entrar em contato pelo email maquinas@unisold.com.br.\n\n"

            to = ['<maquinas@unisold.com.br>', '<ahcmaquinas@gmail.com>']

            return saudacao, msg_final, to

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email(self, num_reg, lista_produtos):
        try:
            saudacao, msg_final, to = self.mensagem_email()

            subject = f'PRO - Cadastro de Produtos Registro {num_reg}!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de produtos cadastrados:\n\n" \

            body += f"{lista_produtos}\n{msg_final}"

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

    def manipula_dados_prod(self):
        try:
            tabela = []
            registros_agrupados = {}

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, registro, obs, descricao, descr_compl, referencia, um, ncm, "
                           f"kg_mt, fornecedor, data_criacao, codigo "
                           f"FROM PRODUTOPRELIMINAR "
                           f"WHERE (codigo IS NOT NULL OR codigo <> 0) AND (entregue IS NULL OR entregue = '' OR entregue = 'N');")
            dados_banco = cursor.fetchall()

            if dados_banco:
                for i in dados_banco:
                    print(i)
                    id_pre, registro, obs, descr, compl, ref, um, ncm, kg_mt, forn, emissao, codigo = i

                    datis = emissao.strftime("%d/%m/%Y")

                    dados = (datis, registro, obs, codigo, descr, compl, ref, um, ncm, kg_mt, forn)
                    tabela.append(dados)

                    if registro in registros_agrupados:
                        registros_agrupados[registro].append(dados)
                    else:
                        registros_agrupados[registro] = [dados]

            for registro, dados_grupo in registros_agrupados.items():
                string_inicial = ""
                for dadoss in dados_grupo:
                    emi_c, reg_c, obs_c, cod_c, descr_c, compl_c, ref_c, um_c, ncm_c, kg_mt_c, forn_c = dadoss

                    if not string_inicial:
                        string_inicial = (
                            f"Data: {emi_c} - Nº Registro: {reg_c} - Observação: {obs_c}:\n\n"
                        )

                    string_inicial += f"- Cód: {cod_c}\n" \
                                      f"- Descrição: {descr_c}\n" \
                                      f"- Referência: {ref_c}\n" \
                                      f"- UM: {um_c}\n" \
                                      f"- Fornecedor: {forn_c}\n\n"

                self.envia_email(registro, string_inicial)

                for dados1 in dados_grupo:
                    reg_c1 = dados1[1]
                    cod_c1 = dados1[3]

                    self.update_pre_cadastro(reg_c1, cod_c1)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def update_pre_cadastro(self, num_registro, cod_produto):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"UPDATE PRODUTOPRELIMINAR SET ENTREGUE = 'S' "
                           f"WHERE registro = {num_registro} and codigo = {cod_produto};")

            print(f"Registro {num_registro} e Código {cod_produto} enviado com sucesso!")

            conecta.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_registros_excluir(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, registro, obs, descricao, descr_compl, referencia, um, ncm, "
                           f"kg_mt, fornecedor, data_criacao, codigo "
                           f"FROM PRODUTOPRELIMINAR "
                           f"WHERE (codigo IS NOT NULL OR codigo <> 0) AND entregue = 'S';")
            dados_banco = cursor.fetchall()
            if dados_banco:
                for i in dados_banco:
                    id_pre, registro, obs, descr, compl, ref, um, ncm, kg_mt, forn, emissao, codigo = i

                    cursor = conecta.cursor()
                    cursor.execute(f"DELETE FROM PRODUTOPRELIMINAR WHERE ID = {id_pre};")

                    print(f"Id Pré {id_pre} excluído com sucesso!")

                    conecta.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaCadastroProduto()
chama_classe.verifica_registros_excluir()
chama_classe.manipula_dados_prod()
