from core.banco import conecta, conecta_engenharia
from core.erros import grava_erro_banco
import re
import sys
import os
import traceback
from dados_email import email_user, password
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import datetime


class GerarFilaValidacaoERP:
    def __init__(self):
        print("🚀 Gerador fila validação ERP iniciado")
        self.nome_arquivo = os.path.basename(__file__)

        self.processar()

    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

        except Exception as e:
            nome_funcao_trat = sys._getframe().f_code.co_name
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

            msg_final = ""

            msg_final += f"Att,\n"
            msg_final += f"Suzuki Máquinas Ltda\n"
            msg_final += f"Fone (51) 3561.2583/(51) 3170.0965\n\n"
            msg_final += f"🟦 Mensagem gerada automaticamente pelo sistema de Planejamento e Controle da Produção (PCP) do ERP Suzuki.\n"
            msg_final += "🔸Por favor, não responda este e-mail diretamente."


            return saudacao, msg_final, to

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def caminho_para_link(self, caminho):
        try:
            return "file:///" + caminho.replace("\\", "/")

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_problemas(self, texto):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'ENGENHARIA VINCULO ERP - PROBLEMA ENCONTRADOS'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"<p>{saudacao}</p>"
            body += "<p>Foram encontrados problemas com propriedades dos projetos!</p>"

            for item in texto:
                link = self.caminho_para_link(item["caminho"])

                body += f"""
                <p>
                <b>Código PI: {item["Código Pai"]} - {item["Referência"]}</b><br>
                <b>{item["id"]} - {item["nome"]}</b><br>
                <a href="{link}">{item["caminho"]}</a><br>
                {item["erro"]}
                </p>
                """

            body += f"<br><p>{msg_final}</p>"

            msg.attach(MIMEText(body, 'html'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado com problemas")

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def processar(self):
        try:
            lista_itens = set()

            cursor_erp = conecta.cursor()
            cursor_eng = conecta_engenharia.cursor()

            cursor_erp.execute("""
                SELECT prod.codigo, prod.obs
                FROM PRODUTOPEDIDOINTERNO prodint
                JOIN produto prod ON prodint.id_produto = prod.id
                WHERE prodint.status = 'A' AND prod.descricao NOT LIKE '%KIT%'
            """)
            registros = cursor_erp.fetchall()

            print(f"📦 Total pedidos ativos: {len(registros)}")

            for codigo, obs in registros:
                ref = self.extrair_referencia(obs)

                if not ref:
                    print(f"⚠️ Produto {codigo} sem referência válida")
                    continue

                cursor_eng.execute("""
                    SELECT ID, TIPO_ARQUIVO
                    FROM ARQUIVOS
                    WHERE NOME_BASE = ?
                      AND TIPO_ARQUIVO IN ('IPT', 'IAM')
                """, (ref,))

                resultados = cursor_eng.fetchall()

                if not resultados:
                    print(f"❌ Sem desenho: {ref}")
                    continue

                if len(resultados) > 1:
                    print(f"⚠️ Duplicado: {ref}")
                    continue

                id_arquivo, tipo = resultados[0]

                ids = self.buscar_toda_estrutura(cursor_eng, id_arquivo)

                if ids:
                    for id_item in ids:
                        dados = (codigo, obs, id_item)
                        lista_itens.add(dados)

            if lista_itens:
                self.verificar_propriedades(lista_itens)

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def ignorar_arquivos(self):
        try:
            lista_ignorados = []

            return lista_ignorados

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def inserir_na_fila(self, id_arquivo):
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                            SELECT 1 FROM FILA_CONFERENCIA
                            WHERE ID_ARQUIVO = ?
                        """, (id_arquivo,))

            if cursor.fetchone():
                print("⚠️ Já está na fila:", id_arquivo)
                return

            # 🔹 insere o próprio arquivo
            cursor.execute("""
                            INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO, ORIGEM)
                            VALUES (?, ?)
                        """, (id_arquivo, "ALTERADOS"))

            print("📥 Inserido na fila:", id_arquivo)

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_estrutura_onde_usa(self, id_arquivo):
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                            SELECT estrut.ID_PAI, arq.ARQUIVO
                            FROM ESTRUTURA as estrut
                            INNER JOIN arquivos as arq ON estrut.ID_PAI = arq.id
                            WHERE estrut.ID_FILHO=?
                        """, (id_arquivo,))
            estrutura = cursor.fetchall()

            return estrutura or []

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_propriedades_iam(self, id_arquivo):
        try:
            cursor_eng = conecta_engenharia.cursor()
            cursor_eng.execute("""
                                SELECT AUTHORITY, DESCRIPTION 
                                FROM PROPRIEDADES_IAM
                                WHERE ID_ARQUIVO=?
                            """, (id_arquivo,))
            prop_iam = cursor_eng.fetchall()

            return prop_iam or []

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_propriedades_ipt(self, id_arquivo, classificacao):
        try:
            cursor_eng = conecta_engenharia.cursor()
            if classificacao == "TERCEIROS":
                cursor_eng.execute("""
                                SELECT AUTHORITY, DESCRIPTION 
                                FROM PROPRIEDADES_IPT 
                                WHERE ID_ARQUIVO=?
                            """, (id_arquivo,))
                prop_ipt = cursor_eng.fetchall()

            else:
                cursor_eng.execute("""
                                    SELECT AUTHORITY, DESCRIPTION, COST_CENTER, REVISION_NUMBER 
                                    FROM PROPRIEDADES_IPT 
                                    WHERE ID_ARQUIVO=?
                                """, (id_arquivo,))
                prop_ipt = cursor_eng.fetchall()

            return prop_ipt or []

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_arquivos(self, id_arquivo):
        try:
            cursor_eng = conecta_engenharia.cursor()
            cursor_eng.execute("""
                            SELECT ARQUIVO, TIPO_ARQUIVO, CLASSIFICACAO, caminho
                            FROM arquivos where ID = ?
                            """, (id_arquivo,))
            arquivo = cursor_eng.fetchall()

            return arquivo or []

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_erp(self, codigo):
        try:
            cursor_erp = conecta.cursor()
            cursor_erp.execute("""
                            SELECT id, descricao, COALESCE(obs, ' ') as obs, unidade, id_versao 
                            FROM produto where codigo = ?
                            """, (codigo,))
            produto = cursor_erp.fetchall()

            return produto or []

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verificar_propriedades(self, lista_itens):
        try:
            lista_final_email = []
            lista_ignorados = self.ignorar_arquivos() or []

            print(len(lista_itens))

            for codigo, ref, id_arquivo in lista_itens:
                if id_arquivo in lista_ignorados:
                    continue

                arquivo = self.consulta_arquivos(id_arquivo)

                if arquivo:
                    for nome, tipo, classificacao, caminho in arquivo:
                        msg_arquivo = ""

                        if tipo == "IAM":
                            prop_iam = self.consulta_propriedades_iam(id_arquivo)

                            soma_problema = 0

                            if prop_iam:
                                for cod_produto, descricao in prop_iam:
                                    if not cod_produto or cod_produto == " ":
                                        soma_problema += 1
                                        msg_arquivo += " Sem Código <br>"
                                    if not descricao:
                                        soma_problema += 1
                                        msg_arquivo += " Sem Descrição <br>"
                                    if cod_produto and descricao:
                                        if cod_produto != " ":
                                            dados_produto = self.consulta_erp(cod_produto)
                                            if dados_produto:
                                                if descricao != dados_produto[0][1]:
                                                    print("DESCRIÇÃO ERRADA!", cod_produto, descricao, dados_produto, caminho)
                                            else:
                                                print("CODIGO ERRADO!", cod_produto, len(cod_produto), descricao, dados_produto, caminho)

                                if soma_problema > 0:
                                    estrutura = self.consulta_estrutura_onde_usa(id_arquivo)
                                    if estrutura:
                                        for estrut in estrutura:
                                            msg_arquivo += f"Onde é usado: {estrut}<br>"

                            else:
                                print("sem propriedades IAM CADASTRADOS", nome, tipo, classificacao, caminho)

                                self.inserir_na_fila(id_arquivo)

                                estrutura = self.consulta_estrutura_onde_usa(id_arquivo)
                                if estrutura:
                                    for estrut in estrutura:
                                        print("ESTRUTURA:" ,estrut)

                        elif tipo == "IPT":
                            if classificacao == "TERCEIROS":
                                prop_ipt = self.consulta_propriedades_ipt(id_arquivo, classificacao)

                                soma_problema = 0

                                if prop_ipt:
                                    for cod_produto, descricao in prop_ipt:
                                        if "BORRACHA IND." in descricao:
                                            continue
                                        if not cod_produto:
                                            soma_problema += 1
                                            msg_arquivo += " Sem Código <br>"
                                        if not descricao:
                                            soma_problema += 1
                                            msg_arquivo += " Sem Descrição <br>"

                                    if soma_problema > 0:
                                        estrutura = self.consulta_estrutura_onde_usa(id_arquivo)
                                        if estrutura:
                                            for estrut in estrutura:
                                                msg_arquivo += f"Onde é usado: {estrut}<br>"

                            else:
                                prop_ipt = self.consulta_propriedades_ipt(id_arquivo, classificacao)

                                soma_problema = 0

                                if prop_ipt:
                                    for cod_produto, descricao, cod_mat, descr_mat in prop_ipt:
                                        if not cod_produto:
                                            soma_problema += 1
                                            msg_arquivo += " Sem Código <br>"
                                        if not descricao:
                                            soma_problema += 1
                                            msg_arquivo += " Sem Descrição <br>"
                                        if not cod_mat:
                                            if "chapa" in descr_mat or "CHAPA" in descr_mat:
                                                pass
                                            elif "fundido" in descr_mat or "FUNDIDO" in descr_mat:
                                                pass
                                            else:
                                                soma_problema += 1
                                                msg_arquivo += " Sem Código de matéria-prima <br>"
                                        if not descr_mat:
                                            soma_problema += 1
                                            msg_arquivo += " Sem Descrição de materia-prima <br>"

                                else:
                                    self.inserir_na_fila(id_arquivo)
                                    print("sem propriedades IPT CADASTRADOS", nome, tipo, classificacao, caminho)

                                if soma_problema > 0:
                                    estrutura = self.consulta_estrutura_onde_usa(id_arquivo)
                                    if estrutura:
                                        for estrut in estrutura:
                                            msg_arquivo += f"Onde é usado: {estrut}<br>"

                        else:
                            self.inserir_na_fila(id_arquivo)
                            print("ENTROU NO ELSE!!!", id_arquivo)

                        if msg_arquivo:
                            lista_final_email.append({
                                "Código Pai": codigo,
                                "Referência": ref,
                                "id": id_arquivo,
                                "nome": nome,
                                "tipo": tipo,
                                "classificacao": classificacao,
                                "caminho": caminho,
                                "erro": msg_arquivo,
                            })
            conecta_engenharia.commit()
            if lista_final_email:
                pass
                #self.envia_email_problemas(lista_final_email)

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def buscar_toda_estrutura(self, cursor, id_pai):
        try:
            visitados = set()
            fila = [id_pai]

            while fila:
                atual = fila.pop()

                if atual in visitados:
                    continue

                visitados.add(atual)

                # 🔴 NOVO: verifica classificação
                cursor.execute("""
                    SELECT CLASSIFICACAO
                    FROM ARQUIVOS
                    WHERE ID = ?
                """, (atual,))
                row = cursor.fetchone()

                if row:
                    classificacao = row[0]

                    if classificacao == "TERCEIROS":
                        continue

                # 🔽 Só chega aqui se NÃO for terceiros
                cursor.execute("""
                    SELECT ID_FILHO
                    FROM ESTRUTURA
                    WHERE ID_PAI = ?
                """, (atual,))

                filhos = [row[0] for row in cursor.fetchall()]

                fila.extend(filhos)

            return list(visitados)

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def extrair_referencia(self, obs):
        try:
            if not obs:
                return None

            s = re.sub(r"[^\d.]", "", obs)
            s = re.sub(r"\.+$", "", s)

            return s if s else None

        except Exception as e:
            nome_funcao = sys._getframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

if __name__ == "__main__":
    GerarFilaValidacaoERP()