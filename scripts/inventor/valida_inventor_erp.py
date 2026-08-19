import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta, conecta_engenharia
from core.erros import trata_excecao
from core.email_service import dados_email
from core.inventor import padrao_desenho, normalizar_texto
import re
from core.conversores import valores_para_float
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import date
import socket


class GerarFilaValidacaoERP:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

        self.processar()

    def enviar_email_referencia_erp_divergente(self, dados_duplicados, desenho, codigo, referencia):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA/ERP - REFERENCIA ERP DIVERGENTE DO DESENHO: {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f'{saudacao}\n\nNo código {codigo} a referência {referencia} está divergente do desenho {desenho}.\n\n'

            if dados_duplicados:
                for i in dados_duplicados:
                    id_arquivo, caminho, num_desenho, descr = i

                    body += f"{descr} - {num_desenho}: '{caminho}'\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado REFERENCIA ERP DIVERGENTE")

        except Exception as e:
            trata_excecao(e)
            raise

    def enviar_email_pre_cadastro(self, caminho, desenho, dados):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA/ERP - CRIAR PRE CADASTRO {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO desenho {desenho} não tem codigo e precisa de Pré Cadastro!\n\n"
            body += f"{caminho}\n\n"
            body += f"{dados}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado SEM codigo (pré cadastro)")

        except Exception as e:
            print(e)

    def enviar_email_cadastra_propriedade(self, caminho, desenho, dados):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA/ERP - LANÇAR PROPRIEDADE INVENTOR {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO desenho {desenho} precisa lançar proriedades!\n\n"
            body += f"{caminho}\n\n"
            body += f"{dados}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado cadastra propriedade no inventor")

        except Exception as e:
            print(e)

    def enviar_email_atualizar_estrutura(self, dados_estrutura):
        try:
            cod_prod, descricao, nome_base, id_estrutura, caminho, dados_op = dados_estrutura

            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA/ERP - ATUALIZAR ESTRUTURA {nome_base}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO produto Código: {cod_prod} - {descricao} está com estrutura diferente\n\n"

            if id_estrutura:
                body += f"Esta versão já existe pelo ID: {id_estrutura}\n\n"

            body += f"{caminho}\n\n"

            if dados_op:
                body += f"Existe Ordens de Produção abertas para este produto:\n\n"
                for i in dados_op:
                    body += f"{i}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado atualizar estrutura!")

        except Exception as e:
            print(e)

    def insert_divergencia(self, dados):
        try:
            id_divergencia, id_arquivo, obs, id_origem = dados

            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                            SELECT ID_TIPO_DIVERGENCIA, ID_ARQUIVO, OBS
                            FROM DIVERGENCIAS
                            WHERE ID_TIPO_DIVERGENCIA = ? and ID_ARQUIVO = ?
                        """, (id_divergencia, id_arquivo,))
            tem_divergencia = cursor.fetchall()

            if not tem_divergencia:
                sql = """
                        INSERT INTO DIVERGENCIAS (ID_TIPO_DIVERGENCIA, ID_ARQUIVO, OBS, ID_ORIGEM)
                        VALUES (?, ?, ?, ?);
                        """
                cursor.execute(sql, (id_divergencia, id_arquivo, obs, id_origem))

                conecta_engenharia.commit()

                print("\n")
                print("Divergencia Inserida com sucesso!", id_divergencia, id_arquivo, obs, id_origem)
                print("\n")

        except Exception as e:
            trata_excecao(e)
            raise

    def inserir_fila_conferencia(self, id_arquivo):
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                SELECT 1 FROM FILA_CONFERENCIA WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            if cursor.fetchone():
                return

            cursor.execute("""
                INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO, ORIGEM)
                VALUES (?, ?)
            """, (id_arquivo, "ALTERADOS"))

            print(f"INSERIDO DA FILA DE CONFERENCIA: {id_arquivo}")

            conecta_engenharia.commit()

        except Exception as e:
            trata_excecao(e)
            raise

    def insert_pre_cadastro(self, id_arquivo, descr, ref, id_origem):
        try:
            cursor = conecta.cursor()
            cursor.execute("""
                                        SELECT ID, OBS, DESCRICAO
                                        FROM PRODUTOPRELIMINAR
                                        WHERE REFERENCIA = ?
                                    """, (ref,))
            tem_preliminar = cursor.fetchall()

            if not tem_preliminar:
                dados = (24, id_arquivo, f"Descrição: {descr} - Referência: {ref}", id_origem)
                self.insert_divergencia(dados)

        except Exception as e:
            trata_excecao(e)
            raise

    def insert_propriedade_inventor(self, dados_produto, id_origem):
        try:
            id_arquivo, nome_prop, valor_prop, caminho_arquivo, nome_base = dados_produto

            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                            SELECT ID, ID_ARQUIVO, NOME_PROPRIEDADE
                            FROM FILA_LANCA_PROPRIEDADE
                            WHERE ID_ARQUIVO = ? and NOME_PROPRIEDADE = ?
                        """, (id_arquivo, nome_prop))
            tem_na_fila = cursor.fetchall()

            if not tem_na_fila:
                dados = (27, id_arquivo, f"{nome_prop} - {valor_prop}", id_origem)
                self.insert_divergencia(dados)

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_estrutura_eng(self, cursor, id_arquivo):
        try:
            cursor.execute("""
                SELECT ID_PAI, ID_FILHO
                FROM ESTRUTURA
                WHERE ID_PAI = ?
            """, (id_arquivo,))
            filhos = cursor.fetchall()

            return filhos

        except Exception as e:
            trata_excecao(e)
            raise

    def tratar_referencia(self, obs):
        try:
            if not obs:
                return None

            match = padrao_desenho.search(obs)

            if not match:
                return None

            s = re.sub(r"[^\d.]", "", obs)
            s = re.sub(r"\.+$", "", s)

            return s if s else None

        except Exception as e:
            trata_excecao(e)
            raise

    def buscar_toda_estrutura(self, cursor, id_pai, id_origem=None):
        try:
            if id_origem is None:
                id_origem = id_pai

            visitados = set()
            fila = [id_pai]

            itens = []

            while fila:
                atual = fila.pop()

                if atual in visitados:
                    continue

                visitados.add(atual)

                itens.append({
                    "id": atual,
                    "id_origem": id_origem
                })

                cursor.execute("""
                    SELECT CLASSIFICACAO
                    FROM ARQUIVOS
                    WHERE ID = ?
                """, (atual,))
                row = cursor.fetchone()

                if row and row[0] == "TERCEIROS":
                    continue

                cursor.execute("""
                    SELECT ID_FILHO
                    FROM ESTRUTURA
                    WHERE ID_PAI = ?
                """, (atual,))
                filhos = [r[0] for r in cursor.fetchall()]
                print(atual, filhos)
                fila.extend(filhos)

            return itens

        except Exception as e:
            trata_excecao(e)
            raise

    def processar(self):
        try:
            lista_itens = []
            ids_unicos = set()

            cursor_erp = conecta.cursor()
            cursor_eng = conecta_engenharia.cursor()

            cursor_erp.execute("""
                SELECT prod.codigo, prod.descricao, prod.obs, prod.conjunto 
                FROM PRODUTOPEDIDOINTERNO prodint
                JOIN produto prod ON prodint.id_produto = prod.id
                WHERE prodint.status = 'A' 
            """)
            registros = cursor_erp.fetchall()

            print(f"📦 Total pedidos ativos: {len(registros)}")

            for codigo, descricao, obs, conj in registros:
                ref = self.tratar_referencia(obs)

                if not ref:
                    continue

                cursor_eng.execute("""
                    SELECT ID, TIPO_ARQUIVO
                    FROM ARQUIVOS
                    WHERE NOME_BASE = ?
                      AND TIPO_ARQUIVO IN ('IPT', 'IAM')
                """, (ref,))
                resultados = cursor_eng.fetchall()

                if not resultados or len(resultados) > 1:
                    continue

                id_arquivo, tipo = resultados[0]

                ids = self.buscar_toda_estrutura(cursor_eng, id_arquivo)

                for item_estrutura in ids:
                    id_item = item_estrutura["id"]
                    id_origem = item_estrutura["id_origem"]

                    if id_item not in ids_unicos:
                        ids_unicos.add(id_item)

                        lista_itens.append({
                            "codigo": codigo,
                            "obs": obs,
                            "id": id_item,
                            "id_origem": id_origem
                        })

            cursor = conecta_engenharia.cursor()
            cursor.execute(f"SELECT proj.id, proj.id_arquivo, arq.nome_base "
                           f"FROM PROJETO as proj "
                           f"LEFT JOIN ARQUIVOS as arq ON proj.ID_ARQUIVO = arq.id "
                           f"where proj.status = 'A';")
            dados_projetos = cursor.fetchall()
            print("PROJETOS", dados_projetos)

            if dados_projetos:
                for i in dados_projetos:
                    id_projeto, id_arquivo_p, nome_base_p = i

                    if id_arquivo_p:
                        ids = self.buscar_toda_estrutura(cursor_eng, id_arquivo_p)

                        for item_estrutura in ids:
                            id_item = item_estrutura["id"]
                            id_origem = item_estrutura["id_origem"]

                            if id_item not in ids_unicos:
                                ids_unicos.add(id_item)

                                lista_itens.append({
                                    "codigo": "",
                                    "obs": nome_base_p,
                                    "id": id_item,
                                    "id_origem": id_origem
                                })

            if lista_itens:
                lista = self.montar_lista(lista_itens)

                # 🔥 debug / consulta
                self.tratar_resultado(cursor_eng, lista)

        except Exception as e:
            trata_excecao(e)
            raise

    def modelo_props(self):
        return {
            "codigo": None,
            "descricao": None,
            "cod_mat": None,
            "desc_mat": None,
            "num_desenho": None,
            "ncm": None,
            "tot_itens": None,
            "compr": None
        }

    def montar_lista(self, lista_itens):
        try:
            lista_final = []

            for dados in lista_itens:
                codigo = dados["codigo"]
                ref = dados["obs"]
                id_arquivo = dados["id"]
                id_origem = dados["id_origem"]
                dados_arq = self.montar_contexto_item(id_arquivo, codigo, ref)

                if not dados_arq:
                    continue

                props = self.buscar_props(dados_arq)

                # 🔥 começa com modelo vazio
                props_unificado = self.modelo_props()

                # 🔥 se tiver propriedades, mescla todas
                for p in props:
                    for k, v in p.items():
                        if v not in (None, "", 0):
                            props_unificado[k] = v

                # 🔥 junta tudo
                item = {
                    **dados_arq,
                    **props_unificado,
                    "id_origem": id_origem
                }

                lista_final.append(item)

            return lista_final

        except Exception as e:
            trata_excecao(e)
            raise

    def montar_contexto_item(self, id_arquivo, codigo, referencia):
        arquivo = self.consulta_arquivos(id_arquivo)

        if not arquivo:
            return None

        nome, nome_base, tipo, classificacao, caminho = arquivo[0]

        return {
            "id": id_arquivo,
            "codigo_pai": codigo,
            "referencia": referencia,
            "nome": nome,
            "nome_base": nome_base,
            "tipo": tipo,
            "classificacao": classificacao,
            "caminho": caminho
        }

    def consulta_arquivos(self, id_arquivo):
        cursor = conecta_engenharia.cursor()
        cursor.execute("""
            SELECT ARQUIVO, NOME_BASE, TIPO_ARQUIVO, CLASSIFICACAO, caminho
            FROM arquivos where ID = ?
        """, (id_arquivo,))
        return cursor.fetchall() or []

    def consulta_arquivos_idw(self, nome_base):
        cursor = conecta_engenharia.cursor()
        cursor.execute("""
            SELECT id, ARQUIVO, NOME_BASE, TIPO_ARQUIVO, CLASSIFICACAO, caminho
            FROM arquivos 
            where NOME_BASE = ? 
            AND TIPO_ARQUIVO = 'IDW'
        """, (nome_base,))
        return cursor.fetchall() or []

    def consulta_codigo_prod_erp(self, codigo):
        try:
            cursor_erp = conecta.cursor()
            cursor_erp.execute("""
                            SELECT prod.id, prod.descricao, COALESCE(prod.obs, ''), prod.unidade, 
                            prod.id_versao, prod.KILOSMETRO, prod.conjunto, tip.DESENHO, prod.ID_SERVICO_INTERNO  
                            FROM produto as prod 
                            LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id
                            where prod.codigo = ?
                            """, (codigo,))
            produto = cursor_erp.fetchall()

            return produto or []

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_referencia_prod_erp(self, ref):
        try:
            cursor_erp = conecta.cursor()
            cursor_erp.execute("""
                            SELECT id, codigo, descricao, COALESCE(obs, '') as obs, unidade, id_versao 
                            FROM produto where obs = ?
                            """, (ref,))
            produto = cursor_erp.fetchall()

            return produto or []

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_codigo_propriedade_eng(self, codigo):
        try:
            lista_final = []

            cursor_eng = conecta_engenharia.cursor()
            cursor_eng.execute("""
                            SELECT ipt.id_arquivo, arq.caminho, ipt.PART_NUMBER, ipt.DESCRIPTION 
                            FROM PROPRIEDADES_IPT ipt
                            INNER JOIN ARQUIVOS arq ON ipt.id_arquivo = arq.id
                            where ipt.AUTHORITY = ?
                            """, (codigo,))
            dados_ipt = cursor_eng.fetchall()
            if dados_ipt:
                for i in dados_ipt:
                    lista_final.append(i)

            cursor_eng.execute("""
                                SELECT iam.id_arquivo, arq.caminho, iam.PART_NUMBER, iam.DESCRIPTION 
                                FROM PROPRIEDADES_IAM iam
                                INNER JOIN ARQUIVOS arq ON iam.id_arquivo = arq.id
                                where iam.AUTHORITY = ?
                                """, (codigo,))
            dados_iam = cursor_eng.fetchall()

            if dados_iam:
                for ii in dados_iam:
                    lista_final.append(ii)

            return lista_final or []

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_arquivo_duplicado_ipt_iam(self, nome_base):
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                        SELECT id, ARQUIVO, NOME_BASE, TIPO_ARQUIVO, CLASSIFICACAO, caminho
                        FROM arquivos where NOME_BASE = ? AND TIPO_ARQUIVO IN ('IPT', 'IAM')
                    """, (nome_base,))
            return cursor.fetchall() or []

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_arquivo_divergente(self, id_arquivo):
        cursor = conecta_engenharia.cursor()
        cursor.execute("""
            SELECT id, ID_TIPO_DIVERGENCIA, ID_ARQUIVO, OBS
            FROM DIVERGENCIAS where ID_ARQUIVO = ?
        """, (id_arquivo,))
        return cursor.fetchall() or []

    def buscar_props(self, dados_arq):
        cursor = conecta_engenharia.cursor()

        if dados_arq["tipo"] == "IAM":
            cursor.execute("""
                SELECT AUTHORITY, DESCRIPTION, COST_CENTER, REVISION_NUMBER, PART_NUMBER, ENGINEER, STOCK_NUMBER, TOTAL_ITENS  
                FROM PROPRIEDADES_IAM
                WHERE ID_ARQUIVO=?
            """, (dados_arq["id"],))
            dados_iam = cursor.fetchall()

            return [
                {
                    "codigo": cod,
                    "descricao": desc,
                    "cod_mat": cod_mat,
                    "desc_mat": desc_mat,
                    "num_desenho": part,
                    "ncm": ncm,
                    "estoque_item": est_item,
                    "tot_itens": tot_itens
                }
                for cod, desc, cod_mat, desc_mat, part, ncm, est_item, tot_itens in dados_iam
            ]

        elif dados_arq["tipo"] == "IPT":
            cursor.execute("""
                SELECT AUTHORITY, DESCRIPTION, COST_CENTER, REVISION_NUMBER, PART_NUMBER, ENGINEER, STOCK_NUMBER, COMPRIMENTO    
                FROM PROPRIEDADES_IPT
                WHERE ID_ARQUIVO=?
            """, (dados_arq["id"],))

            return [
                {
                    "codigo": cod,
                    "descricao": desc,
                    "cod_mat": cod_mat,
                    "desc_mat": desc_mat,
                    "num_desenho": part,
                    "ncm": ncm,
                    "estoque_item": est_item,
                    "compr": compr
                }
                for cod, desc, cod_mat, desc_mat, part, ncm, est_item, compr in cursor.fetchall()
            ]

        return []

    def props_vazias(self, props):
        for p in props:
            if any(v is not None for v in p.values()):
                return False
        return True

    def limpar_texto(self, txt):
        txt = normalizar_texto(txt)

        txt = txt.replace(",", ".")
        txt = re.sub(r"[^\w\s.]", " ", txt)

        # remove palavras irrelevantes
        stopwords = [
            "COM", "DE", "DA", "DO", "PARA", "C/", "P/",
            "TIPO", "MODELO", "REF", "MM", "POL", "POLEGADA"
        ]

        palavras = [
            p for p in txt.split()
            if p not in stopwords
        ]

        return " ".join(palavras)

    def similaridade(self, a, b):
        from difflib import SequenceMatcher

        return SequenceMatcher(None, a, b).ratio()

    def classificar(self, desc_inv, desc_erp):
        a = self.limpar_texto(desc_inv)
        b = self.limpar_texto(desc_erp)

        palavras_a = set(a.split())
        palavras_b = set(b.split())

        intersecao = palavras_a & palavras_b
        qtd_iguais = len(intersecao)

        # 🔥 REGRA NOVA — 3 palavras iguais → OK direto
        if qtd_iguais >= 2:
            return "OK", 1.0

        # 🔴 REGRA — nenhuma palavra em comum → erro
        if qtd_iguais == 0:
            return "ERRO_GRAVE", 0

        # 🔴 REGRA — números diferentes → erro
        nums_a = set(re.findall(r"\d+", a))
        nums_b = set(re.findall(r"\d+", b))

        # 🔴 só erro se NÃO houver interseção
        if nums_a and nums_b and nums_a.isdisjoint(nums_b):
            return "ERRO_GRAVE", 0

        # 🔧 fallback com similaridade
        score = self.similaridade(a, b)

        if score > 0.7:
            return "DUVIDOSO", score

        return "ERRO_GRAVE", score

    def palavras(self, txt):
        return set(self.limpar_texto(txt).split())

    def tratar_resultado(self, cursor_eng, lista):
        lista_validos = []

        print("\n                 ============================")
        print("                    tratar_resultado")
        for item in lista:
            id_arquivo = item['id']

            tipo_arquivo = item['tipo']
            classificacao = item['classificacao']
            caminho_arquivo = item['caminho']

            nome_base = item['nome_base']

            id_origem = item["id_origem"]

            arq_duplicados = self.consulta_arquivo_duplicado_ipt_iam(nome_base)

            if len(arq_duplicados) > 1:
                if "\\inventor\\biblioteca" not in caminho_arquivo:
                    match = padrao_desenho.search(nome_base)
                    if match:
                        for linha in arq_duplicados:
                            ides_arqs = linha[0]
                            dados = (1, ides_arqs, f"{id_arquivo} CADA ARQUIVO SEPARADO - IPT/IAM", id_origem)
                            self.insert_divergencia(dados)
                        continue

            codigo = str(item.get("codigo") or "").strip()
            descricao = str(item.get("descricao") or "").strip()

            cod_mat = str(item.get("cod_mat") or "").strip()
            desc_mat = str(item.get("desc_mat") or "").strip()

            num_desenho = item["num_desenho"]
            ncm = item["ncm"]

            est_item = str(item.get("estoque_item") or "").strip()

            tot_itens_iam = item["tot_itens"]

            compr_ipt = item["compr"]

            if est_item == "FANTASMA":
                continue

            if codigo:
                # 🔒 só aceita número
                if not codigo.isdigit():
                    dados_cod = []
                else:
                    dados_cod = self.consulta_codigo_prod_erp(codigo)
            else:
                dados_cod = []

            if cod_mat:
                # 🔒 só aceita número
                if not cod_mat.isdigit():
                    dados_cod_mat = []
                else:
                    dados_cod_mat = self.consulta_codigo_prod_erp(cod_mat)
            else:
                dados_cod_mat = []

            if "\\inventor\\biblioteca" in caminho_arquivo:
                tem_propriedade = self.consulta_existe_propriedades(id_arquivo)
                if not tem_propriedade:
                    self.inserir_fila_conferencia(id_arquivo)
                    dados = (33, id_arquivo, "BIBLIOTECA", id_origem)
                    self.insert_divergencia(dados)
                    continue
                else:
                    if not descricao:
                        dados = (7, id_arquivo, "BIBLIOTECA", id_origem)
                        self.insert_divergencia(dados)
                        continue
                    if not codigo:
                        dados = (30, id_arquivo, f"{descricao}", id_origem)
                        self.insert_divergencia(dados)
                        continue

                    if not dados_cod:
                        print("- BIBLIOTECA: codigo ERP", codigo, descricao, dados_cod, id_arquivo, "caminho:",
                              caminho_arquivo)
                        continue

                descricao_erp = dados_cod[0][1]

                status, score = self.classificar(descricao, descricao_erp)

                if status == "ERRO_GRAVE":
                    dados = (5, id_arquivo, f"BIBLIOTECA ERRO_GRAVE {descricao}", id_origem)
                    self.insert_divergencia(dados)
                    continue
            else:
                tem_propriedade = self.consulta_existe_propriedades(id_arquivo)
                if not tem_propriedade:
                    self.inserir_fila_conferencia(id_arquivo)
                    dados = (33, id_arquivo, "BIBLIOTECA", id_origem)
                    self.insert_divergencia(dados)
                    continue
                else:
                    if not num_desenho:
                        dados = (26, id_arquivo, f"NOSSO", id_origem)
                        self.insert_divergencia(dados)
                        continue
                    if nome_base != num_desenho:
                        dados = (16, id_arquivo, f"NOSSO - {nome_base}/{num_desenho}", id_origem)
                        self.insert_divergencia(dados)
                        continue
                    if not descricao:
                        dados = (7, id_arquivo, "NOSSO", id_origem)
                        self.insert_divergencia(dados)
                        continue
                    if not codigo:
                        match = padrao_desenho.search(nome_base)
                        if match:
                            ref = f"D {nome_base}"

                            dados_ref = self.consulta_referencia_prod_erp(ref)
                            if not dados_ref:
                                if tipo_arquivo == "IAM" and classificacao == "NOSSO":
                                    filhos = self.consulta_estrutura_eng(cursor_eng, id_arquivo)
                                    if filhos:
                                        if ncm:
                                            if len(descricao) > 30:
                                                dados = (23, id_arquivo, f"Descrição: {descricao}", id_origem)
                                                self.insert_divergencia(dados)
                                                continue
                                            if len(ref) > 20:
                                                dados = (23, id_arquivo, f"Referência: {ref}", id_origem)
                                                self.insert_divergencia(dados)
                                                continue
                                            self.insert_pre_cadastro(id_arquivo, descricao, ref, id_origem)
                                            continue
                                        else:
                                            dados = (2, id_arquivo, "IA, NOSSO, SEM CÓDIGO", id_origem)
                                            self.insert_divergencia(dados)
                                            continue
                                    else:
                                        print("- IAM SEM ESTRUTURA")
                                        continue
                                if tipo_arquivo == "IPT" and classificacao == "NOSSO":
                                    if ncm:
                                        if len(descricao) > 30:
                                            dados = (23, id_arquivo, f"Descrição: {descricao}", id_origem)
                                            self.insert_divergencia(dados)
                                            continue
                                        if len(ref) > 20:
                                            dados = (23, id_arquivo, f"Referência: {ref}", id_origem)
                                            self.insert_divergencia(dados)
                                            continue
                                        self.insert_pre_cadastro(id_arquivo, descricao, ref, id_origem)
                                        continue
                                    else:
                                        dados = (2, id_arquivo, "IPT, NOSSO, SEM CÓDIGO", id_origem)
                                        self.insert_divergencia(dados)
                                        continue
                            elif len(dados_ref) > 1:
                                print("- DESENHO ENCONTRADO EM MAIS PRODUTOS:", dados_ref)
                                continue
                            else:
                                cod_erp_ref = dados_ref[0][1]
                                dados_inventor = (id_arquivo, "Authority", cod_erp_ref, caminho_arquivo, nome_base)
                                self.insert_propriedade_inventor(dados_inventor, id_origem)
                                continue

                        else:
                            dados = (10, id_arquivo, "VERIFICAR SE É BORRACHA E DEFINIR PADRÃO", id_origem)
                            self.insert_divergencia(dados)
                            continue

                    if not dados_cod:
                        dados = (12, id_arquivo, f"CÓDIGO LANÇADO NO INVENTOR: {codigo}", id_origem)
                        self.insert_divergencia(dados)
                        continue

                    descricao_erp = dados_cod[0][1]
                    status, score = self.classificar(descricao, descricao_erp)
                    if status == "ERRO_GRAVE":
                        dados = (5, id_arquivo, f"ERRO_GRAVE {descricao}", id_origem)
                        self.insert_divergencia(dados)
                        continue

                    match = padrao_desenho.search(nome_base)
                    if match:
                        ref_erp = dados_cod[0][2]
                        ref_erp = re.sub(r"[^\d.]", "", ref_erp)  # remove tudo que não é número ou ponto
                        ref_erp_padrao = re.sub(r"\.+$", "", ref_erp)  # saída: 47.00.014.07
                        if ref_erp_padrao:
                            if ref_erp_padrao != nome_base:
                                dados_duplicados = self.consulta_codigo_propriedade_eng(codigo)
                                if len(dados_duplicados) > 1:
                                    for linha in dados_duplicados:
                                        ides_arqs = linha[0]
                                        dados = (11, ides_arqs, f"Código ERP duplicado nas propriedades: {codigo}", id_origem)
                                        self.insert_divergencia(dados)
                                else:
                                    dados = (4, id_arquivo, f"COD PRODUTO {codigo}", id_origem)
                                    self.insert_divergencia(dados)
                                continue

                    conj = dados_cod[0][6]
                    if conj == 10:
                        if not desc_mat:
                            if tipo_arquivo == "IAM" and tot_itens_iam == 1:
                                dados = (8, id_arquivo, f"ACABADO, IAM 1 ITEM, PRODUTO {codigo} - {descricao}", id_origem)
                                self.insert_divergencia(dados)
                                continue
                            if tipo_arquivo == "IPT":
                                dados = (8, id_arquivo, f"ACABADO, IPT, PRODUTO {codigo} - {descricao}", id_origem)
                                self.insert_divergencia(dados)
                                continue

                        if not cod_mat:
                            if tipo_arquivo == "IAM" and tot_itens_iam == 1:
                                filhos = self.consulta_estrutura_eng(cursor_eng, id_arquivo)
                                if filhos:
                                    if len(desc_mat) > 30:
                                        dados = (23, id_arquivo, f"Descrição Mat: {desc_mat}", id_origem)
                                        self.insert_divergencia(dados)
                                        continue

                                    cursor = conecta.cursor()
                                    cursor.execute("""
                                                    SELECT ID, OBS, DESCRICAO
                                                    FROM PRODUTOPRELIMINAR
                                                    WHERE DESCRICAO = ?
                                                """, (desc_mat,))
                                    tem_preliminar = cursor.fetchall()

                                    if not tem_preliminar:
                                        dados = (13, id_arquivo, f"ACABADO, IPT, PRODUTO {codigo} - {descricao}", id_origem)
                                        self.insert_divergencia(dados)
                                        continue
                                else:
                                    print("sem filhos")
                                    continue

                            if tipo_arquivo == "IPT":
                                if len(desc_mat) > 30:
                                    dados = (23, id_arquivo, f"Descrição Mat: {desc_mat}", id_origem)
                                    self.insert_divergencia(dados)
                                    continue

                                cursor = conecta.cursor()
                                cursor.execute("""
                                                SELECT ID, OBS, DESCRICAO
                                                FROM PRODUTOPRELIMINAR
                                                WHERE DESCRICAO = ?
                                            """, (desc_mat,))
                                tem_preliminar = cursor.fetchall()

                                if not tem_preliminar:
                                    dados = (13, id_arquivo, f"ACABADO, IPT, PRODUTO {codigo} - {descricao}", id_origem)
                                    self.insert_divergencia(dados)
                                    continue

                        if (tipo_arquivo == "IAM" and tot_itens_iam == 1) or (tipo_arquivo == "IPT"):
                            if not dados_cod_mat:
                                dados = (15, id_arquivo, f"CÓDIGO LANÇADO NO INVENTOR: {cod_mat}", id_origem)
                                self.insert_divergencia(dados)
                                continue

                            descricao_mat_erp = dados_cod_mat[0][1]
                            status, score = self.classificar(desc_mat, descricao_mat_erp)
                            if status == "ERRO_GRAVE":
                                dados = (14, id_arquivo, f"ERRO_GRAVE {desc_mat}", id_origem)
                                self.insert_divergencia(dados)
                                continue
                            else:
                                um = dados_cod_mat[0][3]
                                if um == "KG" or um == "MT" or um == "MM":
                                    kg_mt = dados_cod_mat[0][5]
                                    if not compr_ipt:
                                        dados = (6, id_arquivo, f"PRODUTO {codigo} - {desc_mat} - {um}", id_origem)
                                        self.insert_divergencia(dados)
                                        continue

                                    if um == "KG":
                                        if not kg_mt:
                                            dados = (9, id_arquivo, f"PRODUTO {codigo} - {desc_mat} - {um} - COMPR: {compr_ipt}", id_origem)
                                            self.insert_divergencia(dados)
                                            continue

            # if id_arquivo == 16733:
            #     print("\n", id_arquivo, num_desenho, ncm, tipo_arquivo, classificacao, num_desenho)
            #     print(caminho_arquivo)
            #     print(dados_cod, dados_cod_mat)
            #     print("           Cód:", codigo, "Descr:", descricao, "CodMat:", cod_mat, "DescrMat:", desc_mat)
            #     print("ID ORIGEM:", id_origem)
            #     print("\n")

            lista_validos.append(item)

        if lista_validos:
            self.tratar_estruturas(lista_validos)

        self.definir_projeto_para_pi()

    def definir_projeto_para_pi(self):
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute(f"SELECT ID, ID_ARQUIVO, QTDE, ID_CLIENTE, SOLICITANTE, OBS, PREVISAO_ENTREGA "
                           f"FROM PROJETO "
                           f"where status = 'A';")
            dados_projetos = cursor.fetchall()

            if dados_projetos:
                for i in dados_projetos:
                    id_proj, id_arq, qtde, id_cliente, solicitante, obs, data_entrega = i

                    sql = """
                            SELECT id, ID_TIPO_DIVERGENCIA, ID_ARQUIVO, OBS, RESOLVIDO
                            FROM DIVERGENCIAS where id_origem = ?
                            """
                    cur = conecta_engenharia.cursor()
                    cur.execute(sql, (id_arq,))
                    dados_divergencias = cur.fetchall()
                    if not dados_divergencias:
                        hoje = date.today()

                        if not qtde:
                            dados = (29, id_arq, f"Projeto: {id_proj} - Falta a quantidade para criar PI", id_arq)
                            self.insert_divergencia(dados)
                        elif not id_cliente:
                            dados = (29, id_arq, f"Projeto: {id_proj} - Falta definir o cliente para criar PI", id_arq)
                            self.insert_divergencia(dados)
                        elif not solicitante:
                            dados = (29, id_arq, f"Projeto: {id_proj} - Falta definir o Solicitante para criar PI", id_arq)
                            self.insert_divergencia(dados)
                        elif not obs:
                            dados = (29, id_arq, f"Projeto: {id_proj} - Falta definir o campo Observação para criar PI", id_arq)
                            self.insert_divergencia(dados)
                        elif not data_entrega or data_entrega <= hoje:
                            dados = (29, id_arq, f"Projeto: {id_proj} - Data de entrega menor do que a data atual para criar PI", id_arq)
                            self.insert_divergencia(dados)
                        else:
                            cursor = conecta.cursor()
                            cursor.execute(f"select id, razao "
                                           f"from clientes "
                                           f"where id = {id_cliente};")
                            dados_cliente = cursor.fetchall()

                            if not dados_cliente:
                                dados = (29, id_arq, f"Projeto: {id_proj} - Falta definir o cliente existente para criar PI",
                                         id_arq)
                                self.insert_divergencia(dados)
                            else:
                                codigo_erp = self.consulta_propriedade_ipt_iam(id_arq, "AUTHORITY")
                                if codigo_erp:
                                    print(id_proj, id_arq, qtde, id_cliente, solicitante, obs, data_entrega)
                                    dados = (hoje, id_cliente, solicitante, obs, codigo_erp, qtde, data_entrega, id_proj)
                                    self.salvar_pedido_interno(dados)
                                else:
                                    print("CÓDIGO PARA INCLUIR NO PI NÃO ENCONTRADO")


        except Exception as e:
            trata_excecao(e)
            raise

    def salvar_pedido_interno(self, dados):
        try:
            hoje, id_cliente, solicitante, obs, codigo_erp, qtde, entrega, id_proj  = dados

            nome_computador = socket.gethostname()

            cursor = conecta.cursor()
            cursor.execute("select GEN_ID(GEN_PEDIDOINTERNO_ID,0) from rdb$database;")
            ultimo_ped0 = cursor.fetchall()
            ultimo_ped1 = ultimo_ped0[0]
            ultimo_ped = int(ultimo_ped1[0]) + 1

            cursor = conecta.cursor()
            cursor.execute(f"Insert into pedidointerno (ID, EMISSAO, ID_CLIENTE, SOLICITANTE, OBS, NOME_PC, "
                           f"STATUS) "
                           f"values (GEN_ID(GEN_PEDIDOINTERNO_ID,1), "
                           f"'{hoje}', '{id_cliente}', '{solicitante}', '{obs}', "
                           f"'{nome_computador}', 'A');")

            qtdezinha_float = valores_para_float(qtde)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, embalagem FROM produto where codigo = '{codigo_erp}';")
            dados_produto = cursor.fetchall()
            id_produto, codigo, embalagem = dados_produto[0]

            cursor = conecta.cursor()
            cursor.execute(f"Insert into produtopedidointerno (ID_PRODUTO, ID_PEDIDOINTERNO, QTDE, "
                           f"DATA_PREVISAO, STATUS) "
                           f"values ({id_produto}, {ultimo_ped}, {qtdezinha_float}, '{entrega}', "
                           f"'A');")

            cur = conecta_engenharia.cursor()
            cur.execute("""
                UPDATE PROJETO
                SET status = 'B'
                WHERE ID = ?
            """, (id_proj,))

            conecta.commit()
            conecta_engenharia.commit()
            print("salvado pedido interno")

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_existe_propriedades(self, id_arquivo):
        try:
            tem_propriedade = ""

            cursor_eng = conecta_engenharia.cursor()
            sql = f"""
                SELECT id, id_arquivo
                FROM PROPRIEDADES_IPT
                WHERE id_arquivo = ?
            """
            cursor_eng.execute(sql, (id_arquivo,))
            dados_ipt = cursor_eng.fetchall()
            if dados_ipt:
                tem_propriedade = dados_ipt

            sql = f"""
                SELECT id, id_arquivo
                FROM PROPRIEDADES_IAM
                WHERE id_arquivo = ?
            """
            cursor_eng.execute(sql, (id_arquivo,))
            dados_iam = cursor_eng.fetchall()

            if dados_iam:
                tem_propriedade = dados_iam

            return tem_propriedade

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_propriedade_ipt_iam(self, id_arquivo, propr):
        try:
            propriedade_escolhida = ""

            cursor_eng = conecta_engenharia.cursor()
            sql = f"""
                SELECT ipt.id_arquivo, ipt.{propr}
                FROM PROPRIEDADES_IPT ipt
                WHERE ipt.id_arquivo = ?
            """
            cursor_eng.execute(sql, (id_arquivo,))
            dados_ipt = cursor_eng.fetchall()
            if dados_ipt:
                for i in dados_ipt:
                    propriedade_escolhida = i[1].strip()

            sql = f"""
                SELECT iam.id_arquivo, iam.{propr}
                FROM PROPRIEDADES_IAM iam
                WHERE iam.id_arquivo = ?
            """
            cursor_eng.execute(sql, (id_arquivo,))
            dados_iam = cursor_eng.fetchall()

            if dados_iam:
                for ii in dados_iam:
                    propriedade_escolhida = ii[1].strip()

            return propriedade_escolhida

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_estrutura_eng_atual(self, cursor, id_pai):
        try:
            cursor.execute("""
                SELECT e.ID_FILHO, e.QTDE
                FROM ESTRUTURA e
                WHERE e.ID_PAI = ?
                  AND NOT EXISTS (
                      SELECT 1
                      FROM PROPRIEDADES_IPT p
                      WHERE p.ID_ARQUIVO = e.ID_FILHO
                        AND p.STOCK_NUMBER = 'FANTASMA'
                  )
                  AND NOT EXISTS (
                      SELECT 1
                      FROM PROPRIEDADES_IAM p
                      WHERE p.ID_ARQUIVO = e.ID_FILHO
                        AND p.STOCK_NUMBER = 'FANTASMA'
                  )
            """, (id_pai,))

            estrutura = cursor.fetchall()

            return estrutura

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_estrutura_erp_atual(self, cursor, id_produto):
        try:
            cursor.execute("""
                SELECT id, codigo, id_versao 
                FROM produto 
                WHERE id = ?
            """, (id_produto,))
            select_prod = cursor.fetchall()
            id_pai, cod, id_versao = select_prod[0]

            if id_versao:
                cursor = conecta.cursor()
                cursor.execute("""
                    SELECT prod.codigo, estprod.quantidade
                    FROM estrutura_produto estprod
                    INNER JOIN produto prod ON estprod.id_prod_filho = prod.id
                    WHERE estprod.id_estrutura = ?
                """, (id_versao,))
                sel_estrutura = cursor.fetchall()

                return sel_estrutura

            return []

        except Exception as e:
            trata_excecao(e)
            raise

    def tratar_estruturas(self, lista):
        cursor_eng = conecta_engenharia.cursor()

        print("\n                  ============================")
        print("                     tratar_estruturas")

        for item in lista:
            id_arquivo = item['id']
            tipo_arquivo = item['tipo']

            id_origem = item["id_origem"]

            caminho_arquivo = item['caminho']

            codigo = str(item.get("codigo") or "").strip()
            dados_erp_prod = self.consulta_codigo_prod_erp(codigo)

            conj = dados_erp_prod[0][6]
            if conj == 10:
                cod_mat = str(item.get("cod_mat") or "").strip()

                compr_ipt = item.get("compr")
                compr_ipt_float = None

                if tipo_arquivo == "IPT":
                    compr_ipt_float = self.extrair_numero(compr_ipt)

                if tipo_arquivo == "IAM":
                    estrutura_eng = self.consulta_estrutura_eng_atual(cursor_eng, id_arquivo)

                    if not estrutura_eng:
                        dados = (28, id_arquivo, f"", id_origem)
                        self.insert_divergencia(dados)
                        continue

                    estrutura_nova = []
                    erro_estrutura = False

                    for id_arquivo_f, qtde_f in estrutura_eng:
                        item_f = next((i for i in lista if i["id"] == id_arquivo_f), None)

                        if not item_f:
                            dados_f_diverg = self.consulta_arquivo_divergente(id_arquivo_f)

                            if not dados_f_diverg:
                                ref = self.consulta_propriedade_ipt_iam(id_arquivo_f, "PART_NUMBER")

                                if ref:
                                    ref = f"D {ref}"

                                    cursor = conecta.cursor()
                                    cursor.execute("""
                                                    SELECT ID, OBS, DESCRICAO
                                                    FROM PRODUTOPRELIMINAR
                                                    WHERE REFERENCIA = ?
                                                """, (ref,))
                                    tem_preliminar = cursor.fetchall()

                                    if not tem_preliminar:
                                        dados = (17, id_arquivo_f, f"ID DO PAI: {id_arquivo}", id_origem)
                                        self.insert_divergencia(dados)
                                else:
                                    dados = (17, id_arquivo_f, f"ID DO PAI: {id_arquivo}", id_origem)
                                    self.insert_divergencia(dados)

                            erro_estrutura = True
                            break

                        codigo_f = str(item_f.get("codigo") or "").strip()

                        if not codigo_f:
                            print("IAM Filho sem código: item_f:", item_f)
                            erro_estrutura = True
                            break

                        tipo_arquivo_f = item_f['tipo']
                        compr_ipt_f = item_f.get("compr")
                        compr_ipt_f_float = None

                        if tipo_arquivo_f == "IPT":
                            compr_ipt_f_float = self.extrair_numero(compr_ipt_f)

                        qtde_calc = self.calcular_qtde_erp(codigo_f, qtde_f, compr_ipt_f_float, id_arquivo_f)

                        if qtde_calc is None:
                            print("IAM Erro ao calcular quantidade: codigo_f:", codigo_f)
                            erro_estrutura = True
                            break

                        estrutura_nova.append((codigo_f, qtde_calc))

                    if erro_estrutura:
                        continue

                else:
                    if not cod_mat:
                        print("IPT sem matéria-prima: Código: ", codigo)
                        continue

                    dados_cod_mat = self.consulta_codigo_prod_erp(cod_mat)
                    if not dados_cod_mat:
                        print("IPT Matéria-prima não existe no ERP: codMat:", cod_mat)
                        continue

                    qtde_calc = self.calcular_qtde_erp(cod_mat, 1, compr_ipt_float, id_arquivo)

                    if qtde_calc is None:
                        print("IPT Erro ao calcular quantidade IPT: codMat:", cod_mat)
                        continue

                    estrutura_nova = [(cod_mat, qtde_calc)]

                if estrutura_nova:
                    self.atualiza_estrutura_erp_v2(codigo, dados_erp_prod, item, estrutura_nova)

            else:
                if "\\inventor\\biblioteca" not in caminho_arquivo:
                    self.tratar_idw(item)

        return None

    def calcular_qtde_erp(self, cod_prod, qtde_eng, compr_ipt, id_arquivo):
        cursor = conecta.cursor()
        cursor.execute("""
            SELECT unidade, KILOSMETRO
            FROM produto
            WHERE codigo = ?
        """, (cod_prod,))
        row = cursor.fetchone()

        if not row:
            print("Produto não encontrado (estrutura):", cod_prod)
            return None

        unidade, kg_mt = row

        unidade = (unidade or "").upper().strip()

        # 🔹 UNIDADE simples (UN, PC, PÇ...)
        if unidade in ("UN", "PC", "PÇ"):
            valor = float(qtde_eng)

            if not valor.is_integer():
                print("Quantidade fracionada inválida para unidade:", unidade, cod_prod, valor)
                return None

            qtde_int = int(valor)
            return qtde_int

        # 🔹 KG → precisa converter
        if unidade == "KG":
            if not kg_mt:
                print("Produto KG sem KILOSMETRO:", cod_prod)
                return None

            if not compr_ipt:
                print("Falta comprimento para conversão KG:", cod_prod)
                return None

            qtde_f = self.arredondar_qtde((valores_para_float(kg_mt) * (compr_ipt / 1000) * float(qtde_eng)), 2)
            return qtde_f

        if unidade == "MT":
            if not compr_ipt:
                print("Falta comprimento para unidade:", unidade, cod_prod, id_arquivo)
                return None

            compr_m = self.arredondar_qtde((valores_para_float(compr_ipt / 1000) * float(qtde_eng)), 2)

            return compr_m

        if unidade == "M²":
            if not compr_ipt:
                print("Falta comprimento para unidade:", unidade, cod_prod, id_arquivo)
                return None

            compr_m = self.arredondar_qtde(valores_para_float(compr_ipt), 2)
            return compr_m

        if unidade == "MM":
            if not compr_ipt:
                print("Falta comprimento para unidade:", unidade, cod_prod)
                return None

            return self.arredondar_qtde((compr_ipt * float(qtde_eng)), 2)

        if unidade in ("CT", "CN"):
            qtde_final = self.arredondar_qtde((valores_para_float(qtde_eng / 100)), 2)

            return qtde_final

        # 🔴 qualquer outra unidade não tratada
        print("Unidade não tratada:", unidade, cod_prod, id_arquivo)
        return None

    def extrair_numero(self, valor):
        if valor is None:
            return None

        s = str(valor).lower().strip()

        # mantém números, ponto e vírgula
        s = re.sub(r"[^\d.,]", "", s)

        if not s:
            return None

        # se tem vírgula e não tem ponto → vírgula é decimal
        if "," in s and "." not in s:
            s = s.replace(",", ".")
        else:
            # remove separador de milhar (vírgula)
            s = s.replace(",", "")

        try:
            return float(s)
        except:
            return None

    def arredondar_qtde(self, qtde, casas_decimais):
        qtde_final = round(qtde, casas_decimais)
        return qtde_final

    def atualiza_estrutura_erp_v2(self, cod_prod, dados_erp_prod, item, estrutura_nova):
        id_arquivo = item['id']
        id_origem = item["id_origem"]

        cursor = conecta.cursor()

        cursor.execute(
            f"""
            SELECT ord.numero, produto.codigo, produto.descricao,
                   COALESCE(produto.obs, '') as obs,
                   produto.unidade,
                   ord.quantidade,
                   ord.datainicial
            FROM ordemservico as ord
            INNER JOIN produto ON ord.produto = produto.id
            WHERE produto.codigo = {cod_prod}
              AND ord.status = 'A'
            """
        )
        dados_op = cursor.fetchall()

        descricao = str(item.get("descricao") or "").strip()

        cursor.execute("""
            SELECT id, id_versao
            FROM produto
            WHERE codigo = ?
        """, (cod_prod,))

        row = cursor.fetchone()

        if not row:
            print("Produto não encontrado:", cod_prod)
            return

        id_prod, id_versao_atual = row

        # ==========================================================
        # Estrutura nova do Inventor
        # ==========================================================

        estrutura_nova_set = set(
            (cod, valores_para_float(self.arredondar_qtde(qtde, 2)))
            for cod, qtde in estrutura_nova
        )

        # ==========================================================
        # Busca todas as versões do ERP
        # ==========================================================

        cursor.execute("""
            SELECT id
            FROM estrutura
            WHERE id_produto = ?
        """, (id_prod,))

        estruturas_existentes = cursor.fetchall()

        # Não existe nenhuma estrutura
        if not estruturas_existentes:

            tipo_divergencia = 18  # Criar estrutura
        else:
            id_estrutura_igual = None

            for (id_estrutura,) in estruturas_existentes:

                cursor.execute("""
                    SELECT prod.codigo, est.quantidade
                    FROM estrutura_produto est
                    JOIN produto prod
                         ON prod.id = est.id_prod_filho
                    WHERE est.id_estrutura = ?
                """, (id_estrutura,))

                estrutura_erp = cursor.fetchall()

                estrutura_erp_set = set(
                    (cod, valores_para_float(self.arredondar_qtde(qtde, 2)))
                    for cod, qtde in estrutura_erp
                )

                if estrutura_erp_set == estrutura_nova_set:
                    id_estrutura_igual = id_estrutura
                    break

            # ======================================================
            # Encontrou estrutura igual
            # ======================================================

            if id_estrutura_igual:
                # Já é a versão ativa
                if id_estrutura_igual == id_versao_atual:
                    self.tratar_idw(item)
                    return

                # Existe versão igual, mas não está ativa
                tipo_divergencia = 31  # Ativar versão existente

            # ======================================================
            # Não encontrou nenhuma igual
            # ======================================================

            else:
                tipo_divergencia = 32  # Criar nova versão

        # ==========================================================
        # Grava divergência
        # ==========================================================

        if dados_op:
            msg = f"TEM OP ABERTA - Cód: {cod_prod} - {descricao}"
        else:
            msg = f"SEM OP ABERTA - Cód: {cod_prod} - {descricao}"

        dados = (
            tipo_divergencia,
            id_arquivo,
            msg,
            id_origem
        )

        self.insert_divergencia(dados)

    def atualiza_estrutura_erp(self, cod_prod, dados_erp_prod, item, estrutura_nova):
        id_arquivo = item['id']

        id_origem = item["id_origem"]

        cursor = conecta.cursor()
        cursor.execute(f"SELECT ord.numero, produto.codigo, produto.descricao, "
                       f"COALESCE(produto.obs, '') as obs, produto.unidade, ord.quantidade, ord.datainicial "
                       f"FROM ordemservico as ord "
                       f"INNER JOIN produto ON ord.produto = produto.id "
                       f"where produto.codigo = {cod_prod} and ord.status = 'A';")
        dados_op = cursor.fetchall()

        descricao = str(item.get("descricao") or "").strip()

        cursor = conecta.cursor()
        cursor.execute("""
            SELECT id, id_versao
            FROM produto
            WHERE codigo = ?
        """, (cod_prod,))
        row = cursor.fetchone()

        if not row:
            print("Produto não encontrado:", cod_prod)
            return

        id_prod, id_versao_atual = row

        # 🔹 normaliza nova estrutura
        estrutura_nova_set = set(
            (cod, valores_para_float(self.arredondar_qtde(qtde, 2)))
            for cod, qtde in estrutura_nova
        )

        # 🔹 busca TODAS as versões
        cursor.execute("""
            SELECT id
            FROM estrutura
            WHERE id_produto = ?
        """, (id_prod,))

        estruturas_existentes = cursor.fetchall()

        for (id_estrutura,) in estruturas_existentes:
            cursor.execute("""
                SELECT prod.codigo, est.quantidade
                FROM estrutura_produto est
                JOIN produto prod ON prod.id = est.id_prod_filho
                WHERE est.id_estrutura = ?
            """, (id_estrutura,))

            estrutura_erp = cursor.fetchall()

            estrutura_erp_set = set(
                (cod, valores_para_float(self.arredondar_qtde(qtde, 2)))
                for cod, qtde in estrutura_erp
            )

            # 🔥 ACHOU IGUAL → só ativa
            if estrutura_erp_set == estrutura_nova_set:
                self.tratar_idw(item)
                return

        if dados_op:
            dados = (18, id_arquivo, f"TEM OP ABERTA - Cód: {cod_prod} - {descricao}", id_origem)
            self.insert_divergencia(dados)
        else:
            dados = (18, id_arquivo, f"SEM OP ABERTA - Cód: {cod_prod} - {descricao}", id_origem)
            self.insert_divergencia(dados)

    def tratar_idw(self, item):
        id_arquivo = item['id']
        caminho_arquivo = item['caminho']
        nome_base = item['nome_base']
        tipo_arquivo = item['tipo']

        id_origem = item["id_origem"]

        codigo = str(item.get("codigo") or "").strip()
        dados_cod = self.consulta_codigo_prod_erp(codigo)
        conj = dados_cod[0][6]
        tem_desenho = dados_cod[0][7]
        if conj == 10 or tem_desenho:
            match = padrao_desenho.search(nome_base)
            if match:
                dados_arq_idw = self.consulta_arquivos_idw(nome_base)
                if dados_arq_idw:
                    if len(dados_arq_idw) == 1:
                        id_arquivo_idw = dados_arq_idw[0][0]
                        nome_base_idw = dados_arq_idw[0][2]
                        caminho_idw = dados_arq_idw[0][5]

                        cursor_eng = conecta_engenharia.cursor()
                        cursor_eng.execute("""
                            SELECT ID_ARQUIVO, ID_ARQUIVO_REFERENCIA
                            FROM PROPRIEDADES_IDW
                            WHERE ID_ARQUIVO = ?
                        """, (id_arquivo_idw,))
                        dados_prop_idw = cursor_eng.fetchall()

                        if dados_prop_idw:
                            if len(dados_prop_idw) == 1:
                                id_arq_dentro = dados_prop_idw[0][1]
                                arquivo_dentro = self.consulta_arquivos(id_arq_dentro)

                                if not arquivo_dentro:
                                    print("PROPRIEDADE IDW - NÃO TEM ARQUIVO DE REFERENCIA", caminho_arquivo, nome_base)
                                else:
                                    nome_dentro, nome_base_dentro, tipo_dentro, classificacao_dentro, caminho_dentro = arquivo_dentro[0]
                                    if nome_base == nome_base_dentro:
                                        caminho_pdf = rf"\\Publico\C\OP\Projetos\{nome_base}.pdf"
                                        if tipo_arquivo == "IPT":
                                            if conj == 10:
                                                tem_cota = self.comparar_cotas_idw(cursor_eng, item, id_arquivo_idw)
                                                if tem_cota:
                                                    if not os.path.exists(caminho_pdf):
                                                        self.inserir_pdf_fila(id_arquivo_idw, caminho_idw)
                                                else:
                                                    dados = (19, id_arquivo_idw, f"COLOCAR COTA NO DESENHO: {nome_base_idw}.idw", id_origem)
                                                    self.insert_divergencia(dados)
                                        else:
                                            if not os.path.exists(caminho_pdf):
                                                self.inserir_pdf_fila(id_arquivo_idw, caminho_idw)
                                    else:
                                        dados = (20, id_arquivo_idw, f"VINCULO/IDW - {nome_base}/{nome_base_dentro}", id_origem)
                                        self.insert_divergencia(dados)
                            else:
                                ids = [linha[0] for linha in dados_prop_idw]
                                dados = (25, id_arquivo, f"IDS dos arquivos IDW: {ids}", id_origem)
                                self.insert_divergencia(dados)
                        else:
                            dados = (21, dados_arq_idw[0][0], f"DESENHO {nome_base}", id_origem)
                            self.insert_divergencia(dados)
                            self.inserir_fila_conferencia(dados_arq_idw[0][0])
                    else:
                        for linha in dados_arq_idw:
                            ides_arqs = linha[0]
                            dados = (1, ides_arqs, f"{id_arquivo} CADA ARQUIVO SEPARADO - IDW", id_origem)
                            self.insert_divergencia(dados)

                else:
                    dados = (3, id_arquivo, "", id_origem)
                    self.insert_divergencia(dados)

    def comparar_cotas_idw(self, cursor_eng, item_ipt, id_arquivo_idw):
        try:
            tem_cota = False

            compr_ipt = item_ipt.get("compr")
            if compr_ipt:
                compr_ipt = self.extrair_numero(compr_ipt)
                if compr_ipt:
                    compr_ipt_float = valores_para_float(compr_ipt)
                    cursor_eng.execute("""
                                        SELECT ID_ARQUIVO, VALOR_COTA
                                        FROM COTAS_IDW
                                        WHERE ID_ARQUIVO = ?
                                    """, (id_arquivo_idw,))
                    cotas_idw = cursor_eng.fetchall()

                    if cotas_idw:
                        corte_preciso = compr_ipt_float
                        corte_sobra = compr_ipt_float - 2.0

                        for id_arq, cota in cotas_idw:
                            if float(cota) == corte_preciso:
                                tem_cota = True
                            if float(cota) == corte_sobra:
                                tem_cota = True
            else:
                tem_cota = True

            return tem_cota

        except Exception as e:
            trata_excecao(e)
            raise

    def inserir_pdf_fila(self, id_arquivo, caminho_arquivo):
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                SELECT ID, ID_ARQUIVO
                FROM FILA_GERAR_PDF
                WHERE ID_ARQUIVO = ?
            """, (id_arquivo,))

            tem_na_fila = cursor.fetchall()

            if not tem_na_fila:
                sql = """
                    INSERT INTO FILA_GERAR_PDF (ID, ID_ARQUIVO) 
                    VALUES (GEN_ID(GEN_FILA_GERAR_PDF_ID,1), ?);
                """

                print(sql)

                cursor.execute(sql, (id_arquivo,)) # ✅ AQUI TAMBÉM
                conecta_engenharia.commit()

                print(" AAAAAAAAAAAAAAAAAAAAA Produto inserido na fila de PDF!", caminho_arquivo)

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    GerarFilaValidacaoERP()