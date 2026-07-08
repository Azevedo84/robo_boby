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
from core.inventor import normalizar_texto, padrao_desenho
import re
from core.conversores import valores_para_float
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import date


class ConfereDivergencias:
    def __init__(self):

        self.destinatario = ['<maquinas@unisold.com.br>']

        self.lista_itens = []

        self.processar_estruturas()

        self.processar()

    def processar_estruturas(self):
        try:
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

                        self.lista_itens.append({
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

                                self.lista_itens.append({
                                    "codigo": "",
                                    "obs": nome_base_p,
                                    "id": id_item,
                                    "id_origem": id_origem
                                })

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
                fila.extend(filhos)

            return itens

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

    def envia_email_tipo_diveregencia(self, num_div, obs_div):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA DIVEREGENCIA - DIVEREGENCIA NÃO TRATADA {num_div}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nA divergência Nº {num_div} não foi tratada!\n\n"

            body += f"{obs_div}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado DIVERGENCIA NÃO TRATADA")

        except Exception as e:
            trata_excecao(e)
            raise

    def envia_email_div_nao_resolvida(self, dados):
        try:
            id_tipo_div, id_arquivo, nome_base, descr_div, obs_div, caminho = dados

            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA DIVEREGENCIA - DIVERGÊNCIA NÃO RESOLVIDA {nome_base}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nA divergência foi marcada como resolvida mas está pendente!\n\n"
            body += f"{descr_div}\n\n"
            body += f"{obs_div}\n\n"
            body += f"{caminho}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado DIVERGENCIA RESOLVIDA MAS PENDENTE")

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

    def comparar_cotas_idw(self, cursor_eng, compr_ipt, id_arquivo_idw):
        try:
            tem_cota = False

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

    def consulta_arquivos_idw(self, nome_base):
        cursor = conecta_engenharia.cursor()
        cursor.execute("""
            SELECT id, ARQUIVO, NOME_BASE, TIPO_ARQUIVO, CLASSIFICACAO, caminho
            FROM arquivos 
            where NOME_BASE = ? 
            AND TIPO_ARQUIVO = 'IDW'
        """, (nome_base,))
        return cursor.fetchall() or []

    def atualiza_estrutura_erp(self, cod_prod, estrutura_nova, id_arquivo, id_verificar):

        from collections import Counter

        resultado_final = False

        def arredondar_qtde(qtde, casas_decimais):
            return round(qtde, casas_decimais)

        cursor = conecta.cursor()

        cursor.execute("""
            SELECT id, id_versao
            FROM produto
            WHERE codigo = ?
        """, (cod_prod,))

        row = cursor.fetchone()

        if row:
            id_prod, id_versao_atual = row

            estrutura_nova_lista = [
                (str(cod).strip(),
                 valores_para_float(arredondar_qtde(qtde, 2)))
                for cod, qtde in estrutura_nova
            ]

            cursor.execute("""
                SELECT prod.codigo, est.quantidade
                FROM estrutura_produto est
                JOIN produto prod ON prod.id = est.id_prod_filho
                WHERE est.id_estrutura = ?
            """, (id_versao_atual,))

            estrutura_erp = cursor.fetchall()

            estrutura_erp_lista = [
                (str(cod).strip(),
                 valores_para_float(arredondar_qtde(qtde, 2)))
                for cod, qtde in estrutura_erp
            ]

            contador_novo = Counter(estrutura_nova_lista)
            contador_erp = Counter(estrutura_erp_lista)

            if contador_novo == contador_erp:
                resultado_final = True
            else:
                if id_arquivo == id_verificar:
                    print("\nDiferenças encontradas:\n")

                    somente_novo = contador_novo - contador_erp
                    somente_erp = contador_erp - contador_novo

                    if somente_novo:
                        print("Existe no Inventor e não no ERP:")
                        for item, qtd in somente_novo.items():
                            print(item)

                    if somente_erp:
                        print("\nExiste no ERP e não no Inventor:")
                        for item, qtd in somente_erp.items():
                            print(item)

        return resultado_final

    def calcular_qtde_erp(self, cod_prod, qtde_eng, compr_ipt, id_arquivo):
        def arredondar_qtde(qtde, casas_decimais):
            qtde_final = round(qtde, casas_decimais)
            return qtde_final

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

            qtde_f = arredondar_qtde((valores_para_float(kg_mt) * (compr_ipt / 1000) * float(qtde_eng)), 2)
            return qtde_f

        if unidade == "MT":
            if not compr_ipt:
                print("Falta comprimento para unidade:", unidade, cod_prod, id_arquivo)
                return None

            compr_m = arredondar_qtde((valores_para_float(compr_ipt / 1000) * float(qtde_eng)), 2)
            return compr_m

        if unidade == "MM":
            if not compr_ipt:
                print("Falta comprimento para unidade:", unidade, cod_prod)
                return None

            return arredondar_qtde((compr_ipt * float(qtde_eng)), 2)

        if unidade in ("CT", "CN"):
            qtde_final = arredondar_qtde((valores_para_float(qtde_eng / 100)), 2)

            return qtde_final

        # 🔴 qualquer outra unidade não tratada
        print("Unidade não tratada:", unidade, cod_prod, id_arquivo)
        return None

    def consulta_arquivos(self, id_arquivo):
        cursor = conecta_engenharia.cursor()
        cursor.execute("""
            SELECT ARQUIVO, NOME_BASE, TIPO_ARQUIVO, CLASSIFICACAO, caminho
            FROM arquivos where ID = ?
        """, (id_arquivo,))
        return cursor.fetchall() or []

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

    def consulta_arquivo_divergente(self, id_arquivo):
        cursor = conecta_engenharia.cursor()
        cursor.execute("""
            SELECT id, ID_TIPO_DIVERGENCIA, ID_ARQUIVO, OBS
            FROM DIVERGENCIAS where ID_ARQUIVO = ?
        """, (id_arquivo,))
        return cursor.fetchall() or []

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

    def consulta_estrutura_eng_atual(self, id_pai):
        try:
            cursor = conecta_engenharia.cursor()
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

    def inserir_fila_conferencia(self, id_arquivo):
        try:
            cursor = conecta_engenharia.cursor()
            # 🔹 evita duplicado
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

            # 🔹 pega dados do arquivo
            cursor.execute("""
                SELECT NOME_BASE, TIPO_ARQUIVO
                FROM ARQUIVOS
                WHERE ID = ?
            """, (id_arquivo,))

            row = cursor.fetchone()

            if not row:
                return

            nome_base, tipo = row

            # 🔥 só IPT e IAM precisam garantir IDW
            if tipo not in ("IPT", "IAM"):
                return

            # 🔹 extrai código do desenho
            match = padrao_desenho.search(nome_base)

            if not match:
                return

            codigo = match.group()

            # 🔹 busca IDW pelo código
            cursor.execute("""
                SELECT ID, TIPO_ARQUIVO, CAMINHO
                FROM ARQUIVOS
                WHERE NOME_BASE = ?
                  AND TIPO_ARQUIVO = 'IDW'
            """, (codigo,))

            resultados = cursor.fetchall()

            if len(resultados) == 1:
                id_idw = resultados[0][0]

                # 🔹 evita duplicado
                cursor.execute("""
                    SELECT 1 FROM FILA_CONFERENCIA
                    WHERE ID_ARQUIVO = ?
                """, (id_idw,))

                if not cursor.fetchone():
                    cursor.execute("""
                        INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO, ORIGEM)
                        VALUES (?, ?)
                    """, (id_idw, "ALTERADOS"))

                    print("📄 IDW inserido na fila:", id_idw)

                    conecta_engenharia.commit()

        except Exception as e:
            print("Erro ao inserir na fila:", e)

    def classificar(self, desc_inv, desc_erp):
        def limpar_texto(txt):
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

        def similaridade(a, b):
            from difflib import SequenceMatcher

            return SequenceMatcher(None, a, b).ratio()
        a = limpar_texto(desc_inv)
        b = limpar_texto(desc_erp)

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
        score = similaridade(a, b)

        if score > 0.7:
            return "DUVIDOSO", score

        return "ERRO_GRAVE", score

    def delete_divergencia(self, id_divergencia, numero):
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                DELETE FROM DIVERGENCIAS
                WHERE ID = ?
            """, (id_divergencia,))

            conectados = cursor.rowcount  # 👈 importante

            conecta_engenharia.commit()

            if conectados > 0:
                print("Divergencia DELETE com sucesso!", id_divergencia, numero)
                return True

            return False

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_duplicados_01(self, tipo_arq, nome_base):
        try:
            tem_campo = False
            dados = []

            cursor = conecta_engenharia.cursor()

            if tipo_arq == "IPT" or tipo_arq == "IAM":
                cursor.execute("""
                                        SELECT id, ARQUIVO, NOME_BASE, TIPO_ARQUIVO, CLASSIFICACAO, caminho
                                        FROM arquivos where NOME_BASE = ? AND TIPO_ARQUIVO IN ('IPT', 'IAM')
                                    """, (nome_base,))
                dados = cursor.fetchall()
            elif tipo_arq == "IDW":
                cursor.execute("""
                            SELECT id, ARQUIVO, NOME_BASE, TIPO_ARQUIVO, CLASSIFICACAO, caminho
                            FROM arquivos 
                            where NOME_BASE = ? 
                            AND TIPO_ARQUIVO = 'IDW'
                        """, (nome_base,))
                dados = cursor.fetchall()

            if len(dados) < 2:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_ncm_eng_02(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                SELECT ipt.id_arquivo, ipt.ENGINEER
                                FROM PROPRIEDADES_IPT ipt
                                where ipt.id_arquivo = ?
                                """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        ncm = i[1]
                        if ncm:
                            tem_campo = True
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.ENGINEER  
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        ncm = ii[1]
                        if ncm:
                            tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_sem_idw_03(self, nome_base):
        try:
            tem_campo = False

            cursor_eng = conecta_engenharia.cursor()
            cursor_eng.execute("""
                            SELECT NOME_BASE, NOME_BASE
                            FROM ARQUIVOS
                            where NOME_BASE = ? 
                            and TIPO_ARQUIVO = 'IDW' 
                            """, (nome_base,))
            dados = cursor_eng.fetchall()
            if dados:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_ref_erp_com_nome_base_04(self, id_arquivo, tipo_arq, nome_base):
        try:
            tem_campo = False
            codigo_prod = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        codigo_prod = i[1]
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY  
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        codigo_prod = ii[1]

            if codigo_prod:
                lista_final = []

                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.AUTHORITY = ?
                                    """, (codigo_prod,))
                dados_ipt_cod = cursor_eng.fetchall()
                if dados_ipt_cod:
                    for titi in dados_ipt_cod:
                        lista_final.append(titi)

                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.AUTHORITY = ?
                                    """, (codigo_prod,))
                dados_iam_cod = cursor_eng.fetchall()
                if dados_iam_cod:
                    for tiitii in dados_iam_cod:
                        lista_final.append(tiitii)

                if len(lista_final) > 1:
                    tem_campo = True

                cursor_erp = conecta.cursor()
                cursor_erp.execute("""
                                    SELECT prod.id, prod.descricao, COALESCE(prod.obs, ''), prod.unidade, 
                                    prod.id_versao, prod.KILOSMETRO, prod.conjunto, tip.DESENHO, prod.ID_SERVICO_INTERNO  
                                    FROM produto as prod 
                                    LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id
                                    where prod.codigo = ?
                                    """, (codigo_prod,))
                dados = cursor_erp.fetchall()
                if dados:
                    ref_erp = dados[0][2]
                    ref_erp = re.sub(r"[^\d.]", "", ref_erp)  # remove tudo que não é número ou ponto
                    ref_erp_padrao = re.sub(r"\.+$", "", ref_erp)  # saída: 47.00.014.07
                    if ref_erp_padrao:
                        if ref_erp_padrao == nome_base:
                            tem_campo = True
            else:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_descricao_erp_com_eng_05(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False
            codigo_prod = ""
            descr_eng = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, ipt.DESCRIPTION
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        codigo_prod = i[1]
                        descr_eng = i[2]
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY, iam.DESCRIPTION 
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        codigo_prod = ii[1]
                        descr_eng = ii[2]

            if codigo_prod:
                if descr_eng:
                    cursor_erp = conecta.cursor()
                    cursor_erp.execute("""
                                        SELECT prod.id, prod.descricao, COALESCE(prod.obs, ''), prod.unidade, 
                                        prod.id_versao, prod.KILOSMETRO, prod.conjunto, tip.DESENHO, prod.ID_SERVICO_INTERNO  
                                        FROM produto as prod 
                                        LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id
                                        where prod.codigo = ?
                                        """, (codigo_prod,))
                    dados = cursor_erp.fetchall()
                    if dados:
                        descr_erp = dados[0][1]

                        status, score = self.classificar(descr_eng, descr_erp)

                        if status != "ERRO_GRAVE":
                            tem_campo = True
                else:
                    tem_campo = True
            else:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_comprimento_ipt_06(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.COMPRIMENTO
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        compr = i[1]
                        if compr:
                            tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def confere_se_precisa_descricao_eng_07(self, id_arquivo, caminho, tipo_arq, nome_base):
        try:
            tem_campo = False

            num_desenho = ""
            descr_eng = ""
            est_item = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.PART_NUMBER, ipt.DESCRIPTION, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        num_desenho = i[1]
                        descr_eng = i[2].strip()
                        est_item = i[3].strip()
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.PART_NUMBER, iam.DESCRIPTION, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        num_desenho = ii[1]
                        descr_eng = ii[2].strip()
                        est_item = ii[3].strip()

            if est_item == "FANTASMA":
                print("FANTASMA", descr_eng)
                tem_campo = True
            else:
                if "\\inventor\\biblioteca" in caminho:
                    if descr_eng:
                        tem_campo = True
                else:
                    if nome_base != num_desenho:
                        pass
                    elif descr_eng:
                        tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_descr_mat_eng_08(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cod_eng = ""
            descr__mat_eng = ""
            est_item = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, ipt.REVISION_NUMBER, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
                        descr__mat_eng = i[2].strip()
                        est_item = i[3].strip()
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY, iam.REVISION_NUMBER, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        cod_eng = ii[1].strip()
                        descr__mat_eng = ii[2].strip()
                        est_item = ii[3].strip()

            if est_item == "FANTASMA":
                print("FANTASMA", descr__mat_eng)
                tem_campo = True
            else:
                if descr__mat_eng:
                    tem_campo = True
                else:
                    if cod_eng:
                        cursor_erp = conecta.cursor()
                        cursor_erp.execute("""
                                            SELECT prod.id, prod.descricao, COALESCE(prod.obs, ''), prod.unidade, 
                                            prod.id_versao, prod.KILOSMETRO, prod.conjunto, tip.DESENHO, prod.ID_SERVICO_INTERNO  
                                            FROM produto as prod 
                                            LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id
                                            where prod.codigo = ?
                                            """, (cod_eng,))
                        dados = cursor_erp.fetchall()
                        if dados:
                            conj_erp = dados[0][6]

                            if conj_erp != 10:
                                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_kg_mt_prod_erp_09(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False
            codigo_prod = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.COST_CENTER
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        codigo_prod = i[1]
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.COST_CENTER 
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                                """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        codigo_prod = ii[1]

            if codigo_prod:
                cursor_erp = conecta.cursor()
                cursor_erp.execute("""
                                    SELECT prod.id, prod.KILOSMETRO, prod.conjunto   
                                    FROM produto as prod 
                                    where prod.codigo = ?
                                    """, (codigo_prod,))
                dados = cursor_erp.fetchall()
                if dados:
                    kg_mt = dados[0][1]
                    if kg_mt:
                        tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_codigo_e_fantasma_eng_10(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cod_eng = ""
            est_item = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
                        est_item = i[2].strip()
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        cod_eng = ii[1].strip()
                        est_item = ii[2].strip()

            if est_item == "FANTASMA":
                tem_campo = True
            else:
                if cod_eng:
                    tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_codigo_duplicado_eng_11(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cod_eng = ""
            lista_final = []

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
                        est_item = i[2].strip()
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        cod_eng = ii[1].strip()
                        est_item = ii[2].strip()

            if cod_eng:
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.AUTHORITY = ?
                                    """, (cod_eng,))
                dados_ipt_cod = cursor_eng.fetchall()
                if dados_ipt_cod:
                    for titi in dados_ipt_cod:
                        lista_final.append(titi)

                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.AUTHORITY = ?
                                    """, (cod_eng,))
                dados_iam_cod = cursor_eng.fetchall()
                if dados_iam_cod:
                    for tiitii in dados_iam_cod:
                        lista_final.append(tiitii)


            if len(lista_final) == 1 or not lista_final:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_codigo_eng_12(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cod_eng = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        cod_eng = ii[1].strip()

            if cod_eng:
                cursor_erp = conecta.cursor()
                cursor_erp.execute("""
                                            SELECT prod.id, prod.descricao, prod.conjunto  
                                            FROM produto as prod 
                                            where prod.codigo = ?
                                            """, (cod_eng,))
                produto = cursor_erp.fetchall()

                if produto:
                    tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_codigo_materia_prima_13(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cod_eng = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.COST_CENTER, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.COST_CENTER, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        cod_eng = ii[1].strip()

            if cod_eng:
                cursor_erp = conecta.cursor()
                cursor_erp.execute("""
                                            SELECT prod.id, prod.descricao, prod.conjunto  
                                            FROM produto as prod 
                                            where prod.codigo = ?
                                            """, (cod_eng,))
                produto = cursor_erp.fetchall()

                if produto:
                    tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_descricao_mat_erp_com_eng_14(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False
            codigo_prod = ""
            descr_eng = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.COST_CENTER, ipt.REVISION_NUMBER
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        codigo_prod = i[1]
                        descr_eng = i[2]
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.COST_CENTER, iam.REVISION_NUMBER
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        codigo_prod = ii[1]
                        descr_eng = ii[2]

            if codigo_prod and descr_eng:
                cursor_erp = conecta.cursor()
                cursor_erp.execute("""
                                    SELECT prod.id, prod.descricao, COALESCE(prod.obs, ''), prod.unidade, 
                                    prod.id_versao, prod.KILOSMETRO, prod.conjunto, tip.DESENHO, prod.ID_SERVICO_INTERNO  
                                    FROM produto as prod 
                                    LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id
                                    where prod.codigo = ?
                                    """, (codigo_prod,))
                dados = cursor_erp.fetchall()
                if dados:
                    descr_erp = dados[0][1]

                    status, score = self.classificar(descr_eng, descr_erp)

                    if status != "ERRO_GRAVE":
                        tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_codigo_mat_eng_15(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cod_eng = ""
            estrutura = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.COST_CENTER, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
            else:
                estrutura = self.consulta_estrutura_eng_atual(id_arquivo)

                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.COST_CENTER, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        cod_eng = ii[1].strip()

            if estrutura:
                if len(estrutura) > 1:
                    tem_campo = True

            if cod_eng:
                cursor_erp = conecta.cursor()
                cursor_erp.execute("""
                                            SELECT prod.id, prod.descricao, prod.conjunto  
                                            FROM produto as prod 
                                            where prod.codigo = ?
                                            """, (cod_eng,))
                produto = cursor_erp.fetchall()

                if produto:
                    tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_num_desenho_prop_eng_16(self, id_arquivo, nome_base, tipo_arq):
        try:
            tem_campo = False

            num_desenho = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.PART_NUMBER, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        num_desenho = i[1].strip()
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.PART_NUMBER, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        num_desenho = ii[1].strip()

            if num_desenho == nome_base:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_estruturas_18(self, id_arquivo, obs_div, tipo_arq):
        try:
            id_verificar = 1

            tem_campo = False

            cod_eng = ""

            if id_arquivo == id_verificar:
                print(id_arquivo, obs_div, tipo_arq)

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, COMPRIMENTO, COALESCE(ipt.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
            else:
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.AUTHORITY, COALESCE(iam.STOCK_NUMBER, '')
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        cod_eng = ii[1].strip()

            if cod_eng:
                cursor_erp = conecta.cursor()
                cursor_erp.execute("""
                                        SELECT prod.id, prod.descricao, prod.conjunto
                                        FROM produto as prod
                                        where prod.codigo = ?
                                        """, (cod_eng,))
                produto = cursor_erp.fetchall()

                conj = produto[0][2]

                if conj == 10:
                    estrutura_nova = []
                    erro_estrutura = 0

                    if tipo_arq == "IAM":
                        estrutura_eng = self.consulta_estrutura_eng_atual(id_arquivo)

                        if estrutura_eng:
                            for id_arquivo_f, qtde_f in estrutura_eng:
                                if id_arquivo == id_verificar:
                                    print("INVENTOR", id_arquivo_f, qtde_f)

                                codigo_f = self.consulta_propriedade_ipt_iam(id_arquivo_f, "AUTHORITY")

                                if codigo_f:
                                    dados_f = self.consulta_arquivos(id_arquivo_f)
                                    tipo_arquivo_f = dados_f[0][2]

                                    compr_ipt_f_float = None

                                    if tipo_arquivo_f == "IPT":
                                        cursor_eng = conecta_engenharia.cursor()
                                        sql = f"""
                                                        SELECT ipt.id_arquivo, ipt.COMPRIMENTO
                                                        FROM PROPRIEDADES_IPT ipt
                                                        WHERE ipt.id_arquivo = ?
                                                    """
                                        cursor_eng.execute(sql, (id_arquivo_f,))
                                        dados_ipt = cursor_eng.fetchall()
                                        if dados_ipt:
                                            for i in dados_ipt:
                                                compr = i[1]
                                                if compr:
                                                    compr_ipt_f = compr.strip()
                                                    compr_ipt_f_float = self.extrair_numero(compr_ipt_f)

                                    qtde_calc = self.calcular_qtde_erp(codigo_f, qtde_f, compr_ipt_f_float, id_arquivo_f)

                                    if qtde_calc is None:
                                        erro_estrutura += 1
                                    else:
                                        if id_arquivo == id_verificar:
                                            print(" - ", codigo_f, qtde_calc)
                                        estrutura_nova.append((codigo_f, qtde_calc))

                    else:
                        compr_ipt_float = None

                        cod_mat = self.consulta_propriedade_ipt_iam(id_arquivo, "COST_CENTER")

                        cursor_eng = conecta_engenharia.cursor()
                        sql = f"""
                                SELECT ipt.id_arquivo, ipt.COMPRIMENTO
                                FROM PROPRIEDADES_IPT ipt
                                WHERE ipt.id_arquivo = ?
                            """
                        cursor_eng.execute(sql, (id_arquivo,))
                        dados_ipt = cursor_eng.fetchall()
                        if dados_ipt:
                            for i in dados_ipt:
                                compr = i[1]
                                if compr:
                                    compr_ipt_f = compr.strip()
                                    compr_ipt_float = self.extrair_numero(compr_ipt_f)

                        if cod_mat:
                            qtde_calc = self.calcular_qtde_erp(cod_mat, 1, compr_ipt_float, id_arquivo)

                            if qtde_calc is None:
                                erro_estrutura += 1
                            else:
                                if id_arquivo == id_verificar:
                                    print(" - ", cod_mat, qtde_calc)
                                estrutura_nova = [(cod_mat, qtde_calc)]

                    if not erro_estrutura:
                        if estrutura_nova:

                            resultado = self.atualiza_estrutura_erp(cod_eng, estrutura_nova, id_arquivo, id_verificar)

                            if resultado:
                                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_medida_corte_idw_19(self, id_arquivo, nome_base):
        try:
            tem_campo = False

            cursor_eng = conecta_engenharia.cursor()
            cursor_eng.execute("""
                                SELECT ID_ARQUIVO, ID_ARQUIVO_REFERENCIA
                                FROM PROPRIEDADES_IDW
                                WHERE ID_ARQUIVO = ?
                            """, (id_arquivo,))
            dados_prop_idw = cursor_eng.fetchall()

            if dados_prop_idw:
                id_arquivo_ipt = dados_prop_idw[0][1]

                cod_eng = ""
                compr = ""

                cursor_eng = conecta_engenharia.cursor()
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.AUTHORITY, ipt.COMPRIMENTO
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo_ipt,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
                        compr = i[2].strip()

                if cod_eng:
                    cursor_erp = conecta.cursor()
                    cursor_erp.execute("""
                                        SELECT prod.id, prod.descricao, prod.conjunto  
                                        FROM produto as prod 
                                        where prod.codigo = ?
                                        """, (cod_eng,))
                    produto = cursor_erp.fetchall()

                    if produto:
                        conj = produto[0][2]
                        if conj == 10:
                            tem_cota = self.comparar_cotas_idw(cursor_eng, compr, id_arquivo)

                            if tem_cota:
                                tem_campo = True

                return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_idw_divergente_desenho_20(self, nome_base):
        try:
            tem_campo = False

            dados_arq_idw = self.consulta_arquivos_idw(nome_base)
            if dados_arq_idw:
                if len(dados_arq_idw) == 1:
                    id_arquivo_idw = dados_arq_idw[0][0]

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

                            if arquivo_dentro:
                                nome_dentro, nome_base_dentro, tipo_dentro, classificacao_dentro, caminho_dentro = \
                                arquivo_dentro[0]
                                if nome_base == nome_base_dentro:
                                    tem_campo = True
                        else:
                            tem_campo = True
                    else:
                        tem_campo = True

            else:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_idw_sem_propriedades_21(self, nome_base):
        try:
            tem_campo = False

            dados_arq_idw = self.consulta_arquivos_idw(nome_base)
            if dados_arq_idw:
                if len(dados_arq_idw) == 1:
                    id_arquivo_idw = dados_arq_idw[0][0]

                    cursor_eng = conecta_engenharia.cursor()
                    cursor_eng.execute("""
                                        SELECT ID_ARQUIVO, ID_ARQUIVO_REFERENCIA
                                        FROM PROPRIEDADES_IDW
                                        WHERE ID_ARQUIVO = ?
                                    """, (id_arquivo_idw,))
                    dados_prop_idw = cursor_eng.fetchall()

                    if dados_prop_idw:
                        tem_campo = True
                else:
                    tem_campo = True
            else:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_muitos_caracteres_23(self, id_arquivo, nome_base, obs_div):
        try:
            tem_campo = False

            if "Descrição" in obs_div:
                descricao = self.consulta_propriedade_ipt_iam(id_arquivo, "DESCRIPTION")

                if descricao:
                    if len(descricao) < 31:
                        tem_campo = True
                else:
                    tem_campo = True

            if "Referência" in obs_div:
                ref = f"D {nome_base}"

                if len(ref) < 21:
                    tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_preliminar_erp_24(self, nome_base, id_arquivo):
        try:
            tem_campo = False

            ref = f"D {nome_base}"

            dados = self.consulta_referencia_prod_erp(ref)
            if dados:
                tem_campo = True

            cursor = conecta.cursor()
            cursor.execute("""
                            SELECT ID, OBS, DESCRICAO
                            FROM PRODUTOPRELIMINAR
                            WHERE REFERENCIA = ?
                        """, (ref,))
            tem_preliminar = cursor.fetchall()
            if tem_preliminar:
                tem_campo = True

            descricao = self.consulta_propriedade_ipt_iam(id_arquivo, "DESCRIPTION")

            if not descricao:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_num_desenho_26(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False
            num_desenho = ""

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                    SELECT ipt.id_arquivo, ipt.PART_NUMBER
                                    FROM PROPRIEDADES_IPT ipt
                                    where ipt.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                if dados_ipt:
                    for i in dados_ipt:
                        num_desenho = i[1]
            elif tipo_arq == "IAM":
                cursor_eng.execute("""
                                    SELECT iam.id_arquivo, iam.PART_NUMBER
                                    FROM PROPRIEDADES_IAM iam
                                    where iam.id_arquivo = ?
                                    """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    for ii in dados_iam:
                        num_desenho = ii[1]

            if num_desenho:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_propriedade_inventor_27(self, id_arquivo, obs_div):
        try:
            tem_campo = False

            if obs_div:
                texto, numero = obs_div.split(" - ")

                if texto:
                    resultado = self.consulta_propriedade_ipt_iam(id_arquivo, texto)
                    if resultado:
                        result = resultado[0].strip()
                        if result:
                            tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_estrutura_inventor_28(self, id_arquivo):
        try:
            tem_campo = False

            estrutura_eng = self.consulta_estrutura_eng_atual(id_arquivo)

            if estrutura_eng:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_projeto_29(self, id_arquivo):
        try:
            tem_campo = False

            cursor = conecta_engenharia.cursor()
            cursor.execute(f"SELECT ID, ID_ARQUIVO, QTDE, ID_CLIENTE, SOLICITANTE, OBS, PREVISAO_ENTREGA "
                           f"FROM PROJETO "
                           f"where status = 'A' and ID_ARQUIVO = {id_arquivo};")
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

                        if qtde and id_cliente and solicitante and obs and data_entrega and data_entrega > hoje:
                            cursor = conecta.cursor()
                            cursor.execute(f"select id, razao "
                                           f"from clientes "
                                           f"where id = {id_cliente};")
                            dados_cliente = cursor.fetchall()

                            if dados_cliente:
                                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_codigo_biblioteca_30(self, id_arquivo, tipo_arq):
        try:
            tem_campo = False

            cod_eng = ""

            print(id_arquivo, tipo_arq)

            cursor_eng = conecta_engenharia.cursor()
            if tipo_arq == "IPT":
                cursor_eng.execute("""
                                                SELECT ipt.id_arquivo, ipt.AUTHORITY, COALESCE(ipt.STOCK_NUMBER, '')
                                                FROM PROPRIEDADES_IPT ipt
                                                where ipt.id_arquivo = ?
                                                """, (id_arquivo,))
                dados_ipt = cursor_eng.fetchall()
                print(dados_ipt)
                if dados_ipt:
                    for i in dados_ipt:
                        cod_eng = i[1].strip()
            else:
                cursor_eng.execute("""
                                                SELECT iam.id_arquivo, iam.AUTHORITY, COALESCE(iam.STOCK_NUMBER, '')
                                                FROM PROPRIEDADES_IAM iam
                                                where iam.id_arquivo = ?
                                                """, (id_arquivo,))
                dados_iam = cursor_eng.fetchall()

                if dados_iam:
                    print(dados_iam)
                    for ii in dados_iam:
                        cod_eng = ii[1].strip()

            if cod_eng:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def consulta_propriedade_33(self, id_arquivo):
        try:
            tem_campo = False

            cursor_eng = conecta_engenharia.cursor()
            sql = f"""
                            SELECT id, id_arquivo
                            FROM PROPRIEDADES_IPT
                            WHERE id_arquivo = ?
                        """
            cursor_eng.execute(sql, (id_arquivo,))
            dados_ipt = cursor_eng.fetchall()
            if dados_ipt:
                tem_campo = True

            sql = f"""
                            SELECT id, id_arquivo
                            FROM PROPRIEDADES_IAM
                            WHERE id_arquivo = ?
                        """
            cursor_eng.execute(sql, (id_arquivo,))
            dados_iam = cursor_eng.fetchall()

            if dados_iam:
                tem_campo = True

            return tem_campo

        except Exception as e:
            trata_excecao(e)
            raise

    def processar(self):
        try:
            sql = """
                    SELECT div.id, div.ID_TIPO_DIVERGENCIA, div.ID_ARQUIVO, ARQ.NOME_BASE, 
                    arq.TIPO_ARQUIVO, tip_div.DESCRICAO, div.OBS, arq.CAMINHO, arq.CLASSIFICACAO, div.RESOLVIDO
                    FROM DIVERGENCIAS as div
                    INNER JOIN ARQUIVOS as arq ON div.ID_ARQUIVO = arq.id 
                    INNER JOIN TIPO_DIVERGENCIA as tip_div ON div.ID_TIPO_DIVERGENCIA = tip_div.id 
                    """
            cur = conecta_engenharia.cursor()
            cur.execute(sql)
            dados_divergencias = cur.fetchall()

            tem_campo = False

            if dados_divergencias:
                for i in dados_divergencias:
                    id_div, id_tipo_div, id_arquivo, nome_base, tipo_arq, descr_div, obs_div, caminho, classifica, resolvido = i

                    if id_tipo_div != 1 and id_tipo_div != 11 and id_tipo_div != 21 and id_tipo_div != 19:
                        if not any(item["id"] == id_arquivo for item in self.lista_itens):
                            print("não tem mais na lista!!!", id_arquivo, id_div)
                            print("tipo de divergencia", id_tipo_div)
                            print(caminho)

                            self.delete_divergencia(id_div, id_tipo_div)
                            continue

                    if id_tipo_div == 1:
                        tem_campo = self.consulta_duplicados_01(tipo_arq, nome_base)
                        if tem_campo:
                            self.delete_divergencia(id_div, 1)

                    if id_tipo_div == 19:
                        tem_campo = self.consulta_medida_corte_idw_19(id_arquivo, nome_base)
                        if tem_campo:
                            self.delete_divergencia(id_div, 19)
                    if id_tipo_div == 2:
                        tem_campo = self.consulta_ncm_eng_02(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 2)
                    if id_tipo_div == 3:
                        tem_campo = self.consulta_sem_idw_03(nome_base)
                        if tem_campo:
                            self.delete_divergencia(id_div, 3)
                    if id_tipo_div == 4:
                        tem_campo = self.consulta_ref_erp_com_nome_base_04(id_arquivo, tipo_arq, nome_base)
                        if tem_campo:
                            self.delete_divergencia(id_div, 4)
                    if id_tipo_div == 5:
                        tem_campo = self.consulta_descricao_erp_com_eng_05(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 5)
                    if id_tipo_div == 6:
                        tem_campo = self.consulta_comprimento_ipt_06(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 6)
                    if id_tipo_div == 7:
                        tem_campo = self.confere_se_precisa_descricao_eng_07(id_arquivo, caminho, tipo_arq, nome_base)
                        if tem_campo:
                            self.delete_divergencia(id_div, 7)
                    if id_tipo_div == 8:
                        tem_campo = self.consulta_descr_mat_eng_08(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 8)
                    if id_tipo_div == 9:
                        tem_campo = self.consulta_kg_mt_prod_erp_09(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 9)
                    if id_tipo_div == 10:
                        tem_campo = self.consulta_codigo_e_fantasma_eng_10(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 10)
                    if id_tipo_div == 11:
                        tem_campo = self.consulta_codigo_duplicado_eng_11(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 11)
                    if id_tipo_div == 12:
                        tem_campo = self.consulta_codigo_eng_12(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 12)
                    if id_tipo_div == 13:
                        tem_campo = self.consulta_codigo_materia_prima_13(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 13)
                    if id_tipo_div == 14:
                        tem_campo = self.consulta_descricao_mat_erp_com_eng_14(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 14)
                    if id_tipo_div == 15:
                        tem_campo = self.consulta_codigo_mat_eng_15(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 15)
                    if id_tipo_div == 16:
                        tem_campo = self.consulta_num_desenho_prop_eng_16(id_arquivo, nome_base, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 16)
                    if id_tipo_div == 17:
                        self.envia_email_tipo_diveregencia(id_tipo_div, obs_div)
                    if id_tipo_div == 18:
                        tem_campo = self.consulta_estruturas_18(id_arquivo, obs_div, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 18)
                    if id_tipo_div == 20:
                        tem_campo = self.consulta_idw_divergente_desenho_20(nome_base)
                        if tem_campo:
                            self.delete_divergencia(id_div, 20)
                    if id_tipo_div == 21:
                        tem_campo = self.consulta_idw_sem_propriedades_21(nome_base)
                        if tem_campo:
                            self.delete_divergencia(id_div, 21)
                    if id_tipo_div == 22:
                        self.envia_email_tipo_diveregencia(id_tipo_div, obs_div)

                    if id_tipo_div == 23:
                        tem_campo = self.consulta_muitos_caracteres_23(id_arquivo, nome_base, obs_div)
                        if tem_campo:
                            self.delete_divergencia(id_div, 23)
                    if id_tipo_div == 24:
                        tem_campo = self.consulta_preliminar_erp_24(nome_base, id_arquivo)
                        if tem_campo:
                            self.delete_divergencia(id_div, 24)

                    if id_tipo_div == 25:
                        self.envia_email_tipo_diveregencia(id_tipo_div, obs_div)

                    if id_tipo_div == 26:
                        tem_campo = self.consulta_num_desenho_26(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 26)

                    if id_tipo_div == 27:
                        tem_campo = self.consulta_propriedade_inventor_27(id_arquivo, obs_div)
                        if tem_campo:
                            self.delete_divergencia(id_div, 27)
                    if id_tipo_div == 28:
                        tem_campo = self.consulta_estrutura_inventor_28(id_arquivo)
                        if tem_campo:
                            self.delete_divergencia(id_div, 28)

                    if id_tipo_div == 29:
                        tem_campo = self.consulta_projeto_29(id_arquivo)
                        if tem_campo:
                            self.delete_divergencia(id_div, 29)

                    if id_tipo_div == 30:
                        tem_campo = self.consulta_codigo_biblioteca_30(id_arquivo, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 30)

                    if id_tipo_div == 31:
                        tem_campo = self.consulta_estruturas_18(id_arquivo, obs_div, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 31)

                    if id_tipo_div == 32:
                        tem_campo = self.consulta_estruturas_18(id_arquivo, obs_div, tipo_arq)
                        if tem_campo:
                            self.delete_divergencia(id_div, 32)

                    if id_tipo_div == 33:
                        tem_campo = self.consulta_propriedade_33(id_arquivo)
                        if tem_campo:
                            self.delete_divergencia(id_div, 33)

                    # if resolvido == "S" and not tem_campo:
                    #     dados = (id_tipo_div, id_arquivo, nome_base, descr_div, obs_div, caminho)
                    #     self.envia_email_div_nao_resolvida(dados)
                    #     print(f"⚠️ NÃO FOI RESOLVIDA:", id_tipo_div, id_arquivo, nome_base, descr_div, obs_div, caminho)

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    ConfereDivergencias()