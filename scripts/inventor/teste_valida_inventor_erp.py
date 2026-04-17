from core.banco import conecta, conecta_engenharia
from core.erros import trata_excecao
from core.email_service import dados_email
from core.inventor import padrao_desenho, normalizar_texto
import re
from core.conversores import valores_para_float
from datetime import date
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


class GerarFilaValidacaoERP:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

        self.processar()

    def envia_email_sem_medida_corte(self, caminho, desenho):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA/ERP - MEDIDA DE CORTE DIVERGENTE {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f'{saudacao}\n\nO desenho {desenho} IDW não possui medida de corte conforme propriedade "Comprimento"!\n\n'

            body += f"'{caminho}'\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado SEM MEDIDA DE CORTE")

        except Exception as e:
            trata_excecao(e)
            raise

    def envia_email_sem_codigo_mat_prima(self, caminho, desenho):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ENGENHARIA/ERP - SEM CÓDIGO MATÉRIA PRIMA {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f'{saudacao}\n\nO desenho {desenho} não possui código de matréia-prima!\n\n'

            body += f"'{caminho}'\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado SEM MEDIDA DE CORTE")

        except Exception as e:
            trata_excecao(e)
            raise

    def insert_pre_cadastro(self, dados_produto):
        try:
            descr, decr_compl, ref, um, ncm, qtdezinha_float, fornecedor = dados_produto

            descr_padrao = normalizar_texto(descr)

            cursor = conecta.cursor()
            cursor.execute("""
                                        SELECT ID, OBS, DESCRICAO
                                        FROM PRODUTOPRELIMINAR
                                        WHERE DESCRICAO = ?
                                    """, (descr_padrao,))
            tem_preliminar = cursor.fetchall()

            if not tem_preliminar:
                obs = "CRIADO PELO BOBY"

                sql = """
                        INSERT INTO PRODUTOPRELIMINAR (ID, OBS, DESCRICAO, DESCR_COMPL, 
                        REFERENCIA, UM, NCM, KG_MT, FORNECEDOR) 
                        VALUES (GEN_ID(GEN_PRODUTOPRELIMINAR_ID,1), ?, ?, ?, ?, ?, ?, ?, ?);
                        """
                print(sql)

                #cursor.execute(sql, (obs, descr_padrao, decr_compl, ref, um, ncm, qtdezinha_float, fornecedor))

                #conecta.commit()

                print("Produto preliminar inserido!", descr, ref)

        except Exception as e:
            trata_excecao(e)
            raise

    def insert_propriedade_inventor(self, dados_produto):
        try:
            id_arquivo, nome_prop, valor_prop = dados_produto

            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                            SELECT ID, ID_ARQUIVO, NOME_PROPRIEDADE
                            FROM FILA_LANCA_PROPRIEDADE
                            WHERE ID_ARQUIVO = ? and NOME_PROPRIEDADE = ?
                        """, (id_arquivo, nome_prop))
            tem_na_fila = cursor.fetchall()

            if not tem_na_fila:
                sql = """
                        INSERT INTO FILA_LANCA_PROPRIEDADE (ID, ID_ARQUIVO, NOME_PROPRIEDADE, VALOR_PROPRIEDADE) 
                        VALUES (GEN_ID(GEN_FILA_LANCA_PROPRIEDADE_ID,1), ?, ?, ?);
                        """
                print(sql)
                #cursor.execute(sql, (id_arquivo, nome_prop, valor_prop))

                #conecta_engenharia.commit()

                print("Produto inserido na fila propriedade inventor!", nome_prop, valor_prop)

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

    def tratar_referencia(self, codigo, descricao, obs):
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

    def buscar_toda_estrutura(self, cursor, id_pai):
        try:
            visitados = set()
            fila = [id_pai]

            while fila:
                atual = fila.pop()

                if atual in visitados:
                    continue

                visitados.add(atual)

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
                print(id_pai, atual, filhos)
                fila.extend(filhos)

            return list(visitados)

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
                AND prod.descricao NOT LIKE '%KIT%' 
                AND prod.codigo = '22218'
            """)
            registros = cursor_erp.fetchall()

            print(f"📦 Total pedidos ativos: {len(registros)}")

            for codigo, descricao, obs, conj in registros:

                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs "
                               f"FROM produto as prod "
                               f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                               f"where prod.codigo = {21981};")
                detalhes_pai = cursor.fetchall()
                codigo, descricao, obs = detalhes_pai[0]
                print(detalhes_pai[0])

                ref = self.tratar_referencia(codigo, descricao, obs)

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

                for id_item in ids:
                    if id_item not in ids_unicos:
                        ids_unicos.add(id_item)
                        lista_itens.append((codigo, obs, id_item))

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

            for codigo, ref, id_arquivo in lista_itens:

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
                    **props_unificado
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

    def consulta_codigo_prod_erp(self, codigo):
        try:
            cursor_erp = conecta.cursor()
            cursor_erp.execute("""
                            SELECT id, descricao, COALESCE(obs, ''), unidade, id_versao, KILOSMETRO, conjunto 
                            FROM produto where codigo = ?
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

    def buscar_props(self, dados_arq):
        cursor = conecta_engenharia.cursor()

        if dados_arq["tipo"] == "IAM":
            cursor.execute("""
                SELECT AUTHORITY, DESCRIPTION, COST_CENTER, REVISION_NUMBER, PART_NUMBER, ENGINEER, TOTAL_ITENS  
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
                    "tot_itens": tot_itens
                }
                for cod, desc, cod_mat, desc_mat, part, ncm, tot_itens in dados_iam
            ]

        elif dados_arq["tipo"] == "IPT":
            cursor.execute("""
                SELECT AUTHORITY, DESCRIPTION, COST_CENTER, REVISION_NUMBER, PART_NUMBER, ENGINEER, COMPRIMENTO    
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
                    "compr": compr
                }
                for cod, desc, cod_mat, desc_mat, part, ncm, compr in cursor.fetchall()
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

        print("\n============================")
        print("tratar_resultado")
        for item in lista:
            id_arquivo = item['id']
            tipo_arquivo = item['tipo']
            classificacao = item['classificacao']
            caminho_arquivo = item['caminho']

            nome_base = item['nome_base']

            codigo = str(item.get("codigo") or "").strip()
            descricao = str(item.get("descricao") or "").strip()

            cod_mat = str(item.get("cod_mat") or "").strip()
            desc_mat = str(item.get("desc_mat") or "").strip()

            num_desenho = item["num_desenho"]
            ncm = item["ncm"]

            tot_itens_iam = item["tot_itens"]

            compr_ipt = item["compr"]

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
                if not descricao:
                    print("- BIBLIOTECA: descricao", codigo, descricao, cod_mat, desc_mat, "caminho:", caminho_arquivo)
                    continue
                if not codigo:
                    print("- BIBLIOTECA: codigo", codigo, descricao, cod_mat, desc_mat, id_arquivo, "caminho:", caminho_arquivo)
                    continue

                if not dados_cod:
                    print("- BIBLIOTECA: codigo ERP", codigo, descricao, dados_cod, id_arquivo, "caminho:",
                          caminho_arquivo)
                    continue

                descricao_erp = dados_cod[0][1]

                status, score = self.classificar(descricao, descricao_erp)

                if status == "ERRO_GRAVE":
                    print("- BIBLIOTECA: DESCRIÇÃO DIFERENTES!", codigo, descricao, dados_cod, id_arquivo, "caminho:",
                          caminho_arquivo)
                continue
            else:
                if nome_base != num_desenho:
                    print("- NOSSOS ITENS: NOME DO ARQUIVO DIVERGENTE", nome_base, num_desenho, id_arquivo, "caminho:", caminho_arquivo)
                    continue
                if not descricao:
                    print("- NOSSOS ITENS: SEM DESCRIÇÃO DO PRODUTO!", codigo, descricao, cod_mat, desc_mat, id_arquivo, "caminho:", caminho_arquivo)
                    continue
                if not codigo:
                    ref = f"D {nome_base}"

                    dados_ref = self.consulta_referencia_prod_erp(ref)
                    if not dados_ref:
                        if tipo_arquivo == "IAM" and classificacao == "NOSSO":
                            filhos = self.consulta_estrutura_eng(cursor_eng, id_arquivo)
                            if filhos:
                                print("- NOSSOS ITENS: tem estrutura:", len(filhos))
                            else:
                                print("- NOSSOS ITENS: IAM SEM ESTRUTURA")
                        if tipo_arquivo == "IPT" and classificacao == "NOSSO":
                            if ncm:
                                dados_produto_nosso = (descricao, "", ref, "UN", ncm, 0, "")
                                self.insert_pre_cadastro(dados_produto_nosso)
                            else:
                                print("SEM NCM", ncm, "caminho:", caminho_arquivo)
                                continue
                    elif len(dados_ref) > 1:
                        print("- NOSSOS ITENS: DESENHO ENCONTRADO EM MAIS DPRODUTOS:", dados_ref)
                    else:
                        cod_erp_ref = dados_ref[0][1]
                        dados_inventor = (id_arquivo, "Authority", cod_erp_ref)
                        self.insert_propriedade_inventor(dados_inventor)

                if not dados_cod:
                    print("- NOSSOS ITENS: codigo ERP", codigo, descricao, dados_cod, id_arquivo, "caminho:", caminho_arquivo)
                    continue

                descricao_erp = dados_cod[0][1]
                status, score = self.classificar(descricao, descricao_erp)
                if status == "ERRO_GRAVE":
                    print("- NOSSOS ITENS: DESCRIÇÃO DIFERENTES!", codigo, descricao, dados_cod, id_arquivo, "caminho:",
                          caminho_arquivo)
                    continue

                if not desc_mat:
                    if tipo_arquivo == "IAM" and tot_itens_iam == 1:
                        print("- NOSSOS ITENS: NOSSO SEM DESCRIÇÃO materia-prima! IAM 1 ITEM")
                        print("           Cód:", codigo, "Descr:", descricao, "CodMat:", cod_mat, "DescrMat:", desc_mat)
                        print("caminho:", caminho_arquivo, id_arquivo)
                        continue
                    if tipo_arquivo == "IPT":
                        print("- NOSSOS ITENS: NOSSO SEM DESCRIÇÃO materia-prima! IPT", codigo, descricao, cod_mat, desc_mat,
                              id_arquivo, "caminho:", caminho_arquivo)
                        continue

                conj = dados_cod[0][6]
                if conj == 10:
                    if not cod_mat:
                        if tipo_arquivo == "IAM" and tot_itens_iam == 1:
                            filhos = self.consulta_estrutura_eng(cursor_eng, id_arquivo)
                            if filhos:
                                arq_filho = filhos[0][1]
                                dados_f = self.consulta_arquivos(arq_filho)
                                if dados_f:
                                    classif_f = dados_f[0][3]
                                    if classif_f == "TERCEIROS":
                                        self.envia_email_sem_codigo_mat_prima()
                            continue
                        if tipo_arquivo == "IPT":
                            print("- NOSSOS ITENS: SEM CODIGO MATERIA PRIMA IPT")
                            print("           Cód:", codigo, "Descr:", descricao, "CodMat:", cod_mat, "DescrMat:", desc_mat)
                            print("caminho:", caminho_arquivo)
                            continue

                    if (tipo_arquivo == "IAM" and tot_itens_iam == 1) or (tipo_arquivo == "IPT"):
                        if not dados_cod_mat:
                            print("- NOSSOS ITENS: codigo Matéria-prima ERP", codigo, descricao, cod_mat, desc_mat, id_arquivo, "caminho:",
                                  caminho_arquivo)
                            continue

                        descricao_mat_erp = dados_cod_mat[0][1]
                        status, score = self.classificar(desc_mat, descricao_mat_erp)
                        if status == "ERRO_GRAVE":
                            print("- NOSSOS ITENS: DESCRIÇÃO MATERIA-PRIMA DIFERENTES!", codigo, descricao, cod_mat, desc_mat, id_arquivo, "caminho:",
                                  caminho_arquivo)
                            continue
                        else:
                            um = dados_cod_mat[0][3]
                            if um == "KG" or um == "MT" or um == "MM":
                                kg_mt = dados_cod_mat[0][5]
                                if not compr_ipt:
                                    print("- NOSSOS ITENS: DEVE ter comprimento", descricao_mat_erp, desc_mat, compr_ipt)
                                    continue

                                if not kg_mt:
                                    print("- NOSSOS ITENS: DEVE ter kg/mt", descricao_mat_erp, desc_mat, compr_ipt)
                                    continue
            lista_validos.append(item)

        if lista_validos:
            self.tratar_estruturas(lista_validos)

    def consulta_estrutura_eng_atual(self, cursor, id_pai):
        try:
            cursor.execute("""
                SELECT ID_FILHO, QTDE
                FROM ESTRUTURA
                WHERE ID_PAI=?
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

        print("\n============================")
        print("tratar_estruturas")

        for item in lista:
            id_arquivo = item['id']
            tipo_arquivo = item['tipo']

            codigo = str(item.get("codigo") or "").strip()
            descricao = str(item.get("descricao") or "").strip()

            cod_mat = str(item.get("cod_mat") or "").strip()

            compr_ipt = item.get("compr")
            compr_ipt_float = None

            if tipo_arquivo == "IPT":
                compr_ipt_float = self.extrair_numero(compr_ipt)

            if tipo_arquivo == "IAM":
                estrutura_eng = self.consulta_estrutura_eng_atual(cursor_eng, id_arquivo)

                if not estrutura_eng:
                    print("IAM sem estrutura:", codigo)
                    continue

                estrutura_nova = []
                erro_estrutura = False

                for id_arquivo_f, qtde in estrutura_eng:
                    item_f = next((i for i in lista if i["id"] == id_arquivo_f), None)

                    if not item_f:
                        print("IAM Filho não encontrado na lista:", id_arquivo_f)
                        erro_estrutura = True
                        break

                    codigo_f = str(item_f.get("codigo") or "").strip()

                    if not codigo_f:
                        print("IAM Filho sem código:", item_f)
                        erro_estrutura = True
                        break

                    qtde_calc = self.calcular_qtde_erp(codigo_f, qtde, compr_ipt_float)

                    if qtde_calc is None:
                        print("IAM Erro ao calcular quantidade:", codigo_f)
                        erro_estrutura = True
                        break

                    estrutura_nova.append((codigo_f, qtde_calc))

                if erro_estrutura:
                    print("IAM Estrutura ignorada por erro:", codigo)
                    continue

            else:
                if not cod_mat:
                    print("IPT sem matéria-prima:", codigo)
                    continue

                dados_cod_mat = self.consulta_codigo_prod_erp(cod_mat)
                if not dados_cod_mat:
                    print("IPT Matéria-prima não existe no ERP:", cod_mat)
                    continue

                qtde_base = compr_ipt_float if compr_ipt_float else 1
                qtde_calc = self.calcular_qtde_erp(cod_mat, qtde_base, compr_ipt_float)

                if qtde_calc is None:
                    print("IPT Erro ao calcular quantidade IPT:", cod_mat)
                    continue

                estrutura_nova = [(cod_mat, qtde_calc)]

            if estrutura_nova:
                self.atualiza_estrutura_erp(codigo, item, estrutura_nova)

        return None

    def calcular_qtde_erp(self, cod_prod, qtde_eng, compr_ipt=None):
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

            qtde_f = self.arredondar_qtde(valores_para_float(kg_mt) * (compr_ipt / 1000), 2)
            return qtde_f

        if unidade == "MT":
            if not compr_ipt:
                print("Falta comprimento para unidade:", unidade, cod_prod)
                return None

            compr_m = self.arredondar_qtde((valores_para_float(compr_ipt / 1000)), 2)
            return compr_m

        if unidade == "MM":
            if not compr_ipt:
                print("Falta comprimento para unidade:", unidade, cod_prod)
                return None

            return compr_ipt

        # 🔴 qualquer outra unidade não tratada
        print("Unidade não tratada:", unidade, cod_prod)
        return None

    def extrair_numero(self, valor):
        if valor is None:
            return None

        s = str(valor).lower().strip()

        # remove tudo que não é número ou ponto
        s = re.sub(r"[^\d.]", "", s)

        if not s:
            return None

        try:
            return valores_para_float(s)
        except:
            return None

    def arredondar_qtde(self, qtde, casas_decimais):
        qtde_final = round(qtde, casas_decimais)
        return qtde_final

    def atualiza_estrutura_erp(self, cod_prod, item, estrutura_nova):
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
                #print("Estrutura ERP", estrutura_erp_set)
                #print("Estrutura Engenharia", estrutura_nova_set)
                if id_versao_atual != id_estrutura:
                    # cursor.execute("""
                    #     UPDATE produto
                    #     SET id_versao = ?
                    #     WHERE id = ?
                    # """, (id_estrutura, id_prod))
                    #
                    # conecta.commit()

                    print("Estrutura já existe Ativando:", cod_prod)

                self.tratar_idw(item)
                return

        # 🔥 NÃO EXISTE → cria nova versão
        print("Criando nova versão:", cod_prod)

        # cursor.execute("""
        #     SELECT COALESCE(MAX(num_versao), 0)
        #     FROM estrutura
        #     WHERE id_produto = ?
        # """, (id_prod,))
        # num_versao = cursor.fetchone()[0] + 1
        #
        # cursor.execute("""
        #     INSERT INTO estrutura (id, id_produto, num_versao, data_versao, obs)
        #     VALUES (GEN_ID(GEN_ESTRUTURA_ID,1), ?, ?, CURRENT_DATE, ?)
        #     RETURNING id
        # """, (id_prod, num_versao, "ATUALIZADO AUTOMATICAMENTE"))
        #
        # id_estrutura = cursor.fetchone()[0]
        #
        # cursor.execute("""
        #     UPDATE produto
        #     SET id_versao = ?
        #     WHERE id = ?
        # """, (id_estrutura, id_prod))
        #
        # for cod_filho, qtde in estrutura_nova:
        #     cursor.execute("SELECT id FROM produto WHERE codigo = ?", (cod_filho,))
        #     row = cursor.fetchone()
        #
        #     if not row:
        #         continue
        #
        #     cursor.execute("""
        #         INSERT INTO estrutura_produto
        #         (id, id_estrutura, id_prod_filho, quantidade)
        #         VALUES (GEN_ID(GEN_ESTRUTURA_PRODUTO_ID,1), ?, ?, ?)
        #     """, (id_estrutura, row[0], qtde))
        #
        # conecta.commit()

        print("Nova estrutura criada:", cod_prod)

    def tratar_idw(self, item):
        id_arquivo = item['id']
        caminho_arquivo = item['caminho']
        nome_base = item['nome_base']
        tipo_arquivo = item['tipo']

        cursor_eng = conecta_engenharia.cursor()
        cursor_eng.execute("""
            SELECT ID_ARQUIVO, ID_ARQUIVO_REFERENCIA
            FROM PROPRIEDADES_IDW
            WHERE ID_ARQUIVO_REFERENCIA = ?
        """, (id_arquivo,))
        dados_idw = cursor_eng.fetchall()

        if dados_idw:
            if len(dados_idw) == 1:
                id_arq_idw = dados_idw[0][0]
                arquivo_idw = self.consulta_arquivos(id_arq_idw)

                if not arquivo_idw:
                    print("tratar_idw - NÃO TEM ARQUIVO DE REFERENCIA")
                else:
                    nome_idw, nome_base_idw, tipo_idw, classificacao_idw, caminho_idw = arquivo_idw[0]
                    if nome_base == nome_base_idw:
                        if tipo_arquivo == "IPT":
                            tem_cota = self.comparar_cotas_idw(cursor_eng, item, id_arq_idw)
                            if tem_cota:
                                caminho_pdf = rf"\\Publico\C\OP\Projetos\{nome_base}.pdf"

                                if not os.path.exists(caminho_pdf):
                                    self.inserir_pdf_fila(id_arq_idw, caminho_idw)
                            else:
                                self.envia_email_sem_medida_corte(caminho_idw, nome_base_idw)
                        else:
                            caminho_pdf = rf"\\Publico\C\OP\Projetos\{nome_base}.pdf"

                            if not os.path.exists(caminho_pdf):
                                self.inserir_pdf_fila(id_arq_idw, caminho_idw)
                    else:
                        print("tratar_idw - NUMERO DE DESENHO DIFERENTE")
            else:
                print("tratar_idw - DESENHOS DUPLICADOS")
        else:
            print("tratar_idw - sem idw", caminho_arquivo)

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
                        for id_arq, cota in cotas_idw:
                            if cota == compr_ipt_float - 2.0:
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

                # cursor.execute(sql, (id_arquivo,)) # ✅ AQUI TAMBÉM
                # conecta_engenharia.commit()

                print("Produto inserido na fila de PDF!", caminho_arquivo)

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    GerarFilaValidacaoERP()