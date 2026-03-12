import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
from dados_email import email_user, password
import os
import traceback
import inspect
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders


class ConferirPreNF:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.manipula_comeco()

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

    def montar_html_divergencia(self, numero_nf, erros):
        try:
            total_erros = len(erros)

            lista_erros_html = ""
            for erro in erros:
                lista_erros_html += f"""
                <tr>
                    <td style="padding:8px;border-bottom:1px solid #ddd;">
                        {erro}
                    </td>
                </tr>
                """

            html = f"""
            <html>
            <body style="font-family: Arial, sans-serif; background-color:#f4f6f8; padding:20px;">
                <div style="background:white; padding:20px; border-radius:8px; max-width:800px; margin:auto;">
    
                    <h2 style="color:#c62828;">🚨 Divergências encontradas na NF {numero_nf}</h2>
    
                    <p><strong>Total de ocorrências:</strong> {total_erros}</p>
    
                    <table style="width:100%; border-collapse:collapse; margin-top:15px;">
                        <thead>
                            <tr>
                                <th style="text-align:left; padding:10px; background:#eeeeee;">
                                    Detalhamento das Divergências
                                </th>
                            </tr>
                        </thead>
                        <tbody>
                            {lista_erros_html}
                        </tbody>
                    </table>
    
                    <br>
    
                    <p style="font-size:13px; color:#555;">
                        Este e-mail foi gerado automaticamente pelo sistema de conferência de Notas Fiscais.<br>
                        Favor verificar as inconsistências acima antes de realizar o lançamento definitivo.
                    </p>
    
                    <hr>
    
                    <p style="font-size:12px; color:#888;">
                        Suzuki Máquinas Ltda<br>
                        Fone (51) 3561.2583 / (51) 3170.0965
                    </p>
    
                </div>
            </body>
            </html>
            """

            return html

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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

            return saudacao, msg_final, to

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def localizar_arquivos_nf(self, numero_nf, data_emissao):
        try:
            if isinstance(data_emissao, str):
                data_emissao = datetime.fromisoformat(data_emissao)

            ano = data_emissao.strftime("%Y")
            mes = data_emissao.strftime("%m")
            dia = data_emissao.strftime("%d")

            base_path = r"\\Publico\g\Pasta Scanner Backup\xml"

            pasta = os.path.join(base_path, ano, mes, dia)

            caminho_xml = os.path.join(pasta, f"{numero_nf}.xml")
            caminho_pdf = os.path.join(pasta, f"{numero_nf}.pdf")

            arquivos = {
                "xml": caminho_xml if os.path.exists(caminho_xml) else None,
                "pdf": caminho_pdf if os.path.exists(caminho_pdf) else None
            }

            return arquivos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return {"xml": None, "pdf": None}

    def envia_email_divergencia_nf(self, numero_nf, data_emissao, erros):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'Divergência na NF {numero_nf}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['To'] = ", ".join(to)
            msg['Subject'] = subject

            body = f"{saudacao}\n\n"
            body += f"Foram encontradas divergências na NF {numero_nf}:\n\n"

            for erro in erros:
                body += f"- {erro}\n"

            body += "\n" + msg_final

            html_body = self.montar_html_divergencia(numero_nf, erros)
            msg.attach(MIMEText(html_body, 'html'))

            # 🔎 Localiza arquivos no servidor
            arquivos = self.localizar_arquivos_nf(numero_nf, data_emissao)

            # 📎 Anexa XML e PDF se existirem
            for tipo, caminho in arquivos.items():
                if caminho:
                    with open(caminho, 'rb') as attachment:
                        part = MIMEBase('application', "octet-stream")
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)

                        nome_arquivo = os.path.basename(caminho)

                        part.add_header(
                            'Content-Disposition',
                            'attachment',
                            filename=Header(nome_arquivo, 'utf-8').encode()
                        )

                        msg.attach(part)

            # 🚀 Envia email
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, msg.as_string())
            server.quit()

            print("Email de divergência enviado com sucesso!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pre_nf_pendentes(self):
        try:
            cursor_nf = conecta.cursor()
            cursor_nf.execute("""
                                SELECT ID, NUMERO_NF, ID_FORNECEDOR
                                FROM PRE_NF_ENTRADA
                                WHERE STATUS IS NULL 
                                OR STATUS = 'PENDENTE' 
                                OR STATUS = 'DIVERGENCIA'
                            """)

            nfs = cursor_nf.fetchall()

            return nfs

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_se_nf_com_oc(self, cfop_item, qtde_cfop_dif, nf_tem_erro):
        try:
            nf_com_oc = False

            cursor_natop = conecta.cursor()
            cursor_natop.execute("""
                                SELECT GERA_OC
                                FROM NATOP
                                WHERE CFOP = ?
                                """, (cfop_item,))

            regra = cursor_natop.fetchone()

            if not regra:
                msg = f"⚠ CFOP {cfop_item} não cadastrado na NATOP."
                nf_tem_erro.append(msg)

            gera_oc = regra[0]

            if gera_oc == 'S':
                nf_com_oc = True
            else:
                qtde_cfop_dif += 1

            return nf_tem_erro, qtde_cfop_dif, nf_com_oc

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_prod_forn_siger(self, id_prod_forn, cod_forn, descr_forn, nf_tem_erro):
        try:
            cursor_vinc = conecta.cursor()
            cursor_vinc.execute("""
            SELECT ID_PRODUTO_SIGER
            FROM PRE_PRODUTO_FORNECEDOR
            WHERE id = ?
            """, (id_prod_forn,))
            vinculo = cursor_vinc.fetchone()

            if not vinculo[0]:
                msg = f"❌ Item Fornecedor {cod_forn} - {descr_forn} sem vínculo com produto Siger."
                nf_tem_erro.append(msg)

            return vinculo, nf_tem_erro

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_oc_e_movimento(self, id_prod_siger, id_prod_nf, dados_nf, dados_prod_nf, nf_tem_erro):
        try:
            id_nf, numero_nf, id_fornecedor = dados_nf
            qtde_nf, valor_nf, ipi_nf = dados_prod_nf

            cur = conecta.cursor()
            cur.execute(f"SELECT codigo, descricao, COALESCE(obs, ''), unidade "
                        f"FROM produto where id = {id_prod_siger};")
            detalhes_produto = cur.fetchall()
            cod_prod, descricao, referencia, unidade = detalhes_produto[0]

            cursor_oc = conecta.cursor()
            cursor_oc.execute("""
            SELECT oc.numero, prodoc.ID, prodoc.QUANTIDADE, prodoc.UNITARIO, prodoc.IPI
            FROM ordemcompra as oc
            INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre
            WHERE oc.entradasaida = 'E' 
            and oc.STATUS = 'A' 
            and prodoc.PRODUTO = ?
            """, (id_prod_siger,))
            oc = cursor_oc.fetchall()

            if not oc:
                cursor = conecta.cursor()
                cursor.execute("""
                SELECT oc.numero, prodoc.ID, ent.QUANTIDADE, ent.MOVIMENTACAO, ent.FORNECEDOR, 
                ent.NOTA, ent.NATUREZA, ent.ORDEMCOMPRA, ent.CODIGO
                FROM ENTRADAPROD as ent
                INNER JOIN ordemcompra as oc ON ent.ORDEMCOMPRA = oc.id
                INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre
                WHERE ent.NOTA = ? 
                and ent.FORNECEDOR = ? 
                and ent.PRODUTO = ?
                """, (numero_nf, id_fornecedor, id_prod_siger,))
                dados_nf_lancada = cursor.fetchone()

                if not dados_nf_lancada:
                    msg = f"❌ Produto {cod_prod} - {descricao} - {referencia} não encontrado na OC."
                    nf_tem_erro.append(msg)
                else:
                    id_prod_oc = dados_nf_lancada[1]
                    qtde_lancada = dados_nf_lancada[2]

                    cursor_check = conecta.cursor()
                    cursor_check.execute("""
                    SELECT 1
                    FROM PRE_NF_OC_VINCULO
                    WHERE ID_PRODUTO_NF = ?
                    AND ID_PRODUTO_OC = ?
                    """, (id_prod_nf, id_prod_oc))

                    ja_existe = cursor_check.fetchone()

                    if not ja_existe:
                        cursor_insert = conecta.cursor()
                        cursor_insert.execute("""
                        INSERT INTO PRE_NF_OC_VINCULO
                        (id, ID_NF_PRE, ID_PRODUTO_NF, ID_PRODUTO_OC, QTDE_FORN, QTDE_SIGER)
                        VALUES (GEN_ID(GEN_PRE_NF_OC_VINCULO_ID,1), ?, ?, ?, ?, ?)
                        """, (id_nf, id_prod_nf, id_prod_oc, qtde_nf, qtde_lancada ))

                        conecta.commit()

            elif len(oc) > 1:
                msg = f"O MESMO ITEM REPETIDO NA OC!"
                nf_tem_erro.append(msg)
            else:
                num_oc, id_prod_oc, qtde_oc, valor_oc, ipi_oc = oc[0]

                if qtde_nf != qtde_oc:
                    msg = f"❌ Divergência de quantidade no item {cod_prod} - {descricao} - {referencia} na OC {num_oc}"
                    nf_tem_erro.append(msg)

                if valor_nf != valor_oc:
                    msg = f"❌ Divergência de valor no item {cod_prod} - {descricao} - {referencia} na OC {num_oc}"
                    nf_tem_erro.append(msg)

                if ipi_nf != ipi_oc:
                    msg = f"❌ Divergência de IPI no item {cod_prod} - {descricao} - {referencia} na OC {num_oc}"
                    nf_tem_erro.append(msg)

                cursor_check = conecta.cursor()
                cursor_check.execute("""
                                    SELECT 1
                                    FROM PRE_NF_OC_VINCULO
                                    WHERE ID_PRODUTO_NF = ?
                                    AND ID_PRODUTO_OC = ?
                                    """, (id_prod_nf, id_prod_oc))

                ja_existe = cursor_check.fetchone()

                if not ja_existe:
                    cursor_insert = conecta.cursor()
                    cursor_insert.execute("""
                                        INSERT INTO PRE_NF_OC_VINCULO
                                        (id, ID_NF_PRE, ID_PRODUTO_NF, ID_PRODUTO_OC, QTDE_FORN, QTDE_SIGER)
                                        VALUES (GEN_ID(GEN_PRE_NF_OC_VINCULO_ID,1), ?, ?, ?, ?, ?)
                                        """, (id_nf, id_prod_nf, id_prod_oc, qtde_nf, qtde_oc))

                    conecta.commit()

            return nf_tem_erro

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def confere_produtos_pre_nf(self, dados_nf, nf_tem_erro):
        try:
            id_nf, numero_nf, id_fornecedor = dados_nf

            cursor_itens = conecta.cursor()
            cursor_itens.execute("""
            SELECT nf_prod.ID, nf_prod.ID_PRODUTO_FORN, nf_prod.CFOP, 
            nf_prod.QTDE, nf_prod.UNIT, nf_prod.IPI, 
            prod_forn.CODIGO_FORNECEDOR, prod_forn.DESCRICAO
            FROM PRE_NF_ENTRADA_PRODUTOS as nf_prod
            INNER JOIN PRE_PRODUTO_FORNECEDOR as prod_forn ON nf_prod.ID_PRODUTO_FORN = prod_forn.id
            WHERE nf_prod.ID_NF_PRE = ?
            """, (id_nf,))
            itens = cursor_itens.fetchall()

            qtde_produtos_nf = len(itens)
            qtde_cfop_dif = 0

            if itens:
                for item in itens:
                    id_prod_nf, id_prod_forn, cfop, qtde_nf, valor_nf, ipi_nf, cod_forn, descr_forn = item

                    nf_tem_erro, qtde_cfop_dif, nf_com_oc = self.verifica_se_nf_com_oc(cfop, qtde_cfop_dif, nf_tem_erro)

                    if nf_com_oc:
                        dados_produto_s, nf_tem_erro = self.verifica_prod_forn_siger(id_prod_forn, cod_forn, descr_forn, nf_tem_erro)

                        if dados_produto_s[0]:
                            id_prod_siger = dados_produto_s[0]

                            dados_prod_nf = (qtde_nf, valor_nf, ipi_nf)

                            nf_tem_erro = self.verifica_oc_e_movimento(id_prod_siger, id_prod_nf, dados_nf, dados_prod_nf, nf_tem_erro)

            return nf_tem_erro, qtde_produtos_nf, qtde_cfop_dif

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_divergencia(self, id_nf, numero_nf, nf_tem_erro):
        try:
            # 🔥 buscar data da NF para localizar XML/PDF
            cursor_data = conecta.cursor()
            cursor_data.execute("""
                    SELECT DATA_EMISSAO
                    FROM PRE_NF_ENTRADA
                    WHERE ID = ?
                """, (id_nf,))
            data_emissao = cursor_data.fetchone()[0]

            # 🔥 envia email
            self.envia_email_divergencia_nf(
                numero_nf=numero_nf,
                data_emissao=data_emissao,
                erros=nf_tem_erro
            )

            status = "DIVERGENCIA"

            cursor = conecta.cursor()
            cursor.execute("UPDATE PRE_NF_ENTRADA "
                           "SET STATUS = ? "
                           "WHERE id = ?",
                           (status, id_nf,))
            conecta.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_comeco(self):
        try:
            nfs = self.pre_nf_pendentes()

            if nfs:
                for nf in nfs:
                    id_nf, numero_nf, id_fornecedor = nf
                    print(f"\n🔎 Conferindo NF {numero_nf}")

                    nf_tem_erro = []

                    dados_nf = (id_nf, numero_nf, id_fornecedor)

                    nf_tem_erro, qtde_produtos_nf, qtde_cfop_dif = self.confere_produtos_pre_nf(dados_nf, nf_tem_erro)

                    # =========================
                    # ATUALIZAR STATUS NF
                    # =========================

                    if nf_tem_erro:
                        self.manipula_divergencia(id_nf, numero_nf, nf_tem_erro)
                    elif qtde_produtos_nf == qtde_cfop_dif:
                        status = "NF DIFERENTE"

                        cursor = conecta.cursor()
                        cursor.execute("UPDATE PRE_NF_ENTRADA "
                                       "SET STATUS = ? "
                                       "WHERE id = ?",
                                       (status, id_nf,))
                        conecta.commit()

                    else:
                        self.consulta_vinculos(id_nf, numero_nf)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_vinculos(self, id_nf, numero_nf):
        try:
            cursor_itens = conecta.cursor()
            cursor_itens.execute("""
                SELECT ID, ID_PRODUTO_FORN, CFOP, QTDE, UNIT, IPI
                FROM PRE_NF_ENTRADA_PRODUTOS
                WHERE ID_NF_PRE = ?
            """, (id_nf,))

            itens = cursor_itens.fetchall()

            if not itens:
                return "SEM_ITENS"

            tem_erro = []

            for item in itens:
                id_item, id_prod_forn, cfop, qtde_nf, valor_nf, ipi_nf = item

                # 🔥 Se for retorno de industrialização, não confere vínculo
                if cfop == '5902':
                    continue

                cursor = conecta.cursor()
                cursor.execute("""
                    SELECT 
                        COALESCE(SUM(QTDE_FORN), 0),
                        COALESCE(SUM(QTDE_SIGER), 0)
                    FROM PRE_NF_OC_VINCULO
                    WHERE ID_PRODUTO_NF = ?
                """, (id_item,))

                qtde_vinc_forn, qtde_vinc_int = cursor.fetchone()

                if qtde_vinc_forn != qtde_nf:
                    msg = f"❌ ERRO: Produto {id_prod_forn} quantidade diferente da NF"
                    tem_erro.append(msg)

            # 🔥 Resultado final da NF
            if tem_erro:
                self.manipula_divergencia(id_nf, numero_nf, tem_erro)
            else:
                status = "VINCULADA"

                cursor = conecta.cursor()
                cursor.execute("UPDATE PRE_NF_ENTRADA "
                               "SET STATUS = ? "
                               "WHERE id = ?",
                               (status, id_nf,))
                conecta.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return "ERRO"



chama_classe = ConferirPreNF()