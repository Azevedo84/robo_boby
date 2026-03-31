import sys
from banco_dados.controle_erros import grava_erro_banco
import traceback
import inspect
import os
import win32com.client
from banco_dados.conexao import conecta_engenharia
from comandos.inventor import padrao_desenho
from datetime import datetime
from multiprocessing import Process, freeze_support
import time

from dados_email import email_user, password
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


class WorkerFilaConferencia:
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
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            return None

    def envia_email_desenho_sem_vinculo(self, texto, desenho):
        try:
            saudacao, msg_final, to = self.dados_email()

            subject = f'ENGENHARIA (FILA) - DESENHO SEM VÍNCULO {desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nO desenho {desenho} tem problemas com vínculos dos arquivos!\n\n"

            for i in texto:
                body += f"{i}\n\n"
            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            server.quit()

            print("email enviado SEM VINCULO")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_comeco(self):
        try:
            freeze_support()
            print("🚀 Worker de fila iniciado")
            self.worker_fila()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def worker_fila(self):
        try:
            cursor = conecta_engenharia.cursor()

            cursor.execute("""
                    SELECT ID_ARQUIVO
                    FROM FILA_CONFERENCIA
                    ORDER BY ID
                """)

            fila = [row[0] for row in cursor.fetchall()]

            print(f"\n📦 Total na fila: {len(fila)}")

            for id_arquivo in fila:

                print(f"\n🚀 Processando ID: {id_arquivo}")

                sucesso = self.processar_com_timeout(id_arquivo)

                if sucesso:
                    cursor.execute("""
                            DELETE FROM FILA_CONFERENCIA
                            WHERE ID_ARQUIVO=?
                        """, (id_arquivo,))

                    conecta_engenharia.commit()

                    print(f"🗑️ Removido da fila: {id_arquivo}")

                else:
                    print(f"❌ Mantido na fila (erro): {id_arquivo}")
                    cursor.execute("""
                                SELECT ID, NOME_BASE, CAMINHO
                                FROM ARQUIVOS
                                WHERE ID = ?
                            """, (id_arquivo,))

                    idw_result = cursor.fetchall()
                    desenho = idw_result[0][1]
                    self.envia_email_desenho_sem_vinculo(idw_result, desenho)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def matar_inventor(self):
        try:
            os.system("taskkill /f /im Inventor.exe >nul 2>&1")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def processar_com_timeout(self, id_arquivo):
        try:
            p = Process(target=self.processar_arquivo_unitario, args=(id_arquivo,))
            p.start()

            tempo_max = 60
            inicio = time.time()

            while p.is_alive():
                if time.time() - inicio > tempo_max:
                    print(f"⏱️ TIMEOUT ID {id_arquivo}")

                    try:
                        p.terminate()
                        p.join(5)
                    except:
                        pass

                    self.matar_inventor()
                    return False

                time.sleep(1)

            p.join()

            if p.exitcode == 0:
                return True
            else:
                self.matar_inventor()
                return False

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def processar_arquivo_unitario(self, id_arquivo):
        import pythoncom

        pythoncom.CoInitialize()  # 🔥 ESSENCIAL

        try:
            sucesso = False
            cursor = conecta_engenharia.cursor()

            inventor = None
            doc = None

            try:
                inventor = win32com.client.Dispatch("Inventor.Application")
                inventor.Visible = False
                inventor.SilentOperation = True

                caminho = self.consulta_caminho_arquivo(cursor, id_arquivo)

                print(f"🔍 Abrindo: {caminho}")

                doc = inventor.Documents.Open(caminho, False)

                self.processar_arquivo(cursor, id_arquivo, caminho, doc)

                sucesso = True

            except Exception as e:
                print(f"❌ ERRO REAL ID {id_arquivo}: {str(e)}")

            finally:
                if doc:
                    try:
                        doc.Close(True)
                    except:
                        pass

                if inventor:
                    try:
                        inventor.Quit()
                    except:
                        pass

            if sucesso:
                conecta_engenharia.commit()
                sys.exit(0)
            else:
                conecta_engenharia.rollback()
                sys.exit(1)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)
            sys.exit(1)

        finally:
            pythoncom.CoUninitialize()  # 🔥 LIMPA COM

    def consulta_caminho_arquivo(self, cursor, id_arquivo):
        try:
            cursor.execute("""
                SELECT CAMINHO FROM ARQUIVOS WHERE ID=?
            """, (id_arquivo,))
            return cursor.fetchone()[0]

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def processar_arquivo(self, cursor, id_arquivo, caminho, doc):
        try:
            print(f"\n📂 Processando: {caminho}")

            props = self.consulta_propriedades_inventor(doc)

            ext = caminho.lower()

            if ext.endswith(".iam"):
                self.processar_iam(cursor, doc, id_arquivo, props)

            elif ext.endswith(".ipt"):
                self.salvar_propriedades_ipt(cursor, id_arquivo, props)

            elif ext.endswith(".idw"):
                self.processar_idw(cursor, doc, id_arquivo)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_propriedades_inventor(self, doc):
        try:
            props = {}

            for prop_set in doc.PropertySets:
                for prop in prop_set:
                    try:
                        props[prop.Name] = str(prop.Value)
                    except:
                        pass

            return props

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def processar_iam(self, cursor, doc, id_arquivo, props):
        try:
            estrutura_nova = self.consulta_estrutura_inventor(doc, cursor)

            total_itens = len(estrutura_nova)

            self.salvar_propriedades_iam(cursor, id_arquivo, props, total_itens)

            self.sincronizar_estrutura(cursor, id_arquivo, estrutura_nova)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_estrutura_inventor(self, doc, cursor):
        try:
            comp = doc.ComponentDefinition
            bom = comp.BOM

            bom.StructuredViewEnabled = True
            bom.StructuredViewFirstLevelOnly = False

            view = None
            for v in bom.BOMViews:
                if v.ViewType == 62465:
                    view = v

            estrutura = {}

            if not view:
                return estrutura

            for row in view.BOMRows:
                if row.ComponentDefinitions.Count == 0:
                    continue

                comp_def = row.ComponentDefinitions.Item(1)
                caminho = comp_def.Document.FullFileName

                id_filho = self.consulta_e_cria_id_arquivo(cursor, caminho)

                self.inserir_fila_conferencia(cursor, id_filho)

                estrutura[id_filho] = float(row.ItemQuantity)

            return estrutura

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def inserir_fila_conferencia(self, cursor, id_arquivo):
        try:
            cursor.execute("""
                SELECT 1 FROM FILA_CONFERENCIA WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            if cursor.fetchone():
                return

            cursor.execute("""
                INSERT INTO FILA_CONFERENCIA (ID_ARQUIVO)
                VALUES (?)
            """, (id_arquivo,))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_e_cria_id_arquivo(self, cursor, caminho):
        try:
            caminho_original = caminho
            caminho = caminho.lower()

            cursor.execute("""
                SELECT ID FROM ARQUIVOS WHERE CAMINHO=?
            """, (caminho,))

            row = cursor.fetchone()

            if row:
                return row[0]

            nome = os.path.basename(caminho_original)
            nome_sem_ext = os.path.splitext(nome)[0]

            classificacao = "ACABADO" if padrao_desenho.search(nome_sem_ext) else "MATERIA_PRIMA"
            dfgfd

            stat = os.stat(caminho_original)

            tamanho = stat.st_size
            data_mod = datetime.fromtimestamp(stat.st_mtime)

            extensao = os.path.splitext(nome)[1].lower()
            tipo_arquivo = extensao.replace(".", "").upper()

            cursor.execute("""
                INSERT INTO ARQUIVOS (
                    ARQUIVO, NOME_BASE, EXTENSAO, TIPO_ARQUIVO,
                    CLASSIFICACAO, CAMINHO, TAMANHO, DATA_MOD
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                RETURNING ID
            """, (
                nome,
                nome_sem_ext,
                extensao,
                tipo_arquivo,
                classificacao,
                caminho,
                tamanho,
                data_mod
            ))

            return cursor.fetchone()[0]

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_propriedades_ipt(self, cursor, id_arquivo, props):
        try:
            cursor.execute("""
                SELECT ID FROM PROPRIEDADES_IPT WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            row = cursor.fetchone()

            dados = (
                props.get("Revision Number"),
                props.get("Part Number"),
                props.get("Description"),
            )

            if row:
                cursor.execute("""
                    UPDATE PROPRIEDADES_IPT
                    SET REVISION_NUMBER=?, PART_NUMBER=?, DESCRIPTION=?
                    WHERE ID_ARQUIVO=?
                """, (*dados, id_arquivo))
            else:
                cursor.execute("""
                    INSERT INTO PROPRIEDADES_IPT
                    (ID_ARQUIVO, REVISION_NUMBER, PART_NUMBER, DESCRIPTION)
                    VALUES (?, ?, ?, ?)
                """, (id_arquivo, *dados))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_propriedades_iam(self, cursor, id_arquivo, props, total_itens):
        try:
            cursor.execute("""
                SELECT ID FROM PROPRIEDADES_IAM WHERE ID_ARQUIVO=?
            """, (id_arquivo,))

            row = cursor.fetchone()

            dados = (
                props.get("Revision Number"),
                props.get("Part Number"),
                props.get("Description"),
                total_itens
            )

            if row:
                cursor.execute("""
                    UPDATE PROPRIEDADES_IAM
                    SET REVISION_NUMBER=?, PART_NUMBER=?, DESCRIPTION=?, TOTAL_ITENS=?
                    WHERE ID_ARQUIVO=?
                """, (*dados, id_arquivo))
            else:
                cursor.execute("""
                    INSERT INTO PROPRIEDADES_IAM
                    (ID_ARQUIVO, REVISION_NUMBER, PART_NUMBER, DESCRIPTION, TOTAL_ITENS)
                    VALUES (?, ?, ?, ?, ?)
                """, (id_arquivo, *dados))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def sincronizar_estrutura(self, cursor, id_pai, estrutura_nova):
        try:
            estrutura_atual = self.consulta_estrutura_atual(cursor, id_pai)

            inseridos = 0
            atualizados = 0
            deletados = 0

            # 🔹 INSERT ou UPDATE
            for id_filho, qtde_nova in estrutura_nova.items():

                if id_filho not in estrutura_atual:
                    cursor.execute("""
                        INSERT INTO ESTRUTURA (ID_PAI, ID_FILHO, QTDE)
                        VALUES (?, ?, ?)
                    """, (id_pai, id_filho, qtde_nova))
                    inseridos += 1

                else:
                    qtde_atual = estrutura_atual[id_filho]

                    if round(qtde_atual, 4) != round(qtde_nova, 4):
                        cursor.execute("""
                            UPDATE ESTRUTURA
                            SET QTDE=?
                            WHERE ID_PAI=? AND ID_FILHO=?
                        """, (qtde_nova, id_pai, id_filho))
                        atualizados += 1

            # 🔻 DELETE só do que não existe mais
            for id_filho in estrutura_atual:

                if id_filho not in estrutura_nova:
                    cursor.execute("""
                        DELETE FROM ESTRUTURA
                        WHERE ID_PAI=? AND ID_FILHO=?
                    """, (id_pai, id_filho))
                    deletados += 1

            print(f"🟡 Estrutura → +{inseridos} | ~{atualizados} | -{deletados}")

            return inseridos, atualizados, deletados

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_estrutura_atual(self, cursor, id_pai):
        try:
            cursor.execute("""
                SELECT ID_FILHO, QTDE
                FROM ESTRUTURA
                WHERE ID_PAI=?
            """, (id_pai,))

            estrutura = {}

            for row in cursor.fetchall():
                id_filho, qtde = row
                estrutura[id_filho] = float(qtde)

            return estrutura

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def processar_idw(self, cursor, doc, id_arquivo):
        try:
            for ref in doc.ReferencedDocuments:
                caminho_ref = ref.FullFileName
                id_ref = self.consulta_e_cria_id_arquivo(cursor, caminho_ref)

                cursor.execute("""
                    DELETE FROM PROPRIEDADES_IDW WHERE ID_ARQUIVO=?
                """, (id_arquivo,))

                cursor.execute("""
                    INSERT INTO PROPRIEDADES_IDW (ID_ARQUIVO, ID_ARQUIVO_REFERENCIA)
                    VALUES (?, ?)
                """, (id_arquivo, id_ref))

                break

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == "__main__":
    freeze_support()
    WorkerFilaConferencia()