import win32com.client
from core.banco import conecta_engenharia
from core.erros import trata_excecao
from multiprocessing import freeze_support

class GravarPropriedadeInventor:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

        self.manipula_comeco()

    def manipula_comeco(self):
        try:
            freeze_support()
            print("🚀 Worker de fila iniciado")
            self.worker_fila()

        except Exception as e:
            trata_excecao(e)
            raise

    def conectar_inventor(self):
        try:
            print("Conectando ao Inventor...")

            try:
                inventor = win32com.client.GetActiveObject("Inventor.Application")
                print("♻️ Reutilizando Inventor já aberto")
            except:
                inventor = win32com.client.Dispatch("Inventor.Application")
                print("🆕 Abrindo novo Inventor")

            inventor.Visible = False
            inventor.SilentOperation = True

            return inventor

        except Exception as e:
            trata_excecao(e)
            raise

    def worker_fila(self):
        inventor = None
        try:
            cursor = conecta_engenharia.cursor()
            cursor.execute("""
                    SELECT fila.ID_ARQUIVO, arq.caminho, arq.NOME_BASE
                    FROM FILA_GERAR_PDF as fila
                    INNER JOIN ARQUIVOS AS arq ON fila.ID_ARQUIVO = arq.id
                    ORDER BY fila.ID
                """)
            fila = cursor.fetchall()

            print(f"\n📦 Total na fila: {len(fila)}")

            if fila:
                inventor = self.conectar_inventor()

                for id_arquivo, caminho, nome_base in fila:
                    print(f"🔍 Abrindo: {caminho}")

                    doc = None

                    try:
                        doc = inventor.Documents.Open(caminho, False)

                        sucesso = self.processar_arquivo(inventor, caminho, nome_base)

                        if sucesso:
                            cursor.execute("""
                                            DELETE FROM FILA_GERAR_PDF
                                            WHERE ID_ARQUIVO=?
                                        """, (id_arquivo,))
                            conecta_engenharia.commit()

                            print(f"PRODUTO ATUALIZADO COM SUCESSO - SAI DA FILA: {caminho}")
                        else:
                            conecta_engenharia.rollback()
                            print(f"PRODUTO COM PROBLEMAS - FICA NA FILA: {caminho}")

                    except Exception as e:
                        conecta_engenharia.rollback()
                        trata_excecao(e)
                        continue

                    finally:
                        if doc is not None:
                            try:
                                doc.Close(True) # type: ignore
                            except Exception as e:
                                trata_excecao(e)

        except Exception as e:
            trata_excecao(e)
            raise

        finally:
            if inventor is not None:
                try:
                    inventor.Quit()
                except Exception as e:
                    trata_excecao(e)
                    raise

    def processar_arquivo(self, inventor, caminho, nome_base):
        try:
            print(f"\n📂 Processando: {caminho}")

            caminho_pdf = rf"\\Publico\C\OP\Projetos\{nome_base}.pdf"
            arquivo_pdf = f"{nome_base}.pdf"

            doc = inventor.Documents.Open(caminho, False)

            pdf_addin = inventor.ApplicationAddIns.ItemById(
                "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"
            )

            if not pdf_addin.Activated:
                pdf_addin.Activate()

            context = inventor.TransientObjects.CreateTranslationContext()
            context.Type = 13059

            options = inventor.TransientObjects.CreateNameValueMap()

            data = inventor.TransientObjects.CreateDataMedium()
            data.FileName = caminho_pdf

            print("Gerando PDF...")
            pdf_addin.SaveCopyAs(doc, context, options, data)
            print("PDF criado:", arquivo_pdf)

            return True

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    freeze_support()
    GravarPropriedadeInventor()