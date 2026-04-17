import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
from core.banco_nuvem import conectar_banco_nuvem
from core.erros import trata_excecao


class AtualizaCustoDaNuvem:
    def __init__(self):
        self.inicia_processo()

    # 🔹 Normaliza código (resolve 99% dos problemas)
    def normalizar_codigo(self, c):
        if c is None:
            return None
        return str(c).strip()

    def inicia_processo(self):
        conecta_nuvem = conectar_banco_nuvem()

        try:
            cursor_local = conecta.cursor()
            cursor_nuvem = conecta_nuvem.cursor()

            # 🔹 1. Dados da nuvem
            cursor_nuvem.execute("""
                SELECT CODIGO_PROD, CUSTO_MEDIO 
                FROM CUSTO_MEDIO_PRODUTO;
            """)
            dados_nuvem = cursor_nuvem.fetchall()

            if not dados_nuvem:
                print("Nenhum dado na nuvem")
                return

            print(f"{len(dados_nuvem)} produtos na nuvem")

            # 🔹 2. Produtos locais
            cursor_local.execute("""
                SELECT id, codigo, conjunto, CUSTOUNITARIO 
                FROM produto;
            """)
            produtos_locais = cursor_local.fetchall()

            # 🔹 Dicionário NORMALIZADO
            produtos_dict = {
                self.normalizar_codigo(p[1]): (p[0], p[2], p[3])
                for p in produtos_locais
            }

            # 🔹 3. Processa updates
            updates = []
            nao_encontrados = 0

            for cod_prod, custo_medio in dados_nuvem:
                cod_prod_norm = self.normalizar_codigo(cod_prod)

                # 🔹 ignora lixo (opcional, mas recomendado)
                if not cod_prod_norm or not cod_prod_norm.replace('.', '').isdigit():
                    continue

                produto_local = produtos_dict.get(cod_prod_norm)

                if not produto_local:
                    nao_encontrados += 1
                    continue

                id_prod, conjunto, custo_atual = produto_local

                if conjunto != 10 and float(custo_medio) != float(custo_atual):
                    updates.append((custo_medio, id_prod))

            # 🔹 4. Executa updates
            if updates:
                cursor_local.executemany(
                    "UPDATE produto SET CUSTOUNITARIO = ? WHERE id = ?",
                    updates
                )
                conecta.commit()
                print(f"{len(updates)} produtos atualizados")
            else:
                print("Nenhum produto precisou ser atualizado")

            print(f"{nao_encontrados} códigos da nuvem não encontrados no ERP")

            # 🔹 5. Limpeza (se quiser ativar)
            cursor_nuvem.execute("SELECT COUNT(*) FROM CUSTO_MEDIO_PRODUTO;")
            total = cursor_nuvem.fetchone()[0]
            print(f"Vai apagar {total} registros da nuvem")

            cursor_nuvem.execute("DELETE FROM CUSTO_MEDIO_PRODUTO;")
            conecta_nuvem.commit()

            print("Processo finalizado")

        except Exception as e:
            trata_excecao(e)
            raise

        finally:
            try:
                conecta_nuvem.close()
            except:
                pass


# 🔹 EXECUÇÃO
if __name__ == "__main__":
    AtualizaCustoDaNuvem()