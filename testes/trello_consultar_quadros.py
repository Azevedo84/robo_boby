from core.banco import conecta
from core.erros import trata_excecao
from trello import Trello

class ConsultarQuadros:
    def __init__(self):
        self.trello = Trello()

        self.id_qd_eletrica = "5ca736f787917954a3235ebf"
        self.id_qd_usinagem = "5ca4da189cf4cd2c3a88ebf1"
        self.id_qd_ajustagem = "5cf7ad02756cf770a921b8db"
        self.id_qd_solda = ""
        self.id_qd_pcp = "5ca731d27ac4f92d562a476b"
        self.id_qd_projeto = ""

        self.iniciar_processo()

        #self.testar_vinculo_op_8843()

    def testar_vinculo_op_8843(self):
        id_os = "6a982d502a20900f4f3f7e6a"

        url_op = "https://trello.com/c/rB6gnnSJ/605-op-8843-conjunto-impressora-flexo-300"

        resultado = self.trello.adicionar_anexo_url(
            cartao_id=id_os,
            url=url_op,
            nome="OP 8843 - CONJUNTO IMPRESSORA FLEXO 300"
        )

        print(resultado)

    def iniciar_processo(self):
        try:
            quadros = self.trello.listar_quadros()

            for i in quadros:
                #print(i)
                pass
            setores = {
                "Elétrica": self.id_qd_eletrica,
                "Usinagem": self.id_qd_usinagem,
                "Ajustagem": self.id_qd_ajustagem,
                "PCP": self.id_qd_pcp,
            }

            faltantes = []

            for nome, id_quadro in setores.items():
                if not any(i["id"] == id_quadro for i in quadros):
                    faltantes.append(nome)

            if faltantes:
                print("Quadros não encontrados:")

                for nome in faltantes:
                    print(f"- {nome}")

                return

            self.analisar_ajustagem()

        except Exception as e:
            trata_excecao(e)
            raise

    def analisar_ajustagem(self):
        try:
            listas = self.trello.listar_listas(self.id_qd_ajustagem)

            print("\n================ AJUSTAGEM ================")

            for lista in listas:
                print(f"\nLISTA: {lista['name']}")

                cartoes = self.trello.listar_cartoes(lista["id"])

                for cartao in cartoes:
                    print(
                        f"CARTÃO DE SERVIÇO: "
                        f"{cartao['id']} - {cartao['name']}"
                    )

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    ConsultarQuadros()