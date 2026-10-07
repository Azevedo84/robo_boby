from trello import Trello


trello = Trello()

# =========================================================
# QUADRO PCP
# =========================================================

quadro_id = "5ca731d27ac4f92d562a476b"

quadro = trello.buscar_quadro(quadro_id)

print("\n================ QUADRO ================")
print("ID:", quadro["id"])
print("Nome:", quadro["name"])


# =========================================================
# LISTAS DO QUADRO
# =========================================================

listas = trello.listar_listas(quadro_id)

print("\n================ LISTAS ================")

for lista in listas:
    print("\nID:", lista["id"])
    print("Nome:", lista["name"])


    # =====================================================
    # CARTÕES DA LISTA
    # =====================================================

    cartoes = trello.listar_cartoes(lista["id"])

    for cartao in cartoes:
        print("   CARTÃO:", cartao["id"], "-", cartao["name"])

# =========================================================
# HISTÓRICO DO CARTÃO
# =========================================================

cartao_id = "6a9043cc25c8b2d4b5d5afa5"

acoes = trello.listar_acoes_cartao(cartao_id)

print("\n================ HISTÓRICO ================")

for acao in acoes:
    print(acao)