from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, black
from reportlab.pdfbase.pdfmetrics import stringWidth

import os


# =========================================================
# CONFIGURAÇÕES
# =========================================================

PAGE_W, PAGE_H = A4

MARGIN_X = 8 * mm
TOP_Y = PAGE_H - 8 * mm

LINE_WIDTH = 0.6

FONT = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

COR_CINZA = HexColor("#cfcfcf")
COR_HEADER = HexColor("#efefef")


# =========================================================
# COLUNAS
# =========================================================

COL_NF = 18 * mm
COL_ITEM = 10 * mm
COL_CODIGO = 18 * mm
COL_QTDE = 16 * mm
COL_DESC = 76 * mm
COL_UN = 10 * mm
COL_UNIT = 22 * mm
COL_TOTAL = 28 * mm

TABLE_W = (
    COL_NF +
    COL_ITEM +
    COL_CODIGO +
    COL_QTDE +
    COL_DESC +
    COL_UN +
    COL_UNIT +
    COL_TOTAL
)


# =========================================================
# ALTURAS
# =========================================================

H_TITLE = 10 * mm
H_ROW = 8 * mm
H_BIG_ROW = 12 * mm
H_DRIVER = 16 * mm

H_HEADER_TABLE = 8 * mm
H_ITEM = 10 * mm

H_FOOTER = 14 * mm


# =========================================================
# HELPERS
# =========================================================

def box(c, x, y, w, h, fill=None):

    if fill:
        c.setFillColor(fill)
        c.rect(x, y, w, h, fill=1)
        c.setFillColor(black)
    else:
        c.rect(x, y, w, h, fill=0)


def line(c, x1, y1, x2, y2):
    c.line(x1, y1, x2, y2)


def text(
        c,
        x,
        y,
        value,
        size=8,
        bold=False,
        center=False,
        right=False
):

    font = FONT_BOLD if bold else FONT

    c.setFont(font, size)

    value = str(value)

    largura = stringWidth(value, font, size)

    if center:
        x -= largura / 2

    if right:
        x -= largura

    c.drawString(x, y, value)


# =========================================================
# DADOS
# =========================================================

itens = [(1, "35629", "71076", "30", "SUPORTE LAMINA PICOTE", "UN", "R$ 60,35", "R$ 1.810,50")]

# for i in range(1, 55):
#
#     itens.append([
#         i,
#         f"53{i:03}",
#         f"17{i:03}",
#         1,
#         f"PRODUTO TESTE {i}",
#         "UN",
#         "500,00",
#         "R$ 500,00"
#     ])


# =========================================================
# PDF
# =========================================================

desktop = os.path.join(
    os.path.expanduser("~"),
    "Desktop"
)

arquivo = os.path.join(
    desktop,
    "solicitacao_nf.pdf"
)

c = canvas.Canvas(
    arquivo,
    pagesize=A4
)

c.setLineWidth(LINE_WIDTH)

headers = [
    ("ITEM", COL_ITEM),
    ("Nº NF", COL_NF),
    ("CÓDIGO", COL_CODIGO),
    ("QTDE", COL_QTDE),
    ("DESCRIÇÃO", COL_DESC),
    ("UN", COL_UN),
    ("VLR UNIT.", COL_UNIT),
    ("VLR TOTAL", COL_TOTAL),
]


# =========================================================
# HEADER DOCUMENTO
# =========================================================

def desenha_topo():

    global y

    y = TOP_Y

    # =====================================================
    # TÍTULO
    # =====================================================

    box(
        c,
        MARGIN_X,
        y - H_TITLE,
        TABLE_W,
        H_TITLE
    )

    text(
        c,
        PAGE_W / 2,
        y - 6.5 * mm,
        "SOLICITAÇÃO DE NOTA FISCAL",
        9,
        True,
        center=True
    )

    y -= H_TITLE

    # =====================================================
    # MATERIAL
    # =====================================================

    w1 = 70 * mm
    w2 = TABLE_W - w1

    box(c, MARGIN_X, y - H_ROW, w1, H_ROW)
    box(c, MARGIN_X + w1, y - H_ROW, w2, H_ROW)

    text(c, MARGIN_X + 2 * mm, y - 5.5 * mm, "Material:", 7, True)
    text(c, MARGIN_X + 24 * mm, y - 5.5 * mm, "NOVO", 8)

    text(c, MARGIN_X + w1 + 2 * mm, y - 5.5 * mm, "Tipo:", 7, True)
    text(c, MARGIN_X + w1 + 18 * mm, y - 5.5 * mm, "INDUSTRIALIZADO", 8)

    y -= H_ROW

    # =====================================================
    # DE
    # =====================================================

    w1 = 70 * mm
    w2 = 50 * mm
    w3 = 35 * mm
    w4 = TABLE_W - (w1 + w2 + w3)

    x = MARGIN_X

    box(c, x, y - H_ROW, w1, H_ROW)
    box(c, x + w1, y - H_ROW, w2, H_ROW)
    box(c, x + w1 + w2, y - H_ROW, w3, H_ROW)
    box(c, x + w1 + w2 + w3, y - H_ROW, w4, H_ROW)

    text(c, x + 2 * mm, y - 5.5 * mm, "De:", 7, True)
    text(c, x + 12 * mm, y - 5.5 * mm, "ACINPLAS", 8)

    text(c, x + w1 + 2 * mm, y - 5.5 * mm, "Solicitante:", 7, True)
    text(c, x + w1 + 30 * mm, y - 5.5 * mm, "ANDERSON", 8)

    text(c, x + w1 + w2 + 2 * mm, y - 5.5 * mm, "N.º.OC:", 7, True)

    y -= H_ROW

    # =====================================================
    # PARA
    # =====================================================

    h = H_BIG_ROW

    w1 = 125 * mm
    w2 = TABLE_W - w1

    x = MARGIN_X

    box(c, x, y - h, w1, h)
    box(c, x + w1, y - h, w2, h)

    text(c, x + 2 * mm, y - 5 * mm, "Para:", 7, True)
    text(c, x + 15 * mm, y - 5 * mm, "Destino", 8)

    text(c, x + w1 + 2 * mm, y - 4.5 * mm, "Código Cliente /", 7, True)
    text(c, x + w1 + 2 * mm, y - 8.5 * mm, "Fornecedor :", 7, True)

    y -= h

    # =====================================================
    # OPERAÇÃO
    # =====================================================

    x = MARGIN_X

    w1 = 105 * mm
    w2 = 45 * mm
    w3 = TABLE_W - (w1 + w2)

    box(c, x, y - h, w1, h)
    box(c, x + w1, y - h, w2, h)
    box(c, x + w1 + w2, y - h, w3, h)

    text(c, x + 2 * mm, y - 5 * mm, "Operação:", 7, True)
    text(c, x + 22 * mm, y - 5 * mm, "VENDA", 8)

    text(c, x + w1 + 2 * mm, y - 5 * mm, "Transporte:", 7, True)

    text(c, x + w1 + w2 + 2 * mm, y - 5 * mm, "Frete:", 7, True)

    y -= h

    # =====================================================
    # OBSERVAÇÕES
    # =====================================================

    obs_h = 18 * mm

    box(
        c,
        MARGIN_X,
        y - obs_h,
        TABLE_W,
        obs_h
    )

    text(
        c,
        MARGIN_X + 2 * mm,
        y - 5 * mm,
        "OBSERVAÇÕES ADICIONAIS DA NF / ITENS DE BAIXA / INFORMAÇÕES IMPORTANTES",
        7,
        True
    )

    y -= obs_h


# =========================================================
# HEADER TABELA
# =========================================================

def desenha_header_tabela(pos_y):

    x = MARGIN_X

    for titulo, largura in headers:

        box(
            c,
            x,
            pos_y - H_HEADER_TABLE,
            largura,
            H_HEADER_TABLE,
            fill=COR_HEADER
        )

        text(
            c,
            x + largura / 2,
            pos_y - 5.2 * mm,
            titulo,
            6.5,
            True,
            center=True
        )

        x += largura

    return pos_y - H_HEADER_TABLE


# =========================================================
# INÍCIO
# =========================================================

desenha_topo()

y = desenha_header_tabela(y)


# =========================================================
# ITENS
# =========================================================

ITENS_MINIMOS = 15

# =========================================================
# COMPLETA SOMENTE SE TIVER MENOS DE 15
# =========================================================

if len(itens) < ITENS_MINIMOS:

    for i in range(len(itens) + 1, ITENS_MINIMOS + 1):

        itens.append([
            i,
            "",
            "",
            "",
            "",
            "",
            "",
            "R$ 0,00"
        ])


# =========================================================
# DESENHA ITENS
# =========================================================

for item in itens:

    # =====================================================
    # NOVA PÁGINA
    # =====================================================

    if y - H_ITEM < H_FOOTER:

        c.showPage()

        c.setLineWidth(LINE_WIDTH)

        y = TOP_Y

        y = desenha_header_tabela(y)

    x = MARGIN_X

    for idx, valor in enumerate(item):

        largura = headers[idx][1]

        fill = COR_CINZA if idx == 7 else None

        box(
            c,
            x,
            y - H_ITEM,
            largura,
            H_ITEM,
            fill=fill
        )

        # =================================================
        # CENTRALIZADAS
        # =================================================

        if idx in [0, 1, 2, 3, 5]:

            text(
                c,
                x + largura / 2,
                y - 6.5 * mm,
                valor,
                7.5,
                center=True
            )

        # =================================================
        # TOTAL
        # =================================================

        elif idx == 7:

            text(
                c,
                x + largura - 2 * mm,
                y - 6.5 * mm,
                valor,
                7.5,
                True,
                right=True
            )

        # =================================================
        # DESCRIÇÃO
        # =================================================

        else:

            text(
                c,
                x + 2 * mm,
                y - 6.5 * mm,
                valor,
                7.5
            )

        x += largura

    y -= H_ITEM


# =========================================================
# FOOTER
# =========================================================

footer_y = y

footer = [
    ("PESO BRUTO (Kg):", "6,41 KG", 45 * mm),
    ("PESO LÍQUIDO (Kg):", "6,41 KG", 45 * mm),
    ("VOLUME:", "4", 40 * mm),
    ("TOTAL GERAL:", "R$ 118,87", 55 * mm),
]

x = MARGIN_X

for idx, (titulo, valor, largura) in enumerate(footer):

    fill = COR_CINZA if idx == 3 else None

    box(
        c,
        x,
        footer_y - H_FOOTER,
        largura,
        H_FOOTER,
        fill=fill
    )

    line(
        c,
        x,
        footer_y - (H_FOOTER / 2),
        x + largura,
        footer_y - (H_FOOTER / 2)
    )

    text(
        c,
        x + largura / 2,
        footer_y - 5 * mm,
        titulo,
        7,
        True,
        center=True
    )

    text(
        c,
        x + largura / 2,
        footer_y - 10.5 * mm,
        valor,
        8,
        idx == 3,
        center=True
    )

    x += largura


# =========================================================
# FINALIZAR
# =========================================================

c.save()

os.startfile(arquivo)