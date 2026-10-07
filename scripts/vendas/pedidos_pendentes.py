import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
from core.erros import trata_excecao
from core.email_service import dados_email
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from datetime import date


class RelatorioPIPendentes:
    def __init__(self):
        self.styles = getSampleStyleSheet()

        self.destinatarios_todos = ['<maquinas@unisold.com.br>']

        self.destinatarios_clientes = {
            'VOTI': ['<maquinas@unisold.com.br>'],
            'HOMEBAG': ['<maquinas@unisold.com.br>'],
            'KOBA': ['<maquinas@unisold.com.br>'],
            'PLASTICOS ARACAJU': ['<maquinas@unisold.com.br>'],
            'PLASTICOS SUZUKI': ['<maquinas@unisold.com.br>'],

        }

        self.iniciar_processo()

    def iniciar_processo(self):
        try:
            cursor = conecta.cursor()

            cursor.execute(
                f"SELECT ped.emissao, ped.id, cli.razao, prod.codigo, prod.DESCRICAO, "
                f"prod.obs, prod.unidade, prodint.qtde, prodint.data_previsao "
                f"FROM PRODUTOPEDIDOINTERNO as prodint "
                f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                f"where prodint.status = 'A' and cli.id <> 16 "
                f"order by ped.emissao;"
            )

            dados_pi = cursor.fetchall()

            self.produtos_sem_imagem = []
            self.dados_pi = dados_pi
            self.dados_pi_por_cliente = {}

            if dados_pi:

                for i_pi in dados_pi:

                    (
                        emissao_pi,
                        num_pi,
                        cliente,
                        cod,
                        descr,
                        ref,
                        um,
                        qtde_pi,
                        entrega_pi
                    ) = i_pi

                    imagem = self.localizar_imagem(cod)

                    if imagem is None:
                        self.produtos_sem_imagem.append(
                            (
                                num_pi,
                                cliente,
                                cod,
                                descr,
                                ref
                            )
                        )

                    # Agrupa por cliente
                    if cliente not in self.dados_pi_por_cliente:
                        self.dados_pi_por_cliente[cliente] = []

                    self.dados_pi_por_cliente[cliente].append(i_pi)

                # Ordena os produtos de cada cliente pela data de entrega
                for cliente in self.dados_pi_por_cliente:
                    self.dados_pi_por_cliente[cliente].sort(
                        key=lambda x: x[8] if x[8] else date.max
                    )

                print("\nPedidos agrupados por cliente:")

                for cliente, produtos in self.dados_pi_por_cliente.items():
                    print(
                        f"{cliente}: {len(produtos)} produtos"
                    )

                if self.produtos_sem_imagem:

                    print("\nProdutos sem imagem:")

                    for produto in self.produtos_sem_imagem:
                        print(produto)

        except Exception as e:
            trata_excecao(e)
            raise

    def enviar_email_produtos_sem_imagem(self):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = "Pedidos Pendentes – Produtos sem imagem"

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = ""

            body += f"{saudacao}\n\n"
            body += "Os seguintes produtos dos Pedidos Pendentes estão sem imagem vinculada:\n\n"

            for num_pi, cliente, cod, descr, ref in self.produtos_sem_imagem:
                body += f"PI: {num_pi} | Cliente: {cliente} | Código: {cod} | Produto: {descr} | Referência: {ref}\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatarios_todos, text)

            server.quit()

            print("Email de produtos sem imagem enviado com sucesso.")

        except Exception as e:
            trata_excecao(e)
            raise

    def localizar_imagem(self, codigo):
        caminho_png = rf"\\Publico\C\OP\Projetos\{codigo}.png"

        if os.path.exists(caminho_png):
            return caminho_png

        return None

    def calcular_posicao(self, codigo, id_produto, conjunto, tipo_material,
                         saldo_produto, qtde_solicitada):
        try:
            cursor = conecta.cursor()

            saldo_produto = saldo_produto or 0
            qtde_solicitada = qtde_solicitada or 0

            # =====================================================
            # ESTOQUE
            # =====================================================

            if saldo_produto >= qtde_solicitada:
                return {
                    "status": "DISPONÍVEL PARA ENVIO",
                    "percentual": None,
                    "detalhe": ""
                }

            # =====================================================
            # COMPRADO
            # =====================================================

            if conjunto != 10:

                sql_oc = """
                    SELECT
                        OC.NUMERO,
                        POC.QUANTIDADE,
                        COALESCE(POC.PRODUZIDO, 0),
                        POC.DATAENTREGA
                    FROM PRODUTOORDEMCOMPRA POC
                    INNER JOIN ORDEMCOMPRA OC
                        ON POC.MESTRE = OC.ID
                    WHERE
                        POC.PRODUTO = ?
                        AND POC.QUANTIDADE > COALESCE(POC.PRODUZIDO, 0)
                        AND OC.STATUS = 'A'
                    ORDER BY POC.DATAENTREGA
                """

                cursor.execute(sql_oc, (id_produto,))
                ordem_compra = cursor.fetchone()

                if ordem_compra:
                    numero_oc, quantidade_oc, produzido_oc, entrega_oc = ordem_compra

                    return {
                        "status": "COMPRA REALIZADA - AGUARDANDO RECEBIMENTO",
                        "percentual": None,
                        "detalhe": (
                            f"OC {numero_oc} | "
                            f"Previsão "
                            f"{entrega_oc.strftime('%d/%m/%Y') if entrega_oc else '-'}"
                        )
                    }

                return {
                    "status": "AGUARDANDO COMPRA",
                    "percentual": None,
                    "detalhe": ""
                }

            # =====================================================
            # FABRICADO
            # =====================================================

            # OP aberta do próprio produto
            sql_op = """
                SELECT
                    OS.NUMERO,
                    OS.QUANTIDADE,
                    OS.DATAPREVISAO
                FROM ORDEMSERVICO OS
                WHERE
                    OS.STATUS = 'A'
                    AND OS.PRODUTO = ?
                ORDER BY OS.DATAPREVISAO, OS.ID
            """

            cursor.execute(sql_op, (id_produto,))
            op = cursor.fetchone()

            # =====================================================
            # PERCENTUAL REAL DA PRODUÇÃO
            # =====================================================

            estrutura_total = self.total_estrutura(
                codigo,
                qtde_solicitada
            )

            total_itens_estrutura = len(estrutura_total)

            consumo_total = self.calculo_3_verifica_estrutura(
                codigo,
                qtde_solicitada
            )

            total_faltantes = len(consumo_total)
            print(total_itens_estrutura, total_faltantes)

            if total_itens_estrutura:
                percentual = int(
                    (
                            (total_itens_estrutura - total_faltantes)
                            / total_itens_estrutura
                    ) * 100
                )
            else:
                percentual = 0

            # =====================================================
            # TEM OP ABERTA
            # =====================================================

            if op:
                numero_op, quantidade_op, previsao_op = op

                previsao = (
                    previsao_op.strftime("%d/%m/%Y")
                    if previsao_op
                    else "-"
                )

                return {
                    "status": "EM PRODUÇÃO",
                    "percentual": percentual,
                    "detalhe": (
                        f"OP {numero_op} | "
                        f"Previsão {previsao}"
                    )
                }

            # =====================================================
            # SEM OP
            # =====================================================

            if percentual == 100:
                return {
                    "status": "PRODUZIDO - AGUARDANDO ENTRADA",
                    "percentual": 100,
                    "detalhe": ""
                }

            if tipo_material == 119:
                return {
                    "status": "AGUARDANDO PRODUÇÃO",
                    "percentual": percentual,
                    "detalhe": ""
                }

            return {
                "status": "AGUARDANDO PROJETO",
                "percentual": percentual,
                "detalhe": ""
            }

        except Exception as e:
            trata_excecao(e)
            raise

    def total_estrutura(self, codigo, qtde):
        try:
            cursor = conecta.cursor()

            cursor.execute(
                f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, tip.tipomaterial, prod.id_versao, prod.quantidade "
                f"FROM produto as prod "
                f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                f"WHERE prod.codigo = '{codigo}';"
            )

            detalhes_pai = cursor.fetchall()

            if detalhes_pai:
                id_pai, cod_pai, descr_pai, ref_pai, um_pai, tipo, id_estrut, saldo = detalhes_pai[0]

                filhos = []

                dadoss = (cod_pai, descr_pai, ref_pai, um_pai, qtde)
                filhos.append(dadoss)

                if id_estrut:
                    cursor.execute(
                        f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                        f"prod.unidade, "
                        f"(estprod.quantidade * {qtde}) as qtde "
                        f"FROM estrutura_produto as estprod "
                        f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                        f"WHERE estprod.id_estrutura = {id_estrut};"
                    )

                    dados_estrutura = cursor.fetchall()

                    if dados_estrutura:
                        for prod in dados_estrutura:
                            cod_f, descr_f, ref_f, um_f, qtde_f = prod

                            filhos_recursivos = self.total_estrutura(
                                cod_f,
                                qtde_f
                            )

                            if filhos_recursivos:
                                filhos.extend(filhos_recursivos)

                return filhos

            return []

        except Exception as e:
            trata_excecao(e)
            raise

    def manipula_consumo_op(self, cod_prod):
        try:
            itens_faltantes = []

            cursor = conecta.cursor()

            cursor.execute(
                f"SELECT ordser.numero, ordser.quantidade, ordser.id_estrutura "
                f"FROM ordemservico AS ordser "
                f"INNER JOIN produto prod ON ordser.produto = prod.id "
                f"WHERE ordser.status = 'A' "
                f"AND prod.codigo = {cod_prod} "
                f"ORDER BY ordser.numero;"
            )

            op_abertas = cursor.fetchall()

            for op, qtde_op, id_estrut in op_abertas:

                if not id_estrut:
                    continue

                cursor = conecta.cursor()

                cursor.execute(
                    f"SELECT estprod.id, "
                    f"(estprod.quantidade * {qtde_op}) AS qtde_necessaria, "
                    f"prod.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                    f"prod.unidade, prod.id_versao "
                    f"FROM estrutura_produto AS estprod "
                    f"INNER JOIN produto AS prod "
                    f"ON estprod.id_prod_filho = prod.id "
                    f"WHERE estprod.id_estrutura = {id_estrut};"
                )

                itens_estrutura = cursor.fetchall()

                for ides, qtde_necessaria, codigo, descricao, ref, um, id_estruti in itens_estrutura:

                    cursor = conecta.cursor()

                    cursor.execute(
                        f"SELECT COALESCE(SUM(prodser.qtde_estrut_prod), 0) "
                        f"FROM produtoos AS prodser "
                        f"WHERE prodser.numero = {op} "
                        f"AND prodser.id_estrut_prod = {ides};"
                    )

                    qtde_consumida = cursor.fetchone()[0] or 0

                    qtde_faltante = qtde_necessaria - qtde_consumida

                    if qtde_faltante > 0:
                        dados = (
                            codigo,
                            descricao,
                            ref,
                            um,
                            id_estrut,
                            qtde_faltante
                        )

                        itens_faltantes.append(dados)

            return itens_faltantes

        except Exception as e:
            trata_excecao(e)
            raise

    def calculo_3_verifica_estrutura(self, codigo, qtde):
        try:
            cursor = conecta.cursor()

            cursor.execute(
                f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, tip.tipomaterial, prod.id_versao, prod.quantidade "
                f"FROM produto as prod "
                f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                f"WHERE prod.codigo = '{codigo}';"
            )

            detalhes_pai = cursor.fetchall()

            if detalhes_pai:
                id_pai, cod_pai, descr_pai, ref_pai, um_pai, tipo, id_estrut, saldo = detalhes_pai[0]

                filhos = []

                itens_faltantes = self.manipula_consumo_op(cod_pai)

                dadoss = (
                    cod_pai,
                    descr_pai,
                    ref_pai,
                    um_pai,
                    qtde
                )

                filhos.append(dadoss)

                if itens_faltantes:

                    for titi in itens_faltantes:

                        cod_f, descricao, ref, um, id_estrut_falt, qtde_f = titi

                        if id_estrut_falt:

                            filhos_recursivos = self.calculo_3_verifica_estrutura(
                                cod_f,
                                qtde_f
                            )

                            if filhos_recursivos:
                                filhos.extend(filhos_recursivos)

                return filhos

            return []

        except Exception as e:
            trata_excecao(e)
            raise

    def gerar_pdf(self):
        try:
            from reportlab.platypus import (
                SimpleDocTemplate,
                Paragraph,
                Spacer,
                Image,
                Table,
                TableStyle,
                KeepTogether
            )
            from reportlab.lib import colors
            from reportlab.lib.styles import ParagraphStyle
            from reportlab.lib.units import mm
            from datetime import date, datetime
            import fitz

            def formatar_data(data):
                if isinstance(data, (date, datetime)):
                    return data.strftime("%d/%m/%Y")

                return str(data)

            # ---------------------------------------------------------
            # ESTILOS
            # ---------------------------------------------------------

            estilo_titulo = ParagraphStyle(
                "TituloCliente",
                parent=self.styles["Heading1"],
                fontName="Helvetica-Bold",
                fontSize=17,
                leading=20,
                textColor=colors.HexColor("#1F2937"),
                spaceAfter=0
            )

            estilo_subtitulo = ParagraphStyle(
                "Subtitulo",
                parent=self.styles["Normal"],
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                textColor=colors.HexColor("#6B7280")
            )

            estilo_pi = ParagraphStyle(
                "PI",
                parent=self.styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9,
                leading=11,
                textColor=colors.HexColor("#111827")
            )

            estilo_campo = ParagraphStyle(
                "Campo",
                parent=self.styles["Normal"],
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                textColor=colors.HexColor("#374151")
            )

            estilo_destaque = ParagraphStyle(
                "Destaque",
                parent=self.styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9,
                leading=11,
                textColor=colors.HexColor("#111827")
            )

            # ---------------------------------------------------------
            # PDFs POR CLIENTE
            # ---------------------------------------------------------

            for cliente, dados_cliente in self.dados_pi_por_cliente.items():

                arquivo = os.path.join(
                    os.getcwd(),
                    f"PI Pendentes - {cliente}.pdf"
                )

                doc = SimpleDocTemplate(
                    arquivo,
                    pagesize=A4,
                    rightMargin=25,
                    leftMargin=25,
                    topMargin=25,
                    bottomMargin=25
                )

                elementos = []

                # IMPORTANTE:
                # As imagens temporárias NÃO podem ser apagadas
                # durante a montagem dos elementos.
                imagens_temporarias = []

                # -----------------------------------------------------
                # CABEÇALHO
                # -----------------------------------------------------

                logo = self.localizar_logo_cliente(cliente)

                quantidade_itens = len(dados_cliente)

                if logo:
                    logo_img = Image(
                        logo,
                        width=35 * mm,
                        height=18 * mm,
                        kind="proportional"
                    )

                    cabecalho_cliente = Table(
                        [[
                            logo_img,
                            [
                                Paragraph(
                                    "Pedidos Pendentes",
                                    estilo_titulo
                                ),
                                Spacer(1, 3),
                                Paragraph(
                                    f"{quantidade_itens} itens pendentes",
                                    estilo_subtitulo
                                )
                            ]
                        ]],
                        colWidths=[
                            40 * mm,
                            125 * mm
                        ]
                    )

                else:
                    cabecalho_cliente = Table(
                        [[
                            Paragraph(
                                "Pedidos Pendentes",
                                estilo_titulo
                            ),
                            Paragraph(
                                f"{quantidade_itens} itens pendentes",
                                estilo_subtitulo
                            )
                        ]],
                        colWidths=[
                            85 * mm,
                            80 * mm
                        ]
                    )

                cabecalho_cliente.setStyle(
                    TableStyle([
                        (
                            "VALIGN",
                            (0, 0),
                            (-1, -1),
                            "MIDDLE"
                        ),
                        (
                            "LEFTPADDING",
                            (0, 0),
                            (-1, -1),
                            0
                        ),
                        (
                            "RIGHTPADDING",
                            (0, 0),
                            (-1, -1),
                            0
                        ),
                        (
                            "TOPPADDING",
                            (0, 0),
                            (-1, -1),
                            0
                        ),
                        (
                            "BOTTOMPADDING",
                            (0, 0),
                            (-1, -1),
                            0
                        ),
                    ])
                )

                elementos.append(cabecalho_cliente)

                elementos.append(
                    Spacer(1, 8)
                )

                linha = Table(
                    [[""]],
                    colWidths=[165 * mm],
                    rowHeights=[1]
                )

                linha.setStyle(
                    TableStyle([
                        (
                            "BACKGROUND",
                            (0, 0),
                            (-1, -1),
                            colors.HexColor("#D1D5DB")
                        ),
                    ])
                )

                elementos.append(linha)

                elementos.append(
                    Spacer(1, 10)
                )


                # -----------------------------------------------------
                # PRODUTOS
                # -----------------------------------------------------

                for i_pi in dados_cliente:

                    (
                        emissao_pi,
                        num_pi,
                        clie_pi,
                        cod,
                        descr,
                        ref,
                        um,
                        qtde_pi,
                        entrega_pi
                    ) = i_pi

                    # -------------------------------------------------
                    # POSIÇÃO DO PEDIDO
                    # -------------------------------------------------

                    cursor = conecta.cursor()

                    cursor.execute("""
                                   SELECT ID,
                                          CONJUNTO,
                                          TIPOMATERIAL,
                                          QUANTIDADE
                                   FROM PRODUTO
                                   WHERE CODIGO = ?
                                   """, (cod,))

                    produto = cursor.fetchone()

                    if produto:
                        (
                            id_produto,
                            conjunto,
                            tipo_material,
                            saldo_produto
                        ) = produto

                        posicao = self.calcular_posicao(
                            cod,
                            id_produto,
                            conjunto,
                            tipo_material,
                            saldo_produto,
                            qtde_pi
                        )
                    else:
                        posicao = {
                            "status": "POSIÇÃO NÃO IDENTIFICADA",
                            "percentual": None,
                            "detalhe": ""
                        }

                    imagem = self.localizar_imagem(
                        cod,
                    )

                    # -------------------------------------------------
                    # CONVERTER PDF DO DESENHO PARA PNG
                    # -------------------------------------------------

                    if imagem and imagem.lower().endswith(".pdf"):
                        doc_imagem = fitz.open(imagem)

                        page = doc_imagem.load_page(0)

                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(1.5, 1.5),
                            alpha=False
                        )

                        caminho_temp = os.path.join(
                            os.getcwd(),
                            f"_imagem_{cod}.png"
                        )

                        pix.save(caminho_temp)

                        doc_imagem.close()

                        imagem = caminho_temp

                        # Guarda o arquivo para apagar
                        # somente depois do doc.build()
                        imagens_temporarias.append(
                            caminho_temp
                        )

                    # -------------------------------------------------
                    # IMAGEM
                    # -------------------------------------------------

                    if imagem and os.path.exists(imagem):

                        imagem_produto = Image(
                            imagem,
                            width=48 * mm,
                            height=34 * mm,
                            kind="proportional"
                        )

                    else:

                        imagem_produto = Paragraph(
                            "Imagem não disponível",
                            estilo_subtitulo
                        )

                    # -------------------------------------------------
                    # DADOS
                    # -------------------------------------------------

                    dados_produto = []

                    dados_produto.append(
                        Paragraph(
                            f"<b>PI {num_pi}</b>",
                            estilo_pi
                        )
                    )

                    dados_produto.append(
                        Spacer(1, 3)
                    )

                    dados_produto.append(
                        Paragraph(
                            f"<b>Código:</b> {cod}",
                            estilo_campo
                        )
                    )

                    dados_produto.append(
                        Spacer(1, 2)
                    )

                    dados_produto.append(
                        Paragraph(
                            f"<b>Produto:</b> {descr}",
                            estilo_campo
                        )
                    )

                    dados_produto.append(
                        Spacer(1, 2)
                    )

                    dados_produto.append(
                        Paragraph(
                            f"<b>Referência:</b> {ref}",
                            estilo_campo
                        )
                    )

                    dados_produto.append(
                        Spacer(1, 2)
                    )

                    dados_produto.append(
                        Paragraph(
                            f"<b>Quantidade:</b> {qtde_pi} {um}",
                            estilo_campo
                        )
                    )

                    dados_produto.append(
                        Spacer(1, 2)
                    )

                    dados_produto.append(
                        Paragraph(
                            f"<b>Entrega:</b> "
                            f"{formatar_data(entrega_pi)}",
                            estilo_destaque
                        )
                    )

                    # -------------------------------------------------
                    # POSIÇÃO DO PEDIDO
                    # -------------------------------------------------

                    status = posicao["status"]
                    percentual = posicao["percentual"]
                    detalhe = posicao["detalhe"]

                    if percentual is not None:
                        texto_posicao = f"<b>{status}</b> — {percentual}%"
                    else:
                        texto_posicao = f"<b>{status}</b>"

                    dados_produto.append(
                        Spacer(1, 5)
                    )

                    dados_produto.append(
                        Paragraph(
                            texto_posicao,
                            estilo_destaque
                        )
                    )

                    if detalhe:
                        dados_produto.append(
                            Spacer(1, 2)
                        )

                        dados_produto.append(
                            Paragraph(
                                detalhe,
                                estilo_campo
                            )
                        )

                    # -------------------------------------------------
                    # CARD DO PRODUTO
                    # -------------------------------------------------

                    tabela_produto = Table(
                        [[
                            dados_produto,
                            imagem_produto
                        ]],
                        colWidths=[
                            112 * mm,
                            53 * mm
                        ]
                    )

                    tabela_produto.setStyle(
                        TableStyle([
                            (
                                "BACKGROUND",
                                (0, 0),
                                (-1, -1),
                                colors.HexColor("#F9FAFB")
                            ),
                            (
                                "BOX",
                                (0, 0),
                                (-1, -1),
                                0.7,
                                colors.HexColor("#D1D5DB")
                            ),
                            (
                                "VALIGN",
                                (0, 0),
                                (-1, -1),
                                "MIDDLE"
                            ),
                            (
                                "ALIGN",
                                (1, 0),
                                (1, 0),
                                "CENTER"
                            ),
                            (
                                "LEFTPADDING",
                                (0, 0),
                                (-1, -1),
                                9
                            ),
                            (
                                "RIGHTPADDING",
                                (0, 0),
                                (-1, -1),
                                9
                            ),
                            (
                                "TOPPADDING",
                                (0, 0),
                                (-1, -1),
                                8
                            ),
                            (
                                "BOTTOMPADDING",
                                (0, 0),
                                (-1, -1),
                                8
                            ),
                        ])
                    )

                    elementos.append(
                        KeepTogether(
                            tabela_produto
                        )
                    )

                    elementos.append(
                        Spacer(1, 7)
                    )

                # -----------------------------------------------------
                # RODAPÉ
                # -----------------------------------------------------

                elementos.append(
                    Spacer(1, 5)
                )

                rodape = Table(
                    [[
                        Paragraph(
                            f"Documento gerado automaticamente • {cliente}",
                            estilo_subtitulo
                        )
                    ]],
                    colWidths=[165 * mm]
                )

                rodape.setStyle(
                    TableStyle([
                        (
                            "ALIGN",
                            (0, 0),
                            (-1, -1),
                            "RIGHT"
                        ),
                        (
                            "LEFTPADDING",
                            (0, 0),
                            (-1, -1),
                            0
                        ),
                        (
                            "RIGHTPADDING",
                            (0, 0),
                            (-1, -1),
                            0
                        ),
                        (
                            "TOPPADDING",
                            (0, 0),
                            (-1, -1),
                            5
                        ),
                        (
                            "BOTTOMPADDING",
                            (0, 0),
                            (-1, -1),
                            0
                        ),
                    ])
                )

                elementos.append(rodape)

                # -----------------------------------------------------
                # GERA O PDF
                # -----------------------------------------------------

                doc.build(elementos)

                print(
                    "PDF gerado:",
                    arquivo
                )

                # -----------------------------------------------------
                # AGORA SIM APAGA AS IMAGENS TEMPORÁRIAS
                # -----------------------------------------------------

                for caminho_temp in imagens_temporarias:

                    try:

                        if os.path.exists(caminho_temp):
                            os.remove(caminho_temp)

                    except Exception:
                        pass

        except Exception as e:
            trata_excecao(e)
            raise

    def pode_enviar_pedidos(self):
        try:
            esta_tudo_certo = not self.produtos_sem_imagem

            if not esta_tudo_certo:
                return False

            return True

        except Exception as e:
            trata_excecao(e)
            raise

    def enviar_email(self):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_user, password)

            for cliente, destinatarios in self.destinatarios_clientes.items():

                nome_cliente = cliente.replace("/", "_").replace("\\", "_")

                arquivo = os.path.join(
                    os.getcwd(),
                    f"PI Pendentes - {nome_cliente}.pdf"
                )

                if not os.path.exists(arquivo):
                    continue

                subject = f"Pedidos Pendentes – {cliente}"

                msg = MIMEMultipart()
                msg["From"] = email_user
                msg["Subject"] = subject

                body = ""

                body += f"{saudacao}\n\n"
                body += (
                    f"Segue em anexo o relatório de Pedidos Pendentes "
                    f"do cliente {cliente}.\n\n"
                )
                body += f"{msg_final}"

                msg.attach(MIMEText(body, "plain"))

                with open(arquivo, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())

                encoders.encode_base64(part)

                nome_arquivo = f"PI Pendentes - {nome_cliente}.pdf"

                part.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=Header(
                        nome_arquivo,
                        "utf-8"
                    ).encode()
                )

                msg.attach(part)

                text = msg.as_string()

                server.sendmail(
                    email_user,
                    destinatarios,
                    text
                )

                print(f"Email enviado para {cliente}.")

            server.quit()

            self.apagar_arquivos_criados()

        except Exception as e:
            trata_excecao(e)
            raise

    def apagar_arquivos_criados(self):
        try:
            arquivos = [
                arquivo
                for arquivo in os.listdir(os.getcwd())
                if arquivo.startswith("PI Pendentes - ")
                   and arquivo.endswith(".pdf")
            ]

            for arquivo in arquivos:
                caminho = os.path.join(
                    os.getcwd(),
                    arquivo
                )

                if os.path.exists(caminho):
                    os.remove(caminho)
                    print("Arquivo apagado:", caminho)

        except Exception as e:
            trata_excecao(e)
            raise

    def localizar_logo_cliente(self, cliente):
        try:
            caminho_logo = rf"\\Publico\C\OP\Projetos\{cliente}.png"

            if os.path.exists(caminho_logo):
                return caminho_logo

            return None

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    try:
        rel = RelatorioPIPendentes()

        if rel.pode_enviar_pedidos():
            rel.gerar_pdf()
            rel.enviar_email()
        else:
            rel.enviar_email_produtos_sem_imagem()

    except Exception as e:
        trata_excecao(e)
        raise