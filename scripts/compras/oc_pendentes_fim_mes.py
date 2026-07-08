import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta, conecta_robo
from core.email_service import dados_email
from collections import defaultdict
from datetime import datetime
from core.erros import trata_excecao
from datetime import date, timedelta
from email.mime.base import MIMEBase
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepTogether,
    Image
)
from reportlab.lib.enums import TA_LEFT


class RelatorioOCPendentes:
    def __init__(self):
        self.styles = getSampleStyleSheet()

        self.destinatario = ['<maquinas@unisold.com.br>']

    def buscar_dados(self):
        try:
            cursor = conecta.cursor()
            cursor.execute("""
                SELECT
                    oc.numero,
                    oc.data,
                    forn.razao,
                    prod.codigo,
                    prod.descricao,
                    COALESCE(prod.obs,''),
                    prod.unidade,
                    prodoc.quantidade,
                    prodoc.produzido,
                    prodoc.unitario,
                    prodoc.ipi,
                    prodoc.dataentrega
                FROM ordemcompra oc
                JOIN produtoordemcompra prodoc ON oc.id = prodoc.mestre
                JOIN produto prod ON prod.id = prodoc.produto
                JOIN fornecedores forn ON forn.id = oc.fornecedor
                WHERE oc.entradasaida='E'
                  AND oc.status='A'
                  AND prodoc.produzido < prodoc.quantidade
                ORDER BY oc.numero
            """)
            return cursor.fetchall()

        except Exception as e:
            trata_excecao(e)
            raise

    def agrupar(self, dados):
        try:
            grupos = defaultdict(list)

            for linha in dados:
                (
                    oc,
                    emissao,
                    fornecedor,
                    codigo,
                    descricao,
                    ref,
                    um,
                    qtde,
                    produzido,
                    unit,
                    ipi,
                    entrega
                ) = linha

                falta = float(qtde) - float(produzido)

                valor_mercadoria = float(qtde) * float(unit)
                valor_ipi = valor_mercadoria * (float(ipi) / 100)
                total_oc = valor_mercadoria + valor_ipi

                valor_pendente = falta * float(unit)
                ipi_pendente = valor_pendente * (float(ipi) / 100)
                total_pendente = valor_pendente + ipi_pendente

                grupos[oc].append({
                    "emissao": emissao,
                    "fornecedor": fornecedor,
                    "codigo": codigo,
                    "descricao": descricao,
                    "ref": ref,
                    "um": um,
                    "falta": falta,
                    "unit": float(unit),

                    "valor_mercadoria": valor_mercadoria,
                    "valor_ipi": valor_ipi,

                    "total_oc": total_oc,
                    "total_pendente": total_pendente,

                    "entrega": entrega
                })

            return grupos

        except Exception as e:
            trata_excecao(e)
            raise

    def moeda(self, valor):
        try:
            return f'R$ {valor:,.2f}'.replace(",", "X").replace(".", ",").replace("X",".")

        except Exception as e:
            trata_excecao(e)
            raise

    def gerar_pdf(self, arquivo="Relatorio_OC_Pendentes.pdf"):
        try:
            dados = self.buscar_dados()
            grupos = self.agrupar(dados)

            doc = SimpleDocTemplate(
                arquivo,
                pagesize=(21*cm,29.7*cm),
                leftMargin=1.2*cm,
                rightMargin=1.2*cm,
                topMargin=1.2*cm,
                bottomMargin=1.2*cm
            )

            story=[]

            titulo = self.styles["Title"]
            titulo.textColor = colors.HexColor("#1F4E79")
            titulo.alignment = TA_LEFT

            subtitulo = self.styles["Heading2"]
            subtitulo.alignment = TA_LEFT

            data_style = ParagraphStyle(
                "Data",
                parent=self.styles["Normal"],
                fontSize=9,
                textColor=colors.grey
            )

            data_style.alignment = TA_CENTER

            # Logo
            base = os.path.dirname(
                os.path.dirname(
                    os.path.dirname(os.path.abspath(__file__))
                )
            )

            caminho_logo = os.path.join(
                base,
                "files",
                "imagens",
                "logo_erp.png"
            )

            logo = Image(caminho_logo)
            logo.drawHeight = 1.8 * cm
            logo.drawWidth = 5.0 * cm

            story.append(Spacer(1, 0.3 * cm))
            story.append(logo)
            story.append(Spacer(1, 0.25 * cm))

            titulo.alignment = TA_CENTER

            story.append(
                Paragraph(
                    "<b>RELATÓRIO DE ORDENS DE COMPRA PENDENTES</b>",
                    titulo
                )
            )

            story.append(
                Paragraph(
                    f"Emitido em {datetime.now().strftime('%d/%m/%Y às %H:%M')}",
                    data_style
                )
            )

            story.append(Spacer(1, 0.2 * cm))

            # Linha azul
            linha = Table([[""]], colWidths=[16 * cm])
            linha.setStyle(TableStyle([
                ("LINEBELOW", (0, 0), (0, 0), 2, colors.HexColor("#1F4E79")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]))

            story.append(linha)
            story.append(Spacer(1, 0.5 * cm))

            # =====================================================
            # RESUMO POR FORNECEDOR
            # =====================================================

            fornecedores = defaultdict(list)

            for oc, itens in grupos.items():
                fornecedor = itens[0]["fornecedor"]

                if oc not in fornecedores[fornecedor]:
                    fornecedores[fornecedor].append(oc)

            story.append(Paragraph("<b>FORNECEDORES COM PENDÊNCIAS</b>", self.styles["Heading2"]))
            story.append(Spacer(1, 0.3 * cm))

            tabela = [["Fornecedor", "OCs Pendentes"]]

            lista = []

            for fornecedor, ocs in fornecedores.items():
                lista.append((min(ocs), fornecedor, ocs))

            lista.sort(key=lambda x: x[0])

            for _, fornecedor, ocs in lista:
                lista_oc = ", ".join(str(x) for x in sorted(fornecedores[fornecedor]))

                tabela.append([
                    fornecedor,
                    lista_oc
                ])

            tb = Table(tabela, colWidths=[11 * cm, 6 * cm])

            tb.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 7),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]))

            story.append(tb)
            story.append(Spacer(1, 0.7 * cm))

            for oc,itens in grupos.items():

                primeiro=itens[0]

                valor_oc=sum(i["total_oc"] for i in itens)
                valor_p=sum(i["total_pendente"] for i in itens)

                bloco = []

                titulo_oc = Table(
                    [[f"ORDEM DE COMPRA Nº {oc}"]],
                    colWidths=[17 * cm]
                )

                titulo_oc.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1F4E79")),
                    ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                    ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 12),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                    ("TOPPADDING", (0, 0), (-1, -1), 7),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ]))

                bloco.append(titulo_oc)
                dados_oc = [
                    [
                        Paragraph(
                            f"<font color='#1F4E79'><b>Fornecedor:</b></font> {primeiro['fornecedor']}",
                            self.styles["BodyText"]
                        ),
                        Paragraph(
                            f"<font color='#1F4E79'><b>Valor Total:</b></font> {self.moeda(valor_oc)}",
                            self.styles["BodyText"]
                        )
                    ],
                    [
                        Paragraph(
                            f"<font color='#1F4E79'><b>Emissão:</b></font> "
                            f"{primeiro['emissao'].strftime('%d/%m/%Y') if primeiro['emissao'] else ''}",
                            self.styles["BodyText"]
                        ),
                        Paragraph(
                            f"<font color='#1F4E79'><b>Valor Pendente:</b></font> {self.moeda(valor_p)}",
                            self.styles["BodyText"]
                        )
                    ]
                ]

                quadro = Table(dados_oc, colWidths=[10 * cm, 7 * cm])

                quadro.setStyle(TableStyle([
                    ("BOX", (0, 0), (-1, -1), 0, colors.white),
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7F9FC")),
                    ("LINEBELOW", (0, 1), (-1, 1), 0, colors.white),
                    ("INNERGRID", (0, 0), (-1, 0), 0.3, colors.HexColor("#D9D9D9")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
                    ("LINEBELOW", (0, 1), (-1, 1), 0.4, colors.HexColor("#BFBFBF")),
                ]))

                bloco.append(quadro)

                tabela=[[
                    "Código",
                    "Descrição",
                    "Ref.",
                    "UM",
                    "Falta",
                    "Unit.",
                    "Total"
                ]]

                for item in itens:
                    tabela.append([
                        item["codigo"],
                        item["descricao"],
                        item["ref"],
                        item["um"],
                        f'{item["falta"]:.3f}',
                        self.moeda(item["unit"]),
                        self.moeda(item["total_pendente"])
                    ])

                tb = Table(
                    tabela,
                    colWidths=[
                        2.0 * cm,  # Código
                        5.3 * cm,  # Descrição
                        2.4 * cm,  # Ref.
                        1.0 * cm,  # UM
                        1.7 * cm,  # Falta
                        2.2 * cm,  # Unit.
                        2.4 * cm  # Total
                    ]
                )
                estilo=[
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF1FB")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                    ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#1F4E79")),
                    ("LINEBELOW", (0, 1), (-1, -1), 0.15, colors.HexColor("#DDDDDD")),
                    ("LINEABOVE", (0, 0), (-1, 0), 0.4, colors.HexColor("#D0D0D0")),
                    ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                    ("FONTSIZE",(0,0),(-1,-1),8),
                    ("ALIGN", (0, 0), (0, -1), "CENTER"),  # Código
                    ("ALIGN", (2, 0), (3, -1), "CENTER"),  # Ref e UM
                    ("ALIGN", (4, 0), (6, -1), "RIGHT"),  # Falta, Unit e Total
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("BOTTOMPADDING",(0,0),(-1,0),6),
                ]

                for l in range(1,len(tabela)):
                    if l%2==0:
                        estilo.append(("BACKGROUND",(0,l),(-1,l),colors.whitesmoke))

                estilo.extend([
                    ("TOPPADDING", (0, 0), (-1, 0), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
                ])

                tb.setStyle(TableStyle(estilo))
                bloco.append(tb)
                bloco.append(Spacer(1, 0.6 * cm))

                story.append(KeepTogether(bloco))

            doc.build(story)
            print("PDF gerado:",arquivo)

        except Exception as e:
            trata_excecao(e)
            raise

    def pode_enviar_oc_fim_mes(self):
        try:
            referencia = self.referencia_pendente()

            if referencia is None:
                return False

            cursor = conecta_robo.cursor()
            cursor.execute("""
                SELECT FIRST 1 ID
                FROM ENVIA_OC_FIM_MES
                WHERE REFERENCIA = ?
            """, (referencia,))

            registro = cursor.fetchone()

            if registro:
                return False

            return True

        except Exception as e:
            trata_excecao(e)
            raise

    def gravar_envio(self):
        try:
            referencia = self.referencia_pendente()

            cursor = conecta_robo.cursor()
            cursor.execute("""
                INSERT INTO ENVIA_OC_FIM_MES
                (REFERENCIA, DATA_ENVIO)
                VALUES
                (?, CURRENT_TIMESTAMP)
            """, (referencia,))
            conecta_robo.commit()

            print("Envio registrado com sucesso.")

        except Exception as e:
            trata_excecao(e)
            raise

    def referencia_pendente(self):
        try:
            hoje = date.today()

            # Sábado ou domingo nunca envia
            if hoje.weekday() >= 5:
                return None

            # Último dia do mês atual
            if hoje.month == 12:
                ultimo = date(hoje.year + 1, 1, 1) - timedelta(days=1)
            else:
                ultimo = date(hoje.year, hoje.month + 1, 1) - timedelta(days=1)

            # Último dia útil
            while ultimo.weekday() >= 5:
                ultimo -= timedelta(days=1)

            # Se hoje já passou do último dia útil,
            # a competência pendente é o mês atual.
            if hoje >= ultimo:
                return hoje.strftime("%m/%Y")

            # Caso contrário, a competência pendente é o mês anterior.
            referencia = hoje.replace(day=1) - timedelta(days=1)
            return referencia.strftime("%m/%Y")

        except Exception as e:
            trata_excecao(e)
            raise

    def enviar_email(self, arquivo_pdf):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            referencia = self.referencia_pendente()

            assunto = f"Ordens de Compra Pendentes - {referencia}"

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['To'] = ", ".join(self.destinatario)
            msg['Subject'] = assunto

            body = (
                f"{saudacao}\n\n"
                f"Segue em anexo o relatório de Ordens de Compra Pendentes do mês de {referencia}.\n\n"
                f"{msg_final}"
            )

            msg.attach(MIMEText(body, "plain"))

            # ==========================
            # Anexa o PDF
            # ==========================
            with open(arquivo_pdf, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            encoders.encode_base64(part)

            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(arquivo_pdf)}"'
            )

            msg.attach(part)

            # ==========================
            # Envia email
            # ==========================
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(
                email_user,
                self.destinatario,
                msg.as_string()
            )

            server.quit()

            # ==========================
            # Grava no banco
            # ==========================
            self.gravar_envio()

            print("Relatório enviado com sucesso.")

        except Exception as e:
            trata_excecao(e)
            raise

if __name__ == "__main__":
    try:
        rel = RelatorioOCPendentes()

        if rel.pode_enviar_oc_fim_mes():
            arquivo = "Relatorio_OC_Pendentes.pdf"

            rel.gerar_pdf(arquivo)
            rel.enviar_email(arquivo)

        else:
            print("Hoje não é o último dia útil do mês ou o relatório já foi enviado.")

    except Exception as e:
        trata_excecao(e)
        raise
