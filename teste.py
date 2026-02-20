import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QPushButton, QFrame, QTableWidget, QTableWidgetItem,
    QScrollArea, QSizePolicy
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class Card(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setObjectName("card")
        self.setFrameShape(QFrame.StyledPanel)

        layout = QVBoxLayout()
        self.setLayout(layout)

        titulo = QLabel(title)
        titulo.setObjectName("cardTitle")
        layout.addWidget(titulo)

        self.content = QVBoxLayout()
        layout.addLayout(self.content)


class DashboardNF(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Detalhes da Nota Fiscal")
        self.resize(600, 900)  # Ideal para tablet 10"

        main_layout = QVBoxLayout(self)

        # Scroll principal
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        main_layout.addWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)

        layout = QVBoxLayout(container)

        # =========================
        # HEADER
        # =========================
        header = QFrame()
        header.setObjectName("header")
        header_layout = QVBoxLayout(header)

        lbl_nf = QLabel("NF Nº: 12345  |  Série: 1")
        lbl_nf.setObjectName("headerText")

        lbl_data = QLabel("Data Emissão: 12/04/2024")
        lbl_total = QLabel("Valor Total: R$ 15.420,00")
        lbl_total.setObjectName("valorTotal")

        status = QLabel("PENDENTE")
        status.setObjectName("status")

        header_layout.addWidget(lbl_nf)
        header_layout.addWidget(lbl_data)
        header_layout.addWidget(lbl_total)
        header_layout.addWidget(status)

        layout.addWidget(header)

        # =========================
        # FORNECEDOR
        # =========================
        card_fornecedor = Card("Dados do Fornecedor")
        card_fornecedor.content.addWidget(QLabel("TechProdutos S.A."))
        card_fornecedor.content.addWidget(QLabel("CNPJ: 12.345.678/0001-90"))
        card_fornecedor.content.addWidget(QLabel("Tel: (11) 3333-4444"))
        card_fornecedor.content.addWidget(QLabel("Email: contato@tech.com"))

        layout.addWidget(card_fornecedor)

        # =========================
        # ITENS
        # =========================
        card_itens = Card("Itens da Nota Fiscal")

        tabela = QTableWidget()
        tabela.setColumnCount(4)
        tabela.setHorizontalHeaderLabels(["Código", "Produto", "Qtde", "Valor"])

        itens = [
            ("1001", "Notebook Lenovo", "5", "R$ 2.500,00"),
            ("2002", "Mouse Óptico", "10", "R$ 30,00"),
            ("3050", "Teclado Mecânico", "8", "R$ 150,00"),
        ]

        tabela.setRowCount(len(itens))

        for row, item in enumerate(itens):
            for col, dado in enumerate(item):
                tabela.setItem(row, col, QTableWidgetItem(dado))

        tabela.resizeColumnsToContents()
        card_itens.content.addWidget(tabela)

        layout.addWidget(card_itens)

        # =========================
        # RESUMO
        # =========================
        card_resumo = Card("Resumo da Nota")

        card_resumo.content.addWidget(QLabel("Total Produtos: R$ 15.400,00"))
        card_resumo.content.addWidget(QLabel("ICMS: R$ 1.200,00"))
        card_resumo.content.addWidget(QLabel("IPI: R$ 300,00"))
        card_resumo.content.addWidget(QLabel("Frete: R$ 100,00"))
        card_resumo.content.addWidget(QLabel("Desconto: R$ 80,00"))
        card_resumo.content.addWidget(QLabel("Valor Final: R$ 15.420,00"))

        layout.addWidget(card_resumo)

        # =========================
        # OBSERVAÇÕES
        # =========================
        card_obs = Card("Observações")
        card_obs.content.addWidget(QLabel("Verificar quantidades e conferir mercadoria."))

        layout.addWidget(card_obs)

        # =========================
        # BOTÕES
        # =========================
        botoes_layout = QHBoxLayout()

        btn_entrada = QPushButton("Dar Entrada")
        btn_aprovar = QPushButton("Aprovar")
        btn_rejeitar = QPushButton("Rejeitar")

        btn_entrada.setObjectName("btnAzul")
        btn_aprovar.setObjectName("btnVerde")
        btn_rejeitar.setObjectName("btnVermelho")

        botoes_layout.addWidget(btn_entrada)
        botoes_layout.addWidget(btn_aprovar)
        botoes_layout.addWidget(btn_rejeitar)

        layout.addLayout(botoes_layout)

        # =========================
        # ESTILO (QSS)
        # =========================
        self.setStyleSheet("""
        QWidget {
            font-size: 16px;
        }

        #header {
            background-color: #2c3e50;
            color: white;
            padding: 15px;
            border-radius: 10px;
        }

        #valorTotal {
            font-size: 20px;
            font-weight: bold;
        }

        #status {
            background-color: orange;
            padding: 5px;
            border-radius: 5px;
            max-width: 120px;
        }

        #card {
            background: #f4f6f7;
            padding: 15px;
            border-radius: 10px;
        }

        #cardTitle {
            font-weight: bold;
            font-size: 18px;
        }

        QPushButton {
            padding: 15px;
            border-radius: 8px;
            color: white;
            font-weight: bold;
        }

        #btnAzul { background-color: #2980b9; }
        #btnVerde { background-color: #27ae60; }
        #btnVermelho { background-color: #c0392b; }
        """)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DashboardNF()
    window.show()
    sys.exit(app.exec_())
