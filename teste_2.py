import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QFrame, QTableWidget, QTableWidgetItem,
    QPushButton, QSpacerItem, QSizePolicy
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class TelaProdutos(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("ERP Suzuki - Produtos do Fornecedor")
        self.resize(900, 500)

        self.setStyleSheet("""
            QWidget {
                background-color: #f4f6f9;
                font-family: Segoe UI;
                font-size: 10pt;
            }

            QTableWidget {
                background-color: white;
                border: 1px solid #dcdcdc;
                gridline-color: #eeeeee;
            }

            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 4px;
                border: none;
                font-weight: bold;
            }

            QPushButton {
                background-color: #2d89ef;
                color: white;
                border-radius: 6px;
                padding: 6px 12px;
            }

            QPushButton:hover {
                background-color: #1b5fbd;
            }
        """)

        layout_principal = QVBoxLayout(self)
        layout_principal.setContentsMargins(20, 20, 20, 20)
        layout_principal.setSpacing(15)

        # 🔹 Título com linhas laterais
        layout_principal.addWidget(self.criar_titulo("Produtos do Fornecedor"))

        # 🔹 Tabela
        self.tabela = QTableWidget()
        self.tabela.setColumnCount(4)
        self.tabela.setHorizontalHeaderLabels(
            ["Código", "Descrição", "Quantidade", "Valor"]
        )
        self.tabela.horizontalHeader().setStretchLastSection(True)
        self.tabela.setRowCount(5)

        # Dados exemplo
        dados = [
            ("001", "Parafuso 10mm", "150", "0,50"),
            ("002", "Porca 10mm", "200", "0,30"),
            ("003", "Arruela 10mm", "300", "0,20"),
            ("004", "Chapa Aço", "50", "25,00"),
            ("005", "Tinta Industrial", "20", "120,00"),
        ]

        for linha, valores in enumerate(dados):
            for coluna, valor in enumerate(valores):
                self.tabela.setItem(linha, coluna, QTableWidgetItem(valor))

        layout_principal.addWidget(self.tabela)

        # 🔹 Linha separadora inferior
        linha = QFrame()
        linha.setFixedHeight(1)
        linha.setStyleSheet("background-color: #dcdcdc;")
        layout_principal.addWidget(linha)

        # 🔹 Botões
        layout_botoes = QHBoxLayout()
        layout_botoes.addSpacerItem(
            QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        )

        btn_adicionar = QPushButton("Adicionar")
        btn_remover = QPushButton("Remover")
        btn_fechar = QPushButton("Fechar")

        btn_fechar.clicked.connect(self.close)

        layout_botoes.addWidget(btn_adicionar)
        layout_botoes.addWidget(btn_remover)
        layout_botoes.addWidget(btn_fechar)

        layout_principal.addLayout(layout_botoes)

    # 🔹 Método para criar o título moderno com linhas laterais
    def criar_titulo(self, texto):
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)

        linha1 = QFrame()
        linha1.setFixedHeight(1)
        linha1.setStyleSheet("background-color: #dcdcdc;")

        titulo = QLabel(texto)
        fonte = QFont()
        fonte.setBold(True)
        fonte.setPointSize(11)
        titulo.setFont(fonte)
        titulo.setStyleSheet("color: #333333;")
        titulo.setAlignment(Qt.AlignCenter)

        linha2 = QFrame()
        linha2.setFixedHeight(1)
        linha2.setStyleSheet("background-color: #dcdcdc;")

        layout.addWidget(linha1)
        layout.addWidget(titulo)
        layout.addWidget(linha2)

        return container


if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = TelaProdutos()
    janela.show()
    sys.exit(app.exec_())