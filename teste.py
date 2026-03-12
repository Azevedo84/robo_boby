import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton,
    QLineEdit, QLabel, QHeaderView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor


class TelaEntradaNF(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Entrada de NF de Compra")
        self.resize(1100, 500)

        layout_principal = QVBoxLayout(self)

        # ===== TABELA =====
        self.tabela = QTableWidget()
        self.tabela.setColumnCount(6)
        self.tabela.setHorizontalHeaderLabels([
            "Produto",
            "Qtd NF",
            "Qtd OC",
            "Qtd Conferida",
            "Destino",
            "Status"
        ])

        self.tabela.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout_principal.addWidget(self.tabela)

        # ===== BOTÃO CONFIRMAR =====
        self.btn_confirmar = QPushButton("Confirmar Entrada")
        self.btn_confirmar.setEnabled(False)
        self.btn_confirmar.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-size: 16px;
                padding: 8px;
                border-radius: 6px;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)

        layout_principal.addWidget(self.btn_confirmar)

        # Dados simulados
        self.carregar_dados_exemplo()

    # ==========================================================
    # CARREGA ITENS DE EXEMPLO
    # ==========================================================
    def carregar_dados_exemplo(self):

        dados = [
            ("Parafuso 5/16", 10, 10, "OP-4587"),
            ("Chapa de Aço 3mm", 5, 5, "OP-4587"),
            ("Tubo de Ferro 2\"", 8, 8, "Estoque")
        ]

        self.tabela.setRowCount(len(dados))

        for linha, (produto, qtd_nf, qtd_oc, destino) in enumerate(dados):

            self.tabela.setItem(linha, 0, QTableWidgetItem(produto))
            self.tabela.setItem(linha, 1, QTableWidgetItem(str(qtd_nf)))
            self.tabela.setItem(linha, 2, QTableWidgetItem(str(qtd_oc)))
            self.tabela.setItem(linha, 4, QTableWidgetItem(destino))

            # Criar célula personalizada para conferência
            widget = self.criar_widget_conferencia(linha, qtd_nf)
            self.tabela.setCellWidget(linha, 3, widget)

            # Status inicial
            status = QTableWidgetItem("Pendente")
            status.setBackground(QColor("#f8d7da"))  # vermelho claro
            self.tabela.setItem(linha, 5, status)

    # ==========================================================
    # CRIA QLineEdit + BOTÃO ✔ DENTRO DA CÉLULA
    # ==========================================================
    def criar_widget_conferencia(self, linha, valor_inicial):

        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(2, 2, 2, 2)

        line = QLineEdit()
        line.setText(str(valor_inicial))
        line.setFixedWidth(60)

        btn_ok = QPushButton("✔")
        btn_ok.setFixedWidth(30)
        btn_ok.setStyleSheet("background-color: #ffc107;")

        # ENTER valida
        line.returnPressed.connect(lambda: self.validar_item(linha))

        # Clique valida
        btn_ok.clicked.connect(lambda: self.validar_item(linha))

        layout.addWidget(line)
        layout.addWidget(btn_ok)

        return container

    # ==========================================================
    # VALIDA ITEM
    # ==========================================================
    def validar_item(self, linha):

        qtd_nf = int(self.tabela.item(linha, 1).text())

        widget = self.tabela.cellWidget(linha, 3)
        line = widget.layout().itemAt(0).widget()
        qtd_conferida = int(line.text())

        status_item = self.tabela.item(linha, 5)

        if qtd_conferida == qtd_nf:
            status_item.setText("OK")
            status_item.setBackground(QColor("#d4edda"))  # verde claro
            self.pintar_linha(linha, QColor("#e9f7ef"))
        else:
            status_item.setText("Divergente")
            status_item.setBackground(QColor("#fff3cd"))  # amarelo
            self.pintar_linha(linha, QColor("#fff8e1"))

        self.verificar_se_pode_confirmar()

    # ==========================================================
    # PINTA LINHA
    # ==========================================================
    def pintar_linha(self, linha, cor):

        for col in range(self.tabela.columnCount()):
            item = self.tabela.item(linha, col)
            if item:
                item.setBackground(cor)

    # ==========================================================
    # VERIFICA SE TODOS ITENS ESTÃO OK
    # ==========================================================
    def verificar_se_pode_confirmar(self):

        todos_ok = True

        for linha in range(self.tabela.rowCount()):
            status = self.tabela.item(linha, 5).text()
            if status != "OK":
                todos_ok = False
                break

        self.btn_confirmar.setEnabled(todos_ok)


# ==========================================================
# EXECUTAR
# ==========================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = TelaEntradaNF()
    janela.show()
    sys.exit(app.exec_())
