import sys
from banco_dados.controle_erros import grava_erro_banco
from arquivos.chamar_arquivos import definir_caminho_arquivo
from comandos.cores import cabecalho_tela, widgets, textos, fonte_botao, fundo_botao, widgets_escuro
from comandos.cores import fundo_tela, fundo_tela_menu, cor_branco
from PyQt5.QtWidgets import QDesktopWidget, QPushButton, QVBoxLayout, QLabel, QSizePolicy
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import QPixmap, QFont, QIcon
import os
import inspect
import traceback

nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
nome_arquivo = os.path.basename(nome_arquivo_com_caminho)


def trata_excecao(nome_funcao, mensagem, arquivo, excecao):
    try:
        tb = traceback.extract_tb(excecao)
        num_linha_erro = tb[-1][1]

        traceback.print_exc()
        print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')

        grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

    except Exception as e:
        nome_funcao_trat = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        tb = traceback.extract_tb(exc_traceback)
        num_linha_erro = tb[-1][1]
        print(f'Houve um problema no arquivo: {nome_arquivo} na função: "{nome_funcao_trat}"\n'
              f'{e} {num_linha_erro}')
        grava_erro_banco(nome_funcao_trat, e, nome_arquivo, num_linha_erro)


def icone(self, nome_imagem):
    try:
        camino = os.path.join('..', 'arquivos', 'icones', nome_imagem)
        caminho_arquivo = definir_caminho_arquivo(camino)

        self.setWindowIcon(QIcon(caminho_arquivo))

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def criar_botao(nome_imagem, texto_botao):
    try:
        botao = QPushButton()
        botao.setFixedSize(60, 60)
        botao.setStyleSheet(f"background-color: {cor_branco};")

        botao.setCursor(Qt.PointingHandCursor)

        layout = QVBoxLayout()
        layout.setSpacing(3)

        img_label = QLabel()
        camino = os.path.join('..', 'arquivos', 'icones', nome_imagem)
        caminho_imagem = definir_caminho_arquivo(camino)

        # Verificação se a imagem existe
        if not os.path.exists(caminho_imagem):
            print(f"Erro: A imagem {caminho_imagem} não foi encontrada.")
            return None

        icon = QPixmap(caminho_imagem)
        img_label.setPixmap(icon.scaled(QSize(25, 25), aspectRatioMode=Qt.KeepAspectRatio))
        img_label.setAlignment(Qt.AlignCenter)

        text_label = QLabel(texto_botao)
        text_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setPointSize(8)
        text_label.setFont(font)

        layout.addWidget(img_label)
        layout.addWidget(text_label)

        botao.setLayout(layout)

        return botao

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def editar_botao(botao, nome_imagem, texto_botao, tamanho):
    try:
        layout = QVBoxLayout()
        layout.setSpacing(3)

        img_label = QLabel()
        camino = os.path.join('..', 'arquivos', 'icones', nome_imagem)
        caminho_imagem = definir_caminho_arquivo(camino)

        if not os.path.exists(caminho_imagem):
            print(f"Erro: A imagem {caminho_imagem} não foi encontrada.")
            return None

        icon = QPixmap(caminho_imagem)
        img_label.setPixmap(icon.scaled(QSize(tamanho, tamanho), aspectRatioMode=Qt.KeepAspectRatio))
        img_label.setAlignment(Qt.AlignCenter)

        text_label = QLabel(texto_botao)
        text_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setPointSize(8)
        text_label.setFont(font)

        layout.addWidget(img_label)
        layout.addWidget(text_label)

        botao.setLayout(layout)

        # Ajustando o tamanho mínimo do botão
        botao.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def tamanho_aplicacao(self):
    try:
        monitor = QDesktopWidget().screenGeometry()
        monitor_width = monitor.width()
        monitor_height = monitor.height()

        if monitor_width > 1919 and monitor_height > 1079:
            interface_width = 1300
            interface_height = 750

        elif monitor_width > 1365 and monitor_height > 767:
            interface_width = 1050
            interface_height = 585
        else:
            interface_width = monitor_width - 165
            interface_height = monitor_height - 90

        x = (monitor_width - interface_width) // 2
        y = (monitor_height - interface_height) // 2

        self.setGeometry(x, y, interface_width, interface_height)

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def cor_widget_cab(nome_widget):
    try:
        nome_widget.setStyleSheet(f"background-color: {cabecalho_tela};")

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def cor_widget(nome_widget):
    try:
        nome_widget.setStyleSheet(f"background-color: {widgets};")

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def cor_widget_escuro(nome_widget):
    try:
        nome_widget.setStyleSheet(f"background-color: {widgets_escuro};")

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def cor_fonte(nome_campo):
    try:
        nome_campo.setStyleSheet(f"color: {textos};")

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def cor_btn(nome_botao):
    try:
        nome_botao.setStyleSheet(f"background-color: {fundo_botao}; color: {fonte_botao};")

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def cor_fundo_tela(nome_widget):
    try:
        nome_widget.setStyleSheet(f"background-color: {fundo_tela};")

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def cor_fundo_tela_menu(nome_widget):
    try:
        nome_widget.setStyleSheet(f"background-color: {fundo_tela_menu};")

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
