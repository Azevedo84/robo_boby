import sys
from banco_dados.controle_erros import grava_erro_banco
from comandos.cores import fundo_cabecalho_tab, fonte_cabecalho_tab, zebra_tab
from PyQt5.QtWidgets import QAbstractItemView, QTableWidget, QStyledItemDelegate, QTableWidgetItem
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt
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


def layout_cabec_tab(nome_tabela):
    try:
        nome_tabela.horizontalHeader().setStyleSheet(
            f"QHeaderView::section {{ "
            f"background-color: {fundo_cabecalho_tab}; "
            f"font-weight: bold; "
            f"color: {fonte_cabecalho_tab}; }}"
        )

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def lanca_tabela(nome_tabela, dados_tab, altura_linha=25, zebra=True, largura_auto=True, bloqueia_texto=True):
    try:
        linhas_est = (len(dados_tab))
        colunas_est = (len(dados_tab[0]))
        nome_tabela.setRowCount(linhas_est)
        nome_tabela.setColumnCount(colunas_est)

        for i, linha in enumerate(dados_tab):
            nome_tabela.setRowHeight(i, altura_linha)
            for j, dado in enumerate(linha):
                item = QTableWidgetItem(str(dado))
                nome_tabela.setItem(i, j, item)
                alinha_cetralizado = AlignDelegate(nome_tabela)
                nome_tabela.setItemDelegateForColumn(j, alinha_cetralizado)

        nome_tabela.setSelectionBehavior(QAbstractItemView.SelectRows)
        nome_tabela.setSelectionBehavior(QAbstractItemView.SelectRows)

        if largura_auto:
            nome_tabela.resizeColumnsToContents()

        if bloqueia_texto:
            nome_tabela.setEditTriggers(QTableWidget.NoEditTriggers)

        if zebra:
            for row in range(nome_tabela.rowCount()):
                if row % 2 == 0:
                    for col in range(nome_tabela.columnCount()):
                        item = nome_tabela.item(row, col)
                        item.setBackground(QColor(zebra_tab))

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def extrair_tabela(nome_tabela):
    try:
        lista_final_itens = []

        total_linhas = nome_tabela.rowCount()
        if total_linhas:
            total_colunas = nome_tabela.columnCount()
            lista_final_itens = []
            linha = []
            for row in range(total_linhas):
                for column in range(total_colunas):
                    widget_item = nome_tabela.item(row, column)
                    if widget_item is not None:
                        lista_item = widget_item.text()
                        linha.append(lista_item)
                        if len(linha) == total_colunas:
                            lista_final_itens.append(linha)
                            linha = []
        return lista_final_itens

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


class AlignDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super(AlignDelegate, self).initStyleOption(option, index)
        option.displayAlignment = Qt.AlignCenter
