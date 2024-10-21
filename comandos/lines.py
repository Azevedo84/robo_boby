import sys
from banco_dados.controle_erros import grava_erro_banco
from PyQt5.QtCore import QLocale, QRegExp
from PyQt5.QtGui import QDoubleValidator, QIntValidator, QRegExpValidator
from datetime import date
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


def validador_decimal(nome_line, numero, decimal=3):
    try:
        validator = QDoubleValidator(0, numero, decimal, nome_line)
        locale = QLocale("pt_BR")
        validator.setLocale(locale)
        validator.setBottom(0.001)

        nome_line.setValidator(validator)

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def validador_inteiro(nome_line, numero):
    try:
        validator = QIntValidator(0, numero, nome_line)
        locale = QLocale("pt_BR")
        validator.setLocale(locale)
        nome_line.setValidator(validator)

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def validador_so_numeros(nome_line):
    try:
        validator = QRegExpValidator(QRegExp(r'\d+'), nome_line)
        nome_line.setValidator(validator)

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)


def definir_data_atual(nome_line):
    try:
        data_hoje = date.today()
        nome_line.setDate(data_hoje)

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
