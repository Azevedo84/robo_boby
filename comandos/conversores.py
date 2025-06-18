import sys
from banco_dados.controle_erros import grava_erro_banco
import os
import inspect
import locale
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


def valores_para_float(string):
    try:
        if string is None:
            return 0.00

        string_certo = str(string)

        if "R$ " in string_certo:
            limpa_string = string_certo.replace("R$ ", '')
        elif "%" in string_certo:
            limpa_string = string_certo.replace("%", '')
        else:
            limpa_string = string_certo

        if limpa_string:
            if "," in limpa_string:
                string_com_ponto = limpa_string.replace(',', '.')
                valor_float = float(string_com_ponto)
            else:
                valor_float = float(limpa_string)
        else:
            valor_float = 0.00

        return valor_float

    except ValueError:
        # Se a conversão falhar, retornar 0.00 e logar o erro
        nome_funcao = inspect.currentframe().f_code.co_name
        print(f"Valor inválido para conversão: {string}", nome_funcao, nome_arquivo)
        return 0.00
    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
        return 0.00


def valores_para_virgula(string):
    try:
        if "R$ " in string:
            limpa_string = string.replace("R$ ", '')
        elif "%" in string:
            limpa_string = string.replace("%", '')
        else:
            limpa_string = string

        if limpa_string:
            if "." in limpa_string:
                string_com_virgula = limpa_string.replace('.', ',')
            else:
                string_com_virgula = limpa_string
        else:
            string_com_virgula = "0,00"

        return string_com_virgula

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
        return None


def float_para_moeda_reais(valor):
    try:
        valor_float = valores_para_float(valor)

        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

        valor_final = locale.currency(valor_float, grouping=True, symbol=True)

        return valor_final

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
        return None


def float_para_porcentagem(valor):
    try:
        if valor:
            ipi_2casas = ("%.2f" % valor)
            valor_string = valores_para_virgula(ipi_2casas)
            valor_final = valor_string + "%"
        else:
            valor_final = ""

        return valor_final

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
        return None


def moeda_reais_para_float(valor_moeda):
    try:
        # Remove o símbolo da moeda e os separadores de milhar
        valor_moeda = valor_moeda.replace('R$', '').replace('.', '').replace(',', '.')

        # Converte a string para float
        valor_float = float(valor_moeda.strip())

        return valor_float

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
        return None


def float_para_virgula(valor_float):
    try:
        if valor_float:
            limpa_string = str(valor_float)
        else:
            limpa_string = "0"

        if limpa_string:
            if "." in limpa_string:
                string_com_virgula = limpa_string.replace('.', ',')
            else:
                string_com_virgula = limpa_string
        else:
            string_com_virgula = "0,00"

        return string_com_virgula

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
        return None


def timestamp_brasileiro(data_e_tempo):
    try:
        if data_e_tempo:
            data_formatada = data_e_tempo.strftime("%d/%m/%Y %H:%M:%S")
        else:
            data_formatada = ""

        return data_formatada

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)
        return None
