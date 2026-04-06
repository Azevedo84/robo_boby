from core.erros import trata_excecao
import locale

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

    except Exception as e:
        trata_excecao(e)
        raise

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
        trata_excecao(e)
        raise

def float_para_moeda_reais(valor):
    try:
        valor_float = valores_para_float(valor)

        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

        valor_final = locale.currency(valor_float, grouping=True, symbol=True)

        return valor_final

    except Exception as e:
        trata_excecao(e)
        raise

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
        trata_excecao(e)
        raise

def moeda_reais_para_float(valor_moeda):
    try:
        # Remove o símbolo da moeda e os separadores de milhar
        valor_moeda = valor_moeda.replace('R$', '').replace('.', '').replace(',', '.')

        # Converte a string para float
        valor_float = float(valor_moeda.strip())

        return valor_float

    except Exception as e:
        trata_excecao(e)
        raise

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
        trata_excecao(e)
        raise

def timestamp_brasileiro(data_e_tempo):
    try:
        if data_e_tempo:
            data_formatada = data_e_tempo.strftime("%d/%m/%Y %H:%M:%S")
        else:
            data_formatada = ""

        return data_formatada

    except Exception as e:
        trata_excecao(e)
        raise

def data_banco_para_brasileiro(data_banco):
    try:
        data_brasil = data_banco.strftime("%d/%m/%Y")

        return data_brasil

    except Exception as e:
        trata_excecao(e)
        raise
