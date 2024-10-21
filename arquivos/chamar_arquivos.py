import sys
import os
import inspect

nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
nome_arquivo = os.path.basename(nome_arquivo_com_caminho)


def definir_caminho_arquivo(caminho):
    def get_file_path(file_name):
        if getattr(sys, 'frozen', False):
            return os.path.join(sys._MEIPASS, file_name)
        else:
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

    caminho_arquivo = get_file_path(caminho)

    return caminho_arquivo
