import re

pasta_arq = r"\\Publico\c\Inventor"
extensoes = (".iam", ".ipt", ".idw")
padrao_desenho = re.compile(r'\b\d+(\.\d+)+\b')
padrao_terceiro_09 = re.compile(r'^\d{2}\.\d{2}\.\d{3}\.09$')
ignorar_pastas = {
    "oldversions",
    "design data",
    "content center",
    ".svn",
    ".tmp.drivedownload",
    ".tmp.driveupload",
    "00 - temporario",
    "alterações",
    "churrasqueira",
    "componentes importados",
    "ztemp",
    "teste"
}

def normalizar_caminho(caminho):
    caminho = caminho.strip()
    caminho = caminho.replace("/", "\\")
    caminho = caminho.lower()

    # padrão da rede
    caminho = caminho.replace("\\publico\\inventor", "\\publico\\c\\inventor")

    return caminho

def definir_classificacao(caminho, nome_sem_ext):
    # prioridade 1
    if "\\inventor\\biblioteca" in caminho:
        return "TERCEIROS"

    # prioridade 2
    if padrao_terceiro_09.match(nome_sem_ext):
        return "TERCEIROS"

    # prioridade 3
    if padrao_desenho.search(nome_sem_ext):
        return "NOSSO"

    # fallback
    return "TERCEIROS"
