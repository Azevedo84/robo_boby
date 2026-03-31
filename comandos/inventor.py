import re

pasta_arq = r"\\Publico\c\Inventor"
extensoes = (".iam", ".ipt", ".idw")
padrao_desenho = re.compile(r'\b\d+(\.\d+)+\b')
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
