import re
import unicodedata

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
    "testes"
}

def extrair_tipo_desenho(ref):
    if not isinstance(ref, str):
        return None

    partes = ref.rsplit(".", 1)

    if len(partes) != 2:
        return None

    final = partes[1]

    # valida: tem que ser exatamente 2 dígitos
    if final.isdigit() and len(final) == 2:
        return final

    return None

def padronizar_caminho(caminho: str) -> str | None:
    if not caminho:
        return None

    original = caminho

    caminho = caminho.strip()
    caminho = caminho.replace("/", "\\")
    caminho = caminho.lower()

    if not caminho.startswith("\\\\publico\\c\\"):
        print(f"⚠️ FORA DO PADRÃO: {original}")
        return None

    return caminho

def corrigir_caminho_inventor(caminho: str) -> str:
    if not caminho:
        return caminho

    c = caminho.strip().replace("/", "\\")
    c_lower = c.lower()

    # 🎯 REGRA SEGURA: só mexe se começar exatamente com isso
    if c_lower.startswith("\\\\publico\\inventor\\"):
        return "\\\\publico\\c\\inventor\\" + c[ len("\\\\publico\\inventor\\") : ]

    return c

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


def extrair_referencia(referencia):
    if not referencia:
        return None

    s = re.sub(r"[^\d.]", "", referencia)
    s = re.sub(r"\.+$", "", s)

    return s if s else None

def normalizar_texto(texto):
    if not texto:
        return ""

    # 1. remove acentos (ç, ã, é…)
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(c for c in texto if not unicodedata.combining(c))

    # 2. maiúsculo
    texto = texto.upper()

    # 3. remove espaços duplicados
    texto = re.sub(r'\s+', ' ', texto)

    # 4. strip
    texto = texto.strip()

    return texto

def formatar_descricao_para_inventor(texto):
    if not texto:
        return ""

    minusculas = {"de", "da", "do", "das", "dos", "e"}

    palavras = texto.lower().split()
    resultado = []

    for i, p in enumerate(palavras):
        if i > 0 and p in minusculas:
            resultado.append(p)
        else:
            resultado.append(p.capitalize())

    return " ".join(resultado)
