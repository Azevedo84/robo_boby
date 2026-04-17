import os

# Caminho onde começar a busca
caminho_inicial = r"\\PUBLICO\Inventor\1 - Folha A4"
nome_arquivo = "57.01.226.03.idw"

# Função para buscar o arquivo
def encontrar_arquivo(caminho, arquivo):
    for raiz, diretorios, arquivos in os.walk(caminho):
        if arquivo in arquivos:
            return os.path.join(raiz, arquivo)
    return None

# Executa a busca
resultado = encontrar_arquivo(caminho_inicial, nome_arquivo)

if resultado:
    print(f"Arquivo encontrado em: {resultado}")
else:
    print("Arquivo não encontrado.")
