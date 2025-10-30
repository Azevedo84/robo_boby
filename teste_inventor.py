import os
import pandas as pd
import re
import win32com.client

# Função para buscar o arquivo
def encontrar_arquivo(caminho, arquivo):
    for raiz, diretorios, arquivos in os.walk(caminho):
        if arquivo in arquivos:
            return os.path.join(raiz, arquivo)
    return None


# Caminho do arquivo na área de trabalho
usuario = os.getlogin()
caminho = fr"C:\Users\{usuario}\Desktop\teste.xlsx"

# Lê a planilha
df = pd.read_excel(caminho, sheet_name="Planilha1")

# Garante que as colunas estejam corretas
df.columns = ["desenho"]

# Lista para guardar os desenhos não encontrados
nao_encontrados = []

# Caminho base de busca
caminho_inicial = r"\\PUBLICO\Inventor\1 - Folha A4"

# Inicializa o Inventor (apenas uma vez)
inv_app = win32com.client.Dispatch("Inventor.Application")
inv_app.Visible = False  # False para rodar em background

# Loop pelos desenhos
for _, linha in df.iterrows():
    desenho = str(linha["desenho"])

    # Limpa o texto, mantendo apenas números e pontos
    s = re.sub(r"[^\d.]", "", desenho)
    s = re.sub(r"\.+$", "", s)
    print(s)

    nome_arquivo = f"{s}.idw"

    # Busca o arquivo
    resultado = encontrar_arquivo(caminho_inicial, nome_arquivo)

    if resultado:
        print(f"Arquivo encontrado em: {resultado}")

        # Caminho do PDF na Área de Trabalho
        desktop_pdf = rf"C:\Users\{usuario}\Desktop\Nova_pasta\{s}.pdf"
        os.makedirs(os.path.dirname(desktop_pdf), exist_ok=True)

        # Abre o arquivo .idw
        doc = inv_app.Documents.Open(resultado)

        # Salva diretamente como PDF
        doc.SaveAs(desktop_pdf, True)

        # Fecha o documento
        doc.Close(True)

        print(f"✅ PDF gerado com sucesso em: {desktop_pdf}")
    else:
        print("Arquivo não encontrado.")
        nao_encontrados.append(s)

# Fecha o Inventor
inv_app.Quit()

# Se houver desenhos não encontrados, gera um Excel na área de trabalho
if nao_encontrados:
    df_nao = pd.DataFrame(nao_encontrados, columns=["Desenhos não encontrados"])
    caminho_excel_nao = fr"C:\Users\{usuario}\Desktop\desenhos_nao_encontrados.xlsx"
    df_nao.to_excel(caminho_excel_nao, index=False)
    print(f"\n❌ Lista de desenhos não encontrados salva em:\n{caminho_excel_nao}")
else:
    print("\n✅ Todos os desenhos foram encontrados e convertidos com sucesso!")
