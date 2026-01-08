from PyPDF2 import PdfReader

pasta = r'C:/ordens/'
arquivo = "OC 18860 Delta Laser.pdf"
caminho = pasta + arquivo

reader = PdfReader(caminho)

texto = ""

for page in reader.pages:
    page_content = page.extract_text()
    if page_content:
        texto += page_content
        print(texto)

import re

oc_match = re.search(r'O\.C\.: \s*(\d+)', texto)
num_oc = oc_match.group(1) if oc_match else None
print("Nº OC: ", num_oc)

emissao = re.search(r'Emissão:\s*(\d{2}/\d{2}/\d{4})', texto)
data_emissao = emissao.group(1) if emissao else None
print("Data Emissão: ", data_emissao)

codigo_fornecedor = re.search(r'C[oó]digo:\s*(\d+)', texto)
codigo_forn = codigo_fornecedor.group(1) if codigo_fornecedor else None
print("Cód. Fornecedor: ", codigo_forn)

# --- Itens ---
padrao_item = re.findall(
    r'''
        \n\d+\s+[A-Z]\s+[\d.]+\s+      # nº item + tipo + referência
        (\d+)\s+                       # código produto
        .*?                            # descrição (ignorada)
        \nNCM:.*?\nLC:.*?\n            # linhas fixas
        \s*([\d.,]+)\s+([A-Z]+)\s+     # quantidade + unidade
        ([\d.,]+)\s+([\d.,]+)\s+       # vl_unitario + vl_total
        ([\d.,]+)\s+                   # IPI
        (\d{2}/\d{2}/\d{4})            # data entrega
        ''',
    texto,
    re.DOTALL | re.VERBOSE
)

itens = []
for item in padrao_item:
    itens.append({
        "codigo_produto": item[0],
        "quantidade": item[1],
        "unidade": item[2],
        "vl_unitario": item[3],
        "vl_total": item[4],
        "ipi": item[5],
        "data_entrega": item[6]
    })

for i in itens:
    print(i)

# --- Totais da OC ---
total_ipi_match = re.search(r'Total IPI:\s*([\d.,]+)', texto)
total_ipi = total_ipi_match.group(1) if total_ipi_match else None

frete_match = re.search(r'Valor\s+frete:\s*([\d.,]+)', texto)
frete = frete_match.group(1) if frete_match else None

outras_despesas_match = re.search(r'Outras despesas:\s*([\d.,]+)', texto)
outras_despesas = outras_despesas_match.group(1) if outras_despesas_match else None

total_mercadorias_match = re.search(r'Total\s+Mercadorias:\s*([\d.,]+)', texto)
total_mercadorias = total_mercadorias_match.group(1) if total_mercadorias_match else None

descontos_match = re.search(r'Descontos:\s*([\d.,]+)', texto)
descontos = descontos_match.group(1) if descontos_match else None

acresc_financeiro_match = re.search(r'Acrésc\. Financeiro:\s*([\d.,]+)', texto)
acresc_financeiro = acresc_financeiro_match.group(1) if acresc_financeiro_match else None

cond_pgto_match = re.search(r'Cond\.\s+pgto\.:\s*(\d+)', texto)
cond_pgto = cond_pgto_match.group(1) if cond_pgto_match else None

total_geral_match = re.search(r'TOTAL GERAL:\s*([\d.,]+)', texto)
total_geral = total_geral_match.group(1) if total_geral_match else None

print("Frete: ", frete, "Outras Despesas: ", outras_despesas, "Descontos: ", descontos)
print("Total: ", total_geral)