import win32com.client

arquivo = r"\\Publico\c\Inventor\21 - Ponto Bobina\21.01.03.00.iam"

inventor = win32com.client.Dispatch("Inventor.Application")
inventor.Visible = False

doc = inventor.Documents.Open(arquivo)

comp = doc.ComponentDefinition
bom = comp.BOM

# habilita BOM estruturada
bom.StructuredViewEnabled = True
bom.StructuredViewFirstLevelOnly = False

view_estruturada = None

for view in bom.BOMViews:
    if view.ViewType == 62465:  # Structured
        view_estruturada = view


estrutura = []

for row in view_estruturada.BOMRows:

    comp_def = row.ComponentDefinitions.Item(1)
    doc_item = comp_def.Document

    qtde = row.ItemQuantity

    props = doc_item.PropertySets
    design = props.Item("Design Tracking Properties")

    try:
        codigo = design.Item("Authority").Value
    except:
        codigo = ""

    try:
        referencia = design.Item("Part Number").Value
    except:
        referencia = ""

    try:
        descricao = design.Item("Description").Value
    except:
        descricao = ""

    estrutura.append({
        "codigo": codigo,
        "referencia": referencia,
        "descricao": descricao,
        "quantidade": qtde
    })


doc.Close()

print("\nESTRUTURA DO CONJUNTO\n")

for item in estrutura:
    print(item)