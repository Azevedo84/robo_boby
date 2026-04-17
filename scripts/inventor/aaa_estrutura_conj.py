import win32com.client

#arquivo = r"\\PUBLICO\Inventor\18 - Extrusora Continua\18.00.001.iam"
arquivo = r"\\PUBLICO\Inventor\18 - Extrusora Continua\18.00.001.iam"

inventor = None
doc = None

try:

    inventor = win32com.client.Dispatch("Inventor.Application")
    inventor.Visible = False

    doc = inventor.Documents.Open(arquivo)

    # -------------------------------
    # DADOS DO PROPRIO CONJUNTO
    # -------------------------------

    design = doc.PropertySets.Item("Design Tracking Properties")

    try:
        codigo_conjunto = design.Item("Authority").Value
    except:
        codigo_conjunto = ""

    try:
        referencia_conjunto = design.Item("Part Number").Value
    except:
        referencia_conjunto = ""

    try:
        descricao_conjunto = design.Item("Description").Value
    except:
        descricao_conjunto = ""

    print("\nDADOS DO CONJUNTO")
    print("-----------------------------")
    print("codigo:", codigo_conjunto)
    print("referencia:", referencia_conjunto)
    print("descricao:", descricao_conjunto)

    # -------------------------------
    # ESTRUTURA
    # -------------------------------

    comp = doc.ComponentDefinition
    bom = comp.BOM

    bom.StructuredViewEnabled = True
    bom.StructuredViewFirstLevelOnly = False

    view_estruturada = None

    for view in bom.BOMViews:
        if view.ViewType == 62465:
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

    print("\nESTRUTURA DO CONJUNTO\n")

    for item in estrutura:
        print(item)

except Exception as e:
    print("ERRO:", e)

finally:

    if doc:
        try:
            doc.Close(True)
        except:
            pass

    if inventor:
        try:
            inventor.Quit()
        except:
            pass