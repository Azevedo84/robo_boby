import win32com.client

#arquivo = r"\\Publico\C\Inventor\21 - Ponto Bobina\21.01.03.01.ipt"
#arquivo = r"\\Publico\C\Inventor\21 - Ponto Bobina\21.01.03.00.iam"
#arquivo = r"\\Publico\C\Inventor\1 - Folha A4\Cód. 21 - Ponto de Bobina\21.01.03.00 - Conj. Ponto Bob Rol\01 - 21.01.03.01.idw"
arquivo = r"\\Publico\C\Inventor\Biblioteca\Molas\MC13X10X29\MC13X10X29.ipt"

inventor = None
doc = None

try:

    inventor = win32com.client.Dispatch("Inventor.Application")
    inventor.Visible = False

    doc = inventor.Documents.Open(arquivo)

    print("\nTODAS AS PROPRIEDADES")
    print("====================================")

    for prop_set in doc.PropertySets:

        print(f"\nPROPERTY SET: {prop_set.Name}")
        print("--------------------------------")

        for prop in prop_set:
            try:
                print(f"{prop.Name}: {prop.Value}")
            except:
                print(f"{prop.Name}: (sem valor)")

    ext = arquivo.lower()

    if ext.endswith(".iam"):

        comp = doc.ComponentDefinition
        bom = comp.BOM

        bom.StructuredViewEnabled = True
        bom.StructuredViewFirstLevelOnly = False

        view_estruturada = None

        for view in bom.BOMViews:
            if view.ViewType == 62465:
                view_estruturada = view

        quantidade_itens = 0

        for row in view_estruturada.BOMRows:

            if row.ComponentDefinitions.Count == 0:
                continue

            quantidade_itens += 1

        print("\nTOTAL ITENS NA ESTRUTURA:", quantidade_itens)

    if arquivo.lower().endswith(".idw"):

        print("\nARQUIVOS REFERENCIADOS:")
        print("--------------------------------")

        for ref in doc.ReferencedDocuments:
            print(ref.FullFileName)

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