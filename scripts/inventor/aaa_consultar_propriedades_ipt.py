import win32com.client

#arquivo = r"\\publico\c\inventor\24 - corte e solda\24.00.001 - corte e solda\24.00.019.01\24.01.053.05.ipt"
arquivo = r"\\publico\c\inventor\1 - folha a4\cód. 24 - corte e solda\24.00.001\24.00.019.01\24.01.053.05.idw"

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

    # 🔥 AQUI É O QUE TU QUER
    if arquivo.lower().endswith(".idw"):

        print("\nARQUIVOS REFERENCIADOS:")
        print("--------------------------------")

        for ref in doc.ReferencedDocuments:
            print(ref.FullFileName)

        # 🔥 NOVO: pegar documento referenciado (.ipt)
        ref_doc = None
        try:
            ref_doc = doc.ReferencedDocuments[0]
        except:
            pass

        # 🔥 NOVO: listar parâmetros do modelo
        if ref_doc:
            print("\nPARAMETROS DO MODELO (.ipt):")
            print("--------------------------------")

            try:
                params = ref_doc.ComponentDefinition.Parameters

                for p in params:
                    try:
                        valor_mm = round(p.Value * 10, 2)
                        print(f"{p.Name}: {valor_mm} mm")
                    except:
                        print(f"{p.Name}: (sem valor)")
            except:
                print("Não conseguiu acessar parâmetros")

        # 🔹 COTAS DO DESENHO
        print("\nCOTAS DO DESENHO:")
        print("--------------------------------")

        for sheet in doc.Sheets:
            print(f"\nSHEET: {sheet.Name}")

            for dim in sheet.DrawingDimensions:

                try:
                    texto = dim.Text.Text
                except:
                    texto = "(sem texto)"

                try:
                    valor = dim.ModelValue  # cm
                    valor_mm = round((valor * 10), 2)
                except:
                    valor_mm = "(sem valor)"

                # 🔥 NOVO: pegar parâmetro da cota
                try:
                    param = dim.Parameter
                    nome_param = param.Name
                except:
                    nome_param = "SEM_PARAM"

                print(f"Param: {nome_param} | Texto: {texto} | Valor (mm): {valor_mm}")

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