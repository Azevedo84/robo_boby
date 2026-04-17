import win32com.client

#arquivo = r"\\PUBLICO\c\Inventor\71 - Seladora Grande\71.00.001.00 - CONJ SELADORA CAIXA\71.00.002.01 - CONJ BR SELA 1550\71.00.013.01\71.01.053.03.ipt"
arquivo = r"\\PUBLICO\c\Inventor\71 - Seladora Grande\71.00.001.00 - CONJ SELADORA CAIXA\71.00.002.01 - CONJ BR SELA 1550\71.00.013.01\71.01.053.03.ipt"

inventor = None
doc = None

try:

    inventor = win32com.client.Dispatch("Inventor.Application")
    inventor.Visible = False

    doc = inventor.Documents.Open(arquivo)

    design = doc.PropertySets.Item("Design Tracking Properties")
    summary = doc.PropertySets.Item("Inventor Summary Information")

    referencia = design.Item("Part Number").Value
    codigo_produto = design.Item("Authority").Value
    descricao = design.Item("Description").Value

    codigo_materia = design.Item("Cost Center").Value
    medida_corte = design.Item("Vendor").Value
    material = design.Item("Material").Value

    descricao_materia = summary.Item("Revision Number").Value

    print("\nDADOS DO IPT")
    print("---------------------------")

    print("referencia:", referencia)
    print("codigo produto:", codigo_produto)
    print("descricao:", descricao)

    print("\nMATERIA PRIMA")
    print("---------------------------")

    print("codigo:", codigo_materia)
    print("material:", material)
    print("descricao:", descricao_materia)
    print("corte:", medida_corte)

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