import win32com.client

arquivo_conj = r"\\Publico\c\Inventor\21 - Ponto Bobina\21.01.03.00.iam"
#arquivo_conj = r"\\Publico\c\Inventor\21 - Ponto Bobina\21.01.03.00.iam"
arquivo_peca = r"\\Publico\c\Inventor\21 - Ponto Bobina\21.01.03.01.ipt"

inventor = win32com.client.Dispatch("Inventor.Application")

# deixa invisível
inventor.Visible = False
inventor.SilentOperation = True

doc = None

try:

    doc = inventor.Documents.Open(arquivo_peca, False)

    props = doc.PropertySets.Item("Design Tracking Properties")

    print("Authority atual:", props.Item("Authority").Value)

    props.Item("Authority").Value = "163349"

    doc.Save()

except Exception as e:
    print("Erro:", e)

finally:
    if doc:
        doc.Close()

    inventor.Quit()