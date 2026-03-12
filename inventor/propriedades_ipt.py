import win32com.client

arquivo = r"\\Publico\c\Inventor\21 - Ponto Bobina\21.01.02.ipt"

inventor = win32com.client.Dispatch("Inventor.Application")
inventor.Visible = True

doc = inventor.Documents.Open(arquivo)

props = doc.PropertySets.Item("Design Tracking Properties")

print("PROPRIEDADES DESIGN TRACKING")
print("-----------------------------")

for prop in props:
    try:
        print(prop.Name, "=", prop.Value)
    except:
        pass