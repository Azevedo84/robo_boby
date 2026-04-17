import win32com.client
import os
import time

#idw = r"\\Publico\c\Inventor\1 - Folha A4\Cód. 21 - Ponto de Bobina\21.01.03.00 - Conj. Ponto Bob Rol\01 - 21.01.03.01.idw"
idw = r"\\Publico\c\Inventor\1 - Folha A4\Cód. 21 - Ponto de Bobina\21.01.03.00 - Conj. Ponto Bob Rol\01 - 21.01.03.01.idw"

desktop = os.path.join(os.path.expanduser("~"), "Desktop")
pdf = os.path.join(desktop, "teste_inventor.pdf")

inventor = None
doc = None

inicio_total = time.time()

try:

    t = time.time()
    print("Abrindo Inventor...")
    inventor = win32com.client.Dispatch("Inventor.Application")
    inventor.Visible = False
    print("Tempo abrir Inventor:", round(time.time() - t, 2), "seg")

    t = time.time()
    print("Abrindo desenho...")
    doc = inventor.Documents.Open(idw, False)
    print("Tempo abrir desenho:", round(time.time() - t, 2), "seg")

    t = time.time()
    print("Carregando addin PDF...")

    pdf_addin = inventor.ApplicationAddIns.ItemById(
        "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"
    )

    if not pdf_addin.Activated:
        pdf_addin.Activate()

    context = inventor.TransientObjects.CreateTranslationContext()
    context.Type = 13059

    options = inventor.TransientObjects.CreateNameValueMap()

    data = inventor.TransientObjects.CreateDataMedium()
    data.FileName = pdf

    print("Gerando PDF...")
    pdf_addin.SaveCopyAs(doc, context, options, data)

    print("Tempo gerar PDF:", round(time.time() - t, 2), "seg")

    print("PDF criado:", pdf)

except Exception as e:
    print("ERRO:", e)

finally:

    try:
        if doc:
            doc.Close()
    except:
        pass

    try:
        if inventor:
            inventor.Quit()
    except:
        pass

    print("Tempo total:", round(time.time() - inicio_total, 2), "seg")