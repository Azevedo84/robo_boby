import sys
import io
import fitz

from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


SCOPES = ["https://www.googleapis.com/auth/drive"]


class Janela(QWidget):

    def __init__(self):
        super().__init__()

        self.resize(900, 700)

        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)

        layout = QVBoxLayout(self)
        layout.addWidget(self.label)

        id_pdf = "11AotqP30qRAe8jJ54iTRZrOv71fulqxI"

        self.mostrar_pdf(id_pdf)

    def mostrar_pdf(self, id_pdf):

        creds = Credentials.from_authorized_user_file(
            "token.json",
            SCOPES
        )

        service = build(
            "drive",
            "v3",
            credentials=creds
        )

        arquivos = {}

        while True:

            resultado = service.files().list(
                q="name='01 - 9.01.09.pdf' and trashed=false",
                fields="files(id,name)"
            ).execute()

            print(resultado)

            for arquivo in resultado.get("files", []):
                arquivos[arquivo["name"]] = arquivo

            pagina = resultado.get("nextPageToken")

            if not pagina:
                break

        print(f"Total: {len(arquivos)}")

        nome = "01 - 9.01.09.pdf"

        if nome not in arquivos:
            print("NÃO ENCONTROU")
            return

        id_pdf = arquivos[nome]["id"]

        request = service.files().get_media(fileId=id_pdf)

        buffer = io.BytesIO()

        downloader = MediaIoBaseDownload(buffer, request)

        concluido = False

        while not concluido:
            status, concluido = downloader.next_chunk()

        pdf_bytes = buffer.getvalue()

        doc = fitz.open(
            stream=pdf_bytes,
            filetype="pdf"
        )

        page = doc.load_page(0)

        pix = page.get_pixmap(
            matrix=fitz.Matrix(1.4, 1.2)
        )

        image = QImage(
            pix.samples,
            pix.width,
            pix.height,
            pix.stride,
            QImage.Format_RGB888
        )

        pixmap = QPixmap.fromImage(image)

        self.label.setPixmap(pixmap)
        self.label.setScaledContents(True)

app = QApplication(sys.argv)

janela = Janela()
janela.show()

sys.exit(app.exec_())