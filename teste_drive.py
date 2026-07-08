import sys
import io
import fitz

from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


SCOPES = ["https://www.googleapis.com/auth/drive.file"]

ID_PDF = "1y9Nab-j3GWFIBzEOLieo_nwKSmR3HsB5"


class Janela(QWidget):

    def __init__(self):
        super().__init__()

        self.resize(900, 700)

        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)

        layout = QVBoxLayout(self)
        layout.addWidget(self.label)

        self.mostrar_pdf()


    def mostrar_pdf(self):

        creds = Credentials.from_authorized_user_file(
            "token.json",
            SCOPES
        )

        service = build(
            "drive",
            "v3",
            credentials=creds
        )

        request = service.files().get_media(
            fileId=ID_PDF
        )

        buffer = io.BytesIO()

        downloader = MediaIoBaseDownload(
            buffer,
            request
        )

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