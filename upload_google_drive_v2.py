import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from datetime import datetime, timezone
from core.credenciais import pasta_local, pasta_google


SCOPES = ["https://www.googleapis.com/auth/drive.file"]


class GoogleDriveSync:

    def __init__(self):
        self.service = self.login()

        self.arquivos_google = {}

        print("teste")

    def login(self):

        creds = None

        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file(
                "token.json",
                SCOPES
            )

        if not creds or not creds.valid:

            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())

            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json",
                    SCOPES
                )

                creds = flow.run_local_server(port=0)

            with open("token.json", "w") as token:
                token.write(creds.to_json())

        return build(
            "drive",
            "v3",
            credentials=creds
        )

    def procurar_arquivo(self, nome_arquivo):

        resultado = self.service.files().list(
            q=f"'{pasta_google}' in parents and name='{nome_arquivo}' and trashed=false",
            fields="files(id, name, modifiedTime)"
        ).execute()

        arquivos = resultado.get("files", [])

        if arquivos:
            return arquivos[0]

        return None

    def upload_arquivo(self, caminho_arquivo):

        metadata = {
            "name": os.path.basename(caminho_arquivo),
            "parents": [pasta_google]
        }

        media = MediaFileUpload(
            caminho_arquivo,
            resumable=True
        )

        arquivo = self.service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name"
        ).execute()

        return arquivo["id"]

    def atualizar_arquivo(self, id_google, caminho_arquivo):

        media = MediaFileUpload(
            caminho_arquivo,
            resumable=True
        )

        self.service.files().update(
            fileId=id_google,
            media_body=media
        ).execute()

    def sincronizar(self):

        self.carregar_arquivos_google()

        uploads = 0
        atualizacoes = 0

        limite_uploads = 100

        for nome_arquivo in os.listdir(pasta_local):

            caminho_arquivo = os.path.join(
                pasta_local,
                nome_arquivo
            )

            if not os.path.isfile(caminho_arquivo):
                continue

            # -------------------------------
            # Arquivo já existe no Google
            # -------------------------------
            if nome_arquivo in self.arquivos_google:

                arquivo_google = self.arquivos_google[nome_arquivo]

                data_local = datetime.fromtimestamp(
                    os.path.getmtime(caminho_arquivo),
                    timezone.utc
                )

                data_google = datetime.strptime(
                    arquivo_google["modifiedTime"],
                    "%Y-%m-%dT%H:%M:%S.%fZ"
                ).replace(tzinfo=timezone.utc)

                diferenca = (data_local - data_google).total_seconds()

                if diferenca > 5:

                    print(f"ATUALIZANDO -> {nome_arquivo}")

                    self.atualizar_arquivo(
                        arquivo_google["id"],
                        caminho_arquivo
                    )

                    atualizacoes += 1

                else:

                    # Se quiser ver os arquivos ignorados, descomente:
                    # print(f"OK -> {nome_arquivo}")
                    pass

            # -------------------------------
            # Arquivo não existe no Google
            # -------------------------------
            else:

                print(f"ENVIANDO -> {nome_arquivo}")

                self.upload_arquivo(caminho_arquivo)

                uploads += 1

                if uploads >= limite_uploads:
                    print(f"\nLimite de {limite_uploads} uploads atingido.")
                    break

        print()
        print(f"Uploads........: {uploads}")
        print(f"Atualizações...: {atualizacoes}")

    def carregar_arquivos_google(self):

        self.arquivos_google = {}

        pagina = None

        while True:

            resultado = self.service.files().list(
                q=f"'{pasta_google}' in parents and trashed=false",
                fields="nextPageToken, files(id,name,modifiedTime)",
                pageToken=pagina
            ).execute()

            for arquivo in resultado.get("files", []):
                self.arquivos_google[arquivo["name"]] = arquivo

            pagina = resultado.get("nextPageToken")

            if not pagina:
                break

        print(f"Arquivos encontrados no Google Drive: {len(self.arquivos_google)}")



if __name__ == "__main__":

    drive = GoogleDriveSync()

    drive.sincronizar()