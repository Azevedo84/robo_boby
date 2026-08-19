import os
from core.banco_nuvem import conectar_banco_nuvem
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.exceptions import RefreshError
from datetime import datetime, timezone
from core.credenciais import pasta_local, pasta_google
from typing import cast


SCOPES = ["https://www.googleapis.com/auth/drive"]


class GoogleDriveSync:
    def __init__(self):
        self.service = self.login()

        about = self.service.about().get(
            fields="user"
        ).execute()

        print(about)

        self.arquivos_google = {}
        self.arquivos_banco = {}

    def carregar_arquivos_banco(self):
        self.arquivos_banco = {}

        conexao = conectar_banco_nuvem()
        cursor = conexao.cursor(dictionary=True)

        cursor.execute("""
            SELECT
                ID,
                TIPO,
                NOME,
                EXTENSAO,
                GOOGLE_FILE_ID
            FROM ARQUIVOS
        """)

        for registro in cursor.fetchall():
            chave = f"{registro['NOME']}.{registro['EXTENSAO']}"

            self.arquivos_banco[chave] = registro

        print(f"BANCO: {len(self.arquivos_banco)} arquivos")

        cursor.close()
        conexao.close()

    def inserir_arquivo_banco(self, arquivo, cursor):
        nome = arquivo["name"]
        google_id = arquivo["id"]

        if "." in nome:
            nome_sem_ext, extensao = nome.rsplit(".", 1)
        else:
            nome_sem_ext = nome
            extensao = ""

        cursor.execute("""
            INSERT INTO ARQUIVOS
            (
                TIPO,
                NOME,
                EXTENSAO,
                GOOGLE_FILE_ID
            )
            VALUES (%s,%s,%s,%s)
        """, (
            "PRODUTO",
            nome_sem_ext,
            extensao.lower(),
            google_id
        ))

        print(f"INSERT ARQUIVO {nome}!")

    def excluir_arquivo_banco(self, registro_banco, cursor):

        cursor.execute("""
            DELETE FROM ARQUIVOS
            WHERE ID = %s
        """, (
            registro_banco["ID"],
        ))

        print(f"DELETE ARQUIVO {registro_banco['NOME']}.{registro_banco['EXTENSAO']}!")

    def login(self):

        creds = None

        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file(
                "token.json",
                SCOPES
            )

        if not creds or not creds.valid:

            if creds and creds.expired and creds.refresh_token:

                try:
                    creds.refresh(Request())

                except RefreshError as erro:
                    if "invalid_grant" not in str(erro):
                        raise

                    if os.path.exists("token.json"):
                        os.remove("token.json")

                    flow = cast(
                        InstalledAppFlow,
                        InstalledAppFlow.from_client_secrets_file(
                            "credentials.json",
                            SCOPES
                        )
                    )

                    creds = flow.run_local_server(port=0)

            else:
                flow = cast(
                    InstalledAppFlow,
                    InstalledAppFlow.from_client_secrets_file(
                        "credentials.json",
                        SCOPES
                    )
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
            fields="nextPageToken, files(id,name,modifiedTime,mimeType)"
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

    def sincronizar_drive(self):
        self.carregar_arquivos_google()

        uploads = 0
        atualizacoes = 0

        limite_uploads = 999

        for nome_arquivo in os.listdir(pasta_local):

            caminho_arquivo = os.path.join(
                pasta_local,
                nome_arquivo
            )

            if not os.path.isfile(caminho_arquivo):
                continue

            # -----------------------------------
            # Arquivo já existe no Google
            # -----------------------------------
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

            # -----------------------------------
            # Arquivo não existe no Google
            # -----------------------------------
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

    def sincronizar_banco(self):
        self.carregar_arquivos_banco()

        conexao = conectar_banco_nuvem()
        cursor = conexao.cursor()

        inseridos = 0

        for nome, arquivo in self.arquivos_google.items():
            if nome not in self.arquivos_banco:
                self.inserir_arquivo_banco(arquivo, cursor)
                inseridos += 1

        excluidos = 0

        for nome, registro in self.arquivos_banco.items():
            if nome not in self.arquivos_google:
                self.excluir_arquivo_banco(registro, cursor)
                excluidos += 1

        conexao.commit()
        cursor.close()
        conexao.close()

        print(f"Inseridos...: {inseridos}")
        print(f"Excluídos..: {excluidos}")

    def carregar_arquivos_google(self):

        self.arquivos_google = {}

        pagina = None
        pagina_num = 1

        while True:
            resultado = self.service.files().list(
                q=f"'{pasta_google}' in parents and trashed=false",
                fields="nextPageToken,files(id,name,modifiedTime,md5Checksum,owners(emailAddress),shortcutDetails)",
                pageToken=pagina,
                pageSize=1000,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()

            arquivos = resultado.get("files", [])

            print(f"Página {pagina_num}: {len(arquivos)} arquivos")

            for arquivo in arquivos:
                nome = arquivo["name"]
                email = arquivo["owners"][0]["emailAddress"]

                if nome not in self.arquivos_google:
                    self.arquivos_google[nome] = []

                self.arquivos_google[nome].append(arquivo)

            pagina = resultado.get("nextPageToken")

            if not pagina:
                break

            pagina_num += 1

        # ------------------------------------------------------------------
        # Remove duplicados (mantém somente o mais recente)
        # ------------------------------------------------------------------

        total_duplicados = 0

        for nome, lista in list(self.arquivos_google.items()):

            if len(lista) == 1:
                self.arquivos_google[nome] = lista[0]
                continue

            print(f"\nDUPLICADO: {nome}")

            for arq in lista:
                print(arq["owners"][0]["emailAddress"], arq["id"], arq["name"])

            lista.sort(
                key=lambda x: x["modifiedTime"],
                reverse=True
            )

            manter = lista[0]

            for excluir in lista[1:]:
                print(f"Excluindo: {excluir['id']}")

                self.service.files().delete(
                    fileId=excluir["id"]
                ).execute()

                total_duplicados += 1

            self.arquivos_google[nome] = manter

        print(f"Duplicados removidos: {total_duplicados}")
        print(f"TOTAL: {len(self.arquivos_google)}")



if __name__ == "__main__":

    drive = GoogleDriveSync()

    drive.sincronizar_drive()
    drive.sincronizar_banco()