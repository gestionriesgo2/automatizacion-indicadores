# drive_reader.py
import io
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from dotenv import load_dotenv

# ------------------------------------------
# CARGAR VARIABLES DE ENTORNO DESDE .ENV
# ------------------------------------------
load_dotenv()  # Busca el archivo .env automáticamente

SCOPES = ["https://www.googleapis.com/auth/drive"]

# Tomar la ruta del JSON desde la variable de entorno
CRED_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
if not CRED_FILE:
    raise ValueError("No se encontró la variable de entorno GOOGLE_APPLICATION_CREDENTIALS")

# ------------------------------------------
# CARGA DE CREDENCIALES
# ------------------------------------------
creds = Credentials.from_service_account_file(
    CRED_FILE,
    scopes=SCOPES
)

# ------------------------------------------
# SERVICIO DE DRIVE
# ------------------------------------------
drive_service = build("drive", "v3", credentials=creds)

# ------------------------------------------
# LISTAR ARCHIVOS EN UNA CARPETA DE DRIVE
# ------------------------------------------
def list_files_in_folder(folder_id):
    archivos = []
    page_token = None

    while True:
        response = drive_service.files().list(
            q=f"'{folder_id}' in parents",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token
        ).execute()

        archivos.extend(response.get("files", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    return archivos

# ------------------------------------------
# LEER EXCEL DESDE DRIVE
# ------------------------------------------
def read_excel_from_drive(file_id):
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    fh.seek(0)
    return fh

# ------------------------------------------
# SUBIR (SOBRESCRIBIR) UN ARCHIVO EN DRIVE
# ------------------------------------------
def upload_bytes_to_drive(bytes_data, file_id):
    fh = io.BytesIO(bytes_data)

    media = MediaIoBaseUpload(
        fh,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False
    )

    updated_file = drive_service.files().update(
        fileId=file_id,
        media_body=media
    ).execute()

    return updated_file
