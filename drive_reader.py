import io
import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from dotenv import load_dotenv

# Cargar variables de entorno desde .env
load_dotenv()

SCOPES = ["https://www.googleapis.com/auth/drive"]

# Leer credenciales desde el JSON almacenado en la variable de entorno
creds_json = os.environ.get("GOOGLE_DRIVE_JSON")
if not creds_json:
    raise ValueError("No se encontró la variable de entorno GOOGLE_DRIVE_JSON")

creds_dict = json.loads(creds_json)
creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)

# Servicio de Drive
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
