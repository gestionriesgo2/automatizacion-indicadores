import io
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ------------------------------------------
# CONFIGURACIÓN GOOGLE DRIVE
# ------------------------------------------
SCOPES = ["https://www.googleapis.com/auth/drive"]

# Ruta al archivo de credenciales
CREDENTIALS_FILE = "python-drive-service-a7c2f08eb564.json"

# Crear credenciales desde archivo JSON
creds = Credentials.from_service_account_file(
    CREDENTIALS_FILE,
    scopes=SCOPES
)

# Servicio de Google Drive
drive_service = build("drive", "v3", credentials=creds)

# ------------------------------------------------
# OBTENER ID DEL BANCO DESDE UNA CARPETA
# ------------------------------------------------
def get_banco_file_id_from_folder(folder_id):
    query = (
        f"'{folder_id}' in parents and "
        f"(mimeType='application/vnd.google-apps.spreadsheet' or "
        f"mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') and "
        f"trashed = false"
    )

    response = drive_service.files().list(
        q=query,
        fields="files(id, name, mimeType)",
        pageSize=1
    ).execute()

    files = response.get("files", [])

    if not files:
        raise ValueError("❌ No se encontró ningún archivo banco en la carpeta")

    print(f"📄 Banco detectado: {files[0]['name']}")
    return files[0]["id"]

# ------------------------------------------
# LISTAR ARCHIVOS EN UNA CARPETA DE DRIVE
# ------------------------------------------
def list_files_in_folder(folder_id):
    archivos = []
    page_token = None

    while True:
        response = drive_service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
            supportsAllDrives=True,          # 🔥 CLAVE
            includeItemsFromAllDrives=True   # 🔥 CLAVE
        ).execute()

        archivos.extend(response.get("files", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    return archivos


# ------------------------------------------
# LEER ARCHIVO EXCEL DESDE DRIVE
# ------------------------------------------
def read_excel_from_drive(file_id):

    file = drive_service.files().get(
        fileId=file_id,
        fields="id, name, mimeType, shortcutDetails",
        supportsAllDrives=True
    ).execute()

    mime_type = file["mimeType"]

    # Resolver shortcut
    if mime_type == "application/vnd.google-apps.shortcut":
        file_id = file["shortcutDetails"]["targetId"]
        file = drive_service.files().get(
            fileId=file_id,
            fields="id, name, mimeType",
            supportsAllDrives=True
        ).execute()
        mime_type = file["mimeType"]

    fh = io.BytesIO()

    if mime_type == "application/vnd.google-apps.spreadsheet":
        request = drive_service.files().export(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            supportsAllDrives=True
        )
    elif mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        request = drive_service.files().get_media(
            fileId=file_id,
            supportsAllDrives=True
        )
    else:
        raise ValueError(f"❌ Tipo de archivo no soportado: {mime_type}")

    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    fh.seek(0)
    return fh




# ------------------------------------------
# SUBIR / SOBRESCRIBIR ARCHIVO EN DRIVE
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


# ------------------------------------------
# VERIFICAR SI UN ARCHIVO EXISTE EN DRIVE
# ------------------------------------------
def file_exists(file_id):
    try:
        drive_service.files().get(
            fileId=file_id,
            supportsAllDrives=True
        ).execute()
        return True
    except:
        return False



# ------------------------------------------
# CREAR O ACTUALIZAR ARCHIVO EN DRIVE
# ------------------------------------------
def create_or_update_file(
    bytes_data,
    file_id=None,
    filename="archivo.xlsx",
    parent_folder_id=None,
    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
):

    media = MediaIoBaseUpload(
        io.BytesIO(bytes_data),
        mimetype=mimetype,
        resumable=False
    )

    # 🔎 VALIDAR SI EL ID EXISTE
    if file_id and file_exists(file_id):
        print(f"♻ Actualizando archivo: {filename}")
        return drive_service.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True
        ).execute()

    # 🔄 Si el ID no existe, crear nuevo archivo
    print(f"🆕 Creando archivo (ID inválido o no existe): {filename}")

    metadata = {"name": filename}
    if parent_folder_id:
        metadata["parents"] = [parent_folder_id]

    return drive_service.files().create(
        body=metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True
    ).execute()



def get_file_id_by_name(folder_id, filename):
    query = (
        f"'{folder_id}' in parents and "
        f"name = '{filename}' and trashed = false"
    )

    response = drive_service.files().list(
        q=query,
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()

    files = response.get("files", [])
    return files[0]["id"] if files else None

