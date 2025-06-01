import os
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from crear_documentos import resource_path

# Configuración
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

CEDULA_COLUMN = 'E'  # Columna donde están las cédulas (A, B, C,...)
CEDULA_FIELD_NAME = 'Numero de Cedula'  # Nombre del campo en la hoja de cálculo





def get_authenticated_service(api_name, api_version):
    """Obtiene el servicio autenticado para diferentes APIs de Google"""
    creds = None
    
    if os.path.exists(resource_path('service_account.json')):
        creds = service_account.Credentials.from_service_account_file(
            resource_path('service_account.json'), scopes=SCOPES)
    elif os.path.exists(resource_path('token.json')):
        creds = Credentials.from_authorized_user_file(resource_path('token.json'), SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            credentials_path = resource_path('credentials.json')
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        token_path=resource_path('token.json')
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    
    return build(api_name, api_version, credentials=creds)

def get_client_by_cedula(sheets_service, spreadsheet_id, sheet_name, cedula):
    """
    Busca directamente un cliente por cédula y devuelve sus datos como diccionario
    Retorna None si no encuentra el cliente
    """
    try:
        # Primero obtenemos las cabeceras
        headers_range = f"{sheet_name}!1:1"
        headers_result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=headers_range
        ).execute()
        
        headers = headers_result.get('values', [[]])[0]
        if not headers:
            return None
        
        # Buscar la fila que contiene la cédula
        range_to_search = f"{sheet_name}!{CEDULA_COLUMN}:{CEDULA_COLUMN}"
        cedulas_result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_to_search
        ).execute()
        
        cedulas = cedulas_result.get('values', [])
        
        for row_num, row in enumerate(cedulas):
            if row and row[0] == str(cedula):
                data_range = f"{sheet_name}!{row_num+1}:{row_num+1}"
                data_result = sheets_service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=data_range
                ).execute()
                
                row_data = data_result.get('values', [[]])[0]
                
                return {headers[i]: row_data[i] if i < len(row_data) else '' for i in range(len(headers))}
        return None
        
    except HttpError as error:
        print(f"Error al buscar cliente: {error}")
        return None

def process_client_data(drive_service, parent_folder_id, client_data):
    try:
        # Verificar que tenemos el campo de cédula
        if CEDULA_FIELD_NAME not in client_data:
            print(f"Error: No se encontró el campo '{CEDULA_FIELD_NAME}' en los datos del cliente")
            return
        
        cedula = client_data[CEDULA_FIELD_NAME]
        
        # Buscar carpeta con el número de cédula
        query = f"'{parent_folder_id}' in parents and name='{cedula}' and mimeType='application/vnd.google-apps.folder'"
        results = drive_service.files().list(
            q=query,
            fields="files(id, name)"
        ).execute()
        items = results.get('files', [])
        
        if items:
            folder_id = items[0]['id']
            print(f"La carpeta ya existe: {items[0]['name']} (ID: {folder_id})")
        else:
            # Crear nueva carpeta
            folder_metadata = {
                'name': cedula,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_folder_id]
            }
            folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
            folder_id = folder.get('id')
            # Se guarda en la variable inicializada como Id de carpeta personal
            print(f"Carpeta creada: {cedula} (ID: {folder_id})")
        return folder_id
            
            
        
        
    except HttpError as error:
        print(f"Error al procesar cliente: {error}")

def subir_archivos_a_drive(drive_service,carpeta_local, carpeta_drive_id):
    """Sube todos los archivos de una carpeta local a una carpeta de Google Drive, 
    eliminando primero los archivos existentes en la carpeta de Drive."""
    
    carpeta_local = resource_path(carpeta_local)
    # Primero, borrar todos los archivos existentes en la carpeta de Drive
    print("Eliminando archivos existentes en la carpeta de Drive...")
    query = f"'{carpeta_drive_id}' in parents and trashed=false"
    resultados = drive_service.files().list(
        q=query,
        fields="files(id, name)"
    ).execute()
    
    for archivo in resultados.get('files', []):
        drive_service.files().delete(fileId=archivo['id']).execute()
        print(f"Eliminado de Drive: {archivo['name']} (ID: {archivo['id']})")
    
    # Luego, subir los nuevos archivos
    print("\nSubiendo nuevos archivos...")
    for nombre_archivo in os.listdir(carpeta_local):
        ruta = os.path.join(carpeta_local, nombre_archivo)
        if os.path.isfile(ruta):
            file_metadata = {
                'name': nombre_archivo,
                'parents': [carpeta_drive_id]
            }
            media = MediaFileUpload(ruta, resumable=True)
            archivo = drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            print(f'Subido: {nombre_archivo} (ID: {archivo["id"]})')
    print('Archivos subidos!!!')


