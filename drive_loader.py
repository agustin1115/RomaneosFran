"""
drive_loader.py — Lee PDFs de romaneo desde Google Drive.
Organiza por mes y semana según la fecha del romaneo.
"""
import os
import io
import tempfile
from datetime import datetime, date
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

DRIVE_FOLDER_ID = '1PpSEKdjQGk3PmU4TAZc6JqngSivvcLwz'
SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.metadata.readonly',
]


def get_drive_service(credentials_path):
    """Crea el servicio de Google Drive."""
    creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)


def listar_pdfs_drive(credentials_path, folder_id=None):
    """
    Lista todos los PDFs en la carpeta de Drive.
    Retorna lista de dicts con id, name, createdTime, size.
    """
    if not folder_id:
        folder_id = DRIVE_FOLDER_ID

    service = get_drive_service(credentials_path)
    all_files = []
    page_token = None

    while True:
        results = service.files().list(
            q=f"'{folder_id}' in parents and mimeType='application/pdf'",
            fields='nextPageToken, files(id, name, createdTime, modifiedTime, size)',
            pageSize=100,
            pageToken=page_token,
            orderBy='createdTime desc'
        ).execute()

        all_files.extend(results.get('files', []))
        page_token = results.get('nextPageToken')
        if not page_token:
            break

    return all_files


def descargar_pdf(credentials_path, file_id):
    """
    Descarga un PDF de Drive a un archivo temporal.
    Retorna el path del archivo temporal.
    """
    service = get_drive_service(credentials_path)
    request = service.files().get_media(fileId=file_id)

    tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    content = request.execute()
    tmp.write(content)
    tmp.close()
    return tmp.name


def obtener_semana_mes(fecha_dt):
    """Devuelve el número de semana dentro del mes (1-5)."""
    if not fecha_dt:
        return 1
    dia = fecha_dt.day
    if dia <= 7:
        return 1
    elif dia <= 14:
        return 2
    elif dia <= 21:
        return 3
    elif dia <= 28:
        return 4
    else:
        return 5


def organizar_por_mes_semana(archivos_con_fecha):
    """
    Recibe lista de dicts con 'nombre', 'fecha' (DD/MM/YYYY), y otros datos.
    Retorna dict anidado: {mes: {semana: [archivos]}}.
    Formato mes: "2026-03 Marzo", semana: "Semana 1 (1-7)"
    """
    meses_nombre = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }

    organizado = {}

    for arch in archivos_con_fecha:
        fecha_str = arch.get('fecha', '')
        fecha_dt = None
        if fecha_str:
            try:
                parts = fecha_str.split('/')
                if len(parts) >= 2:
                    dd = int(parts[0])
                    mm = int(parts[1])
                    yyyy = int(parts[2]) if len(parts) >= 3 else 2026
                    fecha_dt = date(yyyy, mm, dd)
            except (ValueError, IndexError):
                pass

        if not fecha_dt:
            # Intentar desde createdTime de Drive
            ct = arch.get('createdTime', '')
            if ct:
                try:
                    fecha_dt = datetime.fromisoformat(ct.replace('Z', '+00:00')).date()
                except (ValueError, TypeError):
                    pass

        if not fecha_dt:
            fecha_dt = date.today()

        mes_key = f"{fecha_dt.year}-{fecha_dt.month:02d} {meses_nombre.get(fecha_dt.month, '?')}"
        semana = obtener_semana_mes(fecha_dt)
        rangos = {1: '1-7', 2: '8-14', 3: '15-21', 4: '22-28', 5: '29-31'}
        semana_key = f"Semana {semana} ({rangos[semana]})"

        if mes_key not in organizado:
            organizado[mes_key] = {}
        if semana_key not in organizado[mes_key]:
            organizado[mes_key][semana_key] = []

        arch['fecha_dt'] = fecha_dt
        arch['semana'] = semana
        organizado[mes_key][semana_key].append(arch)

    # Ordenar
    organizado = dict(sorted(organizado.items(), reverse=True))
    for mes in organizado:
        organizado[mes] = dict(sorted(organizado[mes].items()))

    return organizado
