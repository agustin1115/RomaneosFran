"""
frigorifico_real.py — Lee la cuenta corriente del frigorífico (Google Drive,
xlsx) y extrae el REAL por servicio (VEP, Cuarteo, Despostada, Congelado),
por semana y por mes. Se conecta en vivo: cada vez que se cargan romaneos,
la app puede tomar el real desde acá y cruzarlo contra el proyectado.

IMPORTANTE: para que la app lo lea sola, el archivo debe estar compartido con
la cuenta de servicio (el mismo email que ya lee la carpeta de romaneos).
"""
import io
import re
import unicodedata

import openpyxl

# Archivo de cuenta corriente del frigorífico en Drive (actualizable)
FILE_ID_DEFAULT = '16ZWnFSVidlVVtSN0XMPLBidu7zuec2C_'

MESES = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
         7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}


def _num(v):
    """Convierte '$ 2,100,350.00' o 8254.64 a float."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v)
    s = s.replace('$', '').replace(' ', '').replace(' ', '').strip()
    if not s or s in {'-', '—'}:
        return 0.0
    neg = s.startswith('-')
    s = s.replace('-', '')
    # formato AR/US con miles ',' y decimal '.'
    s = s.replace(',', '')
    try:
        n = float(s)
        return -n if neg else n
    except ValueError:
        return 0.0


def _txt(v):
    if v is None:
        return ''
    t = str(v)
    t = ''.join(c for c in unicodedata.normalize('NFKD', t) if not unicodedata.combining(c))
    return t.upper().strip()


def parse_cuenta_corriente(rows):
    """rows = lista de tuplas (celdas por fila). Columnas: 0=label 1=cantidad
    2=precio 3=monto. Devuelve lista de semanas con sus totales por servicio."""
    semanas = []
    actual = None

    for row in rows:
        if not row:
            continue
        c0 = _txt(row[0] if len(row) > 0 else '')
        c1 = row[1] if len(row) > 1 else None
        monto = _num(row[3]) if len(row) > 3 else 0.0

        # Fin de los bloques semanales: empieza el resumen contable
        if c0.startswith('RESUMEN VEP') or c0.startswith('TOTAL VEP A PAGAR'):
            break

        # Encabezado de semana: "SEMANA X AL Y-MM-26 (215)"
        if c0.startswith('SEMANA') and '(' in c0:
            # cab y kg pueden estar en c0 o en la celda siguiente
            blob = ' '.join(_txt(x) for x in row if x)
            mcab = re.search(r'CAB\s+([\d.]+)', blob)
            mkg = re.search(r'KG\s+([\d.]+)', blob)
            mmes = re.search(r'AL\s+\d+-(\d{2})-', c0)
            actual = {
                'semana': str(row[0]).strip(),
                'mes': int(mmes.group(1)) if mmes else 0,
                'cab': int(_num(mcab.group(1))) if mcab else 0,
                'kg_faena': _num(mkg.group(1)) if mkg else 0,
                'vep': 0.0, 'cuarteo': 0.0, 'despostada': 0.0, 'congelado': 0.0,
                # cantidades físicas (para comparar volúmenes, no solo plata)
                'vep_cab': 0.0, 'cuarteo_und': 0.0, 'despostada_kg': 0.0,
                'congelado_tunel_kg': 0.0, 'congelado_stock_kg': 0.0,
                'congelado_tunel': 0.0, 'congelado_stock': 0.0,
            }
            semanas.append(actual)
            continue

        if actual is None:
            continue

        # Líneas de servicio (deben tener cantidad en col1 → distingue del resumen)
        tiene_cant = c1 not in (None, '')
        cant = _num(c1)
        if c0 == 'VEP' and tiene_cant:
            actual['vep'] += monto
            actual['vep_cab'] += cant           # cabezas facturadas
        elif c0.startswith('CUARTEO') and tiene_cant:
            actual['cuarteo'] += monto
            actual['cuarteo_und'] += cant        # unidades cuarteadas (= medias)
        elif c0.startswith('DESPOSTADA') and tiene_cant:
            actual['despostada'] += monto
            actual['despostada_kg'] += cant      # kg despostados
        elif c0.startswith('SERVIC CONGELADO') and tiene_cant:
            actual['congelado'] += monto
            if 'STOCK' in c0:
                actual['congelado_stock'] += monto
                actual['congelado_stock_kg'] += cant
            else:                                # túnel (congelado del mes)
                actual['congelado_tunel'] += monto
                actual['congelado_tunel_kg'] += cant

    return semanas


def agrupar_por_mes(semanas):
    meses = {}
    for s in semanas:
        m = s['mes']
        d = meses.setdefault(m, {'mes': m, 'nombre': MESES.get(m, str(m)),
                                 'cab': 0, 'kg_faena': 0.0,
                                 'vep': 0.0, 'cuarteo': 0.0, 'despostada': 0.0, 'congelado': 0.0,
                                 'vep_cab': 0.0, 'cuarteo_und': 0.0, 'despostada_kg': 0.0,
                                 'congelado_tunel_kg': 0.0, 'congelado_stock_kg': 0.0,
                                 'congelado_tunel': 0.0, 'congelado_stock': 0.0, 'semanas': 0})
        for k in ('cab', 'kg_faena', 'vep', 'cuarteo', 'despostada', 'congelado',
                  'vep_cab', 'cuarteo_und', 'despostada_kg',
                  'congelado_tunel_kg', 'congelado_stock_kg', 'congelado_tunel', 'congelado_stock'):
            d[k] += s[k]
        d['semanas'] += 1
    for d in meses.values():
        d['total'] = d['vep'] + d['cuarteo'] + d['despostada'] + d['congelado']
    return dict(sorted(meses.items()))


# ───────── conexión viva a Drive ─────────
def cargar_real_drive(credentials_path, file_id=FILE_ID_DEFAULT):
    """Baja el xlsx de Drive (cuenta de servicio) y devuelve (semanas, meses)."""
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    scopes = ['https://www.googleapis.com/auth/drive.readonly']
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    service = build('drive', 'v3', credentials=creds)
    content = service.files().get_media(fileId=file_id).execute()

    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    # Una solapa por mes: recorro todas
    semanas = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        parsed = parse_cuenta_corriente(rows)
        for s in parsed:
            s['solapa'] = ws.title
        semanas += parsed
    return semanas, agrupar_por_mes(semanas)
