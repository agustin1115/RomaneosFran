"""
insumos_real.py — Lee la planilla de Stock + Compras de insumos (Google Drive,
xlsx) y calcula el REAL: consumo por SKU (stock inicial − actual) valorizado a
LIFO (última compra). Agrupa en bolsas / etiquetas / cajas para comparar contra
el teórico de los romaneos.

IMPORTANTE: el archivo debe estar compartido con la cuenta de servicio (el mismo
email que lee los romaneos).
"""
import io
import re

import openpyxl

FILE_ID_DEFAULT = '1G1KvQRNxqqJCu1SZO-ioXQclXPvF9dpA'


def _num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace('$', '').replace(',', '').replace('#N/A', '').strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def categoria_sku(codigo, nombre):
    """Agrupa el insumo en: bolsa_hueso / bolsa_chica / bolsa_grande /
    etiqueta_alto / etiqueta_auto / caja / otro."""
    cod = (codigo or '').upper()
    n = (nombre or '').upper()
    if cod.startswith('CAJ') or 'CAJA' in n or 'TAPA' in n or 'FONDO' in n:
        return 'caja'
    if 'HUESO' in n or cod.startswith('BOLVACH'):
        return 'bolsa_hueso'
    if 'BOLSA DE VACIO' in n or cod.startswith('BOLVAC'):
        m = re.search(r'(\d{2,3})\s*[Xx]\s*(\d{2,3})', n)
        if m:
            w, h = int(m.group(1)), int(m.group(2))
            return 'bolsa_chica' if w * h <= 90000 else 'bolsa_grande'
        return 'bolsa_grande'
    if 'ALTO' in n and 'IMPACTO' in n:
        return 'etiqueta_alto'
    if 'AUTOADHESIVA' in n:
        return 'etiqueta_auto'
    if cod.startswith('ETIQ') or 'ETIQUETA' in n or cod.startswith('RIBB'):
        return 'etiqueta_alto'
    return 'otro'


def parse_planilla(rows):
    """Devuelve (stock, compras). stock: {sku:{nombre,inicial,actual,cat}}.
    compras: {sku:{mesnum, precio(LIFO), und_mes, total_mes}} (compras del último mes)."""
    stock = {}
    compras_list = []
    modo = 'stock'

    for row in rows:
        if not row:
            continue
        celdas = [(_c if _c is not None else '') for _c in row]
        joined = ' '.join(str(x).upper() for x in celdas)

        if 'MES FACTURA' in joined and 'X UNIDAD' in joined:
            modo = 'compras'
            continue
        if 'STOCK INICIAL' in joined and 'STOCK ACTUAL' in joined:
            continue

        if modo == 'stock':
            cod = str(celdas[0]).strip() if len(celdas) > 0 else ''
            if not re.match(r'^[A-Z]{3,}\d', cod):
                continue
            nombre = str(celdas[1]).strip() if len(celdas) > 1 else ''
            inicial = _num(celdas[6]) if len(celdas) > 6 else 0
            actual = _num(celdas[7]) if len(celdas) > 7 else 0
            stock[cod] = {'nombre': nombre, 'inicial': inicial, 'actual': actual,
                          'cat': categoria_sku(cod, nombre)}
        else:  # compras: 1=mes(YYYYMM) 4=Codigo 6=Cant 7=$xUnidad
            cod = str(celdas[4]).strip() if len(celdas) > 4 else ''
            und = _num(celdas[6]) if len(celdas) > 6 else 0
            precio = _num(celdas[7]) if len(celdas) > 7 else 0
            if not re.match(r'^[A-Z]{3,}\d', cod) or precio <= 0:
                continue
            m = re.match(r'(\d{6})', str(celdas[1]).strip())
            compras_list.append({'cod': cod, 'mesnum': int(m.group(1)) if m else 0,
                                 'und': und, 'precio': precio})

    # Último mes global + LIFO y compras de ese mes por SKU
    ultimo_mes = max((c['mesnum'] for c in compras_list), default=0)
    compras = {}
    for c in sorted(compras_list, key=lambda x: x['mesnum']):
        compras[c['cod']] = {'mesnum': c['mesnum'], 'precio': c['precio']}  # último precio = LIFO
    for cod in compras:
        delmes = [c for c in compras_list if c['cod'] == cod and c['mesnum'] == ultimo_mes]
        compras[cod]['und_mes'] = sum(c['und'] for c in delmes)
        compras[cod]['total_mes'] = sum(c['und'] * c['precio'] for c in delmes)
    return stock, compras


def costo_real(stock, compras):
    """Consumo del mes = stock inicial + compras del mes − stock actual, por familia.
    Devuelve $ y UNIDADES por familia (para chequear 1 bolsa/etiqueta por pieza)."""
    fam = {}
    detalle = []
    for cod, s in stock.items():
        c = compras.get(cod, {})
        precio = c.get('precio', 0)
        compras_mes = c.get('und_mes', 0)
        consumo_und = max(0.0, s['inicial'] + compras_mes - s['actual'])
        costo = consumo_und * precio
        cat = s['cat']
        f = fam.setdefault(cat, {'costo': 0.0, 'consumo_und': 0.0})
        f['costo'] += costo
        f['consumo_und'] += consumo_und
        if consumo_und > 0:
            detalle.append({'sku': cod, 'nombre': s['nombre'], 'cat': cat,
                            'consumo_und': consumo_und, 'precio_lifo': precio, 'costo': costo})

    def _suma(pred, key):
        return sum(f[key] for k, f in fam.items() if pred(k))

    bolsas = _suma(lambda k: k.startswith('bolsa'), 'costo')
    bolsas_und = _suma(lambda k: k.startswith('bolsa'), 'consumo_und')
    etiquetas = fam.get('etiqueta_alto', {}).get('costo', 0.0)
    etiquetas_und = fam.get('etiqueta_alto', {}).get('consumo_und', 0.0)
    caja = fam.get('caja', {}).get('costo', 0.0)
    caja_und = fam.get('caja', {}).get('consumo_und', 0.0)
    return {'bolsas': bolsas, 'bolsas_und': bolsas_und,
            'etiquetas': etiquetas, 'etiquetas_und': etiquetas_und,
            'caja': caja, 'caja_und': caja_und, 'total': bolsas + etiquetas + caja,
            'por_categoria': fam, 'detalle': sorted(detalle, key=lambda x: -x['costo'])}


def cargar_insumos_real_drive(credentials_path, file_id=FILE_ID_DEFAULT):
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    scopes = ['https://www.googleapis.com/auth/drive.readonly']
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    service = build('drive', 'v3', credentials=creds)
    content = service.files().get_media(fileId=file_id).execute()
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    stock, compras = parse_planilla(rows)
    return stock, compras, costo_real(stock, compras)
