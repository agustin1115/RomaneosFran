"""
excel_builder.py — Genera Excel de análisis de romaneo con soporte multi-calidad.
Evolución de build_analisis.py v4 con perfiles Búfalo, Premium Black y Exportación.
"""
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

from config import (
    GRUPOS_POR_CALIDAD, CORTES_CAROS, REND_OBJETIVO,
    COSTOS_PERFILES, AMARILLA_CONTRAMARCAS, AMARILLA_PRECIO_DEFAULT,
    CONTRAMARCA_MAP, PRECIOS_FIJOS_RECORTE,
)

# ── Styles ──
HDR_FILL = PatternFill('solid', fgColor='1F4E79')
HDR_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=10)
SEC_FILL = PatternFill('solid', fgColor='D6E4F0')
SEC_FONT = Font(name='Arial', bold=True, size=10)
DATA = Font(name='Arial', size=10)
BLUE = Font(name='Arial', size=10, color='0000FF')
TOTAL_FILL = PatternFill('solid', fgColor='E2EFDA')
TOTAL_FONT = Font(name='Arial', bold=True, size=10)
RED = Font(name='Arial', size=10, color='FF0000')
GREEN = Font(name='Arial', size=10, color='008000')
CF_GREEN = PatternFill('solid', fgColor='E2EFDA')
CF_BLUE = PatternFill('solid', fgColor='D6E4F0')
CF_YELLOW = PatternFill('solid', fgColor='FFF2CC')
CF_RED = PatternFill('solid', fgColor='FCE4EC')
ORANGE_FONT = Font(name='Arial', bold=True, size=10, color='FF6600')
AMARILLA_FILL = PatternFill('solid', fgColor='FFFFCC')

# Colores por perfil de calidad
PERFIL_COLORS = {
    'Standard':      '1F4E79',
    'Búfalo':        '6D4C41',
    'Premium Black': '212121',
    'Exportación':   '1B5E20',
}

PERFIL_LABELS = {
    'Standard':      'ESTÁNDAR — Mercado Interno',
    'Búfalo':        'BÚFALO — Carne de Búfalo',
    'Premium Black': 'PREMIUM BLACK — Selección Premium',
    'Exportación':   'EXPORTACIÓN — Mercado Internacional',
}


def hdr_row(ws, row, headers, fill=HDR_FILL, font=HDR_FONT):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h)
        c.fill, c.font = fill, font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def build_analisis(data, output_path, calidad='Standard', costos_override=None,
                   precio_amarilla=None, price_matrix=None):
    """
    Genera un Excel de 6 hojas con análisis de romaneo.

    Args:
        data: dict con 'romaneo' y 'cortes' (output de pdf_parser)
        output_path: ruta del Excel de salida
        calidad: 'Standard' | 'Búfalo' | 'Premium Black' | 'Exportación'
        costos_override: dict opcional para sobreescribir costos del perfil
        precio_amarilla: precio $/kg amarilla (default 10500)
        price_matrix: dict {GRUPO: {CLIENTE_COL: precio}}
    """
    if price_matrix is None:
        price_matrix = {}

    GRUPOS = GRUPOS_POR_CALIDAD.get(calidad, GRUPOS_POR_CALIDAD['Standard'])
    costos_base = COSTOS_PERFILES.get(calidad, COSTOS_PERFILES['Standard']).copy()
    if costos_override:
        costos_base.update(costos_override)
    costos = costos_base
    precio_am = precio_amarilla or AMARILLA_PRECIO_DEFAULT
    color_perfil = PERFIL_COLORS.get(calidad, '1F4E79')
    label_perfil = PERFIL_LABELS.get(calidad, calidad)

    rom = data['romaneo'] if 'romaneo' in data else data
    cortes = data.get('cortes', rom.get('cortes', []))

    # Contramarca → pricing column mapping
    cm_to_client = {}
    for c in cortes:
        cm = str(c.get('contramarca', ''))
        if cm and cm in CONTRAMARCA_MAP:
            _, pricing_col = CONTRAMARCA_MAP[cm]
            cm_to_client[cm] = pricing_col

    # Rendimiento objetivo
    categoria = rom.get('categoria', 'Vaca')
    cat_clean = categoria.split('(')[0].strip()
    rend_tabla = REND_OBJETIVO.get(calidad, REND_OBJETIVO['Standard'])
    rend_obj_base = rend_tabla.get(cat_clean, 0.66)

    meat_cortes = [c for c in cortes if c.get('grupo') != 'GRASA']
    kg_total_meat = sum(c['kg'] for c in meat_cortes)
    kg_anatomico = sum(c['kg'] for c in meat_cortes if c.get('tipo', '') == 'ANATÓMICO')
    pct_anatomico = kg_anatomico / kg_total_meat if kg_total_meat > 0 else 0
    real_rend = kg_total_meat / rom.get('kg_entrada', 1) if rom.get('kg_entrada', 0) > 0 else 0
    ajuste_anatomico = 0.01 if pct_anatomico > 0.50 else 0
    rend_obj = rend_obj_base + ajuste_anatomico

    wb = Workbook()
    N_P = len(GRUPOS)

    # Clientes únicos
    clientes_romaneo = {}
    for c in cortes:
        if c.get('grupo') == 'GRASA':
            continue
        cli = c.get('cliente', 'SIN ASIGNAR')
        cm = str(c.get('contramarca', ''))
        if cli not in clientes_romaneo:
            pricing_col = cm_to_client.get(cm, 'RESTO CLIENTES AMBA')
            es_amarilla = cm in AMARILLA_CONTRAMARCAS
            clientes_romaneo[cli] = {'contramarca': cm, 'pricing_col': pricing_col, 'es_amarilla': es_amarilla}
    cli_list = sorted(clientes_romaneo.keys())

    # ══════ PARAMETROS ══════
    ws = wb.active
    ws.title = "PARAMETROS"
    ws.sheet_properties.tabColor = color_perfil
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 20

    ws.merge_cells('A1:B1')
    ws['A1'] = f'PARÁMETROS — {label_perfil}'
    ws['A1'].font = Font(name='Arial', bold=True, size=14, color=color_perfil)

    ws['A3'] = 'PERFIL DE CALIDAD'
    ws['A3'].font = Font(name='Arial', bold=True, size=11, color=color_perfil)
    ws['A3'].fill = SEC_FILL; ws['B3'].fill = SEC_FILL
    ws.cell(4, 1, 'Calidad seleccionada').font = DATA
    ws.cell(4, 2, calidad).font = Font(name='Arial', bold=True, size=12, color=color_perfil)

    ws['A6'] = 'COSTOS VARIABLES (por kg de carne producida)'
    ws['A6'].font = Font(name='Arial', bold=True, size=11, color=color_perfil)
    ws['A6'].fill = SEC_FILL; ws['B6'].fill = SEC_FILL

    cost_items = [
        (7, 'Mano de obra (despostada)', costos['mo']),
        (8, 'Insumos (envasado, etiquetado, etc.)', costos['insumos']),
        (9, 'Flete / transporte', costos['flete']),
        (10, 'SENASA + cuarteo', costos['senasa']),
    ]
    for r, label, val in cost_items:
        ws.cell(r, 1, label).font = DATA
        ws.cell(r, 2, val).font = BLUE; ws[f'B{r}'].number_format = '#,##0'

    ws.cell(11, 1, 'IIBB y otros impuestos (% sobre venta)').font = DATA
    ws.cell(11, 2, costos['iibb']).font = BLUE; ws['B11'].number_format = '0.0%'
    ws.cell(12, 1, 'TOTAL COSTO VARIABLE / kg prod. (sin IIBB)').font = TOTAL_FONT
    ws.cell(12, 1).fill = TOTAL_FILL
    ws['B12'] = '=SUM(B7:B10)'; ws['B12'].font = TOTAL_FONT; ws['B12'].fill = TOTAL_FILL
    ws['B12'].number_format = '#,##0'

    ws['A14'] = 'REFERENCIA MERCADO (MAG Cañuelas)'
    ws['A14'].font = Font(name='Arial', bold=True, size=11, color=color_perfil)
    ws['A14'].fill = SEC_FILL; ws['B14'].fill = SEC_FILL
    ws.cell(15, 1, 'Precio MAG referencia $/kg media res').font = DATA
    ws.cell(15, 2, rom.get('precio_mag', 0)).font = BLUE; ws['B15'].number_format = '#,##0'

    ws['A17'] = 'PRECIO AMARILLA ($/kg, todos los cortes)'
    ws['A17'].font = Font(name='Arial', bold=True, size=11, color='FF6600')
    ws['A17'].fill = AMARILLA_FILL; ws['B17'].fill = AMARILLA_FILL
    ws.cell(17, 2, precio_am).font = Font(name='Arial', bold=True, size=12, color='0000FF')
    ws['B17'].number_format = '#,##0'

    ws['A19'] = 'RENDIMIENTO OBJETIVO'
    ws['A19'].font = Font(name='Arial', bold=True, size=11, color=color_perfil)
    ws['A19'].fill = SEC_FILL; ws['B19'].fill = SEC_FILL
    ws.cell(20, 1, f'Categoría: {categoria}').font = DATA
    ws.cell(20, 2, rend_obj_base).font = BLUE; ws['B20'].number_format = '0.0%'
    ws.cell(21, 1, f'Ajuste anatómico (>50% = +1%): {pct_anatomico*100:.0f}% anatómico').font = DATA
    ws.cell(21, 2, ajuste_anatomico).font = BLUE; ws['B21'].number_format = '+0.0%;-0.0%;0.0%'
    ws.cell(22, 1, 'RENDIMIENTO OBJETIVO FINAL').font = TOTAL_FONT; ws.cell(22, 1).fill = TOTAL_FILL
    ws.cell(22, 2, rend_obj).font = Font(name='Arial', bold=True, size=12, color='0000FF')
    ws['B22'].number_format = '0.0%'; ws.cell(22, 2).fill = TOTAL_FILL

    # Tabla comparativa de costos por perfil
    ws['A24'] = 'COMPARATIVA DE PERFILES'
    ws['A24'].font = Font(name='Arial', bold=True, size=11, color=color_perfil)
    ws['A24'].fill = SEC_FILL; ws['B24'].fill = SEC_FILL
    hdr_row(ws, 25, ['Concepto', 'Standard', 'Búfalo', 'Premium Black', 'Exportación'])
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 14

    for i, (key, label) in enumerate([
        ('mo', 'Mano de obra'),
        ('insumos', 'Insumos'),
        ('flete', 'Flete'),
        ('senasa', 'SENASA'),
        ('iibb', 'IIBB'),
    ]):
        r = 26 + i
        ws.cell(r, 1, label).font = DATA
        for j, perfil in enumerate(['Standard', 'Búfalo', 'Premium Black', 'Exportación']):
            val = COSTOS_PERFILES[perfil][key]
            ws.cell(r, 2 + j, val).font = DATA
            if key == 'iibb':
                ws.cell(r, 2 + j).number_format = '0.0%'
            else:
                ws.cell(r, 2 + j).number_format = '#,##0'
            # Resaltar perfil activo
            if perfil == calidad:
                ws.cell(r, 2 + j).font = Font(name='Arial', bold=True, size=10, color=color_perfil)

    # ══════ PRECIOS ══════
    ws_p = wb.create_sheet("PRECIOS")
    ws_p.sheet_properties.tabColor = "2E7D32"
    ws_p.column_dimensions['A'].width = 22
    ws_p.column_dimensions['B'].width = 14
    ws_p.column_dimensions['C'].width = 10

    ws_p.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3 + len(cli_list))
    ws_p['A1'] = f'PRECIOS DE VENTA — {calidad.upper()}'
    ws_p['A1'].font = Font(name='Arial', bold=True, size=12, color='2E7D32')

    hdrs = ['Corte', '% x Tropa', 'Tipo $'] + [
        f'{cli}\n(ctm {clientes_romaneo[cli]["contramarca"]})' for cli in cli_list
    ]
    P_HDR = 3
    hdr_row(ws_p, P_HDR, hdrs, PatternFill('solid', fgColor='2E7D32'))
    for i in range(len(cli_list)):
        col = 4 + i
        ws_p.column_dimensions[get_column_letter(col)].width = 18
        cli = cli_list[i]
        if clientes_romaneo[cli]['es_amarilla']:
            ws_p.cell(P_HDR, col).fill = PatternFill('solid', fgColor='FF8C00')

    PREC_FIRST = P_HDR + 1
    PREC_LAST = PREC_FIRST + N_P - 1

    for i, (grupo, pct) in enumerate(GRUPOS):
        r = PREC_FIRST + i
        ws_p.cell(r, 1, grupo).font = DATA
        ws_p.cell(r, 2, pct).font = BLUE; ws_p[f'B{r}'].number_format = '0.00%'
        es_caro = grupo in CORTES_CAROS
        ws_p.cell(r, 3, 'CARO' if es_caro else 'BARATO').font = Font(
            name='Arial', bold=True, size=9,
            color='C62828' if es_caro else '2E7D32'
        )
        ws_p.cell(r, 3).alignment = Alignment(horizontal='center')

        for j, cli in enumerate(cli_list):
            col = 4 + j
            cli_info = clientes_romaneo[cli]
            # Recortes: precio fijo independiente del cliente
            if grupo in PRECIOS_FIJOS_RECORTE:
                price = PRECIOS_FIJOS_RECORTE[grupo]
                ws_p.cell(r, col, price).font = BLUE
                ws_p.cell(r, col).number_format = '#,##0'
            elif cli_info['es_amarilla']:
                ws_p.cell(r, col).value = '=PARAMETROS!$B$17'
                ws_p.cell(r, col).font = ORANGE_FONT
                ws_p.cell(r, col).fill = AMARILLA_FILL
            else:
                pricing_col = cli_info['pricing_col']
                price = None
                if grupo in price_matrix and pricing_col in price_matrix[grupo]:
                    price = price_matrix[grupo][pricing_col]
                elif grupo in price_matrix:
                    for fb in ['NETOS PEYA', 'RESTO CLIENTES AMBA']:
                        if fb in price_matrix[grupo]:
                            price = price_matrix[grupo][fb]; break
                ws_p.cell(r, col, round(price) if price else 0).font = BLUE if price else RED
                ws_p.cell(r, col).number_format = '#,##0'

    r_tot_p = PREC_LAST + 1
    ws_p.cell(r_tot_p, 1, 'TOTAL').font = TOTAL_FONT
    ws_p.cell(r_tot_p, 1).fill = TOTAL_FILL; ws_p.cell(r_tot_p, 2).fill = TOTAL_FILL
    ws_p[f'B{r_tot_p}'] = f'=SUM(B{PREC_FIRST}:B{PREC_LAST})'
    ws_p[f'B{r_tot_p}'].number_format = '0.00%'
    ws_p[f'B{r_tot_p}'].font = TOTAL_FONT; ws_p[f'B{r_tot_p}'].fill = TOTAL_FILL

    # ══════ ROMANEO ══════
    ws_r = wb.create_sheet("ROMANEO")
    ws_r.sheet_properties.tabColor = "E65100"
    for col, w in {'A': 30, 'B': 22, 'C': 14, 'D': 12, 'E': 12,
                   'F': 14, 'G': 12, 'H': 14, 'I': 18, 'J': 14}.items():
        ws_r.column_dimensions[col].width = w

    ws_r.merge_cells('A1:J1')
    titulo_rom = rom.get('numero', 'S/N')
    fecha_rom = rom.get('fecha', '')
    ws_r['A1'] = f'ROMANEO N° {titulo_rom} — {fecha_rom} — {calidad.upper()}'
    ws_r['A1'].font = Font(name='Arial', bold=True, size=14, color='E65100')

    hdr_items = [
        (3, 'N° Romaneo', rom.get('numero', '')),
        (4, 'Fecha', rom.get('fecha', '')),
        (5, 'Medias reses', rom.get('medias_reses', 0)),
        (6, 'Kg entrada (media res)', rom.get('kg_entrada', 0)),
        (7, 'Categoría (principal)', rom.get('categoria', '')),
        (8, 'Precio compra $/kg media res (s/IVA)', rom.get('precio_compra', 0)),
        (9, f'Rendimiento objetivo ({cat_clean})', rend_obj),
        (10, f'Rendimiento real', round(real_rend, 4)),
        (11, 'Perfil de calidad', calidad),
    ]
    for r, label, val in hdr_items:
        ws_r.cell(r, 1, label).font = SEC_FONT; ws_r.cell(r, 1).fill = SEC_FILL
        ws_r.cell(r, 2, val).font = BLUE
        if isinstance(val, (int, float)) and val != 0:
            ws_r[f'B{r}'].number_format = '#,##0' if val > 1 else '0.0%'

    # Desglose categoría (col D-F)
    desglose_cat = rom.get('desglose_categoria', data.get('desglose_categoria', {}))
    if desglose_cat:
        ws_r.cell(3, 4, 'Desglose categoría').font = SEC_FONT; ws_r.cell(3, 4).fill = SEC_FILL
        ws_r.column_dimensions['D'].width = 20
        ws_r.column_dimensions['E'].width = 12
        r_cat = 4
        for cat_name, cat_data in desglose_cat.items():
            ws_r.cell(r_cat, 4, cat_name).font = DATA
            ws_r.cell(r_cat, 5, round(cat_data['pct'], 1) / 100).font = BLUE
            ws_r[f'E{r_cat}'].number_format = '0.0%'
            r_cat += 1

    # Desglose tipificación (col G-H)
    desglose_tip = rom.get('desglose_tipificacion', data.get('desglose_tipificacion', {}))
    if desglose_tip:
        ws_r.cell(3, 7, 'Tipificación').font = SEC_FONT; ws_r.cell(3, 7).fill = SEC_FILL
        ws_r.column_dimensions['G'].width = 14
        total_tip = sum(desglose_tip.values())
        r_tip = 4
        for tip, kg in sorted(desglose_tip.items()):
            pct = kg / total_tip if total_tip > 0 else 0
            ws_r.cell(r_tip, 7, tip).font = DATA
            ws_r.cell(r_tip, 8, round(pct, 3)).font = BLUE
            ws_r[f'H{r_tip}'].number_format = '0.0%'
            r_tip += 1

    DET_HDR = 13
    hdr_row(ws_r, DET_HDR,
            ['Corte (detalle)', 'Grupo Precio', 'Tipo', 'Piezas', 'Unidades',
             'Kg', 'Destino', 'Rend %', 'Cliente', 'Nro.Venta'],
            PatternFill('solid', fgColor='E65100'))

    N_DET = len(cortes)
    DET_FIRST = DET_HDR + 1
    DET_LAST = DET_FIRST + N_DET - 1

    for i, c in enumerate(cortes):
        r = DET_FIRST + i
        ws_r.cell(r, 1, c['corte']).font = DATA
        ws_r.cell(r, 2, c['grupo']).font = DATA
        ws_r.cell(r, 3, c.get('tipo', 'ANATÓMICO')).font = DATA
        ws_r.cell(r, 4, c.get('piezas', 0)).font = DATA; ws_r[f'D{r}'].number_format = '#,##0'
        ws_r.cell(r, 5, c.get('unidades', 0)).font = DATA
        ws_r.cell(r, 6, c['kg']).font = DATA; ws_r[f'F{r}'].number_format = '#,##0.00'
        ws_r.cell(r, 7, c.get('destino', '')).font = DATA
        ws_r[f'H{r}'] = f'=IF($B$6=0,0,F{r}/$B$6)'
        ws_r[f'H{r}'].number_format = '0.00%'; ws_r[f'H{r}'].font = DATA
        ws_r.cell(r, 9, c.get('cliente', '')).font = DATA
        ws_r.cell(r, 10, c.get('nro_venta', '')).font = DATA
        cm = str(c.get('contramarca', ''))
        if cm in AMARILLA_CONTRAMARCAS:
            for col in range(1, 11):
                ws_r.cell(r, col).fill = AMARILLA_FILL

    R_TOT = DET_LAST + 1
    ws_r.cell(R_TOT, 1, 'TOTAL SALIDAS').font = TOTAL_FONT
    for col in range(1, 11):
        ws_r.cell(R_TOT, col).fill = TOTAL_FILL
    ws_r[f'D{R_TOT}'] = f'=SUM(D{DET_FIRST}:D{DET_LAST})'
    ws_r[f'D{R_TOT}'].number_format = '#,##0'; ws_r[f'D{R_TOT}'].font = TOTAL_FONT
    ws_r[f'E{R_TOT}'] = f'=SUM(E{DET_FIRST}:E{DET_LAST})'
    ws_r[f'E{R_TOT}'].font = TOTAL_FONT
    ws_r[f'F{R_TOT}'] = f'=SUM(F{DET_FIRST}:F{DET_LAST})'
    ws_r[f'F{R_TOT}'].number_format = '#,##0.00'; ws_r[f'F{R_TOT}'].font = TOTAL_FONT
    ws_r[f'H{R_TOT}'] = f'=IF($B$6=0,0,F{R_TOT}/$B$6)'
    ws_r[f'H{R_TOT}'].number_format = '0.00%'; ws_r[f'H{R_TOT}'].font = TOTAL_FONT

    R_MERMA = R_TOT + 1
    ws_r.cell(R_MERMA, 1, 'MERMA (oreo/proceso)').font = Font(
        name='Arial', bold=True, size=10, color='FF0000')
    ws_r[f'F{R_MERMA}'] = f'=B6-F{R_TOT}'
    ws_r[f'F{R_MERMA}'].number_format = '#,##0.00'
    ws_r[f'F{R_MERMA}'].font = Font(name='Arial', bold=True, size=10, color='FF0000')
    ws_r[f'H{R_MERMA}'] = f'=IF($B$6=0,0,F{R_MERMA}/$B$6)'
    ws_r[f'H{R_MERMA}'].number_format = '0.00%'
    ws_r[f'H{R_MERMA}'].font = Font(name='Arial', bold=True, size=10, color='FF0000')

    R_CARNE = R_MERMA + 1
    ws_r.cell(R_CARNE, 1, 'KG CARNE PRODUCIDA (sin grasa/decomiso)').font = Font(
        name='Arial', bold=True, size=10, color='2E7D32')
    ws_r[f'F{R_CARNE}'] = f'=F{R_TOT}-SUMIF(B{DET_FIRST}:B{DET_LAST},"GRASA",F{DET_FIRST}:F{DET_LAST})'
    ws_r[f'F{R_CARNE}'].number_format = '#,##0.00'
    ws_r[f'F{R_CARNE}'].font = Font(name='Arial', bold=True, size=10, color='2E7D32')
    ws_r[f'H{R_CARNE}'] = f'=IF($B$6=0,0,F{R_CARNE}/$B$6)'
    ws_r[f'H{R_CARNE}'].number_format = '0.00%'
    ws_r[f'H{R_CARNE}'].font = Font(name='Arial', bold=True, size=10, color='2E7D32')

    R_ROBJ = R_CARNE + 1
    ws_r.cell(R_ROBJ, 1, 'REND vs OBJETIVO').font = Font(
        name='Arial', bold=True, size=10, color='1F4E79')
    diff_rend = real_rend - rend_obj
    ws_r.cell(R_ROBJ, 6, diff_rend).font = Font(
        name='Arial', bold=True, size=10,
        color='008000' if diff_rend >= 0 else 'FF0000'
    )
    ws_r[f'F{R_ROBJ}'].number_format = '+0.0%;-0.0%'
    ws_r.cell(R_ROBJ, 8, 'OK' if diff_rend >= 0 else 'BAJO OBJETIVO').font = Font(
        name='Arial', bold=True, size=10,
        color='008000' if diff_rend >= 0 else 'FF0000'
    )

    # Amarillas summary
    R_AM = R_ROBJ + 2
    ws_r.cell(R_AM, 1, 'AMARILLAS (ctmarcas 47/73/74)').font = ORANGE_FONT
    ws_r.cell(R_AM, 1).fill = AMARILLA_FILL
    kg_amarilla = sum(c['kg'] for c in cortes
                      if str(c.get('contramarca', '')) in AMARILLA_CONTRAMARCAS)
    kg_total = sum(c['kg'] for c in cortes)
    ws_r.cell(R_AM, 6, kg_amarilla).font = ORANGE_FONT
    ws_r[f'F{R_AM}'].number_format = '#,##0.00'
    pct_am = kg_amarilla / kg_total if kg_total > 0 else 0
    ws_r.cell(R_AM, 8, pct_am).font = ORANGE_FONT
    ws_r[f'H{R_AM}'].number_format = '0.0%'

    # ══════ ANALISIS ══════
    ws_a = wb.create_sheet("ANALISIS")
    ws_a.sheet_properties.tabColor = "6A1B9A"
    for col, w in [('A', 24), ('B', 14), ('C', 14), ('D', 14),
                   ('E', 14), ('F', 14), ('G', 10), ('H', 16)]:
        ws_a.column_dimensions[col].width = w

    ws_a.merge_cells('A1:H1')
    ws_a['A1'] = f'ANÁLISIS DE RENDIMIENTO — {titulo_rom} — {calidad.upper()}'
    ws_a['A1'].font = Font(name='Arial', bold=True, size=14, color='6A1B9A')

    ws_a['A3'] = 'Kg entrada'; ws_a['A3'].font = SEC_FONT
    ws_a['B3'] = '=ROMANEO!B6'; ws_a['B3'].number_format = '#,##0'
    ws_a['B3'].font = Font(name='Arial', bold=True, size=12)
    ws_a['C3'] = 'Kg carne'; ws_a['C3'].font = SEC_FONT
    ws_a['D3'] = f'=ROMANEO!F{R_CARNE}'; ws_a['D3'].number_format = '#,##0.00'
    ws_a['D3'].font = Font(name='Arial', bold=True, size=12, color='2E7D32')
    ws_a['E3'] = 'Rend. cárnico'; ws_a['E3'].font = SEC_FONT
    ws_a['F3'] = '=IF(B3=0,0,D3/B3)'; ws_a['F3'].number_format = '0.00%'
    ws_a['F3'].font = Font(name='Arial', bold=True, size=12, color='6A1B9A')
    ws_a['G3'] = 'Obj.'; ws_a['G3'].font = SEC_FONT
    ws_a['H3'] = rend_obj; ws_a['H3'].number_format = '0.0%'
    ws_a['H3'].font = Font(name='Arial', bold=True, size=12, color='1F4E79')

    AN_HDR = 6
    hdr_row(ws_a, AN_HDR,
            ['Grupo Corte', 'Kg Real', '% Real s/Carne', '% Esperado',
             'Desvío', '% Cumplim.', 'Tipo $', 'Calificación'],
            PatternFill('solid', fgColor='6A1B9A'))

    AN_FIRST = AN_HDR + 1
    AN_LAST = AN_FIRST + N_P - 1
    grp = f'ROMANEO!$B${DET_FIRST}:$B${DET_LAST}'
    kgr = f'ROMANEO!$F${DET_FIRST}:$F${DET_LAST}'

    for i in range(N_P):
        r = AN_FIRST + i
        p_r = PREC_FIRST + i
        grupo_name = GRUPOS[i][0]
        es_caro = grupo_name in CORTES_CAROS

        ws_a[f'A{r}'] = f'=PRECIOS!A{p_r}'; ws_a[f'A{r}'].font = DATA
        ws_a[f'B{r}'] = f'=SUMIF({grp},A{r},{kgr})'
        ws_a[f'B{r}'].number_format = '#,##0.00'; ws_a[f'B{r}'].font = DATA
        ws_a[f'C{r}'] = f'=IF(ROMANEO!$F${R_CARNE}=0,0,B{r}/ROMANEO!$F${R_CARNE})'
        ws_a[f'C{r}'].number_format = '0.00%'; ws_a[f'C{r}'].font = DATA
        ws_a[f'D{r}'] = f'=PRECIOS!B{p_r}'
        ws_a[f'D{r}'].number_format = '0.00%'; ws_a[f'D{r}'].font = DATA
        ws_a[f'E{r}'] = f'=C{r}-D{r}'
        ws_a[f'E{r}'].number_format = '+0.00%;-0.00%;0.00%'; ws_a[f'E{r}'].font = DATA
        ws_a[f'F{r}'] = f'=IF(D{r}=0,IF(B{r}=0,1,1.5),C{r}/D{r})'
        ws_a[f'F{r}'].number_format = '0.0%'; ws_a[f'F{r}'].font = DATA

        ws_a.cell(r, 7, 'CARO' if es_caro else 'BARATO').font = Font(
            name='Arial', bold=True, size=9,
            color='C62828' if es_caro else '2E7D32'
        )
        ws_a.cell(r, 7).alignment = Alignment(horizontal='center')

        if es_caro:
            formula = (
                f'=IF(D{r}=0,IF(B{r}=0,"N/A","EXTRA"),'
                f'IF(F{r}>=1.1,"ÓPTIMO",IF(F{r}>=0.95,"BUENO",IF(F{r}>=0.8,"REGULAR","MALO"))))'
            )
        else:
            formula = (
                f'=IF(D{r}=0,IF(B{r}=0,"N/A","EXTRA"),'
                f'IF(F{r}<=0.8,"ÓPTIMO",IF(F{r}<=0.95,"BUENO",IF(F{r}<=1.1,"REGULAR","MALO"))))'
            )
        ws_a[f'H{r}'] = formula
        ws_a[f'H{r}'].font = Font(name='Arial', bold=True, size=10)
        ws_a[f'H{r}'].alignment = Alignment(horizontal='center')

    R_TOT_AN = AN_LAST + 1
    ws_a.cell(R_TOT_AN, 1, 'TOTAL CARNE').font = TOTAL_FONT
    for col in range(1, 9):
        ws_a.cell(R_TOT_AN, col).fill = TOTAL_FILL
    ws_a[f'B{R_TOT_AN}'] = f'=SUM(B{AN_FIRST}:B{AN_LAST})'
    ws_a[f'B{R_TOT_AN}'].number_format = '#,##0.00'; ws_a[f'B{R_TOT_AN}'].font = TOTAL_FONT
    ws_a[f'C{R_TOT_AN}'] = f'=IF(ROMANEO!$F${R_CARNE}=0,0,B{R_TOT_AN}/ROMANEO!$F${R_CARNE})'
    ws_a[f'C{R_TOT_AN}'].number_format = '0.00%'; ws_a[f'C{R_TOT_AN}'].font = TOTAL_FONT

    cal_range = f'H{AN_FIRST}:H{AN_LAST}'
    ws_a.conditional_formatting.add(cal_range, CellIsRule(operator='equal', formula=['"ÓPTIMO"'], fill=CF_GREEN))
    ws_a.conditional_formatting.add(cal_range, CellIsRule(operator='equal', formula=['"BUENO"'], fill=CF_BLUE))
    ws_a.conditional_formatting.add(cal_range, CellIsRule(operator='equal', formula=['"REGULAR"'], fill=CF_YELLOW))
    ws_a.conditional_formatting.add(cal_range, CellIsRule(operator='equal', formula=['"MALO"'], fill=CF_RED))

    # ══════ VENTAS ══════
    ws_v = wb.create_sheet("VENTAS")
    ws_v.sheet_properties.tabColor = "FF6600"
    for col, w in [('A', 22), ('B', 14), ('C', 14), ('D', 16),
                   ('E', 14), ('F', 18), ('G', 16)]:
        ws_v.column_dimensions[col].width = w

    ws_v.merge_cells('A1:G1')
    ws_v['A1'] = f'DESGLOSE POR CLIENTE — {titulo_rom} — {calidad.upper()}'
    ws_v['A1'].font = Font(name='Arial', bold=True, size=14, color='FF6600')

    V_HDR = 3
    hdr_row(ws_v, V_HDR,
            ['Cliente', 'Ctmarca', 'Kg Totales', '% del Total', 'Amarilla', 'Ingreso Est. $', 'Alerta'],
            PatternFill('solid', fgColor='FF6600'))

    client_data = {}
    for c in cortes:
        if c.get('grupo') == 'GRASA':
            continue
        cli = c.get('cliente', 'SIN ASIGNAR')
        cm = str(c.get('contramarca', ''))
        if cli not in client_data:
            client_data[cli] = {'kg': 0, 'ingreso': 0, 'cm': cm}
        client_data[cli]['kg'] += c['kg']
        grupo = c['grupo']
        if grupo in PRECIOS_FIJOS_RECORTE:
            price = PRECIOS_FIJOS_RECORTE[grupo]
        elif cm in AMARILLA_CONTRAMARCAS:
            price = precio_am
        else:
            pricing_col = cm_to_client.get(cm, 'RESTO CLIENTES AMBA')
            price = 0
            if grupo in price_matrix and pricing_col in price_matrix[grupo]:
                price = price_matrix[grupo][pricing_col]
            elif grupo in price_matrix:
                for fb in ['NETOS PEYA', 'RESTO CLIENTES AMBA']:
                    if fb in price_matrix[grupo]:
                        price = price_matrix[grupo][fb]; break
        client_data[cli]['ingreso'] += c['kg'] * price

    v_row = V_HDR + 1
    v_first = v_row
    for cli_name in sorted(client_data.keys()):
        cd = client_data[cli_name]
        ws_v.cell(v_row, 1, cli_name).font = DATA
        ws_v.cell(v_row, 2, cd['cm']).font = DATA
        ws_v.cell(v_row, 3, round(cd['kg'], 2)).font = DATA
        ws_v[f'C{v_row}'].number_format = '#,##0.00'
        ws_v[f'D{v_row}'] = f'=IF(C{v_row}=0,0,C{v_row}/SUM(C{v_first}:C{v_first+len(client_data)-1}))'
        ws_v[f'D{v_row}'].number_format = '0.0%'; ws_v[f'D{v_row}'].font = DATA
        es_am = cd['cm'] in AMARILLA_CONTRAMARCAS
        ws_v.cell(v_row, 5, 'SÍ' if es_am else 'NO').font = ORANGE_FONT if es_am else DATA
        ws_v.cell(v_row, 6, round(cd['ingreso'])).font = DATA
        ws_v[f'F{v_row}'].number_format = '$#,##0'
        if es_am:
            ws_v.cell(v_row, 7, f'AMARILLA ${precio_am:,.0f}/kg').font = ORANGE_FONT
            for col in range(1, 8):
                ws_v.cell(v_row, col).fill = AMARILLA_FILL
        v_row += 1

    ws_v.cell(v_row, 1, 'TOTAL').font = TOTAL_FONT
    for col in range(1, 8):
        ws_v.cell(v_row, col).fill = TOTAL_FILL
    ws_v[f'C{v_row}'] = f'=SUM(C{v_first}:C{v_row-1})'
    ws_v[f'C{v_row}'].number_format = '#,##0.00'; ws_v[f'C{v_row}'].font = TOTAL_FONT
    ws_v[f'F{v_row}'] = f'=SUM(F{v_first}:F{v_row-1})'
    ws_v[f'F{v_row}'].number_format = '$#,##0'; ws_v[f'F{v_row}'].font = TOTAL_FONT

    # Precio venta promedio
    v_pvp = v_row + 2
    ws_v.cell(v_pvp, 1, 'PRECIO VENTA PROMEDIO ($/kg carne)').font = Font(
        name='Arial', bold=True, size=11, color='1F4E79')
    ws_v.cell(v_pvp, 1).fill = SEC_FILL
    ws_v[f'F{v_pvp}'] = f'=IF(C{v_row}=0,0,F{v_row}/C{v_row})'
    ws_v[f'F{v_pvp}'].number_format = '$#,##0'
    ws_v[f'F{v_pvp}'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
    ws_v[f'F{v_pvp}'].fill = SEC_FILL

    # ══════ RESULTADO ══════
    ws_res = wb.create_sheet("RESULTADO")
    ws_res.sheet_properties.tabColor = "C62828"
    ws_res.column_dimensions['A'].width = 46
    ws_res.column_dimensions['B'].width = 24
    for c_col in ['C', 'D', 'E']:
        ws_res.column_dimensions[c_col].width = 18

    ws_res.merge_cells('A1:E1')
    ws_res['A1'] = f'RESULTADO ECONÓMICO — {titulo_rom} — {calidad.upper()}'
    ws_res['A1'].font = Font(name='Arial', bold=True, size=14, color='C62828')

    ws_res['A3'] = 'COSTOS'
    ws_res['A3'].font = Font(name='Arial', bold=True, size=12, color='C62828')
    ws_res['A3'].fill = PatternFill('solid', fgColor='FCE4EC')
    ws_res['B3'].fill = PatternFill('solid', fgColor='FCE4EC')

    ws_res.cell(4, 1, 'Compra hacienda (Kg entrada x Precio compra)').font = DATA
    ws_res['B4'] = '=ROMANEO!B6*ROMANEO!B8'; ws_res['B4'].number_format = '$#,##0'; ws_res['B4'].font = RED
    ws_res.cell(5, 1, 'Mano de obra').font = DATA
    ws_res['B5'] = f'=ROMANEO!F{R_CARNE}*PARAMETROS!B7'; ws_res['B5'].number_format = '$#,##0'; ws_res['B5'].font = RED
    ws_res.cell(6, 1, 'Insumos').font = DATA
    ws_res['B6'] = f'=ROMANEO!F{R_CARNE}*PARAMETROS!B8'; ws_res['B6'].number_format = '$#,##0'; ws_res['B6'].font = RED
    ws_res.cell(7, 1, 'Flete').font = DATA
    ws_res['B7'] = f'=ROMANEO!F{R_CARNE}*PARAMETROS!B9'; ws_res['B7'].number_format = '$#,##0'; ws_res['B7'].font = RED
    ws_res.cell(8, 1, 'SENASA + cuarteo').font = DATA
    ws_res['B8'] = f'=ROMANEO!F{R_CARNE}*PARAMETROS!B10'; ws_res['B8'].number_format = '$#,##0'; ws_res['B8'].font = RED

    ws_res.cell(10, 1, 'COSTO TOTAL (sin IIBB)').font = Font(
        name='Arial', bold=True, size=11, color='C62828')
    ws_res.cell(10, 1).fill = PatternFill('solid', fgColor='FCE4EC')
    ws_res['B10'] = '=SUM(B4:B8)'; ws_res['B10'].number_format = '$#,##0'
    ws_res['B10'].font = Font(name='Arial', bold=True, size=11, color='C62828')
    ws_res['B10'].fill = PatternFill('solid', fgColor='FCE4EC')

    # INGRESOS POR LÍNEA
    ws_res['A13'] = 'INGRESOS POR LÍNEA DE CORTE'
    ws_res['A13'].font = Font(name='Arial', bold=True, size=12, color='2E7D32')
    ws_res['A13'].fill = PatternFill('solid', fgColor='E2EFDA')
    for c_col in ['B13', 'C13', 'D13', 'E13']:
        ws_res[c_col].fill = PatternFill('solid', fgColor='E2EFDA')

    hdr_row(ws_res, 14, ['Corte', 'Kg', 'Cliente', 'Precio $/kg', 'Ingreso $'],
            PatternFill('solid', fgColor='2E7D32'))

    ING_FIRST = 15
    meat_only = [c for c in cortes if c.get('grupo') != 'GRASA']
    ING_LAST = ING_FIRST + len(meat_only) - 1

    grupo_to_prec_row = {g: PREC_FIRST + i for i, (g, _) in enumerate(GRUPOS)}
    cli_to_col_idx = {cli: 4 + j for j, cli in enumerate(cli_list)}

    for i, c in enumerate(meat_only):
        r = ING_FIRST + i
        ws_res.cell(r, 1, c['corte']).font = DATA
        ws_res.cell(r, 2, c['kg']).font = DATA; ws_res[f'B{r}'].number_format = '#,##0.00'
        cli = c.get('cliente', 'SIN ASIGNAR')
        ws_res.cell(r, 3, cli).font = DATA
        prec_row = grupo_to_prec_row.get(c['grupo'])
        col_idx = cli_to_col_idx.get(cli)
        if prec_row and col_idx:
            col_letter = get_column_letter(col_idx)
            ws_res[f'D{r}'] = f'=PRECIOS!{col_letter}{prec_row}'
            ws_res[f'D{r}'].number_format = '$#,##0'; ws_res[f'D{r}'].font = DATA
        else:
            ws_res.cell(r, 4, 0).font = RED
        ws_res[f'E{r}'] = f'=B{r}*D{r}'
        ws_res[f'E{r}'].number_format = '$#,##0'; ws_res[f'E{r}'].font = GREEN
        cm = str(c.get('contramarca', ''))
        if cm in AMARILLA_CONTRAMARCAS:
            for col in range(1, 6):
                ws_res.cell(r, col).fill = AMARILLA_FILL

    R_TOT_ING = ING_LAST + 1
    ws_res.cell(R_TOT_ING, 1, 'TOTAL INGRESOS BRUTOS').font = TOTAL_FONT
    for col in range(1, 6):
        ws_res.cell(R_TOT_ING, col).fill = TOTAL_FILL
    ws_res[f'B{R_TOT_ING}'] = f'=SUM(B{ING_FIRST}:B{ING_LAST})'
    ws_res[f'B{R_TOT_ING}'].number_format = '#,##0.00'; ws_res[f'B{R_TOT_ING}'].font = TOTAL_FONT
    ws_res[f'E{R_TOT_ING}'] = f'=SUM(E{ING_FIRST}:E{ING_LAST})'
    ws_res[f'E{R_TOT_ING}'].number_format = '$#,##0'; ws_res[f'E{R_TOT_ING}'].font = TOTAL_FONT

    # P&L
    R_PL = R_TOT_ING + 2
    ws_res[f'A{R_PL}'] = 'RESULTADO DEL NEGOCIO'
    ws_res[f'A{R_PL}'].font = Font(name='Arial', bold=True, size=13, color='1F4E79')
    ws_res[f'A{R_PL}'].fill = SEC_FILL; ws_res[f'B{R_PL}'].fill = SEC_FILL

    r = R_PL + 1; R_ING_BRUTO = r
    ws_res.cell(r, 1, 'Ingresos brutos').font = DATA
    ws_res[f'B{r}'] = f'=E{R_TOT_ING}'; ws_res[f'B{r}'].number_format = '$#,##0'; ws_res[f'B{r}'].font = GREEN

    r += 1
    ws_res.cell(r, 1, 'IIBB y otros impuestos (sobre venta)').font = DATA
    ws_res[f'B{r}'] = f'=-B{R_ING_BRUTO}*PARAMETROS!B11'
    ws_res[f'B{r}'].number_format = '$#,##0'; ws_res[f'B{r}'].font = RED

    r += 1; R_ING_NETO = r
    ws_res.cell(r, 1, 'INGRESOS NETOS').font = Font(name='Arial', bold=True, size=11)
    ws_res[f'B{r}'] = f'=B{R_ING_BRUTO}+B{r-1}'
    ws_res[f'B{r}'].number_format = '$#,##0'
    ws_res[f'B{r}'].font = Font(name='Arial', bold=True, size=11, color='008000')

    r += 1; R_COSTO_NEG = r
    ws_res.cell(r, 1, 'Costo total').font = DATA
    ws_res[f'B{r}'] = '=-B10'; ws_res[f'B{r}'].number_format = '$#,##0'; ws_res[f'B{r}'].font = RED

    r += 2; R_CM = r
    ws_res.cell(r, 1, 'CONTRIBUCIÓN MARGINAL').font = Font(
        name='Arial', bold=True, size=14, color='1F4E79')
    ws_res.cell(r, 1).fill = SEC_FILL
    ws_res[f'B{r}'] = f'=B{R_ING_NETO}+B{R_COSTO_NEG}'
    ws_res[f'B{r}'].number_format = '$#,##0'
    ws_res[f'B{r}'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
    ws_res[f'B{r}'].fill = SEC_FILL

    r += 1; R_MARGEN = r
    ws_res.cell(r, 1, 'Margen % (CM / Ingresos netos)').font = Font(
        name='Arial', bold=True, size=10)
    ws_res[f'B{r}'] = f'=IF(B{R_ING_NETO}=0,0,B{R_CM}/B{R_ING_NETO})'
    ws_res[f'B{r}'].number_format = '0.0%'
    ws_res[f'B{r}'].font = Font(name='Arial', bold=True, size=12, color='1F4E79')

    r += 1
    ws_res.cell(r, 1, 'CM por kg de carne producida').font = Font(
        name='Arial', bold=True, size=10)
    ws_res[f'B{r}'] = f'=IF(ROMANEO!F{R_CARNE}=0,0,B{R_CM}/ROMANEO!F{R_CARNE})'
    ws_res[f'B{r}'].number_format = '$#,##0'

    r += 1
    ws_res.cell(r, 1, 'CM por kg de entrada (hacienda)').font = Font(
        name='Arial', bold=True, size=10)
    ws_res[f'B{r}'] = f'=IF(ROMANEO!B6=0,0,B{R_CM}/ROMANEO!B6)'
    ws_res[f'B{r}'].number_format = '$#,##0'

    r += 1
    ws_res.cell(r, 1, 'Precio venta promedio $/kg carne').font = Font(
        name='Arial', bold=True, size=10)
    ws_res[f'B{r}'] = f'=IF(ROMANEO!F{R_CARNE}=0,0,B{R_ING_BRUTO}/ROMANEO!F{R_CARNE})'
    ws_res[f'B{r}'].number_format = '$#,##0'

    # Calificación
    r += 2
    ws_res[f'A{r}'] = 'CALIFICACIÓN GENERAL'
    ws_res[f'A{r}'].font = Font(name='Arial', bold=True, size=13, color='1F4E79')
    ws_res[f'A{r}'].fill = SEC_FILL; ws_res[f'B{r}'].fill = SEC_FILL
    r += 1
    ws_res.cell(r, 1, 'RESULTADO').font = Font(name='Arial', bold=True, size=14)
    ws_res[f'B{r}'] = (
        f'=IF(B{R_CM}<=0,"PÉRDIDA",'
        f'IF(B{R_MARGEN}>=0.15,"ÓPTIMO",'
        f'IF(B{R_MARGEN}>=0.08,"BUENO",'
        f'IF(B{R_MARGEN}>=0.03,"REGULAR","MALO"))))'
    )
    ws_res[f'B{r}'].font = Font(name='Arial', bold=True, size=16)
    ws_res[f'B{r}'].alignment = Alignment(horizontal='center')
    rng = f'B{r}:B{r}'
    ws_res.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"ÓPTIMO"'], fill=CF_GREEN))
    ws_res.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"BUENO"'], fill=CF_BLUE))
    ws_res.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"REGULAR"'], fill=CF_YELLOW))
    ws_res.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"MALO"'], fill=CF_RED))
    ws_res.conditional_formatting.add(rng, CellIsRule(
        operator='equal', formula=['"PÉRDIDA"'], fill=PatternFill('solid', fgColor='FF0000')))

    wb.save(output_path)
    return {
        'R_CARNE': R_CARNE, 'R_CM': R_CM, 'R_MARGEN': R_MARGEN,
        'R_ING_NETO': R_ING_NETO, 'R_TOT_ING': R_TOT_ING,
        'AN_FIRST': AN_FIRST, 'AN_LAST': AN_LAST,
        'R_ING_BRUTO': R_ING_BRUTO, 'rend_obj': rend_obj,
        'calidad': calidad,
    }
