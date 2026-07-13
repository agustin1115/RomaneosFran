"""
pricing.py — Motor de pricing basado en Costeo MADRE v2.

Soporta 3 modalidades de venta con costos y márgenes distintos:
- Media res entera (sin MO/insumos/faena/congelado, con merma adicional 2-3%)
- Cuartos (idem media, pero cuarteado: +3% merma por cuarteo)
- Cortes individuales (todos los costos, precio base = NETOS PEYA)
"""

# ══════════════════════════════════════════════════════════════════════
# FORMATEO ARGENTINO
# ══════════════════════════════════════════════════════════════════════
def fmt_num(v, decimales=0):
    """Número con separador de miles con punto, decimales con coma."""
    try:
        v = float(v)
    except (ValueError, TypeError):
        return str(v)
    decimales = min(max(decimales, 0), 2)
    if decimales == 0:
        s = f"{v:,.0f}"
        return s.replace(',', '.')
    s = f"{v:,.{decimales}f}"
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def fmt_pct(v, decimales=2):
    """Porcentaje con coma y hasta 2 decimales (ej: 10,54%)."""
    try:
        v = float(v)
    except (ValueError, TypeError):
        return str(v)
    decimales = min(max(decimales, 0), 2)
    s = f"{v:.{decimales}f}".replace('.', ',')
    return s + '%'


def fmt_dinero(v, decimales=0):
    """$1.000 o $1.234,50."""
    return '$' + fmt_num(v, decimales)


# ══════════════════════════════════════════════════════════════════════
# DATOS EXTRAÍDOS DEL COSTEO MADRE v2
# ══════════════════════════════════════════════════════════════════════

CUARTOS_PCT_ENTRADA = {
    'Pistola': 0.464,
    'Asado': 0.2691,
    'Pecho': 0.2438,
    'Merma': 0.0231,
}

REND_CATEGORIA = {
    'Vaca': 0.66,
    'Vaquillona': 0.67,
    'Novillito': 0.68,
    'Novillo': 0.68,
    'Toro': 0.66,
    'Bubalino': 0.67,
}

# Cortes agrupados por cuarto con su % de composición dentro del cuarto
# (precios Standard/Black/Búfalo se sobreescriben desde PRECIOS_BASE cuando
# se usa la lista "Standard" — para Black y Búfalo se mantienen los del costeo)
CORTES_POR_CUARTO = {
    'Bife': [
        ('LOMO CON CORDON',    0.1213, 20355, 28000, 21900),
        ('OJO DE BIFE',        0.2434, 18010, 25000, 16500),
        ('CEJA DE BIFE',       0.0,    28000, 33000, 26500),  # premium tipo entraña
        ('BIFE DE CHORIZO',    0.3934, 18143, 22500, 16000),
        ('TAPA DE BIFE',       0.1213, 11128, 13000,     0),
        ('RECORTE 70-30',      0.1206,  6500,  6000,  6500),
    ],
    'Asado': [
        ('ASADO',              0.6076, 15045, 19000, 11000),
        ('VACIO',              0.2617, 15399, 19500, 14600),
        ('RECORTE 70-30',      0.03,    6500,  6000,  6500),
        ('MATAMBRE',           0.0868, 12833, 16000, 11900),
        ('ENTRAÑA',            0.0226, 28763, 33000, 26500),
    ],
    'Pecho': [
        ('CARNE PICADA',       0.834,  11128, 12500,  9100),
        ('ROASTBEEF',          0.0629, 11178, 14000, 10200),
        ('PALETA',             0.0629, 11791, 14000, 10200),
        ('TAPA DE ASADO',      0.0273, 11278, 14000, 10500),
        ('FALDA',              0.0155, 11127, 15000,     0),
    ],
    'Mocho': [
        ('CUADRIL',            0.0341, 14337, 16500, 12500),
        ('COLITA',              0.014, 17523, 20500, 15900),
        ('TAPA DE CUADRIL',     0.015, 15488, 17500, 13500),
        ('PECETO',             0.0174, 15930, 20000, 15000),
        ('TORTUGUITA',          0.023, 11128, 13000,     0),
        ('TAPA DE NALGA',       0.024, 11788, 13500, 10000),
        ('NALGA SIN TAPA',     0.0548, 14308, 17000, 11950),
        ('CUADRADA',            0.047, 12980, 15000, 10900),
        ('BOLA DE LOMO',        0.047, 12982, 15000, 10900),
    ],
}


def calcular_kg_cuarto(kg_carne_total):
    """Distribuye kg carne entre los cuartos comerciales."""
    return {
        'Bife':  kg_carne_total * 0.312,
        'Asado': kg_carne_total * 0.294,
        'Pecho': kg_carne_total * 0.308,
        'Mocho': kg_carne_total * 0.086,
    }


# ══════════════════════════════════════════════════════════════════════
# COSTOS POR DEFECTO
# ══════════════════════════════════════════════════════════════════════
COSTOS_PRICING_DEFAULT = {
    'precio_compra_kg': 6900,
    'peso_media': 150,
    'rendimiento': 0.68,
    'faena_media': 2000,      # aplica a venta por cortes (cuarteo interno)
    'mo_kg': 1000,            # solo venta por cortes
    'insumos_kg': 0,          # solo venta por cortes
    'flete_kg': 120,          # aplica a todas
    'senasa_kg': 1.2,         # por kg egresado (era $/media, ahora $/kg)
    'congelado_kg': 0,        # solo venta por cortes
    'iibb_ganancias': 0.035,
    'imp_cheque': 0.012,
    'comision_1': 0,
    'comision_2': 0,
    'tna': 0.48,
    'dias_financiamiento': 25,
    # Márgenes independientes por modalidad
    'margen_media': 0.04,     # 4%
    'margen_cuartos': 0.05,   # 5%
    'margen_cortes': 0.10,    # 10%
    'merma_media': 0.025,     # 2.5% merma adicional al vender media entera
    'merma_cuarteo': 0.03,    # 3% merma al cortar en cuartos
    # Carne amarilla (se vende a parte a precio fijo)
    'pct_amarilla': 0.0,
    'precio_amarilla': 10500,
}


# ══════════════════════════════════════════════════════════════════════
# MOTOR DE CÁLCULO
# ══════════════════════════════════════════════════════════════════════
def _calcular_base(params, modo):
    """
    Calcula costos base según modalidad de venta.
    - 'media': se vende la media entera CON HUESO. kg_vendibles = kg_ingresados × (1-merma)
               Solo hacienda + SENASA + flete (sin MO/insumos/faena/congelado)
    - 'cuartos': se vende cuarteada CON HUESO. kg_vendibles = kg_ingresados × (1-merma_cuarteo)
                 Solo hacienda + SENASA + flete
    - 'cortes': se vende despostada. kg_vendibles = kg_ingresados × rendimiento
                Todos los costos (incluye MO, insumos, faena, congelado)
    """
    p = {**COSTOS_PRICING_DEFAULT, **(params or {})}
    kg_ingresados = p['peso_media']
    kg_despost = kg_ingresados * p['rendimiento']  # referencia para costos que van por kg carne

    # Merma y kg vendibles según modalidad
    if modo == 'media':
        # Media entera con hueso, solo pierde merma (oreo/manipuleo)
        kg_vendibles = kg_ingresados * (1 - p['merma_media'])
    elif modo == 'cuartos':
        # Cuarteada con hueso, pierde 3% por el cuarteo
        kg_vendibles = kg_ingresados * (1 - p['merma_cuarteo'])
    else:  # cortes
        # Despostada, rinde según categoría
        kg_vendibles = kg_despost

    costo_hacienda = kg_ingresados * p['precio_compra_kg']
    # SENASA: se paga por kg egresado real (lo que sale de planta)
    costo_senasa = kg_vendibles * p['senasa_kg']
    # Flete: sobre lo que se vende
    costo_flete = kg_vendibles * p['flete_kg']

    if modo == 'cortes':
        costo_mo = kg_vendibles * p['mo_kg']
        costo_insumos = kg_vendibles * p['insumos_kg']
        costo_congelado = kg_vendibles * p['congelado_kg']
        costo_faena = p['faena_media']  # cuarteo interno en planta
    else:
        costo_mo = 0
        costo_insumos = 0
        costo_congelado = 0
        costo_faena = 0

    costos_directos = (costo_hacienda + costo_senasa + costo_flete
                       + costo_mo + costo_insumos + costo_congelado + costo_faena)

    return {
        'params': p,
        'kg_ingresados': kg_ingresados,
        'kg_despost': kg_despost,
        'kg_vendibles': kg_vendibles,
        'costo_hacienda': costo_hacienda,
        'costo_senasa': costo_senasa,
        'costo_flete': costo_flete,
        'costo_mo': costo_mo,
        'costo_insumos': costo_insumos,
        'costo_congelado': costo_congelado,
        'costo_faena': costo_faena,
        'costos_directos': costos_directos,
    }


def _aplicar_impuestos_y_financiero(base, venta_estimada):
    """Suma impuestos sobre venta y costo financiero."""
    p = base['params']
    costo_iibb = venta_estimada * p['iibb_ganancias']
    costo_cheque = venta_estimada * p['imp_cheque']
    costo_com1 = venta_estimada * p['comision_1']
    costo_com2 = venta_estimada * p['comision_2']
    costo_financiero = base['costos_directos'] * p['tna'] * p['dias_financiamiento'] / 365

    costo_total = (base['costos_directos'] + costo_iibb + costo_cheque
                   + costo_com1 + costo_com2 + costo_financiero)

    return {
        **base,
        'venta_estimada': venta_estimada,
        'costo_iibb': costo_iibb,
        'costo_cheque': costo_cheque,
        'costo_com1': costo_com1,
        'costo_com2': costo_com2,
        'costo_financiero': costo_financiero,
        'costo_total': costo_total,
    }


def calcular_media_res(params):
    """Precio sugerido venta media res entera."""
    base = _calcular_base(params, 'media')
    margen = base['params']['margen_media']
    # Precio objetivo: costo_total / (1 - margen), pero costo_total depende de impuestos sobre venta.
    # Iteración simple (converge rápido): arrancar con venta = costos_directos/(1-margen)
    venta = base['costos_directos'] / (1 - margen) if margen < 1 else base['costos_directos']
    for _ in range(5):
        cr = _aplicar_impuestos_y_financiero(base, venta)
        venta = cr['costo_total'] / (1 - margen) if margen < 1 else cr['costo_total']
    cr = _aplicar_impuestos_y_financiero(base, venta)
    precio_kg = venta / base['kg_vendibles'] if base['kg_vendibles'] > 0 else 0
    resultado = venta - cr['costo_total']
    margen_real = resultado / venta * 100 if venta > 0 else 0
    return {
        'modo': 'Media res entera',
        'base': cr,
        'precio_kg': precio_kg,
        'venta_total': venta,
        'costo_total': cr['costo_total'],
        'resultado': resultado,
        'margen_pct': margen_real,
    }


def calcular_cuartos(params, lista='Standard', precios_custom=None, ajuste_pct=0):
    """
    Precio de cada cuarto para lograr el margen objetivo.
    precios_custom: dict con precios por corte (los que se usan en lugar del default)
    """
    base = _calcular_base(params, 'cuartos')
    margen = base['params']['margen_cuartos']

    # Calcular kg por cuarto
    kg_cuartos = calcular_kg_cuarto(base['kg_vendibles'])
    col_idx = {'Standard': 2, 'Black': 3, 'Bufalo': 4}.get(lista, 2)

    # Valor estimado de cada cuarto usando precios de la lista (+ ajuste)
    valor_por_cuarto = {}
    kg_por_cuarto = {}
    for cuarto, cortes in CORTES_POR_CUARTO.items():
        kg_c = kg_cuartos.get(cuarto, 0)
        valor = 0
        kg_total_c = 0
        for corte_info in cortes:
            nombre = corte_info[0]
            pct_comp = corte_info[1]
            precio_base = corte_info[col_idx]
            # Si hay precios custom, usar esos
            if precios_custom and nombre in precios_custom:
                precio_base = precios_custom[nombre]
            precio = precio_base * (1 + ajuste_pct)
            kg_corte = kg_c * pct_comp
            valor += kg_corte * precio
            kg_total_c += kg_corte
        valor_por_cuarto[cuarto] = valor
        kg_por_cuarto[cuarto] = kg_total_c

    venta_estimada = sum(valor_por_cuarto.values())
    # Iterar impuestos
    cr = _aplicar_impuestos_y_financiero(base, venta_estimada)
    for _ in range(3):
        cr = _aplicar_impuestos_y_financiero(base, venta_estimada)

    # Ajustar precio de cada cuarto para lograr margen (escalado uniforme)
    # venta_objetivo = costo_total / (1 - margen)
    venta_objetivo = cr['costo_total'] / (1 - margen) if margen < 1 else cr['costo_total']
    factor_ajuste = venta_objetivo / venta_estimada if venta_estimada > 0 else 1

    resultado = []
    total_valor_obj = 0
    total_kg = 0
    for cuarto in CORTES_POR_CUARTO:
        kg_c = kg_por_cuarto[cuarto]
        valor_obj = valor_por_cuarto[cuarto] * factor_ajuste
        precio_obj = valor_obj / kg_c if kg_c > 0 else 0
        resultado.append({
            'Cuarto': cuarto,
            'Kg': kg_c,
            'Valor estimado': valor_por_cuarto[cuarto],
            'Precio $/kg (lista)': valor_por_cuarto[cuarto] / kg_c if kg_c > 0 else 0,
            'Precio $/kg sugerido': precio_obj,
        })
        total_valor_obj += valor_obj
        total_kg += kg_c

    return {
        'modo': 'Cuartos',
        'base': cr,
        'cuartos': resultado,
        'factor_ajuste': factor_ajuste,
        'venta_total': total_valor_obj,
        'kg_total': total_kg,
        'costo_total': cr['costo_total'],
        'resultado': total_valor_obj - cr['costo_total'],
        'margen_pct': (total_valor_obj - cr['costo_total']) / total_valor_obj * 100
                      if total_valor_obj > 0 else 0,
        'venta_objetivo': venta_objetivo,
    }


def _precios_base_netos_peya(price_matrix=None):
    """Construye dict de precios NETOS PEYA por grupo.
    Si se pasa `price_matrix` (live de Google Sheets) se prioriza sobre el
    PRECIOS_BASE hardcodeado.
    Resuelve subcortes con _INHERIT_PARENT al precio del parent."""
    try:
        from config import PRECIOS_BASE, SUBCORTE_TO_PARENT
    except Exception:
        return {}
    fuente = price_matrix if price_matrix else PRECIOS_BASE
    out = {}
    for grupo, clientes in fuente.items():
        if not isinstance(clientes, dict):
            continue
        if 'NETOS PEYA' in clientes and clientes['NETOS PEYA'] > 0:
            out[grupo] = clientes['NETOS PEYA']
        elif 'RESTO CLIENTES AMBA' in clientes and clientes['RESTO CLIENTES AMBA'] > 0:
            out[grupo] = clientes['RESTO CLIENTES AMBA']
    # Resolver subcortes que heredan del parent
    if not price_matrix:
        for grupo, clientes in PRECIOS_BASE.items():
            if isinstance(clientes, dict) and clientes.get('_INHERIT_PARENT'):
                parent = SUBCORTE_TO_PARENT.get(grupo)
                if parent and parent in out:
                    out[grupo] = out[parent]
    else:
        # Live mode: si subcorte no aparece en la planilla, hereda del parent
        for sub, parent in SUBCORTE_TO_PARENT.items():
            if sub not in out and parent in out:
                out[sub] = out[parent]
    return out


def _grupos_por_lista(lista):
    """Devuelve [(nombre, pct_carne), ...] según rendimientos reales."""
    try:
        from config import GRUPOS_STANDARD, GRUPOS_BUFALO, GRUPOS_PREMIUM_BLACK
    except ImportError:
        return []
    return {
        'Standard': GRUPOS_STANDARD,
        'Black': GRUPOS_PREMIUM_BLACK,
        'Bufalo': GRUPOS_BUFALO,
    }.get(lista, GRUPOS_STANDARD)


def rendimiento_real_de_romaneos(parsed_files, mes=None, año=None, calidad=None):
    """
    Agrega cortes de varios romaneos parseados y devuelve [(grupo, % s/carne), ...].
    Filtra por mes/año (formato fecha 'DD/MM/YYYY') y opcionalmente por calidad.
    Excluye GRASA / DECOMISO. Mapea cada corte a su grupo de precio (CORTE_TO_GRUPO).
    Devuelve None si no hay datos suficientes.
    """
    if not parsed_files:
        return None
    try:
        from config import CORTE_TO_GRUPO
    except ImportError:
        CORTE_TO_GRUPO = []
    from collections import defaultdict
    from datetime import datetime as _dt

    def _grupo_de(nombre):
        n = (nombre or '').upper()
        for keys, grp in CORTE_TO_GRUPO:
            for k in keys:
                if k in n:
                    return grp
        return None

    agg = defaultdict(float)
    total_kg = 0.0
    for p in parsed_files:
        if p.get('error'):
            continue
        if calidad and p.get('calidad') != calidad:
            continue
        if mes is not None or año is not None:
            try:
                fdt = _dt.strptime(p.get('fecha', ''), '%d/%m/%Y')
            except Exception:
                continue
            if mes is not None and fdt.month != mes:
                continue
            if año is not None and fdt.year != año:
                continue
        for c in p.get('cortes', []) or []:
            grupo_pre = (c.get('grupo') or '').upper()
            kg = c.get('kg', 0) or 0
            if kg <= 0:
                continue
            if grupo_pre in {'GRASA', 'DECOMISO', 'SUBPRODUCTO'} or 'GRASA' in grupo_pre:
                continue
            grp = _grupo_de(c.get('corte', ''))
            if not grp or grp == 'GRASA':
                continue
            agg[grp] += kg
            total_kg += kg

    if total_kg <= 0:
        return None
    return [(g, kg / total_kg) for g, kg in sorted(agg.items(), key=lambda x: -x[1])]


def _corte_a_cuarto():
    """Mapeo nombre_corte → cuarto (para agrupar visualmente).
    Los subcortes heredan el cuarto del parent."""
    out = {}
    for cuarto, cortes in CORTES_POR_CUARTO.items():
        for ci in cortes:
            out[ci[0]] = cuarto
    try:
        from config import SUBCORTE_TO_PARENT
        for sub, parent in SUBCORTE_TO_PARENT.items():
            if parent in out and sub not in out:
                out[sub] = out[parent]
    except Exception:
        pass
    out['RECORTE 80-20'] = 'Recortes'
    out['RECORTE 90-10'] = 'Recortes'
    out['RECORTE 70-30'] = 'Recortes'
    return out


def _precio_corte(nombre, lista, precios_custom, precios_netos):
    """Resuelve el precio base de un corte aplicando prioridades:
    1) precio custom del usuario
    2) NETOS PEYA (Standard) — ya resuelve _INHERIT_PARENT al precio del padre
    3) CORTES_POR_CUARTO según lista
    4) Si es subcorte, hereda del parent
    5) RECORTE → $6.500 fallback
    """
    if precios_custom and nombre in precios_custom:
        return precios_custom[nombre]
    if lista == 'Standard' and nombre in precios_netos:
        return precios_netos[nombre]
    col_idx = {'Standard': 2, 'Black': 3, 'Bufalo': 4}.get(lista, 2)
    for _cuarto, cortes in CORTES_POR_CUARTO.items():
        for ci in cortes:
            if ci[0] == nombre and ci[col_idx]:
                return ci[col_idx]
    # Subcorte sin precio propio → hereda del parent
    try:
        from config import SUBCORTE_TO_PARENT
        parent = SUBCORTE_TO_PARENT.get(nombre)
        if parent:
            if precios_custom and parent in precios_custom:
                return precios_custom[parent]
            if lista == 'Standard' and parent in precios_netos:
                return precios_netos[parent]
            for _cuarto, cortes in CORTES_POR_CUARTO.items():
                for ci in cortes:
                    if ci[0] == parent and ci[col_idx]:
                        return ci[col_idx]
    except Exception:
        pass
    if 'RECORTE' in nombre.upper():
        return 6500
    return 0


def construir_lista_cortes(lista='Standard', precios_custom=None, ajuste_pct=0,
                            grupos_override=None, price_matrix=None):
    """
    Lista de cortes basada en rendimientos REALES (GRUPOS_STANDARD/BUFALO/BLACK).
    Si lista='Standard' usa NETOS PEYA de PRECIOS_BASE (o de price_matrix si
    se pasa la planilla live).
    precios_custom: dict {nombre: precio} para sobreescribir.
    grupos_override: lista [(nombre, pct), ...] que reemplaza los rendimientos
        hardcodeados (ej. salida de rendimiento_real_de_romaneos).
    price_matrix: dict {grupo: {cliente: precio}} live de Google Sheets.
    """
    precios_netos = _precios_base_netos_peya(price_matrix) if lista == 'Standard' else {}
    grupos = grupos_override if grupos_override else _grupos_por_lista(lista)
    cuarto_map = _corte_a_cuarto()
    rows = []
    for nombre, pct_carne in grupos:
        precio_base = _precio_corte(nombre, lista, precios_custom, precios_netos)
        precio_ajustado = precio_base * (1 + ajuste_pct)
        rows.append({
            'Cuarto': cuarto_map.get(nombre, 'Otros'),
            'Corte': nombre,
            'pct_cuarto': pct_carne,
            'Precio base': precio_base,
            'Precio ajustado': precio_ajustado,
        })
    return rows


def calcular_cortes(params, lista='Standard', precios_custom=None, ajuste_pct=0,
                    grupos_override=None, price_matrix=None):
    """
    Calcula resultado vendiendo por cortes individuales usando rendimientos reales
    (GRUPOS_STANDARD/BUFALO/BLACK), e incluye carne amarilla como categoría aparte.
    grupos_override: lista [(nombre, pct), ...] para usar rendimiento dinámico.
    price_matrix: dict live de Google Sheets para precios actualizados.
    """
    base = _calcular_base(params, 'cortes')
    p = base['params']
    margen = p['margen_cortes']
    kg_vendibles = base['kg_vendibles']

    # Carne amarilla — se separa antes de distribuir el resto entre cortes
    pct_amarilla = max(0.0, min(p.get('pct_amarilla', 0) or 0, 1.0))
    precio_amarilla = p.get('precio_amarilla', 10500) or 0
    kg_amarilla = kg_vendibles * pct_amarilla
    kg_para_cortes = kg_vendibles - kg_amarilla

    precios_netos = _precios_base_netos_peya(price_matrix) if lista == 'Standard' else {}
    grupos = grupos_override if grupos_override else _grupos_por_lista(lista)
    cuarto_map = _corte_a_cuarto()

    cortes_detalle = []
    venta_total = 0.0

    for nombre, pct_carne in grupos:
        if pct_carne <= 0:
            continue
        kg_corte = kg_para_cortes * pct_carne
        precio_base = _precio_corte(nombre, lista, precios_custom, precios_netos)
        precio = precio_base * (1 + ajuste_pct)
        valor = kg_corte * precio
        cortes_detalle.append({
            'Cuarto': cuarto_map.get(nombre, 'Otros'),
            'Corte': nombre,
            'Kg': kg_corte,
            'Precio $/kg': precio,
            'Valor': valor,
        })
        venta_total += valor

    if kg_amarilla > 0:
        valor_am = kg_amarilla * precio_amarilla
        cortes_detalle.append({
            'Cuarto': 'Amarilla',
            'Corte': 'CARNE AMARILLA',
            'Kg': kg_amarilla,
            'Precio $/kg': precio_amarilla,
            'Valor': valor_am,
        })
        venta_total += valor_am

    cr = _aplicar_impuestos_y_financiero(base, venta_total)

    total_kg_carne = sum(c['Kg'] for c in cortes_detalle)
    for c in cortes_detalle:
        c['% del kg total'] = c['Kg'] / total_kg_carne * 100 if total_kg_carne > 0 else 0
        c['% del valor total'] = c['Valor'] / venta_total * 100 if venta_total > 0 else 0
        c['Impacto +1% precio'] = c['Valor'] * 0.01 / venta_total * 100 if venta_total > 0 else 0

    precio_prom = venta_total / total_kg_carne if total_kg_carne > 0 else 0
    resultado = venta_total - cr['costo_total']
    margen_real = resultado / venta_total * 100 if venta_total > 0 else 0

    venta_objetivo = cr['costo_total'] / (1 - margen) if margen < 1 else cr['costo_total']
    factor_necesario = venta_objetivo / venta_total if venta_total > 0 else 1
    ajuste_necesario = (factor_necesario - 1) * 100

    return {
        'modo': 'Cortes individuales',
        'base': cr,
        'cortes': cortes_detalle,
        'kg_total': total_kg_carne,
        'kg_amarilla': kg_amarilla,
        'venta_total': venta_total,
        'precio_prom': precio_prom,
        'costo_total': cr['costo_total'],
        'resultado': resultado,
        'margen_pct': margen_real,
        'margen_objetivo_pct': margen * 100,
        'ajuste_necesario_pct': ajuste_necesario,
        'venta_objetivo': venta_objetivo,
    }


def _norm_corte(s):
    """Normaliza nombre de corte para match (mayúsculas, sin espacios extras)."""
    return ' '.join(str(s).upper().split()) if s else ''


def aplicar_precios_a_romaneo(cortes_romaneo, precios_lista, params):
    """
    Aplica los precios de la lista (precios_lista: {nombre_corte: precio}) a los
    kg reales de los cortes de un romaneo, usando los costos de `params`.

    Devuelve dict con venta total, precio promedio real, resultado, margen,
    cortes con/sin match en la lista (para detectar gap).
    """
    p = {**COSTOS_PRICING_DEFAULT, **(params or {})}

    precios_norm = {_norm_corte(k): v for k, v in (precios_lista or {}).items()
                    if v and v > 0}

    cortes_match = []
    cortes_sin_match = []
    venta_total = 0.0
    kg_total_carne = 0.0
    kg_match = 0.0

    for c in cortes_romaneo or []:
        nombre = c.get('corte', '') if isinstance(c, dict) else ''
        grupo = (c.get('grupo', '') if isinstance(c, dict) else '').upper()
        kg = c.get('kg', 0) if isinstance(c, dict) else 0
        if kg <= 0:
            continue
        # Excluir grasa/decomiso (no se vende a precio de carne)
        if grupo in {'GRASA', 'DECOMISO', 'SUBPRODUCTO'} or 'GRASA' in nombre.upper():
            continue
        kg_total_carne += kg

        precio = precios_norm.get(_norm_corte(nombre), 0)
        if precio > 0:
            valor = kg * precio
            cortes_match.append({
                'corte': nombre, 'kg': kg, 'precio': precio,
                'valor': valor, 'grupo': grupo,
            })
            venta_total += valor
            kg_match += kg
        else:
            cortes_sin_match.append({'corte': nombre, 'kg': kg, 'grupo': grupo})

    # Costos: base 'cortes', pero con kg_vendibles igual al kg real del romaneo
    base = _calcular_base(p, 'cortes')
    base['kg_vendibles'] = kg_total_carne
    base['kg_despost'] = kg_total_carne
    base['costo_senasa'] = kg_total_carne * p['senasa_kg']
    base['costo_flete'] = kg_total_carne * p['flete_kg']
    base['costo_mo'] = kg_total_carne * p['mo_kg']
    base['costo_insumos'] = kg_total_carne * p['insumos_kg']
    base['costo_congelado'] = kg_total_carne * p['congelado_kg']
    base['costos_directos'] = (
        base['costo_hacienda'] + base['costo_senasa'] + base['costo_flete']
        + base['costo_mo'] + base['costo_insumos'] + base['costo_congelado']
        + base['costo_faena']
    )
    cr = _aplicar_impuestos_y_financiero(base, venta_total)
    resultado = venta_total - cr['costo_total']
    margen_real = (resultado / venta_total * 100) if venta_total > 0 else 0
    precio_prom = (venta_total / kg_match) if kg_match > 0 else 0
    precio_prom_total = (venta_total / kg_total_carne) if kg_total_carne > 0 else 0
    cobertura_pct = (kg_match / kg_total_carne * 100) if kg_total_carne > 0 else 0
    kg_sin = sum(c['kg'] for c in cortes_sin_match)

    return {
        'cortes_match': cortes_match,
        'cortes_sin_match': cortes_sin_match,
        'venta_total': venta_total,
        'kg_total_carne': kg_total_carne,
        'kg_match': kg_match,
        'kg_sin_match': kg_sin,
        'cobertura_pct': cobertura_pct,
        'precio_prom': precio_prom,             # solo sobre los matcheados
        'precio_prom_total': precio_prom_total, # diluído sobre todos los kg
        'costo_total': cr['costo_total'],
        'resultado': resultado,
        'margen_pct': margen_real,
        'base': cr,
    }
