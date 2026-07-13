"""
insumos.py — Costo de insumos (bolsas, etiquetas, cajas) TEÓRICO por pieza
(desde el romaneo) vs REAL (consumo LIFO desde la planilla de stock+compras).

Reglas de mapeo corte → bolsa (confirmadas):
  - Porcionado / Feteado (~1 kg)  → bolsa CHICA
  - Anatómico                     → bolsa GRANDE
  - Con hueso (asado, osobuco, tomahawk, prime ribs, T-bone) → bolsa HUESO (mayor micronaje)
  - Asado plancha grande (14 kg)  → SIN bolsa
  - + 1 etiqueta alto impacto por pieza (todos)
  - + etiqueta autoadhesiva solo en cortes de ese nombre (por ahora $0)
  - + caja: valor (tapa+fondo) / 16 kg → $/kg empaquetado
"""

# Precios por categoría (LIFO, últimas compras — editables en la app)
PRECIOS = {
    'bolsa_chica': 205,   # 200x300 / 200x400
    'bolsa_grande': 390,  # 300x400 / 350x500 / 350x600
    'bolsa_hueso': 690,   # P/HUESO
    'etiqueta': 15,       # alto impacto 10x10
    'autoadhesiva': 0,    # RAAL — sin precio cargado por ahora
    'caja_kg': 112,       # (tapa+fondo)/16
}

# Cortes con hueso → bolsa de mayor micronaje
HUESO_KEYS = ['OSOBUCO', 'TOMAHAWK', 'PRIME RIB', 'T-BONE', 'T BONE', 'BALERO',
              'ASADO C/H', 'ASADO CON HUESO', '5 COSTILLAS', '7 COSTILLAS']
# Cortes con etiqueta autoadhesiva propia
AUTOADHESIVA_KEYS = ['T-BONE', 'MEDIALUNA', 'TOMAHAWK', 'CEJA DE OJO', 'CEJA DE BIFE',
                     'VACIO ROJO', 'OJO DE BIFE', 'BIFE DE CHORIZO', 'ASADO', 'PRIME RIB',
                     'OSOBUCO', 'DENVER']


def _es_hueso(desc, grupo):
    d = (desc or '').upper()
    g = (grupo or '').upper()
    if 'S/H' in d or 'SIN HUESO' in d:
        return False
    # Todo lo ASADO (del centro, banderita, tiras, parent) va con hueso → bolsa
    # de mayor micronaje. Excepción: plancha (va sin bolsa, se filtra antes).
    if 'ASADO' in g and 'PLANCHA' not in g:
        return True
    if any(k in d for k in HUESO_KEYS):
        return True
    return False


def categoria_bolsa(corte):
    """Devuelve 'chica' | 'grande' | 'hueso' | None (sin bolsa)."""
    desc = (corte.get('corte', '') or '').upper()
    grupo = corte.get('grupo', '')
    tipo = corte.get('tipo', 'ANATÓMICO')
    if 'PLANCHA' in desc:               # asado plancha grande → sin bolsa
        return None
    if _es_hueso(desc, grupo):
        return 'hueso'
    if tipo in ('PORCIONADO', 'FETEADO'):
        return 'chica'
    return 'grande'


def _usa_autoadhesiva(corte):
    d = (corte.get('corte', '') or '').upper()
    return any(k in d for k in AUTOADHESIVA_KEYS)


def costo_teorico(parsed_list, precios=None):
    """Insumo teórico ('lo que debería gastarse') desde los romaneos.
    Cuenta piezas por corte (sin recorte ni grasa)."""
    p = dict(PRECIOS, **(precios or {}))
    RECORTES = ('RECORTE 70-30', 'RECORTE 80-20', 'RECORTE 90-10', 'GRASA', 'SIN CLASIFICAR')

    tot = {'bolsas': 0.0, 'etiquetas': 0.0, 'autoadhesivas': 0.0, 'caja': 0.0,
           'piezas': 0, 'kg': 0.0, 'por_categoria': {'chica': 0, 'grande': 0, 'hueso': 0, 'sin_bolsa': 0}}

    for rom in parsed_list:
        for c in rom.get('cortes', []):
            if c.get('grupo') in RECORTES:
                continue
            piezas = int(c.get('piezas', 0) or 0)
            kg = c.get('kg', 0) or 0
            tot['piezas'] += piezas
            tot['kg'] += kg

            cat = categoria_bolsa(c)
            if cat is None:
                tot['por_categoria']['sin_bolsa'] += piezas
            else:
                tot['por_categoria'][cat] += piezas
                tot['bolsas'] += piezas * p[f'bolsa_{cat}']

            # 1 etiqueta alto impacto por pieza (todos)
            tot['etiquetas'] += piezas * p['etiqueta']
            # autoadhesiva solo en cortes nombrados
            if _usa_autoadhesiva(c):
                tot['autoadhesivas'] += piezas * p['autoadhesiva']

    tot['caja'] = tot['kg'] * p['caja_kg']
    tot['total'] = tot['bolsas'] + tot['etiquetas'] + tot['autoadhesivas'] + tot['caja']
    tot['costo_por_kg'] = tot['total'] / tot['kg'] if tot['kg'] else 0
    return tot
