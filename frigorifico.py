"""
frigorifico.py — Costos del frigorífico: PROYECTADO (del romaneo) vs REAL
(cuenta corriente). Servicios: VEP, Cuarteo, Despostada (MO), Congelado.

Tarifas confirmadas (abril 2026, editables en la app):
  VEP        $12.002 / cabeza
  Cuarteo    $2.500  / cabeza
  Despostada $790    / kg carne (kg salidos, sin grasa/recorte va incluido en carne)
  Congelado  $100/kg túnel + $30/kg por mes de stock  (solo cortes destino CONGELADO)
Cabezas = medias / 2.
"""

TARIFAS = {
    'vep': 12002,          # $/cabeza
    'cuarteo': 2500,       # $/cabeza
    'despostada': 790,     # $/kg carne
    'congelado_tunel': 100,        # $/kg (una vez)
    'congelado_stock_mes': 30,     # $/kg por mes en stock
}


def _kg_carne(p):
    return sum(c.get('kg', 0) for c in p.get('cortes', []) if c.get('grupo') != 'GRASA')


def _kg_congelado(p):
    """Kg de cortes con destino CONGELADO / CONG (excluye grasa)."""
    tot = 0
    for c in p.get('cortes', []):
        if c.get('grupo') == 'GRASA':
            continue
        d = str(c.get('destino', '')).upper()
        if 'CONG' in d:
            tot += c.get('kg', 0)
    return tot


def costo_proyectado(p, tarifas=None, meses_stock=1, cabezas=None):
    """Costo de frigorífico proyectado para UN romaneo, desde sus drivers físicos."""
    t = dict(TARIFAS, **(tarifas or {}))
    medias = p.get('medias_reses', 0) or 0
    if cabezas is None:
        cabezas = medias / 2
    kg_carne = _kg_carne(p)
    kg_cong = _kg_congelado(p)

    vep = cabezas * t['vep']
    cuarteo = medias * t['cuarteo']          # el cuarteo se hace a TODAS las medias
    despostada = kg_carne * t['despostada']
    congelado = kg_cong * (t['congelado_tunel'] + t['congelado_stock_mes'] * meses_stock)
    total = vep + cuarteo + despostada + congelado

    return {
        'cabezas': cabezas, 'medias': medias, 'kg_carne': kg_carne, 'kg_congelado': kg_cong,
        'vep': vep, 'cuarteo': cuarteo, 'despostada': despostada, 'congelado': congelado,
        'total': total,
        'costo_por_kg_carne': total / kg_carne if kg_carne else 0,
    }


def desvio(proyectado, real):
    """real y proyectado son dicts {servicio: monto}. Devuelve desvío por servicio."""
    servicios = ['vep', 'cuarteo', 'despostada', 'congelado']
    out = {}
    for s in servicios:
        pr = proyectado.get(s, 0)
        rl = real.get(s, 0)
        out[s] = {'proyectado': pr, 'real': rl, 'desvio': rl - pr,
                  'desvio_pct': (rl - pr) / pr * 100 if pr else 0}
    pt = sum(proyectado.get(s, 0) for s in servicios)
    rt = sum(real.get(s, 0) for s in servicios)
    out['total'] = {'proyectado': pt, 'real': rt, 'desvio': rt - pt,
                    'desvio_pct': (rt - pt) / pt * 100 if pt else 0}
    return out
