"""
segmentacion.py — Identidad por contenido + canal/segmento + persistencia.

Resuelve el problema de que todos los PDF se llaman igual: cada romaneo se
identifica por su CONTENIDO (no por el nombre), se le sugiere un canal/segmento
de negocio, y la elección del usuario queda guardada en un registro JSON.

Phase 1: identidad estable, sugerencia confiable (Búfalo vs Consumo) y
persistencia. Las reglas finas de China 6/23 y Hilton se calibran en Phase 2
con romaneos reales de exportación.
"""
import os
import json
import hashlib

# Canales/segmentos de negocio (orden de aparición en los dropdowns)
SEGMENTOS = ['Consumo', 'Vaca China 6', 'Vaca China 23', 'Novillo Hilton', 'Black', 'Búfalo',
             'Producción Especial']


def cabezas_por_categoria(parsed_list):
    """Suma cabezas / medias / kg por categoría (Vaca, Novillo, Vaquillona, Toro,
    Novillito, Bubalino) en el acumulado. Útil para trackear, p.ej., los toros
    que se meten para hacer más carne picada."""
    agg = {}
    for p in parsed_list:
        for cat, d in (p.get('desglose_categoria') or {}).items():
            a = agg.setdefault(cat, {'cabezas': 0.0, 'medias': 0.0, 'kg': 0.0})
            a['cabezas'] += d.get('cabezas', 0)
            a['medias'] += d.get('medias', 0)
            a['kg'] += d.get('kg', 0)
    tot_kg = sum(a['kg'] for a in agg.values()) or 1
    for a in agg.values():
        a['pct_kg'] = a['kg'] / tot_kg * 100
    # ordenar por kg desc
    return dict(sorted(agg.items(), key=lambda x: -x[1]['kg']))


def _kg_carne(p):
    return sum(c.get('kg', 0) for c in p.get('cortes', []) if c.get('grupo') != 'GRASA')


def _fecha_dt(p):
    """Devuelve datetime de la fecha del romaneo (dd/mm/yyyy) o None."""
    import datetime
    f = (p.get('fecha') or '').strip()
    for fmt in ('%d/%m/%Y', '%d/%m/%y'):
        try:
            return datetime.datetime.strptime(f, fmt)
        except Exception:
            pass
    return None


def semana_de(p):
    """Etiqueta de semana ISO del romaneo, p.ej. '2026-S34 (17–23 ago)'.
    Sirve para agrupar/elegir por semana. Devuelve '(sin fecha)' si no parsea."""
    import datetime
    dt = _fecha_dt(p)
    if not dt:
        return '(sin fecha)'
    iso = dt.isocalendar()  # (year, week, weekday)
    lunes = dt - datetime.timedelta(days=dt.weekday())
    domingo = lunes + datetime.timedelta(days=6)
    meses = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic']
    rango = f"{lunes.day} {meses[lunes.month-1]}–{domingo.day} {meses[domingo.month-1]}"
    return f"{iso[0]}-S{iso[1]:02d} ({rango})"


def semanas_disponibles(parsed_list):
    """Lista ordenada (más reciente primero) de semanas presentes en los romaneos."""
    vistas = {}
    for p in parsed_list:
        et = semana_de(p)
        dt = _fecha_dt(p)
        # clave de orden: usa la fecha mínima de la semana
        vistas.setdefault(et, dt or __import__('datetime').datetime.min)
    return [et for et, _ in sorted(vistas.items(), key=lambda x: x[1], reverse=True)]


def rendimiento_por_planta(parsed_list):
    """Agrupa entrada / carne / rinde por planta de desposte (ICO vs Top Meat).
    Permite comparar el rendimiento de cada planta por separado."""
    agg = {}
    for p in parsed_list:
        planta = p.get('planta') or 'Otra'
        a = agg.setdefault(planta, {'romaneos': 0, 'medias': 0.0, 'cabezas': 0.0,
                                    'kg_entrada': 0.0, 'kg_carne': 0.0, 'grasa_kg': 0.0,
                                    'merma_kg': 0.0})
        a['romaneos'] += 1
        a['medias'] += p.get('medias_reses', 0) or 0
        a['cabezas'] += (p.get('medias_reses', 0) or 0) / 2
        a['kg_entrada'] += p.get('kg_entrada', 0) or 0
        a['kg_carne'] += _kg_carne(p)
        a['grasa_kg'] += p.get('grasa_kg', 0) or 0
        a['merma_kg'] += p.get('merma_kg', 0) or 0
    for a in agg.values():
        ent = a['kg_entrada'] or 1
        a['rinde_carne_pct'] = a['kg_carne'] / ent * 100
        a['grasa_pct'] = a['grasa_kg'] / ent * 100
        a['merma_pct'] = a['merma_kg'] / ent * 100
    return dict(sorted(agg.items(), key=lambda x: -x[1]['kg_entrada']))


def grupos_distintos(p):
    """Cortes distintos (sin grasa ni sin-clasificar) — clave para distinguir
    un romaneo de consumo (30-45 cortes) de uno de China (6 o ~23)."""
    return sorted({c.get('grupo') for c in p.get('cortes', [])
                   if c.get('grupo') not in ('GRASA', 'SIN CLASIFICAR')})


def destino_principal(p):
    """Cliente dominante por kg — ayuda a distinguir consumo (PEYA/super) de export."""
    agg = {}
    for c in p.get('cortes', []):
        if c.get('grupo') == 'GRASA':
            continue
        cli = c.get('cliente') or 'SIN ASIGNAR'
        agg[cli] = agg.get(cli, 0) + c.get('kg', 0)
    return max(agg, key=agg.get) if agg else ''


def romaneo_id(p):
    """Huella estable por contenido. Mismo romaneo → mismo id, aunque el archivo
    se llame igual que otro."""
    base = '|'.join([
        str(p.get('numero', '')),
        str(p.get('fecha', '')),
        str(int(round(p.get('kg_entrada', 0) or 0))),
        str(p.get('medias_reses', 0)),
        str(len([c for c in p.get('cortes', []) if c.get('grupo') != 'GRASA'])),
        str(int(round(_kg_carne(p)))),
    ])
    return hashlib.md5(base.encode('utf-8')).hexdigest()[:12]


# Marcadores en el texto del PDF que identifican el canal.
# China: faenador Delta Car (DELTACAR) o planta de desposte TOP MEAT.
MARCADORES_CHINA = ['DELTACAR', 'DELTA CAR', 'TOP MEAT', 'TOPMEAT']
# Umbral de cortes para distinguir China 6 cortes vs 23 cortes.
UMBRAL_CHINA_6 = 12  # <= 12 cortes distintos → "6 cortes"; más → "23 cortes"


def sugerir_segmento(p):
    """Sugerencia automática del canal.
    - Búfalo: categoría Bubalino o corte bubalino.
    - China: el PDF menciona el faenador Delta Car o el desposte TOP Meat.
      Se distingue 6 vs 23 cortes por la cantidad de cortes distintos.
    - Resto: Consumo (Hilton/Black los confirma el usuario hasta calibrar)."""
    if p.get('categoria') == 'Bubalino' or any(c.get('es_bubalino') for c in p.get('cortes', [])):
        return 'Búfalo'

    texto = (p.get('texto_fuente', '') or '').upper()
    if any(m in texto for m in MARCADORES_CHINA):
        n = len(grupos_distintos(p))
        return 'Vaca China 6' if n <= UMBRAL_CHINA_6 else 'Vaca China 23'

    # Hilton: contramarca JC (en los cortes o en el texto del PDF)
    if _tiene_contramarca_hilton(p):
        return 'Novillo Hilton'

    return 'Consumo'


def _tiene_contramarca_hilton(p):
    """Detecta Hilton por la contramarca 'JC'."""
    for c in p.get('cortes', []):
        if str(c.get('contramarca', '')).upper().strip() == 'JC':
            return True
    import re
    t = (p.get('texto_fuente', '') or '').upper()
    return bool(re.search(r'\bJC\s*-\s*\d', t))


def resumen(p):
    """Datos legibles para mostrar en la grilla de selección."""
    return {
        'fecha': p.get('fecha', '—'),
        'categoria': p.get('categoria', '—'),
        'medias': p.get('medias_reses', 0),
        'kg_entrada': int(round(p.get('kg_entrada', 0) or 0)),
        'n_cortes': len(grupos_distintos(p)),
        'destino': destino_principal(p),
    }


# ───────── persistencia del registro de segmentos ─────────
def cargar_registro(path):
    if path and os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def guardar_registro(path, registro):
    if not path:
        return
    d = os.path.dirname(path)
    if d:
        os.makedirs(d, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(registro, f, ensure_ascii=False, indent=2)


def segmento_de(p, registro):
    """Devuelve (segmento, confirmado). Si el usuario ya lo guardó, usa eso;
    si no, la sugerencia automática."""
    rid = romaneo_id(p)
    if rid in registro and registro[rid].get('segmento'):
        return registro[rid]['segmento'], True
    return sugerir_segmento(p), False


def set_segmento(path, registro, p, segmento):
    """Guarda la elección del usuario para este romaneo (persistente)."""
    rid = romaneo_id(p)
    registro[rid] = dict(resumen(p), segmento=segmento, confirmado=True)
    guardar_registro(path, registro)
    return registro
