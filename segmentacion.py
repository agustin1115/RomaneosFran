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
SEGMENTOS = ['Consumo', 'Vaca China 6', 'Vaca China 23', 'Novillo Hilton', 'Black', 'Búfalo']


def _kg_carne(p):
    return sum(c.get('kg', 0) for c in p.get('cortes', []) if c.get('grupo') != 'GRASA')


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
