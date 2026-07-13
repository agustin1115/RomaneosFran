"""
html_builder_produccion.py — Reporte HTML autocontenido para el GERENTE DE
PRODUCCIÓN. Espejo del análisis Excel pero enfocado en rendimiento, calidad
de despostada, alertas operativas y sugerencias de acción.

NO incluye datos comerciales (precios, márgenes, rentabilidad).
"""
from collections import defaultdict
from datetime import datetime

from config import (GRUPOS_POR_CALIDAD, CORTES_CAROS, REND_OBJETIVO,
                    SUBCORTE_TO_PARENT, SUBCORTES, AMARILLA_CONTRAMARCAS)


# ═══════════════════════════════════════════════════════════════════════
# UMBRALES — todos parametrizables vía argument override
# ═══════════════════════════════════════════════════════════════════════
UMBRALES_DEFAULT = {
    # Rendimiento (puntos porcentuales por DEBAJO del objetivo)
    'rend_amarilla_pp':     1.0,
    'rend_roja_pp':         2.5,
    'rend_negra_pp':        4.0,
    # Carne amarilla (% sobre kg carne)
    'amarilla_media_min':   6.0,
    'amarilla_roja_min':    10.0,
    # Recortes — exceso en pp sobre lo esperado
    'recorte_amarilla_pp':  3.0,
    'recorte_roja_pp':      5.0,
    # Días faena → producción
    'dias_amarilla':        4,
    'dias_roja':            9,
    # Peso por unidad porcionada/feteada
    'porcionado_min_kg':    1.0,
    'porcionado_max_kg':    1.5,
    # Grasa + decomiso (% sobre kg entrada)
    'grasa_amarilla_pct':   6.0,
    'grasa_roja_pct':       9.0,
    # Merma total (% sobre kg entrada)
    'merma_amarilla_pct':   1.5,
    'merma_roja_pct':       3.0,
    # Aumento merma vs histórico (pp)
    'merma_alerta_vs_hist': 1.5,
}


# ═══════════════════════════════════════════════════════════════════════
# CSS embed — autocontenido, mobile-first
# ═══════════════════════════════════════════════════════════════════════
CSS = """
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
     background:#F5F7F5;color:#1A1A1A;line-height:1.5;padding-bottom:60px}
header.hero{background:linear-gradient(135deg,#1B4D3E,#2D7D5F);color:#fff;
            padding:18px 22px;display:flex;justify-content:space-between;
            align-items:center;flex-wrap:wrap;gap:14px;
            box-shadow:0 2px 12px rgba(0,0,0,.18);position:sticky;top:0;z-index:50}
header.hero h1{font-size:20px;letter-spacing:2px;font-weight:800;text-transform:uppercase}
header.hero p{font-size:11px;color:#C9A84C;letter-spacing:1.2px;
              text-transform:uppercase;margin-top:2px;font-weight:600}
.estado-pill{padding:8px 16px;border-radius:24px;font-weight:800;
             font-size:12px;letter-spacing:1.4px;text-transform:uppercase;
             box-shadow:0 2px 8px rgba(0,0,0,.18)}
.estado-verde{background:#27AE60;color:#fff}
.estado-amarilla{background:#F9A825;color:#1A1A1A}
.estado-roja{background:#E53935;color:#fff}
.estado-negra{background:#0A0A0A;color:#fff;border:2px solid #FF1744}

.container{max-width:1100px;margin:0 auto;padding:18px}
section{background:#fff;border-radius:14px;padding:18px 20px;margin-bottom:14px;
        box-shadow:0 1px 6px rgba(0,0,0,.06)}
section h2{font-size:14px;text-transform:uppercase;letter-spacing:1.5px;
           color:#1B4D3E;font-weight:700;margin-bottom:14px;padding-bottom:8px;
           border-bottom:2px solid #E8F5E9}
section h3{font-size:13px;color:#444;font-weight:700;margin-top:16px;margin-bottom:8px}

/* KPIs */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:10px;
      background:transparent;box-shadow:none;padding:0}
.kpi{background:#fff;border-radius:12px;padding:14px;box-shadow:0 1px 6px rgba(0,0,0,.05);
     border-left:4px solid #2D7D5F}
.kpi.verde{border-left-color:#27AE60}
.kpi.amarilla{border-left-color:#F9A825;background:#FFFCEF}
.kpi.roja{border-left-color:#E53935;background:#FFEEEE}
.kpi.negra{border-left-color:#0A0A0A;background:#1A1A1A;color:#fff}
.kpi.negra .kpi-label{color:#C9A84C}
.kpi-label{font-size:10px;text-transform:uppercase;letter-spacing:1.2px;
           color:#666;font-weight:700;margin-bottom:4px}
.kpi-value{font-size:24px;font-weight:800;color:#1B4D3E;letter-spacing:-.5px}
.kpi.negra .kpi-value{color:#fff}
.kpi-meta{font-size:10px;color:#888;margin-top:4px}

/* Alertas */
.alertas-list{display:flex;flex-direction:column;gap:8px}
.alerta{padding:10px 14px;border-radius:8px;font-size:13px;
        display:flex;align-items:center;gap:10px;font-weight:500}
.alerta-icon{font-size:14px}
.alerta.amarilla{background:#FFF8E1;border-left:4px solid #F9A825;color:#5C4400}
.alerta.roja{background:#FFEBEE;border-left:4px solid #E53935;color:#8B0000}
.alerta.negra{background:#1A1A1A;color:#fff;border-left:4px solid #FF1744}
.alerta.verde{background:#E8F5E9;border-left:4px solid #27AE60;color:#1B5E20}

.sugerencia{padding:10px 14px;border-radius:8px;background:#E3F2FD;
            border-left:4px solid #1976D2;color:#0D47A1;font-size:13px;margin-top:6px}

/* Tablas */
.tabla{width:100%;border-collapse:collapse;font-size:13px}
.tabla thead th{background:#1B4D3E;color:#fff;padding:8px 10px;text-align:left;
                font-size:11px;text-transform:uppercase;letter-spacing:.5px;font-weight:700}
.tabla tbody td{padding:7px 10px;border-bottom:1px solid #EEE}
.tabla tbody tr:hover{background:#F9FBF9}
.tabla.compact thead th, .tabla.compact tbody td{padding:5px 8px;font-size:12px}
.row-amarilla{background:#FFF8E1!important}
.row-roja{background:#FFEBEE!important}
.row-negra{background:#1A1A1A!important;color:#fff}
.row-negra td{border-bottom-color:#444}

.calif{padding:3px 9px;border-radius:14px;font-size:10px;font-weight:700;
       letter-spacing:.5px;text-transform:uppercase;display:inline-block}
.calif-óptimo{background:#27AE60;color:#fff}
.calif-bueno{background:#42A5F5;color:#fff}
.calif-regular{background:#F9A825;color:#1A1A1A}
.calif-malo{background:#E53935;color:#fff}
.calif-extra{background:#9C27B0;color:#fff}
.calif-n-a{background:#999;color:#fff}

.caro{color:#C62828;font-weight:700}
.barato{color:#2E7D32;font-weight:600}

.estado{padding:2px 8px;border-radius:12px;font-size:10px;font-weight:700;
        text-transform:uppercase}
.estado.verde{background:#E8F5E9;color:#1B5E20}
.estado.amarilla{background:#FFF8E1;color:#5C4400}
.estado.roja{background:#FFEBEE;color:#8B0000}
.alerta-inline{padding:1px 8px;border-radius:10px;font-size:10px;font-weight:700;
               margin-left:8px}
.alerta-inline.amarilla{background:#FFF8E1;color:#5C4400}
.alerta-inline.roja{background:#FFEBEE;color:#8B0000}
.alerta-inline.negra{background:#1A1A1A;color:#fff}

/* Faltantes */
.faltantes-block{margin-bottom:12px}
.faltantes-roja{color:#C62828}
.faltantes-amarilla{color:#F57F17}
.lista-faltantes{padding-left:22px;font-size:13px;line-height:1.7}
.lista-faltantes li{margin-bottom:4px}
.lista-faltantes strong{color:#1B4D3E}

.metric-line{padding:8px 0;border-bottom:1px solid #F0F0F0;font-size:13px}
.metric-line:last-child{border-bottom:none}

.tropas-note{font-size:11px;color:#777;font-style:italic;margin-top:8px}

footer{text-align:center;padding:18px;font-size:11px;color:#999}

@media(max-width:600px){
  .container{padding:12px}
  section{padding:14px 16px}
  header.hero{padding:14px 16px}
  header.hero h1{font-size:17px}
  .kpi-value{font-size:20px}
}
@media print{
  body{background:#fff}
  section{box-shadow:none;border:1px solid #DDD;page-break-inside:avoid}
  header.hero{position:static}
  .kpi.negra{background:#fff;color:#000;border:2px solid #000}
  .kpi.negra .kpi-value{color:#000}
}
"""


# ═══════════════════════════════════════════════════════════════════════
# Helpers de formato
# ═══════════════════════════════════════════════════════════════════════
def _fmt_kg(v):
    try:
        return f'{float(v):,.0f}'.replace(',', '.')
    except (ValueError, TypeError):
        return '—'


def _fmt_kg2(v):
    try:
        return f'{float(v):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return '—'


def _fmt_pct(v, decimales=2):
    try:
        return f'{float(v):.{decimales}f}%'.replace('.', ',')
    except (ValueError, TypeError):
        return '—'


def _esc(s):
    """Escape HTML básico."""
    if s is None:
        return ''
    return (str(s).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))


# ═══════════════════════════════════════════════════════════════════════
# Evaluadores de alertas
# ═══════════════════════════════════════════════════════════════════════
def _eval_rend(desv_pp, U):
    if desv_pp >= U['rend_negra_pp']:
        return 'negra', f'⚠️ PELIGRO: Rendimiento {desv_pp:.2f}pp por debajo del objetivo'
    if desv_pp >= U['rend_roja_pp']:
        return 'roja', f'Rendimiento {desv_pp:.2f}pp por debajo del objetivo'
    if desv_pp >= U['rend_amarilla_pp']:
        return 'amarilla', f'Rendimiento {desv_pp:.2f}pp por debajo del objetivo'
    return 'verde', None


def _eval_amarilla(pct, U):
    if pct >= U['amarilla_roja_min']:
        return 'roja', f'Carne amarilla {pct:.1f}% — alto'
    if pct >= U['amarilla_media_min']:
        return 'amarilla', f'Carne amarilla {pct:.1f}% — atención'
    return 'verde', None


def _eval_dias(dias, U):
    if dias is None:
        return 'verde', None
    if dias >= U['dias_roja']:
        return 'roja', f'{dias:.0f} días entre faena y producción — oreo prolongado'
    if dias >= U['dias_amarilla']:
        return 'amarilla', f'{dias:.0f} días entre faena y producción'
    return 'verde', None


def _eval_grasa(pct, U):
    if pct >= U['grasa_roja_pct']:
        return 'roja', f'Grasa+decomiso {pct:.1f}% — fuera de norma'
    if pct >= U['grasa_amarilla_pct']:
        return 'amarilla', f'Grasa+decomiso {pct:.1f}%'
    return 'verde', None


def _eval_merma(pct, U):
    if pct >= U['merma_roja_pct']:
        return 'roja', f'Merma total {pct:.2f}% — investigar'
    if pct >= U['merma_amarilla_pct']:
        return 'amarilla', f'Merma total {pct:.2f}%'
    return 'verde', None


def _eval_recorte(desv_pp, U):
    if desv_pp >= U['recorte_roja_pp']:
        return 'roja', f'Exceso de recorte +{desv_pp:.1f}pp'
    if desv_pp >= U['recorte_amarilla_pp']:
        return 'amarilla', f'Exceso de recorte +{desv_pp:.1f}pp'
    return 'verde', None


def _peor_nivel(*niveles):
    """Devuelve el nivel más severo de la lista."""
    orden = ['negra', 'roja', 'amarilla', 'verde']
    presentes = [n for n in niveles if n]
    for nv in orden:
        if nv in presentes:
            return nv
    return 'verde'


def _calif_corte(pct_real, pct_esperado, kg_real, es_caro):
    """Calificación igual que el Excel actual."""
    if pct_esperado == 0:
        return 'N/A' if kg_real == 0 else 'EXTRA'
    cumpl = pct_real / pct_esperado
    if es_caro:
        if cumpl >= 1.10: return 'ÓPTIMO'
        if cumpl >= 0.95: return 'BUENO'
        if cumpl >= 0.80: return 'REGULAR'
        return 'MALO'
    if cumpl <= 0.80: return 'ÓPTIMO'
    if cumpl <= 0.95: return 'BUENO'
    if cumpl <= 1.10: return 'REGULAR'
    return 'MALO'


# ═══════════════════════════════════════════════════════════════════════
# Sugerencias de acción
# ═══════════════════════════════════════════════════════════════════════
def _generar_sugerencias(ctx):
    """Construye lista de sugerencias contextuales basadas en alertas detectadas."""
    sug = []
    if ctx['alerta_rend_nivel'] in ('roja', 'negra'):
        sug.append(('🔴 Bajo rendimiento',
                    'Revisar urgentemente al equipo de despostada del turno y la merma de oreo. '
                    'Si la mañana tuvo problemas, comparar contra el turno tarde.'))
    elif ctx['alerta_rend_nivel'] == 'amarilla':
        sug.append(('🟡 Rendimiento bajo objetivo',
                    'Atención: revisar si hubo cortes mal categorizados o desperdicio anormal en el cuarteo.'))

    if ctx['alerta_amarilla_nivel'] in ('roja', 'amarilla'):
        sug.append(('🟠 Carne amarilla alta',
                    'Reclamar al proveedor de hacienda. Revisar criterios de selección en compra '
                    f'(tropa(s): {", ".join(ctx["tropas_ids"]) or "N/A"}).'))

    if ctx['cortes_caros_faltantes']:
        nombres = ', '.join(c[0] for c in ctx['cortes_caros_faltantes'][:5])
        sug.append(('🔴 Cortes caros faltantes',
                    f'Verificar si {nombres} fueron categorizados correctamente o se perdieron en grasa/recorte. '
                    'Pueden estar bajo otro nombre en el romaneo.'))

    if ctx['cortes_baratos_faltantes']:
        nombres = ', '.join(c[0] for c in ctx['cortes_baratos_faltantes'][:5])
        sug.append(('💡 Cortes baratos faltantes',
                    f'{nombres} aparecen como esperados pero llegaron en 0 kg. '
                    'Si nunca llegan en esta calidad, considerá ajustar el % esperado en config '
                    'para que no genere falsa alerta.'))

    if ctx['picada_excesiva']:
        sug.append(('🟡 Exceso de picada',
                    'Mucha picada respecto al esperado. Calibrar despostada: probablemente '
                    'cortes nobles fueron a picada por mala separación.'))

    if ctx['recorte_excesivo']:
        sug.append(('🟡 Exceso de recorte',
                    'Recorte por encima del % esperado. Revisar separación de cortes y mermas en cuarteo.'))

    if ctx['alerta_dias_nivel'] in ('roja', 'amarilla'):
        sug.append(('⏱️ Oreo prolongado',
                    'Faena hace varios días. Esto puede impactar peso y color. '
                    'Coordinar con el frigorífico para ajustar tiempos.'))

    if ctx['alerta_grasa_nivel'] in ('roja', 'amarilla'):
        sug.append(('🛢️ Grasa y decomiso alto',
                    'Mucha grasa. Revisar si hay sobre-recorte o problemas de calidad de la tropa.'))

    if ctx['alerta_merma_nivel'] in ('roja', 'amarilla') and not ctx['mermas_separadas']:
        sug.append(('📊 Mermas no separadas',
                    'La merma total es alta y no está discriminada por etapa. '
                    'Considerá registrar separadamente: <strong>oreo / cuarteo / despostada</strong> '
                    'para identificar dónde se concentra y mejorarla.'))

    if ctx['oportunidad_merma']:
        sug.append(('📉 Oportunidad de mejora — merma',
                    f'En tropas anteriores la merma fue de {ctx["mejor_merma_hist"]:.2f}%. '
                    f'Hoy fue {ctx["merma_pct"]:.2f}% (+{ctx["delta_merma"]:.2f}pp). '
                    'Investigar qué se hizo distinto en aquellas tropas.'))

    if ctx['cortes_porcionado_problemas']:
        nombres = ', '.join(c['nombre'] for c in ctx['cortes_porcionado_problemas'][:5])
        sug.append(('⚖️ Pesos porcionados fuera de rango',
                    f'{nombres} con pesos por unidad fuera del target 1,0–1,5 kg. '
                    'Calibrar las máquinas porcionadoras del turno.'))

    if ctx['cortes_sin_clasificar']:
        sug.append(('❓ Cortes sin clasificar',
                    f'{len(ctx["cortes_sin_clasificar"])} cortes no se reconocieron. '
                    'Reclasificar manualmente o agregar al diccionario para futuros romaneos.'))

    if ctx['n_tropas'] > 1:
        sug.append(('🐂 Múltiples tropas en el romaneo',
                    f'Este romaneo agrupa {ctx["n_tropas"]} tropas. Para análisis más fino, '
                    'considerá separar la despostada por tropa para detectar cuál es problemática.'))

    if not sug:
        sug.append(('✅ Sin sugerencias — todo en orden',
                    'No se detectaron desviaciones significativas. ¡Buen trabajo!'))

    return sug


# ═══════════════════════════════════════════════════════════════════════
# BUILDER PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════
def build_html_produccion(data, calidad='Standard', historial=None,
                          umbrales_override=None, titulo_extra=None):
    """
    Genera HTML autocontenido para el gerente de producción.

    Args:
        data: dict de un romaneo parseado o un acumulado (con 'romaneo' y 'cortes')
        calidad: 'Standard' | 'Búfalo' | 'Premium Black' | 'Exportación'
        historial: lista de romaneos previos (para comparativas), opcional
        umbrales_override: dict para sobreescribir umbrales default
        titulo_extra: texto adicional para el título (ej. "ACUMULADO")

    Returns: HTML string
    """
    U = dict(UMBRALES_DEFAULT)
    if umbrales_override:
        U.update(umbrales_override)

    rom = data['romaneo'] if 'romaneo' in data else data
    cortes = data.get('cortes', rom.get('cortes', [])) or []
    archivo = rom.get('archivo', 'Romaneo')
    fecha = rom.get('fecha', '—')
    medias = rom.get('medias_reses', 0)
    kg_entrada = rom.get('kg_entrada', 0) or 0
    kg_carne_field = rom.get('kg_carne', 0) or 0
    categoria_raw = rom.get('categoria', '—') or '—'
    cat_clean = categoria_raw.split('(')[0].strip() if categoria_raw else '—'

    # Métricas de carne
    meat_cortes = [c for c in cortes if c.get('grupo') != 'GRASA']
    kg_total_meat = sum(c.get('kg', 0) for c in meat_cortes)
    if kg_carne_field == 0:
        kg_carne = kg_total_meat
    else:
        kg_carne = kg_carne_field
    kg_carne_div = kg_total_meat or kg_carne or 1

    rend_real_pct = (kg_carne / kg_entrada * 100) if kg_entrada else 0
    rend_obj_base = REND_OBJETIVO.get(calidad, REND_OBJETIVO['Standard']).get(cat_clean, 0.66) * 100

    kg_anatomico = sum(c.get('kg', 0) for c in meat_cortes if c.get('tipo') == 'ANATÓMICO')
    pct_anatomico = (kg_anatomico / kg_carne_div * 100) if kg_carne_div else 0
    rend_obj_pct = rend_obj_base + (1.0 if pct_anatomico > 50 else 0.0)
    desv_pp = rend_obj_pct - rend_real_pct

    # Carne amarilla
    pct_am = rom.get('pct_amarilla', 0) or 0
    if isinstance(pct_am, (int, float)) and 0 < pct_am < 1:
        pct_am_pct = pct_am * 100
    else:
        pct_am_pct = float(pct_am or 0)
    if pct_am_pct == 0 and cortes:
        kg_am = sum(c.get('kg', 0) for c in cortes
                    if str(c.get('contramarca', '')) in AMARILLA_CONTRAMARCAS)
        pct_am_pct = (kg_am / kg_carne_div * 100) if kg_carne_div else 0

    # Días faena→producción
    dias_faena = rom.get('dias_faena_produccion')

    # Grasa y decomiso
    kg_grasa = sum(c.get('kg', 0) for c in cortes if c.get('grupo') == 'GRASA')
    pct_grasa = (kg_grasa / kg_entrada * 100) if kg_entrada else 0

    # Merma total
    kg_total_salidas = sum(c.get('kg', 0) for c in cortes)
    merma_kg = max(0, kg_entrada - kg_total_salidas)
    merma_pct = (merma_kg / kg_entrada * 100) if kg_entrada else 0
    merma_separada = bool(rom.get('merma_kg', 0)) and bool(rom.get('grasa_kg', 0))

    # Tropas
    tropas_match = rom.get('tropas_match', []) or []
    n_tropas = len(tropas_match)
    tropas_ids = [str(t.get('tropa', '')) for t in tropas_match]

    # ───── Evaluar alertas globales ─────
    nivel_rend, msg_rend = _eval_rend(desv_pp, U)
    nivel_am, msg_am = _eval_amarilla(pct_am_pct, U)
    nivel_dias, msg_dias = _eval_dias(dias_faena, U)
    nivel_grasa, msg_grasa = _eval_grasa(pct_grasa, U)
    nivel_merma, msg_merma = _eval_merma(merma_pct, U)

    # Estado general
    nivel_general = _peor_nivel(nivel_rend, nivel_am, nivel_dias, nivel_grasa, nivel_merma)
    estado_label = {
        'negra': '⚠️ ATENCIÓN URGENTE',
        'roja': '🔴 Revisión necesaria',
        'amarilla': '🟡 Atención',
        'verde': '✅ Todo OK',
    }[nivel_general]

    # ───── Análisis por grupo (cortes faltantes incluidos) ─────
    GRUPOS = GRUPOS_POR_CALIDAD.get(calidad, GRUPOS_POR_CALIDAD['Standard'])
    kg_por_grupo = defaultdict(float)
    for c in cortes:
        g = c.get('grupo')
        if g and g not in ('GRASA',):
            kg_por_grupo[g] += c.get('kg', 0)

    subcortes_set = set(SUBCORTE_TO_PARENT.keys())
    GRUPOS_VIS = [(g, p) for (g, p) in GRUPOS if g not in subcortes_set]

    cortes_caros_faltantes = []
    cortes_baratos_faltantes = []
    cortes_malo_caro = []
    picada_excesiva = False
    recorte_excesivo = False
    filas_rendimiento = []  # tuplas para la tabla

    for grupo_name, pct_esp in GRUPOS_VIS:
        kg_real = kg_por_grupo.get(grupo_name, 0)
        for sc in SUBCORTES.get(grupo_name, []):
            kg_real += kg_por_grupo.get(sc[0], 0)
        pct_real = (kg_real / kg_carne_div * 100) if kg_carne_div else 0
        pct_esp_pct = pct_esp * 100
        desv = pct_real - pct_esp_pct
        cumpl = (pct_real / pct_esp_pct * 100) if pct_esp_pct > 0 else (100 if kg_real == 0 else 150)
        es_caro = grupo_name in CORTES_CAROS
        calif = _calif_corte(pct_real, pct_esp_pct, kg_real, es_caro)

        # Alertas por fila
        alertas_row = []
        if pct_esp_pct > 0 and kg_real == 0:
            if es_caro:
                alertas_row.append(('roja', 'CARO faltante'))
                cortes_caros_faltantes.append((grupo_name, pct_esp_pct, kg_carne_div * pct_esp))
            else:
                alertas_row.append(('amarilla', 'Faltante (revisar parámetro)'))
                cortes_baratos_faltantes.append((grupo_name, pct_esp_pct, kg_carne_div * pct_esp))

        if calif == 'MALO':
            if es_caro:
                alertas_row.append(('roja', 'MALO en CARO'))
                cortes_malo_caro.append(grupo_name)
            elif grupo_name == 'CARNE PICADA':
                alertas_row.append(('amarilla', 'Exceso picada'))
                picada_excesiva = True

        if 'RECORTE' in grupo_name:
            nv_r, msg_r = _eval_recorte(desv, U)
            if nv_r != 'verde':
                alertas_row.append((nv_r, msg_r))
                recorte_excesivo = True

        nivel_row = _peor_nivel(*[a[0] for a in alertas_row])
        filas_rendimiento.append({
            'grupo': grupo_name,
            'kg_real': kg_real,
            'pct_real': pct_real,
            'pct_esp': pct_esp_pct,
            'desv': desv,
            'cumpl': cumpl,
            'es_caro': es_caro,
            'calif': calif,
            'alertas': alertas_row,
            'nivel': nivel_row,
        })

    # Ordenar para que faltantes críticos aparezcan arriba (los caros faltantes primero)
    cortes_caros_faltantes.sort(key=lambda x: -x[2])
    cortes_baratos_faltantes.sort(key=lambda x: -x[2])

    # Cortes porcionados/feteados — control de pesos por unidad
    cortes_porcionado_problemas = []
    porcionados = []
    for c in cortes:
        if c.get('tipo') in ('PORCIONADO', 'FETEADO') and c.get('unidades', 0) > 0 \
                and c.get('kg', 0) > 0:
            kg_pieza = c['kg'] / c['unidades']
            estado = ('verde' if U['porcionado_min_kg'] <= kg_pieza <= U['porcionado_max_kg']
                      else 'amarilla')
            porcionados.append({
                'corte': c['corte'], 'tipo': c['tipo'],
                'unidades': c['unidades'], 'kg': c['kg'],
                'kg_pieza': kg_pieza, 'estado': estado,
            })
            if estado != 'verde':
                cortes_porcionado_problemas.append({'nombre': c['corte'], 'kg_pieza': kg_pieza})

    # Cortes SIN CLASIFICAR
    sin_clasificar = [c for c in cortes if c.get('grupo') == 'SIN CLASIFICAR']

    # ───── Histórico ─────
    avg_rend_mes = avg_am_mes = None
    rend_ant = am_ant = None
    fecha_ant = None
    merma_hist_avg = merma_hist_min = None
    delta_merma = oportunidad_merma = False
    if historial:
        comp = [h for h in historial
                if (h.get('categoria', '').split('(')[0].strip() == cat_clean
                    and h.get('calidad') == calidad
                    and h.get('archivo') != archivo)]
        if comp:
            comp_sorted = sorted(comp, key=lambda h: h.get('fecha', ''), reverse=True)
            ultimo = comp_sorted[0]
            rend_ant = ultimo.get('rendimiento_pct', 0) or 0
            am_ant_raw = ultimo.get('pct_amarilla', 0) or 0
            am_ant = am_ant_raw * 100 if 0 < am_ant_raw < 1 else am_ant_raw
            fecha_ant = ultimo.get('fecha', '')

            # Promedio mes (mismo mes/año del romaneo actual)
            mes_target = fecha[3:] if len(fecha) >= 10 else ''
            comp_mes = [h for h in comp if h.get('fecha', '')[3:] == mes_target] if mes_target else []
            if comp_mes:
                rends = [h.get('rendimiento_pct', 0) or 0 for h in comp_mes]
                avg_rend_mes = sum(rends) / len(rends) if rends else None
                ams = []
                for h in comp_mes:
                    a = h.get('pct_amarilla', 0) or 0
                    ams.append(a * 100 if 0 < a < 1 else a)
                avg_am_mes = sum(ams) / len(ams) if ams else None

            # Histórico mermas: kg_entrada - kg_carne
            mermas = []
            for h in comp:
                kgi = h.get('kg_entrada', 0) or 0
                kgc = h.get('kg_carne', 0) or 0
                if kgi > 0 and kgc > 0:
                    mermas.append((kgi - kgc) / kgi * 100)
            if mermas:
                merma_hist_avg = sum(mermas) / len(mermas)
                merma_hist_min = min(mermas)
                if merma_pct - merma_hist_avg > U['merma_alerta_vs_hist']:
                    delta_merma = True
                if merma_hist_min < merma_pct - 1:
                    oportunidad_merma = True

    # ───── Construir HTML ─────
    H = []
    H.append('<!DOCTYPE html><html lang="es"><head>')
    H.append('<meta charset="UTF-8">')
    H.append('<meta name="viewport" content="width=device-width,initial-scale=1">')
    H.append(f'<title>Producción · {_esc(archivo)}</title>')
    H.append(f'<style>{CSS}</style></head><body>')

    # HERO
    titulo_extra_str = f' · {titulo_extra}' if titulo_extra else ''
    H.append('<header class="hero">')
    H.append('<div>')
    H.append(f'<h1>TF Carnes · Producción</h1>')
    H.append(f'<p>{_esc(archivo)}{titulo_extra_str} · {_esc(fecha)} · {_esc(categoria_raw)} · {_esc(calidad)} · {medias} medias</p>')
    H.append('</div>')
    H.append(f'<div class="estado-pill estado-{nivel_general}">{estado_label}</div>')
    H.append('</header>')

    H.append('<div class="container">')

    # KPIs
    H.append('<section class="kpis">')

    def _kpi(label, value, meta='', nivel=''):
        cls = f'kpi {nivel}' if nivel else 'kpi'
        return (f'<div class="{cls}">'
                f'<div class="kpi-label">{label}</div>'
                f'<div class="kpi-value">{value}</div>'
                f'<div class="kpi-meta">{meta}</div></div>')

    H.append(_kpi('Kg entrada', _fmt_kg(kg_entrada), f'{medias} medias'))
    H.append(_kpi('Kg carne', _fmt_kg(kg_carne), ''))
    H.append(_kpi('Rendimiento', _fmt_pct(rend_real_pct),
                  f'Obj {_fmt_pct(rend_obj_pct)} · Δ {desv_pp:+.2f}pp', nivel_rend))
    H.append(_kpi('% Amarilla', _fmt_pct(pct_am_pct), '', nivel_am))
    anat_meta = '+1pp ajuste obj' if pct_anatomico > 50 else ''
    H.append(_kpi('% Anatómico', _fmt_pct(pct_anatomico), anat_meta))
    if dias_faena is not None:
        H.append(_kpi('Días faena→prod', f'{int(dias_faena)} d', '', nivel_dias))
    H.append(_kpi('Cortes MALO/caros', str(len(cortes_malo_caro)),
                  ', '.join(cortes_malo_caro[:2]) + ('…' if len(cortes_malo_caro) > 2 else ''),
                  'roja' if cortes_malo_caro else 'verde'))
    H.append('</section>')

    # ALERTAS
    alertas_globales = []
    for nv, msg in [(nivel_rend, msg_rend), (nivel_am, msg_am),
                     (nivel_dias, msg_dias), (nivel_grasa, msg_grasa),
                     (nivel_merma, msg_merma)]:
        if msg:
            alertas_globales.append((nv, msg))
    if cortes_caros_faltantes:
        n = len(cortes_caros_faltantes)
        alertas_globales.append(('roja', f'{n} corte(s) caro(s) faltante(s) — crítico'))
    if cortes_malo_caro:
        alertas_globales.append(('roja',
            f'{len(cortes_malo_caro)} corte(s) caro(s) con calificación MALO: '
            + ', '.join(cortes_malo_caro)))
    if picada_excesiva:
        alertas_globales.append(('amarilla', 'Exceso de carne picada'))
    if recorte_excesivo:
        alertas_globales.append(('amarilla', 'Recortes por encima del esperado'))
    if delta_merma:
        alertas_globales.append(('roja',
            f'Merma actual {merma_pct:.2f}% vs promedio histórico {merma_hist_avg:.2f}%'))

    if alertas_globales:
        H.append('<section><h2>⚠️ Alertas críticas</h2>')
        H.append('<div class="alertas-list">')
        ord_nv = {'negra': 0, 'roja': 1, 'amarilla': 2, 'verde': 3}
        alertas_globales.sort(key=lambda x: ord_nv.get(x[0], 99))
        for nv, msg in alertas_globales:
            H.append(f'<div class="alerta {nv}"><span class="alerta-icon">●</span> {_esc(msg)}</div>')
        H.append('</div></section>')

    # RENDIMIENTO POR GRUPO
    H.append('<section><h2>📊 Rendimiento por grupo de corte</h2>')
    H.append('<table class="tabla">')
    H.append('<thead><tr><th>Grupo</th><th>Kg real</th><th>% real</th><th>% esperado</th>'
             '<th>Desv</th><th>% Cumpl</th><th>Tipo</th><th>Calif</th><th>Alerta</th></tr></thead><tbody>')
    # Ordenar: faltantes caros arriba, después por kg desc
    def _key_orden(f):
        if f['nivel'] == 'negra': return (0, 0)
        if f['nivel'] == 'roja': return (1, -f['kg_real'])
        if f['nivel'] == 'amarilla': return (2, -f['kg_real'])
        return (3, -f['kg_real'])
    for f in sorted(filas_rendimiento, key=_key_orden):
        cls = f'row-{f["nivel"]}' if f['nivel'] else ''
        cal_cls = f['calif'].lower().replace(' ', '-').replace('/', '-')
        cal_cls = cal_cls.replace('ó', 'ó')  # keep accented
        tipo_cls = 'caro' if f['es_caro'] else 'barato'
        alerta_txt = ' · '.join(a[1] for a in f['alertas']) if f['alertas'] else '—'
        H.append(f'<tr class="{cls}">')
        H.append(f'<td><strong>{_esc(f["grupo"])}</strong></td>')
        H.append(f'<td>{_fmt_kg2(f["kg_real"])}</td>')
        H.append(f'<td>{_fmt_pct(f["pct_real"])}</td>')
        H.append(f'<td>{_fmt_pct(f["pct_esp"])}</td>')
        H.append(f'<td>{f["desv"]:+.2f}pp</td>')
        H.append(f'<td>{f["cumpl"]:.0f}%</td>')
        H.append(f'<td class="{tipo_cls}">{"CARO" if f["es_caro"] else "BARATO"}</td>')
        H.append(f'<td><span class="calif calif-{cal_cls}">{f["calif"]}</span></td>')
        H.append(f'<td>{_esc(alerta_txt)}</td></tr>')

    # ── Fila SIN CLASIFICAR + listado individual de cortes no mapeados ──
    sin_clasif = [c for c in cortes if c.get('grupo') == 'SIN CLASIFICAR']
    if sin_clasif:
        from collections import defaultdict as _dd_sc
        agrup_sc = _dd_sc(float)
        for c in sin_clasif:
            agrup_sc[(c.get('corte', '?') or '?').strip()] += c.get('kg', 0)
        total_sc_kg = sum(agrup_sc.values())
        total_sc_pct = (total_sc_kg / kg_carne_div * 100) if kg_carne_div else 0
        # Detalle por corte sin clasificar (en italic gris, alineado debajo)
        for nombre, kg_val in sorted(agrup_sc.items(), key=lambda x: -x[1]):
            pct_v = (kg_val / kg_carne_div * 100) if kg_carne_div else 0
            H.append('<tr style="background:#FFF8E1;">'
                     f'<td style="color:#666;font-style:italic;padding-left:18px">⚠️ {_esc(nombre)}</td>'
                     f'<td style="color:#666">{_fmt_kg2(kg_val)}</td>'
                     f'<td style="color:#666">{_fmt_pct(pct_v)}</td>'
                     '<td>—</td><td>—</td><td>—</td><td>—</td><td>—</td>'
                     '<td style="color:#E65100">sin clasificar</td></tr>')
        # Subtotal SIN CLASIFICAR
        H.append('<tr class="row-amarilla">'
                 '<td><strong>⚠️ CORTES SIN CLASIFICAR (subtotal)</strong></td>'
                 f'<td><strong>{_fmt_kg2(total_sc_kg)}</strong></td>'
                 f'<td><strong>{_fmt_pct(total_sc_pct)}</strong></td>'
                 '<td>—</td><td>—</td><td>—</td><td>—</td><td>—</td>'
                 '<td><strong style="color:#E65100">Revisar manualmente</strong></td></tr>')

    H.append('</tbody></table></section>')

    # CORTES FALTANTES
    if cortes_caros_faltantes or cortes_baratos_faltantes:
        H.append('<section><h2>📦 Cortes faltantes</h2>')
        if cortes_caros_faltantes:
            H.append('<div class="faltantes-block">')
            H.append('<h3 class="faltantes-roja">🔴 CRÍTICOS — Cortes caros faltantes</h3>')
            H.append('<ul class="lista-faltantes">')
            for nombre, pct, kg_esp in cortes_caros_faltantes:
                H.append(f'<li><strong>{_esc(nombre)}</strong> — '
                         f'esperado {_fmt_pct(pct)} (~{_fmt_kg(kg_esp)} kg). '
                         f'Verificar si se categorizó mal o se perdió en grasa/recorte.</li>')
            H.append('</ul></div>')
        if cortes_baratos_faltantes:
            H.append('<div class="faltantes-block">')
            H.append('<h3 class="faltantes-amarilla">🟡 Cortes baratos faltantes</h3>')
            H.append('<ul class="lista-faltantes">')
            for nombre, pct, kg_esp in cortes_baratos_faltantes:
                H.append(f'<li><strong>{_esc(nombre)}</strong> — '
                         f'esperado {_fmt_pct(pct)}. '
                         f'Si nunca llega en {_esc(calidad)}, sugerí ajustar el % esperado '
                         f'en config para no generar falsa alerta.</li>')
            H.append('</ul></div>')
        H.append('</section>')

    # POR TROPA
    if n_tropas >= 1:
        H.append(f'<section><h2>🐂 Análisis por tropa ({n_tropas} tropa(s))</h2>')
        if n_tropas > 1:
            H.append('<div class="alerta amarilla"><span class="alerta-icon">●</span> '
                     f'Este romaneo agrupa {n_tropas} tropas. El % amarilla y rendimiento '
                     'mostrados son del romaneo agregado.</div>')
        H.append('<table class="tabla">')
        H.append('<thead><tr><th>Tropa</th><th>Kg</th><th>% del total</th><th>Origen</th>'
                 '<th>Fecha faena</th><th>Partidas</th></tr></thead><tbody>')
        kg_tropas_tot = sum(t.get('kg', 0) for t in tropas_match) or 1
        for t in tropas_match:
            tropa_id = str(t.get('tropa', ''))
            es_compra = t.get('es_compra_media', False) or (
                tropa_id.startswith('7') and len(tropa_id) >= 4)
            origen = '🚚 Compra media' if es_compra else '🐂 Faena propia'
            ff = t.get('fecha_faena_dt')
            ff_str = ff.strftime('%d/%m/%Y') if ff else (t.get('fecha', '—') or '—')
            kg_t = t.get('kg', 0)
            pct_t = (kg_t / kg_tropas_tot * 100) if kg_tropas_tot else 0
            partidas = t.get('partidas', 1)
            part_str = f'{partidas} partidas' if partidas > 1 else '1 partida'
            H.append('<tr>')
            H.append(f'<td><strong>{_esc(tropa_id)}</strong></td>')
            H.append(f'<td>{_fmt_kg2(kg_t)}</td>')
            H.append(f'<td>{_fmt_pct(pct_t)}</td>')
            H.append(f'<td>{origen}</td>')
            H.append(f'<td>{ff_str}</td>')
            H.append(f'<td>{part_str}</td></tr>')
        H.append('</tbody></table>')
        H.append('</section>')

    # CALIDAD DESPOSTADA
    H.append('<section><h2>✂️ Calidad de despostada</h2>')
    H.append('<div class="metric-line">')
    H.append(f'Anatómico: <strong>{_fmt_pct(pct_anatomico)}</strong> · ')
    H.append(f'Porcionado/Feteado: <strong>{_fmt_pct(100-pct_anatomico)}</strong>')
    H.append('</div>')
    if porcionados:
        H.append('<h3>Pesos por unidad porcionada/feteada (target 1,0–1,5 kg)</h3>')
        H.append('<table class="tabla compact">')
        H.append('<thead><tr><th>Corte</th><th>Tipo</th><th>Unidades</th><th>Kg total</th>'
                 '<th>Kg/unidad</th><th>Estado</th></tr></thead><tbody>')
        for p in sorted(porcionados, key=lambda p: -p['kg']):
            H.append('<tr>')
            H.append(f'<td>{_esc(p["corte"])}</td><td>{p["tipo"]}</td>'
                     f'<td>{p["unidades"]}</td><td>{_fmt_kg2(p["kg"])}</td>'
                     f'<td><strong>{p["kg_pieza"]:.2f} kg</strong></td>'
                     f'<td><span class="estado {p["estado"]}">'
                     f'{"OK" if p["estado"]=="verde" else "Fuera rango"}</span></td>')
            H.append('</tr>')
        H.append('</tbody></table>')
    if sin_clasificar:
        H.append('<h3>❓ Cortes SIN CLASIFICAR (revisar)</h3><ul class="lista-faltantes">')
        for c in sin_clasificar:
            H.append(f'<li>{_esc(c.get("corte","?"))} ({_fmt_kg2(c.get("kg",0))} kg)</li>')
        H.append('</ul>')
    H.append('</section>')

    # MERMAS Y SUBPRODUCTOS
    H.append('<section><h2>🛢️ Mermas y subproductos</h2>')
    grasa_alerta = ''
    if msg_grasa:
        grasa_alerta = f' <span class="alerta-inline {nivel_grasa}">{_esc(msg_grasa)}</span>'
    H.append(f'<div class="metric-line">Grasa y decomiso: <strong>{_fmt_kg(kg_grasa)} kg</strong> '
             f'({_fmt_pct(pct_grasa)} s/entrada){grasa_alerta}</div>')
    merma_alerta = ''
    if msg_merma:
        merma_alerta = f' <span class="alerta-inline {nivel_merma}">{_esc(msg_merma)}</span>'
    H.append(f'<div class="metric-line">Merma total (entrada − salidas): '
             f'<strong>{_fmt_kg(merma_kg)} kg</strong> ({_fmt_pct(merma_pct)}){merma_alerta}</div>')

    if not merma_separada:
        H.append('<div class="alerta amarilla"><span class="alerta-icon">●</span> '
                 'Las mermas no están separadas por etapa. Para diagnóstico fino, '
                 'considerá registrar: <strong>oreo / cuarteo / despostada</strong>.</div>')

    if merma_hist_avg is not None:
        H.append(f'<div class="metric-line">Histórico merma {_esc(cat_clean)} ({_esc(calidad)}): '
                 f'promedio <strong>{_fmt_pct(merma_hist_avg)}</strong>, '
                 f'mejor caso <strong>{_fmt_pct(merma_hist_min)}</strong></div>')
        if delta_merma:
            H.append(f'<div class="alerta roja"><span class="alerta-icon">●</span> '
                     f'Merma actual {merma_pct - merma_hist_avg:+.2f}pp arriba del promedio histórico — investigar causa.</div>')
        elif oportunidad_merma:
            H.append(f'<div class="sugerencia">📉 Mejor histórico: {_fmt_pct(merma_hist_min)}. '
                     f'Hoy: {_fmt_pct(merma_pct)}. Oportunidad de mejora: '
                     f'{merma_pct - merma_hist_min:.2f}pp.</div>')
    H.append('</section>')

    # COMPARATIVA HISTÓRICA
    if rend_ant is not None or avg_rend_mes is not None:
        H.append(f'<section><h2>📈 Comparativa — {_esc(cat_clean)} ({_esc(calidad)})</h2>')
        H.append('<table class="tabla compact"><thead><tr><th></th>'
                 '<th>Actual</th><th>Romaneo anterior</th><th>Δ</th>'
                 '<th>Promedio mes</th><th>Δ vs promedio</th></tr></thead><tbody>')
        # Rendimiento
        d_ant = (rend_real_pct - (rend_ant or 0)) if rend_ant else None
        d_mes = (rend_real_pct - (avg_rend_mes or 0)) if avg_rend_mes else None
        H.append(f'<tr><td>Rendimiento</td>'
                 f'<td><strong>{_fmt_pct(rend_real_pct)}</strong></td>'
                 f'<td>{_fmt_pct(rend_ant) if rend_ant else "—"}'
                 f' <small>({_esc(fecha_ant or "")})</small></td>'
                 f'<td>{(d_ant and f"{d_ant:+.2f}pp") or "—"}</td>'
                 f'<td>{_fmt_pct(avg_rend_mes) if avg_rend_mes else "—"}</td>'
                 f'<td>{(d_mes and f"{d_mes:+.2f}pp") or "—"}</td></tr>')
        # Amarilla
        d_ant_am = (pct_am_pct - (am_ant or 0)) if am_ant is not None else None
        d_mes_am = (pct_am_pct - (avg_am_mes or 0)) if avg_am_mes is not None else None
        H.append(f'<tr><td>% Amarilla</td>'
                 f'<td><strong>{_fmt_pct(pct_am_pct)}</strong></td>'
                 f'<td>{_fmt_pct(am_ant) if am_ant is not None else "—"}</td>'
                 f'<td>{(d_ant_am is not None and f"{d_ant_am:+.2f}pp") or "—"}</td>'
                 f'<td>{_fmt_pct(avg_am_mes) if avg_am_mes is not None else "—"}</td>'
                 f'<td>{(d_mes_am is not None and f"{d_mes_am:+.2f}pp") or "—"}</td></tr>')
        H.append('</tbody></table></section>')

    # SUGERENCIAS DE ACCIÓN
    ctx = {
        'alerta_rend_nivel': nivel_rend,
        'alerta_amarilla_nivel': nivel_am,
        'alerta_dias_nivel': nivel_dias,
        'alerta_grasa_nivel': nivel_grasa,
        'alerta_merma_nivel': nivel_merma,
        'tropas_ids': tropas_ids,
        'cortes_caros_faltantes': cortes_caros_faltantes,
        'cortes_baratos_faltantes': cortes_baratos_faltantes,
        'picada_excesiva': picada_excesiva,
        'recorte_excesivo': recorte_excesivo,
        'mermas_separadas': merma_separada,
        'oportunidad_merma': oportunidad_merma,
        'mejor_merma_hist': merma_hist_min if merma_hist_min is not None else 0,
        'merma_pct': merma_pct,
        'delta_merma': (merma_pct - merma_hist_min) if merma_hist_min is not None else 0,
        'cortes_porcionado_problemas': cortes_porcionado_problemas,
        'cortes_sin_clasificar': sin_clasificar,
        'n_tropas': n_tropas,
    }
    sugerencias = _generar_sugerencias(ctx)
    H.append('<section><h2>💡 Sugerencias de acción</h2>')
    for titulo, texto in sugerencias:
        H.append(f'<div class="sugerencia"><strong>{_esc(titulo)}</strong><br>{texto}</div>')
    H.append('</section>')

    # FOOTER
    H.append(f'<footer>Generado {datetime.now().strftime("%d/%m/%Y %H:%M")} '
             '· TF Carnes · Reporte de producción</footer>')
    H.append('</div></body></html>')

    return '\n'.join(H)
