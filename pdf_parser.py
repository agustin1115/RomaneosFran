"""
pdf_parser.py — Parsea PDFs de romaneo (Rendimientos/Resultado de Despostada).
Extrae encabezado, líneas de cortes, grasa/decomiso y merma.

Formatos soportados:
  - "Rendimientos de Despostada": ctm al final como "47- 215"
  - "Resultado Despostada": ctm antes de piezas como "47-200215"
"""
import re
from config import (CORTE_TO_GRUPO, TIPIFICACION_MAP, CONTRAMARCA_MAP,
                    AMARILLA_CONTRAMARCAS, CODIGO_A_GRUPO)


# ══════════════════════════════════════════════════════════════════════
# PIEZAS ESPERADAS POR MEDIA RES (anatómicas)
# Una media res = medio animal, contiene 1 de cada corte principal
# ══════════════════════════════════════════════════════════════════════
PIEZAS_POR_MEDIA = {
    'BIFE DE CHORIZO': 1,
    'NALGA SIN TAPA': 1,
    'CUADRADA': 1,
    'PECETO': 1,
    'TAPA DE BIFE': 1,
    'BOLA DE LOMO': 1,
    'CUADRIL': 1,
    'COLITA': 1,
    'TAPA DE CUADRIL': 1,
    'LOMO CON CORDON': 1,
    'ROASTBEEF': 1,
    'PALETA': 1,
    'TAPA DE ASADO': 1,
    'FALDA': 1,
    'ENTRAÑA': 1,
    'VACIO': 1,
    'MATAMBRE': 1,
    'ASADO': 1,
    'OJO DE BIFE': 1,
    'TORTUGUITA': 1,
    'TAPA DE NALGA': 1,
}

# Ratio aproximado: cuántas piezas porcionadas/feteadas salen de 1 pieza anatómica
RATIO_PORCIONADO = {
    'BIFE DE CHORIZO': 4,    # 1 bife entero → ~4 porciones
    'NALGA SIN TAPA': 6,     # 1 nalga → ~6 feteadas
    'CUADRADA': 6,           # 1 cuadrada → ~6 feteadas
    'PECETO': 3,             # 1 peceto → ~3 porciones
    'TAPA DE BIFE': 3,       # 1 marucha → ~3 porciones
    'BOLA DE LOMO': 6,       # 1 bola → ~6 feteadas
    'CUADRIL': 3,            # 1 cuadril → ~3 churrascos
    'COLITA': 2,             # 1 colita → ~2 porciones
    'TAPA DE CUADRIL': 2,    # 1 tapa/picanha → ~2 porciones
    'LOMO CON CORDON': 3,    # 1 lomo → ~3 porciones
    'OJO DE BIFE': 5,        # 1 ojo → ~5 porcionados
    'TAPA DE NALGA': 3,      # 1 tapa nalga → ~3 porcionadas
    'VACIO': 2,              # 1 vacío → ~2 porciones
    'MATAMBRE': 2,           # 1 matambre → ~2 porciones
    'ENTRAÑA': 2,            # 1 entraña → ~2 porciones
    'ASADO': 3,              # 1 asado → ~3 tiras/planchas
    'ROASTBEEF': 1,          # generalmente sale entero
    'PALETA': 1,             # generalmente sale entera
    'TAPA DE ASADO': 1,      # generalmente sale entera
    'FALDA': 1,              # generalmente sale entera
    'TORTUGUITA': 1,
}


def clasificar_corte(desc, codigo=None):
    """Mapea texto de corte (y opcionalmente código numérico) a grupo de precio.
    El código tiene prioridad porque las descripciones suelen estar truncadas."""
    # 1) Match por código (lo más confiable)
    if codigo:
        cod_str = str(codigo).strip()
        if cod_str in CODIGO_A_GRUPO:
            return CODIGO_A_GRUPO[cod_str]

    desc_upper = (desc or '').upper().strip()

    # 2) Subcortes especiales por nombre — chequear ANTES que sus parents
    if ('CEJA DE OJO DE BIFE' in desc_upper or 'CEJA OJO DE BIFE' in desc_upper
            or 'CEJA DE BIFE' in desc_upper or 'TAPA DE OJO DE BIFE' in desc_upper
            or 'TAPA OJO DE BIFE' in desc_upper):
        return 'CEJA DE BIFE'
    if 'BIFE DE VACIO' in desc_upper or 'BIFE VACIO' in desc_upper:
        return 'BIFE DE VACIO'
    if ('MEDIALUNA DE VACIO' in desc_upper or 'MEDIA LUNA DE VACIO' in desc_upper
            or 'MEDIA-LUNA DE VACIO' in desc_upper or 'MEDIALUNA VACIO' in desc_upper):
        return 'MEDIALUNA DE VACIO'
    if 'ASADO DEL CENTRO' in desc_upper:
        return 'ASADO DEL CENTRO'
    if 'ASADO BANDERITA' in desc_upper or 'BANDERITA' in desc_upper:
        return 'ASADO BANDERITA'
    # ASADO PLANCHA subcorte SOLO si NO tiene "C/H" / "C H" (con hueso) — esos
    # van al parent ASADO.
    if (('ASADO PLANCHA' in desc_upper or 'ASADO FRES PLANCHA' in desc_upper
            or 'ASADO FRESCO PLANCHA' in desc_upper)
            and 'C/H' not in desc_upper and 'C H ' not in desc_upper
            and 'C/HUESO' not in desc_upper and 'CON HUESO' not in desc_upper):
        return 'ASADO PLANCHA'
    if ('ASADO EN TIRAS' in desc_upper or 'ASADO TIRAS' in desc_upper
            or 'ASADO FRES TIRAS' in desc_upper or 'ASADO FRESCO TIRAS' in desc_upper
            or 'ASADO TIRA' in desc_upper or 'ASADO FRES TIRA' in desc_upper):
        return 'ASADO EN TIRAS'
    if 'MARUCHA' in desc_upper:
        return 'MARUCHA'

    # 2b) Variantes específicas con manejo especial
    # Bife Angosto: variante "sin hueso" o sola → BIFE DE CHORIZO.
    # "Bife Angosto Con Hueso" / "Tomahawk" → SIN CLASIFICAR (premium con hueso).
    if 'BIFE ANGOSTO' in desc_upper:
        if ('CON HUESO' in desc_upper or 'C/H' in desc_upper
                or ' C H ' in f' {desc_upper} ' or 'C/HUESO' in desc_upper):
            return 'SIN CLASIFICAR'
        return 'BIFE DE CHORIZO'
    if 'TOMAHAWK' in desc_upper:
        return 'SIN CLASIFICAR'
    # Cuadril plain o "Cuadril sin tapa" → CUADRIL (no confundir con TAPA DE CUADRIL)
    if ('CUADRIL' in desc_upper and 'COLITA' not in desc_upper
            and ('TAPA' not in desc_upper or 'SIN TAPA' in desc_upper)
            and 'PICANHA' not in desc_upper):
        return 'CUADRIL'

    # 3) Prioridad para evitar match parcial sobre cortes principales
    priority_checks = [
        'TAPA DE CUADRIL', 'TAPA DE NALGA', 'TAPA DE BIFE', 'TAPA DE ASADO',
        'BIFE DE CHORIZO', 'OJO DE BIFE', 'BOLA DE LOMO', 'NALGA SIN TAPA',
        'CARNE PICADA', 'CHURRASCOS DE CUADRIL', 'COLITA DE CUADRIL',
        'ASADO C/H', 'ASADO SIN HUESO',
    ]
    for check in priority_checks:
        if check in desc_upper:
            if 'ASADO' in check and 'TAPA' not in check:
                return 'ASADO'
            if 'COLITA' in check:
                return 'COLITA'
            for keywords, grupo in CORTE_TO_GRUPO:
                if check in [k.upper() for k in keywords]:
                    return grupo

    if 'PICANHA' in desc_upper:
        return 'TAPA DE CUADRIL'
    if 'BIFE AMERICANO' in desc_upper:
        return 'TAPA DE BIFE'

    # 4) Matching general
    for keywords, grupo in CORTE_TO_GRUPO:
        for kw in keywords:
            if kw.upper() in desc_upper:
                if grupo == 'NALGA SIN TAPA' and 'TAPA DE' in desc_upper:
                    continue
                if grupo == 'CUADRIL' and ('COLITA' in desc_upper or 'TAPA' in desc_upper):
                    continue
                if grupo == 'ASADO' and 'TAPA' in desc_upper:
                    continue
                return grupo
    return 'SIN CLASIFICAR'


def clasificar_tipo(desc):
    """PORCIONADO / FETEADO / ANATÓMICO."""
    d = desc.upper()
    if 'PORC' in d or 'PORCION' in d:
        return 'PORCIONADO'
    if 'FETEAD' in d:
        return 'FETEADO'
    return 'ANATÓMICO'


def extraer_contramarca(nro_venta):
    """Extrae número de contramarca del Nro.Venta."""
    if not nro_venta:
        return ''
    nro_venta = str(nro_venta).strip()
    m = re.match(r'(\d+)\s*-', nro_venta)
    if m:
        return m.group(1)
    return ''


def resolver_cliente(contramarca):
    """Devuelve (nombre_cliente, columna_precios) para una contramarca."""
    cm = str(contramarca).strip()
    if cm in AMARILLA_CONTRAMARCAS:
        return 'AMARILLA', 'AMARILLA'
    if cm in CONTRAMARCA_MAP:
        return CONTRAMARCA_MAP[cm]
    return 'RESTO CLIENTES', 'RESTO CLIENTES AMBA'


def _parse_kg(s):
    """Convierte string de kg a float. Maneja '1,842.00' y '17,426.00'."""
    s = s.strip()
    # Formato: 1,842.00 → 1842.00  o  17,426.00 → 17426.00
    s = s.replace(',', '')
    return float(s)


def parse_romaneo_pdf(pdf_path):
    """
    Parsea un PDF de romaneo y devuelve un dict con:
    - numero, fecha, medias_reses, kg_entrada, categoria, tipificacion
    - cortes: lista de dicts
    - grasa_kg, merma_kg
    """
    import pdfplumber

    all_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text.append(text)

    full_text = '\n'.join(all_text)
    lines = full_text.split('\n')

    result = {
        'numero': '',
        'fecha': '',
        'medias_reses': 0,
        'kg_entrada': 0,
        'categoria': '',
        'tipificacion': '',
        'cortes': [],
        'grasa_kg': 0,
        'merma_kg': 0,
    }

    # Detectar formato
    es_resultado = 'Resultado Despostada' in full_text

    # Extraer número de romaneo: "Usuario 215" o "N° 215"
    m = re.search(r'Usuario\s+(\d+)\s+al\s+\d+', full_text)
    if m:
        result['numero'] = m.group(1)
    else:
        m = re.search(r'(?:Romaneo|N[°º])\s*:?\s*(\d+)', full_text)
        if m:
            result['numero'] = m.group(1)

    # Si no encontró número, intentar extraerlo del nombre de archivo
    if not result['numero']:
        fname = str(pdf_path).split('/')[-1]
        m = re.search(r'(\d{3,})', fname)
        if m:
            result['numero'] = m.group(1)

    # Extraer fecha del encabezado (primera fecha que aparece)
    m = re.search(r'(\d{2}/\d{2}/\d{4})', full_text)
    if m:
        result['fecha'] = m.group(1)

    # ── ENTRADA ──
    # Buscar "Total >>>" dentro de la sección "Entrada Despostada"
    # El primer Total >>> es siempre el de entrada
    entrada_found = False
    for i, line in enumerate(lines):
        if 'Entrada Despostada' in line:
            entrada_found = True
        if entrada_found and 'Total >>>' in line:
            # Formato: "Total >>> 140 17,426.00" o "Total >>> 143 19,015.00"
            m = re.search(r'Total\s*>>>\s+([\d,]+)\s+([\d,.]+)', line)
            if m:
                result['medias_reses'] = int(m.group(1).replace(',', ''))
                result['kg_entrada'] = _parse_kg(m.group(2))
            break

    # ── CATEGORÍA ──
    # Buscar tipificaciones en la sección de entrada (entre "Entrada" y primer "Total >>>")
    in_entrada = False
    tip_kg = {}
    for line in lines:
        if 'Entrada Despostada' in line:
            in_entrada = True
            continue
        if in_entrada and 'Total >>>' in line:
            break
        if in_entrada:
            for tip_code in ['VA', 'NO', 'NT', 'VQ', 'TO', 'BU', 'BB']:
                # Buscar tip code como palabra suelta en la línea
                if re.search(rf'\b{tip_code}\b', line):
                    # Extraer kg de esta línea de entrada
                    m_kg = re.search(r'(\d+)\s+([\d,.]+)\s*$', line)
                    if m_kg:
                        kg_val = _parse_kg(m_kg.group(2))
                        tip_kg[tip_code] = tip_kg.get(tip_code, 0) + kg_val

    if tip_kg:
        main_tip = max(tip_kg, key=tip_kg.get)
        result['tipificacion'] = main_tip
        result['categoria'] = TIPIFICACION_MAP.get(main_tip, 'Vaca')

        # Desglose por categoría (% de kg)
        total_tip_kg = sum(tip_kg.values())
        result['desglose_categoria'] = {}
        for tip, kg in tip_kg.items():
            cat_name = TIPIFICACION_MAP.get(tip, tip)
            pct = (kg / total_tip_kg * 100) if total_tip_kg > 0 else 0
            if cat_name in result['desglose_categoria']:
                result['desglose_categoria'][cat_name]['kg'] += kg
                result['desglose_categoria'][cat_name]['pct'] = (
                    result['desglose_categoria'][cat_name]['kg'] / total_tip_kg * 100)
            else:
                result['desglose_categoria'][cat_name] = {
                    'kg': kg, 'pct': round(pct, 1), 'tip': tip
                }

        # Desglose por tipificación (D0, C1, C2, etc.) — buscar en líneas de entrada
        result['desglose_tipificacion'] = {}
        in_ent2 = False
        for line in lines:
            if 'Entrada Despostada' in line:
                in_ent2 = True; continue
            if in_ent2 and 'Total >>>' in line:
                break
            if in_ent2:
                # Buscar códigos de tipificación como D0, C1, C2, J1, etc.
                for tip_clase in re.findall(r'\b([A-Z][0-9])\b', line):
                    m_kg2 = re.search(r'(\d+)\s+([\d,.]+)\s*$', line)
                    if m_kg2:
                        kg_v = _parse_kg(m_kg2.group(2))
                        result['desglose_tipificacion'][tip_clase] = (
                            result['desglose_tipificacion'].get(tip_clase, 0) + kg_v)
    else:
        result['categoria'] = 'Vaca'
        result['desglose_categoria'] = {}
        result['desglose_tipificacion'] = {}

    # ── SALIDAS (cortes) ──
    in_salida = False
    in_subproducto = False

    for line in lines:
        stripped = line.strip()

        # Detectar secciones
        if 'Salidas Despostada' in stripped or 'Salida Despostada' in stripped:
            in_salida = True
            in_subproducto = False
            continue

        if 'Sub-Producto' in stripped or 'Sub Producto' in stripped:
            in_salida = False
            in_subproducto = True
            continue

        if re.match(r'-{10,}', stripped) or re.match(r'={10,}', stripped):
            continue

        if 'Total >>>' in stripped:
            continue

        if 'Total SALIDAS' in stripped or 'Total INGRESOS' in stripped:
            continue

        if 'Merma >>>' in stripped or 'Merma >>' in stripped:
            m = re.search(r'([\d,.]+)\s+[\d.]+\s*$', stripped)
            if m:
                result['merma_kg'] = _parse_kg(m.group(1))
            else:
                m = re.search(r'([\d,.]+)\s*$', stripped)
                if m:
                    result['merma_kg'] = _parse_kg(m.group(1))
            continue

        # ── Parsear línea de corte ──
        if in_salida:
            corte = _parse_linea_corte(stripped, es_resultado)
            if corte:
                result['cortes'].append(corte)

        # ── Grasa y decomiso ──
        if in_subproducto and 'GRASA Y DECOMISO' in stripped.upper():
            m = re.search(r'(\d+)\s+(\d+)\s+([\d,.]+)', stripped)
            if m:
                grasa_kg = _parse_kg(m.group(3))
                result['grasa_kg'] = grasa_kg
                result['cortes'].append({
                    'corte': 'GRASA Y DECOMISO',
                    'grupo': 'GRASA',
                    'tipo': 'SUBPRODUCTO',
                    'piezas': int(m.group(1)),
                    'unidades': int(m.group(2)),
                    'kg': grasa_kg,
                    'destino': 'ENFRI',
                    'contramarca': '',
                    'cliente': '',
                    'nro_venta': '',
                })

    # Texto fuente (encabezado + cuerpo) para detectar faenador / planta de
    # desposte y segmentar el canal (China = DELTACAR / TOP MEAT, etc.).
    result['texto_fuente'] = full_text[:3000]

    return result


def _parse_linea_corte(line, es_resultado):
    """
    Parsea una línea de corte individual.
    Soporta múltiples variantes de formato:
    - TF CA, TF BL, BUBAL, u otras marcas
    - Contramarca al final (Formato 1) o antes de piezas (Formato 2)
    - Contramarca ausente (solo "-")
    """
    if not line or line.startswith('-') or line.startswith('='):
        return None
    if not re.match(r'\d+\s', line):
        return None
    if 'Hoja:' in line or 'Total' in line:
        return None

    # Extraer destino (ENFRI/CONGE/REFRI)
    m_dest = re.search(r'\b(ENFRI|CONGE|REFRI)\b', line)
    if not m_dest:
        return None
    destino = m_dest.group(1)

    # Todo antes del destino = codigo + descripción
    pre_destino = line[:m_dest.start()].strip()
    m_code = re.match(r'(\d+)\s+(.+)', pre_destino)
    if not m_code:
        return None
    codigo = m_code.group(1).strip()
    desc = m_code.group(2).strip()
    # Limpiar desc: sacar ARGENTINA si quedó al final
    desc = re.sub(r'\s+ARGENTINA\s*$', '', desc).strip()

    # Todo después del destino = marca + datos numéricos
    post_destino = line[m_dest.end():].strip()

    # Intentar extraer datos numéricos del final de la línea
    # Buscamos: PIEZAS UNIDADES KG RENDI% [CTM- NUM | -]
    # Los números están al final, la marca/texto al principio del post_destino

    if es_resultado:
        # Formato 2: marca CTM-NUM PIEZAS UNI KG RENDI%
        m = re.search(
            r'(\d+)-([\d]+)\s+'      # contramarca-numero
            r'([\d,]+)\s+'           # piezas
            r'(\d+)\s+'             # unidades
            r'([\d,.]+)\s+'          # kg
            r'([\d.]+)\s*$',         # rendimiento
            post_destino
        )
        if m:
            ctm = m.group(1)
            nro_venta = f"{m.group(1)}-{m.group(2)}"
            piezas = int(m.group(3).replace(',', ''))
            unidades = int(m.group(4))
            kg = _parse_kg(m.group(5))
        else:
            return None
    else:
        # Formato 1: marca PIEZAS UNI KG RENDI% CTM- NUM
        # Primero intentar con contramarca al final
        m = re.search(
            r'([\d,]+)\s+'           # piezas
            r'(\d+)\s+'             # unidades
            r'([\d,.]+)\s+'          # kg
            r'[\d.]+\s+'            # rendimiento
            r'(\d+)-\s*(\d+)\s*$',   # contramarca - numero
            post_destino
        )
        if m:
            piezas = int(m.group(1).replace(',', ''))
            unidades = int(m.group(2))
            kg = _parse_kg(m.group(3))
            ctm = m.group(4)
            nro_venta = f"{m.group(4)}- {m.group(5)}"
        else:
            # Sin contramarca (solo "-" al final)
            m = re.search(
                r'([\d,]+)\s+'       # piezas
                r'(\d+)\s+'         # unidades
                r'([\d,.]+)\s+'      # kg
                r'[\d.]+\s+'        # rendimiento
                r'-\s*$',            # guión solo
                post_destino
            )
            if m:
                piezas = int(m.group(1).replace(',', ''))
                unidades = int(m.group(2))
                kg = _parse_kg(m.group(3))
                ctm = ''
                nro_venta = ''
            else:
                return None

    contramarca = ctm
    cliente, _ = resolver_cliente(contramarca)
    grupo = clasificar_corte(desc, codigo=codigo)
    tipo = clasificar_tipo(desc)

    # Detectar bubalino por marca "BUBAL" en la línea
    es_bubalino = 'BUBAL' in line.upper()

    return {
        'corte': desc,
        'codigo': codigo,
        'grupo': grupo,
        'tipo': tipo,
        'piezas': piezas,
        'unidades': unidades,
        'kg': kg,
        'destino': destino,
        'contramarca': contramarca,
        'cliente': cliente,
        'nro_venta': nro_venta,
        'es_bubalino': es_bubalino,
    }


def control_cortes(parsed_data):
    """
    Control de cortes con proporción anatómico/porcionado/feteado.

    Lógica nueva:
    - Si un grupo tiene SOLO anatómico: esperadas = medias × 1
    - Si tiene SOLO porcionado/feteado: esperadas = medias × ratio
    - Si tiene MIX: calcular proporción de medias que fue a cada tipo
      basado en kg, y esperar piezas proporcionalmente.
      Ej: 10 medias, 5 kg anat + 5 kg porc (50/50) →
          esperar 5 anat + (5 × ratio) porc

    Además: peso promedio por pieza porcionada/feteada (target 1-1.5 kg).
    """
    medias = parsed_data.get('medias_reses', 0)
    if medias == 0:
        return []

    cortes = parsed_data.get('cortes', [])
    from collections import defaultdict
    agrupado = defaultdict(lambda: {'ANATÓMICO': 0, 'PORCIONADO': 0, 'FETEADO': 0,
                                     'kg_anat': 0, 'kg_porc': 0, 'kg_fet': 0})
    for c in cortes:
        if c.get('grupo') in ('GRASA', 'SIN CLASIFICAR', 'RECORTE 70-30',
                               'RECORTE 80-20', 'RECORTE 90-10', 'CARNE PICADA'):
            continue
        grupo = c['grupo']
        tipo = c.get('tipo', 'ANATÓMICO')
        pzas = c.get('piezas', 0)
        kg = c.get('kg', 0)
        if tipo == 'ANATÓMICO':
            agrupado[grupo]['ANATÓMICO'] += pzas
            agrupado[grupo]['kg_anat'] += kg
        elif tipo == 'PORCIONADO':
            agrupado[grupo]['PORCIONADO'] += pzas
            agrupado[grupo]['kg_porc'] += kg
        elif tipo == 'FETEADO':
            agrupado[grupo]['FETEADO'] += pzas
            agrupado[grupo]['kg_fet'] += kg

    resultado = []

    for grupo in PIEZAS_POR_MEDIA:
        data = agrupado.get(grupo, {'ANATÓMICO': 0, 'PORCIONADO': 0, 'FETEADO': 0,
                                     'kg_anat': 0, 'kg_porc': 0, 'kg_fet': 0})
        pzas_anat = data['ANATÓMICO']
        pzas_porc = data['PORCIONADO']
        pzas_fet = data['FETEADO']
        kg_anat = data['kg_anat']
        kg_porc = data['kg_porc']
        kg_fet = data['kg_fet']
        kg_total = kg_anat + kg_porc + kg_fet

        if pzas_anat == 0 and pzas_porc == 0 and pzas_fet == 0:
            # Sin registro
            resultado.append({
                'grupo': grupo,
                'tipo': 'SIN REGISTRO',
                'piezas_reales': 0,
                'piezas_esperadas': medias,
                'diferencia': -medias,
                'ratio_por_media': 0,
                'ratio_esperado': 1,
                'peso_promedio': 0,
                'peso_ideal': '',
                'alerta': 'SIN REGISTRO',
                'detalle': f'No salieron {grupo.lower()} ni anatómicos ni porcionados',
            })
            continue

        ratio_esp = RATIO_PORCIONADO.get(grupo, 1)
        # Proporción de medias destinadas a cada tipo (basado en kg)
        if kg_total > 0:
            prop_anat = kg_anat / kg_total
            prop_porc = kg_porc / kg_total
            prop_fet = kg_fet / kg_total
        else:
            prop_anat = prop_porc = prop_fet = 0

        # Anatómico
        if pzas_anat > 0:
            esperadas_anat = round(medias * prop_anat)
            diff = pzas_anat - esperadas_anat
            tol = max(1, esperadas_anat * 0.05)
            if diff < -tol:
                alerta = f'FALTAN {abs(diff):.0f}'
            elif diff > tol:
                alerta = f'SOBRAN {diff:.0f}'
            else:
                alerta = 'OK'
            resultado.append({
                'grupo': grupo,
                'tipo': 'ANATÓMICO',
                'piezas_reales': pzas_anat,
                'piezas_esperadas': esperadas_anat,
                'diferencia': diff,
                'ratio_por_media': round(pzas_anat / medias, 2),
                'ratio_esperado': 1,
                'peso_promedio': round(kg_anat / pzas_anat, 2) if pzas_anat > 0 else 0,
                'peso_ideal': '-',
                'alerta': alerta,
                'detalle': (f'{pzas_anat} pzas de {round(medias * prop_anat)} esperadas '
                            f'({prop_anat*100:.0f}% de medias fue anatómico)'),
            })

        # Porcionado
        if pzas_porc > 0:
            # Las medias que fueron a porcionado × ratio
            medias_a_porc = medias * prop_porc
            esperadas_porc = round(medias_a_porc * ratio_esp)
            diff_porc = pzas_porc - esperadas_porc
            ratio_real = pzas_porc / medias_a_porc if medias_a_porc > 0 else 0
            peso_prom = kg_porc / pzas_porc if pzas_porc > 0 else 0
            tol_p = max(1, esperadas_porc * 0.25)
            if diff_porc < -tol_p:
                alerta_p = f'FALTAN {abs(diff_porc):.0f}'
            elif diff_porc > tol_p:
                alerta_p = f'SOBRAN {diff_porc:.0f}'
            else:
                alerta_p = 'OK'

            # Evaluar peso
            if peso_prom > 1.5:
                peso_nota = f'PESADO ({peso_prom:.2f} kg)'
            elif peso_prom < 1.0:
                peso_nota = f'LIVIANO ({peso_prom:.2f} kg)'
            else:
                peso_nota = f'OK ({peso_prom:.2f} kg)'

            resultado.append({
                'grupo': grupo,
                'tipo': 'PORCIONADO',
                'piezas_reales': pzas_porc,
                'piezas_esperadas': esperadas_porc,
                'diferencia': diff_porc,
                'ratio_por_media': round(ratio_real, 1),
                'ratio_esperado': ratio_esp,
                'peso_promedio': round(peso_prom, 2),
                'peso_ideal': '1.0-1.5 kg',
                'peso_nota': peso_nota,
                'alerta': alerta_p,
                'detalle': (f'{pzas_porc} porc ({kg_porc:.0f} kg) de {esperadas_porc} esperadas '
                            f'({prop_porc*100:.0f}% de medias fue porcionado × ratio {ratio_esp})'),
            })

        # Feteado
        if pzas_fet > 0:
            medias_a_fet = medias * prop_fet
            esperadas_fet = round(medias_a_fet * ratio_esp)
            diff_fet = pzas_fet - esperadas_fet
            ratio_real_f = pzas_fet / medias_a_fet if medias_a_fet > 0 else 0
            peso_prom_f = kg_fet / pzas_fet if pzas_fet > 0 else 0
            tol_f = max(1, esperadas_fet * 0.25)
            if diff_fet < -tol_f:
                alerta_f = f'FALTAN {abs(diff_fet):.0f}'
            elif diff_fet > tol_f:
                alerta_f = f'SOBRAN {diff_fet:.0f}'
            else:
                alerta_f = 'OK'

            if peso_prom_f > 1.5:
                peso_nota_f = f'PESADO ({peso_prom_f:.2f} kg)'
            elif peso_prom_f < 1.0:
                peso_nota_f = f'LIVIANO ({peso_prom_f:.2f} kg)'
            else:
                peso_nota_f = f'OK ({peso_prom_f:.2f} kg)'

            resultado.append({
                'grupo': grupo,
                'tipo': 'FETEADO',
                'piezas_reales': pzas_fet,
                'piezas_esperadas': esperadas_fet,
                'diferencia': diff_fet,
                'ratio_por_media': round(ratio_real_f, 1),
                'ratio_esperado': ratio_esp,
                'peso_promedio': round(peso_prom_f, 2),
                'peso_ideal': '1.0-1.5 kg',
                'peso_nota': peso_nota_f,
                'alerta': alerta_f,
                'detalle': (f'{pzas_fet} fet ({kg_fet:.0f} kg) de {esperadas_fet} esperadas '
                            f'({prop_fet*100:.0f}% de medias fue feteado × ratio {ratio_esp})'),
            })

    return resultado


def detectar_cortes_faltantes(parsed_data):
    """
    Detecta cortes faltantes en un romaneo.
    Retorna lista de {grupo, es_caro, alerta}.
    Los cortes CAROS faltantes son alerta roja.
    """
    from config import CORTES_CAROS, GRUPOS_POR_CALIDAD

    grupos_presentes = set()
    for c in parsed_data.get('cortes', []):
        if c.get('grupo') != 'GRASA':
            grupos_presentes.add(c['grupo'])

    todos = [g for g, _ in GRUPOS_POR_CALIDAD['Standard']
             if g not in ('RECORTE 80-20', 'RECORTE 90-10')]  # estos pueden faltar

    faltantes = []
    for g in todos:
        if g not in grupos_presentes:
            es_caro = g in CORTES_CAROS
            faltantes.append({
                'grupo': g,
                'es_caro': es_caro,
                'alerta': 'CORTE CARO FALTANTE' if es_caro else 'faltante',
            })
    return faltantes


def es_correccion(nombre_archivo):
    """Detecta si un archivo es una corrección de otro romaneo."""
    import re
    n = nombre_archivo.upper()
    if 'CORREC' in n or 'CORRECCIÓN' in n or 'CORRECCION' in n:
        # Extraer el nombre base (sin la parte de corrección)
        base = re.sub(r'\s*\(?\s*CORREC[CIÓN]*\s*\)?\s*', '', nombre_archivo, flags=re.IGNORECASE).strip()
        return True, base
    return False, nombre_archivo


def detectar_tipo_pdf(pdf_path):
    """
    Detecta el tipo de PDF:
    - 'romaneo': Rendimientos/Resultado de Despostada (análisis completo)
    - 'remanejo': Remanejo de cortes (entrada cortes varios, salida cortes específicos)
    - 'entrada': Romaneo de ENTRADA (detalle de medias, sin salidas)
    - 'desconocido': no se pudo determinar
    """
    import pdfplumber
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ''

        if 'Romaneo de ENTRADA' in text:
            return 'entrada'

        # Remanejo: entrada tiene "CORTES VARIOS" en vez de "1/2 RES"
        if 'CORTES VARIOS' in text:
            return 'remanejo'

        if 'Rendimiento' in text and 'Despostada' in text:
            return 'romaneo'
        if 'Resultado Despostada' in text:
            return 'romaneo'

        return 'desconocido'
    except Exception:
        return 'desconocido'


def parse_remanejo_pdf(pdf_path):
    """
    Parsea un PDF de remanejo.
    Entrada: CORTES VARIOS con kg totales.
    Salida: cortes específicos con kg.
    Calcula merma = entrada - salida.
    """
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        text = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    lines = text.split('\n')
    result = {
        'tipo_pdf': 'remanejo',
        'fecha': '',
        'kg_entrada': 0,
        'piezas_entrada': 0,
        'cortes_salida': [],
        'kg_salida': 0,
        'merma_kg': 0,
        'merma_pct': 0,
    }

    # Fecha
    m = re.search(r'(\d{2}/\d{2}/\d{4})', text)
    if m:
        result['fecha'] = m.group(1)

    # Entrada: "CORTES VARIOS" lines then Total >>>
    in_entrada = False
    for line in lines:
        if 'Entrada Despostada' in line:
            in_entrada = True
            continue
        if in_entrada and 'Total >>>' in line:
            m = re.search(r'Total\s*>>>\s+([\d,]+)\s+([\d,.]+)', line)
            if m:
                result['piezas_entrada'] = int(m.group(1).replace(',', ''))
                result['kg_entrada'] = _parse_kg(m.group(2))
            break

    # Salida: cortes normales
    in_salida = False
    for line in lines:
        stripped = line.strip()
        if 'Salidas Despostada' in stripped or 'Salida Despostada' in stripped:
            in_salida = True
            continue
        if in_salida and 'Total >>>' in stripped:
            m = re.search(r'Total\s*>>>\s+[\d,]+\s+\d+\s+([\d,.]+)', stripped)
            if m:
                result['kg_salida'] = _parse_kg(m.group(1))
            break
        if in_salida and not stripped.startswith('-') and not stripped.startswith('='):
            # Intentar parsear línea de corte
            corte = _parse_linea_corte(stripped, False)
            if corte:
                result['cortes_salida'].append(corte)

    # Si no sacamos kg_salida del total, sumar de los cortes
    if result['kg_salida'] == 0 and result['cortes_salida']:
        result['kg_salida'] = sum(c['kg'] for c in result['cortes_salida'])

    # Merma
    m_merma = re.search(r'Merma\s*>>>\s+([\d,.]+)', text)
    if m_merma:
        result['merma_kg'] = _parse_kg(m_merma.group(1))
    else:
        result['merma_kg'] = result['kg_entrada'] - result['kg_salida']

    if result['kg_entrada'] > 0:
        result['merma_pct'] = round(result['merma_kg'] / result['kg_entrada'] * 100, 2)

    return result


def parse_multiple_pdfs(pdf_paths):
    """Parsea múltiples PDFs y retorna lista de resultados."""
    results = []
    for path in pdf_paths:
        try:
            r = parse_romaneo_pdf(path)
            r['archivo'] = str(path).split('/')[-1]
            results.append(r)
        except Exception as e:
            results.append({
                'archivo': str(path).split('/')[-1],
                'error': str(e),
            })
    return results


def acumular_romaneos(parsed_list):
    """
    Combina múltiples romaneos parseados en uno acumulado.
    Cada romaneo aporta su propio precio de compra — el acumulado
    calcula el costo total real y el precio promedio ponderado.
    """
    acum = {
        'numero': 'ACUMULADO',
        'fecha': '',
        'medias_reses': 0,
        'kg_entrada': 0,
        'categoria': '',
        'tipificacion': '',
        'cortes': [],
        'grasa_kg': 0,
        'merma_kg': 0,
        'archivos_incluidos': [],
        'costo_hacienda_total': 0,
        'detalle_compras': [],
    }

    cat_counts = {}
    fechas = []

    for p in parsed_list:
        if 'error' in p:
            continue
        kg_ent = p.get('kg_entrada', 0)
        precio = p.get('precio_compra', 0)
        costo = kg_ent * precio

        acum['medias_reses'] += p.get('medias_reses', 0)
        acum['kg_entrada'] += kg_ent
        acum['cortes'].extend(p.get('cortes', []))
        acum['grasa_kg'] += p.get('grasa_kg', 0)
        acum['merma_kg'] += p.get('merma_kg', 0)
        acum['archivos_incluidos'].append(p.get('archivo', ''))
        acum['costo_hacienda_total'] += costo
        acum['detalle_compras'].append({
            'archivo': p.get('archivo', ''),
            'kg_entrada': kg_ent,
            'precio_compra': precio,
            'costo': costo,
            'categoria': p.get('categoria', ''),
            'medias': p.get('medias_reses', 0),
        })

        cat = p.get('categoria', '')
        if cat:
            cat_counts[cat] = cat_counts.get(cat, 0) + kg_ent
        if p.get('fecha'):
            fechas.append(p['fecha'])

    if cat_counts:
        acum['categoria'] = max(cat_counts, key=cat_counts.get)
    if fechas:
        acum['fecha'] = f"{min(fechas)} a {max(fechas)}"

    # Precio promedio ponderado por kg
    if acum['kg_entrada'] > 0:
        acum['precio_compra'] = round(acum['costo_hacienda_total'] / acum['kg_entrada'])
    else:
        acum['precio_compra'] = 0

    return acum
