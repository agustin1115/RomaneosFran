"""
pdf_parser.py — Parsea PDFs de romaneo (Rendimientos/Resultado de Despostada).
Extrae encabezado, líneas de cortes, grasa/decomiso y merma.

Formatos soportados:
  - "Rendimientos de Despostada": ctm al final como "47- 215"
  - "Resultado Despostada": ctm antes de piezas como "47-200215"
"""
import re
from config import CORTE_TO_GRUPO, TIPIFICACION_MAP, CONTRAMARCA_MAP, AMARILLA_CONTRAMARCAS


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


def clasificar_corte(desc):
    """Mapea texto de corte a grupo de precio."""
    desc_upper = desc.upper().strip()

    # Prioridad especial para evitar match parcial
    priority_checks = [
        'TAPA DE CUADRIL', 'TAPA DE NALGA', 'TAPA DE BIFE', 'TAPA DE ASADO',
        'BIFE DE CHORIZO', 'OJO DE BIFE', 'BOLA DE LOMO', 'NALGA SIN TAPA',
        'CARNE PICADA', 'CHURRASCOS DE CUADRIL', 'COLITA DE CUADRIL',
        'ASADO C/H', 'ASADO EN TIRAS', 'ASADO SIN HUESO', 'ASADO DEL CENTRO',
    ]
    for check in priority_checks:
        if check in desc_upper:
            # Map special asado variants
            if 'ASADO' in check and 'TAPA' not in check:
                return 'ASADO'
            if 'COLITA' in check:
                return 'COLITA'
            for keywords, grupo in CORTE_TO_GRUPO:
                if check in [k.upper() for k in keywords]:
                    return grupo

    # Picanha → TAPA DE CUADRIL
    if 'PICANHA' in desc_upper:
        return 'TAPA DE CUADRIL'

    # Marucha / bife americano → TAPA DE BIFE
    if 'MARUCHA' in desc_upper or 'BIFE AMERICANO' in desc_upper:
        return 'TAPA DE BIFE'

    # General matching
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

    return result


def _parse_linea_corte(line, es_resultado):
    """
    Parsea una línea de corte individual.

    Formato 1 (Rendimientos):
      110 CARNE PICADA ARGENTINA ENFRI TF CA 1,805 101 1,842.00 10.57 87- 215
      Patrón: CODIGO DESC ARGENTINA DESTINO TF CA PIEZAS UNI KG RENDI% CTM- NUM

    Formato 2 (Resultado):
      530 BIFE DE CHORIZO ARGENTINA CONGE TF CA 47-200215 39 8 169.30 0.89
      Patrón: CODIGO DESC ARGENTINA DESTINO TF CA CTM-NUMNUM PIEZAS UNI KG RENDI%
    """
    if not line or line.startswith('-') or line.startswith('='):
        return None

    # Debe empezar con un código numérico
    if not re.match(r'\d+\s', line):
        return None

    # Ignorar líneas de "Hoja:" (encabezados de página)
    if 'Hoja:' in line:
        return None

    if es_resultado:
        # Formato 2: ctm ANTES de piezas
        # 530 BIFE DE CHORIZO ARGENTINA CONGE TF CA 47-200215 39 8 169.30 0.89
        m = re.match(
            r'(\d+)\s+'                           # codigo
            r'(.+?)\s+'                            # descripción
            r'(?:ARGENTINA\s+)?'                   # ARGENTINA (opcional)
            r'(ENFRI|CONGE|REFRI)\s+'              # destino
            r'(?:TF\s+CA\s+)?'                     # TF CA (opcional)
            r'(\d+)-([\d]+)\s+'                    # contramarca-numero
            r'([\d,]+)\s+'                         # piezas
            r'(\d+)\s+'                            # unidades
            r'([\d,.]+)\s+'                        # kg
            r'([\d.]+)',                           # rendimiento %
            line
        )
        if m:
            desc = m.group(2).strip()
            destino = m.group(3)
            ctm = m.group(4)
            nro_venta = f"{m.group(4)}-{m.group(5)}"
            piezas = int(m.group(6).replace(',', ''))
            unidades = int(m.group(7))
            kg = _parse_kg(m.group(8))

            contramarca = ctm
            cliente, _ = resolver_cliente(contramarca)
            grupo = clasificar_corte(desc)
            tipo = clasificar_tipo(desc)

            return {
                'corte': desc,
                'grupo': grupo,
                'tipo': tipo,
                'piezas': piezas,
                'unidades': unidades,
                'kg': kg,
                'destino': destino,
                'contramarca': contramarca,
                'cliente': cliente,
                'nro_venta': nro_venta,
            }
    else:
        # Formato 1: ctm al FINAL
        # 110 CARNE PICADA ARGENTINA ENFRI TF CA 1,805 101 1,842.00 10.57 87- 215
        m = re.match(
            r'(\d+)\s+'                           # codigo
            r'(.+?)\s+'                            # descripción
            r'(?:ARGENTINA\s+)?'                   # ARGENTINA (opcional)
            r'(ENFRI|CONGE|REFRI)\s+'              # destino
            r'(?:TF\s+CA\s+)?'                     # TF CA (opcional)
            r'([\d,]+)\s+'                         # piezas
            r'(\d+)\s+'                            # unidades
            r'([\d,.]+)\s+'                        # kg
            r'[\d.]+\s+'                           # rendimiento % (ignorar)
            r'(\d+)-\s*(\d+)',                     # contramarca - numero
            line
        )
        if m:
            desc = m.group(2).strip()
            destino = m.group(3)
            piezas = int(m.group(4).replace(',', ''))
            unidades = int(m.group(5))
            kg = _parse_kg(m.group(6))
            ctm = m.group(7)
            nro_venta = f"{m.group(7)}- {m.group(8)}"

            contramarca = ctm
            cliente, _ = resolver_cliente(contramarca)
            grupo = clasificar_corte(desc)
            tipo = clasificar_tipo(desc)

            return {
                'corte': desc,
                'grupo': grupo,
                'tipo': tipo,
                'piezas': piezas,
                'unidades': unidades,
                'kg': kg,
                'destino': destino,
                'contramarca': contramarca,
                'cliente': cliente,
                'nro_venta': nro_venta,
            }

    return None


def control_cortes(parsed_data):
    """
    Genera el control de piezas: esperadas vs reales por media res.
    Para anatómicos: N medias → N piezas esperadas.
    Para porcionados/feteados: calcula ratio real vs ratio esperado.

    Retorna lista de dicts con:
      - grupo, tipo, unidades_reales, unidades_esperadas, diferencia,
        ratio_por_media (para porc/fet), ratio_esperado, alerta
    """
    medias = parsed_data.get('medias_reses', 0)
    if medias == 0:
        return []

    cortes = parsed_data.get('cortes', [])

    # Agrupar PIEZAS por grupo y tipo (piezas = cortes individuales, no paquetes)
    from collections import defaultdict
    agrupado = defaultdict(lambda: {'ANATÓMICO': 0, 'PORCIONADO': 0, 'FETEADO': 0,
                                     'kg_anat': 0, 'kg_porc': 0, 'kg_fet': 0})
    for c in cortes:
        if c.get('grupo') in ('GRASA', 'RECORTE 70-30', 'RECORTE 80-20',
                               'RECORTE 90-10', 'CARNE PICADA'):
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
        if grupo not in agrupado and grupo not in ('CARNE PICADA',):
            piezas_esperadas = medias * PIEZAS_POR_MEDIA[grupo]
            resultado.append({
                'grupo': grupo,
                'tipo': 'ANATÓMICO',
                'piezas_reales': 0,
                'piezas_esperadas': piezas_esperadas,
                'diferencia': -piezas_esperadas,
                'ratio_por_media': 0,
                'ratio_esperado': 1,
                'alerta': 'SIN REGISTRO' if piezas_esperadas > 0 else '',
                'detalle': f'Esperado {piezas_esperadas}, salieron 0',
            })
            continue

        data = agrupado.get(grupo, {'ANATÓMICO': 0, 'PORCIONADO': 0, 'FETEADO': 0})
        pzas_anat = data['ANATÓMICO']
        pzas_porc = data['PORCIONADO']
        pzas_fet = data['FETEADO']
        esperadas = medias * PIEZAS_POR_MEDIA[grupo]
        ratio_esp = RATIO_PORCIONADO.get(grupo, 1)

        # Anatómico: control directo (1 pieza por media res)
        if pzas_anat > 0:
            diff = pzas_anat - esperadas
            if diff < -esperadas * 0.05:
                alerta = f'FALTAN {abs(diff):.0f}'
            elif diff > esperadas * 0.05:
                alerta = f'SOBRAN {diff:.0f}'
            else:
                alerta = 'OK'
            resultado.append({
                'grupo': grupo,
                'tipo': 'ANATÓMICO',
                'piezas_reales': pzas_anat,
                'piezas_esperadas': esperadas,
                'diferencia': diff,
                'ratio_por_media': round(pzas_anat / medias, 2),
                'ratio_esperado': 1,
                'alerta': alerta,
                'detalle': f'{pzas_anat} pzas / {medias} medias = {pzas_anat/medias:.1f} por media',
            })

        # Porcionado: ratio piezas porcionadas / media res
        if pzas_porc > 0:
            porc_esperadas = medias * ratio_esp
            ratio_real = pzas_porc / medias
            diff_porc = pzas_porc - porc_esperadas
            if ratio_real < ratio_esp * 0.7:
                alerta_p = f'BAJO ({ratio_real:.1f} vs {ratio_esp}/media)'
            elif ratio_real > ratio_esp * 1.3:
                alerta_p = f'ALTO ({ratio_real:.1f} vs {ratio_esp}/media)'
            else:
                alerta_p = f'OK ({ratio_real:.1f}/media)'
            resultado.append({
                'grupo': grupo,
                'tipo': 'PORCIONADO',
                'piezas_reales': pzas_porc,
                'piezas_esperadas': round(porc_esperadas),
                'diferencia': round(diff_porc),
                'ratio_por_media': round(ratio_real, 1),
                'ratio_esperado': ratio_esp,
                'alerta': alerta_p,
                'detalle': f'{pzas_porc} porc / {medias} medias = {ratio_real:.1f} porc/media (esperado ~{ratio_esp})',
            })

        # Feteado
        if pzas_fet > 0:
            fet_esperadas = medias * ratio_esp
            ratio_real_f = pzas_fet / medias
            diff_fet = pzas_fet - fet_esperadas
            if ratio_real_f < ratio_esp * 0.7:
                alerta_f = f'BAJO ({ratio_real_f:.1f} vs {ratio_esp}/media)'
            elif ratio_real_f > ratio_esp * 1.3:
                alerta_f = f'ALTO ({ratio_real_f:.1f} vs {ratio_esp}/media)'
            else:
                alerta_f = f'OK ({ratio_real_f:.1f}/media)'
            resultado.append({
                'grupo': grupo,
                'tipo': 'FETEADO',
                'piezas_reales': pzas_fet,
                'piezas_esperadas': round(fet_esperadas),
                'diferencia': round(diff_fet),
                'ratio_por_media': round(ratio_real_f, 1),
                'ratio_esperado': ratio_esp,
                'alerta': alerta_f,
                'detalle': f'{pzas_fet} fet / {medias} medias = {ratio_real_f:.1f} fet/media (esperado ~{ratio_esp})',
            })

        # Si solo hay porcionado/feteado y NO anatómico, calcular equivalente
        if pzas_anat == 0 and (pzas_porc > 0 or pzas_fet > 0):
            total_procesado = pzas_porc + pzas_fet
            equiv_anatomico = total_procesado / ratio_esp if ratio_esp > 0 else total_procesado
            diff_equiv = equiv_anatomico - esperadas
            if diff_equiv < -esperadas * 0.15:
                alerta_eq = f'EQUIV. FALTAN ~{abs(diff_equiv):.0f}'
            elif diff_equiv > esperadas * 0.15:
                alerta_eq = f'EQUIV. SOBRAN ~{diff_equiv:.0f}'
            else:
                alerta_eq = f'EQUIV. OK (~{equiv_anatomico:.0f} anat. equiv.)'
            resultado.append({
                'grupo': grupo,
                'tipo': 'EQUIV. ANATÓMICO',
                'piezas_reales': total_procesado,
                'piezas_esperadas': esperadas,
                'diferencia': round(diff_equiv),
                'ratio_por_media': round(equiv_anatomico / medias, 1) if medias > 0 else 0,
                'ratio_esperado': 1,
                'alerta': alerta_eq,
                'detalle': f'{total_procesado} procesadas / {ratio_esp} ratio = ~{equiv_anatomico:.0f} equiv. anatómicas (esperadas {esperadas})',
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
