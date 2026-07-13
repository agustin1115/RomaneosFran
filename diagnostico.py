"""
diagnostico.py — Convierte los desvíos en ALERTAS con causa, recomendación y
dónde investigar. El foco es comparar CANTIDADES FÍSICAS (cabezas, medias, kg,
piezas): a igual cantidad, igual costo. Si la cantidad difiere, se reporta.
"""

SEV_ALTA = 10.0   # % de diferencia de cantidad
SEV_MEDIA = 3.0


def _pct(real, proy):
    return (real - proy) / proy * 100 if proy else (100.0 if real else 0.0)


def _sev(pct):
    a = abs(pct)
    if a >= SEV_ALTA:
        return '🔴 Alta'
    if a >= SEV_MEDIA:
        return '🟡 Media'
    return '🟢 OK'


# ══════════════════════════════════════════════════════════════════════
# FRIGORÍFICO — comparación por cantidad física
# ══════════════════════════════════════════════════════════════════════
def diagnostico_frigorifico(proy, real, kg_carne):
    """proy: salida sumada de costo_proyectado (cabezas, medias, kg_carne,
    kg_congelado + $). real: mes del conector (vep_cab, cuarteo_und,
    despostada_kg, congelado_tunel_kg + $)."""
    filas = [
        ('VEP', 'cabezas', proy.get('cabezas', 0), real.get('vep_cab', 0) or real.get('cab', 0),
         proy.get('vep', 0), real.get('vep', 0),
         "El frigorífico factura VEP por cabeza. Si te cobran más cabezas que las de los romaneos "
         "tomados, estás pagando faena que no aparece en tu producción (o hay un desfase de fechas).",
         "Cruzá las cabezas facturadas del mes contra la suma de cabezas (medias/2) de los romaneos "
         "cargados. A igual cabezas, igual costo: si no cuadra, reclamá al frigorífico."),
        ('Cuarteo', 'medias', proy.get('medias', 0), real.get('cuarteo_und', 0),
         proy.get('cuarteo', 0), real.get('cuarteo', 0),
         "El cuarteo se hace a TODAS las medias, así que la cantidad tiene que dar igual que las medias "
         "de los romaneos. Cualquier diferencia es sobrefacturación o medias no cargadas.",
         "Compará unidades cuarteadas facturadas vs medias de los romaneos del mes."),
        ('Despostada', 'kg', proy.get('kg_carne', 0), real.get('despostada_kg', 0),
         proy.get('despostada', 0), real.get('despostada', 0),
         "Se cobra por kg despostado. Si los kg facturados superan los kg producidos, hay remanejos "
         "(se cobran a $395/kg) o kg de más.",
         "Cruzá kg facturados vs kg carne del romaneo y buscá líneas 'DESPOSTADA R' (remanejo)."),
        ('Congelado (túnel)', 'kg', proy.get('kg_congelado', 0), real.get('congelado_tunel_kg', 0),
         proy.get('congelado', 0), real.get('congelado_tunel', 0) or real.get('congelado', 0),
         "Deberían cobrarte túnel por los MISMOS kg que produjiste con destino congelado. Si el túnel "
         "facturado no coincide con los kg congelados del romaneo, revisalo.",
         "Compará kg de SERVIC CONGELADO (túnel) facturados vs kg con destino CONGELADO del romaneo."),
    ]
    UMBRAL = 3.0  # % de tolerancia (desfasajes de fecha facturación/romaneo)
    out = []
    for concepto, unidad, cp, cr, mp, mr, causa, reco in filas:
        cant_pct = _pct(cr, cp)
        if cant_pct > UMBRAL:      # te facturan MÁS de lo que produjiste → ROJO
            direccion, severidad, box = 'mas', '🔴 Te cobran de MÁS', 'error'
            titulo = f"{concepto}: te facturan de MÁS — te están cobrando lo que no produjiste"
        elif cant_pct < -UMBRAL:   # te facturan MENOS → VERDE (a tu favor)
            direccion, severidad, box = 'menos', '🟢 A tu favor', 'success'
            titulo = f"{concepto}: te facturaron de MENOS — les faltó cobrarte, puede venir después"
        else:
            direccion, severidad, box = 'ok', '✅ OK', 'ok'
            titulo = f"{concepto}: coincide con los romaneos"
        detalle = (f"Facturado: {cr:,.0f} {unidad} · Romaneos: {cp:,.0f} {unidad} · "
                   f"Δ {cr - cp:+,.0f} {unidad} ({cant_pct:+.1f}%)")
        out.append({
            'concepto': concepto, 'unidad': unidad,
            'cant_proy': cp, 'cant_real': cr, 'cant_desvio': cr - cp, 'cant_desvio_pct': cant_pct,
            'proyectado_kg': mp / kg_carne if kg_carne else 0,
            'real_kg': mr / kg_carne if kg_carne else 0,
            'desvio_kg': (mr - mp) / kg_carne if kg_carne else 0,
            'desvio_total': mr - mp, 'direccion': direccion, 'severidad': severidad, 'box': box,
            'titulo': titulo, 'detalle': detalle, 'causa': causa, 'recomendacion': reco,
        })
    return out


# ══════════════════════════════════════════════════════════════════════
# INSUMOS — 1 bolsa / 1 etiqueta por pieza + $/kg real por consumo
# ══════════════════════════════════════════════════════════════════════
def diagnostico_insumos(teorico, real_ins, kg_despostados):
    """teorico: costo_teorico (piezas, por_categoria, $). real_ins: dict con
    consumo de unidades y $ por familia (bolsas/etiquetas/caja)."""
    piezas_bolsa = (teorico['por_categoria']['chica'] + teorico['por_categoria']['grande']
                    + teorico['por_categoria']['hueso'])
    piezas_total = teorico.get('piezas', 0)

    def fam(nombre, esperado_und, real_und, real_costo, causa, reco):
        d = real_und - esperado_und
        ratio = real_und / esperado_und if esperado_und else (2.0 if real_und else 1.0)
        if ratio >= 1.5:
            severidad, box = '🔴 Crítico', 'error'          # ~doble o más
        elif ratio >= 1.15:
            severidad, box = '🟡 Sobreconsumo', 'warning'
        elif ratio >= 0.85:
            severidad, box = '✅ OK', 'ok'
        else:
            severidad, box = '🟢 Menor a lo esperado', 'success'
        return {
            'concepto': nombre, 'esperado_und': esperado_und, 'real_und': real_und,
            'desvio_und': d, 'desvio_pct': _pct(real_und, esperado_und), 'ratio': ratio,
            'real_kg': real_costo / kg_despostados if kg_despostados else 0,
            'costo_total': real_costo, 'severidad': severidad, 'box': box,
            'titulo': f"{nombre}: usás {ratio:.1f}× lo esperado (deberían ser 1 por pieza)",
            'detalle': f"Consumidas: {real_und:,.0f} · Esperadas (1/pieza): {esperado_und:,.0f} · "
                       f"Ratio {ratio:.2f}× · Δ {d:+,.0f}",
            'causa': causa, 'recomendacion': reco,
        }

    out = [
        fam('Bolsas', piezas_bolsa, real_ins.get('bolsas_und', 0), real_ins.get('bolsas', 0),
            "Deberías usar 1 bolsa por pieza. Si consumiste más, hay roturas, reempaque, o se usó "
            "una bolsa (grande/hueso) por falta de stock de la correcta — y encima más cara.",
            "Consumo = stock inicial + compras del mes − stock final. Si supera las piezas, "
            "deberías tener más stock del que hay: contá físico y buscá mermas/sustituciones."),
        fam('Etiquetas alto impacto', piezas_total, real_ins.get('etiquetas_und', 0), real_ins.get('etiquetas', 0),
            "1 etiqueta de alto impacto por pieza. Más que eso = reimpresiones o descartes.",
            "Compará etiquetas consumidas vs piezas producidas; revisá reimpresiones."),
    ]
    return out


# ══════════════════════════════════════════════════════════════════════
# FLETES
# ══════════════════════════════════════════════════════════════════════
def resumen_fletes(af):
    out = []
    for seg, d in af.get('por_segmento', {}).items():
        if d['kg'] == 0:
            continue
        if d['gap_kg'] > 0:
            out.append({
                'concepto': f"Fletes {seg}", 'real_kg': d['real_kg'], 'benchmark': d['bmax'],
                'desvio_kg': d['gap_kg'], 'desvio_total': d['sobrecosto'],
                'severidad': _sev(_pct(d['real_kg'], d['bmax'])),
                'causa': (f"{seg}: cargas chicas, viajes compuestos o proveedor caro para el segmento. "
                          "La gestión (comisiones) también suma."),
                'recomendacion': ("Consolidá cargas, mové a proveedor más barato del segmento y revisá "
                                  "las comisiones de gestión."),
            })
    return out
