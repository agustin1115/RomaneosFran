"""
report_comprador.py — Scoring de compra y generación de PDF para el comprador.
No incluye datos de facturación. Solo precio compra, puntaje y fundamentos.
"""
from fpdf import FPDF
from datetime import datetime


def fmt_ar(valor, decimales=0):
    """Formato argentino: miles con punto, decimales con coma. Máximo 2 decimales."""
    try:
        v = float(valor)
    except (ValueError, TypeError):
        return str(valor)
    if decimales == 0:
        s = f"{v:,.0f}"
    else:
        decimales = min(decimales, 2)
        s = f"{v:,.{decimales}f}"
    # swap: coma→X, punto→coma, X→punto
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def fmt_pct(valor, decimales=1):
    """Formato porcentaje: 1 decimal máximo."""
    try:
        v = float(valor)
    except (ValueError, TypeError):
        return str(valor)
    decimales = min(decimales, 2)
    s = f"{v:.{decimales}f}".replace('.', ',')
    return s + '%'


def calcular_score(reporte, rend_objetivo=66.0):
    """
    Calcula un puntaje del 1 al 10 para una compra.

    Factores (peso):
    1. Rendimiento vs objetivo      (30%) — rinde lo que debería?
    2. % Cortes caros               (25%) — salieron los cortes que valen?
    3. % Amarilla (inverso)         (25%) — poca amarilla = mejor
    4. % Carne picada (inverso)     (10%) — poca picada = mejor
    5. Precio compra vs equilibrio  (10%) — compró barato?

    Cada factor da 0-10, se pondera y se redondea.
    """
    scores = {}
    fundamentos_positivos = []
    fundamentos_negativos = []

    # 1. Rendimiento vs objetivo (30%)
    rend = reporte['rend']
    diff_rend = rend - rend_objetivo
    if diff_rend >= 3:
        scores['rendimiento'] = 10
        fundamentos_positivos.append(f"Rendimiento excelente ({fmt_pct(rend,1)}, +{fmt_pct(diff_rend,1)} sobre objetivo)")
    elif diff_rend >= 1.5:
        scores['rendimiento'] = 8
        fundamentos_positivos.append(f"Buen rendimiento ({fmt_pct(rend,1)}, +{fmt_pct(diff_rend,1)} sobre objetivo)")
    elif diff_rend >= 0:
        scores['rendimiento'] = 6
    elif diff_rend >= -2:
        scores['rendimiento'] = 4
        fundamentos_negativos.append(f"Rendimiento bajo objetivo ({fmt_pct(rend,1)}, {fmt_pct(diff_rend,1)})")
    else:
        scores['rendimiento'] = max(1, 2 + diff_rend)
        fundamentos_negativos.append(f"Rendimiento muy bajo ({fmt_pct(rend,1)}, {fmt_pct(diff_rend,1)} vs objetivo {fmt_pct(rend_objetivo,0)})")

    # 2. % Cortes caros (25%)
    pct_caros = reporte['pct_caros']
    if pct_caros >= 42:
        scores['cortes_caros'] = 10
        fundamentos_positivos.append(f"Excelente proporcion de cortes caros ({fmt_pct(pct_caros,0)})")
    elif pct_caros >= 38:
        scores['cortes_caros'] = 8
        fundamentos_positivos.append(f"Buena proporcion de cortes caros ({fmt_pct(pct_caros,0)})")
    elif pct_caros >= 34:
        scores['cortes_caros'] = 6
    elif pct_caros >= 30:
        scores['cortes_caros'] = 4
        fundamentos_negativos.append(f"Pocos cortes caros ({fmt_pct(pct_caros,0)}) - mas kg fueron a cortes baratos")
    else:
        scores['cortes_caros'] = 2
        fundamentos_negativos.append(f"Muy pocos cortes caros ({fmt_pct(pct_caros,0)}) - la mayoria fue a picada/recorte")

    # 3. Amarilla (25%) — inverso: menos = mejor
    pct_am = reporte['pct_amarilla']
    if pct_am <= 2:
        scores['amarilla'] = 10
        fundamentos_positivos.append("Sin amarilla significativa - toda la carne a precio completo")
    elif pct_am <= 5:
        scores['amarilla'] = 9
        fundamentos_positivos.append(f"Amarilla muy baja ({fmt_pct(pct_am,0)})")
    elif pct_am <= 10:
        scores['amarilla'] = 7
    elif pct_am <= 20:
        scores['amarilla'] = 5
        fundamentos_negativos.append(f"Amarilla moderada ({fmt_pct(pct_am,0)}) - reduce el precio de venta promedio")
    elif pct_am <= 35:
        scores['amarilla'] = 3
        fundamentos_negativos.append(f"Mucha amarilla ({fmt_pct(pct_am,0)}) - castiga fuerte el ingreso")
    else:
        scores['amarilla'] = 1
        fundamentos_negativos.append(f"Amarilla muy alta ({fmt_pct(pct_am,0)}) - gran parte se vende a precio minimo")

    # 4. Carne picada (10%) — inverso: menos = mejor
    pct_picada = reporte['pct_picada']
    if pct_picada <= 12:
        scores['picada'] = 10
        fundamentos_positivos.append(f"Poca picada ({fmt_pct(pct_picada,0)}) - los kg fueron a cortes de valor")
    elif pct_picada <= 16:
        scores['picada'] = 7
    elif pct_picada <= 20:
        scores['picada'] = 5
    elif pct_picada <= 25:
        scores['picada'] = 3
        fundamentos_negativos.append(f"Mucha picada ({fmt_pct(pct_picada,0)}) - exceso de recortes")
    else:
        scores['picada'] = 1
        fundamentos_negativos.append(f"Picada excesiva ({fmt_pct(pct_picada,0)}) - demasiado va al corte mas barato")

    # 5. Precio compra (10%)
    # Comparar contra el precio de equilibrio para 10% margen
    # Aproximación: precio equilibrio ~$6.800 para vaca
    precio = reporte['precio_compra']
    if precio <= 6200:
        scores['precio'] = 10
        fundamentos_positivos.append(f"Precio de compra muy competitivo (${fmt_ar(precio,0)}/kg)")
    elif precio <= 6600:
        scores['precio'] = 8
        fundamentos_positivos.append(f"Buen precio de compra (${fmt_ar(precio,0)}/kg)")
    elif precio <= 7000:
        scores['precio'] = 6
    elif precio <= 7400:
        scores['precio'] = 4
        fundamentos_negativos.append(f"Precio de compra alto (${fmt_ar(precio,0)}/kg)")
    else:
        scores['precio'] = 2
        fundamentos_negativos.append(f"Precio de compra muy alto (${fmt_ar(precio,0)}/kg) - ajusta el margen")

    # Puntaje final ponderado
    puntaje = (
        scores['rendimiento'] * 0.30 +
        scores['cortes_caros'] * 0.25 +
        scores['amarilla'] * 0.25 +
        scores['picada'] * 0.10 +
        scores['precio'] * 0.10
    )
    puntaje = round(min(10, max(1, puntaje)), 1)

    return {
        'puntaje': puntaje,
        'scores': scores,
        'positivos': fundamentos_positivos,
        'negativos': fundamentos_negativos,
    }


def generar_pdf_comprador(reportes_con_score, fecha_reporte=None):
    """
    Genera un PDF con el reporte para el comprador.
    NO incluye datos de facturación — solo precio compra, puntaje y fundamentos.
    """
    if not fecha_reporte:
        fecha_reporte = datetime.now().strftime('%d/%m/%Y')

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)

    # ── Página 1: Resumen ──
    pdf.add_page()
    pdf.set_font('Helvetica', 'B', 22)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 15, 'Reporte de Compras', new_x='LMARGIN', new_y='NEXT', align='C')
    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 8, f'TF Carnes S.A. - {fecha_reporte}', new_x='LMARGIN', new_y='NEXT', align='C')
    pdf.ln(5)

    # Ranking
    sorted_rep = sorted(reportes_con_score, key=lambda x: x['score']['puntaje'], reverse=True)

    pdf.set_font('Helvetica', 'B', 14)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 10, 'Ranking de compras', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(2)

    # Tabla header
    pdf.set_font('Helvetica', 'B', 9)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    col_widths = [8, 52, 20, 18, 18, 18, 16, 20, 20]
    headers = ['#', 'Romaneo', 'Cat.', 'Medias', 'Rend %', 'Amar %', 'Caros %', '$/kg', 'Puntaje']
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 7, h, border=1, align='C', fill=True)
    pdf.ln()

    # Tabla data
    pdf.set_font('Helvetica', '', 8)
    pdf.set_text_color(0, 0, 0)
    for idx, r in enumerate(sorted_rep):
        rep = r['reporte']
        sc = r['score']
        puntaje = sc['puntaje']

        # Color de fondo según puntaje
        if puntaje >= 7:
            pdf.set_fill_color(232, 245, 233)
        elif puntaje >= 5:
            pdf.set_fill_color(255, 243, 224)
        else:
            pdf.set_fill_color(252, 228, 236)

        vals = [
            str(idx + 1),
            rep['nombre'][:28],
            rep['categoria'][:6],
            str(rep['medias']),
            fmt_ar(rep['rend'], 1),
            fmt_ar(rep['pct_amarilla'], 0),
            fmt_ar(rep['pct_caros'], 0),
            f"${fmt_ar(rep['precio_compra'], 0)}",
            fmt_ar(puntaje, 1),
        ]
        for i, v in enumerate(vals):
            pdf.cell(col_widths[i], 6, v, border=1, align='C', fill=True)
        pdf.ln()

    pdf.ln(5)

    # Mejor y peor
    mejor = sorted_rep[0]
    peor = sorted_rep[-1]

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(46, 125, 50)
    pdf.cell(0, 8, f"Mejor compra: {mejor['reporte']['nombre']} ({fmt_ar(mejor['score']['puntaje'], 1)}/10)",
             new_x='LMARGIN', new_y='NEXT')
    pdf.set_font('Helvetica', '', 9)
    pdf.set_text_color(0, 0, 0)
    for f in mejor['score']['positivos']:
        pdf.cell(5, 5, '')
        pdf.cell(0, 5, f'+ {f}', new_x='LMARGIN', new_y='NEXT')

    pdf.ln(3)
    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(198, 40, 40)
    pdf.cell(0, 8, f"Peor compra: {peor['reporte']['nombre']} ({fmt_ar(peor['score']['puntaje'], 1)}/10)",
             new_x='LMARGIN', new_y='NEXT')
    pdf.set_font('Helvetica', '', 9)
    pdf.set_text_color(0, 0, 0)
    for f in peor['score']['negativos']:
        pdf.cell(5, 5, '')
        pdf.cell(0, 5, f'- {f}', new_x='LMARGIN', new_y='NEXT')

    # ── Detalle por romaneo ──
    for r in sorted_rep:
        rep = r['reporte']
        sc = r['score']
        puntaje = sc['puntaje']

        pdf.add_page()
        pdf.set_font('Helvetica', 'B', 16)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(0, 12, f"{rep['nombre']}", new_x='LMARGIN', new_y='NEXT')

        # Puntaje grande
        if puntaje >= 7:
            pdf.set_text_color(46, 125, 50)
            emoji = 'BUENO'
        elif puntaje >= 5:
            pdf.set_text_color(230, 81, 0)
            emoji = 'REGULAR'
        else:
            pdf.set_text_color(198, 40, 40)
            emoji = 'BAJO'

        pdf.set_font('Helvetica', 'B', 36)
        pdf.cell(40, 25, fmt_ar(puntaje, 1), align='C')
        pdf.set_font('Helvetica', 'B', 14)
        pdf.cell(20, 25, f'/ 10')
        pdf.set_font('Helvetica', 'B', 18)
        pdf.cell(0, 25, emoji, new_x='LMARGIN', new_y='NEXT')

        pdf.ln(3)

        # Datos básicos
        pdf.set_text_color(0, 0, 0)
        pdf.set_font('Helvetica', 'B', 10)
        pdf.set_fill_color(214, 228, 240)

        datos = [
            ('Categoria', rep['categoria']),
            ('Medias reses', str(rep['medias'])),
            ('Kg entrada', fmt_ar(rep['kg_entrada'], 0)),
            ('Kg carne', fmt_ar(rep['kg_carne'], 0)),
            ('Rendimiento', fmt_pct(rep['rend'], 1)),
            ('Precio compra', f"${fmt_ar(rep['precio_compra'], 0)}/kg"),
            ('Amarilla', fmt_pct(rep['pct_amarilla'], 1)),
            ('Cortes caros', fmt_pct(rep['pct_caros'], 1)),
            ('Carne picada', fmt_pct(rep['pct_picada'], 1)),
        ]
        for label, val in datos:
            pdf.cell(55, 7, label, border=1, fill=True)
            pdf.set_font('Helvetica', '', 10)
            pdf.cell(50, 7, val, border=1, align='R')
            pdf.set_font('Helvetica', 'B', 10)
            pdf.ln()

        pdf.ln(5)

        # Desglose de puntaje
        pdf.set_font('Helvetica', 'B', 11)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(0, 8, 'Desglose del puntaje', new_x='LMARGIN', new_y='NEXT')

        pdf.set_font('Helvetica', 'B', 8)
        pdf.set_fill_color(31, 78, 121)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(50, 6, 'Factor', border=1, fill=True, align='C')
        pdf.cell(20, 6, 'Peso', border=1, fill=True, align='C')
        pdf.cell(20, 6, 'Nota', border=1, fill=True, align='C')
        pdf.cell(25, 6, 'Ponderado', border=1, fill=True, align='C')
        pdf.ln()

        pdf.set_text_color(0, 0, 0)
        pdf.set_font('Helvetica', '', 8)
        factores = [
            ('Rendimiento', 0.30, sc['scores']['rendimiento']),
            ('Cortes caros', 0.25, sc['scores']['cortes_caros']),
            ('Amarilla (menos=mejor)', 0.25, sc['scores']['amarilla']),
            ('Picada (menos=mejor)', 0.10, sc['scores']['picada']),
            ('Precio compra', 0.10, sc['scores']['precio']),
        ]
        for nombre, peso, nota in factores:
            pond = nota * peso
            if nota >= 7:
                pdf.set_fill_color(232, 245, 233)
            elif nota >= 5:
                pdf.set_fill_color(255, 243, 224)
            else:
                pdf.set_fill_color(252, 228, 236)
            pdf.cell(50, 6, nombre, border=1, fill=True)
            pdf.cell(20, 6, f'{peso*100:.0f}%', border=1, align='C', fill=True)
            pdf.cell(20, 6, f'{nota}/10', border=1, align='C', fill=True)
            pdf.cell(25, 6, fmt_ar(pond, 2), border=1, align='C', fill=True)
            pdf.ln()

        pdf.ln(5)

        # Fundamentos
        if sc['positivos']:
            pdf.set_font('Helvetica', 'B', 10)
            pdf.set_text_color(46, 125, 50)
            pdf.cell(0, 7, 'Lo que mejoro la compra:', new_x='LMARGIN', new_y='NEXT')
            pdf.set_font('Helvetica', '', 9)
            pdf.set_text_color(0, 0, 0)
            for f in sc['positivos']:
                pdf.cell(5, 5, '')
                pdf.multi_cell(0, 5, f'+ {f}', new_x='LMARGIN', new_y='NEXT')

        pdf.ln(2)

        if sc['negativos']:
            pdf.set_font('Helvetica', 'B', 10)
            pdf.set_text_color(198, 40, 40)
            pdf.cell(0, 7, 'Lo que empeoro la compra:', new_x='LMARGIN', new_y='NEXT')
            pdf.set_font('Helvetica', '', 9)
            pdf.set_text_color(0, 0, 0)
            for f in sc['negativos']:
                pdf.cell(5, 5, '')
                pdf.multi_cell(0, 5, f'- {f}', new_x='LMARGIN', new_y='NEXT')

        pdf.ln(3)

        # Recomendación
        pdf.set_font('Helvetica', 'B', 10)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(0, 7, 'Recomendacion para mejorar:', new_x='LMARGIN', new_y='NEXT')
        pdf.set_font('Helvetica', '', 9)
        pdf.set_text_color(0, 0, 0)

        recos = []
        if sc['scores']['amarilla'] <= 4:
            recos.append("Buscar tropas con menos grasa amarilla. La amarilla se vende a $9.500/kg mientras los cortes normales promedian $13.000+. Cada punto menos de amarilla mejora directamente el ingreso.")
        if sc['scores']['rendimiento'] <= 5:
            recos.append("Priorizar animales con mejor terminacion y conformacion. Un rendimiento 2% mayor en una tropa de 100 medias equivale a ~240 kg mas de carne vendible.")
        if sc['scores']['cortes_caros'] <= 5:
            recos.append("Los cortes caros (bife, lomo, ojo, entraña, vacio) generan el mayor ingreso. Buscar animales que den mejor proporcion de estos cortes.")
        if sc['scores']['picada'] <= 4:
            recos.append("Demasiados kg terminan como picada (el corte mas barato). Revisar si el desposte puede optimizarse o si la calidad del animal lo permite.")
        if sc['scores']['precio'] <= 4:
            recos.append("Negociar mejor el precio de compra. Cada $100/kg menos en una tropa de 10.000 kg son $1.000.000 de ahorro directo.")

        if not recos:
            recos.append("Compra bien balanceada. Mantener los criterios de seleccion actuales.")

        for reco in recos:
            pdf.cell(5, 5, '')
            pdf.multi_cell(0, 5, f'> {reco}', new_x='LMARGIN', new_y='NEXT')
            pdf.ln(1)

    # Footer en todas las páginas
    pdf.set_y(-15)
    pdf.set_font('Helvetica', 'I', 7)
    pdf.set_text_color(150, 150, 150)
    pdf.cell(0, 10, f'TF Carnes S.A. - Reporte Comprador - {fecha_reporte} - Confidencial',
             align='C')

    return pdf.output()
