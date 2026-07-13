---
name: romaneo
description: |
  Analiza romaneos de despostada de carne vacuna con soporte multi-calidad.
  Usa este skill cada vez que el usuario mencione "romaneo", "despostada",
  "rendimiento de tropa", "análisis de cortes", o suba un PDF de producción cárnica.
  Soporta 4 perfiles de calidad: Standard, Búfalo, Premium Black y Exportación.
  Permite análisis individual y acumulado de múltiples archivos.
---

# Análisis de Romaneo de Despostada — TF Carnes S.A. (v2.0)

## Qué hace este skill

Toma PDFs de romaneo (partes de producción de despostada), extrae los datos, y genera Excel con 6 hojas:
1. **PARAMETROS** — perfil de calidad activo + costos variables editables + rendimiento objetivo
2. **PRECIOS** — precio de venta por corte y por cliente, con columna "Tipo $" (CARO/BARATO)
3. **ROMANEO** — detalle línea por línea con grupo, tipo, cliente, amarillas resaltadas
4. **ANALISIS** — rendimiento por grupo vs % esperado con calificación (invertida para baratos)
5. **VENTAS** — desglose de kg e ingreso por cliente/contramarca
6. **RESULTADO** — P&L completo con CM, margen, calificación

## Perfiles de calidad

### Standard (Mercado Interno)
- Costos: MO $790 + Insumos $350 + Flete $180 + SENASA $150
- IIBB: 5%
- Rendimiento obj: Vaca 66%, Novillo/Novillito 68%, Vaquillona/Bubalino 67%, Toro 66%

### Búfalo
- Costos: MO $950 + Insumos $420 + Flete $220 + SENASA $180
- IIBB: 5%
- Rendimiento obj: Vaca 64%, Novillo/Novillito 66%, Vaquillona/Bubalino 65%, Toro 64%
- Más magro, rendimientos de cortes premium más altos, menos carne picada

### Premium Black
- Costos: MO $1.100 + Insumos $580 + Flete $250 + SENASA $200
- IIBB: 5%
- Rendimiento obj: Vaca 67%, Novillo/Novillito 69%, Vaquillona/Bubalino 68%, Toro 67%
- Despostada de precisión, envasado premium

### Exportación
- Costos: MO $1.300 + Insumos $720 + Flete $450 + SENASA $350
- IIBB: 2.5% (reducido por exportación)
- Rendimiento obj: Vaca 68%, Novillo/Novillito 70%, Vaquillona/Bubalino 69%, Toro 68%
- Packaging internacional, certificaciones export, Hilton/Cuota

**Todos los costos son editables** en la sidebar de la app o en la hoja PARAMETROS del Excel.

## Paso 1: Recibir PDF y datos

El usuario sube uno o más PDFs. Preguntarle (si no lo dijo):
- **Precio de compra $/kg** de la media res (s/IVA)
- **Perfil de calidad** (Standard / Búfalo / Premium Black / Exportación)

La **categoría** se detecta automáticamente del PDF (VA/NO/NT/VQ/TO/BU/BB).

## Paso 2: Extraer datos del PDF

Usar `pdfplumber`. Dos formatos posibles:
- **"Rendimientos de Despostada"**: contramarca "47- 215"
- **"Resultado Despostada"**: contramarca "87-100215"

Parsear: encabezado (N°, fecha, medias, kg entrada), salidas (código, desc, destino, piezas, unidades, kg, Nro.Venta), sub-productos, merma.

## Paso 3: Clasificar cada línea

### Mapeo corte → grupo: ver config.py CORTE_TO_GRUPO
### Tipo: PORCIONADO / FETEADO / ANATÓMICO
### CARO/BARATO: según precio NETO PEYA ≥ $15.000/kg
### Amarillas: ctm 47/73/74 → precio fijo editable ($10.500/kg default)

## Paso 4: Precio por cliente

Leer **SIEMPRE** de `PRECIOS DE FACTURACIÓN.xlsx`. Mapeo contramarca → columna: ver config.py CONTRAMARCA_MAP.

## Paso 5: Generar Excel

Usar `excel_builder.build_analisis()` pasando calidad, costos editados y price_matrix.

### Rendimiento objetivo dinámico
- Base según categoría + calidad (ver tabla arriba)
- Si >50% anatómico → +1%

### Calificación invertida para cortes baratos
- **CAROS** (más = mejor): ÓPTIMO ≥110% | BUENO 95-110% | REGULAR 80-95% | MALO <80%
- **BARATOS** (menos = mejor): ÓPTIMO ≤80% | BUENO 80-95% | REGULAR 95-110% | MALO ≥110%

## Paso 6: Análisis múltiple

Si hay más de un PDF:
- **Individual**: un Excel por cada PDF
- **Acumulado**: combina todos en uno solo (suma kg, medias, concatena cortes)
- **Ambos**: genera todo

## Paso 7: Actualizar historial

Guardar en `ROMANEOS/historial_romaneos.json` con campo adicional `calidad`.

## Paso 8: Presentar resultados

1. Link/descarga del Excel
2. Resumen: Kg entrada, Kg carne, Rend%, Objetivo%, Costo, Ingreso, CM, Margen%, Calificación
3. Alertas de amarillas (kg y %)
4. Cortes MALO/REGULAR
5. Comparación con historial por categoría y calidad

## Escala de calificación del negocio
ÓPTIMO ≥15% | BUENO 8-15% | REGULAR 3-8% | MALO 0-3% | PÉRDIDA <0%

## Programa standalone

La app Streamlit se ejecuta con:
```bash
cd 11_Produccion/romaneo_app
pip install -r requirements.txt
streamlit run app.py
```
