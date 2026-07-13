"""
historial_costos.py — Guarda cómo cerró el costo por kg de cada mes (frigorífico,
insumos, fletes) para armar la serie histórica y estimar el pricing futuro.
"""
import os
import json


def cargar(path):
    if path and os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []
    return []


def guardar_mes(path, registro):
    """registro: {'mes':'2026-06', 'frigo_kg':.., 'insumos_kg':.., 'fletes_kg':.., ...}.
    Reemplaza el mes si ya existía."""
    hist = [h for h in cargar(path) if h.get('mes') != registro.get('mes')]
    registro['total_op_kg'] = (registro.get('frigo_kg', 0) + registro.get('insumos_kg', 0)
                               + registro.get('fletes_kg', 0))
    hist.append(registro)
    hist.sort(key=lambda h: str(h.get('mes', '')))
    d = os.path.dirname(path)
    if d:
        os.makedirs(d, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)
    return hist


def proyeccion(hist, inflacion_pct=0.0, metodo='ultimo'):
    """Estima el costo/kg del próximo mes.
    - 'ultimo': último mes × (1 + inflación).
    - 'tendencia': proyecta la variación promedio de los últimos meses."""
    if not hist:
        return None
    campos = ['frigo_kg', 'insumos_kg', 'fletes_kg', 'total_op_kg']
    f = 1 + inflacion_pct / 100.0
    ult = hist[-1]
    out = {'base_mes': ult.get('mes'), 'metodo': metodo, 'inflacion_pct': inflacion_pct}

    if metodo == 'tendencia' and len(hist) >= 2:
        for c in campos:
            vals = [h.get(c, 0) for h in hist if h.get(c)]
            if len(vals) >= 2:
                # variación % mensual promedio
                var = []
                for i in range(1, len(vals)):
                    if vals[i - 1]:
                        var.append(vals[i] / vals[i - 1] - 1)
                factor = (1 + sum(var) / len(var)) if var else f
                out[c] = vals[-1] * factor
            else:
                out[c] = (vals[-1] if vals else 0) * f
    else:
        for c in campos:
            out[c] = ult.get(c, 0) * f
    return out
