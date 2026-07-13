"""
TF Carnes — Reporte de Producción para Gerente de Planta.

App Streamlit independiente que:
1. Carga romaneos (Drive automático + manual)
2. Cruza con planilla de compras (Google Sheets)
3. Genera reportes HTML enfocados en producción (sin info comercial):
   rendimiento, cortes faltantes, calidad de despostada, mermas, sugerencias.

Para correr: streamlit run app_produccion.py
"""
import os
import sys
import json
import tempfile
from datetime import datetime, date

import streamlit as st

# ══════════════════════════════════════════════════════════════════════
# IMPORTS COMPARTIDOS
# ══════════════════════════════════════════════════════════════════════
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import REND_OBJETIVO, AMARILLA_CONTRAMARCAS
from pdf_parser import (parse_romaneo_pdf, detectar_tipo_pdf, parse_remanejo_pdf,
                         es_correccion, acumular_romaneos)
from drive_loader import (listar_pdfs_drive, descargar_pdf, organizar_por_mes_semana)
from html_builder_produccion import build_html_produccion


# ══════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════
SPREADSHEET_COMPRAS = '1OVoP_2QE2gWnX6CPL7ys2vjcisMbuSc9kUpfTcWy66M'
CREDENTIALS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  'credentials.json')

# ── Bootstrap credenciales para Streamlit Cloud ───────────────────────────
# Si no existe credentials.json pero hay secrets en Streamlit Cloud, lo
# escribimos en runtime para que la conexión a Drive/Sheets funcione.
if not os.path.exists(CREDENTIALS_PATH):
    try:
        if 'gcp_service_account' in st.secrets:
            with open(CREDENTIALS_PATH, 'w') as _f:
                json.dump(dict(st.secrets['gcp_service_account']), _f)
    except Exception:
        pass
# ──────────────────────────────────────────────────────────────────────────

HISTORIAL_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'ROMANEOS', 'historial_romaneos.json')


st.set_page_config(
    page_title="TF Carnes — Producción",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════════════════════════
# ESTILOS
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
:root {
    --tf-green-dark: #1B4D3E;
    --tf-green: #2D7D5F;
    --tf-gold: #C9A84C;
    --tf-bg: #F5F7F5;
}
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}
.stApp { background: var(--tf-bg); }
header[data-testid="stHeader"] { height: 2.2rem !important; background: transparent !important; }
.main .block-container { padding-top: 1rem !important; }
.tf-hero {
    background: linear-gradient(135deg, #1B4D3E, #2D7D5F);
    color: #fff; padding: 14px 22px; border-radius: 12px; margin-bottom: 18px;
    box-shadow: 0 2px 10px rgba(27,77,62,.22);
    display: flex; justify-content: space-between; align-items: center; gap: 16px;
}
.tf-hero h1 { font-size: 20px; font-weight: 800; letter-spacing: 2.5px;
              text-transform: uppercase; margin: 0; }
.tf-hero p { font-size: 11px; color: #C9A84C; margin: 2px 0 0 0;
             letter-spacing: 1.4px; text-transform: uppercase; font-weight: 600; }
.tf-hero-meta { font-size: 11px; color: rgba(255,255,255,.85);
                letter-spacing: 1.2px; text-transform: uppercase; font-weight: 600;
                text-align: right; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 2px solid #E8F5E9; }
.stTabs [aria-selected="true"] {
    background: #fff !important; color: var(--tf-green-dark) !important;
    border-bottom: 3px solid var(--tf-gold) !important; font-weight: 700;
}
[data-testid="stMetric"] {
    background: #fff; border: 1px solid #DDE5DD; border-radius: 12px;
    padding: 12px 14px; box-shadow: 0 1px 6px rgba(0,0,0,.04);
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #C9A84C, #B89840) !important;
    color: #fff !important; border: none !important; font-weight: 700;
    text-transform: uppercase; letter-spacing: 1.2px;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# ESTADO DE SESIÓN
# ══════════════════════════════════════════════════════════════════════
if 'parsed_files' not in st.session_state:
    st.session_state.parsed_files = []
if 'remanejos' not in st.session_state:
    st.session_state.remanejos = []
if 'html_results' not in st.session_state:
    st.session_state.html_results = []


# ══════════════════════════════════════════════════════════════════════
# HELPERS — cargar compras y cruzar tropas
# (copia simplificada de app.py para que esta app sea standalone)
# ══════════════════════════════════════════════════════════════════════
def _parse_ar_number(s):
    if s is None: return None
    s = str(s).replace('$', '').replace(' ', '').strip()
    if not s or s in {'0', '-', '—'}: return None
    try:
        if '.' in s and ',' in s:
            return float(s.replace('.', '').replace(',', '.'))
        if ',' in s:
            return float(s.replace('.', '').replace(',', '.'))
        if '.' in s:
            parts = s.split('.')
            if len(parts) == 2 and 1 <= len(parts[1]) <= 2:
                return float(s)
            return float(s.replace('.', ''))
        return float(s)
    except (ValueError, TypeError):
        return None


@st.cache_data(ttl=300)
def cargar_compras_google_sheets():
    """Lee 'Compras' de GSheets. Retorna {tropa: [partidas]} con fecha inferida."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        if not os.path.exists(CREDENTIALS_PATH):
            return None, 'Sin credentials.json'

        SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
                  'https://www.googleapis.com/auth/drive.metadata.readonly']
        creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SPREADSHEET_COMPRAS)
        ws = sh.worksheet('Compras')
        all_data = ws.get_all_values()
        if len(all_data) < 2:
            return None, 'Planilla vacía'

        from datetime import date as _date
        hoy = _date.today()

        # Primera pasada
        filas_parsed = []
        for row in all_data[1:]:
            if len(row) < 54: continue
            fecha = row[3].strip() if row[3] else ''
            tipo = row[7].strip() if row[7] else ''
            precio = _parse_ar_number(row[53])
            if not precio or precio <= 0: continue

            # Detectar monto
            monto = None; kg_der = None; best = 0
            for c in range(8, 53):
                if c in {41, 42, 43}: continue
                val = _parse_ar_number(row[c])
                if not val or val <= 0: continue
                kg_try = val / precio
                if 50 <= kg_try <= 100000 and val > best:
                    best = val; monto = val; kg_der = kg_try

            dd = mm = yyyy = None
            if fecha:
                try:
                    parts = fecha.replace('-', '/').split('/')
                    if len(parts) >= 2:
                        dd = int(parts[0]); mm = int(parts[1])
                        if len(parts) >= 3 and parts[2].strip():
                            ry = int(parts[2])
                            yyyy = ry + 2000 if ry < 100 else ry
                except (ValueError, IndexError):
                    dd = mm = yyyy = None

            tropas = set()
            for ct in [41, 42, 43]:
                t = row[ct].strip() if row[ct] else ''
                if t and t != '0' and t != 'XX':
                    tropas.add(t)
            es_compra_media = any(t.startswith('7') and len(t) >= 4 for t in tropas)

            entrada = {'precio': precio, 'monto': monto, 'kg_sheet': kg_der,
                       'tipo': tipo, 'fecha': fecha, 'fecha_faena_dt': None,
                       'es_compra_media': es_compra_media}
            filas_parsed.append((dd, mm, yyyy, entrada, tropas))

        # Inferencia de año (igual que app.py — solo decrementar en saltos ≥6 meses)
        year_cursor = hoy.year
        prev_mm = None
        for i in range(len(filas_parsed) - 1, -1, -1):
            dd, mm, yyyy, entrada, _ = filas_parsed[i]
            if mm is None: continue
            if yyyy is not None:
                year_cursor = yyyy; prev_mm = mm; continue
            if prev_mm is not None and (mm - prev_mm) >= 6:
                year_cursor -= 1
            try:
                cand = date(year_cursor, mm, dd)
            except ValueError:
                cand = None
            if cand and cand > hoy:
                try: cand = date(year_cursor - 1, mm, dd)
                except ValueError: cand = None
            entrada['fecha_faena_dt'] = cand
            prev_mm = mm

        for dd, mm, yyyy, entrada, _ in filas_parsed:
            if (yyyy is not None and dd is not None and mm is not None
                    and entrada['fecha_faena_dt'] is None):
                try: entrada['fecha_faena_dt'] = date(yyyy, mm, dd)
                except ValueError: pass

        compras = {}
        for _, _, _, entrada, tropas in filas_parsed:
            for t in tropas:
                compras.setdefault(t, []).append(entrada)
        return compras, f'{len(compras)} tropas'
    except Exception as e:
        return None, str(e)


def _extraer_tropas_con_kg(pdf_path):
    import pdfplumber, re
    out = {}
    if not pdf_path or not os.path.exists(pdf_path):
        return out
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ''
                in_entrada = False
                for line in text.split('\n'):
                    if 'Entrada Despostada' in line:
                        in_entrada = True; continue
                    if in_entrada and 'Total >>>' in line:
                        return out
                    if in_entrada:
                        m_t = re.search(r'(\d{4,6})-\s*\d+', line)
                        if not m_t: continue
                        tropa = m_t.group(1)
                        nums = re.findall(r'([\d.,]+)', line)
                        kg_val = 0
                        for n in reversed(nums):
                            try:
                                v = float(n.replace(',', ''))
                                if v > 50:
                                    kg_val = v; break
                            except ValueError: continue
                        if kg_val > 0:
                            out[tropa] = out.get(tropa, 0) + kg_val
    except Exception:
        pass
    return out


def parsear_fecha_romaneo(fecha_str):
    if not fecha_str: return None
    try:
        parts = fecha_str.strip().split('/')
        if len(parts) == 3:
            return datetime(int(parts[2]), int(parts[1]), int(parts[0])).date()
        if len(parts) == 2:
            return datetime(2026, int(parts[1]), int(parts[0])).date()
    except (ValueError, IndexError):
        pass
    return None


def cruzar_precio_compra(parsed_data, pdf_path=None):
    compras, status = cargar_compras_google_sheets()
    if not compras:
        return None, None, status
    tropas_kg = _extraer_tropas_con_kg(pdf_path)
    tropas_encontradas = set(tropas_kg.keys())
    if not tropas_encontradas:
        return None, None, 'Sin tropas en PDF'

    fecha_rom_dt = None
    if parsed_data.get('fecha'):
        try:
            fecha_rom_dt = datetime.strptime(parsed_data['fecha'].strip(), '%d/%m/%Y').date()
        except Exception: pass

    matches = []
    sin_match = []
    for tropa in tropas_encontradas:
        kg_t = tropas_kg.get(tropa, 0)
        entradas = compras.get(tropa)
        if not entradas:
            sin_match.append({'tropa': tropa, 'kg': kg_t})
            continue
        if fecha_rom_dt:
            cercanas = [e for e in entradas
                        if e.get('fecha_faena_dt')
                        and 0 <= (fecha_rom_dt - e['fecha_faena_dt']).days <= 120]
            if cercanas:
                entradas = cercanas
            else:
                sin_match.append({'tropa': tropa, 'kg': kg_t,
                                  'razon': 'sin partida cercana'})
                continue
        partidas = len(entradas)
        monto = sum(e['monto'] for e in entradas if e.get('monto'))
        kg_sh = sum(e['kg_sheet'] for e in entradas if e.get('kg_sheet'))
        precio = (monto / kg_sh) if (monto and kg_sh) else (
            sum(e['precio'] for e in entradas) / len(entradas))
        primera = entradas[0]
        es_compra = (tropa.startswith('7') and len(tropa) >= 4) \
                    or any(e.get('es_compra_media') for e in entradas)
        matches.append({
            'tropa': tropa, 'precio': round(precio),
            'tipo': primera['tipo'], 'fecha': primera['fecha'],
            'fecha_faena_dt': primera.get('fecha_faena_dt'),
            'kg': kg_t, 'partidas': partidas,
            'es_compra_media': es_compra,
        })
    if not matches:
        return None, tropas_encontradas, 'sin matches'
    total_kg = sum(m['kg'] for m in matches) or 1
    promedio = sum(m['precio'] * m['kg'] for m in matches) / total_kg
    return ({
        'precio_promedio': round(promedio),
        'matches': matches,
        'tropas_romaneo': tropas_encontradas,
        'tropas_sin_match': sin_match,
        'kg_matcheado': total_kg,
    }, tropas_encontradas, 'ok')


def generar_nombre(parsed_data, calidad):
    fecha = parsed_data.get('fecha', '')
    if '/' in fecha:
        parts = fecha.split('/')
        dd = parts[0]; mm = parts[1]
    else:
        dd = mm = '00'
    cat = parsed_data.get('categoria', 'XX')
    code = {'Vaca': 'VA', 'Novillo': 'NO', 'Novillito': 'NT',
            'Vaquillona': 'VQ', 'Toro': 'TO', 'Bubalino': 'BU'}.get(cat, 'XX')
    medias = parsed_data.get('medias_reses', 0)
    return f"{dd}{mm} {code} {medias}m {calidad}"


# ══════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="tf-hero">
  <div>
    <h1>TF Carnes · Producción</h1>
    <p>Reporte para Gerente de Planta · Sin información comercial</p>
  </div>
  <div class="tf-hero-meta">Rendimientos · Calidad · Mermas</div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    calidad = st.selectbox("Perfil de calidad",
                           ['Standard', 'Búfalo', 'Premium Black', 'Exportación'],
                           help="Define rendimientos esperados y umbrales por calidad")
    st.markdown("---")
    st.markdown("**Umbrales aplicados:**")
    st.caption("• Rendimiento: +1% amarilla, +2.5% roja, +4% NEGRA")
    st.caption("• Amarilla: 6-10% media, >10% roja")
    st.caption("• Días faena: 4-9 amarilla, >9 roja")
    st.caption("• Recorte: +3pp amarilla, +5pp roja")
    st.caption("• Grasa+decom: 6% amarilla, 9% roja")
    st.caption("• Pesos porcionado: 1,0–1,5 kg")

# ── Tabs ──
tab_carga, tab_analisis = st.tabs([
    "📤 Cargar romaneos",
    "📄 Reportes Producción",
])

# ──────────────────────────────────────────────────────
# TAB CARGA
# ──────────────────────────────────────────────────────
with tab_carga:
    st.markdown("### Cargar romaneos")
    st.caption("Drive automático cruza con la planilla de compras y detecta días faena→producción. "
               "Manual: subí los PDFs a mano.")

    fuente = st.radio("Fuente",
                       ['📂 Drive (automático)', '📤 Subir manual'],
                       horizontal=True, key='fuente_prod')

    if fuente.startswith('📂'):
        if st.button("🔄 Listar PDFs en Drive", type="primary",
                     use_container_width=True, key='listar_drive_prod'):
            with st.spinner("Listando..."):
                drive_files = listar_pdfs_drive(CREDENTIALS_PATH)
            if not drive_files:
                st.warning("No se encontraron PDFs.")
            else:
                st.session_state.drive_files_prod = drive_files
                st.success(f"✅ {len(drive_files)} PDF(s) encontrados")

        drive_files = st.session_state.get('drive_files_prod', [])
        if drive_files:
            org = organizar_por_mes_semana(
                [{'archivo': f['name'], 'fecha': f.get('createdTime', '')[:10]
                  if f.get('createdTime') else ''}
                 for f in drive_files])
            meses = list(org.keys()) if org else []
            sel_mes = st.selectbox("Mes (opcional)", ['Todos'] + meses, key='mes_prod')
            files_filt = drive_files
            if sel_mes != 'Todos':
                nombres_mes = {p['archivo']
                               for sem in org[sel_mes].values() for p in sem}
                files_filt = [f for f in drive_files if f['name'] in nombres_mes]

            st.markdown(f"**{len(files_filt)} archivo(s)** a procesar")

            if st.button("⚙️ Procesar PDFs", type="primary",
                         use_container_width=True, key='proc_drive_prod',
                         disabled=not files_filt):
                parsed = []; remanejos = []; descartados = []
                temp_paths = {}
                progress = st.progress(0)
                for i, df in enumerate(files_filt):
                    with st.spinner(f"{df['name'][:50]}..."):
                        try:
                            tmp = descargar_pdf(CREDENTIALS_PATH, df['id'])
                            tipo = detectar_tipo_pdf(tmp)
                            if tipo == 'remanejo':
                                r = parse_remanejo_pdf(tmp); r['archivo'] = df['name']
                                remanejos.append(r); os.unlink(tmp)
                            elif tipo == 'entrada':
                                descartados.append({'a': df['name'], 'm': 'Solo entrada'})
                                os.unlink(tmp)
                            else:
                                r = parse_romaneo_pdf(tmp)
                                meat = [c for c in r.get('cortes', [])
                                        if c.get('grupo') != 'GRASA']
                                if not meat:
                                    descartados.append({'a': df['name'], 'm': 'Sin cortes'})
                                    os.unlink(tmp)
                                else:
                                    r['archivo'] = df['name']
                                    temp_paths[df['name']] = tmp
                                    parsed.append(r)
                        except Exception as e:
                            parsed.append({'archivo': df['name'], 'error': str(e)})
                    progress.progress((i + 1) / len(files_filt))

                # Aplicar correcciones
                corr = []; no_corr = []
                for p in parsed:
                    if 'error' in p: no_corr.append(p); continue
                    es_c, base = es_correccion(p.get('archivo', ''))
                    if es_c:
                        p['es_correccion'] = True; p['corrige_a'] = base
                        corr.append(p)
                    else:
                        no_corr.append(p)
                for c in corr:
                    base = c['corrige_a'].upper().replace('.PDF', '')
                    rep = False
                    for i, o in enumerate(no_corr):
                        if 'error' in o: continue
                        on = o.get('archivo', '').upper().replace('.PDF', '')
                        if base in on or on in base:
                            no_corr[i] = c; rep = True; break
                    if not rep: no_corr.append(c)
                parsed = no_corr

                # Cruzar
                with st.spinner("Cruzando con planilla de compras..."):
                    for p in parsed:
                        if 'error' in p: continue
                        pdf_tmp = temp_paths.get(p.get('archivo'))
                        cruce, tropas, status = cruzar_precio_compra(p, pdf_tmp)
                        if cruce:
                            p['precio_compra_auto'] = cruce['precio_promedio']
                            p['tropas_match'] = cruce['matches']
                            p['tropas_sin_match'] = cruce.get('tropas_sin_match', [])
                            p['precio_compra'] = cruce['precio_promedio']
                            kg_compra = sum(m.get('kg', 0) for m in cruce['matches']
                                            if m.get('es_compra_media'))
                            kg_faena = sum(m.get('kg', 0) for m in cruce['matches']
                                           if not m.get('es_compra_media'))
                            p['kg_compra_media'] = kg_compra
                            p['kg_faena_propia'] = kg_faena
                        p['cruce_status'] = status
                        if cruce and cruce.get('matches'):
                            fdt = parsear_fecha_romaneo(p.get('fecha', ''))
                            if fdt:
                                ds = []
                                for m in cruce['matches']:
                                    ff = m.get('fecha_faena_dt')
                                    if ff: ds.append((fdt - ff).days)
                                if ds:
                                    p['dias_faena_produccion'] = round(sum(ds)/len(ds), 1)

                for tp in temp_paths.values():
                    try: os.unlink(tp)
                    except Exception: pass

                st.session_state.parsed_files = parsed
                st.session_state.remanejos = remanejos
                if descartados:
                    with st.expander(f"⏭️ {len(descartados)} descartados"):
                        for d in descartados:
                            st.caption(f"- {d['a']}: {d['m']}")
                st.success(f"✅ {len([p for p in parsed if 'error' not in p])} romaneos cargados")
                st.rerun()
    else:
        uploaded = st.file_uploader("Subí los PDFs", type=['pdf'],
                                     accept_multiple_files=True, key='upload_prod')
        if uploaded and st.button("🔍 Parsear archivos", type="primary",
                                    use_container_width=True, key='parse_manual_prod'):
            parsed = []; temp_paths = {}
            progress = st.progress(0)
            for i, uf in enumerate(uploaded):
                with st.spinner(f"{uf.name}..."):
                    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                        tmp.write(uf.read()); tp = tmp.name
                    try:
                        r = parse_romaneo_pdf(tp)
                        r['archivo'] = uf.name
                        temp_paths[uf.name] = tp
                        parsed.append(r)
                    except Exception as e:
                        parsed.append({'archivo': uf.name, 'error': str(e)})
                progress.progress((i + 1) / len(uploaded))

            with st.spinner("Cruzando con planilla..."):
                for p in parsed:
                    if 'error' in p: continue
                    cruce, tropas, status = cruzar_precio_compra(p, temp_paths.get(p['archivo']))
                    if cruce:
                        p['precio_compra'] = cruce['precio_promedio']
                        p['tropas_match'] = cruce['matches']
                    p['cruce_status'] = status
                    if cruce and cruce.get('matches'):
                        fdt = parsear_fecha_romaneo(p.get('fecha', ''))
                        if fdt:
                            ds = [(fdt - m['fecha_faena_dt']).days
                                  for m in cruce['matches']
                                  if m.get('fecha_faena_dt')]
                            if ds:
                                p['dias_faena_produccion'] = round(sum(ds)/len(ds), 1)
            for tp in temp_paths.values():
                try: os.unlink(tp)
                except Exception: pass
            st.session_state.parsed_files = parsed
            st.success(f"✅ {len([p for p in parsed if 'error' not in p])} cargados")
            st.rerun()

    # Listado de archivos cargados
    if st.session_state.parsed_files:
        st.markdown("---")
        st.markdown(f"### 📋 Romaneos cargados ({len(st.session_state.parsed_files)})")
        for p in st.session_state.parsed_files:
            if 'error' in p:
                st.error(f"❌ {p['archivo']}: {p['error']}")
                continue
            kg_e = p.get('kg_entrada', 0) or 0
            kg_c = p.get('kg_carne', 0) or sum(c['kg'] for c in p.get('cortes', [])
                                                 if c.get('grupo') != 'GRASA')
            rend = (kg_c / kg_e * 100) if kg_e else 0
            cat = p.get('categoria', '?')
            medias = p.get('medias_reses', 0)
            dias = p.get('dias_faena_produccion')
            dias_str = f"· {int(dias)}d" if dias else ''
            st.markdown(
                f"📄 **{p['archivo']}** — {cat} · {medias} medias · "
                f"{kg_e:,.0f} kg → {kg_c:,.0f} kg ({rend:.1f}%) {dias_str}"
                .replace(',', '.'))

# ──────────────────────────────────────────────────────
# TAB ANÁLISIS
# ──────────────────────────────────────────────────────
with tab_analisis:
    st.markdown("### Generar reportes HTML")
    valid = [p for p in st.session_state.parsed_files if 'error' not in p]

    if not valid:
        st.info("👆 Cargá romaneos en la pestaña anterior.")
    else:
        archivos_nombres = [p.get('archivo', '?') for p in valid]
        sel = st.multiselect("Archivos (vacío = todos)", archivos_nombres,
                              key='sel_arch_prod')
        if sel:
            valid = [p for p in valid if p.get('archivo') in sel]

        st.markdown(f"**{len(valid)}** romaneo(s) seleccionado(s)")

        if len(valid) > 1:
            modo = st.radio("Modo",
                             ['Individual (uno por archivo)',
                              'Acumulado (todos juntos)',
                              'Ambos'],
                             horizontal=True, key='modo_prod')
        else:
            modo = 'Individual (uno por archivo)'

        if st.button("📄 Generar reportes HTML", type="primary",
                     use_container_width=True, key='gen_html_prod'):
            results = []
            historial = []
            try:
                if os.path.exists(HISTORIAL_PATH):
                    with open(HISTORIAL_PATH, 'r') as f:
                        historial = json.load(f)
            except Exception:
                historial = []

            with st.spinner("Generando reportes..."):
                if modo.startswith('Individual') or modo == 'Ambos':
                    for p in valid:
                        try:
                            data = {
                                'archivo': p.get('archivo'),
                                'fecha': p.get('fecha', ''),
                                'numero': p.get('numero', ''),
                                'medias_reses': p.get('medias_reses', 0),
                                'kg_entrada': p.get('kg_entrada', 0),
                                'kg_carne': p.get('kg_carne', 0),
                                'categoria': p.get('categoria', 'Vaca'),
                                'tipificacion': p.get('tipificacion', ''),
                                'precio_compra': p.get('precio_compra', 0),
                                'pct_amarilla': p.get('pct_amarilla', 0),
                                'dias_faena_produccion': p.get('dias_faena_produccion'),
                                'tropas_match': p.get('tropas_match', []),
                                'tropas_sin_match': p.get('tropas_sin_match', []),
                                'cortes': p.get('cortes', []),
                                'merma_kg': p.get('merma_kg', 0),
                                'grasa_kg': p.get('grasa_kg', 0),
                            }
                            html = build_html_produccion(data, calidad=calidad,
                                                         historial=historial)
                            results.append({
                                'name': f'{generar_nombre(p, calidad)}_PRODUCCION.html',
                                'tipo': 'Individual',
                                'html': html, 'data': p,
                            })
                        except Exception as e:
                            st.error(f"Error en {p.get('archivo')}: {e}")

                if modo.startswith('Acumulado') or modo == 'Ambos':
                    if len(valid) > 1:
                        try:
                            acum = acumular_romaneos(valid)
                            data_acum = {
                                'archivo': f'ACUMULADO_{calidad}',
                                'fecha': acum.get('fecha', ''),
                                'medias_reses': acum.get('medias_reses', 0),
                                'kg_entrada': acum.get('kg_entrada', 0),
                                'kg_carne': acum.get('kg_carne', 0),
                                'categoria': acum.get('categoria', 'Vaca'),
                                'precio_compra': acum.get('precio_compra', 0),
                                'pct_amarilla': acum.get('pct_amarilla', 0),
                                'cortes': acum.get('cortes', []),
                                'tropas_match': acum.get('tropas_match', []),
                            }
                            html = build_html_produccion(data_acum, calidad=calidad,
                                                         historial=historial,
                                                         titulo_extra='ACUMULADO')
                            results.append({
                                'name': f'ACUMULADO_{calidad}_PRODUCCION.html',
                                'tipo': 'Acumulado',
                                'html': html, 'data': acum,
                            })
                        except Exception as e:
                            st.error(f"Error acumulado: {e}")

            st.session_state.html_results = results
            if results:
                st.success(f"✅ {len(results)} reporte(s) generado(s)")

        # Resultados
        if st.session_state.html_results:
            st.markdown("---")
            st.markdown("### 📥 Reportes generados")

            # Descargar todos en un ZIP
            if len(st.session_state.html_results) > 1:
                import io, zipfile
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for r in st.session_state.html_results:
                        zf.writestr(r['name'], r['html'])
                buf.seek(0)
                st.download_button(
                    f"⬇️ Descargar TODOS en ZIP ({len(st.session_state.html_results)} archivos)",
                    buf.getvalue(),
                    file_name=f"reportes_produccion_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                    mime='application/zip',
                    use_container_width=True,
                    key='dl_zip_prod',
                )
                st.markdown("---")

            for r in st.session_state.html_results:
                d = r['data']
                kg_e = d.get('kg_entrada', 0) or 0
                meat = [c for c in d.get('cortes', []) if c.get('grupo') != 'GRASA']
                kg_c = d.get('kg_carne', 0) or sum(c['kg'] for c in meat)
                rend = (kg_c / kg_e * 100) if kg_e else 0
                cat = d.get('categoria', '?')

                with st.expander(
                    f"{'📊' if r['tipo']=='Acumulado' else '📄'} {r['name']}",
                    expanded=True,
                ):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Categoría", cat)
                    c2.metric("Kg entrada", f"{kg_e:,.0f}".replace(',', '.'))
                    c3.metric("Kg carne", f"{kg_c:,.0f}".replace(',', '.'))
                    c4.metric("Rendimiento", f"{rend:.2f}%".replace('.', ','))

                    st.download_button(
                        "⬇️ Descargar HTML",
                        r['html'].encode('utf-8'),
                        file_name=r['name'],
                        mime='text/html',
                        use_container_width=True,
                        key=f"dl_{r['name']}_{id(r)}",
                    )

# Footer
st.markdown("""
<div style="text-align:center;padding:20px;color:#999;font-size:11px;">
TF Carnes · Reporte de Producción · Sin información comercial
</div>
""", unsafe_allow_html=True)
