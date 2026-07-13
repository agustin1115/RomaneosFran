# App Romaneo — TF Carnes · Instrucciones de deploy (Streamlit Cloud)

## Qué es
App Streamlit que analiza romaneos de despostada: rendimiento, canales de negocio
(Consumo / China / Hilton / Búfalo), P&L, y control de **Costos Operativos**
(Frigorífico, Insumos, Fletes) comparando proyectado vs real.

## Configuración del deploy

| Parámetro | Valor |
|---|---|
| **Main file path** | **`app_v2.py`** ← IMPORTANTE (no `app.py`) |
| Python version | **3.11** |
| Dependencias | ya están en `requirements.txt` |

## Secrets (los carga el dueño de la cuenta, NO vienen en este zip)

La app lee Google Drive con una **cuenta de servicio**. En
**App → Settings → Secrets** hay que pegar un bloque con este formato:

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = """-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----
"""
client_email = "romaneos-596@tfromaneos.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "..."
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

> **El contenido real de las credenciales NO está en este zip a propósito.**
> Se lo pasa Fran directamente por un canal seguro (nunca por WhatsApp/mail),
> o lo pega él mismo en el panel de Streamlit.
>
> La app escribe ese secreto a `credentials.json` en runtime automáticamente
> (ya está el bootstrap en `app_v2.py`), así que no hay que hacer nada más.

## Requisito en Google Drive
Los archivos que la app lee deben estar compartidos con el email de la cuenta de
servicio (`romaneos-596@tfromaneos.iam.gserviceaccount.com`):
- Carpeta de romaneos (PDFs)
- Planilla de precios de facturación
- Planilla de compras
- `TF CARNES 30-6-26.xlsx` (cuenta corriente del frigorífico)
- `Stock 27-5-2026 Etiquetas y Embalajes.xlsx` (insumos)
- `COSTO DE FLETES` (fletes)

## Privacidad
Dejar la app **privada** (Settings → Sharing) y habilitar solo los emails
autorizados. La app muestra costos, márgenes y precios de compra.

## Nota técnica
En Streamlit Cloud el disco es efímero: `segmentos.json` (canales confirmados) e
`historial_costos.json` (serie de costos) **no persisten entre reinicios**. Para
que persistan hay que moverlos a Drive o a una base — pendiente.
