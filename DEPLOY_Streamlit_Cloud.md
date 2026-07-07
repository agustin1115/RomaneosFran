# Publicar la App de Romaneo en internet (para abrirla en Chrome)

Objetivo: que tu app abra en Chrome desde una URL fija (ej. `romaneos-tfc.streamlit.app`),
sin Terminal, sin Python, con la conexión a Drive funcionando igual que ahora.

Lo hacés **una sola vez**. Después es solo un favorito en Chrome.

---

## PASO 1 — Subir tu versión actual a GitHub

Tu app en GitHub está vieja: tenés cambios locales sin subir.

➡️ **Doble clic en `1_Subir_a_GitHub.command`** (está en esta misma carpeta).

Se abre una ventana negra (Terminal), sube todo y se cierra cuando termina.
Si te pide usuario/contraseña de GitHub, es tu cuenta `FranCastroTFC`.

> Tus secretos (`credentials.json`) NO se suben — el script lo verifica antes.

---

## PASO 2 — Crear la cuenta de Streamlit (gratis)

1. Andá a **https://share.streamlit.io**
2. Clic en **Continue with GitHub** → logueate con tu cuenta de GitHub (`FranCastroTFC`).
3. Autorizá el acceso cuando te lo pida.

---

## PASO 3 — Crear la app

1. Clic en **Create app** (arriba a la derecha).
2. Elegí **"Deploy a public app from GitHub"**.
3. Completá:
   - **Repository:** `FranCastroTFC/romaneos`
   - **Branch:** `main`
   - **Main file path:** `app_v2.py`   ← el que tiene Costos Operativos, segmentación, China, etc.
   - **App URL:** elegí un nombre, ej. `romaneos-tfc` → queda `romaneos-tfc.streamlit.app`
4. **NO le des Deploy todavía.** Primero los secretos (Paso 4).

---

## PASO 4 — Pegar los secretos (la conexión a Drive)

1. En esa misma pantalla, clic en **Advanced settings**.
2. En **Python version** elegí **3.11**.
3. En el cuadro grande de **Secrets**, pegá TODO el contenido del archivo
   **`secrets_PARA_PEGAR.toml`** (está en esta carpeta — abrilo con TextEdit,
   seleccioná todo con ⌘A, copiá con ⌘C, y pegá ahí).
4. Clic en **Save**.

> Ese archivo tiene la llave de la cuenta de servicio que hoy te conecta a Drive.
> Por eso la app en la nube va a leer los romaneos igual que en tu compu.

---

## PASO 5 — Deploy

1. Clic en **Deploy**.
2. Esperá 2-3 minutos (instala todo solo la primera vez).
3. Cuando termina, te muestra tu app andando. ✅

---

## PASO 6 — Guardar en Chrome

1. Con la app abierta en Chrome, clic en la **estrella** de la barra de direcciones
   → guardar como favorito.
2. Listo: cada vez que quieras la app, clic al favorito. Sin Terminal, sin nada.

---

## Cómo se actualiza de ahora en más

- **Datos (romaneos / precios):** se actualizan solos. La app lee Drive y Google Sheets
  en vivo cada vez que la abrís. No tenés que hacer nada.
- **Cambios en la app (código):** si en el futuro modificás la app, volvés a hacer
  doble clic en `1_Subir_a_GitHub.command` y Streamlit Cloud se actualiza solo en 1-2 min.

---

## Datos técnicos (por si los necesitás)

- **Repo:** https://github.com/FranCastroTFC/romaneos
- **Archivo principal:** `app_v2.py`
- **Cuenta de servicio (Drive):** `romaneos-596@tfromaneos.iam.gserviceaccount.com`
  La carpeta de Drive de romaneos ya está compartida con este email (por eso funciona).
  Si alguna vez deja de leer Drive, verificá que esa carpeta siga compartida con ese email.
- **Carpeta de Drive (ID):** `1PpSEKdjQGk3PmU4TAZc6JqngSivvcLwz`

---

## Si algo falla

| Síntoma | Solución |
|---|---|
| "Falta credentials.json" en la app | Los secrets no se pegaron bien. App > Settings > Secrets, repegá `secrets_PARA_PEGAR.toml` completo. |
| Error al instalar / "ModuleNotFound" | Revisá que `requirements.txt` esté en el repo (ya está). Reboot app desde el menú. |
| No lee Drive | La carpeta dejó de estar compartida con la cuenta de servicio. Compartila de nuevo con el email de arriba (rol: Lector). |
| El push falla en el Paso 1 | Tu token de GitHub venció. Generá uno nuevo en GitHub > Settings > Developer settings > Tokens. |
