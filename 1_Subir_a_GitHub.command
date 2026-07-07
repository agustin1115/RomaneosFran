#!/bin/bash
# ============================================================
#  PASO 1 - Subir la version actual de la app a GitHub
#  Doble clic. Sube tus cambios locales al repo FranCastroTFC/romaneos
#  (Streamlit Cloud lee desde ahi)
# ============================================================

cd "$(dirname "$0")" || exit 1

echo "============================================"
echo "  Subiendo app de Romaneo a GitHub"
echo "============================================"
echo ""

# Seguridad: confirmar que credentials.json y secrets NO se van a subir
echo "Chequeando que los secretos esten protegidos..."
if git check-ignore credentials.json >/dev/null 2>&1; then
    echo "  OK - credentials.json esta protegido (no se sube)"
else
    echo "  ATENCION: credentials.json NO esta en .gitignore. Abortando por seguridad."
    read -n 1 -s -r -p "Presiona una tecla para cerrar..."
    exit 1
fi
echo ""

echo "Agregando cambios..."
git add -A

echo "Creando commit..."
git commit -m "Deploy: version actual de la app + bootstrap secrets para Streamlit Cloud"

echo ""
echo "Subiendo a GitHub (puede pedir tu usuario/token de GitHub)..."
git push origin main

echo ""
if [ $? -eq 0 ]; then
    echo "============================================"
    echo "  LISTO - La app ya esta actualizada en GitHub"
    echo "  Siguiente: segui la guia DEPLOY_Streamlit_Cloud"
    echo "============================================"
else
    echo "Hubo un problema al subir. Mira el mensaje de arriba."
fi
echo ""
read -n 1 -s -r -p "Presiona una tecla para cerrar esta ventana..."
