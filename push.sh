#!/bin/bash

# Si pasas un mensaje entre comillas al ejecutar, lo usa. Si no, usa uno automático con la fecha.
MENSAJE=${1:-"🚀 Actualización del sistema: $(date +'%Y-%m-%d %H:%M')"}

echo "☁️  1. Subiendo BACKEND a Google Apps Script..."
clasp push
clasp deploy

echo "📦 2. Preparando archivos del FRONTEND para GitHub..."
git add .

echo "📝 3. Creando el paquete de cambios..."
git commit -m "$MENSAJE"

echo "☁️  4. Subiendo FRONTEND a GitHub..."
git push

echo "✅ ¡Listo! Backend y Frontend están 100% actualizados en la nube."
