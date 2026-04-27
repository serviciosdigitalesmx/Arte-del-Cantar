#!/bin/bash

# Si pasas un mensaje entre comillas al ejecutar, lo usa. Si no, usa uno automático con la fecha.
MENSAJE=${1:-"🚀 Actualización del sistema: $(date +'%Y-%m-%d %H:%M')"}

echo "📦 Preparando todos los archivos..."
git add .

echo "📝 Creando el paquete de cambios..."
git commit -m "$MENSAJE"

echo "☁️  Subiendo a GitHub..."
git push

echo "✅ ¡Listo! Todos los cambios están en la nube."
