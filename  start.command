#!/bin/bash

# Obtener la ruta de la carpeta donde está guardado este archivo .command
DIR="$( cd "$( dirname "${BASH_SOURCE}" )" && pwd )"
cd "$DIR"

echo "=================================================="
echo "🚀 Iniciando Sistema de Conciliación de Pagos..."
echo "=================================================="

# Activar el entorno virtual de Python
source venv/bin/activate

# Forzar a Mac a abrir la pestaña en el navegador predeterminado
open "http://localhost:8501"

# Correr la aplicación de Streamlit
streamlit run src/app.py