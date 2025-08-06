from flask import Flask, render_template, request, send_file, jsonify
import os
import sys
import requests
import pandas as pd

# Obtener el directorio base de la aplicación
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)

# Configuración básica
app.config['UPLOAD_FOLDER'] = os.path.join(base_dir, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# URL de Google Sheets
GOOGLE_SHEETS_CSV_URL = 'https://docs.google.com/spreadsheets/d/1sfOY1Y3dNVCOT8zyCMzpgARv-R_jRE-S/export?format=csv'

def check_database_connection():
    """Verifica si la base de datos está disponible"""
    try:
        response = requests.get(GOOGLE_SHEETS_CSV_URL, timeout=10)
        if response.status_code == 200:
            return True, "Base de datos conectada"
        else:
            return False, f"Error HTTP: {response.status_code}"
    except Exception as e:
        return False, f"Error de conexión: {str(e)}"

@app.route('/')
def index():
    db_status, db_message = check_database_connection()
    return render_template('index.html', db_status=db_status, db_message=db_message)

@app.route('/health')
def health():
    db_status, db_message = check_database_connection()
    return jsonify({
        "status": "healthy", 
        "message": "Site Survey App is running!",
        "database": {
            "connected": db_status,
            "message": db_message
        }
    })

@app.route('/api/version')
def version():
    return jsonify({
        "app": "Site Survey App",
        "version": "1.0.0",
        "status": "deployed",
        "url": "https://github.com/efrenhernandezfangio/proyecto-hernandez"
    })

@app.route('/api/database-status')
def database_status():
    db_status, db_message = check_database_connection()
    return jsonify({
        "connected": db_status,
        "message": db_message,
        "url": GOOGLE_SHEETS_CSV_URL
    })

@app.route('/test-db')
def test_database():
    """Endpoint para probar la conexión a la base de datos"""
    try:
        df = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
        return jsonify({
            "success": True,
            "message": f"Base de datos cargada exitosamente. {len(df)} registros encontrados.",
            "columns": list(df.columns),
            "sample_data": df.head(3).to_dict('records')
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"Error al cargar la base de datos: {str(e)}",
            "url": GOOGLE_SHEETS_CSV_URL
        })

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False) 