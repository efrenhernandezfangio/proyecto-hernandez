from flask import Flask, render_template, request, send_file, jsonify
import os
import sys
import requests
import csv
import io

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
GOOGLE_SHEETS_CSV_URL = 'https://docs.google.com/spreadsheets/d/1sfOY1Y3dNVCOT8zyCMzpgARv-R_jRE-S/export?format=csv&gid=0'

def check_database_connection():
    """Verifica si la base de datos está disponible con debug detallado"""
    try:
        print(f"Intentando conectar a: {GOOGLE_SHEETS_CSV_URL}")
        response = requests.get(GOOGLE_SHEETS_CSV_URL, timeout=10)
        print(f"Código de respuesta: {response.status_code}")
        print(f"Headers: {dict(response.headers)}")
        
        if response.status_code == 200:
            content = response.content[:500]  # Primeros 500 caracteres
            print(f"Contenido recibido: {content}")
            return True, "Base de datos conectada"
        else:
            return False, f"Error HTTP: {response.status_code} - {response.text[:200]}"
    except Exception as e:
        print(f"Excepción: {str(e)}")
        return False, f"Error de conexión: {str(e)}"

def read_csv_data():
    """Lee los datos CSV sin usar pandas"""
    try:
        response = requests.get(GOOGLE_SHEETS_CSV_URL, timeout=10)
        if response.status_code == 200:
            # Decodificar el contenido
            content = response.content.decode('utf-8')
            print(f"Contenido CSV recibido: {content[:500]}...")
            
            csv_reader = csv.reader(io.StringIO(content))
            
            # Leer todas las filas
            rows = list(csv_reader)
            print(f"Filas leídas: {len(rows)}")
            
            if len(rows) > 1:  # Al menos headers + 1 fila de datos
                headers = rows[0]
                data = rows[1:]
                print(f"Headers: {headers}")
                print(f"Datos: {len(data)} filas")
                return True, headers, data
            else:
                return False, [], []
        else:
            return False, [], []
    except Exception as e:
        print(f"Error leyendo CSV: {str(e)}")
        return False, [], []

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
    success, headers, data = read_csv_data()
    
    if success:
        return jsonify({
            "success": True,
            "message": f"Base de datos cargada exitosamente. {len(data)} registros encontrados.",
            "columns": headers,
            "sample_data": data[:3] if data else []
        })
    else:
        return jsonify({
            "success": False,
            "message": "Error al cargar la base de datos",
            "url": GOOGLE_SHEETS_CSV_URL
        })

@app.route('/debug')
def debug():
    """Endpoint de debug completo"""
    try:
        response = requests.get(GOOGLE_SHEETS_CSV_URL, timeout=10)
        return jsonify({
            "url": GOOGLE_SHEETS_CSV_URL,
            "status_code": response.status_code,
            "headers": dict(response.headers),
            "content_length": len(response.content),
            "content_preview": response.text[:1000],
            "encoding": response.encoding
        })
    except Exception as e:
        return jsonify({
            "error": str(e),
            "url": GOOGLE_SHEETS_CSV_URL
        })

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False) 