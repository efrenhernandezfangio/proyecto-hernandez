from flask import Flask, render_template, request, send_file, jsonify
import os
import sys
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

# Archivo CSV local
CSV_FILE_PATH = os.path.join(base_dir, 'base_datos.csv')

def check_database_connection():
    """Verifica si la base de datos está disponible"""
    try:
        if os.path.exists(CSV_FILE_PATH):
            return True, "Base de datos local conectada"
        else:
            return False, "Archivo CSV no encontrado"
    except Exception as e:
        return False, f"Error de conexión: {str(e)}"

def read_csv_data():
    """Lee los datos CSV local"""
    try:
        if not os.path.exists(CSV_FILE_PATH):
            return False, [], []
            
        with open(CSV_FILE_PATH, 'r', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            rows = list(csv_reader)
            
            if len(rows) > 1:  # Al menos headers + 1 fila de datos
                headers = rows[0]
                data = rows[1:]
                return True, headers, data
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
        "file_path": CSV_FILE_PATH
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
            "file_path": CSV_FILE_PATH
        })

@app.route('/api/data')
def get_data():
    """Endpoint para obtener todos los datos"""
    success, headers, data = read_csv_data()
    
    if success:
        # Convertir a formato JSON
        json_data = []
        for row in data:
            if len(row) == len(headers):
                row_dict = dict(zip(headers, row))
                json_data.append(row_dict)
        
        return jsonify({
            "success": True,
            "total_records": len(json_data),
            "data": json_data
        })
    else:
        return jsonify({
            "success": False,
            "message": "Error al cargar los datos"
        })

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False) 