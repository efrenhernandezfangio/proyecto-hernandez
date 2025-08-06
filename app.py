# --- EJEMPLO DE ENDPOINT PARA RENDERIZAR CHECKBOXES DINÁMICOS ---
# Agrega este endpoint o adapta el tuyo para pasar los valores correctos a la plantilla
# --- FIN DEL EJEMPLO ---
import os
import time
import pandas as pd
import xlwings as xw
import win32com.client
from flask import Flask, request, send_file, render_template_string, redirect, url_for, after_this_request, jsonify
from werkzeug.utils import secure_filename
import sys
import re
import dataframe_image as dfi
import matplotlib.pyplot as plt
import textwrap
import unicodedata
import glob
import subprocess

def normaliza_na(valor):
    if isinstance(valor, str) and valor.strip().lower() == "n/a":
        return "N/A"
    elif pd.isna(valor):
        return "N/A"
    elif valor == "" or (isinstance(valor, str) and valor.strip() == ""):
        return "N/A"
    return valor

# Obtener el directorio base de la aplicación
if getattr(sys, 'frozen', False):
    # Si es un ejecutable compilado
    base_dir = os.path.dirname(sys.executable)
else:
    # Si es código Python normal
    base_dir = os.path.dirname(os.path.abspath(__file__))

print(f"Directorio base: {base_dir}")
print("Archivos en el directorio:", os.listdir(base_dir))

# Buscar el archivo en múltiples ubicaciones
llenado_paths = [
    os.path.join(base_dir, 'llenado-automatico.html'),
    os.path.join(base_dir, 'static', 'llenado-automatico.html'),
    os.path.join(base_dir, 'templates', 'llenado-automatico.html'),
    'llenado-automatico.html'  # Directorio actual como fallback
]

html_form = None
for path in llenado_paths:
    try:
        with open(path, encoding='utf-8') as f:
            html_form = f.read()
        print(f"Archivo llenado-automatico.html cargado desde: {path}")
        break
    except Exception as e:
        print(f"No se pudo cargar desde {path}: {e}")
        continue

if html_form is None:
    print("ERROR: No se pudo cargar llenado-automatico.html desde ninguna ubicación")
    print("Ubicaciones probadas:")
    for path in llenado_paths:
        print(f"  - {path}")
    print("Asegúrate de que el archivo existe en una de estas ubicaciones")
    input("Presiona Enter para salir...")
    sys.exit(1)
app = Flask(__name__)
# Usar ruta relativa para que funcione en cualquier computadora
import os
base_dir = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(base_dir, 'site_survey')
GOOGLE_SHEETS_CSV_URL = 'https://docs.google.com/spreadsheets/d/1sfOY1Y3dNVCOT8zyCMzpgARv-R_jRE-S/export?format=csv'
@app.route('/site_survey_checkboxes', methods=['GET'])
def site_survey_checkboxes():
    import pandas as pd
    user_id = request.args.get('user_id', '')
    fila_idx = request.args.get('fila_idx', '')
    chk_urbana = chk_suburbana = chk_rural = chk_ejidal = chk_pueblo_magico = False
    if fila_idx:
        try:
            df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
            row = df_db.loc[int(fila_idx)]
            tipo_zona_original = row.get('Tipo de Zona', '')
            tipo_zona = normaliza_texto(tipo_zona_original)
            chk_urbana = 'urbana' in tipo_zona
            chk_suburbana = 'suburbana' in tipo_zona or 'suburbana' in tipo_zona or 'suburbana' in tipo_zona.replace('sub', '')
            chk_rural = 'rural' in tipo_zona
            chk_ejidal = 'ejidal' in tipo_zona
            chk_pueblo_magico = 'pueblomagico' in tipo_zona
        except Exception as e:
            print(f"Error leyendo base de datos: {e}")
    return render_template(
        'site_survey_checkboxes.html',
        chk_urbana=chk_urbana,
        chk_suburbana=chk_suburbana,
        chk_rural=chk_rural,
        chk_ejidal=chk_ejidal,
        chk_pueblo_magico=chk_pueblo_magico
    )

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# Usar ruta relativa para que funcione en cualquier computadora
TEMPLATE_PATH = os.path.join(base_dir, 'Temp', 'plantillas', 'llenadoauto.xlsx')


from flask import render_template

@app.route('/')
def index():
    import pandas as pd
    db_status = 'ok'
    db_error = ''
    try:
        df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
    except Exception as e:
        db_status = 'error'
        db_error = str(e)
    return render_template('index.html', db_status=db_status, db_error=db_error)
from flask import render_template, redirect, url_for

@app.route('/diseno_solucion', methods=['GET', 'POST'])
def diseno_solucion():
    import pandas as pd
    if request.method == 'GET':
        user_id = request.args.get('user_id')
        fila_idx = request.args.get('fila_idx')
        if user_id and fila_idx:
            try:
                df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
                row = df_db.loc[int(fila_idx)]
                id_sitio = row.get('ID', '') or ''
                sitio_a = row.get('Nombre del sitio A', '') or ''
                sitio_b = row.get('Nombre del sitio B', '') or ''
                if pd.isna(sitio_a): sitio_a = ''
                if pd.isna(sitio_b): sitio_b = ''
            except Exception as e:
                return f"Error leyendo Google Sheets: {e}"
            html_form_mod = html_form.replace(
                '<form id="autoForm" action="/procesar" method="post" enctype="multipart/form-data" autocomplete="off">',
                '<form id="autoForm" action="/procesar" method="post" enctype="multipart/form-data" autocomplete="off">' +
                f'\n<input type="hidden" name="user_id" value="{user_id}">' +
                f'\n<input type="hidden" name="fila_idx" value="{fila_idx}">' +
                f'\n<input type="hidden" name="template_path" value="static/plantillas/llenadoauto.xlsx">'
            )
            # Cambiar el tipo de site_survey a diseno_solucion
            html_form_mod = html_form_mod.replace(
                'name="tipo" value="site_survey"',
                'name="tipo" value="diseno_solucion"'
            )
            mensaje_plantilla = '''
<div class="success-box">
    <i class="fa-solid fa-file-excel"></i>
    Plantilla de llenado cargada correctamente
</div>
'''
            html_form_mod = html_form_mod.replace(
                '<!-- MENSAJE_PLANTILLA_AQUI -->',
                mensaje_plantilla + '\n<!-- MENSAJE_PLANTILLA_AQUI -->'
            )
            html_form_mod = html_form_mod.replace(
                '<!-- ANALISIS_AQUI -->',
                f'''
<div class="analisis-info-box">
    <div class="analisis-title"><i class="fa-solid fa-circle-info"></i> Enlace seleccionado</div>
    <div class="analisis-row"><b style="color:#00c3ff;">ID:</b> {id_sitio}</div>
    <div class="analisis-row"><b style="color:#00c3ff;">Sitio A:</b> {sitio_a}</div>
    <div class="analisis-row"><b style="color:#00c3ff;">Sitio B:</b> {sitio_b}</div>
</div>
''')
            return render_template_string(html_form_mod, plantilla_cargada=True)
        return render_template('fallback_id.html', titulo='Diseño de Solución', mensaje_error='Faltan parámetros para cargar el registro. Por favor, vuelve al inicio e ingresa tu ID.', mostrar_form=True, ruta_form='/diseno_solucion', placeholder='Ingresa tu ID')
    return '', 204

@app.route('/site_survey', methods=['GET'])
def site_survey():
    user_id = request.args.get('user_id')
    fila_idx = request.args.get('fila_idx')
    tipo_documento = request.args.get('tipo', 'ptp')  # Por defecto PtP

    # Recupera los datos de la fila para mostrar los nombres
    import pandas as pd
    df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
    row = df_db.loc[int(fila_idx)]
    nombre_a = row.get('Nombre del sitio A', '')
    nombre_b = row.get('Nombre del sitio B', '')

    html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>FANGIO TELECOM | Documento Generado</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
        <style>
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}

            body {{
                font-family: 'Montserrat', Arial, sans-serif;
                min-height: 100vh;
                background-color: #0a192f;
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
                background-repeat: no-repeat;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                position: relative;
                overflow-x: hidden;
                color: #e0e7ef;
            }}

            /* Overlay para mejorar la legibilidad */
            body::before {{
                content: '';
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: 
                    radial-gradient(ellipse at top, rgba(0, 0, 0, 0.2) 0%, transparent 50%),
                    radial-gradient(ellipse at bottom, rgba(0, 0, 0, 0.4) 0%, transparent 50%);
                z-index: 1;
                pointer-events: none;
            }}

            /* Efecto de estrellas sutiles */
            body::after {{
                content: '';
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background-image: 
                    radial-gradient(1px 1px at 10% 20%, rgba(255, 255, 255, 0.8), transparent),
                    radial-gradient(1px 1px at 20% 80%, rgba(255, 255, 255, 0.6), transparent),
                    radial-gradient(1px 1px at 80% 30%, rgba(255, 255, 255, 0.9), transparent),
                    radial-gradient(1px 1px at 90% 70%, rgba(255, 255, 255, 0.7), transparent);
                background-size: 400px 400px, 300px 300px, 500px 500px, 350px 350px;
                animation: twinkle 8s ease-in-out infinite;
                z-index: 2;
                pointer-events: none;
            }}

            @keyframes twinkle {{
                0%, 100% {{ opacity: 0.3; }}
                50% {{ opacity: 1; }}
            }}

            .header {{
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                background: linear-gradient(135deg, 
                    rgba(10, 15, 30, 0.98) 0%, 
                    rgba(22, 33, 62, 0.95) 50%, 
                    rgba(15, 52, 96, 0.92) 100%);
                backdrop-filter: blur(20px);
                border-bottom: 3px solid rgba(0, 195, 255, 0.5);
                box-shadow: 
                    0 8px 32px rgba(0, 0, 0, 0.4),
                    0 0 0 1px rgba(0, 195, 255, 0.1) inset;
                padding: 0;
                z-index: 100;
                position: relative;
                overflow: hidden;
            }}

            .header::before {{
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                height: 1px;
                background: linear-gradient(90deg, 
                    transparent, 
                    rgba(0, 195, 255, 0.8), 
                    transparent);
                animation: headerScan 3s ease-in-out infinite;
            }}

            @keyframes headerScan {{
                0%, 100% {{ transform: translateX(-100%); opacity: 0.5; }}
                50% {{ transform: translateX(100%); opacity: 1; }}
            }}

            .header-content {{
                max-width: 1400px;
                margin: 0 auto;
                display: flex;
                align-items: center;
                justify-content: space-between;
                padding: 15px 30px;
                position: relative;
                z-index: 10;
            }}

            .logo-container {{
                display: flex;
                align-items: center;
                gap: 15px;
            }}

            .logo {{
                transition: all 0.3s ease;
            }}

            .logo img {{
                height: 40px;
                width: auto;
                filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.3));
            }}

            .logo-text {{
                color: #e0e7ef;
                font-size: 0.9rem;
                font-weight: 500;
                opacity: 0.9;
                text-shadow: 0 0 10px rgba(0, 195, 255, 0.3);
            }}

            .main-container {{
                flex: 1;
                display: flex;
                align-items: center;
                justify-content: center;
                padding: 20px;
                position: relative;
                z-index: 10;
                width: 100%;
                min-height: 100vh;
            }}

            .success-card {{
                background: linear-gradient(135deg, rgba(22, 33, 62, 0.95) 0%, rgba(15, 52, 96, 0.9) 100%);
                backdrop-filter: blur(20px);
                border: 2px solid rgba(0, 195, 255, 0.4);
                box-shadow: 
                    0 20px 60px rgba(0, 0, 0, 0.4),
                    0 8px 32px rgba(0, 195, 255, 0.3),
                    0 0 0 1px rgba(0, 195, 255, 0.2) inset,
                    0 0 0 4px rgba(0, 195, 255, 0.1) inset,
                    0 0 50px rgba(0, 195, 255, 0.2);
                border-radius: 28px;
                padding: 60px 50px;
                min-width: 500px;
                max-width: 600px;
                width: 100%;
                text-align: center;
                position: relative;
                overflow: hidden;
                z-index: 10;
                animation: cardFloat 6s ease-in-out infinite;
            }}

            @keyframes cardFloat {{
                0%, 100% {{ transform: translateY(0px); }}
                50% {{ transform: translateY(-10px); }}
            }}

            .success-card::before {{
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                height: 3px;
                background: linear-gradient(90deg, transparent, #00c3ff, #00e0ff, #00c3ff, transparent);
                animation: scan 4s ease-in-out infinite;
                box-shadow: 0 0 10px rgba(0, 195, 255, 0.8);
            }}

            .success-card::after {{
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: radial-gradient(circle at 30% 20%, rgba(0, 195, 255, 0.1) 0%, transparent 50%),
                            radial-gradient(circle at 70% 80%, rgba(0, 195, 255, 0.05) 0%, transparent 50%);
                pointer-events: none;
                z-index: -1;
            }}

            @keyframes scan {{
                0%, 100% {{ transform: translateX(-100%); opacity: 0.5; }}
                50% {{ transform: translateX(100%); opacity: 1; }}
            }}

            .success-icon {{
                width: 100px;
                height: 100px;
                border-radius: 50%;
                background: linear-gradient(135deg, #00ff88 0%, #00cc6a 50%, #00ff88 100%);
                border: 4px solid #00c3ff;
                box-shadow: 
                    0 0 40px rgba(0, 255, 136, 0.6),
                    0 0 0 1px rgba(255, 255, 255, 0.2) inset,
                    0 8px 25px rgba(0, 0, 0, 0.2);
                margin: 0 auto 30px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 3rem;
                color: #fff;
                font-weight: bold;
                position: relative;
                animation: successPulse 2s ease-in-out infinite;
            }}

            @keyframes successPulse {{
                0%, 100% {{ transform: scale(1); box-shadow: 0 0 40px rgba(0, 255, 136, 0.6); }}
                50% {{ transform: scale(1.05); box-shadow: 0 0 50px rgba(0, 255, 136, 0.8); }}
            }}

            .success-title {{
                color: #00ff88;
                font-size: 2.5rem;
                font-weight: 800;
                margin: 0 0 20px 0;
                letter-spacing: 2px;
                text-shadow: 0 0 20px rgba(0, 255, 136, 0.5);
            }}

            .info-container {{
                background: rgba(0, 195, 255, 0.1);
                border: 1px solid rgba(0, 195, 255, 0.3);
                border-radius: 16px;
                padding: 25px;
                margin: 30px 0;
                text-align: left;
            }}

            .info-item {{
                display: flex;
                align-items: center;
                gap: 10px;
                margin-bottom: 15px;
                color: #e0e7ef;
                font-size: 1.1rem;
            }}

            .info-item:last-child {{
                margin-bottom: 0;
            }}

            .info-label {{
                color: #00c3ff;
                font-weight: 600;
                min-width: 120px;
                display: flex;
                align-items: center;
                gap: 8px;
            }}

            .info-value {{
                color: #e0e7ef;
                font-weight: 500;
            }}

            .buttons-container {{
                display: flex;
                flex-direction: column;
                gap: 15px;
                margin-top: 30px;
            }}

            .download-button, .upload-images-button, .new-document-button, .back-button, .other-button {{
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
                padding: 18px 30px;
                border-radius: 12px;
                text-decoration: none;
                font-weight: 600;
                font-size: 1.1rem;
                transition: all 0.3s ease;
                border: none;
                cursor: pointer;
                position: relative;
                overflow: hidden;
            }}

            .download-button {{
                background: linear-gradient(135deg, #00c3ff 0%, #00e0ff 100%);
                color: #192133;
                box-shadow: 0 8px 25px rgba(0, 195, 255, 0.4);
            }}

            .download-button:hover {{
                background: linear-gradient(135deg, #00e0ff 0%, #00c3ff 100%);
                transform: translateY(-2px);
                box-shadow: 0 12px 35px rgba(0, 195, 255, 0.6);
            }}

            .back-button {{
                background: rgba(0, 195, 255, 0.1);
                color: #00c3ff;
                border: 2px solid rgba(0, 195, 255, 0.3);
            }}

            .back-button:hover {{
                background: rgba(0, 195, 255, 0.2);
                border-color: rgba(0, 195, 255, 0.5);
                transform: translateY(-2px);
            }}

            .other-button {{
                background: linear-gradient(135deg, #4a90e2 0%, #357abd 100%);
                color: #ffffff;
                box-shadow: 0 8px 25px rgba(74, 144, 226, 0.4);
            }}

            .other-button:hover {{
                background: linear-gradient(135deg, #357abd 0%, #4a90e2 100%);
                transform: translateY(-2px);
                box-shadow: 0 12px 35px rgba(74, 144, 226, 0.6);
            }}

            .upload-images-button {{
                background: linear-gradient(135deg, #9b59b6 0%, #8e44ad 100%);
                color: #ffffff;
                box-shadow: 0 8px 25px rgba(155, 89, 182, 0.4);
            }}

            .upload-images-button:hover {{
                background: linear-gradient(135deg, #8e44ad 0%, #9b59b6 100%);
                transform: translateY(-2px);
                box-shadow: 0 12px 35px rgba(155, 89, 182, 0.6);
            }}

            .footer {{
                background: linear-gradient(135deg, 
                    rgba(10, 15, 30, 0.98) 0%, 
                    rgba(22, 33, 62, 0.95) 100%);
                backdrop-filter: blur(20px);
                border-top: 2px solid rgba(0, 195, 255, 0.3);
                padding: 20px 0;
                position: relative;
                z-index: 10;
            }}

            .footer-content {{
                max-width: 1400px;
                margin: 0 auto;
                text-align: center;
                color: #e0e7ef;
                font-size: 0.9rem;
                opacity: 0.8;
            }}

            .separator {{
                margin: 0 10px;
                color: #00c3ff;
            }}

            .company-name {{
                color: #00c3ff;
                font-weight: 600;
            }}

            @media (max-width: 768px) {{
                .success-card {{
                    min-width: auto;
                    max-width: 90%;
                    padding: 40px 30px;
                }}

                .success-title {{
                    font-size: 2rem;
                }}

                .info-item {{
                    flex-direction: column;
                    align-items: flex-start;
                    gap: 5px;
                }}

                .info-label {{
                    min-width: auto;
                }}
            }}
        </style>
    </head>
    <body>
        <header class="header">
            <div class="header-content">
                <div class="logo-container">
                    <div class="logo">
                        <img src="{{ url_for('static', filename='images/fangio-logo.svg') }}?v={{ range(1, 1000) | random }}" alt="FANGIO TELECOM">
                    </div>
                    <div class="logo-text">
                        <p>Redes Seguras Soluciones Estratégicas</p>
                    </div>
                </div>
            </div>
        </header>

        <div class="main-container">
            <div class="success-card">
                <div class="success-icon">
                    <i class="fas fa-check"></i>
                </div>
                <h1 class="success-title">¡Documento Generado!</h1>
                
                <div class="info-container">
                    <div class="info-item">
                        <span class="info-label"><i class="fas fa-id-card"></i> ID:</span>
                        <span class="info-value">{user_id}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label"><i class="fas fa-map-marker-alt"></i> Sitio A:</span>
                        <span class="info-value">{nombre_a}</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label"><i class="fas fa-map-marker-alt"></i> Sitio B:</span>
                        <span class="info-value">{nombre_b}</span>
                    </div>
                </div>
                
                <div class="buttons-container">
                    <a href="{url_for('descargar_site_survey_ptmp' if tipo_documento == 'ptmp' else 'descargar_site_survey', user_id=user_id)}" class="download-button">
                        <i class="fas fa-download"></i>
                        Descargar Archivo Generado {'(PtMP)' if tipo_documento == 'ptmp' else '(PtP)'}
                    </a>
                    
                    {'<a href="' + url_for('subir_imagenes_ptp', user_id=user_id, fila_idx=fila_idx) + '" class="upload-images-button">' if tipo_documento == 'ptp' else ''}
                    {'<i class="fas fa-images"></i>' if tipo_documento == 'ptp' else ''}
                    {'Subir Imágenes - Planos A' if tipo_documento == 'ptp' else ''}
                    {'</a>' if tipo_documento == 'ptp' else ''}
                    
                    {'<a href="' + url_for('subir_imagenes_ptp_planos_b', user_id=user_id, fila_idx=fila_idx) + '" class="upload-images-button">' if tipo_documento == 'ptp' else ''}
                    {'<i class="fas fa-images"></i>' if tipo_documento == 'ptp' else ''}
                    {'Subir Imágenes - Planos B' if tipo_documento == 'ptp' else ''}
                    {'</a>' if tipo_documento == 'ptp' else ''}
                    
                    {'<a href="' + url_for('subir_imagenes_ptp_fotos_a', user_id=user_id, fila_idx=fila_idx) + '" class="upload-images-button">' if tipo_documento == 'ptp' else ''}
                    {'<i class="fas fa-images"></i>' if tipo_documento == 'ptp' else ''}
                    {'Subir Imágenes - Reporte Fotos A' if tipo_documento == 'ptp' else ''}
                    {'</a>' if tipo_documento == 'ptp' else ''}
                    
                    {'<a href="' + url_for('subir_imagenes_ptp_fotos_b', user_id=user_id, fila_idx=fila_idx) + '" class="upload-images-button">' if tipo_documento == 'ptp' else ''}
                    {'<i class="fas fa-images"></i>' if tipo_documento == 'ptp' else ''}
                    {'Subir Imágenes - Reporte Fotos B' if tipo_documento == 'ptp' else ''}
                    {'</a>' if tipo_documento == 'ptp' else ''}
                    
                    <a href="{url_for('reporte_planeacion', user_id=user_id, fila_idx=fila_idx)}" class="other-button">
                        <i class="fas fa-chart-line"></i>
                        Ir a Reporte de Planeación
                    </a>
                    
                    <a href="{url_for('formulario_archivos', user_id=user_id, fila_idx=fila_idx)}" class="other-button">
                        <i class="fas fa-file-upload"></i>
                        Ir a Diseño de Solución
                    </a>
                    
                    <a href="/" class="back-button">
                        <i class="fas fa-home"></i>
                        Volver al Inicio
                    </a>
                </div>
            </div>
        </div>

        <footer class="footer">
            <div class="footer-content">
                <span>&copy; 2025 Realizado por Efren Alexis Hernandez Mendez</span>
                <span class="separator">|</span>
                <span class="company-name">FANGIO TELECOM</span>
            </div>
        </footer>

        <script>
            // Verificar carga de imágenes
            document.addEventListener('DOMContentLoaded', function() {{
                // Verificar imagen de fondo
                const bgImg = new Image();
                bgImg.onload = function() {{
                    console.log('Imagen de fondo cargada correctamente');
                    document.body.style.backgroundImage = 'linear-gradient(rgba(0, 0, 0, 0.4), rgba(0, 0, 0, 0.6)), url("{{ url_for("static", filename="images/earth-background.jpg") }}?v={{ range(1, 1000) | random }}")';
                    document.body.style.backgroundSize = 'cover';
                    document.body.style.backgroundPosition = 'center';
                    document.body.style.backgroundAttachment = 'fixed';
                    document.body.style.backgroundRepeat = 'no-repeat';
                }};
                bgImg.onerror = function() {{
                    console.log('Error al cargar imagen de fondo - usando color de respaldo');
                    document.body.style.backgroundColor = '#0a192f';
                    // Intentar con ruta alternativa
                    const altBgImg = new Image();
                    altBgImg.onload = function() {{
                        console.log('Imagen de fondo cargada con ruta alternativa');
                        document.body.style.backgroundImage = 'linear-gradient(rgba(0, 0, 0, 0.4), rgba(0, 0, 0, 0.6)), url("/static/images/earth-background.jpg")';
                        document.body.style.backgroundSize = 'cover';
                        document.body.style.backgroundPosition = 'center';
                        document.body.style.backgroundAttachment = 'fixed';
                        document.body.style.backgroundRepeat = 'no-repeat';
                    }};
                    altBgImg.onerror = function() {{
                        console.log('Todas las rutas de imagen fallaron');
                    }};
                    altBgImg.src = '/static/images/earth-background.jpg';
                }};
                bgImg.src = '{{ url_for("static", filename="images/earth-background.jpg") }}?v={{ range(1, 1000) | random }}';

                // Verificar logo
                const logoImg = document.querySelector('.logo img');
                if (logoImg) {{
                    logoImg.onerror = function() {{
                        console.log('Error al cargar logo');
                        this.style.display = 'none';
                    }};
                    logoImg.onload = function() {{
                        console.log('Logo cargado correctamente');
                    }};
                }}

                // Efecto de aparición para la tarjeta principal
                const successCard = document.querySelector('.success-card');
                if (successCard) {{
                    successCard.style.opacity = '0';
                    successCard.style.transform = 'translateY(30px)';
                    
                    setTimeout(() => {{
                        successCard.style.transition = 'all 0.8s ease';
                        successCard.style.opacity = '1';
                        successCard.style.transform = 'translateY(0)';
                    }}, 300);
                }}

                // Efecto de hover para el logo
                const logo = document.querySelector('.logo');
                if (logo) {{
                    logo.addEventListener('mouseenter', function() {{
                        this.style.transform = 'translateY(-3px) scale(1.02)';
                    }});
                    
                    logo.addEventListener('mouseleave', function() {{
                        this.style.transform = 'translateY(0) scale(1)';
                    }});
                }}
            }});
        </script>
    </body>
    </html>
    """
    return render_template_string(html)

@app.route('/descargar_site_survey')
def descargar_site_survey():
    user_id = request.args.get('user_id')
    def limpiar_nombre_archivo(nombre):
        return re.sub(r'[^a-zA-Z0-9_-]', '', str(nombre))
    user_id_limpio = limpiar_nombre_archivo(user_id)
    # Usar ruta relativa para que funcione en cualquier computadora
    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}.xlsx')
    
    # Si no existe el archivo sin timestamp, buscar con timestamp
    if not os.path.exists(output_path):
        import glob
        patron_archivo = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}*.xlsx')
        archivos_encontrados = glob.glob(patron_archivo)
        
        if archivos_encontrados:
            # Usar el archivo más reciente
            archivo_mas_reciente = max(archivos_encontrados, key=os.path.getctime)
            output_path = archivo_mas_reciente
            print(f"DEBUG: Descargando archivo encontrado: {output_path}")
        else:
            print(f"Archivo no encontrado para descargar: {output_path}")
            return "El archivo ya no está disponible. Por favor, genera uno nuevo."

    @after_this_request
    def eliminar_archivos_temporales(response):
        try:
            # Borra el Excel generado
            if os.path.exists(output_path):
                os.remove(output_path)
                print(f"Archivo eliminado: {output_path}")
            # Borra otros archivos temporales relacionados con el user_id
            patron = os.path.join(base_dir, 'site_survey', f'*{user_id_limpio}*.*')
            for archivo in glob.glob(patron):
                try:
                    if archivo != output_path:  # Ya se eliminó arriba
                        os.remove(archivo)
                        print(f"Archivo temporal eliminado: {archivo}")
                except Exception as e:
                    print(f"Error al eliminar archivo temporal: {archivo} - {e}")
        except Exception as e:
            print(f"Error al eliminar archivos: {e}")
        return response

    return send_file(output_path, as_attachment=True)

@app.route('/descargar_site_survey_ptmp')
def descargar_site_survey_ptmp():
    import os
    import re
    user_id = request.args.get('user_id', '')
    if not user_id:
        return "Falta el ID"
    
    def limpiar_nombre_archivo(nombre):
        # Eliminar caracteres problemáticos para nombres de archivo
        nombre_limpio = re.sub(r'[<>:"/\\|?*]', '', str(nombre))
        # Reemplazar espacios con guiones bajos
        nombre_limpio = re.sub(r'\s+', '_', nombre_limpio)
        return nombre_limpio
    
    try:
        # Usar ruta relativa para que funcione en cualquier computadora
        base_dir = os.path.dirname(os.path.abspath(__file__))
        user_id_limpio = re.sub(r'[^a-zA-Z0-9_-]', '', str(user_id))
        output_path = os.path.join(base_dir, 'ptmp_site_survey', f'ss_ptmp_{user_id_limpio}.xlsx')
        
        if not os.path.exists(output_path):
            return f"Archivo no encontrado: {output_path}"
        
        @after_this_request
        def eliminar_archivos_temporales(response):
            try:
                # Esperar un poco para asegurar que la descarga se complete
                import time
                time.sleep(3)
                
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                        print(f"✅ Archivo temporal eliminado: {output_path}")
                    except PermissionError:
                        print(f"⚠️ Archivo en uso, programando eliminación posterior: {output_path}")
                        # Programar eliminación posterior
                        import threading
                        def eliminar_luego():
                            time.sleep(10)  # Esperar 10 segundos más
                            try:
                                if os.path.exists(output_path):
                                    os.remove(output_path)
                                    print(f"✅ Archivo temporal eliminado posteriormente: {output_path}")
                            except Exception as e2:
                                print(f"❌ Error en eliminación posterior: {e2}")
                        
                        thread = threading.Thread(target=eliminar_luego)
                        thread.daemon = True
                        thread.start()
                    except Exception as e:
                        print(f"❌ Error eliminando archivo temporal: {e}")
                else:
                    print(f"⚠️ Archivo no encontrado para eliminar: {output_path}")
            except Exception as e:
                print(f"❌ Error en limpieza de archivo temporal: {e}")
            return response
        
        # Limpiar archivos temporales antiguos antes de enviar el archivo
        limpiar_archivos_temporales_ptmp()
        
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return f"Error descargando archivo: {e}"

@app.route('/limpiar_archivos_temp')
def limpiar_archivos_temp():
    """Ruta para limpiar manualmente archivos temporales"""
    try:
        limpiar_archivos_temporales_ptmp()
        return "✅ Limpieza de archivos temporales completada"
    except Exception as e:
        return f"❌ Error en limpieza: {e}"

@app.route('/limpiar_archivos_temp_forzado')
def limpiar_archivos_temp_forzado():
    """Ruta para limpiar forzadamente archivos temporales (incluye cierre de Excel)"""
    try:
        print("🔄 Iniciando limpieza forzada...")
        forzar_cierre_excel()
        time.sleep(3)
        limpiar_archivos_temporales_ptmp()
        return "✅ Limpieza forzada de archivos temporales completada"
    except Exception as e:
        return f"❌ Error en limpieza forzada: {e}"

def forzar_cierre_excel():
    """Fuerza el cierre de procesos de Excel que puedan estar bloqueando archivos"""
    try:
        import subprocess
        import time
        
        # Usar tasklist para buscar procesos de Excel
        try:
            result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq excel.exe'], 
                                  capture_output=True, text=True)
            
            if 'excel.exe' in result.stdout:
                print("🔍 Encontrados procesos de Excel activos")
                # Forzar cierre de todos los procesos de Excel
                subprocess.run(['taskkill', '/f', '/im', 'excel.exe'], capture_output=True)
                print("🔄 Procesos de Excel terminados")
                time.sleep(2)
            else:
                print("✅ No se encontraron procesos de Excel activos")
                
        except Exception as e:
            print(f"⚠️ Error buscando procesos de Excel: {e}")
            
    except Exception as e:
        print(f"⚠️ Error forzando cierre de Excel: {e}")

def limpiar_archivos_temporales_ptmp():
    """Limpia archivos temporales de PtMP que tengan más de 1 hora"""
    import os
    import time
    import glob
    
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        ptmp_dir = os.path.join(base_dir, 'ptmp_site_survey')
        
        if not os.path.exists(ptmp_dir):
            return
            
        # Buscar archivos temporales de PtMP
        patron = os.path.join(ptmp_dir, 'ss_ptmp_*.xlsx')
        archivos_temp = glob.glob(patron)
        
        tiempo_actual = time.time()
        archivos_eliminados = 0
        
        for archivo in archivos_temp:
            try:
                # Obtener el tiempo de modificación del archivo
                tiempo_modificacion = os.path.getmtime(archivo)
                tiempo_diferencia = tiempo_actual - tiempo_modificacion
                
                # Eliminar archivos con más de 1 hora (3600 segundos)
                if tiempo_diferencia > 3600:
                    try:
                        os.remove(archivo)
                        print(f"🗑️ Archivo temporal antiguo eliminado: {os.path.basename(archivo)}")
                        archivos_eliminados += 1
                    except PermissionError:
                        print(f"⚠️ Archivo temporal en uso, no se puede eliminar: {os.path.basename(archivo)}")
                    except Exception as e:
                        print(f"⚠️ Error eliminando archivo temporal {os.path.basename(archivo)}: {e}")
                    
            except Exception as e:
                print(f"⚠️ Error procesando archivo temporal {archivo}: {e}")
                
        if archivos_eliminados > 0:
            print(f"🧹 Limpieza completada: {archivos_eliminados} archivos temporales eliminados")
        else:
            print("✨ No se encontraron archivos temporales para limpiar")
            
    except Exception as e:
        print(f"❌ Error en limpieza de archivos temporales: {e}")

@app.route('/reporte_planeacion')
def reporte_planeacion():
    import pandas as pd
    user_id = request.args.get('user_id', '')
    fila_idx = request.args.get('fila_idx', '')
    datos = None
    if fila_idx:
        try:
            df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
            row = df_db.loc[int(fila_idx)]
            datos = row.to_dict()
        except Exception as e:
            return f"Error leyendo base de datos: {e}"
    if datos:
        html = f"""
        <h2 style='color:#00c3ff;text-align:center;margin-top:40px;'>Reporte de Planeación (Autollenado)</h2>
        <div style='max-width:500px;margin:30px auto;background:#16213e;padding:28px 32px;border-radius:16px;box-shadow:0 4px 24px #00c3ff33;'>
        <b>ID:</b> {datos.get('ID','')}<br>
        <b>Nombre del sitio A:</b> {datos.get('Nombre del sitio A','')}<br>
        <b>Nombre del sitio B:</b> {datos.get('Nombre del sitio B','')}<br>
        <b>Estado:</b> {datos.get('ESTADO','')}<br>
        <b>Tipo de Zona:</b> {datos.get('Tipo de Zona','')}<br>
        </div>
        <div style='text-align:center;margin-top:18px;'><a href='/' style='color:#00c3ff;'>Volver al inicio</a></div>
        """
        return html
    return "<h2 style='color:#00c3ff;text-align:center;margin-top:60px;'>Formulario Reporte de Planeación (en construcción)</h2>"
from flask import render_template, redirect, url_for

@app.route('/formulario_archivos', methods=['GET'])
def formulario_archivos():
    import pandas as pd
    user_id = request.args.get('user_id', '')
    fila_idx = request.args.get('fila_idx', '')
    id_sitio = user_id
    sitio_a = ''
    sitio_b = ''
    if fila_idx:
        try:
            df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
            row = df_db.loc[int(fila_idx)]
            id_sitio = row.get('ID', '') or ''
            sitio_a = row.get('Nombre del sitio A', '') or ''
            sitio_b = row.get('Nombre del sitio B', '') or ''
            if pd.isna(sitio_a): sitio_a = ''
            if pd.isna(sitio_b): sitio_b = ''
        except Exception as e:
            print(f"Error leyendo Google Sheets en formulario_archivos: {e}")
    html_form_mod = html_form.replace(
        '<form id="autoForm" action="/procesar" method="post" enctype="multipart/form-data" autocomplete="off">',
        '<form id="autoForm" action="/procesar" method="post" enctype="multipart/form-data" autocomplete="off">'
        f'\n<input type="hidden" name="user_id" value="{user_id}">'\
        f'\n<input type="hidden" name="fila_idx" value="{fila_idx}">'\
    )
    html_form_mod = html_form_mod.replace(
        '<!-- ANALISIS_AQUI -->',
        f'''
<div class="analisis-info-box">
    <div class="analisis-title"><i class="fa-solid fa-circle-info"></i> Enlace seleccionado</div>
    <div class="analisis-row"><b style="color:#00c3ff;">ID:</b> {id_sitio}</div>
    <div class="analisis-row"><b style="color:#00c3ff;">Sitio A:</b> {sitio_a}</div>
    <div class="analisis-row"><b style="color:#00c3ff;">Sitio B:</b> {sitio_b}</div>
</div>
''')
    return render_template_string(html_form_mod)

@app.route('/seleccion', methods=['POST'])
def seleccion():
    import pandas as pd
    user_id = request.form.get('user_id')
    fila_idx = request.form.get('fila_idx')
    db_path = request.form.get('db_path')

    df_db = pd.read_excel(db_path, engine='openpyxl')
    if not fila_idx:
        return "Falta el índice de fila"
    fila_idx_int = int(fila_idx)
    row = df_db.loc[fila_idx_int]

    id_sitio = row.get('ID', '') or ''
    sitio_a = row.get('Nombre del sitio A', '') or ''
    sitio_b = row.get('Nombre del sitio B', '') or ''
    # Evita mostrar 'nan'
    if pd.isna(sitio_a): sitio_a = ''
    if pd.isna(sitio_b): sitio_b = ''

    return redirect(url_for(
        'formulario_archivos',
        user_id=user_id,
        fila_idx=fila_idx,
        db_path=db_path,
        id_sitio=id_sitio,
        sitio_a=sitio_a,
        sitio_b=sitio_b
    ))
@app.route('/procesar', methods=['POST'])
def procesar():
    import pandas as pd
    import shutil
    import os
    import time
    import xlwings as xw
    from werkzeug.utils import secure_filename

    print("=== INICIO PROCESAR ===")
    print(f"DEBUG: request.files.keys() = {list(request.files.keys())}")
    print(f"DEBUG: request.form.keys() = {list(request.form.keys())}")
    
    # Debug: mostrar todos los archivos recibidos
    print("=== ARCHIVOS RECIBIDOS ===")
    for key, file_list in request.files.lists():
        if isinstance(file_list, list):
            for i, file in enumerate(file_list):
                if file and file.filename:
                    print(f"  {key}[{i}]: {file.filename} ({file.content_type})")
        else:
            if file_list and file_list.filename:
                print(f"  {key}: {file_list.filename} ({file_list.content_type})")
    print("=== FIN ARCHIVOS RECIBIDOS ===")
    
    fila_idx = request.form.get('fila_idx')
    user_id = request.form.get('user_id')
    tipo = request.form.get('tipo', 'site_survey')  # Por defecto site_survey si no se especifica
    print(f"DEBUG: Tipo recibido en procesar: '{tipo}'")

    # Usar siempre la plantilla fija
    template_path = TEMPLATE_PATH
    template_filename = os.path.basename(TEMPLATE_PATH)
    if not os.path.exists(template_path):
        return f"Falta la plantilla fija en la ruta esperada: {template_path}"

    # --- NUEVO: Usar Google Sheets fijo como base de datos ---
    try:
        df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL, keep_default_na=False, na_values=[])
    except Exception as e:
        return f"Error leyendo la base de datos de Google Sheets: {e}"

    # --- 1. Recibe y guarda archivos ---
    imagenes = request.files.getlist('imagenes_electricas')  # Cambiado de 'imagenes' a 'imagenes_electricas'
    pdf_paths = []
    pdfs = request.files.getlist('pdf_file')  # Cambiado de 'pdfs' a 'pdf_file'
    print(f"DEBUG: pdfs recibidos = {len(pdfs)}")
    for idx, pdf in enumerate(pdfs):
        print(f"DEBUG: pdf {idx}: {pdf.filename if pdf else 'None'}")
    
    for idx, pdf in enumerate(pdfs):
        if pdf and pdf.filename:
            filename = secure_filename(f"{idx}_{pdf.filename}")
            pdf_path = os.path.join(UPLOAD_FOLDER, filename)
            pdf.save(pdf_path)
            pdf_paths.append(pdf_path)
            print(f"DEBUG: PDF guardado: {pdf_path}")
        else:
            print(f"DEBUG: PDF {idx} no válido: {pdf}")
    
    print(f"DEBUG: pdf_paths final = {pdf_paths}")
    
    print(f"DEBUG: imagenes recibidas = {len(imagenes)}")
    for idx, img in enumerate(imagenes):
        print(f"DEBUG: imagen {idx}: {img.filename if img else 'None'}")
    
    imagen_paths = []
    for idx, img in enumerate(imagenes):
        if img and img.filename:
            filename = secure_filename(f"{idx}_{img.filename}")
            img_path = os.path.join(UPLOAD_FOLDER, filename)
            img.save(img_path)
            imagen_paths.append(img_path)
            print(f"DEBUG: Imagen guardada: {img_path}")
        else:
            print(f"DEBUG: Imagen {idx} no válida: {img}")
    
    print(f"DEBUG: imagen_paths final = {imagen_paths}")
    
    if not user_id:
        return "Falta el ID"

    # --- SELECCIÓN DEL REGISTRO CORRECTO ---
    if fila_idx is not None and fila_idx != "":
        try:
            datos = df_db.loc[int(fila_idx)]
        except (ValueError, TypeError):
            return "Índice de fila inválido"
    else:
        coincidencias = df_db[df_db['ID'] == user_id]
        if coincidencias.empty:
            return "ID no encontrado en la base de datos."
        datos = coincidencias.iloc[0]

    # Para depuración: imprime las columnas disponibles
    print("Columnas disponibles:", list(datos.index))

    imagen_b_file = request.files.get('imagen_b')
    imagen_b_path = None
    if imagen_b_file and imagen_b_file.filename:
        imagen_b_filename = secure_filename(imagen_b_file.filename or "")
        imagen_b_path = os.path.join(UPLOAD_FOLDER, imagen_b_filename)
        imagen_b_file.save(imagen_b_path)

    archivo_excel_b = request.files.get('archivo_excel_b')
    archivo_excel_b_path = None
    if archivo_excel_b and archivo_excel_b.filename:
        archivo_excel_b_filename = secure_filename(archivo_excel_b.filename or "")
        archivo_excel_b_path = os.path.join(UPLOAD_FOLDER, archivo_excel_b_filename)
        if os.path.exists(archivo_excel_b_path):
            try:
               os.remove(archivo_excel_b_path)
            except Exception as e:
                return f"Error: No se pudo eliminar el archivo de destino. Detalle: {e}"
        archivo_excel_b.save(archivo_excel_b_path)

    
    word_file = request.files.get('word_file')
    word_file_path = None
    if word_file and word_file.filename:
        word_file_filename = secure_filename(word_file.filename or "")
        word_file_path = os.path.join(UPLOAD_FOLDER, word_file_filename)
        word_file.save(word_file_path)
    
    imagenes_electricas = request.files.getlist('imagenes_electricas')
    imagenes_electricas_paths = []
    for idx, img in enumerate(imagenes_electricas):
        if img and img.filename:
            filename = secure_filename(f"electricas_{idx}_{img.filename}")
            img_path = os.path.join(UPLOAD_FOLDER, filename)
            img.save(img_path)
            imagenes_electricas_paths.append(img_path)

    from docx import Document
    import re

    enlace_principal = ""
    nombreEnlace = ""
    if word_file_path and os.path.exists(word_file_path):
        doc = Document(word_file_path)
        word_text = "\n".join([p.text for p in doc.paragraphs])
        print("WORD TEXT:", word_text)
        for line in word_text.splitlines():
            print("LINE:", line)
            m = re.search(r'Transmission details\s*\(([^)]*)\)', line, re.IGNORECASE)
            if m:
                nombreEnlace = m.group(1)
                nombreEnlace = re.sub(r'\s*\(cambio\)\s*', '', nombreEnlace, flags=re.IGNORECASE)
                nombreEnlace = re.sub(r'\.pl6\s*$', '', nombreEnlace, flags=re.IGNORECASE)
                nombreEnlace = nombreEnlace.strip()
                # SEPARA LOS NOMBRES SI HAY UN GUION Y NÚMEROS
                partes = re.split(r'\d+[A-Z]-', nombreEnlace)
                if len(partes) >= 2:
                    nombre1 = partes[0].strip()
                    nombre2 = partes[1].split()[0].strip()
                enlace_principal = f"{nombre1} - {nombre2}"
            else:
                enlace_principal = nombreEnlace
            break  # Solo el primero
    print("ENLACE PRINCIPAL:", enlace_principal)
      
    img_consumo = request.files.get('img_consumo')
    img_configuracion = request.files.get('img_configuracion')
    img_linea_vista = request.files.get('img_linea_vista')

    img_consumo_path = None
    img_configuracion_path = None
    img_linea_vista_path = None

    if img_consumo and img_consumo.filename:
        img_consumo_path = os.path.join(UPLOAD_FOLDER, secure_filename(img_consumo.filename or ""))
        img_consumo.save(img_consumo_path)
    if img_configuracion and img_configuracion.filename:
        img_configuracion_path = os.path.join(UPLOAD_FOLDER, secure_filename(img_configuracion.filename or ""))
        img_configuracion.save(img_configuracion_path)
    if img_linea_vista and img_linea_vista.filename:
        img_linea_vista_path = os.path.join(UPLOAD_FOLDER, secure_filename(img_linea_vista.filename or ""))
        img_linea_vista.save(img_linea_vista_path)

      # Recibe imagen para hoja 3 (Formato KMZ)
    imagen_kmz_file = request.files.get('imagen_kmz')
    print("imagen_kmz_file:", imagen_kmz_file)
    print("imagen_kmz_file.filename:", imagen_kmz_file.filename if imagen_kmz_file else None)
    imagen_kmz_path = None
    if imagen_kmz_file and imagen_kmz_file.filename:
        imagen_kmz_filename = secure_filename(imagen_kmz_file.filename or "")
        imagen_kmz_path = os.path.join(UPLOAD_FOLDER, imagen_kmz_filename)
        imagen_kmz_file.save(imagen_kmz_path)

    
    
    output_path = os.path.join(UPLOAD_FOLDER, f'LLENADO_{template_filename or "template.xlsx"}')
    kmz_file = request.files.get('kmz')
    kml_image_file = request.files.get('kml_image')
    print("kml_image_file:", kml_image_file)
    print("kml_image_file.filename:", kml_image_file.filename if kml_image_file else None)
    kml_image_path = None
    if kml_image_file and kml_image_file.filename:
        kml_image_filename = secure_filename(kml_image_file.filename or "")
        kml_image_path = os.path.join(UPLOAD_FOLDER, kml_image_filename)
        kml_image_file.save(kml_image_path)

    kmz_path = None
    if kmz_file and kmz_file.filename:
        kmz_filename = secure_filename(kmz_file.filename or "")
        kmz_path = os.path.join(UPLOAD_FOLDER, kmz_filename)
        kmz_file.save(kmz_path)

    kml_image_path = None
    if kml_image_file and kml_image_file.filename:
       kml_image_filename = secure_filename(kml_image_file.filename or "")
       kml_image_path = os.path.join(UPLOAD_FOLDER, kml_image_filename)
       kml_image_file.save(kml_image_path)
    
    # --- PROCESAR ARCHIVOS ADICIONALES FALTANTES ---
    
    # Archivos de planos A
    planos_a_img1 = request.files.get('planos_a_img1')
    planos_a_img2 = request.files.get('planos_a_img2')
    planos_a_img3 = request.files.get('planos_a_img3')
    
    planos_a_img1_path = None
    planos_a_img2_path = None
    planos_a_img3_path = None
    
    if planos_a_img1 and planos_a_img1.filename:
        planos_a_img1_filename = secure_filename(planos_a_img1.filename or "")
        planos_a_img1_path = os.path.join(UPLOAD_FOLDER, planos_a_img1_filename)
        planos_a_img1.save(planos_a_img1_path)
        print(f"DEBUG: Plano A img1 guardado: {planos_a_img1_path}")
    
    if planos_a_img2 and planos_a_img2.filename:
        planos_a_img2_filename = secure_filename(planos_a_img2.filename or "")
        planos_a_img2_path = os.path.join(UPLOAD_FOLDER, planos_a_img2_filename)
        planos_a_img2.save(planos_a_img2_path)
        print(f"DEBUG: Plano A img2 guardado: {planos_a_img2_path}")
    
    if planos_a_img3 and planos_a_img3.filename:
        planos_a_img3_filename = secure_filename(planos_a_img3.filename or "")
        planos_a_img3_path = os.path.join(UPLOAD_FOLDER, planos_a_img3_filename)
        planos_a_img3.save(planos_a_img3_path)
        print(f"DEBUG: Plano A img3 guardado: {planos_a_img3_path}")
    
    # Archivos de planos B
    planos_b_img1 = request.files.get('planos_b_img1')
    planos_b_img2 = request.files.get('planos_b_img2')
    planos_b_img3 = request.files.get('planos_b_img3')
    
    planos_b_img1_path = None
    planos_b_img2_path = None
    planos_b_img3_path = None
    
    if planos_b_img1 and planos_b_img1.filename:
        planos_b_img1_filename = secure_filename(planos_b_img1.filename or "")
        planos_b_img1_path = os.path.join(UPLOAD_FOLDER, planos_b_img1_filename)
        planos_b_img1.save(planos_b_img1_path)
        print(f"DEBUG: Plano B img1 guardado: {planos_b_img1_path}")
    
    if planos_b_img2 and planos_b_img2.filename:
        planos_b_img2_filename = secure_filename(planos_b_img2.filename or "")
        planos_b_img2_path = os.path.join(UPLOAD_FOLDER, planos_b_img2_filename)
        planos_b_img2.save(planos_b_img2_path)
        print(f"DEBUG: Plano B img2 guardado: {planos_b_img2_path}")
    
    if planos_b_img3 and planos_b_img3.filename:
        planos_b_img3_filename = secure_filename(planos_b_img3.filename or "")
        planos_b_img3_path = os.path.join(UPLOAD_FOLDER, planos_b_img3_filename)
        planos_b_img3.save(planos_b_img3_path)
        print(f"DEBUG: Plano B img3 guardado: {planos_b_img3_path}")
    
    # Archivos de fotos individuales para hoja 9 (Fotos A)
    fotos9_names = [
        "foto_e11", "foto_v11", "foto_e23", "foto_e48", "foto_v48", "foto_e59", "foto_v59",
        "foto_e70", "foto_v70", "foto_e87", "foto_v87", "foto_e98", "foto_v98", "foto_e127",
        "foto_v127", "foto_e138", "foto_v138", "foto_e150", "foto_v150"
    ]
    
    fotos9_paths = {}
    for name in fotos9_names:
        foto_file = request.files.get(name)
        if foto_file and foto_file.filename:
            filename = secure_filename(f"{name}_{foto_file.filename}")
            foto_path = os.path.join(UPLOAD_FOLDER, filename)
            foto_file.save(foto_path)
            fotos9_paths[name] = foto_path
            print(f"DEBUG: Foto {name} guardada: {foto_path}")
        else:
            fotos9_paths[name] = None
    
    # Archivos de fotos individuales para hoja 10 (Fotos B)
    fotos10_names = [
        "foto_b_1", "foto_b_2", "foto_b_3", "foto_b_4", "foto_b_5", "foto_b_6", "foto_b_7", "foto_b_8", "foto_b_9", "foto_b_10"
    ]
    
    fotos10_paths = {}
    for name in fotos10_names:
        foto_file = request.files.get(name)
        if foto_file and foto_file.filename:
            filename = secure_filename(f"{name}_{foto_file.filename}")
            foto_path = os.path.join(UPLOAD_FOLDER, filename)
            foto_file.save(foto_path)
            fotos10_paths[name] = foto_path
            print(f"DEBUG: Foto {name} guardada: {foto_path}")
        else:
            fotos10_paths[name] = None 
    
    # Validar que la plantilla sea un archivo Excel válido antes de abrirla
    try:
        with open(template_path, 'rb') as f:
            signature = f.read(4)
            if signature != b'PK\x03\x04':
                return "Error: La plantilla fija no es un archivo Excel válido (.xlsx). Reemplaza la plantilla por una correcta y sin daños."
    except Exception as e:
        return f"Error al validar la plantilla fija: {e}"


    # --- 2. Llenado con xlwings ---
    print(f"DEBUG: Intentando abrir plantilla: {template_path}")
    app_excel = xw.App(visible=False)
    if os.path.exists(output_path):
       try:
          os.remove(output_path)
       except Exception as e:
        print(f"Error: No se pudo eliminar el archivo de salida. Detalle: {e}")
        return f"Error: No se pudo eliminar el archivo de salida. Detalle: {e}"
    try:
        wb = app_excel.books.open(template_path)
    except Exception as e:
        print(f"Error: No se pudo abrir la plantilla de Excel. Detalle: {e}")
        app_excel.quit()
        return "Error: No se pudo abrir la plantilla de Excel. Verifica que el archivo no esté dañado ni abierto en otro programa."
    if wb is None:
        print("Error: wb es None después de abrir la plantilla.")
        app_excel.quit()
        return "Error: No se pudo abrir la plantilla de Excel. Verifica que el archivo no esté dañado ni abierto en otro programa."
    required_sheets = [
        '4. Estudio de informacion A',
        '1. Analisis de Red y Frecuencia',
        '2. Electricas - Diseño log- Fis',
        '5. Estudio de informacion B',
        '8. Estudio de factibilidad',
        '3. Formato KMZ',
        '0. Carátula'
    ]
    try:
        sheet_names = [s.name for s in wb.sheets]
    except Exception as e:
        wb.close()
        app_excel.quit()
        return f"Error: No se pudieron enumerar las hojas del archivo Excel. El archivo puede estar dañado o vacío. Detalle: {e}"
    for sheet in required_sheets:
        if sheet not in sheet_names:
            wb.close()
            app_excel.quit()
            return f"Error: La hoja '{sheet}' no existe en la plantilla de Excel."
    ws_a = wb.sheets['4. Estudio de informacion A']
    ws_red = wb.sheets['1. Analisis de Red y Frecuencia']
    ws_electricas = wb.sheets['2. Electricas - Diseño log- Fis']
    ws_electricas.range('B9').value = enlace_principal
    campos_electricas_celdas = {
        'Configuración MW:': ['D9', 'B14', 'C28', 'F28'],
        'Tamaño de la antena (m)': ['C27', 'F27'],
        'Potencia de Transmisión (dBm)': ['C29', 'F29'],
        'Frecuencia (MHz)': ['C32', 'F32'],
        'Nombre del sitio A': 'C33',
        'Nombre del sitio B': 'F33',
        'ID del sitio A': 'C44',
        'ID del sitio B': 'F44',
        'Potencia de Recepción (dBm)': ['C30', 'F30'],
        'Banda': ['C31', 'F31'],
        'Frecuencia (MHz)': ['C32', 'F32'],
        'NOMBRE DEL SITIO':'C33',
        'Nombre del sitio 2': 'F33',
        'ID':'C34',
        'ID 2': 'F34',
        'consumo de potencia':'F9',
    
   }

    for campo, celdas in campos_electricas_celdas.items():
        valor = normaliza_na(datos.get(campo, ""))
        if isinstance(celdas, list):
            for celda in celdas:
                if isinstance(celda, str):
                    ws_electricas.range(celda).value = valor
                else:
                    print(f"Celda inválida para campo {campo}: {celda}")
        elif isinstance(celdas, str):
            ws_electricas.range(celdas).value = valor
        else:
            print(f"Referencia de celda inválida para campo {campo}: {celdas}")
    
    lat_a = datos.get('LATITUD (TORRE)', '')
    lon_a = datos.get('LONGITUD (TORRE)', '')
    coord_a = f"{lat_a}, {lon_a}" if lat_a and lon_a else ""

    lat_b = datos.get('LATITUD (TORRE) 2', '')
    lon_b = datos.get('LONGITUD (TORRE) 2', '')
    coord_b = f"{lat_b}, {lon_b}" if lat_b and lon_b else ""

    # Ajusta las celdas según tu plantilla
    ws_electricas.range('C35').value = coord_a  # Coordenadas sitio A
    ws_electricas.range('F35').value = coord_b  # Coordenadas sitio B

    # Eliminar todas las imágenes existentes en la hoja 2 antes de insertar nuevas
    for pic in ws_electricas.pictures:
        try:
            pic.delete()
        except Exception as e:
            print(f"Error eliminando imagen previa: {e}")

    # Imagen de consumo
    if img_consumo_path and os.path.exists(img_consumo_path):
        cell_range = ws_electricas.range('C14:D20')
        ws_electricas.pictures.add(
            os.path.abspath(img_consumo_path),
            left=cell_range.left,
            top=cell_range.top,
            width=cell_range.width,
            height=cell_range.height
       )
    # Imagen de configuración
    if img_configuracion_path and os.path.exists(img_configuracion_path):
        cell_range = ws_electricas.range('E14:G20')
        ws_electricas.pictures.add(
            os.path.abspath(img_configuracion_path),
            left=cell_range.left,
            top=cell_range.top,
            width=cell_range.width,
            height=cell_range.height
       )
    # Imagen de línea de vista
    if img_linea_vista_path and os.path.exists(img_linea_vista_path):
        cell_range = ws_electricas.range('B39:G55')
        ws_electricas.pictures.add(
            os.path.abspath(img_linea_vista_path),
            left=cell_range.left,
            top=cell_range.top,
            width=cell_range.width,
            height=cell_range.height
       )
    

    import pandas as pd
    import dataframe_image as dfi
    # --- Tabla horizontal de frecuencia para hoja 1 ---
    campos_tabla = {
        "Tamaño de la antena (m)": datos.get("Tamaño de la antena (m)", ""),
        "Potencia de Transmisión (dBm)": datos.get("Potencia de Transmisión (dBm)", ""),
        "Potencia de Recepción (dBm)": datos.get("Potencia de Recepción (dBm)", ""),
        "Banda": datos.get("Banda", ""),
        "Frecuencia (MHz)": datos.get("Frecuencia (MHz)", ""),
        "#1 Canal ID S1": datos.get("#1 Canal ID S1", ""),
        "#1 Frecuencia de Diseño S1": datos.get("#1 Frecuencia de Diseño S1", ""),
        "#2 Frecuencia de Diseño S1": datos.get("#2 Frecuencia de Diseño S1", ""),
        "#1 Canal ID S2": datos.get("#1 Canal ID S2", ""),
        "#1 Frecuencia de Diseño S2": datos.get("#1 Frecuencia de Diseño S2", ""),
        "#2 Frecuencia de Diseño S2": datos.get("#2 Frecuencia de Diseño S2", "")
    }
    df_tabla = pd.DataFrame([campos_tabla])

    def wrap_col_label(label, width=14):
        return '\n'.join(textwrap.wrap(label, width=width, break_long_words=False))

    col_labels_wrapped = [wrap_col_label(str(col), width=14) for col in df_tabla.columns]

    styler = (
        df_tabla.style
        .set_table_styles([
            {'selector': 'th', 'props': [('background-color', '#0074D9'), ('color', 'white'), ('font-size', '12pt')]},
            {'selector': 'td', 'props': [('background-color', '#D9EAF7'), ('color', 'black'), ('font-size', '12pt'), ('text-align', 'center'), ('border', '1.5px solid #2F5597')]}
        ])
        .set_table_attributes('style="border-collapse:collapse; margin:auto;"')
        .hide(axis="index")  # Oculta los números de fila
    )

    img_path = os.path.join(UPLOAD_FOLDER, 'tabla_frecuencia.png')
    try:
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots(figsize=(12, 2.2))
        ax.axis('off')
        tabla = ax.table(
            cellText=df_tabla.values.tolist(),
            colLabels=col_labels_wrapped,
            loc='center',
            cellLoc='center'
        )
        # Estilo: encabezado azul, fila datos azul clarito
        for (row, col), cell in tabla.get_celld().items():
            if row == 0:
                cell.set_facecolor('#0074D9')
                cell.set_text_props(color='white', weight='bold')
            else:
                cell.set_facecolor('#D9EAF7')
                cell.set_text_props(color='black')
        tabla.auto_set_font_size(False)
        tabla.set_fontsize(10)
        tabla.scale(1.5, 2.8)
        plt.savefig(img_path, bbox_inches='tight', pad_inches=0, dpi=200)
        plt.close()
    except Exception as e:
        print(f"Error al exportar tabla como imagen: {e}")
        # Crear una imagen simple si falla
        import matplotlib.pyplot as plt
        plt.figure(figsize=(10, 2))
        plt.text(0.5, 0.5, 'Tabla de Frecuencias', ha='center', va='center', fontsize=16)
        plt.axis('off')
        plt.savefig(img_path, dpi=200, bbox_inches='tight', pad_inches=0)
        plt.close()
    df_tabla.to_csv(os.path.join(UPLOAD_FOLDER, "debug_tabla.csv"), index=False)
    # Guarda el Excel de depuración con estilo
    try:
        styler.to_excel(os.path.join(UPLOAD_FOLDER, "debug_tabla.xlsx"))
    except Exception as e:
        print(f"Error al guardar el Excel estilizado: {e}")
    print("Tabla horizontal guardada en debug_tabla.csv y debug_tabla.xlsx")

    # Insertar la imagen en el rango A23:G28 de la hoja 1
    cell_range = ws_red.range('A23:G28')
    ws_red.pictures.add(
        img_path,
        left=cell_range.left,
        top=cell_range.top,
        width=cell_range.width,
        height=cell_range.height
    )

    cell_range = ws_red.range('D19').value = datos.get('Margen de desvanecimiento ', '')
    ws_red.range('E19').value = datos.get('Disponibilidad anual (%) ', '')
    ws_a.range('AG8').value = user_id
    ws_a.range('B19').value = enlace_principal
    ws_b = wb.sheets['5. Estudio de informacion B']
    ws_b.range('AG8').value = datos.get('ID 2', '')
    ws_factibilidad = wb.sheets['8. Estudio de factibilidad']
    ws_kmz = wb.sheets['3. Formato KMZ']
    #ws_kmz.activate()


    try:
      ws_caratula = wb.sheets['0. Carátula']
      nombre_a = datos.get('Nombre del sitio A', '')
      nombre_b = datos.get('Nombre del sitio B', '')
      ws_caratula.range('A43').value = f"{nombre_a} - {nombre_b}"
      nombre_enlace_caratula = ws_caratula.range('A43').value or ""
      import re
      nombre_enlace_sin_numeros = re.sub(r'\b\d+[A-Z]?\s*', '', nombre_enlace_caratula).strip()
      nombre_enlace_sin_numeros = re.sub(r'\s{2,}', ' ', nombre_enlace_sin_numeros)
      ws_red.range('B19').value = nombre_enlace_sin_numeros
    except Exception as e:
      print(f"Advertencia: No se pudo llenar la hoja 0. Carátula: {e}")

    # Usar las imágenes de fotos 9 ya procesadas anteriormente
    fotos9_names = [
        "foto_e11", "foto_v11", "foto_e23", "foto_e48", "foto_v48", "foto_e59", "foto_v59",
        "foto_e70", "foto_v70", "foto_e87", "foto_v87", "foto_e98", "foto_v98", "foto_e127",
        "foto_v127", "foto_e138", "foto_v138", "foto_e150", "foto_v150"
    ]
    imagenes_fotos9_paths = [fotos9_paths.get(name) for name in fotos9_names]
    
    
    estado_a_region = {
        'Aguascalientes': 'CENTRO',
        'Baja California': 'NORTE',
        'Baja California Sur': 'NORTE',
        'Campeche': 'SURESTE',
        'Chiapas': 'SUR',
        'Chihuahua': 'NORTE',
        'Ciudad de México': 'CENTRO',
        'Coahuila': 'NORTE',
        'Colima': 'OCCIDENTE',
        'Durango': 'NORTE',
        'Estado de México': 'CENTRO',
        'Guanajuato': 'CENTRO',
        'Guerrero': 'SUR',
        'Hidalgo': 'CENTRO',
        'Jalisco': 'OCCIDENTE',
        'Michoacán': 'OCCIDENTE',
        'Morelos': 'CENTRO',
        'Nayarit': 'OCCIDENTE',
        'Nuevo León': 'NORESTE',
        'Oaxaca': 'SUR',
        'Puebla': 'CENTRO',
        'Querétaro': 'CENTRO',
        'Quintana Roo': 'SURESTE',
        'San Luis Potosí': 'CENTRO',
        'Sinaloa': 'NORTE',
        'Sonora': 'NORTE',
        'Tabasco': 'SURESTE',
        'Tamaulipas': 'NORESTE',
        'Tlaxcala': 'CENTRO',
        'Veracruz': 'GOLFO',
        'Yucatán': 'SURESTE',
        'Zacatecas': 'NORTE'
    }
    

    tipo_zona = str(datos.get('Tipo de Zona', '')).strip().lower()
    ws_a.range('M20').value = tipo_zona == 'urbana'
    ws_a.range('Q20').value = tipo_zona == 'sub-urbana'
    ws_a.range('V20').value = tipo_zona == 'rural'
    ws_a.range('Y20').value = tipo_zona == 'ejidal'
    ws_a.range('AB20').value = tipo_zona == 'pueblo mágico'

    tipo_zona2 = str(datos.get('Tipo de Zona 2', '')).strip().lower()
    ws_b.range('M20').value = tipo_zona2 == 'urbana'
    ws_b.range('Q20').value = tipo_zona2 == 'sub-urbana'
    ws_b.range('V20').value = tipo_zona2 == 'rural'
    ws_b.range('Y20').value = tipo_zona2 == 'ejidal'
    ws_b.range('AB20').value = tipo_zona2 == 'pueblo mágico'
  
    visible = str(datos.get('El sitio es visible de día y de noche (libre de maleza y arboles): ', '')).strip().lower()
    ws_a.range('Q21').value = visible == 'si'
    ws_a.range('T21').value = visible == 'no'

    visible2 = str(datos.get('El sitio es visible de día y de noche (libre de maleza y arboles): 2', '')).strip().lower()
    ws_b.range('Q21').value = visible2 == 'si'
    ws_b.range('T21').value = visible2 == 'no'

    tipo_camino = str(datos.get('Tipo de Camino', '')).strip().lower()
    ws_a.range('H22').value = tipo_camino == 'terracería'
    ws_a.range('M22').value = tipo_camino == 'pavimentado'
    ws_a.range('R22').value = tipo_camino == 'empedrado'
    ws_a.range('W22').value = tipo_camino == 'mixto'

    tipo_camino2  = str(datos.get(' Tipo de Camino 2 ', '')).strip().lower()
    ws_b.range('H22').value = tipo_camino2 == 'terracería'
    ws_b.range('M22').value = tipo_camino2 == 'pavimentado'
    ws_b.range('R22').value = tipo_camino2 == 'empedrado'
    ws_b.range('W22').value = tipo_camino2 == 'mixto'

    tipo_torre = str(datos.get('Tipo de Torre', '')).strip().lower()
    ws_a.range('H34').value = tipo_torre == 'autosoportada'
    ws_a.range('P34').value = tipo_torre == 'arriostrada'
    ws_a.range('W34').value = tipo_torre == 'Monopolo'
    ws_a.range('AC34').value = tipo_torre == 'Minipolo'
    ws_a.range('AH34').value = tipo_torre == 'otro'
    
   
    espacio_disponible = str(datos.get('¿Espacio disponible de conexión?', '')).strip().lower()
    ws_a.range('AG51').value = espacio_disponible == 'si'
    ws_a.range('AJ51').value = espacio_disponible == 'no'
    ws_a.range('V40').value = espacio_disponible == 'si'
    ws_a.range('Z40').value = espacio_disponible == 'no'

    cara_propuesta = str(datos.get('Cara de preparación para cableado vertical en torre', '')).strip().lower()
    ws_a.range('S42').value = cara_propuesta == 'a'
    ws_a.range('X42').value = cara_propuesta == 'b'
    ws_a.range('AC42').value = cara_propuesta == 'c'
    ws_a.range('AH42').value = cara_propuesta == 'd'

    barra_tierra = str(datos.get('Barra de Tierra', '')).strip().lower()
    ws_a.range('P53').value = barra_tierra == 'si'
    ws_a.range('S53').value = barra_tierra == 'no'

    tipo_solucion = str(datos.get('Tipo de Solucion', '')).strip().lower()
    ws_a.range('P55').value = tipo_solucion == 'piso'
    ws_a.range('S55').value = tipo_solucion == 'torre'

    existe_break = str(datos.get('¿Existe algun breaker existente en sitio?', '')).strip().lower()
    ws_a.range('Y47').value = existe_break == 'si'
    ws_a.range('AB47').value = existe_break == 'no'
    
    alimentacion_compatible = str(datos.get('Alimentacion compatible con el equipamiento ', '')).strip().lower()
    ws_a.range('Y51').value = alimentacion_compatible == 'si'
    ws_a.range('AB51').value = alimentacion_compatible == 'no'

    sistema_electrico = str(datos.get('SISTEMA ELECTRICO', '')).strip().lower()
    ws_a.range('AG47').value = sistema_electrico == 'monofásica'
    ws_a.range('AJ47').value = sistema_electrico == 'bifásica'

    tipo_torre2 = str(datos.get('Tipo de Torre2', '')).strip().lower()
    ws_a.range('H58').value = tipo_torre2 == 'autosoportada'
    ws_a.range('P58').value = tipo_torre2 == 'arriostrada'
    ws_a.range('W58').value = tipo_torre2 == 'monopolo'
    ws_a.range('AC58').value = tipo_torre2 == 'minipolo'
    ws_a.range('AH58').value = tipo_torre2 == 'otro'

    espacio_disponible2 = str(datos.get('¿Espacio disponible de conexión?2', '')).strip().lower()
    ws_a.range('V64').value = espacio_disponible2 == 'si'
    ws_a.range('Z64').value = espacio_disponible2 == 'no'

    cara_preparacion2 = str(datos.get('Cara de preparación para cableado vertical en torre 2', '')).strip().lower()
    ws_a.range('S66').value = cara_preparacion2 == 'a'
    ws_a.range('X66').value = cara_preparacion2 == 'b'
    ws_a.range('AC66').value = cara_preparacion2 == 'c'
    ws_a.range('AH66').value = cara_preparacion2 == 'd'
    
    existe_tierra2 = str(datos.get('Existe Barra de Tierras 2', '')).strip().lower()
    ws_a.range('P77').value = existe_tierra2 == 'si'
    ws_a.range('S77').value = existe_tierra2 == 'no'

    tipo_solucion2 = str(datos.get('Tipo de solucion 2', '')).strip().lower()
    ws_a.range('P79').value = tipo_solucion2 == 'piso'
    ws_a.range('S79').value = tipo_solucion2 == 'torre'
    
    existe_break2 = str(datos.get('Existe algun breaker existente en sitio 2 ', '')).strip().lower()
    ws_a.range('Y71').value = existe_break2 == 'si'
    ws_a.range('AB71').value = existe_break2 == 'no'

    alimenacion_existente2= str(datos.get('SISTEMA ELECTRICO 2', '')).strip().lower()
    ws_a.range('AG71').value = alimenacion_existente2 == 'monofásica'
    ws_a.range('AJ71').value = alimenacion_existente2 == 'bifásica'
    
    alimenacion_compatible2= str(datos.get('Alimentacion compatible con el equipamiento 2', '')).strip().lower()
    ws_a.range('Y75').value = alimenacion_compatible2 == 'si'
    ws_a.range('AB75').value = alimenacion_compatible2 == 'no'

    espacio_conexion2= str(datos.get('¿Espacio disponible de conexión? 2', '')).strip().lower()
    ws_a.range('AG75').value = espacio_conexion2 == 'si'
    ws_a.range('AJ75').value = espacio_conexion2 == 'no'

    linea_vista = str(datos.get('Linea de vista ', '')).strip().lower()
    motivo = str(datos.get('Motivo ', '')).strip().lower()

    ws_a.range('K82').value = (linea_vista == 'si')
    ws_a.range('O82').value = (linea_vista == 'no')
    ws_a.range('J83').value = False
    ws_a.range('O83').value = False
    ws_a.range('U83').value = False
    ws_a.range('AA83').value = False
    ws_a.range('E84').value = False

    if linea_vista == 'no':
        if motivo == 'arboles':
            ws_a.range('J83').value = True
        elif motivo == 'espectacular':
            ws_a.range('O83').value = True
        elif motivo == 'edificio':
            ws_a.range('U83').value = True
        elif motivo == 'montaña':
            ws_a.range('AA83').value = True
        elif motivo == 'n/a':
            ws_a.range('E84').value = True

    campos_a_celdas = {
        'NOMBRE DEL SITIO': ['K8', 'H33'],
        #'REGION': 'E9',
        'PROPIETARIO': 'N9',
        'ESTADO ': 'AD14',
        'Calle': 'E13',
        'Colonia': 'E14',
        'Municipio': 'F15',
        'C.P': 'AD13',
        'Referencias':'K16',
        'Nombre de contacto en sitio': 'I18',
        'Telefono': 'AC18',
        'Tipo de Zona': 'E16',
        'Tipo de Camino': 'E17',
        'LATITUD (TORRE)': 'L29',
        'LONGITUD (TORRE)': 'AB29',
        'LATITUD (FACHADA)': 'L26',
        'LONGITUD (FACHADA)': 'AB26',
        'Altitud (msnm)': 'N30',
        'Diametro de pierna superior':'L35',
        'Diametro de pierna Inferior':'V35',
        'NCRA RB':'AC35',
        'Franja2RB':'AI35',
        'Altura de la Torre':'L36',
        'Dado':'V36',
        'Altura Edificio1':'AF36',
        'Nivel inferior de franja disponible': 'U37',
        'Nivel superior de franja disponible': 'AI37',
        'Altura de MW conforme a topologia': 'C40',
        'Azimut RB ': 'N40',
        'Propuesta de altura de antena de MW1': 'AC40',
        'Propuesta de altura de antena de MW (SD)1': 'AH40',
        'Altura de soporte para OMB propuesto': 'P45',
        'Longitud del cable de tierra nuevo OMB': 'P46',
        'Longitud del cable de tierra ODU': 'P47',
        'Longitud de cable IF': 'P48',
        'Tipo de soporte para antena MW propuesto': 'P49',
        'Longitud de cable ACDB-Nuevo OMB': 'P50',
        'Longitud de cable RTN - Router':'P51',
        'Longitud de cable RTN - BBU SITE 1': 'P52',
        'MEDICION DE BARRA DE TIERRA (Ohms)':'P54',
        'Nombre del sitio 2': 'H57',
        'Diámetro de Pierna superio2':'L59',
        'Diámetro de Pierna inferior2':'V59',
        ' NCRA2 ':'AC59',
        'Franja2-2':'AI59',
        'Altura torre 2': 'L60',
        'DADO 2':'V60',
        'Altura edificio 2':'AF60',
        'Nivel inferior de franja disponible 2': 'U61',
        'Nivel superior de franja disponible 2': 'AI61',
        'Altura de MW conforme a topologia 2': 'C64',
        'Azimut 2': 'N64',
        'Propuesta de altura de antena de MW2': 'AC64',
        'Propuesta de altura de antena de MW (SD)2':'AH64',
        'Altura de soporte para OMB propuesto2':'P69',
        'Longitud del cable de tierra nuevo OMB 2': 'P70',
        'Longitud del cable de tierra ODU 2': 'P71',
        'Longitud de cable IF 2': 'P72',
        'Tipo de soporte para antena MW propuesto 2': 'P73',
        'Longitud de cable ACDB-Nuevo OMB 2': 'P74',
        'Longitud de cable RTN - Router 2': 'P75',
        'Longitud de cable RTN - BBU 2': 'P76',
        'Medición del Sistema de Tierras 2': 'P78',
        'Nombre del sitio A': ['M117', 'M139'],
        'Nombre del sitio B': ['M162', 'M184'],
     
    }

    print("ESTADO:", datos.get('ESTADO'))
    print("ESTADO 2:", datos.get('ESTADO 2'))
    print("ESTADO2:", datos.get('ESTADO2'))

    estado_b_region = {
        'Aguascalientes': 'CENTRO',
        'Baja California': 'NORTE',
        'Baja California Sur': 'NORTE',
        'Campeche': 'SURESTE',
        'Chiapas': 'SUR',
        'Chihuahua': 'NORTE',
        'Ciudad de México': 'CENTRO',
        'Coahuila': 'NORTE',
        'Colima': 'OCCIDENTE',
        'Durango': 'NORTE',
        'Estado de México': 'CENTRO',
        'Guanajuato': 'CENTRO',
        'Guerrero': 'SUR',
        'Hidalgo': 'CENTRO',
        'Jalisco': 'OCCIDENTE',
        'Michoacán': 'OCCIDENTE',
        'Morelos': 'CENTRO',
        'Nayarit': 'OCCIDENTE',
        'Nuevo León': 'NORESTE',
        'Oaxaca': 'SUR',
        'Puebla': 'CENTRO',
        'Querétaro': 'CENTRO',
        'Quintana Roo': 'SURESTE',
        'San Luis Potosí': 'CENTRO',
        'Sinaloa': 'NORTE',
        'Sonora': 'NORTE',
        'Tabasco': 'SURESTE',
        'Tamaulipas': 'NORESTE',
        'Tlaxcala': 'CENTRO',
        'Veracruz': 'GOLFO',
        'Yucatán': 'SURESTE',
        'Zacatecas': 'NORTE'
    }
    estado_b = datos.get('ESTADO 2')
    if not estado_b or pd.isna(estado_b):
       estado_b = datos.get('ESTADO2')
    if not estado_b or pd.isna(estado_b):
       estado_b = datos.get('ESTADO')
    region_b = estado_b_region.get(str(estado_b).strip(), 'OTRA')
    ws_b.range('E9').value = region_b

    estado_a_region = {
         'Aguascalientes': 'CENTRO',
         'Baja California': 'NORTE',
         'Baja California Sur': 'NORTE',
         'Campeche': 'SURESTE',
         'Chiapas': 'SUR',
         'Chihuahua': 'NORTE',
         'Ciudad de México': 'CENTRO',
         'Coahuila': 'NORTE',
         'Colima': 'OCCIDENTE',
         'Durango': 'NORTE',
         'Estado de México': 'CENTRO',
         'Guanajuato': 'CENTRO',
         'Guerrero': 'SUR',
         'Hidalgo': 'CENTRO',
         'Jalisco': 'OCCIDENTE',
         'Michoacán': 'OCCIDENTE',
         'Morelos': 'CENTRO',
         'Nayarit': 'OCCIDENTE',
         'Nuevo León': 'NORESTE',
         'Oaxaca': 'SUR',
         'Puebla': 'CENTRO',
         'Querétaro': 'CENTRO',
         'Quintana Roo': 'SURESTE',
         'San Luis Potosí': 'CENTRO',
         'Sinaloa': 'NORTE',
         'Sonora': 'NORTE',
         'Tabasco': 'SURESTE',
         'Tamaulipas': 'NORESTE',
         'Tlaxcala': 'CENTRO',
         'Veracruz': 'GOLFO',
         'Yucatán': 'SURESTE',
         'Zacatecas': 'NORTE'
    }

    # Para el sitio A
    estado_a = datos.get('ESTADO ', '').strip()
    region_a = estado_a_region.get(estado_a, 'OTRA')
    ws_a.range('D10').value = region_a  # Ajusta la celda si tu plantilla usa otra

# Para el sitio B (si aplica)
    estado_b = datos.get('ESTADO 2 ', '').strip()
    region_b = estado_a_region.get(estado_b, 'OTRA')
    ws_b.range('D10').value = region_b  # Ajusta la celda si tu plantilla usa otra
    

    campos_b_celdas= {
    'Nombre del sitio 2': 'K8',
    'ID 2': 'AG8',
    #'REGION 2': 'E9',
    'PROPIETARIO 2': 'N9',
    'ESTADO 2': 'AD14',
    'Calle 2': 'E13',
    'Colonia 2': 'E14',
    'Municipio 2': 'F15',
    'C.P 2': 'AD13',
    'Referencias 2':'K16',
    'Nombre de contacto en sitio 2': 'I18',
    'Telefono 2': 'AC18',
    'LATITUD (TORRE) 2': 'L29',
    'LONGITUD (TORRE) 2': 'AB29',
    'LATITUD (FACHADA) 2': 'L26',
    'LONGITUD (FACHADA) 2': 'AB26',
    'Altitud (msnm) 2': 'N30',

    }
    copias_factibilidad= {
    'H33': 'H8',
    'L35': 'L10',
    'V35': 'V10',
    'AC35': 'AD10',
    'AI35': 'AN10',
    'L36': 'L11',
    'V36': 'V11',
    'AF36': 'AF11',
    'U37': 'U12',
    'AI37': 'AL12',
    'C40': 'C15',
    'N40': 'N15',
    'AC40': 'AC15',
    'AH40': 'AK15',
    'P45': 'P20',
    'P46': 'P21',
    'P47': 'P22',
    'P48': 'P23',
    'P49': 'P24',
    'P50': 'P25',
    'P51': 'P26',   
    'P52': 'P27',       
    'P54': 'P29',
    'H57': 'H32',
    'L59': 'L34',
    'V59': 'V34',
    'AC59': 'AD34',
    'AI59': 'AN34',
    'L60': 'L35',
    'V60': 'V35',
    'AF60': 'AF35',
    'U61': 'U36',
    'AI61': 'AL36',
    'C64': 'C39',
    'N64': 'N39',
    'AC64': 'AC39',
    'AH64': 'AK39',
    'P69': 'P44',   
    'P70': 'P45',
    'P71': 'P46',
    'P72': 'P47',
    'P73': 'P48',
    'P74': 'P49',
    'P75': 'P50',
    'P76': 'P51',
    'P78': 'P53',
    }
    copias_checkbox_factibilidad = {
    'H34': 'H9',   # autosoportada
    'P34': 'P9',   # arriostrada
    'W34': 'V9',   # monopolo
    'AC34': 'AC9', # minipolo
    'AH34': 'AH9', # otro
    'V40' : 'V15', # espacio disponible de conexión
    'Z40': 'Z15', # no espacio disponible de conexión
    'S42': 'Z17', # cara de preparación A
    'X42': 'AE17', # cara de preparación B
    'AC42': 'AJ17', # cara de preparación C
    'AH42': 'AO17', # cara de preparación D
    'P53': 'P28', # barra de tierra 
    'S53': 'S28', # no barra de tierra
    'P55': 'P30', # tipo de solución piso
    'S55': 'S30', # tipo de solución torre
    'X47': 'Z22', # existe breaker existente en sitio
    'AB47': 'AC22', # no existe breaker existente en sitio
    'Y51': 'Z26', # alimentacion compatible con el equipamiento
    'AB51': 'AC26', # no alimentacion compatible con el equipamiento
    'AG47': 'AH22', # sistema electrico monofasica
    'AJ47': 'AM22', # sistema electrico bifasica
    'AG51': 'AI26', # espacio disponible de conexión
    'AJ51': 'AL26', # no espacio disponible de conexión
    'H58': 'H33',   # autosoportada 2
    'P58': 'P33',   # arriostrada 2
    'W58': 'W33',   # monopolo 2
    'AC58': 'AC33', # minipolo 2
    'AH58': 'AH33', # otro 2
    'V64': 'V39', # espacio disponible de conexión 2
    'Z64': 'Z39', # no espacio disponible de conexión 2
    'S66': 'Z41', # cara de preparación A 2
    'X66': 'AE41', # cara de preparación B 2
    'AC66': 'AJ41', # cara de preparación C 2
    'AH66': 'AO41', # cara de preparación D 2
    'P77': 'P52', # existe barra de tierra 2
    'S77': 'S52', # no existe barra de tierra 2
    'P79': 'P54', # tipo de solución piso 2
    'S79': 'S54', # tipo de solución torre 2
    'Y71': 'Z46', # existe breaker existente en sitio 2
    'AB71': 'AC46', # no existe breaker existente en sitio 2
    'AG71': 'AH46', # sistema electrico monofasica 2
    'AJ71': 'AM46', # sistema electrico bifasica 2
    'Y75': 'Z50', # alimentacion compatible con el equipamiento 2
    'AB75': 'AC50', # no alimentacion compatible con el equipamiento
    'AG75': 'AI50', # espacio disponible de conexión 2
    'AJ75': 'AL50', # no espacio disponible de conexión 2


    }
    
    
    for campo, celda in campos_a_celdas.items():
        valor = normaliza_na(datos.get(campo, ""))
        if isinstance(celda, list):
            for c in celda:
                ws_a.range(c).value = valor
        else:
            ws_a.range(celda).value = valor

    # --- Asignación de región basada en el estado (MOVIDO AQUÍ) ---
    # Leer directamente el estado que ya está en AD14 (después de que se haya escrito)
    print(f"DEBUG: Antes de leer AD14 - ws_a.name = {ws_a.name}")
    estado = ws_a.range('AD14').value
    print(f"DEBUG: Estado leído de AD14: '{estado}'")
    print(f"DEBUG: Tipo de estado: {type(estado)}")
    print(f"DEBUG: Estado después de strip: '{str(estado).strip()}'")
    print(f"DEBUG: Estado en estado_a_region: {'Sí' if str(estado).strip() in estado_a_region else 'No'}")
    region = estado_a_region.get(str(estado).strip(), 'OTRA')
    print(f"DEBUG: Región asignada: '{region}'")
    ws_a.range('E9').value = region
    print(f"DEBUG: Región escrita en E9: '{region}'")

    

    for campo, celda in campos_b_celdas.items():
        valor = normaliza_na(datos.get(campo, ""))
        if isinstance(celda, list):
            for c in celda:
                ws_b.range(c).value = valor
        else:
            ws_b.range(celda).value = valor
    
    for origen, destino in copias_factibilidad.items():
     ws_factibilidad.range(destino).value = ws_a.range(origen).value

    for origen, destino in copias_checkbox_factibilidad.items():
        if origen == 'Y47' and destino == 'Z22':
            ws_factibilidad.range(destino).value = bool(ws_a.range(origen).value)
        else:
            ws_factibilidad.range(destino).value = ws_a.range(origen).value
    
    print("DEBUG: Archivos recibidos en request.files:", list(request.files.keys()))
    imagenes_torres = request.files.getlist('imagenes_torres')
    
    imagenes_torres = request.files.getlist('imagenes_torres')
    imagenes_torres_paths = []
    for idx, img in enumerate(imagenes_torres):
        if img and img.filename:
            filename = secure_filename(f"torres_{idx}_{img.filename}")
            img_path = os.path.join(UPLOAD_FOLDER, filename)
            img.save(img_path)
            imagenes_torres_paths.append(img_path) 

    imagenes_torres_b = request.files.getlist('imagenes_torres_b')
    imagenes_torres_b_paths = []
    for idx, img in enumerate(imagenes_torres_b):
        if img and img.filename:
            filename = secure_filename(f"torres_b_{idx}_{img.filename}")
            img_path = os.path.join(UPLOAD_FOLDER, filename)
            img.save(img_path)
            imagenes_torres_b_paths.append(img_path)

           
    # Reordena las imágenes según tu lógica visual
    # Asegúrate de que haya al menos 6 imágenes para evitar errores de índice
    print(f"DEBUG: imagen_paths tiene {len(imagen_paths)} elementos")
    print(f"DEBUG: imagen_paths = {imagen_paths}")
    print(f"DEBUG: ws_a.name = {ws_a.name}")
    
    if len(imagen_paths) >= 6:
        ordenadas = [imagen_paths[5], imagen_paths[4], imagen_paths[3], imagen_paths[2], imagen_paths[1], imagen_paths[0]]
        img_cells = ['C87', 'C87', 'E118', 'E140', 'E163', 'E185']
        imagenes_final = ordenadas
    else:
        img_cells = ['C87', 'C87', 'E118', 'E140', 'E163', 'E185'][:len(imagen_paths)]
        imagenes_final = imagen_paths[::-1]  # Invierte el orden si quieres de derecha a izquierda

    # Inserta las imágenes en las celdas correspondientes
    img_ranges = ['C86:O113', 'Y86:AK113', 'E118:O136', 'E140:O158', 'E163:O181', 'E185:O203']
    
    print(f"DEBUG: imagenes_final = {imagenes_final}")
    print(f"DEBUG: img_ranges = {img_ranges}")
    
    # Limpia imágenes previas en la hoja antes de insertar nuevas
    print(f"DEBUG: Limpiando {ws_a.pictures.count} imágenes previas en {ws_a.name}")
    for pic in ws_a.pictures:
        try:
            pic.delete()
            print("DEBUG: Imagen previa eliminada")
        except Exception as e:
            print(f"Error eliminando imagen previa en ws_a: {e}")

    # Depuración: imprime las rutas de las imágenes a insertar
    print("Imágenes a insertar en 4. Estudio de informacion A:")
    for idx, img in enumerate(imagenes_final):
        print(f"Imagen {idx}: {img} - Existe: {os.path.exists(img) if img else False}")

    for idx, img_path in enumerate(imagenes_final):
        print(f"Intentando insertar imagen {idx}: {img_path} ...", end="")
        if idx < len(img_ranges) and img_path and os.path.exists(img_path):
            try:
                cell_range = ws_a.range(img_ranges[idx])
                print(f"DEBUG: cell_range = {img_ranges[idx]}, left={cell_range.left}, top={cell_range.top}")
                ws_a.pictures.add(
                    os.path.abspath(img_path),
                    left=cell_range.left,
                    top=cell_range.top,
                    width=cell_range.width,
                    height=cell_range.height
                )
                print("OK")
            except Exception as e:
                print(f"ERROR: {e}")
        else:
            print("NO EXISTE")
    
    print(f"DEBUG: Después de insertar, {ws_a.pictures.count} imágenes en {ws_a.name}")
    
    # --- INSERTAR IMÁGENES DE PLANOS A Y B ---
    
    # Planos A (en la hoja 4. Estudio de informacion A)
    planos_a_ranges = ['C17:AK60', 'C69:AK123', 'C134:AK173']
    planos_a_paths = [planos_a_img1_path, planos_a_img2_path, planos_a_img3_path]
    
    print("Imágenes de Planos A a insertar:")
    for idx, img_path in enumerate(planos_a_paths):
        print(f"Plano A {idx+1}: {img_path} - Existe: {os.path.exists(img_path) if img_path else False}")
    
    for idx, img_path in enumerate(planos_a_paths):
        if idx < len(planos_a_ranges) and img_path and os.path.exists(img_path):
            try:
                cell_range = ws_a.range(planos_a_ranges[idx])
                ws_a.pictures.add(
                    os.path.abspath(img_path),
                    left=cell_range.left,
                    top=cell_range.top,
                    width=cell_range.width,
                    height=cell_range.height
                )
                print(f"Plano A {idx+1} insertado correctamente")
            except Exception as e:
                print(f"Error insertando Plano A {idx+1}: {e}")
    
    # Planos B (en la hoja 5. Estudio de informacion B)
    ws_b = wb.sheets['5. Estudio de informacion B']
    planos_b_ranges = ['C17:AK60', 'C69:AK123', 'C134:AK173']
    planos_b_paths = [planos_b_img1_path, planos_b_img2_path, planos_b_img3_path]
    
    print("Imágenes de Planos B a insertar:")
    for idx, img_path in enumerate(planos_b_paths):
        print(f"Plano B {idx+1}: {img_path} - Existe: {os.path.exists(img_path) if img_path else False}")
    
    for idx, img_path in enumerate(planos_b_paths):
        if idx < len(planos_b_ranges) and img_path and os.path.exists(img_path):
            try:
                cell_range = ws_b.range(planos_b_ranges[idx])
                ws_b.pictures.add(
                    os.path.abspath(img_path),
                    left=cell_range.left,
                    top=cell_range.top,
                    width=cell_range.width,
                    height=cell_range.height
                )
                print(f"Plano B {idx+1} insertado correctamente")
            except Exception as e:
                print(f"Error insertando Plano B {idx+1}: {e}")

    ws_torres = wb.sheets['6. Estudio torres y antenas A']
    img_torres_ranges = ['C17:AK60', 'C69:AK123', 'C134:AK173']
    # Limpia imágenes previas en la hoja antes de insertar nuevas
    print(f"DEBUG: Limpiando {ws_torres.pictures.count} imágenes previas en {ws_torres.name}")
    for pic in ws_torres.pictures:
        try:
            pic.delete()
            print("DEBUG: Imagen previa eliminada en hoja 6")
        except Exception as e:
            print(f"Error eliminando imagen previa en ws_torres: {e}")
    print("Imágenes a insertar en 6. Estudio torres y antenas A:")
    for idx, img_path in enumerate(imagenes_torres_paths):
        print(f"Imagen {idx}: {img_path} - Existe: {os.path.exists(img_path) if img_path else False}")
    for idx, img_path in enumerate(imagenes_torres_paths):
        print(f"Intentando insertar imagen {idx}: {img_path} ...", end="")
        if idx < len(img_torres_ranges) and os.path.exists(img_path):
            try:
                cell_range = ws_torres.range(img_torres_ranges[idx])
                print(f"DEBUG: cell_range = {img_torres_ranges[idx]}, left={cell_range.left}, top={cell_range.top}")
                ws_torres.pictures.add(
                    os.path.abspath(img_path),
                    left=cell_range.left,
                    top=cell_range.top,
                    width=cell_range.width,
                    height=cell_range.height
                )
                print("OK")
            except Exception as e:
                print(f"ERROR: {e}")
        else:
            print("NO EXISTE")
    print(f"DEBUG: Después de insertar, {ws_torres.pictures.count} imágenes en {ws_torres.name}")

    ws_torres_b = wb.sheets['7. Estudio torres y antenas B']
    img_torres_b_ranges = ['C15:AK57', 'C67:AK119', 'C132:AK169']
    for idx, img_path in enumerate(imagenes_torres_b_paths):
        if idx < len(img_torres_b_ranges) and os.path.exists(img_path):
            cell_range = ws_torres_b.range(img_torres_b_ranges[idx])
            ws_torres_b.pictures.add(
                os.path.abspath(img_path),
                left=cell_range.left,
                top=cell_range.top,
                width=cell_range.width,
                height=cell_range.height
            )

    ws_fotos9 = wb.sheets['9. Factibilidad Reporte Fotos A']
    fotos9_celdas = [
    'H11:L18',  # 1. GPS con coordenadas de la torre
    'Y11:AC18',  # 2. Fachada del sitio
    'H23:L29',  # 3. Foto de torre completa
    'H48:L54',  # 4. Foto desde piso mostrando espacio en torre para MW topología
    'Y48:AC54',  # 5. Medición con cinta del rad center en torre topología
    'H59:L66',  # 6. Foto desde piso mostrando espacio en torre para MW (SD)
    'Y59:AC66',  # 7. Medición con cinta del rad center en torre (SD)
    'H70:L77',  # 8. Foto desde piso mostrando espacio (propuesto) en torre para antena
    'Y70:AC77',  # 9. Foto desde piso mostrando espacio (propuesto) en torre para antena
    'H87:L94',  # 10. Foto línea de Vista de Sitio A a Sitio B
    'Y87:AC94',  # 11. Foto línea de Vista de Sitio A a Sitio B Diversidad
    'H98:L105',  # 12. Foto Barra de Tierra
    'Y98:AC105',  # 13. Foto de escalerilla de torre
    'H127:L134', # 14. Foto del espacio disponible dentro del Gabinete OMB
    'Y127:AC134', # 15. Foto del espacio disponible en torre para OMB adicional
    'H138:L144', # 16. Foto DPU existente
    'Y138:AC144', # 17. Foto del espacio disponible en torre para DPU y Batería
    'H150:L156', # 18. Foto ACDB y Breaker
    'Y150:AC156', # 19. Foto de Agregador (Site Entry)
    ]
    

    while len(imagenes_fotos9_paths) < len(fotos9_celdas):
        imagenes_fotos9_paths.append(None)

    for idx, celda in enumerate(fotos9_celdas):
        img_path = imagenes_fotos9_paths[idx]
        cell_range = ws_fotos9.range(celda)
        if img_path and os.path.exists(img_path):
            ws_fotos9.pictures.add(
            os.path.abspath(img_path),
            left=cell_range.left,
            top=cell_range.top,
            width=cell_range.width,
            height=cell_range.height
    )
        else:
        # Si no hay imagen, coloca un N/A grande centrado en la celda superior izquierda
            cell = ws_fotos9.range(celda.split(':')[0])
            cell.value = "N/A"
            cell.api.Font.Size = 36
            cell.api.HorizontalAlignment = -4108  # xlCenter
            cell.api.VerticalAlignment = -4108    # xlCenter
    

    # Usar las imágenes de fotos 10 ya procesadas anteriormente
    fotos10_names = [
         "foto_b_1", "foto_b_2", "foto_b_3", "foto_b_4", "foto_b_5", "foto_b_6", "foto_b_7", "foto_b_8", "foto_b_9", "foto_b_10"
    ]
    imagenes_fotos10_paths = [fotos10_paths.get(name) for name in fotos10_names]

    ws_fotos10 = wb.sheets['10. Reporte Fotos B']
    fotos10_celdas = [
    'H11:L18', 'Y11:AC18', 'H23:L29', 'H48:L54', 'Y48:AC54', 'H59:L66', 'Y59:AC66',
    'H70:L77', 'Y70:AC77', 'H87:L94', 'Y87:AC94'
    ]
    while len(imagenes_fotos10_paths) < len(fotos10_celdas):
        imagenes_fotos10_paths.append(None)
    for idx, celda in enumerate(fotos10_celdas):
        img_path = imagenes_fotos10_paths[idx]
        cell_range = ws_fotos10.range(celda)
        if img_path and os.path.exists(img_path):
            ws_fotos10.pictures.add(
                os.path.abspath(img_path),
                left=cell_range.left,
                top=cell_range.top,
                width=cell_range.width,
                height=cell_range.height
           )
        else:
            # Si no hay imagen, coloca un N/A grande centrado en la celda superior izquierda
            cell = ws_fotos10.range(celda.split(':')[0])
            cell.value = "N/A"
            cell.api.Font.Size = 36
            cell.api.HorizontalAlignment = -4108  # xlCenter
            cell.api.VerticalAlignment = -4108    # xlCenter

    pdf_icon_cells = ['AB121', 'AB166', 'AB200', 'AB220', 'AB240', 'AB260']

    # --- Inserta KMZ e imagen KML en la hoja de Formato KMZ ---
  #  print("KMZ path:", kmz_path)
   # print("KMZ exists:", os.path.exists(kmz_path) if kmz_path else False)
  #  print("KML image path:", kml_image_path)
   # print("KML image exists:", os.path.exists(kml_image_path) if kml_image_path else False)
   # print("ws_kmz name:", ws_kmz.name)
   # print("Antes de insertar imagen:", ws_kmz.pictures.count)
    # Inserta imagen KML en la hoja de Formato KMZ
   # if kml_image_path and os.path.exists(kml_image_path):
      #  try:
       #    cell = ws_kmz.range('C21')  # Usa 'C21' para centrar la imagen
      #     ws_kmz.pictures.add(
          #     os.path.abspath(kml_image_path),
           #    left=cell.left,
           #    top=cell.top,
           #    width=cell.width * 6,   # Ajusta el ancho para cubrir varias columnas si lo deseas
          #     height=180
        #  )
         #  print("Imagen KML insertada")
      #  except Exception as e:
        #    print(f"Error al insertar imagen KML: {e}")
   # print("Después de insertar imagen:", ws_kmz.pictures.count)
    # Guarda y cierra el archivo solo una vez
    print("KML IMAGE PATH:", kml_image_path)
    print("EXISTS:", os.path.exists(kml_image_path) if kml_image_path else False)
    print("SIZE:", os.path.getsize(kml_image_path) if kml_image_path and os.path.exists(kml_image_path) else "NO FILE")
    # Hoja 3: Formato KMZ
    kmz_img_path = imagen_kmz_path if imagen_kmz_path and os.path.exists(imagen_kmz_path) else (
        kml_image_path if kml_image_path and os.path.exists(kml_image_path) else None
    )
    if kmz_img_path:
        ws_kmz = wb.sheets['3. Formato KMZ']
        cell_range = ws_kmz.range('B21:F38')  # O el rango que desees
        ws_kmz.pictures.add(
            os.path.abspath(kmz_img_path),
            left=cell_range.left,
            top=cell_range.top,
            width=cell_range.width,
            height=cell_range.height
        )

# Hoja 5: Estudio de información B
    if imagen_b_path and os.path.exists(imagen_b_path):
        ws_b = wb.sheets['5. Estudio de informacion B']
        cell_range = ws_b.range('B36:AJ45')
        ws_b.pictures.add(
            os.path.abspath(imagen_b_path),
            left=cell_range.left,
            top=cell_range.top,
            width=cell_range.width,
            height=cell_range.height
      )
    
    
    
    # Intentar guardar con múltiples estrategias
    guardado_exitoso = False
    max_intentos_guardado = 3
    
    for intento in range(max_intentos_guardado):
        try:
            print(f"DEBUG: Intento {intento + 1} de guardar archivo en: {output_path}")
            wb.save(output_path)
            print(f"DEBUG: Guardado exitoso en intento {intento + 1}")
            guardado_exitoso = True
            break
        except Exception as e:
            print(f"DEBUG: Error en intento {intento + 1}: {e}")
            
            if intento < max_intentos_guardado - 1:
                # Intentar guardar como backup
                try:
                    time_module = __import__('time')
                    backup_path = output_path.replace('.xlsx', f'_backup_{int(time_module.time())}.xlsx')
                    print(f"DEBUG: Intentando guardar como backup: {backup_path}")
                    wb.save(backup_path)
                    print(f"DEBUG: Guardado como backup exitoso: {backup_path}")
                    output_path = backup_path
                    guardado_exitoso = True
                    break
                except Exception as backup_e:
                    print(f"DEBUG: Error en backup: {backup_e}")
                    time.sleep(2)
            else:
                print(f"DEBUG: Todos los intentos de guardado fallaron")
    
    if not guardado_exitoso:
        wb.close()
        app_excel.quit()
        return f"Error: No se pudo guardar el archivo después de {max_intentos_guardado} intentos"
    import os
    if not os.path.exists(output_path):
        print(f"Error: El archivo de salida no se pudo guardar correctamente en: {output_path}")
        wb.close()
        app_excel.quit()
        return f"Error: El archivo de salida no se pudo guardar correctamente en: {output_path}"
    time.sleep(1)
    
    # Cerrar Excel de forma segura
    try:
        wb.close()
        print("DEBUG: Workbook cerrado correctamente")
    except Exception as close_error:
        print(f"DEBUG: Error al cerrar workbook: {close_error}")
    
    try:
        app_excel.quit()
        print("DEBUG: Excel cerrado correctamente")
    except Exception as quit_error:
        print(f"DEBUG: Error al cerrar Excel: {quit_error}")
        # Forzar cierre de Excel
        try:
            subprocess.run(['taskkill', '/f', '/im', 'excel.exe'], capture_output=True)
            print("DEBUG: Excel forzado a cerrar")
        except:
            print("DEBUG: No se pudo forzar el cierre de Excel")

    # --- 3. Inserta archivos como OLEObjects (íconos) usando win32com ---
    try:
        try:
            excel = win32com.client.Dispatch("Excel.Application")
        except Exception as e:
            return f"Error al inicializar Excel para OLE: {e}"
        # Chequeo previo: existencia y permisos del archivo
        import os
        if not os.path.exists(output_path):
            return f"Error: El archivo de salida no existe en la ruta esperada: {output_path}"
        if not os.access(output_path, os.R_OK | os.W_OK):
            return f"Error: No tienes permisos de lectura/escritura para el archivo: {output_path}"
        # try/except para Workbooks.Open
        try:
            wb_com = excel.Workbooks.Open(output_path)
        except Exception as e:
            try:
                if 'excel' in locals() and excel is not None:
                    excel.Quit()
            except Exception:
                pass
            return f"Error al abrir el archivo Excel para OLE: {e}"
        if wb_com is None:
            try:
                if 'excel' in locals() and excel is not None:
                    excel.Quit()
            except Exception:
                pass
            # Intenta eliminar el archivo corrupto o bloqueado
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
                return "Error: No se pudo abrir el archivo Excel para incrustar archivos OLE. El archivo fue eliminado automáticamente por estar corrupto o bloqueado. Por favor, vuelve a intentar el proceso."
            except Exception as e:
                return f"Error: No se pudo abrir ni eliminar el archivo Excel de salida. Detalle: {e}"
        # Inserta el archivo Excel como objeto en N52 de la hoja 5
        if archivo_excel_b_path and os.path.exists(archivo_excel_b_path):
            ws_b_com = wb_com.Sheets("5. Estudio de informacion B")
            ws_b_com.OLEObjects().Add(
                Filename=archivo_excel_b_path,
                Link=False,
                DisplayAsIcon=True,
                IconFileName="C:\\Windows\\System32\\shell32.dll",
                IconIndex=1,
                IconLabel=os.path.basename(archivo_excel_b_path),
                Left=ws_b_com.Range("N52").Left,
                Top=ws_b_com.Range("N52").Top
            )

        if word_file_path and os.path.exists(word_file_path):
            ws_word = wb_com.Sheets("1. Analisis de Red y Frecuencia")
            ws_word.OLEObjects().Add(
                Filename=word_file_path,
                Link=False,
                DisplayAsIcon=True,
                IconFileName="C:\\Windows\\System32\\shell32.dll",
                IconIndex=2,  # Cambia el icono si lo deseas
                IconLabel=os.path.basename(word_file_path),
                Left=ws_word.Range("D12").Left,
                Top=ws_word.Range("D12").Top
           ) 
        


        # Incrusta los PDF en la hoja "4. Estudio de informacion A"
        ws_a_com = wb_com.Sheets("4. Estudio de informacion A")
        pdf_icon_cells = ['AB121', 'AB166', 'AB200', 'AB220', 'AB240', 'AB260']
        print(f"DEBUG: Intentando insertar {len(pdf_paths)} PDFs en la hoja 4. Estudio de informacion A")
        for idx, pdf_path in enumerate(pdf_paths):
            print(f"DEBUG: PDF {idx}: {pdf_path} - Existe: {os.path.exists(pdf_path) if pdf_path else False}")
            if idx < len(pdf_icon_cells) and os.path.exists(pdf_path):
                try:
                    ws_a_com.OLEObjects().Add(
                        Filename=pdf_path,
                        Link=False,
                        DisplayAsIcon=True,
                        IconFileName="C:\\Windows\\System32\\shell32.dll",
                        IconIndex=0,
                        IconLabel=os.path.basename(pdf_path),
                        Left=ws_a_com.Range(pdf_icon_cells[idx]).Left,
                        Top=ws_a_com.Range(pdf_icon_cells[idx]).Top
                    )
                    print(f"DEBUG: PDF {idx} insertado correctamente en {pdf_icon_cells[idx]}")
                except Exception as e:
                    print(f"DEBUG: Error insertando PDF {idx}: {e}")
            else:
                print(f"DEBUG: PDF {idx} no se insertó - no existe o índice fuera de rango")

        # Incrusta el KMZ en la hoja "3. Formato KMZ"
        if kmz_path and os.path.exists(kmz_path):
            ws_kmz_com = wb_com.Sheets("3. Formato KMZ")
            ws_kmz_com.OLEObjects().Add(
                Filename=kmz_path,
                Link=False,
                DisplayAsIcon=True,
                IconFileName="C:\\Windows\\System32\\shell32.dll",
                IconIndex=0,
                IconLabel=os.path.basename(kmz_path),
                Left=ws_kmz_com.Range("C12").Left,
                Top=ws_kmz_com.Range("C12").Top
            )

        wb_com.Save()
        wb_com.Close()
        excel.Quit()
    except Exception as e:
        return f"Error al incrustar archivos OLE: {e}"

    time.sleep(1)
    
    # Verificar el tipo de llenado para mostrar la confirmación apropiada
    if tipo == 'diseno_solucion':
        # Generar página de confirmación con el mismo estilo que Site Survey
        html = f"""
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <title>FANGIO TELECOM | Documento Generado</title>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
            <style>
                * {{
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                }}

                body {{
                    font-family: 'Montserrat', Arial, sans-serif;
                    min-height: 100vh;
                    background-color: #0a192f;
                    background-size: cover;
                    background-position: center;
                    background-attachment: fixed;
                    background-repeat: no-repeat;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                    align-items: center;
                    position: relative;
                    overflow-x: hidden;
                    color: #e0e7ef;
                }}

                /* Overlay para mejorar la legibilidad */
                body::before {{
                    content: '';
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: rgba(10, 25, 47, 0.7);
                    z-index: -1;
                }}

                /* Estrellas animadas */
                body::after {{
                    content: '';
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background-image: 
                        radial-gradient(2px 2px at 20px 30px, #eee, transparent),
                        radial-gradient(2px 2px at 40px 70px, rgba(255,255,255,0.8), transparent),
                        radial-gradient(1px 1px at 90px 40px, #fff, transparent),
                        radial-gradient(1px 1px at 130px 80px, rgba(255,255,255,0.6), transparent),
                        radial-gradient(2px 2px at 160px 30px, #ddd, transparent);
                    background-repeat: repeat;
                    background-size: 200px 100px;
                    animation: sparkle 4s linear infinite;
                    z-index: -2;
                }}

                @keyframes sparkle {{
                    from {{ transform: translateY(0px); }}
                    to {{ transform: translateY(-100px); }}
                }}

                .main-container {{
                    width: 100%;
                    max-width: 500px;
                    margin: 0 auto;
                    padding: 20px;
                    z-index: 1;
                }}

                .success-card {{
                    background: rgba(22, 33, 62, 0.95);
                    backdrop-filter: blur(10px);
                    border: 2px solid rgba(0, 195, 255, 0.3);
                    border-radius: 20px;
                    padding: 40px 30px;
                    text-align: center;
                    box-shadow: 0 20px 40px rgba(0, 195, 255, 0.2);
                    position: relative;
                    overflow: hidden;
                }}

                .success-card::before {{
                    content: '';
                    position: absolute;
                    top: 0;
                    left: -100%;
                    width: 100%;
                    height: 100%;
                    background: linear-gradient(90deg, transparent, rgba(0, 195, 255, 0.1), transparent);
                    transition: left 0.5s;
                }}

                .success-card:hover::before {{
                    left: 100%;
                }}

                .success-icon {{
                    width: 80px;
                    height: 80px;
                    background: linear-gradient(135deg, #00c37a 0%, #00a870 100%);
                    border-radius: 50%;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    margin: 0 auto 20px;
                    box-shadow: 0 10px 30px rgba(0, 195, 122, 0.4);
                    animation: pulse 2s infinite;
                }}

                @keyframes pulse {{
                    0% {{ transform: scale(1); }}
                    50% {{ transform: scale(1.05); }}
                    100% {{ transform: scale(1); }}
                }}

                .success-icon i {{
                    font-size: 2.5rem;
                    color: white;
                }}

                .success-title {{
                    color: #00c37a;
                    font-size: 2rem;
                    font-weight: 700;
                    margin-bottom: 30px;
                    text-shadow: 0 2px 10px rgba(0, 195, 122, 0.3);
                }}

                .document-info {{
                    background: rgba(26, 35, 58, 0.8);
                    border-radius: 15px;
                    padding: 25px;
                    margin-bottom: 30px;
                    border-left: 5px solid #00c3ff;
                }}

                .info-row {{
                    display: flex;
                    align-items: center;
                    margin-bottom: 15px;
                    font-size: 1.1rem;
                }}

                .info-row:last-child {{
                    margin-bottom: 0;
                }}

                .info-row i {{
                    color: #00c3ff;
                    margin-right: 12px;
                    font-size: 1.2rem;
                    width: 20px;
                    text-align: center;
                }}

                .info-label {{
                    color: #b5c7e6;
                    font-weight: 600;
                    margin-right: 10px;
                }}

                .info-value {{
                    color: #ffffff;
                    font-weight: 500;
                }}

                .buttons-container {{
                    display: flex;
                    flex-direction: column;
                    gap: 15px;
                }}

                .download-button, .new-document-button, .back-button, .other-button {{
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 10px;
                    padding: 18px 30px;
                    border-radius: 12px;
                    text-decoration: none;
                    font-weight: 600;
                    font-size: 1.1rem;
                    transition: all 0.3s ease;
                    border: none;
                    cursor: pointer;
                    position: relative;
                    overflow: hidden;
                }}

                .download-button {{
                    background: linear-gradient(135deg, #00c3ff 0%, #0099cc 100%);
                    color: #ffffff;
                    box-shadow: 0 8px 25px rgba(0, 195, 255, 0.4);
                }}

                .download-button:hover {{
                    background: linear-gradient(135deg, #0099cc 0%, #00c3ff 100%);
                    transform: translateY(-2px);
                    box-shadow: 0 12px 35px rgba(0, 195, 255, 0.6);
                }}

                .back-button {{
                    background: rgba(0, 195, 255, 0.1);
                    color: #00c3ff;
                    border: 2px solid rgba(0, 195, 255, 0.3);
                }}

                .back-button:hover {{
                    background: rgba(0, 195, 255, 0.2);
                    border-color: rgba(0, 195, 255, 0.5);
                    transform: translateY(-2px);
                }}

                .other-button {{
                    background: linear-gradient(135deg, #4a90e2 0%, #357abd 100%);
                    color: #ffffff;
                    box-shadow: 0 8px 25px rgba(74, 144, 226, 0.4);
                }}

                .other-button:hover {{
                    background: linear-gradient(135deg, #357abd 0%, #4a90e2 100%);
                    transform: translateY(-2px);
                    box-shadow: 0 12px 35px rgba(74, 144, 226, 0.6);
                }}

                .footer {{
                    text-align: center;
                    margin-top: 30px;
                    color: #8892b0;
                    font-size: 0.9rem;
                }}

                .header {{
                    text-align: center;
                    margin-bottom: 30px;
                }}

                .header img {{
                    width: 120px;
                    height: auto;
                    margin-bottom: 15px;
                }}

                .header h2 {{
                    color: #00c3ff;
                    font-size: 1.5rem;
                    font-weight: 600;
                    margin: 0;
                }}
            </style>
        </head>
        <body>
            <div class="main-container">
                <div class="header">
                    <img src="{url_for('static', filename='images/fangio-logo.svg')}" alt="FANGIO TELECOM">
                    <h2>FANGIO TELECOM</h2>
                </div>
                
                <div class="success-card">
                    <div class="success-icon">
                        <i class="fas fa-check"></i>
                    </div>
                    
                    <h1 class="success-title">¡Documento Generado!</h1>
                    
                    <div class="document-info">
                        <div class="info-row">
                            <i class="fas fa-file-alt"></i>
                            <span class="info-label">ID:</span>
                            <span class="info-value">{user_id}</span>
                        </div>
                        <div class="info-row">
                            <i class="fas fa-map-marker-alt"></i>
                            <span class="info-label">Sitio A:</span>
                            <span class="info-value">{datos.get('Nombre del sitio A', '')}</span>
                        </div>
                        <div class="info-row">
                            <i class="fas fa-map-marker-alt"></i>
                            <span class="info-label">Sitio B:</span>
                            <span class="info-value">{datos.get('Nombre del sitio B', '')}</span>
                        </div>
                    </div>
                    
                    <div class="buttons-container">
                        <a href="{url_for('descargar_diseno_solucion', user_id=user_id, fila_idx=fila_idx)}" class="download-button">
                            <i class="fas fa-download"></i>
                            Descargar Archivo Generado
                        </a>
                        
                        <a href="{url_for('site_survey', user_id=user_id, fila_idx=fila_idx)}" class="other-button">
                            <i class="fas fa-clipboard-check"></i>
                            Ir a Site Survey
                        </a>
                        
                        <a href="{url_for('formulario_archivos', user_id=user_id, fila_idx=fila_idx)}" class="other-button">
                            <i class="fas fa-file-upload"></i>
                            Ir a Formulario de Archivos
                        </a>
                        
                        <a href="/" class="back-button">
                            <i class="fas fa-home"></i>
                            Volver al Inicio
                        </a>
                    </div>
                </div>
                
                <div class="footer">
                    <p>© 2024 FANGIO TELECOM. Todos los derechos reservados.</p>
                </div>
            </div>

            <script>
                // Verificar que la imagen de fondo se cargue correctamente
                const backgroundImage = new Image();
                backgroundImage.onload = function() {{
                    document.body.style.backgroundImage = 'url("{url_for('static', filename='images/earth-background.jpg')}")';
                    console.log('Imagen de fondo cargada correctamente');
                }};
                backgroundImage.onerror = function() {{
                    console.log('Error cargando imagen de fondo, usando fallback');
                    document.body.style.backgroundImage = 'url("/static/images/earth-background.jpg")';
                }};
                backgroundImage.src = "{url_for('static', filename='images/earth-background.jpg')}";
            </script>
        </body>
        </html>
        """
        
        return render_template_string(html)
    else:
        # Para site_survey y otros tipos, usar el comportamiento original
        return send_file(output_path, as_attachment=True)

@app.route('/descargar_diseno_solucion')
def descargar_diseno_solucion():
    import pandas as pd
    user_id = request.args.get('user_id')
    fila_idx = request.args.get('fila_idx')
    
    # Obtener datos de la base de datos
    try:
        df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
        if fila_idx:
            datos = df_db.loc[int(fila_idx)]
        else:
            coincidencias = df_db[df_db['ID'] == user_id]
            if coincidencias.empty:
                return "ID no encontrado en la base de datos."
            datos = coincidencias.iloc[0]
    except Exception as e:
        return f"Error leyendo la base de datos: {e}"
    
    # Construir el nombre del archivo de salida
    output_filename = f"WORKING_{int(time.time())}_llenadoauto.xlsx"
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)
    
    if not os.path.exists(output_path):
        return "Archivo no encontrado. Por favor, genera el documento nuevamente."
    
    def limpiar_nombre_archivo(nombre):
        # Eliminar caracteres problemáticos para nombres de archivo
        caracteres_invalidos = '<>:"/\\|?*'
        for char in caracteres_invalidos:
            nombre = nombre.replace(char, '_')
        return nombre
    
    # Crear un nombre de archivo más amigable
    nombre_archivo = f"Diseno_Solucion_{user_id}_{limpiar_nombre_archivo(datos.get('Nombre del sitio A', ''))}.xlsx"
    
    @after_this_request
    def eliminar_archivos_temporales(response):
        try:
            if os.path.exists(output_path):
                os.remove(output_path)
        except Exception as e:
            print(f"Error eliminando archivo temporal: {e}")
        return response
    
    return send_file(output_path, as_attachment=True, download_name=nombre_archivo)

@app.route('/seleccion_tipo_llenado')
def seleccion_tipo_llenado():
    user_id = request.args.get('user_id')
    fila_idx = request.args.get('fila_idx')
    return render_template('seleccion_tipo_llenado.html', user_id=user_id, fila_idx=fila_idx)

@app.route('/seleccion_llenado_ptp')
def seleccion_llenado_ptp():
    user_id = request.args.get('user_id')
    fila_idx = request.args.get('fila_idx')
    return render_template('seleccion_llenado_ptp.html', user_id=user_id, fila_idx=fila_idx)

@app.route('/seleccion_llenado_ptmp')
def seleccion_llenado_ptmp():
    user_id = request.args.get('user_id')
    fila_idx = request.args.get('fila_idx')
    return render_template('seleccion_llenado_ptmp.html', user_id=user_id, fila_idx=fila_idx)

@app.route('/redirigir_tipo_llenado', methods=['POST'])
def redirigir_tipo_llenado():
    tipo = request.form.get('tipo')
    user_id = request.form.get('user_id')
    fila_idx = request.form.get('fila_idx')
    print(f"DEBUG tipo recibido: '{tipo}'")
    if tipo and tipo.strip().lower() == 'ptp':
        # Redirigir a la selección de tipo de llenado para PtP
        return redirect(url_for('seleccion_llenado_ptp', user_id=user_id, fila_idx=fila_idx))
    elif tipo and tipo.strip().lower() == 'ptmp':
        # Redirigir a la selección de tipo de llenado para PtMP
        return redirect(url_for('seleccion_llenado_ptmp', user_id=user_id, fila_idx=fila_idx))
    elif tipo and tipo.strip().lower() == 'site_survey':
        try:
            print('DEBUG: Entrando a bloque site_survey')
            # --- BLOQUE DE LLENADO DE SITE SURVEY ---
            import pandas as pd
            import xlwings as xw
            import os, re

            df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
            row = df_db.loc[int(fila_idx)]
            nombre_a = row.get('Nombre del sitio A', '') if 'Nombre del sitio A' in row else ''
            nombre_b = row.get('Nombre del sitio B', '') if 'Nombre del sitio B' in row else ''

            # Usar ruta relativa para que funcione en cualquier computadora
            import os
            base_dir = os.path.dirname(os.path.abspath(__file__))
            plantilla_path = os.path.join(base_dir, 'site_survey', 'EJEMPLO SS VACIO.xlsx')
            user_id_limpio = re.sub(r'[^a-zA-Z0-9_-]', '', str(user_id))
            output_path = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}.xlsx')
            
            # Eliminar archivo existente si existe
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                    print(f'DEBUG: Archivo existente eliminado: {output_path}')
                    time.sleep(1)
                except Exception as del_error:
                    print(f'DEBUG: Error al eliminar archivo existente: {del_error}')
                    # Intentar con nombre diferente
                    time_module = __import__('time')
                    output_path = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}_{int(time_module.time())}.xlsx')
                    print(f'DEBUG: Usando nuevo nombre: {output_path}')

            print('DEBUG: Intentando iniciar Excel...')
            
            # Cerrar cualquier proceso de Excel que pueda estar interfiriendo
            try:
                subprocess.run(['taskkill', '/f', '/im', 'excel.exe'], capture_output=True)
                print('DEBUG: Procesos de Excel cerrados')
                time.sleep(2)
            except:
                print('DEBUG: No se pudieron cerrar procesos de Excel')
            
            app_excel = xw.App(visible=False)
            print('DEBUG: Excel iniciado, intentando abrir plantilla...')
            
            # Verificar si el archivo está siendo usado
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    print(f'DEBUG: Intento {attempt + 1} de abrir plantilla...')
                    wb = app_excel.books.open(plantilla_path)
                    print('DEBUG: Plantilla abierta correctamente')
                    break
                except Exception as excel_error:
                    print(f'DEBUG: Error en intento {attempt + 1}: {excel_error}')
                    if attempt < max_attempts - 1:
                        print('DEBUG: Esperando 2 segundos antes del siguiente intento...')
                        time.sleep(2)
                    else:
                        print('DEBUG: Todos los intentos fallaron')
                        app_excel.quit()
                        return f'Error al abrir la plantilla después de {max_attempts} intentos: {str(excel_error)}'
            ws_caratula = wb.sheets['0. Carátula']
            ws_info_a = wb.sheets['1. Información General A']
            ws_info_b = wb.sheets['2. Información General B']
            ws_info_c = wb.sheets['3. Espacios en Torre y Piso A-B']
            ws_info_d = wb.sheets['4. Planos A']
            ws_info_e = wb.sheets['5. Planos B']
            ws_info_f = wb.sheets['6. Reporte Fotos A']
            ws_info_g = wb.sheets['7. Reporte Fotos B']

            # Llenado de los checkboxes y campos
            ws_info_a.range('B63').value = 'N/A'
            
            # Función mejorada para checkboxes
            def set_checkbox(worksheet, cell, condition, debug_name=""):
                """Función para establecer checkboxes con debug"""
                value = bool(condition)
                worksheet.range(cell).value = value
                if debug_name:
                    print(f"DEBUG: {debug_name} - Celda {cell}: {value}")
                return value
            
            # Debug para tipo de zona
            tipo_zona_original = row.get('Tipo de Zona', '')
            tipo_zona = normaliza_texto(tipo_zona_original)
            print(f"*** DEBUG TIPO DE ZONA ***")
            print(f"   - Valor original: '{tipo_zona_original}'")
            print(f"   - Valor normalizado: '{tipo_zona}'")
            
            set_checkbox(ws_info_a, 'L21', 'urbana' in tipo_zona, "Tipo Zona - Urbana")
            set_checkbox(ws_info_a, 'P21', 'suburbana' in tipo_zona, "Tipo Zona - Suburbana")
            set_checkbox(ws_info_a, 'U21', 'rural' in tipo_zona, "Tipo Zona - Rural")
            set_checkbox(ws_info_a, 'X21', 'ejidal' in tipo_zona, "Tipo Zona - Ejidal")
            set_checkbox(ws_info_a, 'AB21', 'pueblomagico' in tipo_zona, "Tipo Zona - Pueblo Mágico")
            ws_caratula.range('A43').value = f"{nombre_a} - {nombre_b}"
            tipo_visible_original = row.get('El sitio es visible de día y de noche (libre de maleza y arboles): ', '')
            tipo_visible = normaliza_texto(tipo_visible_original)
            print(f"*** DEBUG VISIBILIDAD ***")
            print(f"   - Valor original: '{tipo_visible_original}'")
            print(f"   - Valor normalizado: '{tipo_visible}'")
            
            set_checkbox(ws_info_a, 'P22', 'si' in tipo_visible, "Visibilidad - Sí")
            set_checkbox(ws_info_a, 'S22', 'no' in tipo_visible, "Visibilidad - No")
            tipo_camino_original = row.get('Tipo de Camino', '')
            tipo_camino = normaliza_texto(tipo_camino_original)
            print(f"*** DEBUG TIPO DE CAMINO ***")
            print(f"   - Valor original: '{tipo_camino_original}'")
            print(f"   - Valor normalizado: '{tipo_camino}'")
            
            set_checkbox(ws_info_a, 'G23', 'terraceria' in tipo_camino, "Tipo Camino - Terracería")
            set_checkbox(ws_info_a, 'L23', 'pavimentado' in tipo_camino, "Tipo Camino - Pavimentado")
            set_checkbox(ws_info_a, 'Q23', 'empedrado' in tipo_camino, "Tipo Camino - Empedrado")
            set_checkbox(ws_info_a, 'V23', 'mixto' in tipo_camino, "Tipo Camino - Mixto")

            # Debug para checkbox de grúa - SITIO A
            print(f"*** DEBUG CHECKBOX GRÚA SITIO A ***")
            
            # Mostrar todos los campos disponibles para debug
            print(f"   - Todos los campos disponibles: {list(row.index)}")
            
            # Buscar el campo correcto (sin prefijo "comentario:")
            campo_grua = 'En caso de requerirse grúa, identifique si es factible el uso de la misma y que no se tenga una posible obstrucción.'
            print(f"   - Campo en row: {campo_grua in row}")
            
            # Buscar variaciones del nombre del campo
            if campo_grua not in row:
                print(f"   - Campo no encontrado, buscando variaciones...")
                for campo in row.index:
                    if 'grúa' in campo.lower() and 'obstrucción' in campo.lower() and not campo.startswith('comentario:'):
                        print(f"   - Posible variación encontrada: '{campo}' = '{row[campo]}'")
                    elif 'grúa' in campo.lower() or 'grua' in campo.lower():
                        print(f"   - Campo con 'grúa' encontrado: '{campo}' = '{row[campo]}'")
            
            tipo_caso_grua_original = row.get(campo_grua, '')
            print(f"   - Valor original: '{tipo_caso_grua_original}'")
            tipo_caso_grua = normaliza_texto(tipo_caso_grua_original)
            print(f"   - Valor normalizado: '{tipo_caso_grua}'")
            
            # Verificar con diferentes variaciones de mayúsculas/minúsculas
            print(f"   - ¿Contiene 'izaje libre'? {'izaje libre' in tipo_caso_grua}")
            print(f"   - ¿Contiene 'izaje con obstaculos'? {'izaje con obstaculos' in tipo_caso_grua}")
            print(f"   - ¿Contiene 'requiere visita del especialista'? {'requiere visita del especialista' in tipo_caso_grua}")
            
            # Verificar también en el valor original
            print(f"   - ¿Contiene 'Izaje libre' en original? {'Izaje libre' in tipo_caso_grua_original}")
            print(f"   - ¿Contiene 'Izaje con obstaculos' en original? {'Izaje con obstaculos' in tipo_caso_grua_original}")
            print(f"   - ¿Contiene 'Requiere visita del especialista' en original? {'Requiere visita del especialista' in tipo_caso_grua_original}")
            
            # Usar comparación case-insensitive
            valor_b61 = 'izaje libre' in tipo_caso_grua or 'Izaje libre' in tipo_caso_grua_original
            valor_j61 = 'izaje con obstaculos' in tipo_caso_grua or 'Izaje con obstaculos' in tipo_caso_grua_original
            valor_s61 = 'requiere visita del especialista' in tipo_caso_grua or 'Requiere visita del especialista' in tipo_caso_grua_original
            
            print(f"   - Valor calculado para B61: {valor_b61}")
            print(f"   - Valor calculado para J61: {valor_j61}")
            print(f"   - Valor calculado para S61: {valor_s61}")
            
            set_checkbox(ws_info_a, 'B61', valor_b61, "Grúa - Izaje libre")
            set_checkbox(ws_info_a, 'J61', valor_j61, "Grúa - Izaje con obstáculos")
            set_checkbox(ws_info_a, 'S61', valor_s61, "Grúa - Requiere visita especialista")


            # Debug para checkbox de grúa - SITIO B
            print(f"*** DEBUG CHECKBOX GRÚA SITIO B ***")
            
            # Buscar el campo correcto (sin prefijo "comentario:")
            campo_grua2 = 'En caso de requerirse grúa, identifique si es factible el uso de la misma y que no se tenga una posible obstrucción2.'
            print(f"   - Campo en row: {campo_grua2 in row}")
            
            # Buscar variaciones del nombre del campo
            if campo_grua2 not in row:
                print(f"   - Campo no encontrado, buscando variaciones...")
                for campo in row.index:
                    if 'grúa' in campo.lower() and 'obstrucción' in campo.lower() and '2' in campo and not campo.startswith('comentario:'):
                        print(f"   - Posible variación encontrada: '{campo}' = '{row[campo]}'")
            
            tipo_caso_grua2_original = row.get(campo_grua2, '')
            print(f"   - Valor original: '{tipo_caso_grua2_original}'")
            tipo_caso_grua2 = normaliza_texto(tipo_caso_grua2_original)
            print(f"   - Valor normalizado: '{tipo_caso_grua2}'")
            
            # Verificar con diferentes variaciones de mayúsculas/minúsculas
            print(f"   - ¿Contiene 'izaje libre'? {'izaje libre' in tipo_caso_grua2}")
            print(f"   - ¿Contiene 'izaje con obstaculos'? {'izaje con obstaculos' in tipo_caso_grua2}")
            print(f"   - ¿Contiene 'requiere visita del especialista'? {'requiere visita del especialista' in tipo_caso_grua2}")
            
            # Usar comparación case-insensitive
            ws_info_b.range('B61').value = 'izaje libre' in tipo_caso_grua2 or 'Izaje libre' in tipo_caso_grua2_original
            ws_info_b.range('J61').value = 'izaje con obstaculos' in tipo_caso_grua2 or 'Izaje con obstaculos' in tipo_caso_grua2_original
            ws_info_b.range('S61').value = 'requiere visita del especialista' in tipo_caso_grua2 or 'Requiere visita del especialista' in tipo_caso_grua2_original
            
            # Verificar que se escribió correctamente
            print(f"   - Valor en B61: '{ws_info_b.range('B61').value}'")
            print(f"   - Valor en J61: '{ws_info_b.range('J61').value}'")
            print(f"   - Valor en S61: '{ws_info_b.range('S61').value}'")

            tipo_zona_original = row.get('Tipo de Zona 2', '')
            tipo_zona = normaliza_texto(tipo_zona_original)
            ws_info_b.range('L21').value = 'urbana' in tipo_zona
            ws_info_b.range('P21').value = 'suburbana' in tipo_zona
            ws_info_b.range('U21').value = 'rural' in tipo_zona
            ws_info_b.range('X21').value = 'ejidal' in tipo_zona
            ws_info_b.range('AB21').value = 'pueblomagico' in tipo_zona
            tipo_visible_original = row.get('El sitio es visible de día y de noche (libre de maleza y arboles): 2', '')
            tipo_visible = normaliza_texto(tipo_visible_original)
            ws_info_b.range('P22').value =  'si' in tipo_visible
            ws_info_b.range('S22').value = 'no' in tipo_visible
            tipo_camino_original = row.get(' Tipo de Camino 2 ', '')
            tipo_camino = normaliza_texto(tipo_camino_original)
            ws_info_b.range('G23').value = 'terraceria' in tipo_camino
            ws_info_b.range('L23').value = 'pavimentado' in tipo_camino
            ws_info_b.range('Q23').value =  'empedrado' in tipo_camino
            ws_info_b.range('V23').value =  'mixto' in tipo_camino




            tipo_Propietario_Administrador_original = row.get('Propietario_Administrador', '')
            tipo_Propietario_Administrador = normaliza_texto(tipo_Propietario_Administrador_original)
            ws_info_a.range('K34').value = 'telesite'in tipo_Propietario_Administrador
            ws_info_a.range('P34').value = 'ctwr' in tipo_Propietario_Administrador
            ws_info_a.range('V34').value = 'mtp' in tipo_Propietario_Administrador
            ws_info_a.range('Z34').value = 'intelesites' in tipo_Propietario_Administrador
            ws_info_a.range('AE34').value = 'even' in tipo_Propietario_Administrador
            ws_info_a.range('A35').value = 'atc' in tipo_Propietario_Administrador
            ws_info_a.range('F35').value = 'temm' in tipo_Propietario_Administrador
            ws_info_a.range('K35').value = 'renta tower' in tipo_Propietario_Administrador
            ws_info_a.range('P35').value = 'torrecom' in tipo_Propietario_Administrador
            ws_info_a.range('V35').value = 'uniti' in tipo_Propietario_Administrador
            ws_info_a.range('A36').value = 'tower one' in tipo_Propietario_Administrador
            ws_info_a.range('F36').value = 'iimt' in tipo_Propietario_Administrador
            ws_info_a.range('K36').value = 'servicom' in tipo_Propietario_Administrador
            ws_info_a.range('A37').value = 'canadian tower' in tipo_Propietario_Administrador
            ws_info_a.range('F37').value = 'mx tower' in tipo_Propietario_Administrador
            ws_info_a.range('K37').value = 'cfe' in tipo_Propietario_Administrador

            tipo_Propietario_Administrador_original = row.get('Propietario_Administrador B', '')
            tipo_Propietario_Administrador = normaliza_texto(tipo_Propietario_Administrador_original)
            ws_info_b.range('K34').value = 'telesite'in tipo_Propietario_Administrador
            ws_info_b.range('P34').value = 'ctwr' in tipo_Propietario_Administrador
            ws_info_b.range('V34').value = 'mtp' in tipo_Propietario_Administrador
            ws_info_b.range('Z34').value = 'intelesites' in tipo_Propietario_Administrador
            ws_info_b.range('AE34').value = 'even' in tipo_Propietario_Administrador
            ws_info_b.range('A35').value = 'atc' in tipo_Propietario_Administrador
            ws_info_b.range('F35').value = 'temm' in tipo_Propietario_Administrador
            ws_info_b.range('K35').value = 'renta tower' in tipo_Propietario_Administrador
            ws_info_b.range('P35').value = 'torrecom' in tipo_Propietario_Administrador
            ws_info_b.range('V35').value = 'uniti' in tipo_Propietario_Administrador
            ws_info_b.range('A36').value = 'tower one' in tipo_Propietario_Administrador
            ws_info_b.range('F36').value = 'iimt' in tipo_Propietario_Administrador
            ws_info_b.range('K36').value = 'servicom' in tipo_Propietario_Administrador
            ws_info_b.range('A37').value = 'canadian tower' in tipo_Propietario_Administrador
            ws_info_b.range('F37').value = 'mx tower' in tipo_Propietario_Administrador
            ws_info_b.range('K37').value = 'cfe' in tipo_Propietario_Administrador

            tipo_tipositio_original = normaliza_texto(row.get('Tipo de sitio', ''))
            tipo_tipositio = normaliza_texto(tipo_tipositio_original)
            print(f"Tipo de sitio original: '{tipo_tipositio_original}'")
            print(f"Tipo de sitio normalizado: '{tipo_tipositio}'")
            print(f"¿Contiene 'terrenogreenfield'? {'terrenogreenfield' in tipo_tipositio}")
            print(f"¿Contiene 'sobresuelorawland'? {'sobresuelorawland' in tipo_tipositio}")
            ws_info_a.range('D39').value = 'terrenogreenfield' in tipo_tipositio
            ws_info_a.range('M39').value = 'sobresuelorawland' in tipo_tipositio
            ws_info_a.range('U39').value = 'sobreazotea' in tipo_tipositio

            tipo_tipositio_original = normaliza_texto(row.get('Tipo de sitio B', ''))
            tipo_tipositio = normaliza_texto(tipo_tipositio_original)
            print(f"Tipo de sitio original: '{tipo_tipositio_original}'")
            print(f"Tipo de sitio normalizado: '{tipo_tipositio}'")
            print(f"¿Contiene 'terrenogreenfield'? {'terrenogreenfield' in tipo_tipositio}")
            print(f"¿Contiene 'sobresuelorawland'? {'sobresuelorawland' in tipo_tipositio}")
            ws_info_b.range('D39').value = 'terrenogreenfield' in tipo_tipositio
            ws_info_b.range('M39').value = 'sobresuelorawland' in tipo_tipositio
            ws_info_b.range('U39').value = 'sobreazotea' in tipo_tipositio


            tipo_riesgo_original = normaliza_texto(row.get('Riesgo', ''))
            tipo_riesgo = normaliza_texto(tipo_riesgo_original)
            ws_info_a.range('Y40').value = 'delitocomunroboatranseuntes' in tipo_riesgo
            ws_info_a.range('P41').value = 'inconformidadvecinalconbloqueo' in tipo_riesgo
            ws_info_a.range('AA41').value = 'delincuenciaorganizada' in tipo_riesgo

            tipo_riesgo_original = normaliza_texto(row.get('Riesgo B', ''))
            tipo_riesgo = normaliza_texto(tipo_riesgo_original)
            ws_info_b.range('Y40').value = 'delitocomunroboatranseuntes' in tipo_riesgo
            ws_info_b.range('P41').value = 'inconformidadvecinalconbloqueo' in tipo_riesgo
            ws_info_b.range('AA41').value = 'delincuenciaorganizada' in tipo_riesgo

            tipo_considera_accesible_original = normaliza_texto(row.get('Considera accesible el sitio de día y de noche?', ''))
            tipo_considera_accesible = normaliza_texto(tipo_considera_accesible_original)
            ws_info_a.range('S43').value = 'solodedia' in tipo_considera_accesible
            ws_info_a.range('W43').value = 'solodenoche' in tipo_considera_accesible
            ws_info_a.range('AB43').value = 'sinproblemadehora' in tipo_considera_accesible

            tipo_considera_accesible_original = normaliza_texto(row.get('Considera accesible el sitio de día y de noche? B', ''))
            tipo_considera_accesible = normaliza_texto(tipo_considera_accesible_original)
            ws_info_b.range('S43').value = 'solodedia' in tipo_considera_accesible
            ws_info_b.range('W43').value = 'solodenoche' in tipo_considera_accesible
            ws_info_b.range('AB43').value = 'sinproblemadehora' in tipo_considera_accesible

            tipo_zonasegura_original = (row.get('El sitio se encuentra construido en zona segura (De NO derrumbes):', ''))
            tipo_zonasegura = normaliza_texto(tipo_zonasegura_original)
            ws_info_a.range('S44').value = 'si' in tipo_zonasegura
            ws_info_a.range('W44').value = 'no' in tipo_zonasegura

            tipo_zonasegura_original = (row.get('El sitio se encuentra construido en zona segura (De NO derrumbes) B:', ''))
            tipo_zonasegura = normaliza_texto(tipo_zonasegura_original)
            ws_info_b.range('S44').value = 'si' in tipo_zonasegura
            ws_info_b.range('W44').value = 'no' in tipo_zonasegura

            tipo_horariocontrolado_original = normaliza_texto(row.get('Horario Controlado', ''))
            tipo_horariocontrolado = normaliza_texto(tipo_horariocontrolado_original)
            ws_info_a.range('B50').value = 'si' == True
            ws_info_a.range('B50').value = 'no' == False

            tipo_horariocontrolado_original = normaliza_texto(row.get('Horario Controlado B', ''))
            tipo_horariocontrolado = normaliza_texto(tipo_horariocontrolado_original)
            ws_info_b.range('B50').value = 'si' == True
            ws_info_b.range('B50').value = 'no' == False
     

            # Acceso al Personal
            acceso_personal_original = normaliza_texto(row.get('TIPO DE ACCESO A SITIO', ''))
            acceso_personal = normaliza_texto(acceso_personal_original)

            ws_info_a.range('B47').value = 'llave' in acceso_personal
            ws_info_a.range('H47').value = 'permiso/memorandum' in acceso_personal
            ws_info_a.range('R47').value = 'candadodecombinacion' in acceso_personal
            ws_info_a.range('B48').value = 'tarjetaelectronica' in acceso_personal
            ws_info_a.range('J48').value = 'otro' in acceso_personal


            # Campos vinculados
            if 'candadodecombinacion' in acceso_personal:
                ws_info_a.range('AF47').value = row.get('Candado de Combinación', '')
                ws_info_a.range('Q48').value = 'N/A'
            elif ('llave' in acceso_personal or 'permiso/memorandum' in acceso_personal or 'tarjetaelectronica' in acceso_personal):
                ws_info_a.range('Q49').value = row.get('Dónde recoger llave/permiso/tarjeta', '')

            tipo_formaingresar_original = normaliza_texto(row.get('Forma de ingresar el equipo al sitio es con:', ''))
            tipo_formaingresar = normaliza_texto(tipo_formaingresar_original)
            ws_info_a.range('U55').value = 'maniobra' in tipo_formaingresar
            ws_info_a.range('AA55').value = 'izajecongarrucha' in tipo_formaingresar
            ws_info_a.range('AG55').value = 'izajecongrua' in tipo_formaingresar

            tipo_formaingresar_original = normaliza_texto(row.get('Forma de ingresar el equipo al sitio es con: B', ''))
            tipo_formaingresar = normaliza_texto(tipo_formaingresar_original)
            ws_info_b.range('U55').value = 'maniobra' in tipo_formaingresar
            ws_info_b.range('AA55').value = 'izajecongarrucha' in tipo_formaingresar
            ws_info_b.range('AG55').value = 'izajecongrua' in tipo_formaingresar


            tipo_requerir_grua_original = normaliza_texto(row.get('Para instalación de grúa, considera necesario que se requiera tramitar permiso con las autoridades locales:', ''))
            tipo_requerir_grua = normaliza_texto(tipo_requerir_grua_original)
            print(f"tipo_requerir_grua original: '{tipo_requerir_grua_original}'")
            print(f"tipo_requerir_grua normalizado: '{tipo_requerir_grua}'")
            ws_info_a.range('AB67').value = 'requieregrua' in tipo_requerir_grua
            ws_info_a.range('AG67').value = 'noaplicagrua' in tipo_requerir_grua



            tipo_requerir_grua_original = normaliza_texto(row.get('Para instalación de grúa, considera necesario que se requiera tramitar permiso con las autoridades locales: B', ''))
            tipo_requerir_grua = normaliza_texto(tipo_requerir_grua_original)
            print(f"tipo_requerir_grua original: '{tipo_requerir_grua_original}'")
            print(f"tipo_requerir_grua normalizado: '{tipo_requerir_grua}'")
            ws_info_b.range('AB67').value = 'requieregrua' in tipo_requerir_grua
            ws_info_b.range('AG67').value = 'noaplicagrua' in tipo_requerir_grua

            acceso_personal_original = normaliza_texto(row.get('TIPO DE ACCESO A SITIO B', ''))
            acceso_personal = normaliza_texto(acceso_personal_original)

            ws_info_b.range('B47').value = 'llave' in acceso_personal
            ws_info_b.range('H47').value = 'permiso/memorandum' in acceso_personal
            ws_info_b.range('R47').value = 'candadodecombinacion' in acceso_personal
            ws_info_b.range('B48').value = 'tarjetaelectronica' in acceso_personal
            ws_info_b.range('J48').value = 'otro' in acceso_personal


                    # Campos vinculados
            if 'candadodecombinacion' in acceso_personal:
                ws_info_b.range('AF47').value = row.get('Candado de Combinación B', '')
                ws_info_b.range('Q48').value = 'N/A'
            elif ('llave' in acceso_personal or 'permiso/memorandum' in acceso_personal or 'tarjetaelectronica' in acceso_personal):
                ws_info_b.range('Q49').value = row.get('Dónde recoger llave/permiso/tarjeta B', '')




            tipo_requiere_grua_original = normaliza_texto(row.get('Requiere Grua (Si / No)', ''))
            tipo_requiere_grua = normaliza_texto(tipo_requiere_grua_original)
            ws_info_a.range('AC66').value = 'si' in tipo_requiere_grua
            ws_info_a.range('AF66').value = 'no' in tipo_requiere_grua

            tipo_requiere_grua_original = normaliza_texto(row.get('Requiere Grua (Si / No) B', ''))
            tipo_requiere_grua = normaliza_texto(tipo_requiere_grua_original)
            ws_info_b.range('AC66').value = 'si' in tipo_requiere_grua
            ws_info_b.range('AF66').value = 'no' in tipo_requiere_grua

            tipo_llegar_original = normaliza_texto(row.get('Para la llegada al sitio con el equipo a instalar, se requiere de:', ''))
            # Separa por coma y normaliza solo el primer valor
            primer_valor = tipo_llegar_original.split(',')[0].strip() if ',' in tipo_llegar_original else tipo_llegar_original.strip()
            primer_tipo = normaliza_texto(primer_valor)
            print(f"primer_valor: '{primer_valor}'")
            print(f"primer_tipo normalizado: '{primer_tipo}'")
            ws_info_a.range('B72').value = (primer_tipo == 'pickup')
            ws_info_a.range('G72').value = (primer_tipo == 'pickup4x4')
            ws_info_a.range('M72').value = (primer_tipo == 'animalesdecarga')


            tipo_llegar_original = normaliza_texto(row.get('Para la llegada al sitio con el equipo a instalar, se requiere de:', ''))
            # Separa por coma y normaliza solo el primer valor
            primer_valor = tipo_llegar_original.split(',')[0].strip() if ',' in tipo_llegar_original else tipo_llegar_original.strip()
            primer_tipo = normaliza_texto(primer_valor)
            print(f"primer_valor: '{primer_valor}'")
            print(f"primer_tipo normalizado: '{primer_tipo}'")
            ws_info_b.range('B72').value = (primer_tipo == 'pickup')
            ws_info_b.range('G72').value = (primer_tipo == 'pickup4x4')
            ws_info_b.range('M72').value = (primer_tipo == 'animalesdecarga')


            policia_original = normaliza_texto(row.get('Existe cerca del sitio alguna comandancia de policía o del ejercito?', '').strip().lower())
            ws_info_a.range('AB105').value = 'si' in policia_original
            ws_info_a.range('AE105').value = 'no' in policia_original

            if 'si' in policia_original:
                ws_info_a.range('U106').value = row.get('Si la respuesta anterior es si, Indique a que distancia  policia',) 
                ws_info_a.range('U107').value = row.get('Se cuenta con algún número de teléfono?, indíquelo', 'N/A') or 'N/A'
            else:
                ws_info_a.range('U106').value = 'N/A'
                ws_info_a.range('U107').value = 'N/A'

    # --- Cruz Roja, Hospital, asistencia médica ---
            cruzroja_original = normaliza_texto(row.get('Existe Cruz Roja, Hospital u otro tipo de asistencia medica cerca del sitio.', '').strip().lower())
            ws_info_a.range('AB108').value = 'si' in cruzroja_original
            ws_info_a.range('AE108').value = 'no' in cruzroja_original

            if 'si' in cruzroja_original:
                ws_info_a.range('U109').value = row.get('Si la respuesta anterior es si, Indique a que distancia cruz ', )  
                ws_info_a.range('U110').value = row.get('Se cuenta con algún numero de teléfono?, indíquelo: cruz', )
            else:
                ws_info_a.range('U109').value = 'N/A'
                ws_info_a.range('U110').value = 'N/A'

          # --- Mapa Nacional de Riesgos ---
            riesgo = normaliza_texto(row.get('Según el Mapa Nacional de Riesgos, indique en que zona se ubica el sitio:', '').strip().lower())
            ws_info_a.range('Y111').value = 'bajo' in riesgo
            ws_info_a.range('AB111').value = 'medio' in riesgo
            ws_info_a.range('AE111').value = 'alto' in riesgo



            policia_original = normaliza_texto(row.get('Existe cerca del sitio alguna comandancia de policía o del ejercito? B', ''))
            ws_info_b.range('AB105').value = 'si' in policia_original
            ws_info_b.range('AE105').value = 'no' in policia_original

            if 'si' in policia_original:
                ws_info_b.range('U106').value = row.get('Si la respuesta anterior es si, Indique a que distancia  policia B',) 
                ws_info_b.range('U107').value = row.get('Se cuenta con algún numero de teléfono?, indíquelo: B',) 
            else:
                ws_info_b.range('U106').value = 'N/A'
                ws_info_b.range('U107').value = 'N/A'

            # --- Cruz Roja, Hospital, asistencia médica ---
            cruzroja_original = normaliza_texto(row.get('Existe Cruz Roja, Hospital u otro tipo de asistencia medica cerca del sitio. B', ''))
            ws_info_b.range('AB108').value = 'si' in cruzroja_original
            ws_info_b.range('AE108').value = 'no' in cruzroja_original

            if 'si' in cruzroja_original:
                ws_info_b.range('U109').value = row.get('Si la respuesta anterior es si, Indique a que distancia cruz B', )  
                ws_info_b.range('U110').value = row.get('Se cuenta con algún numero de teléfono?, indíquelo: cruz B', )
            else:
                ws_info_b.range('U109').value = 'N/A'
                ws_info_b.range('U110').value = 'N/A'

              # --- Mapa Nacional de Riesgos ---
            riesgo = normaliza_texto(row.get('Según el Mapa Nacional de Riesgos, indique en que zona se ubica el sitio: B', ''))
            ws_info_b.range('Y111').value = 'bajo' in riesgo
            ws_info_b.range('AB111').value = 'medio' in riesgo
            ws_info_b.range('AE111').value = 'alto' in riesgo





            tipo_torre = normaliza_texto(row.get('Tipo de Torre', ''))
            ws_info_c.range('G8').value = tipo_torre == 'autosoportada'
            ws_info_c.range('O8').value = tipo_torre == 'arriostrada'
            ws_info_c.range('V8').value = tipo_torre == 'Monopolo'
            ws_info_c.range('AB8').value = tipo_torre == 'Minipolo'
            ws_info_c.range('AG8').value = tipo_torre == 'otro'

            espacio_disponible = normaliza_texto(row.get('¿Espacio disponible de conexión?', ''))
            ws_info_c.range('AH25').value = espacio_disponible == 'si'
            ws_info_c.range('AK25').value = espacio_disponible == 'no'
            ws_info_c.range('U14').value = espacio_disponible == 'si'
            ws_info_c.range('Y14').value = espacio_disponible == 'no'


            existe_break = normaliza_texto(row.get('¿Existe algun breaker existente en sitio? ', ''))
            ws_info_c.range('Y21').value = existe_break == 'si'
            ws_info_c.range('AB21').value = existe_break == 'no'

        
            alimentacion_compatible = normaliza_texto(row.get('Alimentacion compatible con el equipamiento ', ''))
            print(f"*** DEBUG ALIMENTACIÓN COMPATIBLE ***")
            print(f"   - Valor: '{alimentacion_compatible}'")
            set_checkbox(ws_info_c, 'Y25', alimentacion_compatible == 'si', "Alimentación Compatible - Sí")
            set_checkbox(ws_info_c, 'AB25', alimentacion_compatible == 'no', "Alimentación Compatible - No")

            sistema_electrico_original = row.get('SISTEMA ELECTRICO', '')
            sistema_electrico = normaliza_texto(sistema_electrico_original)
            print(f"*** DEBUG SISTEMA ELÉCTRICO ***")
            print(f"   - Valor original: '{sistema_electrico_original}'")
            print(f"   - Valor normalizado: '{sistema_electrico}'")
            set_checkbox(ws_info_c, 'AG21', 'monofasica' in sistema_electrico, "Sistema Eléctrico - Monofásica")
            set_checkbox(ws_info_c, 'AL21', 'bifasica' in sistema_electrico, "Sistema Eléctrico - Bifásica")



            cara_propuesta = normaliza_texto(row.get('Cara de preparación para cableado vertical en torre', ''))
            print(f"*** DEBUG CARA PROPUESTA ***")
            print(f"   - Valor: '{cara_propuesta}'")
            set_checkbox(ws_info_c, 'Y16', cara_propuesta == 'a', "Cara Propuesta - A")
            set_checkbox(ws_info_c, 'AC16', cara_propuesta == 'b', "Cara Propuesta - B")
            set_checkbox(ws_info_c, 'AE16', cara_propuesta == 'c', "Cara Propuesta - C")
            set_checkbox(ws_info_c, 'AN16', cara_propuesta == 'd', "Cara Propuesta - D")

            barra_tierra = normaliza_texto(row.get('Barra de Tierra', ''))
            print(f"*** DEBUG BARRA DE TIERRA ***")
            print(f"   - Valor: '{barra_tierra}'")
            set_checkbox(ws_info_c, 'O27', barra_tierra == 'si', "Barra de Tierra - Sí")
            set_checkbox(ws_info_c, 'R27', barra_tierra == 'no', "Barra de Tierra - No")

            tipo_solucion = normaliza_texto(row.get('Tipo de Solucion', ''))
            print(f"*** DEBUG TIPO DE SOLUCIÓN ***")
            print(f"   - Valor: '{tipo_solucion}'")
            set_checkbox(ws_info_c, 'O29', tipo_solucion == 'piso', "Tipo Solución - Piso")
            set_checkbox(ws_info_c, 'R29', tipo_solucion == 'torre', "Tipo Solución - Torre")



        
            tipo_torre2 = normaliza_texto(row.get('Tipo de Torre2', ''))
            ws_info_c.range('G32').value = tipo_torre2 == 'autosoportada'
            ws_info_c.range('O32').value = tipo_torre2 == 'arriostrada'
            ws_info_c.range('V32').value = tipo_torre2 == 'monopolo'
            ws_info_c.range('AB32').value = tipo_torre2 == 'minipolo'
            ws_info_c.range('AG32').value = tipo_torre2 == 'otro'

            espacio_disponible2 = normaliza_texto(row.get('¿Espacio disponible de conexión?2', ''))
            ws_info_c.range('U38').value = espacio_disponible2 == 'si'
            ws_info_c.range('Y38').value = espacio_disponible2 == 'no'

            cara_preparacion2 = normaliza_texto(row.get('Cara de preparación para cableado vertical en torre 2', ''))
            ws_info_c.range('Y40').value = cara_preparacion2 == 'a'
            ws_info_c.range('AD40').value = cara_preparacion2 == 'b'
            ws_info_c.range('AI40').value = cara_preparacion2 == 'c'
            ws_info_c.range('AN40').value = cara_preparacion2 == 'd'
        
            existe_tierra2 = normaliza_texto(row.get('Existe Barra de Tierras 2', ''))
            ws_info_c.range('O51').value = existe_tierra2 == 'si'
            ws_info_c.range('R51').value = existe_tierra2 == 'no'

            tipo_solucion2 = normaliza_texto(row.get('Tipo de solucion 2', ''))
            ws_info_c.range('O53').value = tipo_solucion2 == 'piso'
            ws_info_c.range('R53').value = tipo_solucion2 == 'torre'
        
            existe_break2 = normaliza_texto(row.get('Existe algun breaker existente en sitio 2 ', ''))
            ws_info_c.range('Y45').value = existe_break2 == 'si'
            ws_info_c.range('AB45').value = existe_break2 == 'no'

            alimenacion_existente2= normaliza_texto(row.get('SISTEMA ELECTRICO 2', ''))
            ws_info_c.range('AG45').value = 'monofasica' in alimenacion_existente2
            ws_info_c.range('AL45').value = 'bifasica' in alimenacion_existente2
        
            alimenacion_compatible2= normaliza_texto(row.get('Alimentacion compatible con el equipamiento 2', ''))
            ws_info_c.range('Y49').value = alimenacion_compatible2 == 'si'
            ws_info_c.range('AB49').value = alimenacion_compatible2 == 'no'

            espacio_conexion2= normaliza_texto(row.get('¿Espacio disponible de conexión? 2', ''))
            ws_info_c.range('AH49').value = espacio_conexion2 == 'si'
            ws_info_c.range('AK49').value = espacio_conexion2 == 'no'

            linea_vista = normaliza_texto(row.get('Linea de vista ', ''))
            motivo = normaliza_texto(row.get('Motivo ', ''))

            ws_info_c.range('R56').value = (linea_vista == 'si')
            ws_info_c.range('V56').value = (linea_vista == 'no')
            ws_info_c.range('Q57').value = False
            ws_info_c.range('V57').value = False
            ws_info_c.range('AC57').value = False
            ws_info_c.range('AI57').value = False
            ws_info_c.range('C58').value = False

            if linea_vista == 'no':
               if motivo == 'arboles':
                  ws_info_c.range('Q57').value = True
               elif motivo == 'espectacular':
                  ws_info_c.range('V57').value = True
               elif motivo == 'edificio':
                  ws_info_c.range('AC57').value = True
               elif motivo == 'montaña':
                  ws_info_c.range('AI57').value = True
               elif motivo == 'n/a':
                  ws_info_c.range('C58').value = True

            
             # Llenado automático de celdas usando el diccionario de mapeo
            campos_a_celdas = {
                'Fecha Inicio Site Survey': 'G8',
                'Fecha final Site Survey': 'AF8',
                'NOMBRE DEL SITIO': 'J9', 
                'PROPIETARIO': 'M10',
                'ID': 'AF9',
                'ESTADO ':'AC15',
                'Calle': 'D14',
                'Colonia': 'D15',
                'Municipio': 'E16',
                'C.P': 'AC14',
                'Referencias':'J17',
                'Nombre de contacto en sitio': 'H19',
                'Telefono': 'AB19',
                'LATITUD (TORRE)': 'K30',
                'LONGITUD (TORRE)': 'AA30',
                'LATITUD (FACHADA)': 'K27',
                'LONGITUD (FACHADA)': 'AA27',
                'Altitud (msnm)': 'M31',
                'Horario de solicitud de accesos': 'Q50',
                'Contacto solicitud de accesos': 'B52',
                'Como o donde obtener permisos/llave/tarjeta': 'O49',
                'Comentario:Forma de ingresar el equipo al sitio es con:': 'B57',
                'comentario:En caso de requerirse grúa, identifique si es factible el uso de la misma y que no se tenga una posible obstrucción.': 'B63',
                'Si la respuesta anterior es si, Indique a que distancia ':'AD72'
            }
            for campo, celdas in campos_a_celdas.items():
                valor = normaliza_na(row.get(campo, ""))
                if campo == 'comentario:En caso de requerirse grúa, identifique si es factible el uso de la misma y que no se tenga una posible obstrucción.':
                    print(f"Valor para {campo}: '{valor}' (celda {celdas})")
                if isinstance(celdas, list):
                    for celda in celdas:
                        ws_info_a.range(celda).value = valor
                else:
                    ws_info_a.range(celdas).value = valor

            campos_b_celdas = {
                'Fecha Inicio Site Survey B': 'G8',
                'Fecha final Site Survey B': 'AF8',
                'Nombre del sitio 2': 'J9', 
                'PROPIETARIO 2': 'M10',
                'ID 2': 'AF9',
                'ESTADO 2 ':'AC15',
                'Calle 2': 'D14',
                'Colonia 2': 'D15',
                'Municipio 2': 'E16',
                'C.P 2': 'AC14',
                'Referencias 2':'J17',
                'Nombre de contacto en sitio 2': 'H19',
                'Telefono 2': 'AB19',
                'LATITUD (TORRE) 2': 'K30',
                'LONGITUD (TORRE) 2': 'AA30',
                'LATITUD (FACHADA) 2': 'K27',
                'LONGITUD (FACHADA) 2': 'AA27',
                'Altitud (msnm) 2': 'M31',
                'Horario de solicitud de accesos B': 'Q50',
                'Contacto solicitud de accesos B': 'B52',
                'Como o donde obtener permisos/llave/tarjeta B': 'O49',
                'Comentario:Forma de ingresar el equipo al sitio es con: B': 'B57',
                'comentario:En caso de requerirse grúa, identifique si es factible el uso de la misma y que no se tenga una posible obstrucción.': 'B63',
                'Si la respuesta anterior es si, Indique a que distancia B':'AD72'
            }
            for campo, celdas in campos_b_celdas.items():
                valor = normaliza_na(row.get(campo, ""))
                if campo == 'comentario:En caso de requerirse grúa, identifique si es factible el uso de la misma y que no se tenga una posible obstrucción.':
                   print(f"Valor para {campo}: '{valor}' (celda {celdas})")
                if isinstance(celdas, list):
                    for celda in celdas:
                        ws_info_b.range(celda).value = valor
                else:
                    ws_info_b.range(celdas).value = valor    

            campos_c_celdas = {
            'NOMBRE DEL SITIO': 'G7',
            'Diametro de pierna superior':'K9',
            'Diametro de pierna Inferior':'U9',
            'NCRA RB':'AC9',
            'Franja2RB':'AM9',
            'Altura de la Torre':'K10',
            'Dado':'U10',
            'Altura Edificio1':'AE10',
            'Nivel inferior de franja disponible': 'T11',
            'Nivel superior de franja disponible': 'AK11',
            'Altura de MW conforme a topologia': 'B14',
            'Azimut RB ': 'M14',
            'Propuesta de altura de antena de MW1': 'AB14',
            'Propuesta de altura de antena de MW (SD)1': 'AJ14',
            'Altura de soporte para OMB propuesto': 'O19',
            'Longitud del cable de tierra nuevo OMB': 'O20',
            'Longitud del cable de tierra ODU': 'O21',
            'Longitud de cable IF': 'O22',
            'Tipo de soporte para antena MW propuesto': 'O23',
            'Longitud de cable ACDB-Nuevo OMB': 'O24',
            'Longitud de cable RTN - Router':'O25',
            'Longitud de cable RTN - BBU SITE 1': 'O26',
            'MEDICION DE BARRA DE TIERRA (Ohms)':'O28',
            'Nombre del sitio 2': 'G31',
            'Diámetro de Pierna superio2':'K33',
            'Diámetro de Pierna inferior2':'U33',
            ' NCRA2 ':'AC33',
            'Franja2-2':'AM33',
            'Altura torre 2': 'K34',
            'DADO 2':'U34',
            'Altura edificio 2':'AE34',
            'Nivel inferior de franja disponible 2': 'T35',
            'Nivel superior de franja disponible 2': 'AK35',
            'Altura de MW conforme a topologia 2': 'B38',
            'Azimut 2': 'M38',
            'Propuesta de altura de antena de MW2': 'AB38',
            'Propuesta de altura de antena de MW (SD)2':'AJ38',
            'Altura de soporte para OMB propuesto2':'O43',
            'Longitud del cable de tierra nuevo OMB 2': 'O44',
            'Longitud del cable de tierra ODU 2': 'O45',
            'Longitud de cable IF 2': 'O46',
            'Tipo de soporte para antena MW propuesto 2': 'O47',
            'Longitud de cable ACDB-Nuevo OMB 2': 'O48',
            'Longitud de cable RTN - Router 2': 'O49',
            'Longitud de cable RTN - BBU 2': 'O50',
            'Medición del Sistema de Tierras 2': 'O52',
         


            }
            for campo, celdas in campos_c_celdas.items():
                valor = normaliza_na(row.get(campo, ""))
                print(f"ws_info_c: Escribiendo en {celdas} el valor '{valor}' para campo '{campo}'")
                if isinstance(celdas, list):
                    for celda in celdas:
                        ws_info_c.range(celda).value = valor
                else:
                    ws_info_c.range(celdas).value = valor   

            wb.save(output_path)
            wb.close()
            app_excel.quit()

            print('DEBUG: Excel cerrado, antes del redirect')
            return redirect(url_for('site_survey', user_id=user_id, fila_idx=fila_idx))
        except Exception as e:
            print(f"ERROR en llenado site_survey: {e}")
            return f"Error en llenado site_survey: {e}"
    elif tipo == 'reporte_planeacion':
        print('DEBUG: Entrando a bloque reporte_planeacion')
        return redirect(url_for('reporte_planeacion', user_id=user_id, fila_idx=fila_idx))
    elif tipo == 'diseno_solucion':
        print('DEBUG: Entrando a bloque diseno_solucion')
        return redirect(url_for('formulario_archivos', user_id=user_id, fila_idx=fila_idx))
    elif tipo == 'ptmp_site_survey':
        try:
            print('DEBUG: Entrando a bloque ptmp_site_survey')
            print(f'DEBUG: user_id = {user_id}, fila_idx = {fila_idx}')
            # --- BLOQUE DE LLENADO DE SITE SURVEY PtMP ---
            import pandas as pd
            import xlwings as xw
            import os, re

            print('DEBUG: Leyendo CSV de Google Sheets...')
            df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
            print(f'DEBUG: CSV leído, fila_idx = {fila_idx}')
            row = df_db.loc[int(fila_idx)]
            print('DEBUG: Fila obtenida correctamente')
            print(f'DEBUG: Campos disponibles en la fila: {list(row.index)}')
            
            # Buscar campos relacionados con altura de edificio
            campos_altura = [campo for campo in row.index if 'altura' in campo.lower() and 'edificio' in campo.lower()]
            print(f'DEBUG: Campos de altura de edificio encontrados: {campos_altura}')
            
            # Verificar el valor exacto del campo Altura Edificio1
            if 'Altura Edificio1' in row:
                valor_altura_edificio1 = row['Altura Edificio1']
                print(f'DEBUG: Valor directo de Altura Edificio1: "{valor_altura_edificio1}" (tipo: {type(valor_altura_edificio1)})')
            else:
                print(f'DEBUG: Campo "Altura Edificio1" NO encontrado en la fila')
                # Buscar variaciones del nombre
                for campo in row.index:
                    if 'altura' in campo.lower() and 'edificio' in campo.lower() and '1' in campo:
                        print(f'DEBUG: Posible variación encontrada: "{campo}" = "{row[campo]}"')
            
            nombre_a = row.get('Nombre del sitio A', '') if 'Nombre del sitio A' in row else ''
            nombre_b = row.get('Nombre del sitio B', '') if 'Nombre del sitio B' in row else ''
            print(f'DEBUG: nombre_a = {nombre_a}, nombre_b = {nombre_b}')

            # Usar ruta relativa para que funcione en cualquier computadora
            import os
            base_dir = os.path.dirname(os.path.abspath(__file__))
            plantilla_path = os.path.join(base_dir, 'ptmp_site_survey', 'EJEMPLO SS PtMP VACIO.xlsx')
            user_id_limpio = re.sub(r'[^a-zA-Z0-9_-]', '', str(user_id))
            output_path = os.path.join(base_dir, 'ptmp_site_survey', f'ss_ptmp_{user_id_limpio}.xlsx')
            print(f'DEBUG: plantilla_path = {plantilla_path}')
            print(f'DEBUG: output_path = {output_path}')
            print(f'DEBUG: ¿Existe plantilla? {os.path.exists(plantilla_path)}')

            print('DEBUG: Intentando iniciar Excel para PtMP...')
            
            # Cerrar cualquier instancia de Excel que pueda estar interfiriendo
            try:
                import subprocess
                subprocess.run(['taskkill', '/f', '/im', 'excel.exe'], capture_output=True)
                print('DEBUG: Procesos de Excel cerrados')
                time.sleep(1)
            except:
                print('DEBUG: No se pudieron cerrar procesos de Excel')
            
            app_excel = xw.App(visible=False)
            print('DEBUG: Excel iniciado, intentando abrir plantilla PtMP...')
            
            # Verificar si el archivo está siendo usado
            import time
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    print(f'DEBUG: Intento {attempt + 1} de abrir plantilla...')
                    wb = app_excel.books.open(plantilla_path)
                    print('DEBUG: Plantilla PtMP abierta correctamente')
                    break
                except Exception as excel_error:
                    print(f'DEBUG: Error en intento {attempt + 1}: {excel_error}')
                    if attempt < max_attempts - 1:
                        print('DEBUG: Esperando 2 segundos antes del siguiente intento...')
                        time.sleep(2)
                    else:
                        print('DEBUG: Todos los intentos fallaron')
                        app_excel.quit()
                        raise excel_error
            
            # Obtener las hojas de trabajo específicas para PtMP
            print('DEBUG: Obteniendo hojas de trabajo...')
            try:
                ws_caratula = wb.sheets['0. Carátula']
                print('DEBUG: Hoja Carátula obtenida')
                ws_sector = wb.sheets['1. Información General SECTOR']
                print('DEBUG: Hoja SECTOR obtenida')
                ws_cpe = wb.sheets['2. Información General CPE']
                print('DEBUG: Hoja CPE obtenida')
                ws_espacios = wb.sheets['3. Espacios en Torre y Piso A-B']
                print('DEBUG: Hoja Espacios obtenida')
                ws_planos_a = wb.sheets['4. Planos A']
                print('DEBUG: Hoja Planos A obtenida')
                ws_planos_b = wb.sheets['5. Planos B']
                print('DEBUG: Hoja Planos B obtenida')
            except Exception as sheet_error:
                print(f'DEBUG: Error al obtener hojas: {sheet_error}')
                print(f'DEBUG: Hojas disponibles: {[sheet.name for sheet in wb.sheets]}')
                wb.close()
                app_excel.quit()
                raise sheet_error
            
            # Llenar nombre en carátula
            print('DEBUG: Llenando carátula...')
            ws_caratula.range('A37').value = f"{nombre_a} - {nombre_b}"
            print('DEBUG: Carátula llenada correctamente')

            # Llenado de los checkboxes y campos para PtMP
            ws_sector.range('B63').value = 'N/A'
            
            # Tipo de Zona - SECTOR
            tipo_zona_original = row.get('Tipo de Zona', '')
            tipo_zona = normaliza_texto(tipo_zona_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('C22').value = 'urbana' in tipo_zona
            ws_sector.range('D21').value = 'suburbana' in tipo_zona
            ws_sector.range('E21').value = 'rural' in tipo_zona
            ws_sector.range('F21').value = 'ejidal' in tipo_zona
            ws_sector.range('H22').value = 'pueblomagico' in tipo_zona
            
            # Título del documento PtMP
            
            # Visibilidad del sitio - SECTOR
            tipo_visible_original = row.get('El sitio es visible de día y de noche (libre de maleza y arboles): ', '')
            tipo_visible = normaliza_texto(tipo_visible_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('E23').value = 'si' in tipo_visible
            ws_sector.range('G21').value = 'no' in tipo_visible
            
            # Tipo de Camino - SECTOR
            tipo_camino_original = row.get('Tipo de Camino', '')
            tipo_camino = normaliza_texto(tipo_camino_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('B25').value = 'terraceria' in tipo_camino
            ws_sector.range('D25').value = 'pavimentado' in tipo_camino
            ws_sector.range('F25').value = 'empedrado' in tipo_camino
            ws_sector.range('H25').value = 'mixto' in tipo_camino

            # Tipo de Zona - CPE
            tipo_zona_original = row.get('Tipo de Zona 2', '')
            tipo_zona = normaliza_texto(tipo_zona_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('C22').value = 'urbana' in tipo_zona
            ws_cpe.range('D21').value = 'suburbana' in tipo_zona
            ws_cpe.range('F21').value = 'rural' in tipo_zona
            ws_cpe.range('G21').value = 'ejidal' in tipo_zona
            ws_cpe.range('H21').value = 'pueblomagico' in tipo_zona
            
            # Visibilidad del sitio - CPE
            tipo_visible_original = row.get('El sitio es visible de día y de noche (libre de maleza y arboles): 2', '')
            tipo_visible = normaliza_texto(tipo_visible_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('E23').value = 'si' in tipo_visible
            ws_cpe.range('F20').value = 'no' in tipo_visible
            
            # Tipo de Camino - CPE
            tipo_camino_original = row.get(' Tipo de Camino 2 ', '')
            tipo_camino = normaliza_texto(tipo_camino_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('B25').value = 'terraceria' in tipo_camino
            ws_cpe.range('D25').value = 'pavimentado' in tipo_camino
            ws_cpe.range('F25').value = 'empedrado' in tipo_camino
            ws_cpe.range('H25').value = 'mixto' in tipo_camino

            # Propietario/Administrador - SECTOR
            tipo_Propietario_Administrador_original = row.get('Propietario_Administrador', '')
            tipo_Propietario_Administrador = normaliza_texto(tipo_Propietario_Administrador_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('A34').value = 'telesite' in tipo_Propietario_Administrador
            ws_sector.range('D36').value = 'ctwr' in tipo_Propietario_Administrador
            ws_sector.range('E34').value = 'mtp' in tipo_Propietario_Administrador
            ws_sector.range('G36').value = 'intelesites' in tipo_Propietario_Administrador
            ws_sector.range('H34').value = 'even' in tipo_Propietario_Administrador
            ws_sector.range('A37').value = 'atc' in tipo_Propietario_Administrador
            ws_sector.range('B38').value = 'temm' in tipo_Propietario_Administrador
            ws_sector.range('C38').value = 'renta tower' in tipo_Propietario_Administrador
            ws_sector.range('E38').value = 'torrecom' in tipo_Propietario_Administrador
            ws_sector.range('F38').value = 'uniti' in tipo_Propietario_Administrador
            ws_sector.range('A40').value = 'tower one' in tipo_Propietario_Administrador
            ws_sector.range('B40').value = 'iimt' in tipo_Propietario_Administrador
            ws_sector.range('C40').value = 'servicom' in tipo_Propietario_Administrador
            ws_sector.range('A42').value = 'canadian tower' in tipo_Propietario_Administrador
            ws_sector.range('B42').value = 'mx tower' in tipo_Propietario_Administrador
            ws_sector.range('C42').value = 'cfe' in tipo_Propietario_Administrador

            # Propietario/Administrador - CPE
            tipo_Propietario_Administrador_original = row.get('Propietario_Administrador B', '')
            tipo_Propietario_Administrador = normaliza_texto(tipo_Propietario_Administrador_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('B37').value = 'banbien' in tipo_Propietario_Administrador
            ws_cpe.range('C37').value = 'sepomex' in tipo_Propietario_Administrador
            ws_cpe.range('D37').value = 'cac cfe' in tipo_Propietario_Administrador
            ws_cpe.range('E37').value = 'telecom' in tipo_Propietario_Administrador
            ws_cpe.range('F37').value = 'ubbj' in tipo_Propietario_Administrador
            ws_cpe.range('F38').value = 'sedena' in tipo_Propietario_Administrador
            ws_cpe.range('G38').value = 'cuartel guardia nacional' in tipo_Propietario_Administrador
            ws_cpe.range('A39').value = 'unidad medico rural' in tipo_Propietario_Administrador
            ws_cpe.range('B39').value = 'hospital rural' in tipo_Propietario_Administrador
            ws_cpe.range('C39').value = 'centro de salud' in tipo_Propietario_Administrador
            ws_cpe.range('A41').value = 'unidad medico familiar' in tipo_Propietario_Administrador
            ws_cpe.range('B41').value = 'telesecundaria' in tipo_Propietario_Administrador
            # ws_cpe.range('').value = 'servicom' in tipo_Propietario_Administrador
            # ws_cpe.range('').value = 'canadian tower' in tipo_Propietario_Administrador
            # ws_cpe.range('').value = 'mx tower' in tipo_Propietario_Administrador
            # ws_cpe.range('').value = 'cfe' in tipo_Propietario_Administrador

            # Tipo de sitio - SECTOR
            tipo_tipositio_original = normaliza_texto(row.get('Tipo de sitio', ''))
            tipo_tipositio = normaliza_texto(tipo_tipositio_original)
            print(f"Tipo de sitio original: '{tipo_tipositio_original}'")
            print(f"Tipo de sitio normalizado: '{tipo_tipositio}'")
            print(f"¿Contiene 'terrenogreenfield'? {'terrenogreenfield' in tipo_tipositio}")
            print(f"¿Contiene 'sobresuelorawland'? {'sobresuelorawland' in tipo_tipositio}")
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('A45').value = 'terrenogreenfield' in tipo_tipositio
            ws_sector.range('C45').value = 'sobresuelorawland' in tipo_tipositio
            ws_sector.range('E45').value = 'sobreazotea' in tipo_tipositio

            # Tipo de sitio - CPE
            tipo_tipositio_original = normaliza_texto(row.get('Tipo de sitio B', ''))
            tipo_tipositio = normaliza_texto(tipo_tipositio_original)
            print(f"Tipo de sitio original: '{tipo_tipositio_original}'")
            print(f"Tipo de sitio normalizado: '{tipo_tipositio}'")
            print(f"¿Contiene 'terrenogreenfield'? {'terrenogreenfield' in tipo_tipositio}")
            print(f"¿Contiene 'sobresuelorawland'? {'sobresuelorawland' in tipo_tipositio}")
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('A44').value = 'terrenogreenfield' in tipo_tipositio
            ws_cpe.range('C44').value = 'sobresuelorawland' in tipo_tipositio
            ws_cpe.range('E44').value = 'sobreazotea' in tipo_tipositio

            # Riesgo - SECTOR
            tipo_riesgo_original = normaliza_texto(row.get('Riesgo', ''))
            tipo_riesgo = normaliza_texto(tipo_riesgo_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('F46').value = 'delitocomunroboatranseuntes' in tipo_riesgo
            ws_sector.range('C46').value = 'inconformidadvecinalconbloqueo' in tipo_riesgo
            ws_sector.range('G47').value = 'delincuenciaorganizada' in tipo_riesgo

            # Riesgo - CPE
            tipo_riesgo_original = normaliza_texto(row.get('Riesgo B', ''))
            tipo_riesgo = normaliza_texto(tipo_riesgo_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('F45').value = 'delitocomunroboatranseuntes' in tipo_riesgo
            ws_cpe.range('C45').value = 'inconformidadvecinalconbloqueo' in tipo_riesgo
            ws_cpe.range('G46').value = 'delincuenciaorganizada' in tipo_riesgo

            # Accesibilidad - SECTOR
            tipo_considera_accesible_original = normaliza_texto(row.get('Considera accesible el sitio de día y de noche?', ''))
            tipo_considera_accesible = normaliza_texto(tipo_considera_accesible_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('C48').value = 'solodedia' in tipo_considera_accesible
            ws_sector.range('D48').value = 'solodenoche' in tipo_considera_accesible
            ws_sector.range('F47').value = 'sinproblemadehora' in tipo_considera_accesible

            # Accesibilidad - CPE
            tipo_considera_accesible_original = normaliza_texto(row.get('Considera accesible el sitio de día y de noche? B', ''))
            tipo_considera_accesible = normaliza_texto(tipo_considera_accesible_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('C47').value = 'solodedia' in tipo_considera_accesible
            ws_cpe.range('D47').value = 'solodenoche' in tipo_considera_accesible
            ws_cpe.range('F48').value = 'sinproblemadehora' in tipo_considera_accesible

            # Zona segura - SECTOR
            tipo_zonasegura_original = (row.get('El sitio se encuentra construido en zona segura (De NO derrumbes):', ''))
            tipo_zonasegura = normaliza_texto(tipo_zonasegura_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('G50').value = 'si' in tipo_zonasegura
            ws_sector.range('H50').value = 'no' in tipo_zonasegura

            # Zona segura - CPE
            tipo_zonasegura_original = (row.get('El sitio se encuentra construido en zona segura (De NO derrumbes) B:', ''))
            tipo_zonasegura = normaliza_texto(tipo_zonasegura_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('G49').value = 'si' in tipo_zonasegura
            ws_cpe.range('H49').value = 'no' in tipo_zonasegura

            # Horario controlado - SECTOR
            tipo_horariocontrolado_original = normaliza_texto(row.get('Horario Controlado', ''))
            tipo_horariocontrolado = normaliza_texto(tipo_horariocontrolado_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_sector.range('B55').value = 'si' in tipo_horariocontrolado

            # Horario controlado - CPE
            tipo_horariocontrolado_original = normaliza_texto(row.get('Horario Controlado B', ''))
            tipo_horariocontrolado = normaliza_texto(tipo_horariocontrolado_original)
            # TODO: Agregar coordenadas correctas para PtMP
            ws_cpe.range('B54').value = 'si' in tipo_horariocontrolado




            # Acceso al Personal
            acceso_personal_original = normaliza_texto(row.get('TIPO DE ACCESO A SITIO', ''))
            acceso_personal = normaliza_texto(acceso_personal_original)

            ws_info_a.range('B47').value = 'llave' in acceso_personal
            ws_info_a.range('H47').value = 'permiso/memorandum' in acceso_personal
            ws_info_a.range('R47').value = 'candadodecombinacion' in acceso_personal
            ws_info_a.range('B48').value = 'tarjetaelectronica' in acceso_personal
            ws_info_a.range('J48').value = 'otro' in acceso_personal


            # Campos vinculados
            if 'candadodecombinacion' in acceso_personal:
                ws_info_a.range('AF47').value = row.get('Candado de Combinación', '')
                ws_info_a.range('Q48').value = 'N/A'
            elif ('llave' in acceso_personal or 'permiso/memorandum' in acceso_personal or 'tarjetaelectronica' in acceso_personal):
                ws_info_a.range('Q49').value = row.get('Dónde recoger llave/permiso/tarjeta', '')

            tipo_formaingresar_original = normaliza_texto(row.get('Forma de ingresar el equipo al sitio es con:', ''))
            tipo_formaingresar = normaliza_texto(tipo_formaingresar_original)
            ws_info_a.range('U55').value = 'maniobra' in tipo_formaingresar
            ws_info_a.range('AA55').value = 'izajecongarrucha' in tipo_formaingresar
            ws_info_a.range('AG55').value = 'izajecongrua' in tipo_formaingresar

            tipo_formaingresar_original = normaliza_texto(row.get('Forma de ingresar el equipo al sitio es con: B', ''))
            tipo_formaingresar = normaliza_texto(tipo_formaingresar_original)
            ws_info_b.range('U55').value = 'maniobra' in tipo_formaingresar
            ws_info_b.range('AA55').value = 'izajecongarrucha' in tipo_formaingresar
            ws_info_b.range('AG55').value = 'izajecongrua' in tipo_formaingresar


            tipo_requerir_grua_original = normaliza_texto(row.get('Para instalación de grúa, considera necesario que se requiera tramitar permiso con las autoridades locales:', ''))
            tipo_requerir_grua = normaliza_texto(tipo_requerir_grua_original)
            print(f"tipo_requerir_grua original: '{tipo_requerir_grua_original}'")
            print(f"tipo_requerir_grua normalizado: '{tipo_requerir_grua}'")
            ws_info_a.range('AB67').value = 'requieregrua' in tipo_requerir_grua
            ws_info_a.range('AG67').value = 'noaplicagrua' in tipo_requerir_grua



            tipo_requerir_grua_original = normaliza_texto(row.get('Para instalación de grúa, considera necesario que se requiera tramitar permiso con las autoridades locales: B', ''))
            tipo_requerir_grua = normaliza_texto(tipo_requerir_grua_original)
            print(f"tipo_requerir_grua original: '{tipo_requerir_grua_original}'")
            print(f"tipo_requerir_grua normalizado: '{tipo_requerir_grua}'")
            ws_info_b.range('AB67').value = 'requieregrua' in tipo_requerir_grua
            ws_info_b.range('AG67').value = 'noaplicagrua' in tipo_requerir_grua

            acceso_personal_original = normaliza_texto(row.get('TIPO DE ACCESO A SITIO B', ''))
            acceso_personal = normaliza_texto(acceso_personal_original)

            ws_info_a.range('B47').value = 'llave' in acceso_personal
            ws_info_a.range('H47').value = 'permiso/memorandum' in acceso_personal
            ws_info_a.range('R47').value = 'candadodecombinacion' in acceso_personal
            ws_info_a.range('B48').value = 'tarjetaelectronica' in acceso_personal
            ws_info_a.range('J48').value = 'otro' in acceso_personal


                    # Campos vinculados
            if 'candadodecombinacion' in acceso_personal:
                ws_info_b.range('AF47').value = row.get('Candado de Combinación B', '')
                ws_info_b.range('Q48').value = 'N/A'
            elif ('llave' in acceso_personal or 'permiso/memorandum' in acceso_personal or 'tarjetaelectronica' in acceso_personal):
                ws_info_b.range('Q49').value = row.get('Dónde recoger llave/permiso/tarjeta B', '')




            tipo_requiere_grua_original = normaliza_texto(row.get('Requiere Grua (Si / No)', ''))
            tipo_requiere_grua = normaliza_texto(tipo_requiere_grua_original)
            ws_info_a.range('AC66').value = 'si' in tipo_requiere_grua
            ws_info_a.range('AF66').value = 'no' in tipo_requiere_grua

            tipo_requiere_grua_original = normaliza_texto(row.get('Requiere Grua (Si / No) B', ''))
            tipo_requiere_grua = normaliza_texto(tipo_requiere_grua_original)
            ws_info_b.range('AC66').value = 'si' in tipo_requiere_grua
            ws_info_b.range('AF66').value = 'no' in tipo_requiere_grua

            tipo_llegar_original = normaliza_texto(row.get('Para la llegada al sitio con el equipo a instalar, se requiere de:', ''))
            # Separa por coma y normaliza solo el primer valor
            primer_valor = tipo_llegar_original.split(',')[0].strip() if ',' in tipo_llegar_original else tipo_llegar_original.strip()
            primer_tipo = normaliza_texto(primer_valor)
            print(f"primer_valor: '{primer_valor}'")
            print(f"primer_tipo normalizado: '{primer_tipo}'")
            ws_info_a.range('B72').value = (primer_tipo == 'pickup')
            ws_info_a.range('G72').value = (primer_tipo == 'pickup4x4')
            ws_info_a.range('M72').value = (primer_tipo == 'animalesdecarga')


            tipo_llegar_original = normaliza_texto(row.get('Para la llegada al sitio con el equipo a instalar, se requiere de:', ''))
            # Separa por coma y normaliza solo el primer valor
            primer_valor = tipo_llegar_original.split(',')[0].strip() if ',' in tipo_llegar_original else tipo_llegar_original.strip()
            primer_tipo = normaliza_texto(primer_valor)
            print(f"primer_valor: '{primer_valor}'")
            print(f"primer_tipo normalizado: '{primer_tipo}'")
            ws_info_b.range('B72').value = (primer_tipo == 'pickup')
            ws_info_b.range('G72').value = (primer_tipo == 'pickup4x4')
            ws_info_b.range('M72').value = (primer_tipo == 'animalesdecarga')


            policia_original = normaliza_texto(row.get('Existe cerca del sitio alguna comandancia de policía o del ejercito?', '').strip().lower())
            ws_info_a.range('AB105').value = 'si' in policia_original
            ws_info_a.range('AE105').value = 'no' in policia_original

            if 'si' in policia_original:
                ws_info_a.range('U106').value = row.get('Si la respuesta anterior es si, Indique a que distancia  policia',) 
                ws_info_a.range('U107').value = row.get('Se cuenta con algún número de teléfono?, indíquelo', 'N/A') or 'N/A'
            else:
                ws_info_a.range('U106').value = 'N/A'
                ws_info_a.range('U107').value = 'N/A'

    # --- Cruz Roja, Hospital, asistencia médica ---
            cruzroja_original = normaliza_texto(row.get('Existe Cruz Roja, Hospital u otro tipo de asistencia medica cerca del sitio.', '').strip().lower())
            ws_info_a.range('AB108').value = 'si' in cruzroja_original
            ws_info_a.range('AE108').value = 'no' in cruzroja_original

            if 'si' in cruzroja_original:
                ws_info_a.range('U109').value = row.get('Si la respuesta anterior es si, Indique a que distancia cruz ', )  
                ws_info_a.range('U110').value = row.get('Se cuenta con algún numero de teléfono?, indíquelo: cruz', )
            else:
                ws_info_a.range('U109').value = 'N/A'
                ws_info_a.range('U110').value = 'N/A'

          # --- Mapa Nacional de Riesgos ---
            riesgo = normaliza_texto(row.get('Según el Mapa Nacional de Riesgos, indique en que zona se ubica el sitio:', '').strip().lower())
            ws_info_a.range('Y111').value = 'bajo' in riesgo
            ws_info_a.range('AB111').value = 'medio' in riesgo
            ws_info_a.range('AE111').value = 'alto' in riesgo



            policia_original = normaliza_texto(row.get('Existe cerca del sitio alguna comandancia de policía o del ejercito? B', ''))
            ws_info_b.range('AB105').value = 'si' in policia_original
            ws_info_b.range('AE105').value = 'no' in policia_original

            if 'si' in policia_original:
                ws_info_b.range('U106').value = row.get('Si la respuesta anterior es si, Indique a que distancia  policia B',) 
                ws_info_b.range('U107').value = row.get('Se cuenta con algún numero de teléfono?, indíquelo: B',) 
            else:
                ws_info_b.range('U106').value = 'N/A'
                ws_info_b.range('U107').value = 'N/A'

            # --- Cruz Roja, Hospital, asistencia médica ---
            cruzroja_original = normaliza_texto(row.get('Existe Cruz Roja, Hospital u otro tipo de asistencia medica cerca del sitio. B', ''))
            ws_info_b.range('AB108').value = 'si' in cruzroja_original
            ws_info_b.range('AE108').value = 'no' in cruzroja_original

            if 'si' in cruzroja_original:
                ws_info_b.range('U109').value = row.get('Si la respuesta anterior es si, Indique a que distancia cruz B', )  
                ws_info_b.range('U110').value = row.get('Se cuenta con algún numero de teléfono?, indíquelo: cruz B', )
            else:
                ws_info_b.range('U109').value = 'N/A'
                ws_info_b.range('U110').value = 'N/A'

              # --- Mapa Nacional de Riesgos ---
            riesgo = normaliza_texto(row.get('Según el Mapa Nacional de Riesgos, indique en que zona se ubica el sitio: B', ''))
            ws_info_b.range('Y111').value = 'bajo' in riesgo
            ws_info_b.range('AB111').value = 'medio' in riesgo
            ws_info_b.range('AE111').value = 'alto' in riesgo



            


            

                        # Llenado de campos específicos de la hoja "3. Espacios en Torre y Piso A-B"
            # Línea de vista
            linea_vista_original = normaliza_texto(row.get('Línea de vista', ''))
            linea_vista = normaliza_texto(linea_vista_original)
            # TODO: Agregar coordenadas correctas para PtMP
            # ws_espacios.range('R56').value = (linea_vista == 'si')
            # ws_espacios.range('V56').value = (linea_vista == 'no')
            # ws_espacios.range('Q57').value = False
            # ws_espacios.range('V57').value = False
            # ws_espacios.range('AC57').value = False
            # ws_espacios.range('AI57').value = False
            # ws_espacios.range('C58').value = False

            if linea_vista == 'no':
                motivo = normaliza_texto(row.get('Motivo de obstrucción', ''))
                # TODO: Agregar coordenadas correctas para PtMP
                # if motivo == 'arboles':
                #    ws_espacios.range('Q57').value = True
                # elif motivo == 'espectacular':
                #    ws_espacios.range('V57').value = True
                # elif motivo == 'edificio':
                #    ws_espacios.range('AC57').value = True
                # elif motivo == 'montaña':
                #    ws_espacios.range('AI57').value = True
                # elif motivo == 'n/a':
                #    ws_espacios.range('C58').value = True

            # Llenado automático de celdas usando el diccionario de mapeo para SECTOR
            campos_sector_celdas = {
                'Fecha Inicio Site Survey': 'C9',
                'Fecha final Site Survey': 'I9',
                'NOMBRE DEL SITIO': 'C10', 
                'PROPIETARIO': 'E11',
                'ID': 'I10',
                'ESTADO ':'I16',
                'Calle': 'B15',
                'Colonia': 'B16',
                'Municipio': 'B17',
                'C.P': 'I15',
                'Referencias':'C18',
                'Nombre de contacto en sitio': 'C20',
                'Telefono': 'H20',
                'LATITUD (TORRE)': 'D31',
                'LONGITUD (TORRE)': 'H31',
                'LATITUD (FACHADA)': 'D28',
                'LONGITUD (FACHADA)': 'H28',
                'Altitud (msnm)': 'E33',
                'Horario de solicitud de accesos': 'C53',
                'Contacto solicitud de accesos': 'B55',
                'Como o donde obtener permisos/llave/tarjeta': 'O49',
                'Comentario:Forma de ingresar el equipo al sitio es con:': 'B57',
                'comentario:En caso de requerirse grúa, identifique si es factible el uso de la misma y que no se tenga una posible obstrucción.': 'B63',
                'Si la respuesta anterior es si, Indique a que distancia ':'H83'
            }
            for campo, celdas in campos_sector_celdas.items():
                valor = normaliza_na(row.get(campo, ""))
                if isinstance(celdas, list):
                    for celda in celdas:
                        ws_sector.range(celda).value = valor
                else:
                    ws_sector.range(celdas).value = valor

            # Llenado automático de celdas usando el diccionario de mapeo para CPE
            campos_cpe_celdas = {
                'Fecha Inicio Site Survey B': 'C9',
                'Fecha final Site Survey B': 'I9',
                'Nombre del sitio 2': 'C10', 
                'ID 2': 'I10',
                'ESTADO 2 ':'I16',
                'Calle 2': 'B15',
                'Colonia 2': 'B16',
                'Municipio 2': 'B17',
                'C.P 2': 'I15',
                'Referencias 2':'C18',
                'Nombre de contacto en sitio 2': 'C20',
                'Telefono 2': 'H20',
                'LATITUD (TORRE) 2': 'D31',
                'LONGITUD (TORRE) 2': 'H31',
                'LATITUD (FACHADA) 2': 'D28',
                'LONGITUD (FACHADA) 2': 'H28',
                'Altitud (msnm) 2': 'E33',
                'Horario de solicitud de accesos B': 'C53',
                'Contacto solicitud de accesos B': 'B55',
                'Como o donde obtener permisos/llave/tarjeta B': 'C49',
                'Comentario:Forma de ingresar el equipo al sitio es con: B': 'B57',
                'comentario:En caso de requerirse grúa, identifique si es factible el uso de la misma y que no se tenga una posible obstrucción.': 'B63',
                'Si la respuesta anterior es si, Indique a que distancia B':'H83'
            }
            for campo, celdas in campos_cpe_celdas.items():
                valor = normaliza_na(row.get(campo, ""))
                print(f"ws_cpe: Escribiendo en {celdas} el valor '{valor}' para campo '{campo}'")
                if isinstance(celdas, list):
                    for celda in celdas:
                        ws_cpe.range(celda).value = valor
                else:
                    ws_cpe.range(celdas).value = valor    

            # Llenado automático de celdas para la hoja "3. Espacios en Torre y Piso A-B"
            campos_espacios_celdas = {
                'NOMBRE DEL SITIO': 'G7',
                'Diametro de pierna superior':'K9',
                'Diametro de pierna Inferior':'U9',
                'NCRA RB':'AC9',
                'Franja2RB':'AM9',
                'Altura de la Torre':'K10',
                'Dado':'U10',
                'Altura Edificio1':'EA10',
                'Nivel inferior de franja disponible': 'T11',
                'Nivel superior de franja disponible': 'AK11',
                'Altura de MW conforme a topologia': 'B14',
                'Azimut RB ': 'M14',
                'Propuesta de altura de antena de MW1': 'AB14',
                'Propuesta de altura de antena de MW (SD)1': 'AJ14',
                'Altura de soporte para OMB propuesto': 'O19',
                'Longitud del cable de tierra nuevo OMB': 'O20',
                'Longitud del cable de tierra ODU': 'O21',
                'Longitud de cable IF': 'O22',
                'Tipo de soporte para antena MW propuesto': 'O23',
                'Longitud de cable ACDB-Nuevo OMB': 'O24',
                'Longitud de cable RTN - Router':'O25',
                'Longitud de cable RTN - BBU SITE 1': 'O26',
                'MEDICION DE BARRA DE TIERRA (Ohms)':'O28',
                'Nombre del sitio 2': 'G31',
                'Diámetro de Pierna superio2':'K33',
                'Diámetro de Pierna inferior2':'U33',
                ' NCRA2 ':'AC33',
                'Franja2-2':'AM33',
                'Altura torre 2': 'K34',
                'DADO 2':'U34',
                'Altura edificio 2':'AE34',
                'Nivel inferior de franja disponible 2': 'T35',
                'Nivel superior de franja disponible 2': 'AK35',
                'Altura de MW conforme a topologia 2': 'B38',
                'Azimut 2': 'M38',
                'Propuesta de altura de antena de MW2': 'AB38',
                'Propuesta de altura de antena de MW (SD)2':'AJ38',
                'Altura de soporte para OMB propuesto2':'O43',
                'Longitud del cable de tierra nuevo OMB 2': 'O44',
                'Longitud del cable de tierra ODU 2': 'O45',
                'Longitud de cable IF 2': 'O46',
                'Tipo de soporte para antena MW propuesto 2': 'O47',
                'Longitud de cable ACDB-Nuevo OMB 2': 'O48',
                'Longitud de cable RTN - Router 2': 'O49',
                'Longitud de cable RTN - BBU 2': 'O50',
                'Medición del Sistema de Tierras 2': 'O52',
            }
            for campo, celdas in campos_espacios_celdas.items():
                valor_original = row.get(campo, "")
                valor = normaliza_na(valor_original)
                print(f"ws_espacios: Campo '{campo}' - Valor original: '{valor_original}' - Valor normalizado: '{valor}' - Escribiendo en {celdas}")
                
                # Log especial para Altura Edificio1
                if campo == 'Altura Edificio1':
                    print(f"*** DEBUG ESPECIAL Altura Edificio1 ***")
                    print(f"   - Campo en row: {'Altura Edificio1' in row}")
                    print(f"   - Valor original: '{valor_original}'")
                    print(f"   - Tipo del valor: {type(valor_original)}")
                    print(f"   - Valor normalizado: '{valor}'")
                    print(f"   - Celda destino: {celdas}")
                
                if isinstance(celdas, list):
                    for celda in celdas:
                        ws_espacios.range(celda).value = valor
                        print(f"ws_espacios: Escrito en celda {celda}: '{ws_espacios.range(celda).value}'")
                else:
                    ws_espacios.range(celdas).value = valor
                    print(f"ws_espacios: Escrito en celda {celdas}: '{ws_espacios.range(celdas).value}'")   

            # Guardar y cerrar
            wb.save(output_path)
            wb.close()
            app_excel.quit()

            print('DEBUG: Excel PtMP cerrado, antes del redirect')
            return redirect(url_for('descargar_site_survey_ptmp', user_id=user_id, fila_idx=fila_idx))
        except Exception as e:
            print(f"ERROR en llenado ptmp_site_survey: {e}")
            return f"Error en llenado ptmp_site_survey: {e}"
    elif tipo == 'ptmp_reporte_planeacion':
        print('DEBUG: Entrando a bloque ptmp_reporte_planeacion')
        # TODO: Implementar llenado de Reporte de Planeación PtMP
        return f"Funcionalidad PtMP Reporte de Planeación en desarrollo para {user_id}"
    elif tipo == 'ptmp_diseno_solucion':
        print('DEBUG: Entrando a bloque ptmp_diseno_solucion')
        # TODO: Implementar llenado de Diseño de Solución PtMP
        return f"Funcionalidad PtMP Diseño de Solución en desarrollo para {user_id}"
    else:
        print('DEBUG: Entrando a bloque else (formulario_archivos)')
        return redirect(url_for('formulario_archivos', user_id=user_id, fila_idx=fila_idx))
    print('DEBUG: Fin de redirigir_tipo_llenado (esto no debería verse si todos los returns están bien)')

@app.route('/subir_imagenes_ptp', methods=['GET', 'POST'])
def subir_imagenes_ptp():
    if request.method == 'GET':
        user_id = request.args.get('user_id')
        fila_idx = request.args.get('fila_idx')
        return render_template('subir_imagenes_ptp.html', user_id=user_id, fila_idx=fila_idx)
    
    if request.method == 'POST':
        user_id = request.form.get('user_id')
        fila_idx = request.form.get('fila_idx')
        
        # Definir base_dir
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Crear directorio para imágenes si no existe
        imagenes_dir = os.path.join(base_dir, 'uploads', 'imagenes_ptp', user_id)
        os.makedirs(imagenes_dir, exist_ok=True)
        
        # Procesar cada imagen subida
        imagenes_procesadas = []
        
        for i in range(1, 4):  # Soporte para hasta 10 imágenes
            imagen_key = f'imagen_{i}'
            if imagen_key in request.files:
                archivo = request.files[imagen_key]
                if archivo and archivo.filename != '':
                    # Generar nombre único para la imagen
                    nombre_archivo = f"imagen_{i}_{user_id}_{int(time.time())}.jpg"
                    ruta_archivo = os.path.join(imagenes_dir, nombre_archivo)
                    
                    # Guardar la imagen
                    archivo.save(ruta_archivo)
                    imagenes_procesadas.append({
                        'numero': i,
                        'ruta': ruta_archivo,
                        'nombre_original': archivo.filename
                    })
        
        # Guardar información de las imágenes en el Excel
        try:
            import xlwings as xw
            plantilla_path = os.path.join(base_dir, 'static', 'plantillas', 'EJEMPLO SS VACIO.xlsx')
            user_id_limpio = re.sub(r'[<>:"/\\|?*]', '', str(user_id))
            output_path = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}.xlsx')
            
            # Verificar que el archivo Excel existe (buscar también con timestamp)
            if not os.path.exists(output_path):
                # Buscar archivos que coincidan con el patrón
                import glob
                patron_archivo = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}*.xlsx')
                archivos_encontrados = glob.glob(patron_archivo)
                
                if archivos_encontrados:
                    print(f"DEBUG: Archivos encontrados: {archivos_encontrados}")
                    # Usar el archivo más reciente
                    archivo_mas_reciente = max(archivos_encontrados, key=os.path.getctime)
                    output_path = archivo_mas_reciente
                    print(f"DEBUG: Usando archivo más reciente: {output_path}")
                else:
                    print(f"DEBUG: No se encontraron archivos con patrón: {patron_archivo}")
                    return jsonify({
                        'success': False,
                        'message': f'El archivo Excel no existe. Primero debes generar el Site Survey. Ruta esperada: {output_path}'
                    })
            
            # Abrir el archivo Excel existente
            app_excel = xw.App(visible=False)
            wb = app_excel.books.open(output_path)
            ws_info_d = wb.sheets['4. Planos A']
            
            # Insertar las imágenes en las celdas correspondientes
            for imagen in imagenes_procesadas:
                # Definir las celdas donde insertar las imágenes
                celdas_imagenes = {
                    1: 'C17:AK60',  # Imagen 1
                    2: 'C69:AK123', # Imagen 2
                    3: 'C134:AK173', # Imagen 3
                }
                
                if imagen['numero'] in celdas_imagenes:
                    rango_celda = celdas_imagenes[imagen['numero']]
                    # Insertar la imagen en el rango específico
                    ws_info_d.pictures.add(imagen['ruta'], 
                                         left=ws_info_d.range(rango_celda).left,
                                         top=ws_info_d.range(rango_celda).top,
                                         width=ws_info_d.range(rango_celda).width,
                                         height=ws_info_d.range(rango_celda).height)
            
            # Guardar y cerrar
            wb.save()
            wb.close()
            app_excel.quit()
            
            return jsonify({
                'success': True,
                'message': f'Se subieron {len(imagenes_procesadas)} imágenes correctamente',
                'imagenes': imagenes_procesadas
            })
            
        except Exception as e:
            error_msg = str(e)
            if "No such file" in error_msg:
                return jsonify({
                    'success': False,
                    'message': f'El archivo Excel no existe. Primero debes generar el Site Survey PTP antes de subir imágenes. Error: {error_msg}'
                })
            else:
                return jsonify({
                    'success': False,
                    'message': f'Error al procesar las imágenes: {error_msg}'
                })
    
    return "Método no permitido"

@app.route('/subir_imagenes_ptp_planos_b', methods=['GET', 'POST'])
def subir_imagenes_ptp_planos_b():
    if request.method == 'GET':
        user_id = request.args.get('user_id')
        fila_idx = request.args.get('fila_idx')
        return render_template('subir_imagenes_ptp_planos_b.html', user_id=user_id, fila_idx=fila_idx)
    
    if request.method == 'POST':
        user_id = request.form.get('user_id')
        fila_idx = request.form.get('fila_idx')
        
        # Definir base_dir
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Crear directorio para imágenes si no existe
        imagenes_dir = os.path.join(base_dir, 'uploads', 'imagenes_ptp_planos_b', user_id)
        os.makedirs(imagenes_dir, exist_ok=True)
        
        # Procesar cada imagen subida
        imagenes_procesadas = []
        
        for i in range(1, 4):  # Soporte para hasta 3 imágenes
            imagen_key = f'imagen_{i}'
            if imagen_key in request.files:
                archivo = request.files[imagen_key]
                if archivo and archivo.filename != '':
                    nombre_archivo = f"imagen_{i}_{user_id}_{int(time.time())}.jpg"
                    ruta_archivo = os.path.join(imagenes_dir, nombre_archivo)
                    archivo.save(ruta_archivo)
                    imagenes_procesadas.append({
                        'numero': i,
                        'ruta': ruta_archivo,
                        'nombre_original': archivo.filename
                    })
        
        # Guardar información de las imágenes en el Excel
        try:
            import xlwings as xw
            user_id_limpio = re.sub(r'[<>:"/\\|?*]', '', str(user_id))
            output_path = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}.xlsx')
            
            # Verificar que el archivo Excel existe (buscar también con timestamp)
            if not os.path.exists(output_path):
                # Buscar archivos que coincidan con el patrón
                import glob
                patron_archivo = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}*.xlsx')
                archivos_encontrados = glob.glob(patron_archivo)
                
                if archivos_encontrados:
                    print(f"DEBUG: Archivos encontrados: {archivos_encontrados}")
                    # Usar el archivo más reciente
                    archivo_mas_reciente = max(archivos_encontrados, key=os.path.getctime)
                    output_path = archivo_mas_reciente
                    print(f"DEBUG: Usando archivo más reciente: {output_path}")
                else:
                    print(f"DEBUG: No se encontraron archivos con patrón: {patron_archivo}")
                    return jsonify({
                        'success': False,
                        'message': f'El archivo Excel no existe. Primero debes generar el Site Survey. Ruta esperada: {output_path}'
                    })
            
            # Abrir el archivo Excel existente
            app_excel = xw.App(visible=False)
            wb = app_excel.books.open(output_path)
            ws_info_e = wb.sheets['5. Planos B']
            
            # Insertar las imágenes en las celdas correspondientes
            for imagen in imagenes_procesadas:
                # Definir las celdas donde insertar las imágenes (mismos rangos que Planos A)
                celdas_imagenes = {
                    1: 'C17:AK60',  # Imagen 1
                    2: 'C69:AK123', # Imagen 2
                    3: 'C134:AK173', # Imagen 3
                }
                
                if imagen['numero'] in celdas_imagenes:
                    rango_celda = celdas_imagenes[imagen['numero']]
                    # Insertar la imagen en el rango específico
                    ws_info_e.pictures.add(imagen['ruta'], 
                                         left=ws_info_e.range(rango_celda).left,
                                         top=ws_info_e.range(rango_celda).top,
                                         width=ws_info_e.range(rango_celda).width,
                                         height=ws_info_e.range(rango_celda).height)
            
            # Guardar y cerrar
            wb.save()
            wb.close()
            app_excel.quit()
            
            return jsonify({
                'success': True,
                'message': f'Se subieron {len(imagenes_procesadas)} imágenes correctamente a Planos B',
                'imagenes': imagenes_procesadas
            })
            
        except Exception as e:
            error_msg = str(e)
            if "No such file" in error_msg:
                return jsonify({
                    'success': False,
                    'message': f'El archivo Excel no existe. Primero debes generar el Site Survey PTP antes de subir imágenes. Error: {error_msg}'
                })
            else:
                return jsonify({
                    'success': False,
                    'message': f'Error al procesar las imágenes: {error_msg}'
                })
    
    return "Método no permitido"

@app.route('/subir_imagenes_ptp_fotos_a', methods=['GET', 'POST'])
def subir_imagenes_ptp_fotos_a():
    if request.method == 'GET':
        user_id = request.args.get('user_id')
        fila_idx = request.args.get('fila_idx')
        return render_template('subir_imagenes_ptp_fotos_a.html', user_id=user_id, fila_idx=fila_idx)
    
    if request.method == 'POST':
        user_id = request.form.get('user_id')
        fila_idx = request.form.get('fila_idx')
        
        # Definir base_dir
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Crear directorio para imágenes si no existe
        imagenes_dir = os.path.join(base_dir, 'uploads', 'imagenes_ptp_fotos_a', user_id)
        os.makedirs(imagenes_dir, exist_ok=True)
        
        # Procesar cada imagen subida
        imagenes_procesadas = []
        
        for i in range(1, 20):  # Soporte para hasta 19 imágenes
            imagen_key = f'imagen_{i}'
            if imagen_key in request.files:
                archivo = request.files[imagen_key]
                if archivo and archivo.filename != '':
                    nombre_archivo = f"imagen_{i}_{user_id}_{int(time.time())}.jpg"
                    ruta_archivo = os.path.join(imagenes_dir, nombre_archivo)
                    archivo.save(ruta_archivo)
                    imagenes_procesadas.append({
                        'numero': i,
                        'ruta': ruta_archivo,
                        'nombre_original': archivo.filename
                    })
        
        # Guardar información de las imágenes en el Excel
        try:
            import xlwings as xw
            user_id_limpio = re.sub(r'[<>:"/\\|?*]', '', str(user_id))
            output_path = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}.xlsx')
            
            # Verificar que el archivo Excel existe (buscar también con timestamp)
            if not os.path.exists(output_path):
                # Buscar archivos que coincidan con el patrón
                import glob
                patron_archivo = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}*.xlsx')
                archivos_encontrados = glob.glob(patron_archivo)
                
                if archivos_encontrados:
                    print(f"DEBUG: Archivos encontrados: {archivos_encontrados}")
                    # Usar el archivo más reciente
                    archivo_mas_reciente = max(archivos_encontrados, key=os.path.getctime)
                    output_path = archivo_mas_reciente
                    print(f"DEBUG: Usando archivo más reciente: {output_path}")
                else:
                    print(f"DEBUG: No se encontraron archivos con patrón: {patron_archivo}")
                    return jsonify({
                        'success': False,
                        'message': f'El archivo Excel no existe. Primero debes generar el Site Survey. Ruta esperada: {output_path}'
                    })
            
            # Abrir el archivo Excel existente
            app_excel = xw.App(visible=False)
            wb = app_excel.books.open(output_path)
            ws_info_f = wb.sheets['6. Reporte Fotos A']
            
            # Definir las celdas donde insertar las imágenes para Reporte Fotos A
            celdas_imagenes = {
                1: 'G11:L18',   # 6.3.1 Foto linea de Vista de Sitio A a Sitio B
                2: 'X11:AD18',  # 6.3.2 Foto linea de Vista de Sitio A a Sitio B Diversidad
                3: 'G23:L30',   # 6.1.3 Foto de torre completa
                4: 'G48:L55',   # 6.2.1 Foto desde piso mostrando el espacio en torre para antena de MW conforme a topologia
                5: 'X48:AD55',  # 6.2.2 Medición con cinta del rad center en torre conforme a topologia
                6: 'G59:L66',   # 6.2.3 Foto desde piso mostrando el espacio en torre para antena de MW (SD)
                7: 'X59:AD66',  # 6.2.4 Medición con cinta del rad center en torre (SD)
                8: 'G70:L77',   # 6.2.5 Foto desde piso mostrando el espacio (propuesto) en torre para antena
                9: 'X70:AD77',  # 6.2.6 Medición con cinta del rad center (propuesto) en torre conforme a topologia
                10: 'G87:L94',  # 6.3.1 Foto linea de Vista de Sitio A a Sitio B
                11: 'X87:AD94', # 6.3.2 Foto linea de Vista de Sitio A a Sitio B Diversidad
                12: 'G98:L105', # 6.3.3 Foto Barra de Tierra
                13: 'X98:AD105', # 6.3.4 Foto de escalerilla de torre
                14: 'G127:L134', # 6.4.1 Foto del espacio disponible dentro del Gabinete OMB
                15: 'X127:AD134', # 6.4.2 Foto del espacio disponible en torre para OMB adicional
                16: 'G138:L145', # 6.4.3 Foto DPU existente
                17: 'X138:AD145', # 6.4.4 Foto del espacio disponible en torre para DPU y Bateria
                18: 'G150:L157', # 6.4.5 Foto ACDB y Breaker
                19: 'X150:AD157', # 6.4.6 Foto de Agregador (Site Entry)
            }
            
            # Insertar las imágenes en las celdas correspondientes
            for imagen in imagenes_procesadas:
                if imagen['numero'] in celdas_imagenes:
                    rango_celda = celdas_imagenes[imagen['numero']]
                    # Insertar la imagen en el rango específico
                    ws_info_f.pictures.add(imagen['ruta'], 
                                         left=ws_info_f.range(rango_celda).left,
                                         top=ws_info_f.range(rango_celda).top,
                                         width=ws_info_f.range(rango_celda).width,
                                         height=ws_info_f.range(rango_celda).height)
            
            # Colocar N/A en las celdas donde no se subieron imágenes
            for i in range(1, 20):
                if i not in [img['numero'] for img in imagenes_procesadas]:
                    if i in celdas_imagenes:
                        rango_celda = celdas_imagenes[i]
                        # Obtener la celda central del rango para colocar N/A
                        rango_obj = ws_info_f.range(rango_celda)
                        celda_central = ws_info_f.range(rango_obj.left + rango_obj.width/2, rango_obj.top + rango_obj.height/2)
                        celda_central.value = "N/A"
            
            # Guardar y cerrar
            wb.save()
            wb.close()
            app_excel.quit()
            
            return jsonify({
                'success': True,
                'message': f'Se subieron {len(imagenes_procesadas)} imágenes correctamente a Reporte Fotos A',
                'imagenes': imagenes_procesadas
            })
            
        except Exception as e:
            error_msg = str(e)
            if "No such file" in error_msg:
                return jsonify({
                    'success': False,
                    'message': f'El archivo Excel no existe. Primero debes generar el Site Survey PTP antes de subir imágenes. Error: {error_msg}'
                })
            else:
                return jsonify({
                    'success': False,
                    'message': f'Error al procesar las imágenes: {error_msg}'
                })
    
    return "Método no permitido"

@app.route('/subir_imagenes_ptp_fotos_b', methods=['GET', 'POST'])
def subir_imagenes_ptp_fotos_b():
    if request.method == 'GET':
        user_id = request.args.get('user_id')
        fila_idx = request.args.get('fila_idx')
        return render_template('subir_imagenes_ptp_fotos_b.html', user_id=user_id, fila_idx=fila_idx)
    
    if request.method == 'POST':
        user_id = request.form.get('user_id')
        fila_idx = request.form.get('fila_idx')
        
        # Definir base_dir
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Crear directorio para imágenes si no existe
        imagenes_dir = os.path.join(base_dir, 'uploads', 'imagenes_ptp_fotos_b', user_id)
        os.makedirs(imagenes_dir, exist_ok=True)
        
        # Procesar cada imagen subida
        imagenes_procesadas = []
        
        print(f"DEBUG: Procesando imágenes para Reporte Fotos B - User ID: {user_id}")
        print(f"DEBUG: Archivos recibidos: {list(request.files.keys())}")
        
        for i in range(1, 20):  # Soporte para hasta 19 imágenes (todas las imágenes de Reporte Fotos B)
            imagen_key = f'imagen_{i}'
            print(f"DEBUG: Verificando {imagen_key}")
            
            if imagen_key in request.files:
                archivo = request.files[imagen_key]
                print(f"DEBUG: Archivo encontrado para {imagen_key}: {archivo.filename}")
                
                if archivo and archivo.filename != '':
                    nombre_archivo = f"imagen_{i}_{user_id}_{int(time.time())}.jpg"
                    ruta_archivo = os.path.join(imagenes_dir, nombre_archivo)
                    print(f"DEBUG: Guardando archivo en: {ruta_archivo}")
                    
                    archivo.save(ruta_archivo)
                    imagenes_procesadas.append({
                        'numero': i,
                        'ruta': ruta_archivo,
                        'nombre_original': archivo.filename
                    })
                    print(f"DEBUG: Imagen {i} procesada correctamente")
                else:
                    print(f"DEBUG: Archivo {imagen_key} está vacío o no tiene nombre")
            else:
                print(f"DEBUG: No se encontró archivo para {imagen_key}")
        
        print(f"DEBUG: Total de imágenes procesadas: {len(imagenes_procesadas)}")
        print(f"DEBUG: Imágenes procesadas: {imagenes_procesadas}")
        
        # Guardar información de las imágenes en el Excel
        try:
            import xlwings as xw
            user_id_limpio = re.sub(r'[<>:"/\\|?*]', '', str(user_id))
            output_path = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}.xlsx')
            
            # Verificar que el archivo Excel existe (buscar también con timestamp)
            if not os.path.exists(output_path):
                # Buscar archivos que coincidan con el patrón
                import glob
                patron_archivo = os.path.join(base_dir, 'site_survey', f'ss_{user_id_limpio}*.xlsx')
                archivos_encontrados = glob.glob(patron_archivo)
                
                if archivos_encontrados:
                    print(f"DEBUG: Archivos encontrados: {archivos_encontrados}")
                    # Usar el archivo más reciente
                    archivo_mas_reciente = max(archivos_encontrados, key=os.path.getctime)
                    output_path = archivo_mas_reciente
                    print(f"DEBUG: Usando archivo más reciente: {output_path}")
                else:
                    print(f"DEBUG: No se encontraron archivos con patrón: {patron_archivo}")
                    return jsonify({
                        'success': False,
                        'message': f'El archivo Excel no existe. Primero debes generar el Site Survey. Ruta esperada: {output_path}'
                    })
            
            # Abrir el archivo Excel existente
            print(f"DEBUG: Abriendo archivo Excel: {output_path}")
            app_excel = xw.App(visible=False)
            wb = app_excel.books.open(output_path)
            ws_info_g = wb.sheets['7. Reporte Fotos B']
            print(f"DEBUG: Hoja '7. Reporte Fotos B' abierta correctamente")
            
            # Definir las celdas donde insertar las imágenes para Reporte Fotos B
            celdas_imagenes = {
                1: 'G11:L18', # 7.4.4 Foto del espacio disponible en torre para DPU y Bateria
                2: 'X11:AD18',  # 7.4.5 Foto ACDB y Breaker
                3: 'G23:L30', # 7.4.6 Foto de Agregador (Site Entry)
                4: 'G48:L55',  # 7.2.1 Foto desde piso mostrando el espacio en torre para antena de MW conforme a topologia
                5: 'X48:AD55', # 7.2.2 Medición con cinta del rad center en torre conforme a topologia
                6: 'G59:L66',  # 7.1.3 Foto de torre completa
                7: 'X59:AD66', # 7.2.3 Foto desde piso mostrando el espacio en torre para antena de MW (SD)
                8: 'G70:L77',  # 7.2.4 Medición con cinta del rad center en torre (SD)
                9: 'X70:AD77', # 7.2.5 Foto desde piso mostrando el espacio (propuesto) en torre para antena
                10: 'G87:L94', # 7.2.6 Medición con cinta del rad center (propuesto) en torre conforme a topologia
                11: 'X87:AD94', # 7.3.1 Foto linea de Vista de Sitio A a Sitio B
                12: 'G98:L105', # 7.3.2 Foto linea de Vista de Sitio A a Sitio B Diversidad
                13: 'X98:AD105', # 7.3.3 Foto Barra de Tierra
                14: 'G127:L134', # 7.3.4 Foto de escalerilla de torre
                15: 'X127:AD134', # 7.4.1 Foto del espacio disponible dentro del Gabinete OMB
                16: 'G138:L145', # 7.4.2 Foto del espacio disponible en torre para OMB adicional
                17: 'X138:AD145', # 7.4.3 Foto DPU existente
                18: 'G150:L157', # 7.4.4 Foto del espacio disponible en torre para DPU y Bateria
                19: 'X150:AD157', # 7.4.5 Foto ACDB y Breaker
            }
            print(f"DEBUG: Rangos de celdas definidos: {celdas_imagenes}")
            
            # Insertar las imágenes en las celdas correspondientes
            for imagen in imagenes_procesadas:
                print(f"DEBUG: Procesando imagen {imagen['numero']}: {imagen['ruta']}")
                if imagen['numero'] in celdas_imagenes:
                    rango_celda = celdas_imagenes[imagen['numero']]
                    print(f"DEBUG: Insertando imagen en rango: {rango_celda}")
                    # Insertar la imagen en el rango específico
                    ws_info_g.pictures.add(imagen['ruta'], 
                                         left=ws_info_g.range(rango_celda).left,
                                         top=ws_info_g.range(rango_celda).top,
                                         width=ws_info_g.range(rango_celda).width,
                                         height=ws_info_g.range(rango_celda).height)
                    print(f"DEBUG: Imagen {imagen['numero']} insertada correctamente en Excel")
                else:
                    print(f"DEBUG: ERROR - No se encontró rango para imagen {imagen['numero']}")
            
            # Colocar N/A en las celdas donde no se subieron imágenes
            for i in range(1, 20):
                if i not in [img['numero'] for img in imagenes_procesadas]:
                    if i in celdas_imagenes:
                        rango_celda = celdas_imagenes[i]
                        print(f"DEBUG: Colocando N/A en rango {rango_celda} para imagen {i}")
                        # Obtener la celda central del rango para colocar N/A
                        rango_obj = ws_info_g.range(rango_celda)
                        celda_central = ws_info_g.range(rango_obj.left + rango_obj.width/2, rango_obj.top + rango_obj.height/2)
                        celda_central.value = "N/A"
                        print(f"DEBUG: N/A colocado correctamente en imagen {i}")
                    else:
                        print(f"DEBUG: No se encontró rango para colocar N/A en imagen {i}")
                else:
                    print(f"DEBUG: Imagen {i} fue subida, no se coloca N/A")
            
            # Guardar y cerrar
            print(f"DEBUG: Guardando archivo Excel...")
            wb.save()
            print(f"DEBUG: Archivo guardado correctamente")
            wb.close()
            app_excel.quit()
            print(f"DEBUG: Excel cerrado correctamente")
            
            return jsonify({
                'success': True,
                'message': f'Se subieron {len(imagenes_procesadas)} imágenes correctamente a Reporte Fotos B',
                'imagenes': imagenes_procesadas
            })
            
        except Exception as e:
            error_msg = str(e)
            if "No such file" in error_msg:
                return jsonify({
                    'success': False,
                    'message': f'El archivo Excel no existe. Primero debes generar el Site Survey PTP antes de subir imágenes. Error: {error_msg}'
                })
            else:
                return jsonify({
                    'success': False,
                    'message': f'Error al procesar las imágenes: {error_msg}'
                })
    
    return "Método no permitido"

@app.route('/seleccion_id', methods=['POST'])
def seleccion_id():
    import pandas as pd
    user_id = request.form.get('user_id')
    if not user_id:
        return "Falta el ID"
    try:
        df_db = pd.read_csv(GOOGLE_SHEETS_CSV_URL, keep_default_na=False, na_values=[])
    except Exception as e:
        return f"Error leyendo la base de datos de Google Sheets: {e}"
    coincidencias = df_db[df_db['ID'] == user_id]
    if coincidencias.empty:
        return "ID no encontrado en la base de datos."
    if len(coincidencias) > 1:
        opciones = []
        for idx, row in coincidencias.iterrows():
            opciones.append({'idx': idx, 'sitio_a': row.get('Nombre del sitio A', ''), 'sitio_b': row.get('Nombre del sitio B', ''), 'analisis': row.get('Análisis', '')})
        return render_template('seleccion_registro.html', user_id=user_id, opciones=opciones)
    fila_idx = coincidencias.index[0]
    return redirect(url_for('seleccion_tipo_llenado', user_id=user_id, fila_idx=fila_idx))

def normaliza_na(valor):
    # Debug: imprimir el valor original y su tipo
    print(f"DEBUG normaliza_na: Valor original: '{valor}' (tipo: {type(valor)})")
    
    # Si es None o NaN de pandas
    if valor is None or pd.isna(valor):
        print(f"DEBUG normaliza_na: Valor es None/NaN, retornando 'N/A'")
        return "N/A"
    
    # Si es string
    if isinstance(valor, str):
        valor_limpio = valor.strip()
        # Casos vacíos
        if valor_limpio == "":
            print(f"DEBUG normaliza_na: String vacío, retornando 'N/A'")
            return "N/A"
        # Casos N/A (diferentes variaciones)
        if valor_limpio.lower() in ["n/a", "na", "n.a", "n.a.", "no aplica", "no disponible"]:
            print(f"DEBUG normaliza_na: String N/A detectado, retornando 'N/A'")
            return "N/A"
        # Si no es N/A, retornar el valor original
        print(f"DEBUG normaliza_na: String válido, retornando '{valor_limpio}'")
        return valor_limpio
    
    # Si es número
    if isinstance(valor, (int, float)):
        print(f"DEBUG normaliza_na: Número detectado, retornando '{valor}'")
        return str(valor)
    
    # Para cualquier otro tipo, convertir a string
    print(f"DEBUG normaliza_na: Otro tipo, convirtiendo a string: '{str(valor)}'")
    return str(valor)

def normaliza_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().lower()
    import unicodedata
    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )
    # Elimina paréntesis y otros signos
    for char in "-_.,;:()[]{}":
        texto = texto.replace(char, '')
    texto = texto.replace(' ', '')
    return texto




if __name__ == "__main__":
    try:
        print("Iniciando servidor Flask en modo producción...")
        print("El servidor estará disponible en:")
        print("  - Local: http://127.0.0.1:5000")
        print("  - Red: http://0.0.0.0:5000")
        print("Presiona Ctrl+C para detener el servidor")
        # Modo producción sin debug para evitar el servidor de desarrollo
        app.run(debug=False, use_reloader=False, host='0.0.0.0', port=5000)
    except KeyboardInterrupt:
        print("\nServidor detenido por el usuario")
    except Exception as e:
        print(f"Error al iniciar el servidor: {e}")
        input("Presiona Enter para salir...")
