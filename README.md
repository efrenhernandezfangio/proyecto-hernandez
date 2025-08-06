# Site Survey App

Aplicación web para la gestión y generación de reportes de Site Survey.

## Características

- Generación automática de reportes de Site Survey
- Soporte para PTP y PTMP
- Interfaz web intuitiva
- Procesamiento de archivos Excel
- Subida de imágenes y documentos

## Instalación

1. Clona el repositorio:
```bash
git clone [URL_DEL_REPOSITORIO]
cd nuevo-baseado
```

2. Instala las dependencias:
```bash
pip install -r requirements.txt
```

## Uso

### Opción 1: Servidor de Producción (Recomendado)
```bash
python run_app_production.py
```

### Opción 2: Servidor de Desarrollo
```bash
python app.py
```

### Opción 3: Archivo Batch (Windows)
Doble clic en `ejecutar_sin_navegador.bat`

## Acceso

Una vez iniciado el servidor, accede a:
- **Local:** http://127.0.0.1:5000
- **Red:** http://[TU_IP]:5000

## Estructura del Proyecto

```
nuevo-baseado/
├── app.py                          # Aplicación principal Flask
├── run_app_production.py           # Servidor de producción con Waitress
├── requirements.txt                # Dependencias Python
├── templates/                      # Plantillas HTML
├── static/                         # Archivos estáticos
├── uploads/                        # Archivos subidos
└── site_survey/                    # Reportes generados
```

## Dependencias Principales

- Flask
- Waitress (servidor WSGI)
- pandas
- xlwings
- openpyxl
- matplotlib

## Licencia

[Especificar licencia] 