from app import app
from waitress import serve
import os

if __name__ == "__main__":
    try:
        print("Iniciando servidor de producción con Waitress...")
        print("El servidor estará disponible en:")
        print("  - Local: http://127.0.0.1:5000")
        print("  - Red: http://0.0.0.0:5000")
        print("Presiona Ctrl+C para detener el servidor")
        
        # Configuración de Waitress para producción
        serve(app, host='0.0.0.0', port=5000, threads=4)
        
    except KeyboardInterrupt:
        print("\nServidor detenido por el usuario")
    except Exception as e:
        print(f"Error al iniciar el servidor: {e}")
        input("Presiona Enter para salir...") 