from app import app
import webbrowser
import threading
import sys
import os

# Corregir el problema de sys.stdin
if hasattr(sys, '_MEIPASS'):
    # Si estamos en un ejecutable PyInstaller
    if sys.stdin is None:
        import io
        sys.stdin = io.StringIO('')

def open_browser():
    try:
        webbrowser.open_new("http://127.0.0.1:5000")
    except Exception as e:
        print(f"No se pudo abrir el navegador: {e}")
        print("Abre manualmente: http://127.0.0.1:5000")

if __name__ == "__main__":
    try:
        # Iniciar el navegador después de 2 segundos
        threading.Timer(2.0, open_browser).start()
        
        # Ejecutar la aplicación Flask
        print("🚀 Iniciando Site Survey App...")
        print("📱 La aplicación se abrirá en tu navegador")
        print("🌐 URL: http://127.0.0.1:5000")
        print("⏹️  Para cerrar, presiona Ctrl+C en esta ventana")
        print("-" * 50)
        
        app.run(debug=False, use_reloader=False, host="127.0.0.1", port=5000)
        
    except KeyboardInterrupt:
        print("\n👋 Aplicación cerrada por el usuario")
    except Exception as e:
        print(f"❌ Error al iniciar la aplicación: {e}")
        input("Presiona Enter para salir...") 