#!/usr/bin/env python3
"""
Servidor para acceso remoto - Equipo CDMX y GDL
"""

import os
import sys
import socket
import webbrowser
from threading import Timer

def get_local_ip():
    """Obtener IP local para acceso remoto"""
    try:
        # Conectar a un servidor externo para obtener IP local
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "127.0.0.1"

def main():
    print("=== Site Survey Enterprise - Servidor Remoto ===")
    
    try:
        # Importar y crear la aplicación
        from app_enterprise import create_app, init_database
        
        # Crear aplicación
        app = create_app('development')
        print("✓ Aplicación creada")
        
        # Inicializar base de datos
        init_database(app)
        print("✓ Base de datos inicializada")
        
        # Obtener IP local
        local_ip = get_local_ip()
        port = 5000
        
        print("\n" + "="*60)
        print("🚀 SERVIDOR INICIADO EXITOSAMENTE!")
        print("="*60)
        print(f"📍 IP Local: {local_ip}")
        print(f"🌐 Puerto: {port}")
        print("\n📱 URLs de acceso:")
        print(f"   🔗 Local (CDMX): http://localhost:{port}")
        print(f"   🌍 Remoto (GDL): http://{local_ip}:{port}")
        print("\n📋 Funcionalidades:")
        print("   🔧 Admin: http://localhost:5000/admin")
        print("   📊 Dashboard: http://localhost:5000/static/dashboard.html")
        print("   🔑 Login: http://localhost:5000/auth/login")
        print("   📋 Site Survey: http://localhost:5000/site-survey")
        print("\n🔑 Credenciales:")
        print("   Usuario: admin")
        print("   Contraseña: admin123")
        print("\n⚠️  IMPORTANTE:")
        print("   1. Asegúrate de que el firewall permita conexiones al puerto 5000")
        print("   2. Tu equipo en GDL debe acceder a: http://" + local_ip + ":5000")
        print("   3. Para detener: Ctrl+C")
        print("="*60)
        
        # Abrir navegador local
        Timer(2.0, lambda: webbrowser.open(f'http://localhost:{port}')).start()
        
        # Ejecutar servidor
        app.run(debug=True, host='0.0.0.0', port=port)
        
    except ImportError as e:
        print(f"❌ Error de importación: {e}")
        print("💡 Ejecuta: pip install -r requirements_dev.txt")
        input("Presiona Enter para continuar...")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        print("💡 Verifica que todos los archivos estén presentes")
        input("Presiona Enter para continuar...")

if __name__ == '__main__':
    main() 