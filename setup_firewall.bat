@echo off
echo === Configurando Firewall para Acceso Remoto ===
echo.

echo Configurando reglas de firewall...
netsh advfirewall firewall add rule name="Site Survey Enterprise" dir=in action=allow protocol=TCP localport=5000

echo.
echo ✅ Firewall configurado exitosamente!
echo.
echo Ahora puedes ejecutar: python run_server_remote.py
echo.
pause 