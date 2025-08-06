@echo off
echo ========================================
echo    INSTALADOR SITE SURVEY APP v1.0
echo ========================================
echo.
echo Instalando Site Survey App...
set "DESKTOP=%USERPROFILE%\Desktop"
set "APP_FOLDER=%DESKTOP%\SiteSurveyApp"

echo Creando carpeta de aplicacion...
if not exist "%APP_FOLDER%" mkdir "%APP_FOLDER%"

echo Copiando archivos...
copy "SiteSurveyApp.exe" "%APP_FOLDER%\"

echo Creando acceso directo...
echo @echo off > "%DESKTOP%\Site Survey App.bat"
echo cd /d "%APP_FOLDER%" >> "%DESKTOP%\Site Survey App.bat"
echo start "" "SiteSurveyApp.exe" >> "%DESKTOP%\Site Survey App.bat"

echo.
echo ========================================
echo    INSTALACION COMPLETADA
echo ========================================
echo.
echo La aplicacion se ha instalado en:
echo %APP_FOLDER%
echo.
echo Puedes ejecutarla desde el acceso directo
echo en tu escritorio: "Site Survey App.bat"
echo.
echo IMPORTANTE:
echo - Si la app no se abre, ejecuta como administrador
echo - Verifica que Windows Defender no la bloquee
echo - La app usa puerto 5000 (cerrar otras apps si hay conflicto)
echo.
pause
