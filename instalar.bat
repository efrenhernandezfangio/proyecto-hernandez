@echo off
echo Instalando Site Survey App...
set "DESKTOP=%USERPROFILE%\Desktop"
copy "SiteSurveyApp.exe" "%DESKTOP%\"
echo @echo off > "%DESKTOP%\Site Survey App.bat"
echo start "" "%DESKTOP%\SiteSurveyApp.exe" >> "%DESKTOP%\Site Survey App.bat"
echo ¡Instalación completada!
pause
