@echo off
cd /d "%~dp0"
echo Iniciando servidor de actualizaciones en http://localhost:3001
echo.
echo Manten esta ventana abierta mientras usas el dashboard.
echo Pulsa Ctrl+C para detener.
echo.
node server.js
pause
