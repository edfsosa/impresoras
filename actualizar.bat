@echo off
title Actualizar Monitor de Impresoras
cd /d C:\impresoras
echo.
echo === 1. Actualizando codigo desde GitHub ===
git pull
if %errorlevel% neq 0 (
    echo Error en git pull. Abortando.
    pause
    exit /b 1
)
echo.
echo === 2. Instalando dependencias ===
call .venv\Scripts\activate
pip install -r requirements.txt
echo.
echo === 3. Compilando .exe ===
pyinstaller MonitorImpresoras.spec
if %errorlevel% neq 0 (
    echo Error al compilar. Abortando.
    pause
    exit /b 1
)
echo.
echo === 4. Copiando al servidor ===
copy /Y "C:\impresoras\dist\MonitorImpresoras.exe" "\\MXL8372J8P\impresoras\dist\"
if %errorlevel% equ 0 (
    echo Copiado exitoso a \\MXL8372J8P\impresoras\dist\
) else (
    echo ERROR: No se pudo copiar al servidor. Verificar conexion.
)
echo.
echo === Listo! ===
pause
