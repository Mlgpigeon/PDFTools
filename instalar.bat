@echo off
chcp 65001 >nul
title PDF Toolkit — Instalador de dependencias
echo.
echo ============================================
echo   PDF Toolkit — Instalador de dependencias
echo ============================================
echo.

:: Verificar que Python existe
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] No se encontro Python en el PATH.
    echo Descargalo de https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" al instalar.
    pause
    exit /b 1
)

:: Crear entorno virtual si no existe
if not exist "venv" (
    echo [1/4] Creando entorno virtual...
    python -m venv venv
) else (
    echo [1/4] Entorno virtual ya existe, reutilizando...
)

:: Activar entorno virtual
echo [2/4] Activando entorno virtual...
call venv\Scripts\activate.bat

echo.
echo [3/4] Instalando dependencias obligatorias...
pip install --upgrade pip
pip install pypdf[crypto] Pillow

echo.
echo [4/4] Instalando dependencias opcionales...
pip install tkinterdnd2 docx2pdf reportlab

echo.
echo ============================================
echo   Instalacion completada.
echo ============================================
echo.
echo Dependencias externas opcionales (instalar manualmente):
echo   - Ghostscript: https://ghostscript.com/releases/gsdnld.html
echo     (necesario para comprimir PDFs con maxima calidad)
echo   - LibreOffice: https://www.libreoffice.org/download/
echo     (alternativa a Word para convertir DOCX a PDF)
echo.
echo Para ejecutar:
echo   venv\Scripts\python pdf_toolkit.py
echo.
echo O activa el venv primero:
echo   venv\Scripts\activate
echo   python pdf_toolkit.py
echo.
pause
