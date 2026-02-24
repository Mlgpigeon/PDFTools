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

echo [1/3] Actualizando pip...
python -m pip install --upgrade pip

echo.
echo [2/3] Instalando dependencias obligatorias...
python -m pip install pypdf[crypto] Pillow

echo.
echo [3/3] Instalando dependencias opcionales...
python -m pip install tkinterdnd2 docx2pdf reportlab

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
echo Para ejecutar: python pdf_toolkit.py
echo.
pause
