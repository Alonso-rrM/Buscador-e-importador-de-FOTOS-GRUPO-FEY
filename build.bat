@echo off
echo.
echo ==================================================
echo   Compilando Unificador de Catalogos a .EXE
echo ==================================================
echo.

:: 1. Instalar PyInstaller en el entorno virtual
echo Paso 1: Verificando e instalando PyInstaller en el entorno virtual...
call .venv\Scripts\python.exe -m pip install pyinstaller

:: 2. Ejecutar PyInstaller para generar el .exe
echo.
echo Paso 2: Compilando la aplicacion a un archivo unico .exe...
call .venv\Scripts\pyinstaller --onefile --noconsole --name "UnificadorPDF" --clean gui.py

echo.
echo ==================================================
echo ¡Proceso completado!
echo El programa ejecutable se ha creado en:
echo   dist\UnificadorPDF.exe
echo.
echo Puedes llevar ese archivo a cualquier otra PC.
echo ==================================================
pause
