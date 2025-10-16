@echo off
setlocal enabledelayedexpansion

:: --- CONFIGURACIÓN ---
set VENV_DIR=env
set PYTHON_VER=3.8.3
set PYTHON_DIR=python-%PYTHON_VER%
set PYTHON_LOCAL=%PYTHON_DIR%\python.exe
set PYTHON_EXE=
set DEST_DIR=%LOCALAPPDATA%\SACER_PROCESAMIENTO
set SHORTCUT_PATH=%USERPROFILE%\Desktop\SACER GUI.lnk

echo 📁 Carpeta de destino: %DEST_DIR%
echo 📦 Iniciando instalación...

:: --- BUSCAR PYTHON EN EL SISTEMA ---
for /f "delims=" %%i in ('python --version 2^>^&1') do set SYS_PYVER=%%i
echo 🧪 Buscando Python en el sistema...
echo 📌 Python del sistema: !SYS_PYVER!

echo !SYS_PYVER! | find "3.8.3" >nul
if %errorlevel%==0 (
    echo ✅ Se usará Python del sistema: python.exe
    set PYTHON_EXE=python
    goto :crear_entorno
)

:: --- BUSCAR PYTHON PORTÁTIL ---
if exist "%PYTHON_LOCAL%" (
    for /f "delims=" %%i in ('"%PYTHON_LOCAL%" --version 2^>^&1') do set LOCAL_PYVER=%%i
    echo 📌 Python portátil encontrado: !LOCAL_PYVER!
    echo !LOCAL_PYVER! | find "3.8.3" >nul
    if %errorlevel%==0 (
        echo ✅ Se usará Python portátil
        set PYTHON_EXE=%PYTHON_LOCAL%
        goto :crear_entorno
    ) else (
        echo ⚠ Python portátil no es la versión correcta. Reinstalando...
    )
) else (
    echo ❌ No se encontró Python 3.8.3 válido. Procediendo a descarga...
)

goto :descargar_python

:descargar_python
echo 🔽 Descargando Python %PYTHON_VER% portable...
curl -L -o python.zip https://www.python.org/ftp/python/%PYTHON_VER%-embed-amd64.zip
mkdir "%PYTHON_DIR%"
tar -xf python.zip -C "%PYTHON_DIR%"
del python.zip
echo ✅ Python descargado y extraído.
set PYTHON_EXE=%PYTHON_LOCAL%

:crear_entorno
:: --- CREAR ENTORNO VIRTUAL ---
if exist "%VENV_DIR%\Scripts\activate.bat" (
    echo ✅ El entorno virtual ya existe. Saltando creación...
    goto :activar_entorno
)

echo 🔧 Creando entorno virtual...
%PYTHON_EXE% -m venv %VENV_DIR%

:activar_entorno
call %VENV_DIR%\Scripts\activate.bat

:: --- VERIFICAR requirements.txt ---
if not exist requirements.txt (
    echo ❌ No se encontró el archivo requirements.txt. Abortando instalación...
    pause
    exit /b
)

:: --- VERIFICAR PAQUETES INSTALADOS ---
echo 🧪 Verificando dependencias instaladas...
set "REQ_OK=1"
for %%p in (pandas openpyxl PyQt5 numpy Pillow) do (
    %VENV_DIR%\Scripts\python.exe -c "import %%p" >nul 2>&1
    if errorlevel 1 (
        set "REQ_OK=0"
    )
)

if !REQ_OK! == 1 (
    echo ✅ Todos los paquetes ya están instalados. Saltando instalación.
) else (
    echo 📦 Instalando dependencias...
    pip install --upgrade pip
    pip install -r requirements.txt
)

:: --- COPIAR ARCHIVOS A DESTINO ---
echo 📁 Copiando archivos a %DEST_DIR%...
mkdir "%DEST_DIR%"
xcopy * "%DEST_DIR%" /E /I /Y

:: --- CREAR lanzador.bat EN DESTINO ---
(
echo @echo off
echo cd /d %%~dp0
echo call env\Scripts\activate.bat
echo python gui.py
) > "%DEST_DIR%\lanzador.bat"

:: --- CREAR ACCESO DIRECTO EN ESCRITORIO ---
echo 🔗 Creando acceso directo...
powershell -Command ^
  "$s=(New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT_PATH%');" ^
  "$s.TargetPath='%DEST_DIR%\lanzador.bat';" ^
  "$s.WorkingDirectory='%DEST_DIR%';" ^
  "$s.IconLocation='%DEST_DIR%\iconos\SACER.ico';" ^
  "$s.Save()"

:: --- CREAR UNINSTALLER ---
echo 🧹 Creando uninstaller...
(
echo @echo off
echo echo 🔄 Eliminando SACER...
echo rmdir /s /q "%DEST_DIR%"
echo del "%SHORTCUT_PATH%"
echo del %%~f0
echo echo 🗑 SACER eliminado correctamente.
echo pause
) > "%DEST_DIR%\uninstaller.bat"

echo ✅ Instalación completada con éxito.
pause
