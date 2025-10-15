@echo off
setlocal enabledelayedexpansion

title Instalador del Sistema de Procesamiento de Mediciones
color 0A

echo ===============================================
echo    SISTEMA DE PROCESAMIENTO DE MEDICIONES
echo ===============================================
echo.

:: Obtener el directorio donde está el instalador
set "INSTALLER_DIR=%~dp0"
cd /d "!INSTALLER_DIR!"
echo Directorio del instalador: !CD!
echo.

:: Verificar que estamos en el directorio correcto buscando archivos con y sin extension
set "FOUND_MAIN=0"
if exist "main.py" set "FOUND_MAIN=1"
if exist "main" set "FOUND_MAIN=1"

if !FOUND_MAIN! equ 0 (
    echo [ERROR] No se encuentra main.py o main en el directorio actual
    echo.
    echo Archivos encontrados en el directorio:
    dir /b
    echo.
    echo Por favor, asegurese de que el instalador esta en la carpeta correcta.
    pause
    exit /b 1
)

echo [✓] Directorio del proyecto verificado correctamente
echo.

:: Verificar si se ejecuta como administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ADVERTENCIA] No se estan ejecutando como administrador.
    echo Algunas funciones podrian requerir privilegios de administrador.
    echo.
    timeout /t 3 >nul
)

:: Configurar rutas
set "APP_NAME=Sistema de Generacion de Reportes - SACER"
set "INSTALL_DIR=%USERPROFILE%\SistemaGeneracionReportes"
set "PYTHON_URL=https://www.python.org/ftp/python/3.8.3/python-3.8.3-amd64.exe"
set "PYTHON_INSTALLER=python-3.8.3-installer.exe"
set "REQUIREMENTS_FILE=requirements.txt"

echo Verificando instalacion de Python 3.8.3...
echo.

:: Verificar si Python 3.8.3 esta instalado
python --version >nul 2>&1
if %errorLevel% equ 0 (
    python -c "import sys; print(sys.version)" | findstr "3.8.3" >nul
    if %errorLevel% equ 0 (
        echo [✓] Python 3.8.3 ya esta instalado.
        goto :INSTALL_APP
    ) else (
        echo [i] Se encontro Python, pero no es la version 3.8.3
        echo.
    )
)

echo [i] Python 3.8.3 no esta instalado o no se encuentra.
echo     Descargando e instalando Python 3.8.3...
echo.

:: Descargar Python 3.8.3
echo Descargando Python 3.8.3...
powershell -Command "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%'"

if not exist "%PYTHON_INSTALLER%" (
    echo [ERROR] No se pudo descargar Python 3.8.3
    echo Por favor, instale manualmente Python 3.8.3 desde:
    echo https://www.python.org/downloads/release/python-383/
    pause
    exit /b 1
)

echo [✓] Python 3.8.3 descargado correctamente.
echo.

:: Instalar Python 3.8.3
echo Instalando Python 3.8.3...
echo NOTA: En el instalador de Python, asegurese de marcar:
echo       [✓] Add Python 3.8 to PATH
echo.
echo Presione cualquier tecla para continuar con la instalacion...
pause >nul

start /wait "" "%PYTHON_INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0

:: Verificar instalacion
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] La instalacion de Python fallo.
    echo Por favor, instale manualmente Python 3.8.3
    pause
    exit /b 1
)

python -c "import sys; print(sys.version)" | findstr "3.8.3" >nul
if %errorLevel% neq 0 (
    echo [ERROR] La version instalada no es Python 3.8.3
    pause
    exit /b 1
)

echo [✓] Python 3.8.3 instalado correctamente.
echo.

:INSTALL_APP
echo ===============================================
echo      INSTALANDO LA APLICACION
echo ===============================================
echo.

:: Crear directorio de instalacion
if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
    echo [✓] Directorio de instalacion creado: %INSTALL_DIR%
) else (
    echo [i] El directorio ya existe. Se sobreescribiran los archivos.
)

:: Copiar archivos de la aplicacion
echo Copiando archivos de la aplicacion...

:: Lista de archivos principales (buscar con y sin extension .py)
set "FILE_NAMES=main main2 main3 SPECTRA_Filtro normalizarTV actualizarConfig gui"
set "FOLDERS_TO_COPY=img SPECTRA_Filtrado iconos"
set "CONFIG_FILES=config_processamiento.json config_ocupacion.json config_unbrales_ciudades.json config_spectra.json config.json config_backup.json"
set "OTHER_FILES=.gitignore"

:: Funcion para copiar archivos con o sin extension
set "COPIED_FILES=0"

for %%f in (%FILE_NAMES%) do (
    :: Buscar archivo con extension .py primero
    if exist "%%f.py" (
        copy "%%f.py" "%INSTALL_DIR%" >nul
        echo [✓] Copiado: %%f.py
        set /a COPIED_FILES+=1
    ) else if exist "%%f" (
        :: Si no tiene extension, copiar como esta y agregar extension .py
        copy "%%f" "%INSTALL_DIR%\%%f.py" >nul
        echo [✓] Copiado: %%f (renombrado a %%f.py)
        set /a COPIED_FILES+=1
    ) else (
        echo [ADVERTENCIA] No se encontro: %%f
    )
)

if !COPIED_FILES! equ 0 (
    echo [ERROR] No se pudieron copiar los archivos principales
    pause
    exit /b 1
)

:: Copiar archivos de configuracion
for %%f in (%CONFIG_FILES%) do (
    if exist "%%f" (
        copy "%%f" "%INSTALL_DIR%" >nul
        echo [✓] Copiado: %%f
    ) else (
        echo [ADVERTENCIA] No se encontro: %%f
    )
)

:: Copiar otros archivos
for %%f in (%OTHER_FILES%) do (
    if exist "%%f" (
        copy "%%f" "%INSTALL_DIR%" >nul
        echo [✓] Copiado: %%f
    )
)

:: Copiar carpetas
for %%d in (%FOLDERS_TO_COPY%) do (
    if exist "%%d" (
        xcopy "%%d" "%INSTALL_DIR%\%%d" /E /I /Y >nul
        echo [✓] Copiada carpeta: %%d
    ) else (
        echo [ADVERTENCIA] No se encontro carpeta: %%d
    )
)

:: Copiar archivo SACER.ico si existe
if exist "SACER.ico" (
    copy "SACER.ico" "%INSTALL_DIR%\iconos\" >nul 2>&1
    echo [✓] Copiado: SACER.ico a carpeta iconos
)

:: Crear y instalar dependencias Python
echo.
echo Creando e instalando dependencias de Python...

(
echo pandas
echo openpyxl
echo PyQt5
echo numpy
echo datetime
echo json
echo os
echo sys
echo subprocess
echo platform
) > "%INSTALL_DIR%\%REQUIREMENTS_FILE%"

echo [✓] Archivo %REQUIREMENTS_FILE% creado

:: Instalar dependencias
echo Instalando dependencias...
pip install -r "%INSTALL_DIR%\%REQUIREMENTS_FILE%"

if %errorLevel% equ 0 (
    echo [✓] Dependencias instaladas correctamente.
) else (
    echo [ADVERTENCIA] Hubo problemas instalando algunas dependencias.
)

:: Crear acceso directo en el escritorio
echo.
echo Creando acceso directo en el escritorio...

set "DESKTOP_DIR=%USERPROFILE%\Desktop"
set "SHORTCUT_PATH=%DESKTOP_DIR%\%APP_NAME%.lnk"
set "TARGET_PATH=%INSTALL_DIR%\gui.py"
set "ICON_PATH=%INSTALL_DIR%\iconos\SACER.ico"

:: Verificar si existe el icono, si no usar Python por defecto
if not exist "%ICON_PATH%" (
    for %%P in (python.exe) do set "ICON_PATH=%%~$PATH:P"
)

:: Crear script VBS para crear acceso directo
set "VBS_SCRIPT=%TEMP%\create_shortcut.vbs"

(
echo Set oWS = WScript.CreateObject^("WScript.Shell"^)
echo sLinkFile = "%SHORTCUT_PATH%"
echo Set oLink = oWS.CreateShortcut^(sLinkFile^)
echo oLink.TargetPath = "python"
echo oLink.Arguments = """"%TARGET_PATH%""""
echo oLink.WorkingDirectory = "%INSTALL_DIR%"
echo oLink.Description = "%APP_NAME%"
echo oLink.IconLocation = "%ICON_PATH%"
echo oLink.Save
) > "%VBS_SCRIPT%"

cscript //nologo "%VBS_SCRIPT%"

if exist "%SHORTCUT_PATH%" (
    echo [✓] Acceso directo creado en el escritorio.
) else (
    echo [ADVERTENCIA] No se pudo crear el acceso directo.
)

:: Crear script de desinstalacion
echo.
echo Creando script de desinstalacion...

set "UNINSTALL_BAT=%INSTALL_DIR%\uninstall.bat"

(
@echo off
setlocal enabledelayedexpansion

title Desinstalador - Sistema de Generacion de Reportes - SACER
color 0C

echo ===============================================
echo    DESINSTALANDO Sistema de Generacion de Reportes - SACER
echo ===============================================
echo.
echo ADVERTENCIA: Esta accion eliminara permanentemente:
echo.
echo [✗] La aplicacion y todos sus archivos
echo [✗] Los accesos directos
echo [✗] Las configuraciones guardadas
echo.
echo [✓] Los archivos de mediciones y reportes NO seran eliminados
echo.
set /p confirm=¿Esta seguro que desea continuar? [s/N]: 

if /i not "!confirm!"=="s" (
    echo.
    echo Desinstalacion cancelada.
    timeout /t 3 >nul
    exit /b 0
)

echo.
echo Iniciando desinstalacion...
echo.

:: Eliminar acceso directo
set "DESKTOP_SHORTCUT=%USERPROFILE%\Desktop\Sistema de Generacion de Reportes - SACER.lnk"
if exist "!DESKTOP_SHORTCUT!" (
    del "!DESKTOP_SHORTCUT!"
    echo [✓] Acceso directo eliminado
) else (
    echo [i] Acceso directo no encontrado
)

:: Eliminar directorio de instalacion
set "INSTALL_DIR=%USERPROFILE%\SistemaGeneracionReportes"
if exist "!INSTALL_DIR!" (
    rmdir /s /q "!INSTALL_DIR!"
    echo [✓] Archivos de aplicacion eliminados
) else (
    echo [i] Directorio de instalacion no encontrado
)

:: Eliminar entrada del registro
set "REG_KEY=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Sistema de Generacion de Reportes - SACER"
reg delete "!REG_KEY!" /f >nul 2>&1
if !errorLevel! equ 0 (
    echo [✓] Entrada de registro eliminada
) else (
    echo [i] Entrada de registro no encontrada
)

echo.
echo ===============================================
echo    DESINSTALACION COMPLETADA
echo ===============================================
echo.
echo La aplicacion ha sido desinstalado completamente.
echo.
pause
) > "%UNINSTALL_BAT%"

:: Crear entrada en "Agregar o quitar programas"
echo.
echo Configurando registro de Windows para desinstalacion...

set "REG_KEY=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Sistema de Generacion de Reportes - SACER"
reg add "%REG_KEY%" /v "DisplayName" /d "%APP_NAME%" /f >nul 2>&1
reg add "%REG_KEY%" /v "UninstallString" /d "\"%UNINSTALL_BAT%\"" /f >nul 2>&1
reg add "%REG_KEY%" /v "InstallLocation" /d "%INSTALL_DIR%" /f >nul 2>&1
reg add "%REG_KEY%" /v "Publisher" /d "MIGBermio" /f >nul 2>&1
reg add "%REG_KEY%" /v "DisplayVersion" /d "1.0.0" /f >nul 2>&1
reg add "%REG_KEY%" /v "NoModify" /d "1" /f >nul 2>&1
reg add "%REG_KEY%" /v "NoRepair" /d "1" /f >nul 2>&1

echo [✓] Entrada de registro creada para desinstalacion

:: Limpiar archivos temporales
del "%PYTHON_INSTALLER%" 2>nul
del "%VBS_SCRIPT%" 2>nul

echo.
echo ===============================================
echo      INSTALACION COMPLETADA
echo ===============================================
echo.
echo [✓] %APP_NAME% ha sido instalado correctamente.
echo.
echo Ubicacion de instalacion: %INSTALL_DIR%
echo Acceso directo creado en el escritorio.
echo.
echo Para ejecutar la aplicacion:
echo 1. Use el acceso directo en el escritorio
echo 2. O navegue a %INSTALL_DIR% y ejecute 'python gui.py'
echo.
echo Para desinstalar, ejecute: %UNINSTALL_BAT%
echo o use 'Agregar o quitar programas' en Windows
echo.

pause