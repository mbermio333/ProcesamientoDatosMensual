@echo off
setlocal enabledelayedexpansion

title Instalador del Sistema de Procesamiento de Mediciones
color 0A

echo ===============================================
echo    SISTEMA DE PROCESAMIENTO DE MEDICIONES
echo ===============================================
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

:: Lista de archivos y carpetas a copiar (ajusta segun tu proyecto)
set "FILES_TO_COPY=main.py main2.py main3.py SPECTRA_Filtro.py normalizarTV.py actualizarConfig.py"
set "FOLDERS_TO_COPY=Img SPECTRA_Filtrado"
set "CONFIG_FILES=config_procesamiento.json config_ocupacion.json config_umbrales_ciudades.json config_spectra.json config.json config_backup.json"

:: Copiar archivos principales
for %%f in (%FILES_TO_COPY%) do (
    if exist "%%f" (
        copy "%%f" "%INSTALL_DIR%" >nul
        echo [✓] Copiado: %%f
    ) else (
        echo [ADVERTENCIA] No se encontro: %%f
    )
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

:: Copiar carpetas
for %%d in (%FOLDERS_TO_COPY%) do (
    if exist "%%d" (
        xcopy "%%d" "%INSTALL_DIR%\%%d" /E /I /Y >nul
        echo [✓] Copiada carpeta: %%d
    ) else (
        echo [ADVERTENCIA] No se encontro carpeta: %%d
    )
)

:: Copiar archivos GUI si existen
if exist "gui.py" (
    copy "gui.py" "%INSTALL_DIR%" >nul
    echo [✓] Copiado: gui.py
)


if exist "iconos" (
    xcopy "iconos" "%INSTALL_DIR%\iconos" /E /I /Y >nul
    echo [✓] Copiada carpeta: iconos
)

:: Instalar dependencias Python
echo.
echo Instalando dependencias de Python...

:: Crear archivo requirements.txt si no existe
if not exist "%REQUIREMENTS_FILE%" (
    echo Creando archivo de dependencias...
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
    ) > "%REQUIREMENTS_FILE%"
)

:: Instalar dependencias
pip install -r "%REQUIREMENTS_FILE%"

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
set "ICON_PATH=%INSTALL_DIR%\iconos\SACER.ico"  :: Ajusta la ruta del icono

:: Si no hay icono, usar el de Python por defecto
if not exist "%ICON_PATH%" (
    set "ICON_PATH=%INSTALL_DIR%\python.exe"
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

title Desinstalador - Sistema de Procesamiento de Mediciones
color 0C

set "APP_NAME=Sistema de Procesamiento de Mediciones"
set "INSTALL_DIR=%USERPROFILE%\SistemaProcesamientoMediciones"
set "DESKTOP_DIR=%USERPROFILE%\Desktop"
set "SHORTCUT_PATH=%DESKTOP_DIR%\%APP_NAME%.lnk"
set "REG_KEY=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%APP_NAME%"

echo ===============================================
echo    DESINSTALANDO %APP_NAME%
echo ===============================================
echo.
echo ADVERTENCIA: Esta accion eliminara permanentemente:
echo.
echo [✗] La aplicacion y todos sus archivos
echo [✗] Los accesos directos
echo [✗] Las configuraciones guardadas
echo.
echo [✓] Los archivos de mediciones y reportes NO seran eliminados
echo     (carpetas MedicionesFmCSV, MedicionesTvCSV, etc.)
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
if exist "!SHORTCUT_PATH!" (
    del "!SHORTCUT_PATH!"
    echo [✓] Acceso directo eliminado
) else (
    echo [i] Acceso directo no encontrado
)

:: Eliminar directorio de instalacion
if exist "!INSTALL_DIR!" (
    rmdir /s /q "!INSTALL_DIR!"
    echo [✓] Archivos de aplicacion eliminados
) else (
    echo [i] Directorio de instalacion no encontrado
)

:: Eliminar entrada del registro
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
echo %APP_NAME% ha sido desinstalado completamente.
echo.
echo Los siguientes elementos se conservan:
echo - Archivos de mediciones en las carpetas originales
echo - Archivos de reportes generados
echo - Python y las librerias instaladas
echo.
pause
) > "%UNINSTALL_BAT%"

:: Crear entrada en "Agregar o quitar programas" (opcional)
echo.
echo Configurando registro de Windows para desinstalacion...

set "REG_KEY=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%APP_NAME%"
reg add "%REG_KEY%" /v "DisplayName" /d "%APP_NAME%" /f >nul 2>&1
reg add "%REG_KEY%" /v "UninstallString" /d "\"%UNINSTALL_BAT%\"" /f >nul 2>&1
reg add "%REG_KEY%" /v "InstallLocation" /d "%INSTALL_DIR%" /f >nul 2>&1
reg add "%REG_KEY%" /v "Publisher" /d "MIGBermio" /f >nul 2>&1  :: Cambia por tu nombre/empresa
reg add "%REG_KEY%" /v "DisplayVersion" /d "1.0.0" /f >nul 2>&1  :: Cambia por tu version

:: Limpiar archivos temporarios
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
echo.

pause