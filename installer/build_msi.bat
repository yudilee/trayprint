@echo off
REM ============================================================================
REM  TrayPrint — MSI Installer Build Script
REM ============================================================================
REM
REM  This script builds the TrayPrint MSI installer using WiX Toolset v3.x.
REM
REM  Prerequisites:
REM    1. WiX Toolset v3.14+ installed (https://wixtoolset.org)
REM       - Ensure candle.exe and light.exe are in your PATH, or
REM         set WIX_TOOLSET_PATH below.
REM    2. PyInstaller build must have been run first:
REM         python build.py
REM       Ensure dist\trayprint.exe exists.
REM    3. trayprint.ico must exist in the project root.
REM    4. config.json must exist in the project root.
REM    5. PowerShell scripts must exist in the installer directory:
REM         - Register-TrayPrintTask.ps1
REM         - Unregister-TrayPrintTask.ps1
REM
REM ============================================================================

setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0
set PROJECT_DIR=%SCRIPT_DIR%..\
set WIX_TOOLSET_PATH=

REM ---- Locate WiX candle.exe ----
where candle.exe >nul 2>&1
if not errorlevel 1 (
    set CANDLE=candle.exe
    set LIGHT=light.exe
) else if defined WIX_TOOLSET_PATH (
    set CANDLE="%WIX_TOOLSET_PATH%\candle.exe"
    set LIGHT="%WIX_TOOLSET_PATH%\light.exe"
) else if defined WIX (
    set CANDLE="%WIX%bin\candle.exe"
    set LIGHT="%WIX%bin\light.exe"
) else (
    echo [ERROR] WiX Toolset not found in PATH.
    echo         Please install WiX v3.14+ from https://wixtoolset.org
    echo         or set the WIX_TOOLSET_PATH environment variable.
    exit /b 1
)
echo CANDLE=%CANDLE%
echo LIGHT=%LIGHT%

REM ---- Check prerequisites ----
if not exist "%PROJECT_DIR%dist\trayprint.exe" (
    echo [ERROR] dist\trayprint.exe not found.
    echo         Run 'python build.py' first to create the PyInstaller executable.
    exit /b 1
)

if not exist "%PROJECT_DIR%trayprint.ico" (
    echo [WARNING] trayprint.ico not found — installer will have no icon.
)

if not exist "%PROJECT_DIR%config.json" (
    echo [WARNING] config.json not found at project root.
    echo           A default config.json will be expected by the installer.
)

if not exist "%SCRIPT_DIR%Register-TrayPrintTask.ps1" (
    echo [ERROR] Register-TrayPrintTask.ps1 not found in installer directory.
    echo         This script is required to register the Scheduled Task.
    exit /b 1
)

if not exist "%SCRIPT_DIR%Unregister-TrayPrintTask.ps1" (
    echo [ERROR] Unregister-TrayPrintTask.ps1 not found in installer directory.
    echo         This script is required to unregister the Scheduled Task.
    exit /b 1
)

REM ---- Build MSI ----
echo ===========================================================================
echo  TrayPrint — MSI Installer Build
echo  Project:  %PROJECT_DIR%
echo  WiX:      %CANDLE% / %LIGHT%
echo ===========================================================================
echo.

pushd "%SCRIPT_DIR%"

echo [1/2] Compiling WiX source...
%CANDLE% -ext WixUtilExtension -dVersion=1.0.0 trayprint.wxs -out trayprint.wixobj
echo [candle rc=!errorlevel!]
if errorlevel 1 (
    echo [ERROR] WiX compilation failed (candle.exe exit code !errorlevel!).
    popd
    exit /b 1
)

echo [2/2] Linking MSI package...
%LIGHT% -ext WixUtilExtension -ext WixUIExtension -out "TrayPrint.msi" trayprint.wixobj
echo [light rc=!errorlevel!]
if errorlevel 1 (
    echo [ERROR] WiX linking failed (light.exe exit code !errorlevel!).
    popd
    exit /b 1
)

echo.
echo ===========================================================================
echo  BUILD SUCCESS!
echo  Installer: %SCRIPT_DIR%TrayPrint.msi
echo ===========================================================================

REM Salin ke dist\ supaya masuk artifact workflow (upload glob dist/*.msi)
if exist "%SCRIPT_DIR%TrayPrint.msi" (
    copy /y "%SCRIPT_DIR%TrayPrint.msi" "%PROJECT_DIR%dist\TrayPrint.msi" >nul
    echo Copied MSI to: %PROJECT_DIR%dist\TrayPrint.msi
) else (
    echo [WARNING] TrayPrint.msi tidak ditemukan utk disalin ke dist\
)

popd
endlocal
exit /b 0
