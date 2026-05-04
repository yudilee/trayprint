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
REM
REM ============================================================================

setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0
set PROJECT_DIR=%SCRIPT_DIR%..\
set WIX_TOOLSET_PATH=

REM ---- Locate WiX candle.exe ----
where candle.exe >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set CANDLE=candle.exe
    set LIGHT=light.exe
) else if defined WIX_TOOLSET_PATH (
    set CANDLE="%WIX_TOOLSET_PATH%\candle.exe"
    set LIGHT="%WIX_TOOLSET_PATH%\light.exe"
) else (
    echo [ERROR] WiX Toolset not found in PATH.
    echo         Please install WiX v3.14+ from https://wixtoolset.org
    echo         or set the WIX_TOOLSET_PATH environment variable.
    exit /b 1
)

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

REM ---- Build MSI ----
echo ===========================================================================
echo  TrayPrint — MSI Installer Build
echo  Project:  %PROJECT_DIR%
echo  WiX:      %CANDLE% / %LIGHT%
echo ===========================================================================
echo.

pushd "%SCRIPT_DIR%"

echo [1/2] Compiling WiX source...
%CANDLE% product.wxs -out product.wixobj
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] WiX compilation failed (candle.exe exit code %ERRORLEVEL%).
    popd
    exit /b 1
)

echo [2/2] Linking MSI package...
%LIGHT% -out "TrayPrint.msi" product.wixobj
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] WiX linking failed (light.exe exit code %ERRORLEVEL%).
    popd
    exit /b 1
)

echo.
echo ===========================================================================
echo  BUILD SUCCESS!
echo  Installer: %SCRIPT_DIR%TrayPrint.msi
echo ===========================================================================

popd
endlocal
exit /b 0
