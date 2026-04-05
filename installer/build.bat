@echo off
setlocal

set PUBLISH_DIR=..\ExcelMerge.GUI\bin\Release\net8.0-windows\win-x64\publish

echo ========================================
echo  ExcelMerge Release Build
echo ========================================

echo.
echo [1/3] Cleaning previous build...
if exist "%PUBLISH_DIR%" rd /s /q "%PUBLISH_DIR%"
if exist "output" rd /s /q "output"

echo.
echo [2/3] Publishing self-contained single-file...
dotnet publish ..\ExcelMerge.GUI\ExcelMerge.GUI.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
if errorlevel 1 (
    echo.
    echo ERROR: dotnet publish failed!
    pause
    exit /b 1
)

echo.
echo [3/3] Building installer...
where iscc >nul 2>&1
if errorlevel 1 (
    echo Inno Setup not found in PATH.
    echo Install from: https://jrsoftware.org/isinfo.php
    echo Then run: iscc ExcelMerge.iss
    echo.
    echo Alternatively, distribute the published files directly from:
    echo   %PUBLISH_DIR%
) else (
    iscc ExcelMerge.iss
    if errorlevel 1 (
        echo ERROR: Inno Setup compilation failed!
        pause
        exit /b 1
    )
    echo.
    echo Installer created in: output\
)

echo.
echo ========================================
echo  Build complete!
echo ========================================
echo.
echo Published files: %PUBLISH_DIR%
echo.
pause
