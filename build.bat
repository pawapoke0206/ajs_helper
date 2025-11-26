@echo off
echo ========================================
echo   AST (AJS Support Tool) - Build
echo   onefile + Defender protection
echo ========================================
echo.

REM --- Clean previous build ---
if exist "dist\AST.exe" (
    echo [1/4] Cleaning previous build...
    del /f "dist\AST.exe"
)
if exist "build\AST" (
    rmdir /s /q "build\AST"
)

REM --- Rebuild bootloader from source ---
echo.
echo [2/4] Rebuilding PyInstaller bootloader from source...
echo       (Falls back to prebuilt if no C++ compiler found)
echo.
pip install --force-reinstall --no-binary pyinstaller pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [WARN] Source build failed. Using prebuilt PyInstaller.
    echo        See BUILD_README.md if Defender still blocks the exe.
    pip install pyinstaller >nul 2>&1
)

REM --- Build exe (onefile + no UPX) ---
echo.
echo [3/4] Building with PyInstaller... (onefile / no UPX)
echo.
pyinstaller build.spec --noconfirm --clean

REM --- Copy data files next to exe ---
echo.
echo [4/4] Copying data files...
if not exist "dist" mkdir "dist"
if exist "AJS_trans.prm"      copy /y "AJS_trans.prm"      "dist\" >nul 2>&1
if exist "io_exceptions.json" copy /y "io_exceptions.json" "dist\" >nul 2>&1
if exist "config.json"        copy /y "config.json"        "dist\" >nul 2>&1
if exist "history.json"       copy /y "history.json"       "dist\" >nul 2>&1

echo.
echo ========================================
echo   Build complete: dist\AST.exe
echo.
echo   Data files copied to dist\ folder.
echo   Distribute the entire dist\ folder.
echo.
echo   Protection layers:
echo     [o] UPX compression disabled
echo     [o] Bootloader source build (attempted)
echo ========================================
pause
