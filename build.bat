@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ================================================
echo   ButterLayer - Build EXE
echo ================================================
echo.

if not exist "venv\Scripts\activate.bat" (
    echo [ERROR] venv not found. Please run install.bat first.
    pause
    exit /b 1
)

call venv\Scripts\activate

echo Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    pip install pyinstaller -q
)

echo [OK] PyInstaller ready
echo.
echo Start building...
echo.

pyinstaller pdf_ocr_translator.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

set "TESS_SRC=C:\Program Files\Tesseract-OCR"
set "TESS_DST=dist\ButterLayer\tesseract"

if exist "%TESS_SRC%\tesseract.exe" (
    echo.
    echo Copying Tesseract OCR...
    xcopy /E /I /Y "%TESS_SRC%" "%TESS_DST%" >nul
    echo [OK] Tesseract copied
) else (
    echo.
    echo [WARNING] Tesseract not found.
)

echo.
echo Copying instruction file...
copy /Y "ButterLayer Operation Instruction.txt" "dist\ButterLayer\" >nul

echo.
echo ================================================
echo   Build completed!
echo ================================================
echo.
echo Output folder: dist\ButterLayer\
echo.
pause
