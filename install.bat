@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================================
echo   PDF OCR 翻譯工具 - 一鍵安裝
echo ================================================
echo.

REM --- 檢查 Python ---
python --version >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 未找到 Python，請先安裝 Python 3.10 以上版本。
    echo 下載位址：https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)
for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo [OK] Python %PYVER%

REM --- 檢查 Tesseract OCR ---
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo [OK] Tesseract OCR 已安裝
) else (
    echo.
    echo ================================================
    echo   [注意] 未偵測到 Tesseract OCR
    echo ================================================
    echo.
    echo   掃描式 PDF（圖片式 PDF）需要 Tesseract 才能辨識文字。
    echo   原生文字 PDF 不需要，可略過。
    echo.
    echo   建議現在先安裝 Tesseract，再繼續：
    echo.
    echo   1. 前往下載頁面：
    echo      https://github.com/UB-Mannheim/tesseract/wiki
    echo.
    echo   2. 安裝時在語言清單勾選以下項目：
    echo      - chi_tra  繁體中文
    echo      - chi_sim  簡體中文
    echo      - jpn      日文（選用）
    echo      - kor      韓文（選用）
    echo      （英文 eng 預設已包含）
    echo.
    echo   3. 安裝完成後，重新執行此 install.bat
    echo.
    echo   若只翻譯原生文字 PDF，可直接按任意鍵略過繼續安裝。
    echo.
    pause
)

REM --- 建立虛擬環境 ---
if not exist "venv\" (
    echo 建立 Python 虛擬環境...
    python -m venv venv
    if errorlevel 1 (
        echo [錯誤] 虛擬環境建立失敗。
        pause
        exit /b 1
    )
    echo [OK] 虛擬環境建立完成
) else (
    echo [OK] 虛擬環境已存在，略過建立
)

REM --- 安裝套件 ---
echo.
echo 安裝 Python 套件（可能需要 1-3 分鐘）...
call venv\Scripts\activate
pip install --upgrade pip -q
pip install -r requirements.txt
if errorlevel 1 (
    echo [錯誤] 套件安裝失敗，請確認網路連線後重試。
    pause
    exit /b 1
)

echo.
echo ================================================
echo   安裝完成！
echo ================================================
echo.
echo 使用方式：
echo   1. 雙擊「start.bat」開啟程式
echo   2. 在網頁介面中輸入您的 Gemini API Key
echo      （取得方式：https://aistudio.google.com/apikey）
echo   3. 上傳 PDF 並開始翻譯
echo.
pause
