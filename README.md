# PDF OCR 翻譯工具

上傳 PDF → 自動 OCR → Gemini AI 翻譯 → 網頁即時檢視 + 下載 Word 文件

---

## 功能一覽

| 功能 | 說明 |
|------|------|
| 智慧 OCR | 自動判斷：原生文字 PDF 直接擷取；掃描圖像 PDF 使用 Tesseract OCR |
| AI 翻譯 | 使用 Google Gemini 2.5 翻譯，支援繁體中文、簡體中文、英文等多語言 |
| 翻譯覆蓋 | 左側 PDF 可疊加翻譯文字，直接對照原文位置 |
| 字詞查詢 | 點擊 PDF 上任意字詞，即時顯示該詞意義 |
| 編輯翻譯 | 右側翻譯可直接修改，插入圖片（支援拖放、複製貼上），自動儲存 |
| 圖片調整 | 已插入的圖片可拖曳右下角調整大小 |
| 下載 Word | 三種格式：雙欄對照 / 純翻譯閱讀版 / 原 PDF 圖片 + OCR 文字 |

---

## 方案一：EXE 版（推薦給朋友使用）

打包後不需要安裝 Python，解壓縮即可使用。

### 開發者：打包步驟

1. 先完成下方「方案二」的安裝步驟，確保虛擬環境與套件就緒
2. 若需要 OCR 功能，先安裝 Tesseract（詳見下方）
3. 雙擊 `build.bat`，等待約 5–10 分鐘
4. 完成後輸出在 `dist\PDF翻譯工具\` 資料夾

### 朋友：使用步驟

1. 解壓縮收到的 ZIP 檔
2. 雙擊 `Pdf_ocr_translator.exe`，瀏覽器會自動開啟
3. 在網頁輸入自己的 Gemini API Key（見下方取得方式）
4. 上傳 PDF 開始翻譯

> 無需安裝任何軟體，開箱即用。

---

## 方案二：原始碼版（開發 / 自用）

### 前置需求

1. **Python 3.10+**
   下載：https://www.python.org/downloads/
   安裝時勾選「Add Python to PATH」

2. **Tesseract OCR**（掃描式 PDF 才需要）
   下載：https://github.com/UB-Mannheim/tesseract/wiki
   安裝時在語言清單勾選 `chi_tra`（繁體中文）、`chi_sim`（簡體中文）

### 安裝

```
install.bat
```

### 啟動

```
start.bat
```

瀏覽器會自動開啟 http://127.0.0.1:5000

---

## 取得 Gemini API Key

本工具使用 Google Gemini AI 進行翻譯，**每位使用者需要自己的免費 API Key**。

1. 前往 https://aistudio.google.com/apikey
2. 登入 Google 帳號
3. 點擊「建立 API 金鑰」
4. 複製金鑰，貼入網頁介面的「Gemini API Key」欄位

> API Key 儲存在本機瀏覽器，不會傳送至任何伺服器。

---

## 使用說明

1. 開啟程式後在「Gemini API Key」欄位輸入金鑰
2. 選擇目標翻譯語言（預設：繁體中文）
3. 拖曳或點擊上傳 PDF 檔案
4. 點擊「開始處理」，等待進度完成
5. 右側可直接編輯翻譯內容
6. 左側點擊「翻譯覆蓋」可將翻譯疊加在 PDF 原文上
7. 點擊 PDF 上的字詞可查看詞義
8. 點擊「下載」選擇 Word 格式

---

## 專案結構

```
pdf_ocr_translator/
├── app.py                    # Flask 後端
├── requirements.txt          # Python 套件清單
├── pdf_ocr_translator.spec   # PyInstaller 打包設定
├── build.bat                 # 打包成 EXE
├── install.bat               # 一鍵安裝（原始碼版）
├── start.bat                 # 啟動（原始碼版）
├── README.md
└── templates/
    └── index.html            # 網頁前端
```

---

## 注意事項

- OCR 品質取決於 PDF 掃描解析度，建議原稿解析度 ≥ 150 DPI
- 每次翻譯會呼叫 Gemini API，免費方案有每分鐘請求次數限制
- 本工具僅供本機使用，不建議部署至公開網路
