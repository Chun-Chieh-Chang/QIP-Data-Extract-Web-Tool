# 📁 QIP Data Extract Web Tool - 專案結構說明

## 🎯 專案概述
本專案是一個基於網頁的 QIP 數據提取工具,功能等同於原本的 VBA 巨集,但無需安裝 Excel 即可使用。

---

## 📂 目錄結構 (MECE 原則)

```
QIP_DataExtract/
├── 📄 README.md                    # 專案說明與使用指南
├── 📄 PROJECT_STRUCTURE.md         # 本檔案 - 專案結構詳細說明
├── 📄 .gitignore                   # Git 忽略規則
│
├── 📁 docs/                        # 🌐 Web 應用程式 (GitHub Pages 發布來源)
│   ├── 📄 index.html               # 主頁面 - UI 結構
│   ├── 📄 README.md                # Web 應用說明
│   ├── 📄 .nojekyll                # GitHub Pages 配置
│   │
│   ├── 📁 css/                     # 樣式表
│   │   └── 📄 style.css            # 主樣式檔案
│   │
│   └── 📁 js/                      # JavaScript 模組
│       ├── 📄 app.js               # 主應用程式 - UI 交互與流程控制
│       ├── 📄 processor.js         # 核心處理器 - 批號合併與數據處理邏輯
│       ├── 📄 data-extractor.js    # 數據提取模組
│       ├── 📄 spec-extractor.js    # 規格提取模組
│       ├── 📄 data-validator.js    # 數據驗證模組
│       ├── 📄 error-logger.js      # 錯誤日誌模組
│       ├── 📄 excel-exporter.js    # Excel 輸出模組
│       └── 📁 lib/                 # 第三方函式庫
│           └── 📄 xlsx.full.min.js # SheetJS - Excel 檔案處理
│
└── 📁 vba-reference/               # 🔧 VBA 參考代碼 (僅供參考,不參與 Web 運行)
    ├── 📄 theCode.bas              # 原始 VBA 完整代碼 (單體檔案)
    ├── 📄 DataExtractor.bas        # VBA 數據提取模組
    ├── 📄 DataValidator.bas        # VBA 數據驗證模組
    ├── 📄 SpecificationExtractor.bas # VBA 規格提取模組
    └── 📄 ErrorLogger.bas          # VBA 錯誤日誌模組
```

---

## 🧩 模組職責劃分 (MECE)

### 1️⃣ **Web 應用層** (`docs/`)
- **目的**: 提供完整的前端應用,可直接部署至 GitHub Pages
- **核心檔案**:
  - `index.html`: UI 結構與佈局
  - `css/style.css`: 視覺樣式
  - `js/app.js`: 使用者交互、檔案上傳、配置管理
  - `js/processor.js`: 核心業務邏輯 (批號合併、跨頁處理)

### 2️⃣ **業務邏輯層** (`docs/js/`)
各模組遵循 **單一職責原則**:

| 模組 | 職責 | 主要功能 |
|------|------|----------|
| `app.js` | UI 控制器 | 檔案上傳、範圍選擇、配置管理、進度顯示 |
| `processor.js` | 核心處理器 | 批號合併、跨頁數據整合、產品資訊提取 |
| `data-extractor.js` | 數據提取 | 從 Excel 工作表提取批號、穴號、數值數據 |
| `spec-extractor.js` | 規格提取 | 提取規格符號、公差、USL/LSL |
| `data-validator.js` | 數據驗證 | 驗證工作表格式、數據完整性 |
| `error-logger.js` | 錯誤處理 | 記錄處理錯誤、生成錯誤報告 |
| `excel-exporter.js` | 輸出生成 | 生成標準化 Excel 輸出檔案 |

### 3️⃣ **VBA 參考層** (`vba-reference/`)
- **目的**: 保留原始 VBA 邏輯作為參考與備份
- **注意**: 此目錄內容 **不參與 Web 應用運行**
- **包含**:
  - `theCode.bas`: 原始完整 VBA 代碼
  - 模組化 VBA 檔案 (與 JS 模組對應)

---

## 🔄 數據流程

```
使用者上傳 Excel
    ↓
app.js (檔案載入 & UI 交互)
    ↓
processor.js (核心處理邏輯)
    ├→ data-extractor.js (提取數據)
    ├→ spec-extractor.js (提取規格)
    ├→ data-validator.js (驗證數據)
    └→ error-logger.js (記錄錯誤)
    ↓
excel-exporter.js (生成輸出)
    ↓
下載結果 Excel 檔案
```

---

## 🚀 部署說明

### GitHub Pages 設定
1. Repository Settings → Pages
2. Source: **Deploy from a branch**
3. Branch: **main** / Folder: **/docs**
4. 訪問: `https://chun-chieh-chang.github.io/QIP-Data-Extract-Web-Tool/docs/`

### 本地開發
```bash
# 直接開啟 docs/index.html (需要本地伺服器以避免 CORS 問題)
# 推薦使用 VS Code Live Server 或 Python HTTP Server
python -m http.server 8000
# 然後訪問 http://localhost:8000/docs/
```

---

## 📝 維護指南

### 新增功能時
1. **UI 變更**: 修改 `docs/index.html` 和 `docs/css/style.css`
2. **業務邏輯**: 修改對應的 JS 模組 (遵循單一職責)
3. **新增模組**: 在 `docs/js/` 建立新檔案,並在 `index.html` 引入

### 測試流程
1. 本地測試: 使用 Live Server 測試完整流程
2. 提交前檢查: 確保所有 JS 模組正確引入
3. 推送後驗證: 檢查 GitHub Pages 部署狀態

---

## 📌 版本資訊
- **建立日期**: 2026-01-07
- **最後更新**: 2026-01-07
- **維護者**: Chun-Chieh Chang
