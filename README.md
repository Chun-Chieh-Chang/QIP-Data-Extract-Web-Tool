# QIP Data Extract Web Tool

> 🌐 **基於網頁的 QIP 數據提取工具** - 無需 Excel,跨平台運行,完整復刻 VBA 功能

[![GitHub Pages](https://img.shields.io/badge/Demo-Live-success?style=flat-square)](https://chun-chieh-chang.github.io/QIP-Data-Extract-Web-Tool/docs/)
[![License](https://img.shields.io/badge/License-MIT-blue?style=flat-square)](LICENSE)

---

## 🚀 快速開始

### 線上使用 (推薦)
**[🔗 點擊開啟工具](https://chun-chieh-chang.github.io/QIP-Data-Extract-Web-Tool/docs/)**

### 本地運行
```bash
# 克隆專案
git clone https://github.com/Chun-Chieh-Chang/QIP-Data-Extract-Web-Tool.git
cd QIP-Data-Extract-Web-Tool

# 啟動本地伺服器 (避免 CORS 問題)
python -m http.server 8000
# 訪問 http://localhost:8000/docs/
```

---

## ✨ 核心特性

### 🔒 **隱私優先**
- ✅ 100% 前端處理,數據不離開您的瀏覽器
- ✅ 無需上傳至伺服器,確保機密安全

### 🎯 **功能完整**
- ✅ **多頁面批號合併**: 正確處理跨工作表的批號數據 (Step Logic)
- ✅ **規格自動提取**: 自動抓取規格符號、公差、USL/LSL
- ✅ **產品資訊識別**: 提取產品名稱與測量單位
- ✅ **靈活配置**: 支援 8/16/24/32/40/48 穴等多種模穴排列

### 💎 **現代化體驗**
- 🖱️ 拖放上傳 Excel 檔案
- 👁️ 即時預覽工作表內容 (支援合併儲存格)
- 🎯 視覺化範圍選擇器
- 💾 配置保存與載入

### 📊 **標準化輸出**
- 產出的 Excel 格式與 VBA 版本完全一致
- 可直接用於 SPC 分析系統

---

## 📖 使用指南

### 基本流程
1. **上傳檔案** → 拖曳或選擇 QIP Excel 報表
2. **設定參數** → 選擇模穴數、穴號範圍、數據範圍
3. **開始處理** → 系統自動解析所有工作表
4. **下載結果** → 取得標準化 Excel 輸出

### 進階功能
- **範圍選擇**: 點擊 "選擇" 按鈕,在預覽表格中框選範圍
- **頁面偏移**: 設定跨頁數據的偏移量 (1=同頁, 2=下頁...)
- **配置管理**: 保存常用配置,快速載入

---

## 📁 專案結構

```
QIP_DataExtract/
├── docs/                    # 🌐 Web 應用 (GitHub Pages)
│   ├── index.html           # 主頁面
│   ├── css/style.css        # 樣式表
│   ├── js/                  # JavaScript 模組
│   │   ├── core/            # 核心邏輯 (MECE: Internal data handling)
│   │   │   ├── extractor.js
│   │   │   ├── spec-extractor.js
│   │   │   ├── processor.js
│   │   │   └── validator.js
│   │   ├── ui/              # 使用者介面 (MECE: Interface management)
│   │   │   ├── app.js
│   │   │   └── logger.js
│   │   ├── utils/           # 工具函式 (MECE: Reusable tools)
│   │   │   └── exporter.js
│   │   └── lib/             # 第三方庫
│   └── notes/               # 設計文件紀錄
└── reference/               # 🔧 原始參考代碼
```

📚 **詳細說明**: [docs/notes/PROJECT_STRUCTURE.md](docs/notes/PROJECT_STRUCTURE.md)

---

## 🛠 技術架構

### 前端技術
- **HTML5 + CSS3 + Tailwind**: 現代化 UI 與高效開發
- **Vanilla JavaScript**: 無框架依賴,輕量高效
- **SheetJS (xlsx.js)**: Excel 檔案處理

### 模組化設計 (MECE 原則)
| 類別 | 模組 | 職責 |
|------|------|------|
| **Core** | `extractor.js` | 數據提取邏輯 |
| **Core** | `spec-extractor.js` | 規格提取邏輯 |
| **Core** | `processor.js` | 業務流程編排 |
| **Core** | `validator.js` | 數據格式驗證 |
| **UI**   | `app.js` | UI 交互與事件綁定 |
| **UI**   | `logger.js` | 錯誤日誌與回饋 |
| **Utils**| `exporter.js` | Excel 格式導出 |

---

## 🤝 貢獻指南

歡迎提交 Issue 或 Pull Request!

### 開發流程
1. Fork 本專案
2. 建立功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交變更 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

---

## 📝 版本歷史

- **v1.0.0** (2026-01-07): 初始版本,完整功能實現

---

## 👤 作者

**Chun-Chieh Chang**

- GitHub: [@Chun-Chieh-Chang](https://github.com/Chun-Chieh-Chang)

---

## 📄 授權

本專案採用 MIT 授權 - 詳見 [LICENSE](LICENSE) 檔案

---

## 🙏 致謝

- [SheetJS](https://sheetjs.com/) - 強大的 Excel 處理函式庫
- 原始 VBA 巨集開發團隊

---

**⭐ 如果這個專案對您有幫助,請給個 Star!**
