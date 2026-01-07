# QIP 檢驗報告數據提取工具

純前端 Excel 檢驗報告數據分析工具，所有處理在瀏覽器本地完成，確保資料隱私。

## 🔒 隱私保障

- **純前端處理**：所有數據處理在您的瀏覽器本地完成
- **無伺服器上傳**：檔案不會上傳至任何伺服器
- **無網路請求**：處理過程不發送任何網路請求
- **本地存儲**：配置僅存於您的設備（localStorage）

## 📋 功能特點

- 支援 Excel 檔案拖放上傳
- 多模穴配置（8/16/24/32/40/48 穴）
- 視覺化範圍選擇或手動輸入
- 自動提取規格數據
- 配置保存與載入
- Excel 結果匯出

## 🚀 使用方法

1. **上傳檔案**：拖放或點擊選擇 Excel 檔案
2. **設定參數**：選擇模穴數，設定數據範圍
3. **開始處理**：點擊「開始處理」按鈕
4. **下載結果**：處理完成後下載 Excel 結果

## 🛠️ 技術架構

- **前端框架**：純 HTML/CSS/JavaScript
- **Excel 解析**：[SheetJS](https://sheetjs.com/)
- **部署**：GitHub Pages

## 📦 檔案結構

```
docs/
├── index.html          # 主頁面
├── css/
│   └── style.css       # 樣式表
├── js/
│   ├── lib/
│   │   └── xlsx.full.min.js  # SheetJS 庫
│   ├── app.js          # 主應用邏輯
│   ├── processor.js    # 數據處理核心
│   ├── data-extractor.js
│   ├── data-validator.js
│   ├── spec-extractor.js
│   ├── error-logger.js
│   └── excel-exporter.js
└── .nojekyll           # 禁用 Jekyll
```

## 📝 版本資訊

- **版本**：1.0.0
- **原始版本**：VBA QIP_DataExtract

## 📄 授權

MIT License
