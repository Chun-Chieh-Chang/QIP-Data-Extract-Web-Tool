# QIP Data Extract Web Tool

這是一個基於網頁的 QIP 數據提取工具，功能等同於原本的 VBA "QIP Data Extract" 巨集，但無需安裝 Excel 即可使用，且支援跨平台。

## 🚀 線上使用 (Live Tool)
**[點擊這裡開啟工具](https://chun-chieh-chang.github.io/QIP-Data-Extract-Web-Tool/docs/)**
*(若出現 404 錯誤，請確保在 GitHub Repository Settings > Pages 中將來源設定為 `/docs` 資料夾)*

## ✨ 主要功能 (Features)
- **純前端處理 (Client-Side)**：所有運算皆在您的瀏覽器中完成，數據不會上傳至伺服器，確保機密安全。
- **完整復刻 VBA 邏輯**：
  - ✅ **多頁面批號合併**：正確處理跨工作表的批號數據 (Step Logic)。
  - ✅ **規格與產品資訊提取**：自動抓取規格 (Specs)、產品名稱 (P2/P3/P2:V3) 與 測量單位 (W23)。
  - ✅ **靈活的版面配置**：支援 8穴、16穴、32穴等多種模穴排列設定。
- **現代化介面**：支援拖放上傳、即時預覽 (Preview)、範圍選擇。
- **標準化輸出**：產出的 Excel 格式與 VBA 版本一致，可直接用於 SPC 分析系統。

## 🛠 使用方法 (Usage)
1. 開啟 [線上工具](https://chun-chieh-chang.github.io/QIP-Data-Extract-Web-Tool/docs/)。
2. **上傳檔案**：拖曳或點擊選擇您的 QIP Excel 報表。
3. **設定參數**：
   - 選擇 Excel 中的數據範圍 (可使用 "選擇" 按鈕直接在預覽窗格中框選)。
   - 設定 "頁面偏移" (1=同頁, 2=下頁...) 以正確合併跨頁資料。
4. **開始處理**：點擊按鈕，系統將自動解析所有工作表。
5. **下載結果**：檢查無誤後，下載生成的 Excel 檔案。

## 📁 專案結構
- `docs/`：網頁應用程式源代碼 (HTML/JS/CSS)，也是 GitHub Pages 的發布來源。
- `*.bas`：原始 VBA 參考代碼 (備份用)。

---
*Created by Chun-Chieh Chang*
