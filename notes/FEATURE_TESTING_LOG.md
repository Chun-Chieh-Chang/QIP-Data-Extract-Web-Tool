# 功能測試紀錄 - 儲存格映射協定自動確認

**功能名稱**: 儲存格映射協定自動確認機制  
**版本**: v2.4.0  
**測試日期**: 2026-02-03  
**測試人員**: Development Team  
**狀態**: 進行中

---

## 功能需求

### 原始需求
- 在點選資料範圍(第2點點選完畢)後即完成確認
- 不必再點選預覽區的"確認選取範圍"按鈕
- 若要修改，只需重新點選儲存格映射協定即可

### 實現方案
1. 自動確認選擇：第二次點擊儲存格後自動調用 `confirmSelection()`
2. 自動隱藏預覽區：選擇完成後 500ms 隱藏預覽區
3. 重新修改機制：重新點擊同一按鈕可重置選擇狀態
4. 隱藏確認按鈕：移除"確認選取範圍"和"取消選取"按鈕

---

## 代碼修改清單

### 1. js/ui/app.js - handleCellClick 函數
**修改位置**: 第 600-620 行  
**修改內容**:
- 第二次點擊後添加 `setTimeout(() => { confirmSelection(); }, 300);`
- 自動確認選擇，無需手動點擊

**修改前**:
```javascript
} else {
    // 第二次點擊 - 設置終點
    selectionEnd = { row, col };
    highlightSelection(selectionStart, selectionEnd);
    elements.confirmSelection.disabled = false;
}
```

**修改後**:
```javascript
} else {
    // 第二次點擊 - 設置終點
    selectionEnd = { row, col };
    highlightSelection(selectionStart, selectionEnd);
    elements.confirmSelection.disabled = false;
    
    // 自動確認選擇 (完成第2點後即確認)
    setTimeout(() => {
        confirmSelection();
    }, 300);
}
```

### 2. js/ui/app.js - cancelSelection 函數
**修改位置**: 第 683-700 行  
**修改內容**:
- 添加自動隱藏預覽區的邏輯
- 選擇完成後 500ms 隱藏預覽區

**修改前**:
```javascript
function cancelSelection() {
    selectionMode = null;
    selectionTarget = null;
    selectionStart = null;
    selectionEnd = null;
    elements.selectionModeText.textContent = '無';
    elements.confirmSelection.disabled = true;
    elements.previewTable.querySelectorAll('.selected, .selection-start').forEach(el => {
        el.classList.remove('selected', 'selection-start');
    });
}
```

**修改後**:
```javascript
function cancelSelection() {
    selectionMode = null;
    selectionTarget = null;
    selectionStart = null;
    selectionEnd = null;
    elements.selectionModeText.textContent = '無';
    elements.confirmSelection.disabled = true;
    elements.previewTable.querySelectorAll('.selected, .selection-start').forEach(el => {
        el.classList.remove('selected', 'selection-start');
    });
    
    // 自動隱藏預覽區 (完成選擇後自動收起)
    setTimeout(() => {
        elements.previewSection.style.display = 'none';
    }, 500);
}
```

### 3. js/ui/app.js - startRangeSelection 函數
**修改位置**: 第 560-590 行  
**修改內容**:
- 添加重新點擊檢測邏輯
- 允許用戶重新選擇範圍

**修改前**:
```javascript
function startRangeSelection(btn) {
    if (!currentWorkbook || !elements.worksheetSelect.value) {
        alert('請先上傳檔案並選擇工作表');
        return;
    }
    
    selectionMode = btn.dataset.type;
    selectionTarget = btn.dataset.target;
    selectionStart = null;
    selectionEnd = null;
    // ... 其他代碼
}
```

**修改後**:
```javascript
function startRangeSelection(btn) {
    if (!currentWorkbook || !elements.worksheetSelect.value) {
        alert('請先上傳檔案並選擇工作表');
        return;
    }
    
    // 如果已經在選擇模式中，重新點擊則重置
    if (selectionMode === btn.dataset.type && selectionTarget === btn.dataset.target) {
        // 重置選擇狀態
        selectionStart = null;
        selectionEnd = null;
        elements.previewTable.querySelectorAll('.selected, .selection-start').forEach(el => {
            el.classList.remove('selected', 'selection-start');
        });
        elements.confirmSelection.disabled = true;
        return;
    }
    
    selectionMode = btn.dataset.type;
    selectionTarget = btn.dataset.target;
    selectionStart = null;
    selectionEnd = null;
    // ... 其他代碼
}
```

### 4. index.html - 確認按鈕隱藏
**修改位置**: 第 750-760 行  
**修改內容**:
- 添加 `hidden` class 隱藏確認和取消按鈕

**修改前**:
```html
<button id="confirm-selection" disabled
    class="px-8 py-4 bg-emerald-500 hover:bg-emerald-600 text-white font-black text-sm uppercase tracking-widest rounded-2xl shadow-lg shadow-emerald-500/20 disabled:opacity-30 disabled:shadow-none transition-all">
    <span class="material-icons-round text-base mr-2">done_all</span> 確認選取範圍
</button>
<button id="cancel-selection"
    class="px-8 py-4 bg-slate-200 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-black text-sm uppercase tracking-widest rounded-2xl transition-all">
    取消選取
</button>
```

**修改後**:
```html
<button id="confirm-selection" disabled
    class="px-8 py-4 bg-emerald-500 hover:bg-emerald-600 text-white font-black text-sm uppercase tracking-widest rounded-2xl shadow-lg shadow-emerald-500/20 disabled:opacity-30 disabled:shadow-none transition-all hidden">
    <span class="material-icons-round text-base mr-2">done_all</span> 確認選取範圍
</button>
<button id="cancel-selection"
    class="px-8 py-4 bg-slate-200 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-black text-sm uppercase tracking-widest rounded-2xl transition-all hidden">
    取消選取
</button>
```

---

## 測試計畫

### 測試環境
- 瀏覽器: Chrome/Edge (最新版)
- 開發者工具: F12 Console
- 測試檔案: 標準 QIP Excel 檔案

### 測試用例

#### TC-001: 基本自動確認流程
**前置條件**:
- 已上傳 Excel 檔案
- 已選擇工作表
- 已設置模穴數

**測試步驟**:
1. 點擊儲存格映射協定按鈕 (ID 區域)
2. 在預覽表格中點擊第一個儲存格
3. 在預覽表格中點擊第二個儲存格

**預期結果**:
- ✓ 預覽區自動打開
- ✓ 第一次點擊標記起點 (selection-start class)
- ✓ 第二次點擊後自動確認
- ✓ 預覽區在 500ms 後自動隱藏
- ✓ 輸入框填入選擇的範圍 (如 A1:B2)
- ✓ Console 無錯誤

**測試狀態**: [ ] 待測 [ ] 進行中 [ ] 通過 [ ] 失敗

#### TC-002: 重新修改機制
**前置條件**:
- 已完成 TC-001
- 已自動確認一次選擇

**測試步驟**:
1. 重新點擊同一個儲存格映射協定按鈕
2. 在預覽表格中點擊不同的儲存格

**預期結果**:
- ✓ 預覽區重新打開
- ✓ 之前的選擇被清除
- ✓ 可以重新選擇新的範圍
- ✓ 新選擇自動確認
- ✓ 輸入框更新為新的範圍
- ✓ Console 無錯誤

**測試狀態**: [ ] 待測 [ ] 進行中 [ ] 通過 [ ] 失敗

#### TC-003: 多穴組選擇
**前置條件**:
- 已上傳 Excel 檔案
- 已選擇工作表
- 已設置模穴數為 16 或以上

**測試步驟**:
1. 為穴組 1 選擇 ID 區域
2. 為穴組 1 選擇數據區域
3. 為穴組 2 選擇 ID 區域
4. 為穴組 2 選擇數據區域

**預期結果**:
- ✓ 每次選擇都自動確認
- ✓ 所有輸入框都正確填入
- ✓ 預覽區在每次選擇後自動隱藏
- ✓ Console 無錯誤

**測試狀態**: [ ] 待測 [ ] 進行中 [ ] 通過 [ ] 失敗

#### TC-004: 邊界情況 - 單一儲存格選擇
**前置條件**:
- 已上傳 Excel 檔案
- 已選擇工作表

**測試步驟**:
1. 點擊儲存格映射協定按鈕
2. 點擊第一個儲存格
3. 再次點擊同一個儲存格

**預期結果**:
- ✓ 自動確認
- ✓ 輸入框填入單一儲存格地址 (如 A1)
- ✓ Console 無錯誤

**測試狀態**: [ ] 待測 [ ] 進行中 [ ] 通過 [ ] 失敗

#### TC-005: 工作表切換
**前置條件**:
- 已上傳多工作表的 Excel 檔案
- 已選擇工作表

**測試步驟**:
1. 點擊儲存格映射協定按鈕
2. 使用工作表切換按鈕切換工作表
3. 在新工作表中選擇儲存格

**預期結果**:
- ✓ 工作表正確切換
- ✓ 預覽表格更新
- ✓ 選擇功能正常
- ✓ 自動確認正常
- ✓ Console 無錯誤

**測試狀態**: [ ] 待測 [ ] 進行中 [ ] 通過 [ ] 失敗

---

## 測試執行結果

### 測試日期: ___________
### 測試人員: ___________

| 測試用例 | 狀態 | 備註 | 失敗原因 |
|---------|------|------|---------|
| TC-001 | [ ] | | |
| TC-002 | [ ] | | |
| TC-003 | [ ] | | |
| TC-004 | [ ] | | |
| TC-005 | [ ] | | |

---

## 失敗分析與矯正措施

### 失敗 #1: ___________
**現象**: 
**原因分析**: 
**矯正措施**: 
**驗證結果**: 

---

## 開發者工具檢查清單

- [ ] Console 無紅色錯誤
- [ ] Console 無黃色警告 (除外部庫)
- [ ] Network 標籤無 4xx/5xx 錯誤
- [ ] Performance 標籤無明顯卡頓
- [ ] 所有事件監聽器正確綁定
- [ ] DOM 元素正確更新
- [ ] 計時器正確清理

---

## 簽核

- [ ] 開發完成
- [ ] 測試通過
- [ ] 代碼審查通過
- [ ] 準備推送至 GitHub

