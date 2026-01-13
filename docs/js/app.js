/**
 * QIP Data Extract Tool - 主應用程式
 * 處理 UI 交互、檔案上傳、配置管理
 */

// 全域變數
let currentWorkbook = null;
let currentFileName = '';
let selectedFiles = []; // 儲存所有選取的檔案
let processingResults = null;
let selectionMode = null;
let selectionTarget = null;
let selectionStart = null;
let selectionEnd = null;

// DOM 元素緩存
const elements = {};

/**
 * 初始化應用程式
 */
document.addEventListener('DOMContentLoaded', () => {
    console.log('QIP Data Extract Tool 初始化...');

    // 緩存 DOM 元素
    cacheElements();

    // 綁定事件
    bindEvents();

    // 載入已保存的配置
    loadSavedConfigs();

    console.log('初始化完成');
});

/**
 * 緩存 DOM 元素
 */
function cacheElements() {
    elements.fileInput = document.getElementById('file-input');
    elements.browseBtn = document.getElementById('browse-btn');
    elements.uploadArea = document.getElementById('upload-area');
    elements.fileInfo = document.getElementById('file-info');
    elements.selectedFileName = document.getElementById('selected-file-name');
    elements.removeFile = document.getElementById('remove-file');
    elements.fullReset = document.getElementById('full-reset');
    elements.workbookInfo = document.getElementById('workbook-info');

    elements.productCode = document.getElementById('product-code');
    elements.cavityCount = document.getElementById('cavity-count');

    elements.worksheetGroup = document.getElementById('worksheet-group');
    elements.worksheetSelect = document.getElementById('worksheet-select');
    elements.previewWorksheet = document.getElementById('preview-worksheet');

    elements.rangeGroup = document.getElementById('range-group');
    elements.previewSection = document.getElementById('preview-section');
    elements.previewTable = document.getElementById('preview-table');
    elements.prevSheetBtn = document.getElementById('prev-sheet');
    elements.nextSheetBtn = document.getElementById('next-sheet');
    elements.currentSheetLabel = document.getElementById('current-sheet-name');
    elements.selectionModeText = document.getElementById('selection-mode-text');
    elements.confirmSelection = document.getElementById('confirm-selection');
    elements.cancelSelection = document.getElementById('cancel-selection');

    elements.configName = document.getElementById('config-name');
    elements.saveConfig = document.getElementById('save-config');
    elements.loadConfig = document.getElementById('load-config');
    elements.resetConfig = document.getElementById('reset-config');

    elements.startProcess = document.getElementById('start-process');
    elements.progressContainer = document.getElementById('progress-container');
    elements.progressFill = document.getElementById('progress-fill');
    elements.progressText = document.getElementById('progress-text');

    elements.resultSection = document.getElementById('result-section');
    elements.resultSummary = document.getElementById('result-summary');
    elements.downloadExcel = document.getElementById('download-excel');
    elements.errorLog = document.getElementById('error-log');
    elements.errorList = document.getElementById('error-list');

    elements.configDialog = document.getElementById('config-dialog');
    elements.configList = document.getElementById('config-list');
    elements.closeConfigDialog = document.getElementById('close-config-dialog');
}

/**
 * 綁定事件處理器
 */
function bindEvents() {
    // 檔案上傳
    elements.browseBtn.addEventListener('click', () => elements.fileInput.click());
    elements.fileInput.addEventListener('change', handleFileSelect);
    elements.removeFile.addEventListener('click', removeFile);

    // 完全重置
    if (elements.fullReset) {
        elements.fullReset.addEventListener('click', () => {
            if (confirm('確定要完全重置嗎？這將清空所有資料並重新載入頁面。')) {
                location.reload();
            }
        });
    }

    // 拖放上傳
    elements.uploadArea.addEventListener('dragover', handleDragOver);
    elements.uploadArea.addEventListener('dragleave', handleDragLeave);
    elements.uploadArea.addEventListener('drop', handleDrop);
    elements.uploadArea.addEventListener('click', (e) => {
        if (e.target === elements.uploadArea || e.target.classList.contains('upload-icon')) {
            elements.fileInput.click();
        }
    });

    // 模穴數變更
    elements.cavityCount.addEventListener('change', handleCavityCountChange);

    // 工作表選擇變更時自動預覽
    elements.worksheetSelect.addEventListener('change', () => {
        if (elements.worksheetSelect.value) {
            previewWorksheet();
        }
    });

    // 預覽按鈕 (保留作為手動刷新)
    elements.previewWorksheet.addEventListener('click', previewWorksheet);

    // 範圍選擇按鈕
    document.querySelectorAll('.select-range-btn').forEach(btn => {
        btn.addEventListener('click', () => startRangeSelection(btn));
    });

    // 輸入框聚焦時自動啟動選擇模式
    document.querySelectorAll('.range-input').forEach(input => {
        input.addEventListener('focus', (e) => {
            // 確保預覽已開啟
            if (elements.previewSection.style.display === 'none') {
                previewWorksheet();
            }

            // 設定選擇模式
            const btn = e.target.nextElementSibling; // 假設按鈕在輸入框後面
            if (btn && btn.classList.contains('select-range-btn')) {
                startRangeSelection(btn);
            }
        });
    });

    // 預覽表格選擇
    elements.confirmSelection.addEventListener('click', confirmSelection);
    elements.cancelSelection.addEventListener('click', cancelSelection);

    // 工作表切換按鈕
    elements.prevSheetBtn.addEventListener('click', () => switchSheet(-1));
    elements.nextSheetBtn.addEventListener('click', () => switchSheet(1));

    // 配置管理
    elements.saveConfig.addEventListener('click', saveConfiguration);
    elements.loadConfig.addEventListener('click', showConfigDialog);
    elements.resetConfig.addEventListener('click', resetConfiguration);
    elements.closeConfigDialog.addEventListener('click', () => elements.configDialog.close());

    // 處理和下載
    elements.startProcess.addEventListener('click', startProcessing);
    elements.downloadExcel.addEventListener('click', downloadResults);

    // 監聽所有配置輸入框的變更 (支援手動輸入)
    document.querySelectorAll('.config-input').forEach(input => {
        input.addEventListener('input', (e) => {
            const hasVal = !!e.target.value.trim();
            e.target.classList.toggle('has-value', hasVal);
            updateStartButton();
        });
    });
}

/**
 * 處理檔案選擇
 */
async function handleFileSelect(e) {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
        await loadFiles(files);
    }
}

/**
 * 處理拖放
 */
function handleDragOver(e) {
    e.preventDefault();
    elements.uploadArea.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    elements.uploadArea.classList.remove('dragover');
}

async function handleDrop(e) {
    e.preventDefault();
    elements.uploadArea.classList.remove('dragover');

    const files = Array.from(e.dataTransfer.files).filter(file => isExcelFile(file));
    if (files.length > 0) {
        await loadFiles(files);
    } else {
        alert('請上傳 Excel 檔案 (.xls, .xlsx, .xlsm)');
    }
}

/**
 * 檢查是否為 Excel 檔案
 */
function isExcelFile(file) {
    const validTypes = [
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel.sheet.macroEnabled.12'
    ];
    const validExtensions = ['.xls', '.xlsx', '.xlsm'];
    const extension = '.' + file.name.split('.').pop().toLowerCase();

    return validTypes.includes(file.type) || validExtensions.includes(extension);
}

/**
 * 載入多個檔案
 */
async function loadFiles(files) {
    try {
        console.log(`載入 ${files.length} 個檔案`);

        // 如果是第一次載入，使用第一個檔案作為預覽範本
        const isFirstLoad = selectedFiles.length === 0;

        // 加入新的檔案至列表 (避免重複)
        for (const file of files) {
            if (!selectedFiles.find(f => f.name === file.name && f.size === file.size)) {
                selectedFiles.push(file);
            }
        }

        if (isFirstLoad && selectedFiles.length > 0) {
            const firstFile = selectedFiles[0];
            const data = await firstFile.arrayBuffer();
            currentWorkbook = XLSX.read(data, { type: 'array' });
            currentFileName = firstFile.name;

            // 更新工作表選擇器
            updateWorksheetSelector();

            // 自動填充產品品號
            if (!elements.productCode.value) {
                const baseName = firstFile.name.replace(/\.[^/.]+$/, '');
                elements.productCode.value = baseName;
                elements.productCode.classList.add('has-value');
            }
        }

        // 更新 UI
        elements.fileInfo.style.display = 'block';
        elements.uploadArea.style.display = 'none';

        const fileCount = selectedFiles.length;
        if (fileCount === 1) {
            elements.selectedFileName.textContent = selectedFiles[0].name;
        } else {
            elements.selectedFileName.textContent = `已選取 ${fileCount} 個檔案`;
        }

        // 顯示工作簿信息 (顯示目前作為範本的檔案)
        if (currentWorkbook) {
            const sheetCount = currentWorkbook.SheetNames.length;
            elements.workbookInfo.innerHTML = `
                <p>📄 範本檔案: <strong>${currentFileName}</strong></p>
                <p>📊 工作表數量: <strong>${sheetCount}</strong></p>
                <p>工作表: ${currentWorkbook.SheetNames.slice(0, 5).join(', ')}${sheetCount > 5 ? '...' : ''}</p>
                <p class="mt-2 text-primary font-bold">已就緒，將提取共 ${fileCount} 個檔案的數據</p>
            `;
        }

        // 顯示相關區段
        elements.worksheetGroup.style.display = 'block';

        // 檢查是否可以開始處理
        updateStartButton();

        console.log('檔案載入完成');
    } catch (error) {
        console.error('載入檔案失敗:', error);
        alert('載入檔案失敗: ' + error.message);
    }
}

/**
 * 移除檔案
 */
function removeFile() {
    currentWorkbook = null;
    currentFileName = '';
    selectedFiles = [];
    elements.fileInput.value = '';
    elements.fileInfo.style.display = 'none';
    elements.uploadArea.style.display = 'block';
    elements.worksheetGroup.style.display = 'none';
    elements.rangeGroup.style.display = 'none';
    elements.previewSection.style.display = 'none';
    elements.resultSection.style.display = 'none';
    updateStartButton();
}

/**
 * 更新工作表選擇器
 */
function updateWorksheetSelector() {
    elements.worksheetSelect.innerHTML = '<option value="">-- 請選擇工作表 --</option>';

    if (currentWorkbook) {
        for (const name of currentWorkbook.SheetNames) {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            elements.worksheetSelect.appendChild(option);
        }
    }
}

/**
 * 處理模穴數變更
 */
function handleCavityCountChange() {
    const count = parseInt(elements.cavityCount.value) || 0;

    // 顯示/隱藏對應的穴組
    for (let i = 2; i <= 6; i++) {
        const group = document.getElementById(`cavity-group-${i}`);
        if (group) {
            const minCavities = i * 8;
            group.classList.toggle('hidden', count < minCavities);
        }
    }

    // 顯示範圍設定區
    if (count > 0) {
        elements.rangeGroup.style.display = 'block';
    }

    elements.cavityCount.classList.toggle('has-value', count > 0);
    updateStartButton();
}

/**
 * 預覽工作表
 */
function previewWorksheet() {
    const sheetName = elements.worksheetSelect.value;
    if (!sheetName || !currentWorkbook) {
        alert('請先選擇工作表');
        return;
    }

    const worksheet = currentWorkbook.Sheets[sheetName];
    renderPreviewTable(worksheet);

    // 更新顯示名稱
    if (elements.currentSheetLabel) {
        elements.currentSheetLabel.textContent = sheetName;
    }

    elements.previewSection.style.display = 'block';
}

/**
 * 切換工作表
 */
function switchSheet(offset) {
    const select = elements.worksheetSelect;
    if (!select || select.options.length === 0) return;

    const newIndex = select.selectedIndex + offset;
    if (newIndex >= 1 && newIndex < select.options.length) { // Skip index 0 (placeholder)
        select.selectedIndex = newIndex;
        select.dispatchEvent(new Event('change'));
    }
}

/**
 * 渲染預覽表格
 */
function renderPreviewTable(worksheet) {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    const maxRows = Math.min(range.e.r + 1, 400); // Increase to 400 rows
    const maxCols = Math.min(range.e.c + 1, 150); // Increase to 150 cols (approx ET)

    // Prepare merge map
    const mergeMap = {};
    if (worksheet['!merges']) {
        worksheet['!merges'].forEach(merge => {
            if (merge.s.r < maxRows && merge.s.c < maxCols) {
                const key = `${merge.s.r},${merge.s.c}`;
                mergeMap[key] = {
                    rowspan: Math.min(merge.e.r, maxRows - 1) - merge.s.r + 1,
                    colspan: Math.min(merge.e.c, maxCols - 1) - merge.s.c + 1
                };

                // Mark covered cells
                for (let r = merge.s.r; r <= Math.min(merge.e.r, maxRows - 1); r++) {
                    for (let c = merge.s.c; c <= Math.min(merge.e.c, maxCols - 1); c++) {
                        if (r === merge.s.r && c === merge.s.c) continue;
                        mergeMap[`${r},${c}`] = { hidden: true };
                    }
                }
            }
        });
    }

    let html = '<thead><tr><th></th>';

    // Column Headers
    for (let c = 0; c < maxCols; c++) {
        html += `<th>${XLSX.utils.encode_col(c)}</th>`;
    }
    html += '</tr></thead><tbody>';

    // Rows
    for (let r = 0; r < maxRows; r++) {
        html += `<tr><th>${r + 1}</th>`;
        for (let c = 0; c < maxCols; c++) {
            const key = `${r},${c}`;
            if (mergeMap[key]?.hidden) continue;

            const cellAddr = XLSX.utils.encode_cell({ r, c });
            const cell = worksheet[cellAddr];
            const value = cell ? (cell.w || cell.v || '') : '';

            let attrs = `data-row="${r}" data-col="${c}" title="${value}"`;
            let classes = [];

            // Add classes for styling if needed
            if (cell?.t === 'n') classes.push('numeric');

            if (mergeMap[key]) {
                if (mergeMap[key].rowspan > 1) attrs += ` rowspan="${mergeMap[key].rowspan}"`;
                if (mergeMap[key].colspan > 1) attrs += ` colspan="${mergeMap[key].colspan}"`;
            }

            if (classes.length) attrs += ` class="${classes.join(' ')}"`;

            let displayValue = String(value);
            if (displayValue.length > 20) displayValue = displayValue.substring(0, 20) + '...';

            html += `<td ${attrs}>${displayValue}</td>`;
        }
        html += '</tr>';
    }
    html += '</tbody>';

    elements.previewTable.innerHTML = html;

    // 綁定儲存格選擇事件
    elements.previewTable.querySelectorAll('td').forEach(td => {
        td.addEventListener('click', handleCellClick);
    });
}

/**
 * 開始範圍選擇
 */
function startRangeSelection(btn) {
    if (!currentWorkbook || !elements.worksheetSelect.value) {
        alert('請先上傳檔案並選擇工作表');
        return;
    }

    selectionMode = btn.dataset.type;
    selectionTarget = btn.dataset.target;
    selectionStart = null;
    selectionEnd = null;

    elements.selectionModeText.textContent = selectionMode === 'cavity' ? '選擇穴號範圍' : '選擇數據範圍';
    elements.confirmSelection.disabled = true;

    // 清除之前的選擇
    elements.previewTable.querySelectorAll('.selected, .selection-start').forEach(el => {
        el.classList.remove('selected', 'selection-start');
    });

    // 確保預覽表格可見
    if (elements.previewSection.style.display === 'none') {
        previewWorksheet();
    }

    elements.previewSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * 處理儲存格點擊
 */
function handleCellClick(e) {
    if (!selectionMode) return;

    const td = e.target;
    const row = parseInt(td.dataset.row);
    const col = parseInt(td.dataset.col);

    if (!selectionStart) {
        // 第一次點擊 - 設置起點
        selectionStart = { row, col };
        td.classList.add('selection-start');
        elements.confirmSelection.disabled = true;
    } else {
        // 第二次點擊 - 設置終點
        selectionEnd = { row, col };

        // 高亮選擇範圍
        highlightSelection(selectionStart, selectionEnd);
        elements.confirmSelection.disabled = false;
    }
}

/**
 * 高亮選擇範圍
 */
function highlightSelection(start, end) {
    // 清除之前的選擇
    elements.previewTable.querySelectorAll('.selected').forEach(el => {
        el.classList.remove('selected');
    });

    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);

    elements.previewTable.querySelectorAll('td').forEach(td => {
        const row = parseInt(td.dataset.row);
        const col = parseInt(td.dataset.col);

        if (row >= minRow && row <= maxRow && col >= minCol && col <= maxCol) {
            td.classList.add('selected');
        }
    });
}

/**
 * 確認選擇
 */
function confirmSelection() {
    if (!selectionStart || !selectionEnd || !selectionTarget) return;

    const minRow = Math.min(selectionStart.row, selectionEnd.row);
    const maxRow = Math.max(selectionStart.row, selectionEnd.row);
    const minCol = Math.min(selectionStart.col, selectionEnd.col);
    const maxCol = Math.max(selectionStart.col, selectionEnd.col);

    const startCell = XLSX.utils.encode_cell({ r: minRow, c: minCol });
    const endCell = XLSX.utils.encode_cell({ r: maxRow, c: maxCol });
    const rangeStr = startCell === endCell ? startCell : `${startCell}:${endCell}`;

    const input = document.getElementById(selectionTarget);
    if (input) {
        input.value = rangeStr;
        input.classList.add('has-value');

        // 觸發 input 事件以更新狀態 (如按鈕啟用)
        const event = new Event('input', { bubbles: true });
        input.dispatchEvent(event);
    }

    cancelSelection();
    updateStartButton();
}

/**
 * 取消選擇
 */
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

/**
 * 保存配置
 */
function saveConfiguration() {
    const name = elements.configName.value.trim();
    if (!name) {
        alert('請輸入配置名稱');
        return;
    }

    const config = gatherConfiguration();
    config.name = name;
    config.savedAt = new Date().toISOString();

    // 從 localStorage 載入現有配置
    const configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');

    // 檢查是否已存在同名配置
    const existingIndex = configs.findIndex(c => c.name === name);
    if (existingIndex >= 0) {
        configs[existingIndex] = config;
    } else {
        configs.push(config);
    }

    localStorage.setItem('qip_configs', JSON.stringify(configs));
    alert('配置已保存');
}

/**
 * 收集當前配置
 */
function gatherConfiguration() {
    const config = {
        productCode: elements.productCode.value,
        cavityCount: elements.cavityCount.value,
        cavityGroups: {}
    };

    for (let i = 1; i <= 6; i++) {
        config.cavityGroups[i] = {
            cavityIdRange: document.getElementById(`cavity-id-${i}`)?.value || '',
            dataRange: document.getElementById(`data-range-${i}`)?.value || '',
            pageOffset: i === 1 ? 0 : parseInt(document.getElementById(`offset-${i}`)?.value || '1') - 1
        };
    }

    return config;
}

/**
 * 顯示配置對話框
 */
function showConfigDialog() {
    const configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');

    if (configs.length === 0) {
        alert('尚未保存任何配置');
        return;
    }

    let html = '<ul class="space-y-3">';
    for (let i = 0; i < configs.length; i++) {
        const c = configs[i];
        const savedAt = new Date(c.savedAt).toLocaleString();
        html += `
            <li class="flex items-center justify-between p-3 bg-slate-50 dark:bg-slate-800 rounded-xl border border-slate-100 dark:border-slate-700 hover:border-primary/50 transition-colors">
                <div class="flex flex-col">
                    <strong class="text-sm text-slate-700 dark:text-slate-200">${c.name}</strong>
                    <span class="text-[10px] text-slate-400 font-medium">${c.cavityCount || 0} 穴 | ${savedAt}</span>
                </div>
                <div class="flex gap-2">
                    <button class="px-3 py-1 bg-primary text-white text-[11px] font-bold rounded-lg" onclick="loadConfiguration(${i})">載入</button>
                    <button class="px-3 py-1 bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-300 text-[11px] font-bold rounded-lg" onclick="deleteConfiguration(${i})">刪除</button>
                </div>
            </li>
        `;
    }
    html += '</ul>';

    elements.configList.innerHTML = html;
    elements.configDialog.showModal();
}

/**
 * 載入配置
 */
function loadConfiguration(index) {
    const configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
    const config = configs[index];

    if (!config) return;

    elements.productCode.value = config.productCode || '';
    elements.cavityCount.value = config.cavityCount || '';
    elements.configName.value = config.name || '';

    // 觸發模穴數變更
    handleCavityCountChange();

    // 填充穴組配置
    for (let i = 1; i <= 6; i++) {
        const group = config.cavityGroups[i];
        if (group) {
            const cavityIdInput = document.getElementById(`cavity-id-${i}`);
            const dataRangeInput = document.getElementById(`data-range-${i}`);
            const offsetInput = document.getElementById(`offset-${i}`);

            if (cavityIdInput) {
                cavityIdInput.value = group.cavityIdRange || '';
                cavityIdInput.classList.toggle('has-value', !!group.cavityIdRange);
            }
            if (dataRangeInput) {
                dataRangeInput.value = group.dataRange || '';
                dataRangeInput.classList.toggle('has-value', !!group.dataRange);
            }
            if (offsetInput && i > 1) {
                offsetInput.value = (group.pageOffset || 0) + 1;
            }
        }
    }

    elements.configDialog.close();
    updateStartButton();
    alert('配置已載入');
}

/**
 * 刪除配置
 */
function deleteConfiguration(index) {
    if (!confirm('確定要刪除這個配置嗎？')) return;

    const configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
    configs.splice(index, 1);
    localStorage.setItem('qip_configs', JSON.stringify(configs));

    showConfigDialog(); // 刷新列表
}

/**
 * 重置配置
 */
function resetConfiguration() {
    if (!confirm('確定要重置所有設定嗎？')) return;

    elements.productCode.value = '';
    elements.cavityCount.value = '';
    elements.configName.value = '';

    for (let i = 1; i <= 6; i++) {
        const cavityIdInput = document.getElementById(`cavity-id-${i}`);
        const dataRangeInput = document.getElementById(`data-range-${i}`);
        const offsetInput = document.getElementById(`offset-${i}`);

        if (cavityIdInput) cavityIdInput.value = '';
        if (dataRangeInput) dataRangeInput.value = '';
        if (offsetInput) offsetInput.value = '1';
    }

    document.querySelectorAll('.config-input').forEach(input => {
        input.classList.remove('has-value');
    });

    handleCavityCountChange();
    updateStartButton();
}

/**
 * 載入已保存的配置列表
 */
function loadSavedConfigs() {
    // 檢查是否有配置
    const configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
    console.log(`發現 ${configs.length} 個已保存的配置`);
}

/**
 * 更新開始處理按鈕狀態
 */
function updateStartButton() {
    const hasFile = selectedFiles.length > 0;
    const hasCavityCount = parseInt(elements.cavityCount.value) > 0;
    const hasCavityId = document.getElementById('cavity-id-1')?.value?.trim() !== '';
    const hasDataRange = document.getElementById('data-range-1')?.value?.trim() !== '';

    const canStart = hasFile && hasCavityCount && hasCavityId && hasDataRange;
    elements.startProcess.disabled = !canStart;
}

/**
 * 開始處理
 */
async function startProcessing() {
    if (selectedFiles.length === 0) {
        alert('請先上傳檔案');
        return;
    }

    const config = gatherConfiguration();
    console.log('開始處理，配置:', config);

    // 顯示進度
    elements.progressContainer.style.display = 'block';
    elements.startProcess.disabled = true;
    elements.resultSection.style.display = 'none';

    try {
        const processor = new QIPProcessor(config);
        const workbooks = [];

        // 逐一讀取檔案 (避免一次讀取太多檔案造成記憶體壓力)
        for (let i = 0; i < selectedFiles.length; i++) {
            const file = selectedFiles[i];
            elements.progressText.textContent = `正在讀取檔案 (${i + 1}/${selectedFiles.length}): ${file.name}`;

            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            workbook.fileName = file.name; // 用於錯誤記錄
            workbooks.push(workbook);

            // 讓 UI 有機會更新
            await new Promise(resolve => setTimeout(resolve, 0));
        }

        processingResults = await processor.processMultipleWorkbooks(workbooks, (progress) => {
            elements.progressFill.style.width = `${progress.percent}%`;
            elements.progressText.textContent = progress.message;
        });

        // 顯示結果
        showResults(processingResults);

    } catch (error) {
        console.error('處理失敗:', error);
        alert('處理失敗: ' + error.message);
    } finally {
        elements.startProcess.disabled = false;
    }
}

/**
 * 顯示處理結果
 */
function showResults(results) {
    elements.progressFill.style.width = '100%';
    elements.progressText.textContent = '處理完成！';

    elements.resultSection.style.display = 'block';

    const itemCount = Object.keys(results.inspectionItems).length;

    elements.resultSummary.innerHTML = `
        <h3 class="font-bold flex items-center gap-2 mb-3">
             <span class="material-icons-round">check_circle</span> 數據提取成功
        </h3>
        <div class="grid grid-cols-3 gap-4">
            <div class="bg-white/50 dark:bg-black/10 p-3 rounded-lg text-center">
                <p class="text-[10px] uppercase font-bold opacity-60">檢驗項目</p>
                <p class="text-xl font-bold">${itemCount}</p>
            </div>
            <div class="bg-white/50 dark:bg-black/10 p-3 rounded-lg text-center">
                <p class="text-[10px] uppercase font-bold opacity-60">處理工作表</p>
                <p class="text-xl font-bold">${results.processedSheets}</p>
            </div>
            <div class="bg-white/50 dark:bg-black/10 p-3 rounded-lg text-center">
                <p class="text-[10px] uppercase font-bold opacity-60">總穴數</p>
                <p class="text-xl font-bold">${results.totalCavities}</p>
            </div>
        </div>
        <p class="mt-4 text-xs opacity-80">資料處理已完成，您可以點擊下方按鈕下載 Excel 結果檔案。</p>
    `;

    // 顯示錯誤日誌（如果有）
    if (results.hasErrors) {
        elements.errorLog.style.display = 'block';
        elements.errorList.innerHTML = results.errors.map(e =>
            `<div class="error-item">
                <strong>${e.sheetName}</strong>: ${e.errorMessage}
            </div>`
        ).join('');
    } else {
        elements.errorLog.style.display = 'none';
    }

    elements.resultSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * 下載結果
 */
function downloadResults() {
    if (!processingResults) {
        alert('沒有處理結果可供下載');
        return;
    }

    try {
        const exporter = new ExcelExporter();
        const productCode = elements.productCode.value || 'QIP';

        exporter.createFromResults(processingResults, productCode);

        const filename = `${productCode}_數據提取結果_${formatDate(new Date())}`;
        exporter.download(filename);

    } catch (error) {
        console.error('導出失敗:', error);
        alert('導出失敗: ' + error.message);
    }
}

/**
 * 格式化日期
 */
function formatDate(date) {
    return date.toISOString().slice(0, 10).replace(/-/g, '');
}
