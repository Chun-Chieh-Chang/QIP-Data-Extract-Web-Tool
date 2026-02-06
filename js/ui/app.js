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
let groupSheetIndices = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 }; // 紀錄各穴組對應的工作表索引

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

    // 啟動滾動監聽 (Scroll Spy)
    if (typeof setupScrollSpy === 'function') {
        setupScrollSpy();
    }

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
    elements.exportConfigs = document.getElementById('export-configs');
    elements.importConfigsBtn = document.getElementById('import-configs-btn');
    elements.importConfigsFile = document.getElementById('import-configs-file');
    elements.clearAllConfigs = document.getElementById('clear-all-configs');
    elements.searchConfig = document.getElementById('search-config');

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

    elements.helpBtn = document.getElementById('help-btn');
    elements.helpDialog = document.getElementById('help-dialog');
    elements.closeHelpDialog = document.getElementById('close-help-dialog');

    elements.runtimeSpinner = document.getElementById('runtime-spinner');
    elements.globalResetBtn = document.getElementById('global-reset-btn');
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

    // 全域一鍵重置按鈕
    if (elements.globalResetBtn) {
        elements.globalResetBtn.addEventListener('click', performGlobalReset);
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
    if (elements.saveConfig) elements.saveConfig.addEventListener('click', saveConfiguration);
    if (elements.loadConfig) elements.loadConfig.addEventListener('click', () => showConfigDialog());
    if (elements.resetConfig) elements.resetConfig.addEventListener('click', resetConfiguration);

    if (elements.exportConfigs) {
        elements.exportConfigs.addEventListener('click', () => {
            console.log('點擊導出按鈕');
            exportConfigurations();
        });
    }

    if (elements.importConfigsBtn && elements.importConfigsFile) {
        elements.importConfigsBtn.addEventListener('click', () => {
            console.log('點擊導入按鈕，觸發檔案選擇');
            elements.importConfigsFile.click();
        });
        elements.importConfigsFile.addEventListener('change', (e) => {
            console.log('偵測到導入檔案變更');
            importConfigurations(e);
        });
    }

    if (elements.clearAllConfigs) elements.clearAllConfigs.addEventListener('click', clearAllConfigurations);
    if (elements.closeConfigDialog) elements.closeConfigDialog.addEventListener('click', () => elements.configDialog.close());

    // 使用說明
    if (elements.helpBtn) elements.helpBtn.addEventListener('click', () => elements.helpDialog.showModal());
    if (elements.closeHelpDialog) elements.closeHelpDialog.addEventListener('click', () => elements.helpDialog.close());

    // 點擊對話框外部關閉
    [elements.configDialog, elements.helpDialog].forEach(dialog => {
        if (dialog) {
            dialog.addEventListener('click', (e) => {
                if (e.target === dialog) dialog.close();
            });
        }
    });

    // 搜尋功能
    if (elements.searchConfig) {
        elements.searchConfig.addEventListener('input', (e) => {
            const term = e.target.value.toLowerCase().trim();
            renderConfigList(term);
        });
    }

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

            // 預設選取第一個工作表並預覽
            if (currentWorkbook.SheetNames.length > 0) {
                elements.worksheetSelect.selectedIndex = 1; // 索引 1 是第一個 Sheet (0 是佔位符)
                previewWorksheet();
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
            const sizeInMB = selectedFiles.reduce((acc, f) => acc + f.size, 0) / (1024 * 1024);
            elements.workbookInfo.innerHTML = `
                <div class="space-y-4 w-full">
                    <div class="flex flex-col gap-1 w-full max-w-full">
                        <span class="text-slate-400 flex-shrink-0 text-xs uppercase tracking-wider">範本來源:</span>
                        <span class="text-slate-700 dark:text-slate-200 font-black break-all text-sm leading-relaxed block w-full" title="${currentFileName}">${currentFileName}</span>
                    </div>
                    <div class="flex flex-col gap-1 w-full max-w-full">
                        <span class="text-slate-400 flex-shrink-0 text-xs uppercase tracking-wider">工作表數:</span>
                        <span class="text-slate-700 dark:text-slate-200 font-black text-sm">${sheetCount} 個分頁</span>
                    </div>
                    <div class="flex flex-col gap-1 w-full max-w-full">
                        <span class="text-slate-400 flex-shrink-0 text-xs uppercase tracking-wider">分頁清單:</span>
                        <span class="text-slate-500 dark:text-slate-400 break-all text-sm leading-relaxed block w-full">${currentWorkbook.SheetNames.slice(0, 8).join(', ')}${sheetCount > 8 ? '...' : ''}</span>
                    </div>
                    <div class="pt-4 flex flex-wrap items-center gap-3">
                        <span class="px-3 py-2 bg-primary/10 text-primary text-sm font-black uppercase rounded-xl">共 ${fileCount} 檔</span>
                        <span class="px-3 py-2 bg-slate-100 dark:bg-white/5 text-slate-500 dark:text-slate-400 text-sm font-black uppercase rounded-xl">${sizeInMB.toFixed(2)} MB</span>
                    </div>
                </div>
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

    // 自動捲動到指定位置 (Row 24 & Most Right)
    requestAnimationFrame(() => {
        const container = elements.previewTable.parentElement;
        if (container) {
            // 水平捲動到最右邊
            container.scrollLeft = container.scrollWidth;

            // 垂直捲動到第 24 行
            // 注意：tbody 的 index 是從 0 開始，所以第 24 行是 index 23
            const rows = elements.previewTable.querySelectorAll('tbody tr');
            if (rows.length > 23) {
                const targetRow = rows[23];
                // 使用 offsetTop 定位
                container.scrollTop = targetRow.offsetTop;
            }
        }
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

        // 自動確認選擇 (完成第2點後即確認)
        setTimeout(() => {
            confirmSelection();
        }, 300);
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

        // 紀錄該穴組目前選取時的工作表索引
        const groupId = parseInt(selectionTarget.split('-').pop());
        if (!isNaN(groupId)) {
            groupSheetIndices[groupId] = elements.worksheetSelect.selectedIndex - 1;
            console.log(`穴組 ${groupId} 已綁定至工作表索引: ${groupSheetIndices[groupId]}`);
        }

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

    // 自動隱藏預覽區 (完成選擇後自動收起)
    setTimeout(() => {
        elements.previewSection.style.display = 'none';
    }, 500);
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

    const baseIndex = groupSheetIndices[1] || 0;

    for (let i = 1; i <= 6; i++) {
        config.cavityGroups[i] = {
            cavityIdRange: document.getElementById(`cavity-id-${i}`)?.value || '',
            dataRange: document.getElementById(`data-range-${i}`)?.value || '',
            pageOffset: (groupSheetIndices[i] || 0) - baseIndex
        };
    }

    return config;
}

/**
 * 顯示配置對話框
 */
function showConfigDialog() {
    // 重置搜尋框
    if (elements.searchConfig) elements.searchConfig.value = '';

    renderConfigList();
    elements.configDialog.showModal();
}

/**
 * 渲染配置列表
 * @param {string} searchTerm 搜尋關鍵字
 */
function renderConfigList(searchTerm = '') {
    const configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');

    if (configs.length === 0) {
        elements.configList.innerHTML = `
            <div class="py-10 text-center text-slate-400">
                <span class="material-icons-round text-3xl mb-2 opacity-20">inventory_2</span>
                <p class="text-xs">尚無保存的配置</p>
            </div>
        `;
        return;
    }

    const filteredConfigs = searchTerm
        ? configs.filter(c => c.name.toLowerCase().includes(searchTerm))
        : configs;

    // 按時間降序排列 (最新的在前)
    filteredConfigs.sort((a, b) => new Date(b.savedAt) - new Date(a.savedAt));

    if (filteredConfigs.length === 0) {
        elements.configList.innerHTML = `
            <div class="py-10 text-center text-slate-400">
                <p class="text-xs">找不到符合 "${searchTerm}" 的配置</p>
            </div>
        `;
        return;
    }

    // 取得原始索引以確保操作正確
    let html = '';
    for (const c of filteredConfigs) {
        const originalIndex = configs.findIndex(orig => orig.name === c.name && orig.savedAt === c.savedAt);
        const savedAt = new Date(c.savedAt).toLocaleString('zh-TW', { hour12: false, month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });

        html += `
            <li class="flex items-center justify-between p-3 bg-slate-50 dark:bg-slate-800 rounded-xl border border-slate-100 dark:border-slate-700 hover:border-primary/40 transition-all group">
                <div class="flex flex-col min-w-0 flex-1 mr-3">
                    <strong class="text-sm text-slate-700 dark:text-slate-200 truncate" title="${c.name}">${c.name}</strong>
                    <div class="flex items-center gap-2 mt-0.5">
                        <span class="px-1.5 py-0.5 bg-slate-200 dark:bg-slate-700 text-[9px] font-bold text-slate-500 rounded lowercase">${c.cavityCount || 0} CAVITY</span>
                        <span class="text-[10px] text-slate-400 font-medium">${savedAt}</span>
                    </div>
                </div>
                <div class="flex gap-1.5 flex-shrink-0">
                    <button class="px-3 py-1.5 bg-primary/10 hover:bg-primary text-primary hover:text-white text-[11px] font-bold rounded-lg transition-colors" onclick="loadConfiguration(${originalIndex})">載入</button>
                    <button class="p-1.5 text-slate-400 hover:text-rose-500 transition-colors" onclick="deleteConfiguration(${originalIndex})">
                        <span class="material-icons-round text-base">delete</span>
                    </button>
                </div>
            </li>
        `;
    }

    // 如果數量很大，顯示計數
    if (configs.length > 5) {
        const infoHtml = `<p class="text-[10px] text-slate-400 mb-2 px-1">顯示 ${filteredConfigs.length} / ${configs.length} 筆配置</p>`;
        elements.configList.innerHTML = infoHtml + `<ul class="space-y-2">${html}</ul>`;
    } else {
        elements.configList.innerHTML = `<ul class="space-y-2">${html}</ul>`;
    }
}

/**
 * 導出所有配置
 */
function exportConfigurations() {
    const configs = localStorage.getItem('qip_configs');
    if (!configs || configs === '[]') {
        alert('目前沒有任何配置可以導出');
        return;
    }

    try {
        const blob = new Blob([configs], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');

        a.href = url;
        a.download = `QIP_Configurations_Backup_${date}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        console.log('配置已導出至本地硬碟');
    } catch (error) {
        console.error('導出失敗:', error);
        alert('導出失敗: ' + error.message);
    }
}

/**
 * 導入配置
 */
function importConfigurations(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (event) {
        try {
            const importedConfigs = JSON.parse(event.target.result);
            if (!Array.isArray(importedConfigs)) {
                throw new Error('不正確的配置格式');
            }

            if (confirm(`確定要導入 ${importedConfigs.length} 筆配置嗎？這將會與現有配置合併。`)) {
                const currentConfigs = JSON.parse(localStorage.getItem('qip_configs') || '[]');

                // 合併並去重 (以名稱和時間戳記為準)
                const merged = [...currentConfigs];
                let addedCount = 0;

                for (const imp of importedConfigs) {
                    const exists = merged.find(c => c.name === imp.name && c.savedAt === imp.savedAt);
                    if (!exists) {
                        merged.push(imp);
                        addedCount++;
                    }
                }

                localStorage.setItem('qip_configs', JSON.stringify(merged));
                alert(`導入成功！新增了 ${addedCount} 筆配置，總計 ${merged.length} 筆。`);

                // 更新列表（如果對話框開著）
                if (elements.configDialog.open) {
                    renderConfigList();
                }
            }
        } catch (error) {
            console.error('導入失敗:', error);
            alert('導入失敗: ' + error.message);
        } finally {
            elements.importConfigsFile.value = ''; // 重置 input
        }
    };
    reader.readAsText(file);
}

/**
 * 清空所有配置 (重置時清除檔案)
 */
function clearAllConfigurations() {
    const count = JSON.parse(localStorage.getItem('qip_configs') || '[]').length;
    if (count === 0) {
        alert('目前沒有任何配置。');
        return;
    }

    if (confirm(`危險：確定要刪除所有 ${count} 筆已保存的配置嗎？此操作無法恢復！\n\n建議先執行「導出」備份。`)) {
        localStorage.removeItem('qip_configs');
        alert('所有配置已清空。');

        if (elements.configDialog.open) {
            renderConfigList();
        }
    }
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
            // 恢復索引記憶 (根據 pageOffset 恢復，假設基底是索引 0)
            groupSheetIndices[i] = group.pageOffset || 0;
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
    groupSheetIndices = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };

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
 * 執行全域一鍵重置
 * 清空所有欄位、釋放記憶體、重置應用狀態
 */
function performGlobalReset() {
    // 確認對話框
    if (!confirm('確定要執行一鍵重置嗎？\n\n這將會：\n✓ 清空所有輸入欄位\n✓ 移除已上傳的檔案\n✓ 釋放記憶體空間\n✓ 重置所有設定\n\n此操作無法復原。')) {
        return;
    }

    console.log('開始執行全域重置...');

    // 顯示狀態更新
    if (typeof updateStatus === 'function') {
        updateStatus('processing', '正在重置系統...');
    }

    try {
        // 1. 清空全域變數，釋放記憶體
        currentWorkbook = null;
        currentFileName = '';
        selectedFiles = [];
        processingResults = null;
        selectionMode = null;
        selectionTarget = null;
        selectionStart = null;
        selectionEnd = null;
        groupSheetIndices = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };

        // 2. 清空檔案輸入
        if (elements.fileInput) {
            elements.fileInput.value = '';
        }

        // 3. 重置所有配置輸入欄位
        if (elements.productCode) elements.productCode.value = '';
        if (elements.cavityCount) elements.cavityCount.value = '';
        if (elements.configName) elements.configName.value = '';

        // 4. 清空所有穴組範圍輸入
        for (let i = 1; i <= 6; i++) {
            const cavityIdInput = document.getElementById(`cavity-id-${i}`);
            const dataRangeInput = document.getElementById(`data-range-${i}`);
            const offsetInput = document.getElementById(`offset-${i}`);

            if (cavityIdInput) cavityIdInput.value = '';
            if (dataRangeInput) dataRangeInput.value = '';
            if (offsetInput) offsetInput.value = '1';
        }

        // 5. 移除所有 has-value 類別
        document.querySelectorAll('.config-input').forEach(input => {
            input.classList.remove('has-value');
        });

        // 6. 隱藏所有穴組（除了第一組）
        for (let i = 2; i <= 6; i++) {
            const group = document.getElementById(`cavity-group-${i}`);
            if (group) {
                group.classList.add('hidden');
            }
        }

        // 7. 重置 UI 顯示狀態
        if (elements.fileInfo) elements.fileInfo.style.display = 'none';
        if (elements.uploadArea) elements.uploadArea.style.display = 'block';
        if (elements.worksheetGroup) elements.worksheetGroup.style.display = 'none';
        if (elements.rangeGroup) elements.rangeGroup.style.display = 'none';
        if (elements.previewSection) elements.previewSection.style.display = 'none';
        if (elements.resultSection) elements.resultSection.style.display = 'none';
        if (elements.progressContainer) elements.progressContainer.style.display = 'none';

        // 8. 清空預覽表格
        if (elements.previewTable) {
            elements.previewTable.innerHTML = '';
        }

        // 9. 清空結果區域
        if (elements.resultSummary) {
            elements.resultSummary.innerHTML = '';
        }
        if (elements.errorList) {
            elements.errorList.innerHTML = '';
        }
        if (elements.errorLog) {
            elements.errorLog.style.display = 'none';
        }

        // 10. 重置工作表選擇器
        if (elements.worksheetSelect) {
            elements.worksheetSelect.innerHTML = '<option value="">-- 請選擇工作表 --</option>';
        }

        // 11. 重置進度條
        if (elements.progressFill) {
            elements.progressFill.style.width = '0%';
        }
        if (elements.progressText) {
            elements.progressText.textContent = '等待指令中...';
        }

        // 12. 更新按鈕狀態
        updateStartButton();

        // 13. 強制垃圾回收提示（瀏覽器會自動處理）
        console.log('記憶體釋放完成');

        // 14. 顯示成功訊息
        if (typeof updateStatus === 'function') {
            updateStatus('success', '系統已完全重置');
            setTimeout(() => {
                updateStatus('ready', '系統就緒');
            }, 2000);
        }

        console.log('全域重置完成');

        // 15. 可選：滾動到頁面頂部
        window.scrollTo({ top: 0, behavior: 'smooth' });

    } catch (error) {
        console.error('重置過程發生錯誤:', error);
        if (typeof updateStatus === 'function') {
            updateStatus('error', '重置失敗');
        }
        alert('重置過程發生錯誤: ' + error.message);
    }
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
    // 顯示進度
    elements.progressContainer.style.display = 'block';
    elements.startProcess.disabled = true;
    elements.resultSection.style.display = 'none';

    // 更新狀態為處理中
    if (typeof statusManager !== 'undefined') {
        statusManager.setStatus('processing');
    }

    // 啟動動畫
    if (elements.runtimeSpinner) {
        elements.runtimeSpinner.classList.add('animate-spin');
    }

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

        // 更新狀態為錯誤
        if (typeof statusManager !== 'undefined') {
            statusManager.setStatus('error');
        }
    } finally {
        elements.startProcess.disabled = false;
        // 停止動畫
        if (elements.runtimeSpinner) {
            elements.runtimeSpinner.classList.remove('animate-spin');
        }

        // 狀態更新將在 showResults 中處理成功，或者在這裡處理完成態
    }
}

/**
 * 顯示處理結果
 */
function showResults(results) {
    elements.progressFill.style.width = '100%';
    elements.progressText.textContent = '數據提取執行完成！';

    // 更新狀態為成功
    if (typeof statusManager !== 'undefined') {
        statusManager.setStatus('success');
    }

    elements.resultSection.style.display = 'block';

    const itemCount = Object.keys(results.inspectionItems).length;

    elements.resultSummary.innerHTML = `
        <h3 class="text-lg font-black flex items-center gap-3 mb-6 text-emerald-600 dark:text-emerald-400">
             <span class="material-icons-round text-2xl">check_circle</span> 數據提取成功
        </h3>
        <div class="grid grid-cols-3 gap-6">
            <div class="bg-white/50 dark:bg-slate-800/50 p-5 rounded-2xl text-center border border-slate-200 dark:border-white/5">
                <p class="text-xs uppercase font-black text-slate-500 dark:text-slate-400 tracking-widest mb-2">檢驗項目數</p>
                <p class="text-3xl font-black text-slate-800 dark:text-white">${itemCount}</p>
            </div>
            <div class="bg-white/50 dark:bg-slate-800/50 p-5 rounded-2xl text-center border border-slate-200 dark:border-white/5">
                <p class="text-xs uppercase font-black text-slate-500 dark:text-slate-400 tracking-widest mb-2">處理工作表</p>
                <p class="text-3xl font-black text-slate-800 dark:text-white">${results.processedSheets}</p>
            </div>
            <div class="bg-white/50 dark:bg-slate-800/50 p-5 rounded-2xl text-center border border-slate-200 dark:border-white/5">
                <p class="text-xs uppercase font-black text-slate-500 dark:text-slate-400 tracking-widest mb-2">總計穴數</p>
                <p class="text-3xl font-black text-slate-800 dark:text-white">${results.totalCavities}</p>
            </div>
        </div>
        <p class="mt-8 text-sm font-bold text-slate-600 dark:text-slate-400 leading-loose">
            所有數據提取任務已順利完成。您可以下載生成的 Excel 報表以進行進一步的分析與存檔。
        </p>
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
