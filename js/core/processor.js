/**
 * QIPProcessor - 核心處理邏輯
 * 對應 VBA theCode.bas 中的處理核心
 */
class QIPProcessor {
    constructor(config) {
        this.config = config;
        this.errorLogger = new ErrorLogger();
        this.results = {
            inspectionItems: {},
            totalBatches: 0,
            totalCavities: 0,
            processedSheets: 0,
            productInfo: { productName: '', measurementUnit: '' }
        };
    }

    /**
     * 處理工作簿
     * @param {Object} workbook - SheetJS workbook
     * @param {Function} progressCallback - 進度回調函數
     * @returns {Object} 處理結果
     */
    async processWorkbook(workbook, progressCallback = () => { }) {
        return this.processMultipleWorkbooks([workbook], progressCallback);
    }

    /**
     * 處理多個工作簿
     * @param {Array} workbooks - Array of SheetJS workbooks
     * @param {Function} progressCallback - 進度回調函數
     * @returns {Object} 處理結果
     */
    async processMultipleWorkbooks(workbooks, progressCallback = () => { }) {
        console.log(`開始處理 ${workbooks.length} 個工作簿...`);
        console.log('配置:', this.config);

        let totalFiles = workbooks.length;
        this.results.processedSheets = 0; // 重置計數 (針對單次批次處理)

        for (let fileIndex = 0; fileIndex < totalFiles; fileIndex++) {
            const workbook = workbooks[fileIndex];
            const fileName = workbook.fileName || `File ${fileIndex + 1}`;
            const sheetCount = workbook.SheetNames.length;

            // 為每個 workbook 分配唯一 ID
            const workbookId = `WB_${fileIndex}_${Date.now()}`;

            // 計算最大頁面偏移量，決定步長
            let maxOffset = 0;
            if (this.config.cavityGroups) {
                for (let g = 1; g <= 6; g++) {
                    if (this.config.cavityGroups[g]) {
                        maxOffset = Math.max(maxOffset, this.config.cavityGroups[g].pageOffset || 0);
                    }
                }
            }
            const step = maxOffset + 1;

            // 遍歷所有工作表
            for (let i = 0; i < sheetCount; i += step) {
                const sheetName = workbook.SheetNames[i];
                const worksheet = workbook.Sheets[sheetName];

                try {
                    // 更新進度
                    progressCallback({
                        current: fileIndex + 1,
                        total: totalFiles,
                        message: `[${fileIndex + 1}/${totalFiles}] 處理: ${sheetName}`,
                        percent: Math.round(((fileIndex * sheetCount + i + 1) / (totalFiles * sheetCount)) * 100)
                    });

                    // 提取並彙整數據 (會自動加入 this.results.inspectionItems)
                    // 傳遞 workbookId 以區分不同檔案
                    await this.processWorksheet(workbook, worksheet, sheetName, i, workbookId, fileName);
                    this.results.processedSheets++;

                } catch (error) {
                    console.error(`處理工作表 ${sheetName} 時發生錯誤:`, error);
                    this.errorLogger.logError(`${fileName}: ${sheetName}`, error.message);
                }

                // 讓 UI 有機會更新
                await this.sleep(5);
            }

            // 如果是第一個工作簿，提取產品資訊與規格 (假設後續檔案格式相同)
            if (fileIndex === 0) {
                // 提取產品資訊
                this.results.productInfo = DataExtractor.extractProductInfo(workbook);
                // 提取規格
                await this.extractSpecifications(workbook, progressCallback);
            }
        }

        console.log('處理完成', this.results);
        return this.getResults();
    }

    /**
     * 處理單個工作表
     * @param {Object} workbook 
     * @param {Object} worksheet 
     * @param {string} sheetName 
     * @param {number} sheetIndex 
     * @param {string} workbookId - 工作簿唯一識別碼
     * @param {string} fileName - 檔案名稱
     */
    async processWorksheet(workbook, worksheet, sheetName, sheetIndex, workbookId, fileName) {
        // 獲取批號（使用工作表名稱作為批號）
        const batchName = sheetName;

        // 處理每個穴組
        for (let groupIndex = 1; groupIndex <= 6; groupIndex++) {
            const groupConfig = this.config.cavityGroups[groupIndex];

            if (!groupConfig || !groupConfig.cavityIdRange || !groupConfig.dataRange) {
                continue;
            }

            // 計算目標工作表索引
            const targetSheetIndex = sheetIndex + (groupConfig.pageOffset || 0);

            if (targetSheetIndex < 0 || targetSheetIndex >= workbook.SheetNames.length) {
                continue;
            }

            const targetWs = workbook.Sheets[workbook.SheetNames[targetSheetIndex]];

            // 從該穴組提取檢驗項目數據
            const items = DataExtractor.extractInspectionItemsFromGroup(targetWs, groupConfig);

            for (const item of items) {
                // 傳遞 workbookId 和 fileName 以區分不同檔案的批次
                this.addToResults(item.inspectionItem, batchName, item.data, workbookId, fileName);
            }
        }
    }

    // extractInspectionItemsFromGroup 已遷移至 DataExtractor.js

    /**
     * 將數據添加到結果
     * @param {string} inspectionItem 
     * @param {string} batchName 
     * @param {Object} data 
     * @param {string} workbookId - 工作簿唯一識別碼
     * @param {string} fileName - 檔案名稱（用於顯示）
     */
    addToResults(inspectionItem, batchName, data, workbookId, fileName) {
        if (!inspectionItem || Object.keys(data).length === 0) return;

        if (!this.results.inspectionItems[inspectionItem]) {
            this.results.inspectionItems[inspectionItem] = {
                batches: {},
                allCavities: new Set(),
                specification: null
            };
        }

        const item = this.results.inspectionItems[inspectionItem];

        // 使用複合鍵：workbookId + batchName
        // 這確保了：同檔案 + 基準名稱相同 = 合併；不同檔案 = 獨立保存
        const batchKey = `${workbookId}::${batchName}`;

        // 如果批次已存在（同一檔案內的相同批號），合併數據
        if (item.batches[batchKey]) {
            Object.assign(item.batches[batchKey].data, data);
        } else {
            // 新批次：儲存數據及元資訊
            item.batches[batchKey] = {
                data: { ...data },
                batchName: batchName,  // 原始批號名稱
                workbookId: workbookId, // 所屬檔案 ID
                fileName: fileName      // 檔案名稱（用於顯示）
            };
            this.results.totalBatches++;
        }

        // 記錄所有穴號
        for (const cavityId of Object.keys(data)) {
            item.allCavities.add(cavityId);
        }

        // 計算最大穴號
        const currentMaxCavity = Array.from(item.allCavities)
            .map(Number)
            .filter(n => !isNaN(n))
            .reduce((max, val) => Math.max(max, val), 0);

        // 如果配置中有模穴數，優先使用配置中的模穴數，否則使用發現的最大穴號
        const configCavityCount = parseInt(this.config.cavityCount) || 0;
        
        this.results.totalCavities = Math.max(
            this.results.totalCavities,
            currentMaxCavity,
            configCavityCount
        );
    }

    /**
     * 提取規格數據
     * @param {Object} workbook 
     * @param {Function} progressCallback 
     */
    async extractSpecifications(workbook, progressCallback) {
        progressCallback({
            current: 100,
            total: 100,
            message: '提取規格數據...',
            percent: 95
        });

        for (const itemName of Object.keys(this.results.inspectionItems)) {
            const spec = SpecificationExtractor.extractSpecification(workbook, itemName);
            this.results.inspectionItems[itemName].specification = spec;
        }
    }

    // extractProductInfo 已遷移至 DataExtractor.js

    /**
     * 獲取處理結果
     * @returns {Object}
     */
    getResults() {
        return {
            inspectionItems: this.results.inspectionItems,
            totalBatches: this.results.totalBatches,
            totalCavities: this.results.totalCavities,
            cavityCount: parseInt(this.config.cavityCount) || 0,
            processedSheets: this.results.processedSheets,
            productInfo: this.results.productInfo,
            itemCount: Object.keys(this.results.inspectionItems).length,
            errors: this.errorLogger.getErrors(),
            hasErrors: this.errorLogger.hasErrors()
        };
    }

    /**
     * 輔助函數：延遲
     * @param {number} ms 
     */
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = QIPProcessor;
}
