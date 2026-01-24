/**
 * SpecificationExtractor - 規格數據提取模組
 * 對應 VBA SpecificationExtractor.bas
 */
class SpecificationExtractor {

    /**
     * 規格數據結構
     */
    static createSpecificationData() {
        return {
            symbol: '',
            nominalValue: 0,
            upperTolerance: 0,
            lowerTolerance: 0,
            usl: 0,
            lsl: 0,
            target: 0,
            isValid: false
        };
    }

    /**
     * 尋找包含規格數據的工作表
     * @param {Object} workbook - SheetJS workbook
     * @returns {Object|null} worksheet
     */
    static findSpecificationWorksheet(workbook) {
        try {
            const specKeywords = ['規格', 'spec', 'specification', '檢驗標準', '檢驗規格'];

            // 優先尋找包含規格關鍵字的工作表
            for (const sheetName of workbook.SheetNames) {
                const lowerName = sheetName.toLowerCase();
                for (const keyword of specKeywords) {
                    if (lowerName.includes(keyword.toLowerCase())) {
                        return workbook.Sheets[sheetName];
                    }
                }
            }

            // 尋找包含規格數據的工作表
            for (const sheetName of workbook.SheetNames) {
                const ws = workbook.Sheets[sheetName];
                if (this.hasSpecificationData(ws)) {
                    return ws;
                }
            }

            // 回退到第一個工作表
            if (workbook.SheetNames.length > 0) {
                return workbook.Sheets[workbook.SheetNames[0]];
            }

            return null;
        } catch (error) {
            console.error('SpecificationExtractor.findSpecificationWorksheet error:', error);
            return null;
        }
    }

    /**
     * 檢查工作表是否包含規格數據
     * @param {Object} worksheet 
     * @returns {boolean}
     */
    static hasSpecificationData(worksheet) {
        try {
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const maxRow = Math.min(range.e.r, 99); // 只檢查前100行

            for (let r = 0; r <= maxRow; r++) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: 0 });
                const cell = worksheet[cellAddr];

                if (cell && cell.v !== undefined) {
                    const value = String(cell.v).trim().toLowerCase();
                    if (value.includes('檢驗項目') || value.includes('inspection') ||
                        value.includes('規格') || value.includes('spec')) {
                        return true;
                    }
                }
            }

            return false;
        } catch (error) {
            console.error('SpecificationExtractor.hasSpecificationData error:', error);
            return false;
        }
    }

    /**
     * 根據檢驗項目查找規格數據
     * @param {Object} worksheet 
     * @param {string} itemName - 檢驗項目名稱
     * @returns {Object} SpecificationData
     */
    static findSpecificationByItem(worksheet, itemName) {
        try {
            const spec = this.createSpecificationData();

            if (!itemName || !worksheet) return spec;

            // 清理檢驗項目名稱
            const cleanItemName = itemName.replace(/[()]/g, '').trim();
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const maxRow = Math.min(range.e.r, 99);

            for (let r = 0; r <= maxRow; r++) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: 0 });
                const cell = worksheet[cellAddr];

                if (cell) {
                    const cellRawValue = cell.w || String(cell.v || '').trim();
                    const cellCleanValue = cellRawValue.replace(/[()]/g, '').trim();

                    // 檢查是否匹配 (支援精確匹配或去括號後的匹配)
                    if (cellRawValue === itemName ||
                        cellCleanValue === cleanItemName ||
                        cellRawValue.includes(itemName) ||
                        itemName.includes(cellRawValue)) {

                        const rowSpec = this.readSpecificationFromRow(worksheet, r);
                        if (rowSpec.isValid) {
                            return rowSpec;
                        }
                    }
                }
            }

            return spec;
        } catch (error) {
            console.error('SpecificationExtractor.findSpecificationByItem error:', error);
            return this.createSpecificationData();
        }
    }

    /**
     * 從指定行讀取規格數據
     * @param {Object} worksheet 
     * @param {number} row - 0-indexed row
     * @returns {Object} SpecificationData
     */
    static readSpecificationFromRow(worksheet, row) {
        const spec = this.createSpecificationData();

        try {
            // 使用 DataValidator 定義的欄位索引 (0-indexed)
            const cols = DataValidator.SPEC_COLUMNS;

            // 檢查是否有檢測工具代碼（識別規格行）
            const toolCell = worksheet[XLSX.utils.encode_cell({ r: row, c: cols.TOOL })];
            const toolVal = toolCell ? String(toolCell.v || '').trim() : '';
            if (!toolVal) return spec;

            // 讀取符號
            const symbolCell = worksheet[XLSX.utils.encode_cell({ r: row, c: cols.SYMBOL })];
            spec.symbol = symbolCell ? String(symbolCell.v || '').trim() : '';

            // 讀取基準值（採用與 VBA 相同的搜尋順序：E(row) -> F(row) -> E(row+1) -> F(row+1)）
            let nominalValue = null;
            const searchTargets = [
                { r: row, c: cols.NOMINAL_1 }, { r: row, c: cols.NOMINAL_2 },
                { r: row + 1, c: cols.NOMINAL_1 }, { r: row + 1, c: cols.NOMINAL_2 }
            ];

            for (const target of searchTargets) {
                const cell = worksheet[XLSX.utils.encode_cell(target)];
                // 優先使用 w (formatted text) 進行數值檢查，避免 J 等被誤判
                const val = cell ? (cell.t === 'n' ? cell.v : parseFloat(String(cell.w || cell.v).trim())) : NaN;
                if (!isNaN(val) && val !== null && val !== '') {
                    nominalValue = val;
                    break;
                }
            }

            if (nominalValue === null) return spec;

            spec.nominalValue = nominalValue;
            spec.target = nominalValue;

            // 讀取上公差 (第一行 G, H 欄)
            const upperSignCell = worksheet[XLSX.utils.encode_cell({ r: row, c: cols.UPPER_SIGN })];
            const upperTolCell = worksheet[XLSX.utils.encode_cell({ r: row, c: cols.UPPER_TOL })];
            const upperSign = upperSignCell ? String(upperSignCell.v || '').trim() : '+';
            let upperTolVal = (upperTolCell && !isNaN(parseFloat(upperTolCell.v))) ? Math.abs(parseFloat(upperTolCell.v)) : 0;

            // 讀取下公差 (第二行 G, H 欄)
            const lowerSignCell = worksheet[XLSX.utils.encode_cell({ r: row + 1, c: cols.LOWER_SIGN })];
            const lowerTolCell = worksheet[XLSX.utils.encode_cell({ r: row + 1, c: cols.LOWER_TOL })];
            const lowerSign = lowerSignCell ? String(lowerSignCell.v || '').trim() : '-';
            let lowerTolVal = (lowerTolCell && !isNaN(parseFloat(lowerTolCell.v))) ? Math.abs(parseFloat(lowerTolCell.v)) : 0;

            // 根據符號計算實際的偏移量 (Offsets)
            // 理解符號意義：+ 代表增加，- 代表減少，± 代表雙向
            let upperOffset = 0;
            let lowerOffset = 0;

            if (upperSign === '±') {
                upperOffset = upperTolVal;
                lowerOffset = -upperTolVal;
            } else {
                // 上公差行：如果是 + 符號則為正偏移，如果是 - 符號則為負偏移
                upperOffset = (upperSign === '-') ? -upperTolVal : upperTolVal;
                // 下公差行：如果是 - 符號則為負偏移，如果是 + 符號則為正偏移
                lowerOffset = (lowerSign === '+') ? lowerTolVal : -lowerTolVal;
            }

            spec.upperTolerance = Math.abs(upperOffset);
            spec.lowerTolerance = Math.abs(lowerOffset);

            // 計算 USL 和 LSL
            // 邏輯上：上限 = 基準 + 偏移1, 下限 = 基準 + 偏移2
            spec.usl = spec.nominalValue + upperOffset;
            spec.lsl = spec.nominalValue + lowerOffset;

            // 確保 usl 永遠是大於等於 lsl 的那個
            if (spec.usl < spec.lsl) {
                const temp = spec.usl;
                spec.usl = spec.lsl;
                spec.lsl = temp;
            }

            spec.isValid = true;

            return spec;
        } catch (error) {
            console.error('SpecificationExtractor.readSpecificationFromRow error:', error);
            return spec;
        }
    }

    /**
     * 嘗試從工作表自動提取規格
     * @param {Object} workbook 
     * @param {string} inspectionItem 
     * @returns {Object} SpecificationData
     */
    static extractSpecification(workbook, inspectionItem) {
        try {
            const specWs = this.findSpecificationWorksheet(workbook);
            if (!specWs) {
                console.log('未找到規格工作表');
                return this.createSpecificationData();
            }

            const spec = this.findSpecificationByItem(specWs, inspectionItem);
            if (spec.isValid) {
                console.log(`成功提取規格: ${inspectionItem}`, spec);
            } else {
                console.log(`未找到規格: ${inspectionItem}`);
            }

            return spec;
        } catch (error) {
            console.error('SpecificationExtractor.extractSpecification error:', error);
            return this.createSpecificationData();
        }
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SpecificationExtractor;
}
