/**
 * DataExtractor - 數據提取模組
 * 對應 VBA DataExtractor.bas
 */
class DataExtractor {

    /**
     * 提取所有生產批號
     * @param {Object} worksheet - SheetJS worksheet object
     * @returns {Array<string>}
     */
    static extractBatchNumbers(worksheet) {
        try {
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const startRow = DataValidator.getDataStartRow();
            const result = [];

            for (let r = startRow - 1; r <= range.e.r; r++) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: 0 });
                const cell = worksheet[cellAddr];

                if (cell && cell.v !== undefined && cell.v !== '') {
                    result.push(String(cell.v));
                }
            }

            return result;
        } catch (error) {
            console.error('DataExtractor.extractBatchNumbers error:', error);
            return [];
        }
    }

    /**
     * 提取每個批次的平均值
     * @param {Object} worksheet 
     * @returns {Array<number>}
     */
    static extractAverageValues(worksheet) {
        try {
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const startRow = DataValidator.getDataStartRow();
            const cavityCount = DataValidator.getCavityColumnCount(worksheet);
            const startCol = DataValidator.getCavityStartColumn();
            const result = [];

            if (cavityCount === 0) return [];

            for (let r = startRow - 1; r <= range.e.r; r++) {
                let sum = 0;
                let count = 0;

                for (let c = startCol - 1; c < startCol - 1 + cavityCount; c++) {
                    const cellAddr = XLSX.utils.encode_cell({ r: r, c: c });
                    const cell = worksheet[cellAddr];

                    if (cell && typeof cell.v === 'number') {
                        sum += cell.v;
                        count++;
                    } else if (cell && !isNaN(parseFloat(cell.v))) {
                        sum += parseFloat(cell.v);
                        count++;
                    }
                }

                result.push(count > 0 ? sum / count : 0);
            }

            return result;
        } catch (error) {
            console.error('DataExtractor.extractAverageValues error:', error);
            return [];
        }
    }

    /**
     * 提取特定穴號的所有數據
     * @param {Object} worksheet 
     * @param {number} cavityIndex - 穴號索引（1-based）
     * @returns {Array}
     */
    static extractCavityData(worksheet, cavityIndex) {
        try {
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const startRow = DataValidator.getDataStartRow();
            const colNum = DataValidator.getCavityStartColumn() + cavityIndex - 1;
            const result = [];

            for (let r = startRow - 1; r <= range.e.r; r++) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: colNum - 1 });
                const cell = worksheet[cellAddr];

                if (cell && typeof cell.v === 'number') {
                    result.push(cell.v);
                } else if (cell && !isNaN(parseFloat(cell.v))) {
                    result.push(parseFloat(cell.v));
                } else {
                    result.push(null);
                }
            }

            return result;
        } catch (error) {
            console.error('DataExtractor.extractCavityData error:', error);
            return [];
        }
    }

    /**
     * 獲取所有穴號的標題
     * @param {Object} worksheet 
     * @returns {Array<string>}
     */
    static getCavityHeaders(worksheet) {
        try {
            const cavityCount = DataValidator.getCavityColumnCount(worksheet);
            if (cavityCount === 0) return [];

            const startCol = DataValidator.getCavityStartColumn();
            const headers = [];

            for (let i = 0; i < cavityCount; i++) {
                const cellAddr = XLSX.utils.encode_cell({ r: 0, c: startCol - 1 + i });
                const cell = worksheet[cellAddr];

                if (cell && cell.v !== undefined && String(cell.v).trim() !== '') {
                    headers.push(String(cell.v));
                } else {
                    headers.push(`穴${i + 1}`);
                }
            }

            return headers;
        } catch (error) {
            console.error('DataExtractor.getCavityHeaders error:', error);
            return [];
        }
    }

    /**
     * 從指定範圍提取數據
     * @param {Object} worksheet 
     * @param {string} rangeStr - 如 "K3:R10"
     * @returns {Array<Array>}
     */
    static extractDataFromRange(worksheet, rangeStr) {
        try {
            const range = XLSX.utils.decode_range(rangeStr);
            const result = [];

            for (let r = range.s.r; r <= range.e.r; r++) {
                const row = [];
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cellAddr = XLSX.utils.encode_cell({ r: r, c: c });
                    const cell = worksheet[cellAddr];
                    row.push(cell ? cell.v : null);
                }
                result.push(row);
            }

            return result;
        } catch (error) {
            console.error('DataExtractor.extractDataFromRange error:', error);
            return [];
        }
    }

    /**
     * 清理儲存格值
     * @param {*} value 
     * @returns {string}
     */
    static cleanCellValue(value) {
        if (value === null || value === undefined) return '';
        let str = String(value).trim();
        // 移除常見的非數字字符但保留小數點和負號
        str = str.replace(/[^\d.\-]/g, '');
        return str;
    }

    /**
     * 從穴組提取檢驗項目數據
     * @param {Object} worksheet 
     * @param {Object} groupConfig 
     * @returns {Array}
     */
    static extractInspectionItemsFromGroup(worksheet, groupConfig) {
        const result = [];

        try {
            const idRangeParsed = DataValidator.parseRangeString(groupConfig.cavityIdRange);
            const dataRangeParsed = DataValidator.parseRangeString(groupConfig.dataRange);

            if (!idRangeParsed || !dataRangeParsed) {
                return result;
            }

            // 遍歷數據範圍的每一行（每行可能是不同的檢驗項目）
            for (let rowOffset = 0; rowOffset <= dataRangeParsed.endRow - dataRangeParsed.startRow; rowOffset++) {
                const dataRow = dataRangeParsed.startRow - 1 + rowOffset; // 0-indexed
                const data = {};
                let itemName = '';

                // 提取檢驗項目名稱（從 A 或 B 欄）
                for (let c = 0; c < dataRangeParsed.startCol - 1; c++) {
                    const cellAddr = XLSX.utils.encode_cell({ r: dataRow, c: c });
                    const cell = worksheet[cellAddr];

                    if (cell) {
                        // 優先使用格式化後的內容 (cell.w)，以處理 (1), (2) 等會被 Excel 視為數字的格式
                        const cellValue = cell.w || String(cell.v || '').trim();
                        if (cellValue && cellValue !== '0') {
                            // 如果包含括號如 (1)，isNumericString 會回傳 false，從而正確識別為檢驗項目
                            if (!this.isNumericString(cellValue)) {
                                itemName = cellValue;
                                break;
                            }
                        }
                    }
                }

                // 提取穴號數據
                const idRow = idRangeParsed.startRow - 1; // 0-indexed
                let hasData = false;

                for (let colOffset = 0; colOffset < idRangeParsed.endCol - idRangeParsed.startCol + 1; colOffset++) {
                    const col = idRangeParsed.startCol - 1 + colOffset;

                    // 獲取穴號 ID
                    const idCellAddr = XLSX.utils.encode_cell({ r: idRow, c: col });
                    const idCell = worksheet[idCellAddr];
                    let cavityId = idCell ? String(idCell.v || '').trim() : '';

                    if (!cavityId) continue;

                    // 提取穴號數字
                    const numMatch = cavityId.match(/\d+/);
                    if (numMatch) {
                        cavityId = numMatch[0];
                    }

                    // 獲取數據
                    const dataCellAddr = XLSX.utils.encode_cell({ r: dataRow, c: col });
                    const dataCell = worksheet[dataCellAddr];

                    if (dataCell && dataCell.v !== undefined) {
                        const cleanValue = this.cleanCellValue(dataCell.v);
                        if (cleanValue !== '' && !isNaN(parseFloat(cleanValue))) {
                            data[cavityId] = parseFloat(cleanValue);
                            hasData = true;
                        }
                    }
                }

                if (hasData && itemName) {
                    result.push({
                        inspectionItem: itemName,
                        data: data
                    });
                }
            }

        } catch (error) {
            console.error('DataExtractor.extractInspectionItemsFromGroup error:', error);
        }

        return result;
    }

    /**
     * 提取產品資訊 (Product Name & Measurement Unit)
     * @param {Object} workbook 
     * @returns {Object}
     */
    static extractProductInfo(workbook) {
        let productName = '';
        let measurementUnit = '';

        for (const sheetName of workbook.SheetNames) {
            const ws = workbook.Sheets[sheetName];

            // 提取產品名稱 (P2 或 P3)
            if (!productName) {
                const cellP2 = ws['P2'];
                if (cellP2 && cellP2.v) productName = String(cellP2.v).trim();

                if (!productName) {
                    const cellP3 = ws['P3'];
                    if (cellP3 && cellP3.v) productName = String(cellP3.v).trim();
                }

                // 掃描 P2:V3 區域
                if (!productName) {
                    for (let r = 1; r <= 2; r++) {
                        for (let c = 15; c <= 21; c++) {
                            const addr = XLSX.utils.encode_cell({ r, c });
                            const cell = ws[addr];
                            if (cell && cell.v && String(cell.v).trim() !== '0' && String(cell.v).trim() !== 'False') {
                                productName = String(cell.v).trim();
                                break;
                            }
                        }
                        if (productName) break;
                    }
                }
            }

            // 提取測量單位 (W23)
            if (!measurementUnit) {
                const cellW23 = ws['W23'];
                if (cellW23 && cellW23.v) {
                    let val = String(cellW23.v).trim();
                    // 移除 "單位：" 前綴
                    val = val.replace(/單位[:：]/g, '').trim();
                    measurementUnit = val;
                }
            }

            if (productName && measurementUnit) break;
        }

        return { productName, measurementUnit };
    }

    /**
     * 檢查字串是否為數字
     * @param {string} str 
     * @returns {boolean}
     */
    static isNumericString(str) {
        return !isNaN(parseFloat(str)) && isFinite(str);
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DataExtractor;
}
