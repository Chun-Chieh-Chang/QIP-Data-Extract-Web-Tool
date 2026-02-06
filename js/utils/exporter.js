/**
 * ExcelExporter - Excel 輸出模組
 * 使用 SheetJS 生成 Excel 檔案，輸出格式與 VBA 版本相同或更優
 */
class ExcelExporter {
    constructor() {
        this.workbook = XLSX.utils.book_new();
        this.workbook.Props = {
            Title: 'QIP 數據提取結果',
            Author: 'QIP Data Extract Tool',
            CreatedDate: new Date()
        };
    }

    /**
     * 從處理結果創建 Excel
     * @param {Object} results - QIPProcessor 的處理結果
     * @param {string} productCode - 產品品號
     * @returns {ExcelExporter}
     */
    createFromResults(results, productCode = '') {
        console.log('開始創建 Excel...', results);

        for (const [itemName, itemData] of Object.entries(results.inspectionItems)) {
            this.addInspectionSheet(itemName, itemData, productCode, results.productInfo);
        }

        return this;
    }

    /**
     * 添加檢驗項目工作表
     * @param {string} sheetName - 工作表名稱（檢驗項目）
     * @param {Object} itemData - 項目數據 { batches, allCavities, specification }
     * @param {string} productCode - 產品品號
     */
    /**
     * 添加檢驗項目工作表
     * @param {string} sheetName - 工作表名稱（檢驗項目）
     * @param {Object} itemData - 項目數據 { batches, allCavities, specification }
     * @param {string} productCode - 產品品號
     * @param {Object} productInfo - 產品資訊 { productName, measurementUnit }
     */
    addInspectionSheet(sheetName, itemData, productCode = '', productInfo = null) {
        // 清理工作表名稱
        const cleanName = this.cleanSheetName(sheetName);

        // 獲取並排序所有穴號
        const cavities = Array.from(itemData.allCavities)
            .map(Number)
            .sort((a, b) => a - b);

        // 獲取批號列表（現在是複合鍵）
        const batchKeys = Object.keys(itemData.batches);

        // 提取批次資訊（包含原始批號名稱和檔案資訊）
        const batchEntries = batchKeys.map(key => ({
            key: key,
            ...itemData.batches[key]
        }));

        // 構建數據陣列
        const data = [];

        // 1. 欄位佈置 (Column Layout)
        // A: Target, B: USL, C: LSL, D: 生產批號, E+: 穴號
        const headerRow = ['Target', 'USL', 'LSL', '生產批號'];
        for (const cavityNum of cavities) {
            headerRow.push(`${cavityNum}號穴`);
        }
        data.push(headerRow);

        // 2. 資料排列規則 (Data Arrangement) & 3. 產品資訊 (Fixed Metadata Position)
        // 確保至少展開至第 6 行以顯示完整的產品資訊標籤 (Row 1 Header + min 5 data/label rows)
        const rowCount = Math.max(batchEntries.length + 1, 6);

        for (let i = 0; i < rowCount - 1; i++) {
            const rowIdx = i + 2; // 1-indexed Excel row number
            const batchIdx = i;   // 0-indexed batch array index
            const row = new Array(4 + cavities.length).fill(''); // A, B, C, D, E...

            // Specs 僅在第 2 行 (Row 2) 出現
            if (rowIdx === 2) {
                if (itemData.specification && itemData.specification.isValid) {
                    row[0] = itemData.specification.target;
                    row[1] = itemData.specification.usl;
                    row[2] = itemData.specification.lsl;
                } else {
                    row[0] = '未設定';
                    row[1] = '未設定';
                    row[2] = '未設定';
                }
            }

            // 產品資訊固定在 Row 5 & Row 6 的 A, B 欄
            if (rowIdx === 5) {
                row[0] = 'ProductName';
                row[1] = productInfo ? (productInfo.productName || '') : '';
            } else if (rowIdx === 6) {
                row[0] = 'MeasurementUnit';
                row[1] = productInfo ? (productInfo.measurementUnit || '') : '';
            }

            // 批次數據從 Col D (index 3) 開始填充
            if (batchIdx < batchEntries.length) {
                const batchEntry = batchEntries[batchIdx];
                // 使用原始批號名稱作為顯示（如果有多個同名批次，可選擇性添加檔案名稱）
                const displayName = batchEntry.batchName;
                row[3] = displayName;

                // 從 batch.data 中提取穴號數據
                const batchData = batchEntry.data;
                for (let j = 0; j < cavities.length; j++) {
                    const cavityNum = cavities[j];
                    const value = batchData[String(cavityNum)];
                    row[4 + j] = value !== undefined ? value : '';
                }
            }

            data.push(row);
        }

        // 創建工作表
        const worksheet = XLSX.utils.aoa_to_sheet(data);

        // 設置列寬 (A:12, B:12, C:12, D:15, E+:10)
        const colWidths = [
            { wch: 12 },  // Target
            { wch: 12 },  // USL
            { wch: 12 },  // LSL
            { wch: 15 }   // 生產批號
        ];
        for (let i = 0; i < cavities.length; i++) {
            colWidths.push({ wch: 10 });
        }
        worksheet['!cols'] = colWidths;

        // 設置儲存格樣式
        this.setHeaderStyles(worksheet, headerRow.length);
        this.setSpecificationStyles(worksheet, itemData.specification);

        // 添加到工作簿
        XLSX.utils.book_append_sheet(this.workbook, worksheet, cleanName);

        console.log(`創建工作表: ${cleanName}, 批次數: ${batchEntries.length}, 穴數: ${cavities.length}`);
    }

    /**
     * 設置標題行樣式
     * @param {Object} worksheet 
     * @param {number} colCount 
     */
    setHeaderStyles(worksheet, colCount) {
        // SheetJS 免費版樣式支援有限，這裡主要設置數字格式
        // 使用 xlsx-style 或 exceljs 可以獲得完整樣式支援
        for (let c = 0; c < colCount; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r: 0, c: c });
            if (worksheet[cellAddr]) {
                // 標記為標題（供後續處理）
                worksheet[cellAddr].s = {
                    font: { bold: true },
                    fill: { fgColor: { rgb: '92D050' } },
                    alignment: { horizontal: 'center' }
                };
            }
        }
    }

    /**
     * 設置規格行樣式
     * @param {Object} worksheet 
     * @param {Object} specification 
     */
    setSpecificationStyles(worksheet, specification) {
        // 設置規格數字格式 (A2, B2, C2)
        for (let c = 0; c <= 2; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r: 1, c: c });
            if (worksheet[cellAddr] && typeof worksheet[cellAddr].v === 'number') {
                worksheet[cellAddr].z = '0.0000';
            }
        }
    }

    /**
     * 清理工作表名稱
     * @param {string} name 
     * @returns {string}
     */
    cleanSheetName(name) {
        // 移除不允許的字符
        let result = name.replace(/[\\/:*?"<>|]/g, '_');
        // 限制長度（Excel 最多31字符）
        if (result.length > 31) {
            result = result.substring(0, 31);
        }
        return result.trim() || '未命名項目';
    }

    /**
     * 導出 Excel 檔案
     * @param {string} filename - 檔案名稱（不含副檔名）
     */
    export(filename = 'QIP_數據提取結果') {
        const fullFilename = `${filename}.xlsx`;
        XLSX.writeFile(this.workbook, fullFilename);
        console.log(`Excel 檔案已導出: ${fullFilename}`);
    }

    /**
     * 獲取 Excel 二進制數據
     * @returns {ArrayBuffer}
     */
    getArrayBuffer() {
        return XLSX.write(this.workbook, {
            bookType: 'xlsx',
            type: 'array'
        });
    }

    /**
     * 獲取 Blob 物件（用於下載）
     * @returns {Blob}
     */
    getBlob() {
        const buffer = this.getArrayBuffer();
        return new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
    }

    /**
     * 觸發下載
     * @param {string} filename 
     */
    download(filename = 'QIP_數據提取結果') {
        const blob = this.getBlob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${filename}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        console.log(`開始下載: ${filename}.xlsx`);
    }

    /**
     * 獲取工作表數量
     * @returns {number}
     */
    getSheetCount() {
        return this.workbook.SheetNames.length;
    }

    /**
     * 重置（創建新工作簿）
     */
    reset() {
        this.workbook = XLSX.utils.book_new();
        this.workbook.Props = {
            Title: 'QIP 數據提取結果',
            Author: 'QIP Data Extract Tool',
            CreatedDate: new Date()
        };
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ExcelExporter;
}
