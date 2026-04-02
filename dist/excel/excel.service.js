"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelService = void 0;
const common_1 = require("@nestjs/common");
const XLSX = require("xlsx");
let ExcelService = class ExcelService {
    constructor() {
        this.brandColumns = {
            'Brand A': ['Model', 'Price', 'Subsidy A', 'Stock', 'Color'],
            'Brand B': ['Model', 'Price', 'Subsidy B', 'Subsidy B2', 'Stock'],
            'Brand C': ['Model', 'Price', 'Subsidy C', 'Warranty', 'Stock'],
        };
        this.results = new Map();
    }
    getBrandColumns() {
        return this.brandColumns;
    }
    storeResult(id, data) {
        this.results.set(id, data);
        setTimeout(() => this.results.delete(id), 10 * 60 * 1000);
    }
    getResult(id) {
        return this.results.get(id);
    }
    async filterBySubIds(fileBuffer, subIds, filterColumnName) {
        const workbook = XLSX.read(fileBuffer, {
            type: 'buffer',
            cellFormula: false,
            cellHTML: false,
            cellNF: false,
            cellStyles: false,
            cellDates: false,
            raw: false
        });
        const sheetName = workbook.SheetNames[0];
        const sourceWorksheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sourceWorksheet['!ref']);
        const headers = [];
        const headerRow = range.s.r;
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
            const cell = sourceWorksheet[cellAddress];
            headers.push(cell ? String(cell.v) : '');
        }
        let filterColIndex = -1;
        if (filterColumnName) {
            filterColIndex = headers.indexOf(filterColumnName);
        }
        else {
            const commonNames = ['SubID', 'subid', 'subId', 'SUBID', 'ExtraParam', 'extraParam', 'extParam2', 'extparam2'];
            for (const name of commonNames) {
                filterColIndex = headers.indexOf(name);
                if (filterColIndex !== -1)
                    break;
            }
        }
        if (filterColIndex === -1) {
            throw new Error('Filter column not found in the Excel file');
        }
        const rowsBySubId = new Map();
        for (let row = headerRow + 1; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: filterColIndex });
            const cell = sourceWorksheet[cellAddress];
            const subIdValue = cell ? String(cell.v) : '';
            if (subIdValue && subIdValue !== 'undefined' && subIdValue !== 'null' && subIdValue.trim() !== '') {
                if (!rowsBySubId.has(subIdValue)) {
                    rowsBySubId.set(subIdValue, []);
                }
                rowsBySubId.get(subIdValue).push(row);
            }
        }
        const result = {
            sheets: [],
            hasData: false,
            sourceWorksheet,
            headers,
            range,
            workbook
        };
        let subIdsToProcess = [];
        if (!subIds || subIds.length === 0) {
            subIdsToProcess = Array.from(rowsBySubId.entries())
                .filter(([_, rows]) => rows.length > 10)
                .map(([subId, _]) => subId);
        }
        else {
            subIdsToProcess = subIds;
        }
        for (const subId of subIdsToProcess) {
            const rows = rowsBySubId.get(subId);
            if (rows && rows.length > 0) {
                const sheetLabel = `${filterColumnName || 'extParam2'}_${subId}`;
                const previewData = [];
                for (const sourceRow of rows) {
                    const rowData = {};
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: sourceRow, c: col });
                        const cell = sourceWorksheet[cellAddress];
                        const header = headers[col - range.s.c];
                        if (cell && cell.w !== undefined && cell.w !== null) {
                            rowData[header] = cell.w;
                        }
                        else if (cell && cell.v !== undefined && cell.v !== null) {
                            rowData[header] = String(cell.v);
                        }
                        else {
                            rowData[header] = '';
                        }
                    }
                    previewData.push(rowData);
                }
                result.sheets.push({
                    name: sheetLabel,
                    rows: rows,
                    data: previewData,
                    subId: subId
                });
                result.hasData = true;
            }
        }
        if (!result.hasData) {
            result.sheets.push({
                name: 'No Results',
                rows: [],
                data: [{ Message: 'No matching records found for the provided IDs or no SubIDs with more than 10 records' }],
                subId: null
            });
        }
        return result;
    }
    async createExcelFromData(sheetsData) {
        const newWorkbook = XLSX.utils.book_new();
        const sourceWorksheet = sheetsData.sourceWorksheet;
        const range = sheetsData.range;
        for (const sheet of sheetsData.sheets) {
            if (sheet.rows && sheet.rows.length > 0) {
                const newWorksheet = {};
                if (sourceWorksheet['!cols']) {
                    newWorksheet['!cols'] = JSON.parse(JSON.stringify(sourceWorksheet['!cols']));
                }
                if (sourceWorksheet['!rows']) {
                    newWorksheet['!rows'] = JSON.parse(JSON.stringify(sourceWorksheet['!rows']));
                }
                if (sourceWorksheet['!merges']) {
                    newWorksheet['!merges'] = JSON.parse(JSON.stringify(sourceWorksheet['!merges']));
                }
                const copyCellWithFormat = (sourceCell) => {
                    if (!sourceCell)
                        return null;
                    const newCell = {
                        v: sourceCell.v,
                        t: sourceCell.t
                    };
                    if (sourceCell.w !== undefined)
                        newCell.w = sourceCell.w;
                    if (sourceCell.z !== undefined)
                        newCell.z = sourceCell.z;
                    if (sourceCell.f !== undefined)
                        newCell.f = sourceCell.f;
                    if (sourceCell.F !== undefined)
                        newCell.F = sourceCell.F;
                    if (sourceCell.r !== undefined)
                        newCell.r = sourceCell.r;
                    if (sourceCell.h !== undefined)
                        newCell.h = sourceCell.h;
                    if (sourceCell.c !== undefined)
                        newCell.c = JSON.parse(JSON.stringify(sourceCell.c));
                    if (sourceCell.s !== undefined)
                        newCell.s = JSON.parse(JSON.stringify(sourceCell.s));
                    if (sourceCell.l !== undefined)
                        newCell.l = JSON.parse(JSON.stringify(sourceCell.l));
                    return newCell;
                };
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const sourceCellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
                    const targetCellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                    const sourceCell = sourceWorksheet[sourceCellAddress];
                    if (sourceCell) {
                        newWorksheet[targetCellAddress] = copyCellWithFormat(sourceCell);
                    }
                }
                let targetRow = 1;
                for (const sourceRow of sheet.rows) {
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const sourceCellAddress = XLSX.utils.encode_cell({ r: sourceRow, c: col });
                        const targetCellAddress = XLSX.utils.encode_cell({ r: targetRow, c: col });
                        const sourceCell = sourceWorksheet[sourceCellAddress];
                        if (sourceCell) {
                            newWorksheet[targetCellAddress] = copyCellWithFormat(sourceCell);
                        }
                    }
                    targetRow++;
                }
                newWorksheet['!ref'] = XLSX.utils.encode_range({
                    s: { r: 0, c: range.s.c },
                    e: { r: targetRow - 1, c: range.e.c }
                });
                XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheet.name);
            }
            else {
                const noResultSheet = XLSX.utils.json_to_sheet([{ Message: 'No matching records found' }]);
                XLSX.utils.book_append_sheet(newWorkbook, noResultSheet, sheet.name);
            }
        }
        return XLSX.write(newWorkbook, {
            type: 'buffer',
            bookType: 'xlsx',
            cellStyles: true
        });
    }
};
exports.ExcelService = ExcelService;
exports.ExcelService = ExcelService = __decorate([
    (0, common_1.Injectable)()
], ExcelService);
//# sourceMappingURL=excel.service.js.map