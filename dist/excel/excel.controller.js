"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelController = void 0;
const common_1 = require("@nestjs/common");
const platform_express_1 = require("@nestjs/platform-express");
const excel_service_1 = require("./excel.service");
let ExcelController = class ExcelController {
    constructor(excelService) {
        this.excelService = excelService;
    }
    getBrands() {
        return this.excelService.getBrandColumns();
    }
    async uploadAndFilter(file, subIds, filterColumnName) {
        let subIdArray = [];
        if (subIds && subIds.trim()) {
            subIdArray = subIds.split(',').map(id => id.trim());
        }
        const result = await this.excelService.filterBySubIds(file.buffer, subIdArray, filterColumnName);
        const resultId = Date.now().toString();
        this.excelService.storeResult(resultId, result);
        return { ...result, resultId };
    }
    async downloadExcel(resultId, res) {
        const data = this.excelService.getResult(resultId);
        if (!data) {
            res.status(404).send('Result not found');
            return;
        }
        const buffer = await this.excelService.createExcelFromData(data);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=filtered_results.xlsx');
        res.send(buffer);
    }
    async downloadSheet(resultId, sheetIndex, res) {
        const data = this.excelService.getResult(resultId);
        if (!data) {
            res.status(404).json({ error: 'Result not found', message: 'The result may have expired or the file needs to be uploaded again' });
            return;
        }
        if (!data.sheets[sheetIndex]) {
            console.error('Sheet not found at index:', sheetIndex);
            res.status(404).json({ error: 'Sheet not found', message: `Sheet at index ${sheetIndex} does not exist` });
            return;
        }
        const singleSheetData = {
            sheets: [data.sheets[sheetIndex]],
            hasData: true,
            sourceWorksheet: data.sourceWorksheet,
            headers: data.headers,
            range: data.range,
            workbook: data.workbook
        };
        const buffer = await this.excelService.createExcelFromData(singleSheetData);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${data.sheets[sheetIndex].name}.xlsx`);
        res.send(buffer);
    }
};
exports.ExcelController = ExcelController;
__decorate([
    (0, common_1.Get)('brands'),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", []),
    __metadata("design:returntype", void 0)
], ExcelController.prototype, "getBrands", null);
__decorate([
    (0, common_1.Post)('upload'),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Body)('subIds')),
    __param(2, (0, common_1.Body)('filterColumnName')),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, String, String]),
    __metadata("design:returntype", Promise)
], ExcelController.prototype, "uploadAndFilter", null);
__decorate([
    (0, common_1.Post)('download'),
    __param(0, (0, common_1.Body)('resultId')),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [String, Object]),
    __metadata("design:returntype", Promise)
], ExcelController.prototype, "downloadExcel", null);
__decorate([
    (0, common_1.Post)('download-sheet'),
    __param(0, (0, common_1.Body)('resultId')),
    __param(1, (0, common_1.Body)('sheetIndex')),
    __param(2, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [String, Number, Object]),
    __metadata("design:returntype", Promise)
], ExcelController.prototype, "downloadSheet", null);
exports.ExcelController = ExcelController = __decorate([
    (0, common_1.Controller)('excel'),
    __metadata("design:paramtypes", [excel_service_1.ExcelService])
], ExcelController);
//# sourceMappingURL=excel.controller.js.map