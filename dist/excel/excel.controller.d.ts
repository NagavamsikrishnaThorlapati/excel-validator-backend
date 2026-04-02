import { Response } from 'express';
import { ExcelService } from './excel.service';
export declare class ExcelController {
    private readonly excelService;
    constructor(excelService: ExcelService);
    getBrands(): {
        'Brand A': string[];
        'Brand B': string[];
        'Brand C': string[];
    };
    uploadAndFilter(file: Express.Multer.File, subIds: string, filterColumnName: string): Promise<{
        resultId: string;
        sheets: any[];
        hasData: boolean;
        sourceWorksheet: import("xlsx").WorkSheet;
        headers: any[];
        range: import("xlsx").Range;
        workbook: import("xlsx").WorkBook;
    }>;
    downloadExcel(resultId: string, res: Response): Promise<void>;
    downloadSheet(resultId: string, sheetIndex: number, res: Response): Promise<void>;
}
