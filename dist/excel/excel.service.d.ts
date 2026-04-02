import * as XLSX from 'xlsx';
export declare class ExcelService {
    private brandColumns;
    private results;
    getBrandColumns(): {
        'Brand A': string[];
        'Brand B': string[];
        'Brand C': string[];
    };
    storeResult(id: string, data: any): void;
    getResult(id: string): any;
    filterBySubIds(fileBuffer: Buffer, subIds: string[], filterColumnName?: string): Promise<{
        sheets: any[];
        hasData: boolean;
        sourceWorksheet: XLSX.WorkSheet;
        headers: any[];
        range: XLSX.Range;
        workbook: XLSX.WorkBook;
    }>;
    createExcelFromData(sheetsData: any): Promise<Buffer>;
}
