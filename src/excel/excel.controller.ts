import { Controller, Post, Get, UploadedFile, UseInterceptors, Body, Res } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { ExcelService } from './excel.service';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Get('brands')
  getBrands() {
    return this.excelService.getBrandColumns();
  }

  @Post('upload')
  @UseInterceptors(FileInterceptor('file'))
  async uploadAndFilter(@UploadedFile() file: Express.Multer.File, @Body('subIds') subIds: string, @Body('filterColumnName') filterColumnName: string) {

    let subIdArray = [];
    if (subIds && subIds.trim()) {
      subIdArray = subIds.split(',').map(id => id.trim());
    }
    
    const result = await this.excelService.filterBySubIds(file.buffer, subIdArray, filterColumnName);
    const resultId = Date.now().toString();
    this.excelService.storeResult(resultId, result);    
    return { ...result, resultId };
  }

  @Post('download')
  async downloadExcel(@Body('resultId') resultId: string, @Res() res: Response) {
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

  @Post('download-sheet')
  async downloadSheet(@Body('resultId') resultId: string, @Body('sheetIndex') sheetIndex: number, @Res() res: Response ) {    
    const data = this.excelService.getResult(resultId);
    if (!data) {
      res.status(404).json({ error: 'Result not found', message: 'The result may have expired or the file needs to be uploaded again'});
      return;
    }
    
    if (!data.sheets[sheetIndex]) {
      console.error('Sheet not found at index:', sheetIndex);
      res.status(404).json({ error: 'Sheet not found', message: `Sheet at index ${sheetIndex} does not exist`});
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
}
