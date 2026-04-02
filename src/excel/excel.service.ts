import { Injectable } from '@nestjs/common';
import * as XLSX from 'xlsx';

@Injectable()
export class ExcelService {
  private brandColumns = {
    'Brand A': ['Model', 'Price', 'Subsidy A', 'Stock', 'Color'],
    'Brand B': ['Model', 'Price', 'Subsidy B', 'Subsidy B2', 'Stock'],
    'Brand C': ['Model', 'Price', 'Subsidy C', 'Warranty', 'Stock'],
  };

  private results = new Map<string, any>();
  getBrandColumns() {
    return this.brandColumns;
  }

  storeResult(id: string, data: any) {
    this.results.set(id, data);
    setTimeout(() => this.results.delete(id), 10 * 60 * 1000);
  }

  getResult(id: string) {
    return this.results.get(id);
  }

  async filterBySubIds(
    fileBuffer: Buffer, 
    subIds: string[], 
    filterColumnName?: string
  ) {
    // Read with raw: false to get formatted values in cell.w
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
    
    // Read headers
    const headers = [];
    const headerRow = range.s.r;
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
      const cell = sourceWorksheet[cellAddress];
      headers.push(cell ? String(cell.v) : '');
    }
    
    // Find the filter column index
    let filterColIndex = -1;
    if (filterColumnName) {
      filterColIndex = headers.indexOf(filterColumnName);
    } else {
      // Try common variations
      const commonNames = ['SubID', 'subid', 'subId', 'SUBID', 'ExtraParam', 'extraParam', 'extParam2', 'extparam2'];
      for (const name of commonNames) {
        filterColIndex = headers.indexOf(name);
        if (filterColIndex !== -1) break;
      }
    }
    
    if (filterColIndex === -1) {
      throw new Error('Filter column not found in the Excel file');
    }
    
    // Group rows by SubID value
    const rowsBySubId = new Map<string, number[]>();
    
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
    
    // Filter based on provided SubIDs or auto-detect
    let subIdsToProcess: string[] = [];
    
    if (!subIds || subIds.length === 0) {
      // Auto-detect SubIDs with more than 10 records
      subIdsToProcess = Array.from(rowsBySubId.entries())
        .filter(([_, rows]) => rows.length > 10)
        .map(([subId, _]) => subId);
    } else {
      subIdsToProcess = subIds;
    }
    
    // Create sheets for each SubID
    for (const subId of subIdsToProcess) {
      const rows = rowsBySubId.get(subId);
      if (rows && rows.length > 0) {
        const sheetLabel = `${filterColumnName || 'extParam2'}_${subId}`;
        
        // Create preview data using formatted text
        const previewData = [];
        for (const sourceRow of rows) {
          const rowData = {};
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: sourceRow, c: col });
            const cell = sourceWorksheet[cellAddress];
            const header = headers[col - range.s.c];
            
            // Always use cell.w (formatted text) which has the display value
            if (cell && cell.w !== undefined && cell.w !== null) {
              rowData[header] = cell.w;
            } else if (cell && cell.v !== undefined && cell.v !== null) {
              rowData[header] = String(cell.v);
            } else {
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
    
    // If no data was found
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

  async createExcelFromData(sheetsData: any): Promise<Buffer> {
    const newWorkbook = XLSX.utils.book_new();
    const sourceWorksheet = sheetsData.sourceWorksheet;
    const range = sheetsData.range;
    
    for (const sheet of sheetsData.sheets) {
      if (sheet.rows && sheet.rows.length > 0) {
        // Create new worksheet
        const newWorksheet = {};
        
        // Copy worksheet properties with deep clone
        if (sourceWorksheet['!cols']) {
          newWorksheet['!cols'] = JSON.parse(JSON.stringify(sourceWorksheet['!cols']));
        }
        if (sourceWorksheet['!rows']) {
          newWorksheet['!rows'] = JSON.parse(JSON.stringify(sourceWorksheet['!rows']));
        }
        if (sourceWorksheet['!merges']) {
          newWorksheet['!merges'] = JSON.parse(JSON.stringify(sourceWorksheet['!merges']));
        }
        
        // Helper function to deep copy a cell
        const copyCellWithFormat = (sourceCell) => {
          if (!sourceCell) return null;
          
          const newCell: any = {
            v: sourceCell.v,
            t: sourceCell.t
          };
          
          // Copy all formatting properties
          if (sourceCell.w !== undefined) newCell.w = sourceCell.w;
          if (sourceCell.z !== undefined) newCell.z = sourceCell.z;
          if (sourceCell.f !== undefined) newCell.f = sourceCell.f;
          if (sourceCell.F !== undefined) newCell.F = sourceCell.F;
          if (sourceCell.r !== undefined) newCell.r = sourceCell.r;
          if (sourceCell.h !== undefined) newCell.h = sourceCell.h;
          if (sourceCell.c !== undefined) newCell.c = JSON.parse(JSON.stringify(sourceCell.c));
          if (sourceCell.s !== undefined) newCell.s = JSON.parse(JSON.stringify(sourceCell.s));
          if (sourceCell.l !== undefined) newCell.l = JSON.parse(JSON.stringify(sourceCell.l));
          
          return newCell;
        };
        
        // Copy header row
        for (let col = range.s.c; col <= range.e.c; col++) {
          const sourceCellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
          const targetCellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
          const sourceCell = sourceWorksheet[sourceCellAddress];
          
          if (sourceCell) {
            newWorksheet[targetCellAddress] = copyCellWithFormat(sourceCell);
          }
        }
        
        // Copy data rows
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
        
        // Set the range
        newWorksheet['!ref'] = XLSX.utils.encode_range({
          s: { r: 0, c: range.s.c },
          e: { r: targetRow - 1, c: range.e.c }
        });
        
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheet.name);
      } else {
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
}
