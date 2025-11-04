import * as XLSX from 'xlsx';
import type { IWorkbookData } from '@univerjs/core';
import { LocaleType } from '@univerjs/core';

/**
 * Converts Univer workbook data to XLSX format and downloads it
 */
export async function exportWorkbookToXLSX(workbookData: IWorkbookData, filename: string = 'workbook.xlsx'): Promise<void> {
  try {
    // Create a new XLSX workbook
    const wb = XLSX.utils.book_new();

    // Get all sheets from Univer workbook
    const sheetOrder = workbookData.sheetOrder || [];
    const sheets = workbookData.sheets || {};

    // Convert each sheet
    for (const sheetId of sheetOrder) {
      const sheet = sheets[sheetId];
      if (!sheet) continue;

      // Get sheet name
      const sheetName = sheet.name || sheetId;

      // Get cell matrix data
      const cellMatrix = sheet.cellData || {};
      
      // Find the maximum row and column
      let maxRow = 0;
      let maxCol = 0;
      
      for (const rowKey in cellMatrix) {
        const row = parseInt(rowKey);
        if (row > maxRow) maxRow = row;
        
        const rowData = cellMatrix[rowKey];
        if (rowData) {
          for (const colKey in rowData) {
            const col = parseInt(colKey);
            if (col > maxCol) maxCol = col;
          }
        }
      }

      // Create a 2D array for the sheet data
      const sheetData: any[][] = [];
      
      // Initialize all rows
      for (let r = 0; r <= maxRow; r++) {
        sheetData[r] = [];
        for (let c = 0; c <= maxCol; c++) {
          sheetData[r][c] = '';
        }
      }

      // Fill in the data from Univer's cell matrix
      for (const rowKey in cellMatrix) {
        const row = parseInt(rowKey);
        const rowData = cellMatrix[rowKey];
        
        if (rowData) {
          for (const colKey in rowData) {
            const col = parseInt(colKey);
            const cell = rowData[colKey];
            
            if (cell) {
              // Get the value from the cell
              // Cell structure: { v: value, t: type, s: style, ... }
              let value: any = '';
              
              if (cell.v !== undefined && cell.v !== null) {
                // Get the raw value - XLSX cells can have various types
                // Use type assertion to work around TypeScript's strict type checking
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const rawValue: any = cell.v;
                
                // Check cell type first, then handle value accordingly
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                if ((cell as any).t === 'n') {
                  // Number type
                  value = Number(rawValue);
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                } else if ((cell as any).t === 'b') {
                  // Boolean type
                  value = Boolean(rawValue);
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                } else if ((cell as any).t === 's') {
                  // String type
                  const strValue = String(rawValue);
                  // Check if it's a formula (starts with =)
                  if (strValue.startsWith('=')) {
                    value = strValue;
                  } else {
                    value = strValue;
                  }
                } else {
                  // Default: convert to string
                  value = String(rawValue);
                }
              }
              
              sheetData[row][col] = value;
            }
          }
        }
      }

      // Create XLSX worksheet from the 2D array
      const ws = XLSX.utils.aoa_to_sheet(sheetData);

      // Add column widths if available
      const columnProperties = sheet.columnData;
      if (columnProperties) {
        const cols: { wch: number }[] = [];
        for (let c = 0; c <= maxCol; c++) {
          const colProp = columnProperties[c];
          if (colProp && colProp.w !== undefined) {
            // Convert pixels to character width (approximate)
            cols.push({ wch: Math.max(colProp.w / 7, 10) });
          } else {
            cols.push({ wch: 10 });
          }
        }
        ws['!cols'] = cols;
      }

      // Add row heights if available
      const rowProperties = sheet.rowData;
      if (rowProperties) {
        const rows: { hpt: number }[] = [];
        for (let r = 0; r <= maxRow; r++) {
          const rowProp = rowProperties[r];
          if (rowProp && rowProp.h !== undefined) {
            rows.push({ hpt: rowProp.h });
          }
        }
        if (rows.length > 0) {
          ws['!rows'] = rows;
        }
      }

      // Add merge ranges if available
      const mergeData = sheet.mergeData;
      if (mergeData && mergeData.length > 0) {
        const merges: XLSX.Range[] = [];
        for (const merge of mergeData) {
          merges.push({
            s: { r: merge.startRow || 0, c: merge.startColumn || 0 },
            e: { r: merge.endRow || 0, c: merge.endColumn || 0 },
          });
        }
        ws['!merges'] = merges;
      }

      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    }

    // Write and download the file
    XLSX.writeFile(wb, filename);
    console.log(`✅ Exported workbook to ${filename}`);
  } catch (error) {
    console.error('Error exporting workbook to XLSX:', error);
    throw error;
  }
}

/**
 * Converts XLSX file to Univer workbook data format
 */
export async function importXLSXToWorkbookData(file: File): Promise<IWorkbookData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) {
          reject(new Error('Failed to read file'));
          return;
        }

        // Read the XLSX workbook
        // XLSX.read() reads all cells by default, regardless of !ref range
        // The !ref property is just a convenience hint, but all cells are available in the ws object
        const wb = XLSX.read(data, { 
          type: 'binary', 
          cellFormula: true, 
          cellDates: true,
        });

        // Convert to Univer workbook format
        const workbookData: IWorkbookData = {
          id: `workbook-${Date.now()}`,
          name: file.name.replace(/\.(xlsx|xls)$/i, ''),
          appVersion: '0.10.8',
          sheets: {},
          sheetOrder: [],
          locale: LocaleType.EN_US,
          styles: {},
        };

        // Convert each sheet
        wb.SheetNames.forEach((sheetName) => {
          const ws = wb.Sheets[sheetName];
          if (!ws) return;

          // Create cell matrix for Univer
          const cellMatrix: { [row: string]: { [col: string]: any } } = {};
          
          // Track maximum row and column to determine sheet dimensions
          let maxRow = -1;
          let maxCol = -1;

          // Iterate through ALL cell addresses in the worksheet object
          // In sparse mode (dense: false), ALL cells are included as keys, even those outside !ref
          // Cell addresses are keys like "A1", "B2", etc.
          const cellAddressPattern = /^[A-Z]+[0-9]+$/;
          let cellCount = 0;
          
          for (const cellAddress in ws) {
            // Skip special properties that start with !
            if (cellAddress.startsWith('!')) {
              continue;
            }
            
            // Only process valid cell addresses
            if (!cellAddressPattern.test(cellAddress)) {
              continue;
            }

            const cell = ws[cellAddress];
            if (!cell) continue;

            cellCount++;

            // Decode the cell address to get row and column indices
            const decoded = XLSX.utils.decode_cell(cellAddress);
            const R = decoded.r;
            const C = decoded.c;
            
            // Update max row and column (this ensures we capture all cells with data)
            if (R > maxRow) maxRow = R;
            if (C > maxCol) maxCol = C;

            const rowKey = String(R);
            const colKey = String(C);

            // Initialize row if needed
            if (!cellMatrix[rowKey]) {
              cellMatrix[rowKey] = {};
            }

            // Create cell data object
            const cellData: any = {};
            
            // Set value
            if (cell.v !== undefined) {
              cellData.v = cell.v;
            } else if (cell.w !== undefined) {
              // Use formatted text if value is not available
              cellData.v = cell.w;
            }
            
            // Set type
            if (cell.t) {
              cellData.t = cell.t; // 'n' for number, 's' for string, 'b' for boolean, 'e' for error
            } else if (typeof cellData.v === 'number') {
              cellData.t = 'n';
            } else if (typeof cellData.v === 'boolean') {
              cellData.t = 'b';
            } else {
              cellData.t = 's';
            }
            
            // Set formula if present
            if (cell.f) {
              cellData.f = cell.f;
              // For formulas, keep the calculated value if available
              if (cell.v !== undefined) {
                cellData.v = cell.v;
              }
            }

            cellMatrix[rowKey][colKey] = cellData;
          }

          // Get the range from !ref for comparison/debugging (but don't rely on it)
          const rangeFromRef: XLSX.Range | null = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
          
          // If maxRow from actual cells is higher than !ref range, we found cells beyond !ref
          if (rangeFromRef && maxRow > rangeFromRef.e.r) {
            console.log(`⚠️ Found ${maxRow - rangeFromRef.e.r} additional rows beyond !ref range (${rangeFromRef.e.r + 1} to ${maxRow})`);
          }

          // Note: We don't use sheet_to_json anymore because it's limited by !ref
          // By iterating through all cell addresses in the ws object directly,
          // we capture ALL cells, even those beyond the !ref range

          // Debug logging to understand what we're importing
          console.log(`📊 Importing sheet "${sheetName}":`, {
            rangeFromRef: ws['!ref'] || 'none',
            maxRowFromCells: maxRow,
            maxColFromCells: maxCol,
            totalRowsInMatrix: Object.keys(cellMatrix).length,
            totalCellAddresses: cellCount,
            cellsPerRow: Object.keys(cellMatrix).map(rowKey => Object.keys(cellMatrix[rowKey] || {}).length).reduce((a, b) => a + b, 0) / Math.max(Object.keys(cellMatrix).length, 1),
          });

          // Create sheet data
          // Set rowCount and columnCount to at least 100,000 rows and 1,000 columns
          // This ensures users can scroll to these limits. Univer's virtual scrolling will handle rendering efficiently.
          const MIN_ROWS = 100000;
          const MIN_COLS = 1000;
          const calculatedRowCount = Math.max(maxRow + 1, MIN_ROWS);
          const calculatedColCount = Math.max(maxCol + 1, MIN_COLS);
          
          const sheetId = `sheet-${Date.now()}-${wb.SheetNames.indexOf(sheetName)}`;
          workbookData.sheets[sheetId] = {
            id: sheetId,
            name: sheetName,
            cellData: cellMatrix,
            rowCount: calculatedRowCount,
            columnCount: calculatedColCount,
            showGridlines: 1,
            rowData: {},
            columnData: {},
            mergeData: [],
          };

          // Add column widths
          if (ws['!cols']) {
            const columnData: { [col: string]: any } = {};
            ws['!cols'].forEach((col, index) => {
              if (col && col.wch) {
                // Convert character width to pixels (approximate)
                columnData[String(index)] = { w: col.wch * 7 };
              }
            });
            if (Object.keys(columnData).length > 0) {
              workbookData.sheets[sheetId].columnData = columnData;
            }
          }

          // Add row heights
          if (ws['!rows']) {
            const rowData: { [row: string]: any } = {};
            ws['!rows'].forEach((row, index) => {
              if (row && row.hpt) {
                rowData[String(index)] = { h: row.hpt };
              }
            });
            if (Object.keys(rowData).length > 0) {
              workbookData.sheets[sheetId].rowData = rowData;
            }
          }

          // Add merge ranges
          if (ws['!merges']) {
            const mergeData: any[] = [];
            ws['!merges'].forEach((merge: XLSX.Range) => {
              mergeData.push({
                startRow: merge.s.r,
                startColumn: merge.s.c,
                endRow: merge.e.r,
                endColumn: merge.e.c,
              });
            });
            if (mergeData.length > 0) {
              workbookData.sheets[sheetId].mergeData = mergeData;
            }
          }

          workbookData.sheetOrder.push(sheetId);
        });

        console.log(`✅ Imported XLSX file: ${file.name}`, {
          sheetCount: workbookData.sheetOrder.length,
          sheetNames: workbookData.sheetOrder.map(id => workbookData.sheets[id]?.name),
          sheets: workbookData.sheetOrder.map(id => ({
            name: workbookData.sheets[id]?.name,
            rowCount: workbookData.sheets[id]?.rowCount,
            columnCount: workbookData.sheets[id]?.columnCount,
          })),
        });

        resolve(workbookData);
      } catch (error) {
        console.error('Error importing XLSX file:', error);
        reject(error);
      }
    };

    reader.onerror = () => {
      reject(new Error('Failed to read file'));
    };

    // Read file as binary string for XLSX
    reader.readAsBinaryString(file);
  });
}
