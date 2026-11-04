/**
 * Export Sheet Structure Script
 * This script exports the first 3 rows (headers + 2 sample rows) from each sheet
 * as CSV content that can be saved to files for database structure analysis.
 * 
 * To use this script:
 * 1. Copy this function to your Google Apps Script project
 * 2. Run the exportAllSheetStructures() function
 * 3. Check the execution log to see the CSV content
 * 4. Copy each CSV block and save as separate .csv files in your workspace
 */

function exportAllSheetStructures() {
  try {
    // Get the spreadsheet
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      Logger.log("ERROR: Could not access spreadsheet");
      return;
    }
    
    Logger.log("=== EXPORTING SHEET STRUCTURES ===");
    Logger.log("Spreadsheet: " + spreadsheet.getName());
    Logger.log("Spreadsheet ID: " + spreadsheet.getId());
    
    // List of sheets to export (from config.js)
    const sheetsToExport = [
      { key: 'PRODUCTS', name: SHEETS.PRODUCTS },
      { key: 'ORDERS', name: SHEETS.ORDERS },
      { key: 'DEPOSITS', name: SHEETS.DEPOSITS },
      { key: 'EXPENSES', name: SHEETS.EXPENSES },
      { key: 'EXPENSE_ACCOUNTS', name: SHEETS.EXPENSE_ACCOUNTS },
      { key: 'PAYMENT_SOURCES', name: SHEETS.PAYMENT_SOURCES },
      { key: 'DEPOSIT_ACCOUNTS', name: SHEETS.DEPOSIT_ACCOUNTS },
      { key: 'PURCHASES', name: SHEETS.PURCHASES },
      { key: 'SUPPLIERS', name: SHEETS.SUPPLIERS },
      { key: 'INVESTMENTS', name: SHEETS.INVESTMENTS },
      { key: 'ACCOUNTS', name: SHEETS.ACCOUNTS },
      { key: 'PEOPLE', name: SHEETS.PEOPLE },
      { key: 'USERS', name: SHEETS.USERS },
      { key: 'INVENTORY_CONSUMPTION', name: SHEETS.INVENTORY_CONSUMPTION }
    ];
    
    // Also check for additional sheets that might exist
    const allSheets = spreadsheet.getSheets();
    const additionalSheets = allSheets.filter(sheet => {
      const sheetName = sheet.getName();
      return !sheetsToExport.some(s => s.name === sheetName) && 
             !sheetName.startsWith('_') && // Skip system sheets
             sheetName !== 'Outlets'; // Will handle separately
    });
    
    // Add Outlets sheet if it exists
    const outletsSheet = spreadsheet.getSheetByName('Outlets');
    if (outletsSheet) {
      sheetsToExport.push({ key: 'OUTLETS', name: 'Outlets' });
    }
    
    // Add any other discovered sheets
    additionalSheets.forEach(sheet => {
      sheetsToExport.push({ 
        key: sheet.getName().toUpperCase().replace(/[^A-Z0-9]/g, '_'), 
        name: sheet.getName() 
      });
    });
    
    Logger.log("Found " + sheetsToExport.length + " sheets to export");
    
    // Export each sheet
    sheetsToExport.forEach(sheetInfo => {
      exportSheetStructure(spreadsheet, sheetInfo.key, sheetInfo.name);
    });
    
    Logger.log("=== EXPORT COMPLETE ===");
    Logger.log("Copy the CSV blocks above and save them as separate .csv files in your workspace");
    Logger.log("Suggested folder: database_structure/");
    
  } catch (error) {
    Logger.log("ERROR in exportAllSheetStructures: " + error.toString());
  }
}

function exportSheetStructure(spreadsheet, sheetKey, sheetName) {
  try {
    Logger.log("\n" + "=".repeat(50));
    Logger.log("SHEET: " + sheetName + " (" + sheetKey + ")");
    Logger.log("=".repeat(50));
    
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("Sheet '" + sheetName + "' not found - SKIPPING");
      return;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      Logger.log("Sheet '" + sheetName + "' is empty - SKIPPING");
      return;
    }
    
    Logger.log("Sheet dimensions: " + lastRow + " rows x " + lastCol + " columns");
    
    // Get first 3 rows (headers + 2 sample rows) or all rows if less than 3
    const rowsToExport = Math.min(3, lastRow);
    const range = sheet.getRange(1, 1, rowsToExport, lastCol);
    const data = range.getValues();
    
    Logger.log("Exporting " + rowsToExport + " rows");
    
    // Convert to CSV format
    Logger.log("\n--- CSV CONTENT START (save as: " + sheetName.toLowerCase() + "_structure.csv) ---");
    
    data.forEach((row, index) => {
      // Clean and escape CSV values
      const csvRow = row.map(cell => {
        if (cell === null || cell === undefined) return '';
        let value = cell.toString();
        // Escape quotes and wrap in quotes if contains comma, quote, or newline
        if (value.includes(',') || value.includes('"') || value.includes('\n')) {
          value = '"' + value.replace(/"/g, '""') + '"';
        }
        return value;
      }).join(',');
      
      Logger.log(csvRow);
    });
    
    Logger.log("--- CSV CONTENT END ---");
    
    // Also log column analysis
    if (data.length > 0) {
      Logger.log("\n--- COLUMN ANALYSIS ---");
      const headers = data[0];
      headers.forEach((header, index) => {
        const columnLetter = getColumnLetter(index + 1);
        Logger.log("Column " + columnLetter + " (" + (index + 1) + "): " + header);
      });
      
      if (data.length > 1) {
        Logger.log("\n--- SAMPLE DATA TYPES ---");
        const sampleRow = data[1];
        headers.forEach((header, index) => {
          const value = sampleRow[index];
          const type = typeof value;
          const actualType = value instanceof Date ? 'Date' : 
                           typeof value === 'number' ? 'Number' :
                           typeof value === 'boolean' ? 'Boolean' :
                           'String';
          Logger.log(header + ": " + actualType + " (example: " + value + ")");
        });
      }
    }
    
  } catch (error) {
    Logger.log("ERROR exporting sheet '" + sheetName + "': " + error.toString());
  }
}

function getColumnLetter(columnNumber) {
  let columnLetter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}

/**
 * Alternative function to export just the headers of all sheets
 * Use this if you only need column names without sample data
 */
function exportSheetHeaders() {
  try {
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      Logger.log("ERROR: Could not access spreadsheet");
      return;
    }
    
    Logger.log("=== SHEET HEADERS EXPORT ===");
    
    const allSheets = spreadsheet.getSheets();
    
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith('_')) return; // Skip system sheets
      
      Logger.log("\n" + sheetName + ":");
      Logger.log("-".repeat(sheetName.length + 1));
      
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) {
        Logger.log("(Empty sheet)");
        return;
      }
      
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      headers.forEach((header, index) => {
        Logger.log((index + 1) + ". " + header);
      });
    });
    
  } catch (error) {
    Logger.log("ERROR in exportSheetHeaders: " + error.toString());
  }
}

/**
 * Export a specific sheet by name
 * Usage: exportSpecificSheet('Products')
 */
function exportSpecificSheet(sheetName) {
  try {
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      Logger.log("ERROR: Could not access spreadsheet");
      return;
    }
    
    exportSheetStructure(spreadsheet, sheetName.toUpperCase(), sheetName);
    
  } catch (error) {
    Logger.log("ERROR in exportSpecificSheet: " + error.toString());
  }
}
