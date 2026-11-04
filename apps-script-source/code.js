/**
 * Web app entry point (Snowbee A3.1).
 * Shared helpers live in utils.js — do not redeclare getSheetData, monthKey, etc.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function escapeHtmlForOutput_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function doGet(e) {
  if (!initializeSpreadsheet()) {
    var detail = getLastSpreadsheetInitError_();
    var body = 'Error: Could not initialize spreadsheet';
    if (detail) body += '\n\n' + detail;
    return HtmlService.createHtmlOutput('<pre style="font:14px/1.5 monospace;padding:16px">' +
      escapeHtmlForOutput_(body) + '</pre>');
  }
  // Production entry: ui_index only. Legacy index.html is deprecated (Snowbee A1.6).
  const html = HtmlService.createTemplateFromFile('ui_index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function getPaymentSources() {
  try {
    Logger.log("=== Getting Payment Sources ===");

    if (typeof SHEETS === 'undefined') {
      Logger.log("SHEETS configuration not accessible");
      return [];
    }

    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      Logger.log("Unable to access spreadsheet");
      return [];
    }

    const sheet = spreadsheet.getSheetByName(SHEETS.PAYMENT_SOURCES);
    if (!sheet) {
      Logger.log("Payment Sources sheet not found");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    Logger.log("Retrieved " + data.length + " rows from Payment Sources sheet");

    if (data.length < 2) {
      Logger.log("No data rows found (only headers or empty sheet)");
      return [];
    }

    data.shift();
    const nameCol = COLUMNS.PAYMENT_SOURCES.NAME.col;
    const sources = data.map(function(row) { return row[nameCol]; }).filter(Boolean);
    Logger.log("Returning " + sources.length + " payment sources: " + JSON.stringify(sources));

    return sources;

  } catch (error) {
    Logger.log("Error getting payment sources: " + error.toString());
    return [];
  }
}

function getTutorials() {
  try {
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      throw new Error("Could not access spreadsheet");
    }

    const sheet = spreadsheet.getSheetByName(SHEETS.TUTORIALS);
    if (!sheet) {
      throw new Error("Tutorials sheet not found");
    }

    const range = sheet.getDataRange();
    const values = range.getValues();

    if (values.length < 2) {
      return [];
    }

    const headers = values[0];
    const tutorials = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const tutorial = {
        courseId: getFixedRowField(row, 'TUTORIALS', 'COURSE_ID') || '',
        description: getFixedRowField(row, 'TUTORIALS', 'DESCRIPTION') || '',
        category: getFixedRowField(row, 'TUTORIALS', 'CATEGORY') || '',
        supplier: getFixedRowField(row, 'TUTORIALS', 'SUPPLIER') || '',
        status: getFixedRowField(row, 'TUTORIALS', 'STATUS') || '',
        url: getFixedRowField(row, 'TUTORIALS', 'URL') || ''
      };

      if (tutorial.courseId || tutorial.description) {
        tutorials.push(tutorial);
      }
    }

    Logger.log(`Retrieved ${tutorials.length} tutorials from Tutorials sheet`);
    return tutorials;

  } catch (error) {
    Logger.log('Error in getTutorials: ' + error.toString());
    throw error;
  }
}
