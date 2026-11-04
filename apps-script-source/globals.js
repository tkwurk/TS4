// Global spreadsheet instance
var GLOBAL = {
  spreadsheet: null
};

function resolveSpreadsheetIdForOpen_() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) return SPREADSHEET_ID;
  var stored = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (stored) return stored;
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      var id = active.getId();
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
      return id;
    }
  } catch (e) {
    Logger.log('resolveSpreadsheetIdForOpen_ active spreadsheet: ' + e.message);
  }
  return '';
}

function initSpreadsheet() {
  if (!GLOBAL.spreadsheet) {
    var spreadsheetId = resolveSpreadsheetIdForOpen_();
    if (!spreadsheetId) {
      throw new Error(
        'SPREADSHEET_ID is not configured. Open this script from the bound spreadsheet ' +
        '(Extensions → Apps Script) and run cacheSpreadsheetIdForWebApp(), or set Script Property SPREADSHEET_ID.'
      );
    }
    try {
      GLOBAL.spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      Logger.log('Spreadsheet initialized successfully with ID: ' + spreadsheetId);
    } catch (error) {
      Logger.log('Error initializing spreadsheet: ' + error.toString());
      throw new Error('Failed to open spreadsheet ' + spreadsheetId + ': ' + error.message);
    }
  }
  return GLOBAL.spreadsheet;
}

/** Run once from the bound spreadsheet Apps Script editor to cache ID for web app /exec. */
function cacheSpreadsheetIdForWebApp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Open this project from the spreadsheet (Extensions → Apps Script), then run again.');
  }
  var id = ss.getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  GLOBAL.spreadsheet = null;
  return 'Cached SPREADSHEET_ID: ' + id;
}

function getGlobalSpreadsheet() {
  try {
    if (GLOBAL.spreadsheet) {
      return GLOBAL.spreadsheet;
    }
    return initSpreadsheet();
  } catch (error) {
    Logger.log("Error in getGlobalSpreadsheet: " + error.toString());
    throw error;
  }
}

// Debug function to test spreadsheet access
function testSpreadsheetAccess() {
  try {
    Logger.log("Testing spreadsheet access...");
    Logger.log("SPREADSHEET_ID defined: " + (typeof SPREADSHEET_ID !== 'undefined'));
    Logger.log("SPREADSHEET_ID value: " + SPREADSHEET_ID);
    
    const spreadsheet = getGlobalSpreadsheet();
    Logger.log("Spreadsheet object: " + spreadsheet);
    Logger.log("Spreadsheet name: " + spreadsheet.getName());
    
    const sheet = spreadsheet.getSheetByName("Orders");
    Logger.log("Orders sheet found: " + (sheet !== null));
    
    return "Spreadsheet access test completed successfully";
  } catch (error) {
    Logger.log("Spreadsheet access test failed: " + error.toString());
    return "Error: " + error.toString();
  }
}

// Runtime reset and diagnostic functions
function resetRuntime() {
  try {
    // Clear the global spreadsheet
    GLOBAL.spreadsheet = null;
    
    // Force re-initialization
    Logger.log("=== RUNTIME RESET ===");
    Logger.log("Global spreadsheet cleared");
    
    // Test basic variable access
    Logger.log("SPREADSHEET_ID accessible: " + (typeof SPREADSHEET_ID !== 'undefined'));
    Logger.log("SHEETS accessible: " + (typeof SHEETS !== 'undefined'));
    
    if (typeof SHEETS !== 'undefined') {
      Logger.log("SHEETS.INVESTMENTS: " + SHEETS.INVESTMENTS);
    }
    
    // Test spreadsheet access
    const spreadsheet = getGlobalSpreadsheet();
    Logger.log("Spreadsheet reinitialized: " + spreadsheet.getName());
    
    return "Runtime reset successful";
  } catch (error) {
    Logger.log("Runtime reset failed: " + error.toString());
    return "Error: " + error.toString();
  }
}

function diagnosticCheck() {
  try {
    Logger.log("=== DIAGNOSTIC CHECK ===");
    
    // Check all global variables
    var globals = [
      'SPREADSHEET_ID', 'SHEETS', 'GLOBAL',
      'SpreadsheetApp', 'Logger', 'Utilities'
    ];
    
    globals.forEach(function(varName) {
      try {
        var exists = (typeof eval(varName) !== 'undefined');
        Logger.log(varName + " defined: " + exists);
        if (varName === 'SHEETS' && exists) {
          Logger.log("SHEETS keys: " + Object.keys(SHEETS).join(', '));
        }
      } catch (e) {
        Logger.log(varName + " error: " + e.toString());
      }
    });
    
    return "Diagnostic check completed - see logs";
  } catch (error) {
    Logger.log("Diagnostic check failed: " + error.toString());
    return "Error: " + error.toString();
  }
}