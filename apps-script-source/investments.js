function recordInvestment(investment) {
  try {
    // Debug logging
    Logger.log("=== Recording Investment ===");
    Logger.log("Investment data: " + JSON.stringify(investment));
    
    // Check if SHEETS is accessible
    if (typeof SHEETS === 'undefined') {
      throw new Error("SHEETS configuration not accessible");
    }
    
    // Check if INVESTMENTS sheet name is defined
    if (!SHEETS.INVESTMENTS) {
      throw new Error("INVESTMENTS sheet name not defined in configuration");
    }
    
    Logger.log("Looking for sheet: " + SHEETS.INVESTMENTS);
    
    // Get spreadsheet and sheet
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      throw new Error("Unable to access spreadsheet");
    }
    
    Logger.log("Spreadsheet name: " + spreadsheet.getName());
    
    const sheet = spreadsheet.getSheetByName(SHEETS.INVESTMENTS);
    if (!sheet) {
      // Try to create the sheet if it doesn't exist
      Logger.log("Investments sheet not found, attempting to create it");
      try {
        const newSheet = spreadsheet.insertSheet(SHEETS.INVESTMENTS);
        // Add headers
        newSheet.appendRow([
          'Timestamp', 'Detail', 'Qty', 'Unit Price', 'Total', 
          'Outlet ID', 'Outlet Name', 'Payment Source', 'Account', 
          'Session Outlet', 'Username'
        ]);
        Logger.log("Created new Investments sheet with headers");
      } catch (createError) {
        throw new Error("Investments sheet not found and could not be created: " + createError.toString());
      }
    }
    
    // Parse investment data if it's a string
    if (typeof investment === 'string') {
      try {
        investment = JSON.parse(investment);
      } catch (parseError) {
        throw new Error("Invalid investment data format: " + parseError.toString());
      }
    }

    // Validate session parameters
    const { deviceId, username } = investment;
    if (!deviceId || !username) {
      throw new Error("Device ID and username are required to record an investment.");
    }

    // Get user session
    var session = getUserSessionEnhanced(deviceId, username) || {};
    var timestamp = new Date();
    Logger.log("Timestamp:", timestamp);

    var row = [
      formatTimestampForSheet(timestamp),
      investment.detail || "",
      Number(investment.qty) || 0,
      Number(investment.unitPrice) || 0,
      Number(investment.total) || 0,
      investment.outletId || "",
      investment.outletName || "",
      investment.paymentSource || "",
      investment.account || "",
      session.outlet || "",
      session.username || ""
    ];
    
    Logger.log("Appending row: " + JSON.stringify(row));
    sheet.appendRow(row);
    Logger.log("Investment recorded successfully");
    
    return { success: true, message: "Investment recorded successfully" };
    
  } catch (error) {
    Logger.log("Error recording investment: " + error.toString());
    throw new Error("Failed to record investment: " + error.toString());
  }
}

function getInvestmentReportData() {
  try {
    Logger.log("=== Getting Investment Report Data ===");
    
    // Check if SHEETS is accessible
    if (typeof SHEETS === 'undefined') {
      Logger.log("SHEETS configuration not accessible");
      return [];
    }
    
    // Get spreadsheet and sheet
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      Logger.log("Unable to access spreadsheet");
      return [];
    }
    
    const sheet = spreadsheet.getSheetByName(SHEETS.INVESTMENTS);
    if (!sheet) {
      Logger.log("Investments sheet not found");
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log("Retrieved " + data.length + " rows from Investments sheet");
    
    if (data.length < 2) {
      Logger.log("No data rows found (only headers or empty sheet)");
      return [];
    }
    
    const headers = data.shift();
    Logger.log("Headers: " + JSON.stringify(headers));
    
    const outletIdIdx = headers.findIndex(h => /outlet id/i.test(h));
    const outletNameIdx = headers.findIndex(h => /outlet name/i.test(h));
    const accountIdx = headers.findIndex(h => /account/i.test(h));
    const detailIdx = headers.findIndex(h => /^detail$/i.test(h));
    const dateIdx = headers.findIndex(h => /timestamp|date/i.test(h));
    const amountIdx = headers.findIndex(h => /total/i.test(h));
    
    Logger.log("Column indices - Outlet: " + outletNameIdx + ", Account: " + accountIdx + ", Detail: " + detailIdx + ", Date: " + dateIdx + ", Amount: " + amountIdx);

    const reportRows = {};
    data.forEach(row => {
      const outlet = row[outletNameIdx] || 'Unknown';
      const account = row[accountIdx] || 'Unknown';
      const detail = detailIdx >= 0 ? (row[detailIdx] || '') : '';
      const rawDate = row[dateIdx] || '';
      const total = Number(row[amountIdx]) || 0;
      
      let dt;
      let month = 'Unknown';
      
      try {
        if (rawDate instanceof Date) {
          dt = rawDate;
        } else if (typeof rawDate === 'string' && rawDate.trim() !== '') {
          // Try Utilities.parseDate first for Google Sheets compatibility
          try {
            dt = Utilities.parseDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!dt || isNaN(dt.getTime())) {
              dt = Utilities.parseDate(rawDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            }
            if (!dt || isNaN(dt.getTime())) {
              dt = Utilities.parseDate(rawDate, Session.getScriptTimeZone(), 'dd-MM-yyyy');
            }
          } catch (e) {
            Logger.log('Utilities.parseDate failed, trying standard Date: ' + e);
          }
          
          // Fallback to standard Date parsing if Utilities.parseDate failed
          if (!dt || isNaN(dt.getTime())) {
            dt = new Date(rawDate);
          }
        } else if (rawDate) {
          dt = new Date(rawDate);
        }
        
        if (dt && !isNaN(dt.getTime())) {
          month = dt.getFullYear() + '-' + ('0'+(dt.getMonth()+1)).slice(-2);
        } else {
          Logger.log('Failed to parse date in investments: ' + JSON.stringify(rawDate));
        }
      } catch (e) {
        Logger.log('Error parsing date in investments: ' + e + ', dateRaw: ' + JSON.stringify(rawDate));
      }

      // Combine by Outlet, Account, Detail, Month
      const key = outlet + "|" + account + "|" + detail + "|" + month;
      if (!reportRows[key]) {
        reportRows[key] = {
          outlet,
          account,
          detail,
          month,
          total: 0
        };
      }
      reportRows[key].total += total;
    });
    
    const result = Object.values(reportRows);
    Logger.log("Returning " + result.length + " report rows");
    return result;
    
  } catch (error) {
    Logger.log("Error getting investment report data: " + error.toString());
    return [];
  }
}

function parseInvestmentDate(dateString) {
  let date = null;
  
  if (dateString instanceof Date) {
    return dateString;
  }
  
  if (typeof dateString === 'string') {
    // Try Utilities.parseDate first for Google Sheets compatibility
    try {
      // Try dd-MM-yyyy format first (our standard format)
      date = Utilities.parseDate(dateString, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
      if (!date || isNaN(date.getTime())) {
        date = Utilities.parseDate(dateString, Session.getScriptTimeZone(), 'dd-MM-yyyy');
      }
      if (!date || isNaN(date.getTime())) {
        date = Utilities.parseDate(dateString, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      if (!date || isNaN(date.getTime())) {
        date = Utilities.parseDate(dateString, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      }
    } catch (e) {
      Logger.log('Utilities.parseDate failed, trying standard Date: ' + e);
    }
    
    // Fallback to standard Date parsing if Utilities.parseDate failed
    if (!date || isNaN(date.getTime())) {
      date = new Date(dateString);
    }
  } else {
    date = new Date(dateString);
  }
  
  if (!date || isNaN(date.getTime())) {
    Logger.log('Failed to parse date: ' + dateString + ', using current time');
    date = new Date();
  }
  
  return date;
}

/**
 * SecurityWrapper-compatible investment recording function
 * @param {Object} params - Parameters object from SecurityWrapper (contains _authContext and investment data)
 * @returns {Object} Success response or throws error
 */
function recordInvestmentSecure(params) {
  try {
    Logger.log("recordInvestmentSecure called with params: " + JSON.stringify(params));
    
    // Extract the investment data and auth context from SecurityWrapper enhanced parameters
    var investment = params;
    var authContext = params._authContext;
    
    Logger.log("Auth context: " + JSON.stringify(authContext));
    Logger.log("Investment data: " + JSON.stringify(investment));
    
    // Validate authentication context
    if (!authContext || !authContext.username || !authContext.outlet) {
      throw new Error("Authentication required: Missing username or outlet information");
    }
    
    // Create investment record with authentication context
    const investmentWithAuth = {
      ...investment,
      deviceId: authContext.deviceId,
      username: authContext.username
    };
    
    // Call the original recordInvestment function
    const result = recordInvestment(investmentWithAuth);
    
    Logger.log("recordInvestmentSecure completed successfully for user: " + authContext.username);
    return { 
      success: true, 
      message: "Investment recorded successfully",
      username: authContext.username,
      outlet: authContext.outlet,
      ...result
    };
    
  } catch (error) {
    Logger.log("recordInvestmentSecure error: " + error.toString());
    throw error;
  }
}
/**
 * Investment transactions report source. Reads investment purchases from the
 * Purchases sheet (rows flagged Is Investment = Y) so the Investment tab can show
 * and group them. Returns one object per line with date (yyyy-MM-dd), supplier,
 * product, category (product group), qty, price, total, account and outlet.
 */
function getInvestmentTransactions() {
  try {
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = data[0];
    var invIdx = (typeof purchaseInvestmentColIndex_ === 'function') ? purchaseInvestmentColIndex_(headers) : -1;
    if (invIdx < 0) return [];

    function col(re) { return headers.findIndex(function (h) { return re.test(String(h)); }); }
    var dateIdx = col(/timestamp|^date/i);
    var supplierIdx = col(/supplier/i);
    var productIdx = col(/product.*name/i);
    var categoryIdx = col(/category/i);
    var qtyIdx = col(/quantity|qty/i);
    var priceIdx = col(/^price$/i);
    var totalIdx = col(/^total$|amount/i);
    var accountIdx = col(/^account$/i);
    var outletIdx = col(/outlet/i);

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!isInvestmentPurchaseRow_(row, invIdx)) continue;
      out.push({
        date: dateIdx >= 0 ? formatDateForJSON(row[dateIdx]) : '',
        supplier: supplierIdx >= 0 ? (row[supplierIdx] || '') : '',
        product: productIdx >= 0 ? (row[productIdx] || '') : '',
        category: categoryIdx >= 0 ? (row[categoryIdx] || '') : '',
        qty: qtyIdx >= 0 ? (Number(row[qtyIdx]) || 0) : 0,
        price: priceIdx >= 0 ? (Number(row[priceIdx]) || 0) : 0,
        total: totalIdx >= 0 ? (Number(row[totalIdx]) || 0) : 0,
        account: accountIdx >= 0 ? (row[accountIdx] || '') : '',
        outlet: outletIdx >= 0 ? (row[outletIdx] || '') : ''
      });
    }
    return out;
  } catch (error) {
    Logger.log('getInvestmentTransactions error: ' + error);
    return [];
  }
}
