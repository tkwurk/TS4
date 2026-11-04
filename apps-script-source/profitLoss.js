function getProfitLossReport(deviceId, username, year) {
  year = Number(year) || new Date().getFullYear();
  return getPnLBundle(deviceId, username, year).daily;
}

// Helper function to extract day key from date (YYYY-MM-DD format)
function getDayKey(date) {
  if (!date) {
    Logger.log('getDayKey: null/undefined date provided');
    return null;
  }
  
  // Log the original date for debugging June 30th issues
  var originalDate = date;
  if ((typeof date === 'string' && date.includes('06-30')) || 
      (typeof date === 'string' && date.includes('30-06')) ||
      (typeof date === 'string' && date.includes('June')) ||
      (typeof date === 'string' && date.includes('30'))) {
    Logger.log('🔍 getDayKey processing potential June 30th date: ' + JSON.stringify(originalDate));
  }
  
  var d;
  if (date instanceof Date) {
    d = date;
    Logger.log('getDayKey: Date object provided: ' + d.toISOString());
  } else if (typeof date === 'string') {
    Logger.log('getDayKey: String date provided: "' + date + '"');
    
    // Enhanced date parsing for multiple formats
    try {
      // First try the most common Google Sheets formats
      d = Utilities.parseDate(date, Session.getScriptTimeZone(), 'MM-dd-yyyy HH:mm:ss');
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(date, Session.getScriptTimeZone(), 'MM-dd-yyyy');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(date, Session.getScriptTimeZone(), 'dd-MM-yyyy');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
    } catch (e) {
      Logger.log('getDayKey: Utilities.parseDate failed, trying standard Date: ' + e);
    }
    
    // Fallback to standard Date parsing if Utilities.parseDate failed
    if (!d || isNaN(d.getTime())) {
      Logger.log('getDayKey: Trying standard Date constructor');
      d = new Date(date);
    }
  } else {
    Logger.log('getDayKey: Non-string, non-Date object provided');
    d = new Date(date);
  }
  
  if (!d || isNaN(d.getTime())) {
    Logger.log('getDayKey: Failed to parse date: ' + JSON.stringify(originalDate));
    return null;
  }
  
  var year = d.getFullYear();
  var month = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  var result = year + '-' + month + '-' + day;
  
  // Enhanced logging for June 30, 2025 tracking
  if (result === '2025-06-30') {
    Logger.log('🎯✅ SUCCESS: June 30, 2025 successfully parsed in getDayKey!');
    Logger.log('   Original input: ' + JSON.stringify(originalDate));
    Logger.log('   Parsed Date object: ' + d.toISOString());
    Logger.log('   Final result: ' + result);
  }
  
  // Also log if we're getting June 2025 dates to track the pattern
  if (result && result.includes('2025-06')) {
    Logger.log('📅 June 2025 date processed: ' + result + ' (from: ' + JSON.stringify(originalDate) + ')');
  }
  
  return result;
}

// Function to get weekly breakdown for a specific month
function getWeeklyBreakdown(deviceId, username, year, month) {
  Logger.log("=== Generating Weekly Breakdown ===");
  Logger.log("DeviceId: " + (deviceId || 'NOT PROVIDED'));
  Logger.log("Username: " + (username || 'NOT PROVIDED'));
  Logger.log("Year: " + year + ", Month: " + month);
  
  // Get user session for outlet filtering using clean login approach
  var session = {};
  
  if (deviceId && username) {
    // Use clean login session retrieval
    session = getUserSessionEnhanced(deviceId, username) || {};
    Logger.log("Using clean login session retrieval");
  } else {
    // Fallback to enhanced session for backward compatibility
    session = getUserSessionEnhanced(deviceId) || {};
    Logger.log("Using enhanced session retrieval (backward compatibility)");
  }
  
  var userOutlet = session.outlet;
  Logger.log("User outlet: " + (userOutlet || 'NONE SET'));
  
  // Calculate the first and last day of the month
  var startDate = new Date(year, month - 1, 1); // month is 1-based, but Date() is 0-based
  var endDate = new Date(year, month, 0); // Last day of the month
  var daysInMonth = endDate.getDate();
  
  Logger.log("Month range: " + startDate.toISOString() + " to " + endDate.toISOString());
  Logger.log("Days in month: " + daysInMonth);
  
  var salesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
  var expenseSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.EXPENSES);
  var purchaseSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
  
  // Initialize weekly data structure
  var weeks = [];
  var currentDate = new Date(startDate);
  var weekNumber = 1;
  
  // Generate weeks for the month
  while (currentDate <= endDate) {
    var weekStart = new Date(currentDate);
    var weekEnd = new Date(currentDate);
    weekEnd.setDate(weekEnd.getDate() + 6); // Add 6 days for a 7-day week
    
    // Don't let week end go beyond month end
    if (weekEnd > endDate) {
      weekEnd = new Date(endDate);
    }
    
    weeks.push({
      weekNumber: weekNumber,
      startDate: new Date(weekStart),
      endDate: new Date(weekEnd),
      sales: 0,
      purchases: 0,
      expenses: 0,
      profit: 0
    });
    
    // Move to next week
    currentDate.setDate(currentDate.getDate() + 7);
    weekNumber++;
  }
  
  Logger.log("Generated " + weeks.length + " weeks for the month");
  
  // Process sales data
  if (salesSheet) {
    var salesData = salesSheet.getDataRange().getValues();
    if (salesData.length > 1) {
      var salesHeaders = salesData.shift();
      var salesCols = getTransactionColumns(salesHeaders, 'ORDERS');
      var dateIdx = salesCols.date;
      var totalIdx = salesCols.total;
      var outletIdx = salesCols.outlet;
      
      salesData.forEach(function(row) {
        // Filter by outlet only if user has one set
        if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;
        if (!row || !row[dateIdx]) return;
        
        var transactionDate = parseTransactionDate(row[dateIdx]);
        if (!transactionDate) return;
        
        // Check if transaction is within our month
        if (transactionDate >= startDate && transactionDate <= endDate) {
          var amount = Number(row[totalIdx]) || 0;
          
          // Find which week this transaction belongs to
          for (var i = 0; i < weeks.length; i++) {
            if (transactionDate >= weeks[i].startDate && transactionDate <= weeks[i].endDate) {
              weeks[i].sales += amount;
              break;
            }
          }
        }
      });
    }
  }
  
  // Process purchases data
  if (purchaseSheet) {
    var purchaseData = purchaseSheet.getDataRange().getValues();
    if (purchaseData.length > 1) {
      var purchaseHeaders = purchaseData.shift();
      var purchaseCols = getTransactionColumns(purchaseHeaders, 'PURCHASES');
      var dateIdx = purchaseCols.date;
      var totalIdx = purchaseCols.total;
      var outletIdx = purchaseCols.outlet;
      var purchaseInvIdx = purchaseInvestmentColIndex_(purchaseHeaders);
      
      purchaseData.forEach(function(row) {
        // Filter by outlet only if user has one set
        if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;
        if (isInvestmentPurchaseRow_(row, purchaseInvIdx)) return; // capital spend, not P&L cost
        if (!row || !row[dateIdx]) return;
        
        var transactionDate = parseTransactionDate(row[dateIdx]);
        if (!transactionDate) return;
        
        // Check if transaction is within our month
        if (transactionDate >= startDate && transactionDate <= endDate) {
          var amount = Number(row[totalIdx]) || 0;
          
          // Find which week this transaction belongs to
          for (var i = 0; i < weeks.length; i++) {
            if (transactionDate >= weeks[i].startDate && transactionDate <= weeks[i].endDate) {
              weeks[i].purchases += amount;
              break;
            }
          }
        }
      });
    }
  }
  
  // Process expenses data
  if (expenseSheet) {
    var expenseData = expenseSheet.getDataRange().getValues();
    if (expenseData.length > 1) {
      var expenseHeaders = expenseData.shift();
      var expenseCols = getTransactionColumns(expenseHeaders, 'EXPENSES');
      var dateIdx = expenseCols.date;
      var amountIdx = expenseCols.amount;
      var outletIdx = expenseCols.outlet;
      
      expenseData.forEach(function(row) {
        // Filter by outlet
        if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;
        if (!row || !row[dateIdx]) return;
        
        var transactionDate = parseTransactionDate(row[dateIdx]);
        if (!transactionDate) return;
        
        // Check if transaction is within our month
        if (transactionDate >= startDate && transactionDate <= endDate) {
          var amount = Number(row[amountIdx]) || 0;
          
          // Find which week this transaction belongs to
          for (var i = 0; i < weeks.length; i++) {
            if (transactionDate >= weeks[i].startDate && transactionDate <= weeks[i].endDate) {
              weeks[i].expenses += amount;
              break;
            }
          }
        }
      });
    }
  }
  
  // Calculate profit for each week and format output
  var result = weeks.map(function(week) {
    week.profit = week.sales - week.purchases - week.expenses;
    
    // Format dates for display
    var weekStartDay = week.startDate.getDate();
    var weekEndDay = week.endDate.getDate();
    
    // Determine if this week is in the future or partial
    var currentDate = new Date();
    var isFuture = week.startDate > currentDate;
    var isPartial = week.startDate <= currentDate && week.endDate > currentDate;
    
    return {
      label: "Week " + week.weekNumber + " (" + weekStartDay + "-" + weekEndDay + ")",
      sales: week.sales,
      purchases: week.purchases,
      expenses: week.expenses,
      profit: week.profit,
      isFuture: isFuture,
      isPartial: isPartial,
      startDate: week.startDate.getTime(), // Send as timestamp for client
      endDate: week.endDate.getTime()
    };
  });
  
  Logger.log("Weekly breakdown result: " + JSON.stringify(result));
  return result;
}

// Helper function to parse transaction dates consistently
function parseTransactionDate(dateValue) {
  if (!dateValue) return null;
  
  var d;
  if (dateValue instanceof Date) {
    d = dateValue;
  } else if (typeof dateValue === 'string') {
    // Try multiple date formats (same order as getDayKey for consistency)
    try {
      d = Utilities.parseDate(dateValue, Session.getScriptTimeZone(), 'MM-dd-yyyy HH:mm:ss');
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(dateValue, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(dateValue, Session.getScriptTimeZone(), 'MM-dd-yyyy');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(dateValue, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(dateValue, Session.getScriptTimeZone(), 'dd-MM-yyyy');
      }
      if (!d || isNaN(d.getTime())) {
        d = Utilities.parseDate(dateValue, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
    } catch (e) {
      Logger.log('Utilities.parseDate failed for transaction date: ' + e);
    }
    
    // Fallback to standard Date parsing
    if (!d || isNaN(d.getTime())) {
      d = new Date(dateValue);
    }
  } else {
    d = new Date(dateValue);
  }
  
  if (!d || isNaN(d.getTime())) {
    Logger.log('Failed to parse transaction date: ' + JSON.stringify(dateValue));
    return null;
  }
  
  // Normalize to midnight so date-range comparisons (which use midnight boundaries)
  // correctly include transactions that occur at any time of day.
  d.setHours(0, 0, 0, 0);
  return d;
}

// Debug function to analyze data and find missing June 30th
function debugProfitLossData(deviceId, username) {
  Logger.log("=== DEBUGGING P&L DATA FOR JUNE 30, 2025 ISSUE ===");
  
  // Get user session for outlet filtering
  var session = {};
  if (deviceId && username) {
    session = getUserSessionEnhanced(deviceId, username) || {};
  } else {
    session = getUserSessionEnhanced(deviceId) || {};
  }
  
  var userOutlet = session.outlet;
  Logger.log("Debug: User outlet filter: " + (userOutlet || 'NONE'));
  
  var salesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
  var expenseSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.EXPENSES);
  var purchaseSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
  
  var currentYear = new Date().getFullYear();
  var targetDate = currentYear + '-06-30';
  Logger.log("Looking for target date: " + targetDate);
  
  var debugReport = {
    salesData: { total: 0, june2025: 0, june30: 0, june30Entries: [] },
    purchaseData: { total: 0, june2025: 0, june30: 0, june30Entries: [] },
    expenseData: { total: 0, june2025: 0, june30: 0, june30Entries: [] },
    dateFormats: [],
    dateRange: '',
    currentYear: currentYear
  };
  
  // Debug sales data
  if (salesSheet) {
    var salesData = salesSheet.getDataRange().getValues();
    if (salesData.length > 1) {
      var salesHeaders = salesData.shift();
      var salesCols = getTransactionColumns(salesHeaders, 'ORDERS');
      var dateIdx = salesCols.date;
      var totalIdx = salesCols.total;
      var outletIdx = salesCols.outlet;
      
      Logger.log("DEBUG Sales - Headers: " + JSON.stringify(salesHeaders));
      Logger.log("DEBUG Sales - Date column index: " + dateIdx + ", Total column: " + totalIdx + ", Outlet column: " + outletIdx);
      
      debugReport.salesData.total = salesData.length;
      
      salesData.forEach(function(row, index) {
        if (!row || !row[dateIdx]) return;
        
        // Apply outlet filter
        if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;
        
        var originalDate = row[dateIdx];
        var dateKey = getDayKey(originalDate);
        
        // Track unique date formats
        if (typeof originalDate === 'string' && debugReport.dateFormats.indexOf(originalDate.slice(0, 10)) === -1) {
          debugReport.dateFormats.push(originalDate.slice(0, 10));
        }
        
        if (dateKey && dateKey.startsWith(currentYear + '-06')) {
          debugReport.salesData.june2025++;
          
          if (dateKey === targetDate) {
            debugReport.salesData.june30++;
            debugReport.salesData.june30Entries.push({
              rowIndex: index + 2, // +2 because we shifted headers and 1-based indexing
              originalDate: originalDate,
              parsedDate: dateKey,
              total: row[totalIdx],
              outlet: outletIdx !== -1 ? row[outletIdx] : 'N/A'
            });
            Logger.log("🎯 SALES: Found June 30, " + currentYear + " entry at row " + (index + 2) + ": " + JSON.stringify(row));
          }
        }
      });
    }
  }
  
  // Debug purchase data
  if (purchaseSheet) {
    var purchaseData = purchaseSheet.getDataRange().getValues();
    if (purchaseData.length > 1) {
      var purchaseHeaders = purchaseData.shift();
      var purchaseCols = getTransactionColumns(purchaseHeaders, 'PURCHASES');
      var dateIdx = purchaseCols.date;
      var totalIdx = purchaseCols.total;
      var outletIdx = purchaseCols.outlet;
      
      Logger.log("DEBUG Purchases - Headers: " + JSON.stringify(purchaseHeaders));
      Logger.log("DEBUG Purchases - Date column index: " + dateIdx + ", Total column: " + totalIdx + ", Outlet column: " + outletIdx);
      
      debugReport.purchaseData.total = purchaseData.length;
      
      purchaseData.forEach(function(row, index) {
        if (!row || !row[dateIdx]) return;
        
        // Apply outlet filter
        if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;
        
        var originalDate = row[dateIdx];
        var dateKey = getDayKey(originalDate);
        
        if (dateKey && dateKey.startsWith(currentYear + '-06')) {
          debugReport.purchaseData.june2025++;
          
          if (dateKey === targetDate) {
            debugReport.purchaseData.june30++;
            debugReport.purchaseData.june30Entries.push({
              rowIndex: index + 2,
              originalDate: originalDate,
              parsedDate: dateKey,
              total: row[totalIdx],
              outlet: outletIdx !== -1 ? row[outletIdx] : 'N/A'
            });
            Logger.log("🎯 PURCHASES: Found June 30, " + currentYear + " entry at row " + (index + 2) + ": " + JSON.stringify(row));
          }
        }
      });
    }
  }
  
  // Debug expense data  
  if (expenseSheet) {
    var expenseData = expenseSheet.getDataRange().getValues();
    if (expenseData.length > 1) {
      var expenseHeaders = expenseData.shift();
      var expenseCols = getTransactionColumns(expenseHeaders, 'EXPENSES');
      var dateIdx = expenseCols.date;
      var amountIdx = expenseCols.amount;
      var outletIdx = expenseCols.outlet;
      
      Logger.log("DEBUG Expenses - Headers: " + JSON.stringify(expenseHeaders));
      Logger.log("DEBUG Expenses - Date column index: " + dateIdx + ", Amount column: " + amountIdx + ", Outlet column: " + outletIdx);
      
      debugReport.expenseData.total = expenseData.length;
      
      expenseData.forEach(function(row, index) {
        if (!row || !row[dateIdx]) return;
        
        // Apply outlet filter
        if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;
        
        var originalDate = row[dateIdx];
        var dateKey = getDayKey(originalDate);
        
        if (dateKey && dateKey.startsWith(currentYear + '-06')) {
          debugReport.expenseData.june2025++;
          
          if (dateKey === targetDate) {
            debugReport.expenseData.june30++;
            debugReport.expenseData.june30Entries.push({
              rowIndex: index + 2,
              originalDate: originalDate,
              parsedDate: dateKey,
              amount: row[amountIdx],
              outlet: outletIdx !== -1 ? row[outletIdx] : 'N/A'
            });
            Logger.log("🎯 EXPENSES: Found June 30, " + currentYear + " entry at row " + (index + 2) + ": " + JSON.stringify(row));
          }
        }
      });
    }
  }
  
  // Summary
  Logger.log("=== DEBUG SUMMARY ===");
  Logger.log("Sales data: " + debugReport.salesData.total + " total, " + debugReport.salesData.june2025 + " June " + currentYear + ", " + debugReport.salesData.june30 + " June 30");
  Logger.log("Purchase data: " + debugReport.purchaseData.total + " total, " + debugReport.purchaseData.june2025 + " June " + currentYear + ", " + debugReport.purchaseData.june30 + " June 30");
  Logger.log("Expense data: " + debugReport.expenseData.total + " total, " + debugReport.expenseData.june2025 + " June " + currentYear + ", " + debugReport.expenseData.june30 + " June 30");
  Logger.log("Sample date formats: " + JSON.stringify(debugReport.dateFormats.slice(0, 10)));
  
  return debugReport;
}

// Debug function to test profit/loss report
function testProfitLossReport() {
  Logger.log("=== Testing Profit/Loss Report ===");
  
  try {
    var result = getProfitLossReport();
    Logger.log("Test result: " + JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log("Error in test: " + error.toString());
    return [];
  }
}

function getMonthlyBreakdown(deviceId, username, year) {
  year = Number(year) || new Date().getFullYear();
  return getPnLBundle(deviceId, username, year).monthly;
}

function getMonthlyCategoryBreakdown(deviceId, username, year, month) {
  return getPnLMonthlyCategoryFromBundle(deviceId, username, year, month);
}

function getCategoryPnLBreakdown(deviceId, username, year) {
  year = Number(year) || new Date().getFullYear();
  return getPnLBundle(deviceId, username, year).category;
}

function getSupplierPnLBreakdown(deviceId, username, year) {
  Logger.log("=== Generating Supplier Purchase Breakdown ===");
  Logger.log("Year: " + year);

  var session = {};
  if (deviceId && username) {
    session = getUserSessionEnhanced(deviceId, username) || {};
  } else {
    session = getUserSessionEnhanced(deviceId) || {};
  }

  var userOutlet = session.outlet;
  Logger.log("User outlet: " + (userOutlet || 'NONE SET'));

  var purchaseSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
  var supplierData = {};
  var totalPurchases = 0;
  var totalSales = 0;

  function normalizeLookupValue(value) {
    return (value || '').toString().trim().toUpperCase();
  }

  function findObjectValue(obj, matcher) {
    var keys = Object.keys(obj || {});
    for (var i = 0; i < keys.length; i++) {
      if (matcher.test(keys[i])) {
        return obj[keys[i]];
      }
    }
    return '';
  }

  function addProductCodeToSupplier(bucket, productCode) {
    var cleanCode = (productCode || '').toString().trim();
    if (!cleanCode) return;
    if (!bucket.productCodeMap[cleanCode]) {
      bucket.productCodeMap[cleanCode] = true;
      bucket.productCodes.push(cleanCode);
    }
  }

  function ensureSupplierBucket(name) {
    if (!supplierData[name]) {
      supplierData[name] = {
        supplier: name,
        purchases: 0,
        sales: 0,
        count: 0,
        productCodes: [],
        productCodeMap: {}
      };
    }
    return supplierData[name];
  }

  var productLookup = {
    byCode: {},
    byCategoryName: {}
  };

  try {
    var products = JSON.parse(getProducts(null, userOutlet) || '[]');
    products.forEach(function(product) {
      var supplierValue = findObjectValue(product, /supplier|vendor/i);
      var cleanSupplier = (supplierValue || '').toString().trim();
      var category = findObjectValue(product, /^category$/i) || findObjectValue(product, /category/i);
      var name = findObjectValue(product, /^name$/i) || findObjectValue(product, /product.*name|item.*name/i);
      var explicitCode = findObjectValue(product, /product ?code|^code$/i);
      var productCode = generateProductCode({
        productCode: (explicitCode || '').toString().trim(),
        category: (category || '').toString().trim(),
        name: (name || '').toString().trim()
      });

      if (productCode) {
        productLookup.byCode[normalizeLookupValue(productCode)] = cleanSupplier;
      }

      if (category || name) {
        productLookup.byCategoryName[
          normalizeLookupValue(category) + '||' + normalizeLookupValue(name)
        ] = cleanSupplier;
      }
    });
  } catch (lookupError) {
    Logger.log('Supplier lookup build failed: ' + lookupError.toString());
  }

  if (purchaseSheet) {
    var purchaseRows = purchaseSheet.getDataRange().getValues();
    if (purchaseRows.length > 1) {
      var purchaseHeaders = purchaseRows.shift();
      var purchaseCols = getTransactionColumns(purchaseHeaders, 'PURCHASES');
      var dateIdx = purchaseCols.date;
      var totalIdx = purchaseCols.total;
      var outletIdx = purchaseCols.outlet;
      var supplierIdx = purchaseCols.supplier;
      var productIdx = purchaseCols.productName;
      var categoryIdx = purchaseCols.category;
      var codeIdx = purchaseCols.productCode;

      Logger.log("Purchases supplier column index: " + supplierIdx);

      purchaseRows.forEach(function(row) {
        if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;
        if (!row || !row[dateIdx]) return;

        var dayKey = getDayKey(row[dateIdx]);
        if (!dayKey) return;
        if (parseInt(dayKey.substring(0, 4), 10) !== year) return;

        var directSupplier = (supplierIdx !== -1 && row[supplierIdx]) ? row[supplierIdx].toString().trim() : '';
        var productName = productIdx !== -1 ? (row[productIdx] || '').toString().trim() : '';
        var category = categoryIdx !== -1 ? (row[categoryIdx] || '').toString().trim() : '';
        var rawCode = codeIdx !== -1 ? (row[codeIdx] || '').toString().trim() : '';
        var generatedCode = generateProductCode({
          productCode: rawCode,
          category: category,
          name: productName
        });

        var supplier = directSupplier;
        if (!supplier && generatedCode) {
          supplier = productLookup.byCode[normalizeLookupValue(generatedCode)] || '';
        }
        if (!supplier && (category || productName)) {
          supplier = productLookup.byCategoryName[
            normalizeLookupValue(category) + '||' + normalizeLookupValue(productName)
          ] || '';
        }
        if (!supplier) supplier = 'Not found supplier';

        var bucket = ensureSupplierBucket(supplier);
        if (generatedCode) {
          addProductCodeToSupplier(bucket, generatedCode);
        }

        var amount = Number(row[totalIdx]) || 0;
        bucket.purchases += amount;
        bucket.count += 1;
        totalPurchases += amount;
      });
    }
  }

  // Process Orders (sales) sheet — map sales to suppliers via product code lookup
  var salesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
  if (salesSheet) {
    var salesRows = salesSheet.getDataRange().getValues();
    if (salesRows.length > 1) {
      var salesHeaders = salesRows.shift();
      var salesCols = getTransactionColumns(salesHeaders, 'ORDERS');
      var sDateIdx = salesCols.date;
      var sTotalIdx = salesCols.total;
      var sOutletIdx = salesCols.outlet;
      var sProductNameIdx = salesCols.productName;
      var sCategoryIdx = salesCols.category;
      var sCodeIdx = salesCols.productCode;

      salesRows.forEach(function(row) {
        if (sOutletIdx !== -1 && userOutlet && row[sOutletIdx] !== userOutlet) return;
        if (!row || !row[sDateIdx]) return;
        var dayKey = getDayKey(row[sDateIdx]);
        if (!dayKey) return;
        if (parseInt(dayKey.substring(0, 4), 10) !== year) return;

        var productName = sProductNameIdx !== -1 ? (row[sProductNameIdx] || '').toString().trim() : '';
        var category = sCategoryIdx !== -1 ? (row[sCategoryIdx] || '').toString().trim() : '';
        var rawCode = sCodeIdx !== -1 ? (row[sCodeIdx] || '').toString().trim() : '';
        var generatedCode = generateProductCode({ productCode: rawCode, category: category, name: productName });

        var supplier = '';
        if (generatedCode) {
          supplier = productLookup.byCode[normalizeLookupValue(generatedCode)] || '';
        }
        if (!supplier && (category || productName)) {
          supplier = productLookup.byCategoryName[normalizeLookupValue(category) + '||' + normalizeLookupValue(productName)] || '';
        }
        if (!supplier) supplier = 'Not found supplier';

        var bucket = ensureSupplierBucket(supplier);
        var amount = Number(row[sTotalIdx]) || 0;
        bucket.sales += amount;
        totalSales += amount;
      });
    }
  }

  var suppliers = Object.keys(supplierData).map(function(name) {
    var item = supplierData[name];
    return {
      supplier: item.supplier,
      purchases: item.purchases,
      sales: item.sales,
      profit: item.sales - item.purchases,
      count: item.count,
      productCodes: item.productCodes.sort(),
      pct: totalPurchases > 0 ? Math.round((item.purchases / totalPurchases) * 1000) / 10 : 0
    };
  }).sort(function(a, b) {
    return b.purchases - a.purchases;
  });

  Logger.log("Supplier breakdown: " + suppliers.length + " suppliers, total purchases: " + totalPurchases + ", total sales: " + totalSales);
  return { suppliers: suppliers, totalPurchases: totalPurchases, totalSales: totalSales };
}

// Debug function to test weekly breakdown
function testWeeklyBreakdown() {
  Logger.log("=== Testing Weekly Breakdown ===");
  
  try {
    var currentDate = new Date();
    var year = currentDate.getFullYear();
    var month = currentDate.getMonth() + 1; // Convert to 1-based month
    
    var result = getWeeklyBreakdown(null, null, year, month);
    Logger.log("Weekly test result: " + JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log("Error in weekly test: " + error.toString());
    return [];
  }
}