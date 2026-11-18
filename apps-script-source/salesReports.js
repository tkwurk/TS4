function getSalesSummaryData(deviceId, username) {
  // Get user's outlet
  const session = getUserSessionEnhanced(deviceId, username) || {};
  const userOutlet = session.outlet;

  const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const outletIdx = headers.findIndex(h => /outlet/i.test(h));
  const dateIdx = headers.findIndex(h => /date|timestamp/i.test(h));
  const productIdx = headers.findIndex(h => /product|item|name/i.test(h));
  const qtyIdx = headers.findIndex(h => /qty|quantity/i.test(h));
  const totalIdx = headers.findIndex(h => /total|amount/i.test(h));
  const productCodeIdx = headers.findIndex(h => /product.*code|code/i.test(h));

  // Filter data by outlet and map to required format
  return data
    .filter(row => outletIdx === -1 || (typeof outletRowMatchesUser_ === 'function'
      ? outletRowMatchesUser_(row[outletIdx], userOutlet)
      : row[outletIdx] === userOutlet))
    .map(row => {
      const productCode = productCodeIdx >= 0 && row[productCodeIdx] ? 
        row[productCodeIdx] : 
        ((row[headers.findIndex(h => /category/i.test(h))] || 'OTHER') + '_' + (row[productIdx] || 'UNKNOWN')).replace(/\s+/g, '_').toUpperCase();
      
      return {
        date: dateIdx >= 0 ? row[dateIdx] : '',
        product: productIdx >= 0 ? row[productIdx] : '',
        qty: qtyIdx >= 0 ? row[qtyIdx] : '',
        total: totalIdx >= 0 ? row[totalIdx] : '',
        productCode: productCode
      };
    });
}

/**
 * Reads only the rows of `sheet` that fall on one business day. Date-ordered
 * (ascending) sheets are binary-searched on a single date column, then just that
 * row block is read — far cheaper than getDataRange() on the whole sheet. If the
 * date column is not usable for the search (blank/invalid cells, not ascending),
 * falls back to reading every row, so results are always correct.
 * @return {{headers: Array, rows: Array[], firstRow: number}} firstRow = 1-based sheet row of rows[0].
 */
function readRowsForDay_(sheet, dateHeaderRe, startOfDay, endOfDay) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { headers: lastCol >= 1 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [], rows: [], firstRow: 2 };
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dateIdx = headers.findIndex(function(h) { return h && dateHeaderRe.test(h); });

  function fullRead() {
    return { headers: headers, rows: sheet.getRange(2, 1, lastRow - 1, lastCol).getValues(), firstRow: 2 };
  }
  if (dateIdx < 0) return fullRead();

  var col = sheet.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues();
  var n = col.length;
  var ms = new Array(n);
  for (var i = 0; i < n; i++) {
    var v = col[i][0];
    var t = v instanceof Date ? v.getTime() : (v ? new Date(v).getTime() : NaN);
    if (isNaN(t)) return fullRead();
    if (i > 0 && t < ms[i - 1]) return fullRead();
    ms[i] = t;
  }

  var startMs = startOfDay.getTime();
  var endMs = endOfDay.getTime();
  var lo = 0, hi = n;
  while (lo < hi) { var m = (lo + hi) >> 1; if (ms[m] < startMs) lo = m + 1; else hi = m; }
  var first = lo;
  hi = n;
  while (lo < hi) { var m2 = (lo + hi) >> 1; if (ms[m2] <= endMs) lo = m2 + 1; else hi = m2; }
  var count = lo - first;
  if (count <= 0) return { headers: headers, rows: [], firstRow: first + 2 };
  return {
    headers: headers,
    rows: sheet.getRange(first + 2, 1, count, lastCol).getValues(),
    firstRow: first + 2
  };
}

function getSalesSummaryByDate(params) {
  try {
    // Handle both SecurityWrapper format and legacy direct calls
    let dateString, deviceId, username;
    
    if (params && typeof params === 'object' && params._authContext) {
      // SecurityWrapper format
      dateString = params.dateString || params.date;
      deviceId = params._authContext.deviceId;
      username = params._authContext.username;
    } else if (arguments.length >= 3) {
      // Legacy format: getSalesSummaryByDate(dateString, deviceId, username)
      dateString = arguments[0];
      deviceId = arguments[1];
      username = arguments[2];
    } else if (typeof params === 'string') {
      // Simple call with just date string - this is the problematic case
      dateString = params;
      return emptySalesSummary(dateString);
    } else {
      return emptySalesSummary("unknown");
    }
    
    // Fix corrupted date parameter if needed
    let fixedDateString = dateString;
    if (typeof dateString === 'object' && dateString !== null) {
      if (dateString.hasOwnProperty('0')) {
        // Reconstruct from indexed object
        const chars = [];
        for (let i = 0; dateString.hasOwnProperty(i.toString()); i++) {
          chars.push(dateString[i.toString()]);
        }
        fixedDateString = chars.join('');
      } else {
        fixedDateString = new Date().toISOString().split('T')[0]; // Today's date as fallback
      }
    } else if (typeof dateString !== 'string') {
      fixedDateString = String(dateString);
    }
    
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    
    if (!sheet) {
      return emptySalesSummary(fixedDateString);
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      return emptySalesSummary(fixedDateString);
    }
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Get current user's outlet
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;
    
    // Handle case where outlet might be the string "undefined"
    let actualUserOutlet = userOutlet;
    if (!userOutlet || userOutlet === "undefined" || userOutlet === undefined) {
      // Try to derive outlet from username as fallback
      if (username && username.toLowerCase() === 'cineam') {
        actualUserOutlet = 'Cineam';
      } else {
        return emptySalesSummary(fixedDateString);
      }
    } else {
      actualUserOutlet = userOutlet;
    }

    const outletIdx = headers.findIndex(h => /outlet/i.test(h));
    const dateIdx = headers.findIndex(h => /timestamp|date|tanggal|waktu|time/i.test(h));
    const productIdx = headers.findIndex(h => /product|item|name/i.test(h));
    const qtyIdx = headers.findIndex(h => /qty|quantity/i.test(h));
    const priceIdx = headers.findIndex(h => /price/i.test(h));
    const totalIdx = headers.findIndex(h => /total|amount/i.test(h));
    const categoryIdx = headers.findIndex(h => /category/i.test(h));
    const orderNumberIdx = headers.findIndex(h => /order\s*number/i.test(h));
    const paymentIdx = headers.findIndex(h => /payment/i.test(h));
    const productCodeIdx = headers.findIndex(h => /product.*code|code/i.test(h));

    // Parse target date - create date object for the target day
    let targetDate;
    try {
      if (fixedDateString.includes('-')) {
        // Handle YYYY-MM-DD format
        const dateParts = fixedDateString.split('-');
        if (dateParts.length === 3) {
          const year = parseInt(dateParts[0]);
          const month = parseInt(dateParts[1]) - 1; // Month is 0-based
          const day = parseInt(dateParts[2]);
          targetDate = new Date(year, month, day);
        }
      }
      
      // Fallback to standard Date parsing if above fails
      if (!targetDate || isNaN(targetDate.getTime())) {
        targetDate = new Date(fixedDateString);
      }
    } catch (parseError) {
      targetDate = new Date(fixedDateString);
    }
    
    if (!targetDate || isNaN(targetDate.getTime())) {
      return emptySalesSummary(fixedDateString);
    }
    
    // Set target date to start of day (00:00:00) for comparison
    const targetStartOfDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
    const targetEndOfDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 23, 59, 59, 999);

    const paymentConfigMap = typeof loadPaymentMethodConfigMap_ === 'function' ? loadPaymentMethodConfigMap_() : {};
    const summary = {};
    let totalSales = 0;
    const ordersPaymentTotals = {};
    const orderNumbers = new Set();
    const baseOrderNumbers = new Set(); // Track unique base orders
    let processedRows = 0;
    let outletFilteredRows = 0;
    let dateFilteredRows = 0;

    const ORDER_DATE_RE = /timestamp|date|tanggal|waktu|time/i;
    const dayOrders = readRowsForDay_(sheet, ORDER_DATE_RE, targetStartOfDay, targetEndOfDay);

    dayOrders.rows.forEach((row, index) => {
      processedRows++;
      
      // Filter by outlet first - only process rows that match user's outlet
      if (outletIdx >= 0) {
        const rowOutlet = row[outletIdx];
        
        if (!actualUserOutlet) {
          outletFilteredRows++;
          return;
        }
        
        // Convert both to strings for comparison to handle type mismatches
        const rowOutletStr = String(rowOutlet).trim();
        const userOutletStr = String(actualUserOutlet).trim();
        
        if (typeof outletRowMatchesUser_ === 'function') {
          if (!outletRowMatchesUser_(rowOutlet, actualUserOutlet)) {
            outletFilteredRows++;
            return;
          }
        } else if (rowOutletStr !== userOutletStr) {
          outletFilteredRows++;
          return;
        }
      }

      // Get order date from the row
      const orderDateRaw = row[dateIdx];
      if (!orderDateRaw) {
        return;
      }
      
      let orderDate;
      if (orderDateRaw instanceof Date) {
        orderDate = orderDateRaw;
      } else {
        // Convert to Date object regardless of format
        orderDate = new Date(orderDateRaw);
      }
      
      if (isNaN(orderDate.getTime())) {
        return;
      }
      
      // Compare using date values instead of string formatting
      // Check if order date falls within the target day (00:00:00 to 23:59:59)
      if (orderDate < targetStartOfDay || orderDate > targetEndOfDay) {
        dateFilteredRows++;
        return;
      }

      const orderNumber = row[orderNumberIdx];
      const itemName = row[productIdx];
      const qty = Number(row[qtyIdx]) || 0;
      const price = Number(row[priceIdx]) || 0;
      const total = Number(row[totalIdx]) || 0;
      const category = row[categoryIdx] || 'Other';
      const payment = row[paymentIdx] || '';
      const productCode = row[productCodeIdx] || (category + '_' + itemName).replace(/\s+/g, '_').toUpperCase();

      // Calculate actual sales amount using quantity × price
      const calculatedTotal = qty * price;
      const salesAmount = calculatedTotal > 0 ? calculatedTotal : total; // Use calculated amount, fallback to stored total

      // Skip only if missing essential data (allow zero sales amount for items like "Ayam Sisa")
      if (!itemName || !qty) {
        return;
      }
      
      orderNumbers.add(orderNumber);
      
      // Extract base order number (remove item suffix) for order counting
      const baseOrderNumber = orderNumber.replace(/-\d{3}$/, '');
      baseOrderNumbers.add(baseOrderNumber);
      totalSales += salesAmount;

      // Non-auto-deposit methods only (Cash/split recorded as cash on Orders — excluded here)
      const paymentLower = payment.toLowerCase().trim();
      if (paymentLower && paymentLower !== 'none' &&
          !isCashLikePaymentMethodName_(paymentLower) &&
          !isPartialPaymentLabel_(payment) &&
          typeof paymentMethodHasAutoDeposit_ === 'function' &&
          !paymentMethodHasAutoDeposit_(payment, paymentConfigMap)) {
        ordersPaymentTotals[paymentLower] = (ordersPaymentTotals[paymentLower] || 0) + salesAmount;
      }

      if (!summary[category]) summary[category] = {};
      // Use product code as unique key instead of just item name
      const uniqueKey = productCode;
      // Create separate entries for paid vs free items of the same type
      if (price === 0) {
        // This is a leftover/free item - create separate leftover key
        const leftoverKey = uniqueKey + '_LEFTOVER';
        if (!summary[category][leftoverKey]) {
          summary[category][leftoverKey] = { 
            qty: 0, 
            sales: 0, 
            itemName: itemName + ' (Leftover)',
            productCode: productCode,
            price: 0,
            isLeftover: true
          };
        }
        summary[category][leftoverKey].qty += qty;
        summary[category][leftoverKey].sales += 0; // Always 0 for leftover
      } else {
        // This is a regular paid item
        if (!summary[category][uniqueKey]) {
          summary[category][uniqueKey] = { 
            qty: 0, 
            sales: 0, 
            itemName: itemName,
            productCode: productCode,
            price: price,
            isLeftover: false
          };
        }
        summary[category][uniqueKey].qty += qty;
        summary[category][uniqueKey].sales += salesAmount;
      }
    });
    
    // Deposits for the same date — auto-deposit methods for payment breakdown
    let totalDeposits = 0;
    let totalDepositsByQris = 0;
    const depositsPaymentTotals = {};
    const accountTypeMap = typeof loadAccountTypeByNameMap_ === 'function'
      ? loadAccountTypeByNameMap_()
      : {};
    try {
      const depositSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.DEPOSITS);
      if (depositSheet) {
        const dayDeposits = readRowsForDay_(depositSheet, /date|timestamp/i, targetStartOfDay, targetEndOfDay);
        const depositData = [dayDeposits.headers].concat(dayDeposits.rows);
        if (depositData.length > 1) {
          const depositHeaders = depositData[0];
          const depositDateIdx = depositHeaders.findIndex(h => /date|timestamp/i.test(h));
          const depositAmountIdx = depositHeaders.findIndex(h => /amount/i.test(h));
          const depositAccountIdx = depositHeaders.findIndex(h => /account/i.test(h));
          const depositOutletIdx = depositHeaders.findIndex(h => /outlet/i.test(h));
          const depositPayMethodIdx = depositHeaders.findIndex(h => /paymentmethod|payment.method/i.test(h));
          const depositAccum = {
            totalDeposits: 0,
            totalDepositsByQris: 0,
            depositsPaymentTotals: depositsPaymentTotals
          };
          
          depositData.slice(1).forEach((row) => {
            if (depositOutletIdx !== -1) {
              const rowOutlet = row[depositOutletIdx];
              if (!actualUserOutlet) return;
              if (typeof outletRowMatchesUser_ === 'function') {
                if (!outletRowMatchesUser_(rowOutlet, actualUserOutlet)) return;
              } else if (String(rowOutlet).trim() !== String(actualUserOutlet).trim()) {
                return;
              }
            }
            
            if (depositDateIdx === -1 || !row[depositDateIdx]) return;
            const depositDateRaw = row[depositDateIdx];
            let depositDate = depositDateRaw instanceof Date ? depositDateRaw : new Date(depositDateRaw);
            if (isNaN(depositDate.getTime())) return;
            if (depositDate < targetStartOfDay || depositDate > targetEndOfDay) return;

            accumulateDepositForSalesBreakdown_(row, {
              amountIdx: depositAmountIdx,
              accountIdx: depositAccountIdx,
              payMethodIdx: depositPayMethodIdx
            }, accountTypeMap, paymentConfigMap, depositAccum);
          });
          totalDeposits = depositAccum.totalDeposits;
          totalDepositsByQris = depositAccum.totalDepositsByQris;
        }
      }
    } catch (depositError) {
      // Continue without deposits data
    }

    const paymentBreakdown = buildSalesByPaymentMethodBreakdown_({
      totalSales: totalSales,
      ordersPaymentTotals: ordersPaymentTotals,
      depositsPaymentTotals: depositsPaymentTotals,
      configMap: paymentConfigMap
    });
    const salesByPaymentMethod = paymentBreakdown.salesByPaymentMethod;
    const totalSalesByCash = paymentBreakdown.totalSalesByCash;
    const totalSalesByNonCash = paymentBreakdown.totalSalesByNonCash;
    
    // Get missed orders for the same date
    let missedOrdersCount = 0;
    let missedOrdersTotal = 0;
    try {
      const missedSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.MISSED_ORDERS);
      if (missedSheet) {
        const dayMissed = readRowsForDay_(missedSheet, /timestamp|date/i, targetStartOfDay, targetEndOfDay);
        const missedData = [dayMissed.headers].concat(dayMissed.rows);
        if (missedData.length > 1) {
          const mHeaders = missedData[0];
          const mDateIdx = mHeaders.findIndex(h => /timestamp|date/i.test(h));
          const mOutletIdx = mHeaders.findIndex(h => /outlet/i.test(h));
          const mTotalIdx = mHeaders.findIndex(h => /^total$/i.test(h));
          const mQtyIdx = mHeaders.findIndex(h => /quantity|qty/i.test(h));

          missedData.slice(1).forEach(row => {
            if (mOutletIdx !== -1 && actualUserOutlet) {
              if (String(row[mOutletIdx]).trim() !== String(actualUserOutlet).trim()) return;
            }
            if (mDateIdx !== -1 && row[mDateIdx]) {
              let mDate = row[mDateIdx] instanceof Date ? row[mDateIdx] : new Date(row[mDateIdx]);
              if (!isNaN(mDate.getTime()) && mDate >= targetStartOfDay && mDate <= targetEndOfDay) {
                missedOrdersCount += (parseFloat(row[mQtyIdx]) || 0);
                missedOrdersTotal += (parseFloat(row[mTotalIdx]) || 0);
              }
            }
          });
        }
      }
    } catch (missedError) {
      // Continue without missed orders data
    }
    
    const result = {
      date: dateString,
      summary: summary,
      totalSales: totalSales,
      totalSalesByCash: totalSalesByCash,
      totalSalesByNonCash: totalSalesByNonCash,
      salesByPaymentMethod: salesByPaymentMethod,
      ordersCount: baseOrderNumbers.size,
      itemsCount: orderNumbers.size,
      totalDeposits: totalDeposits,
      totalDepositsByQris: totalDepositsByQris,
      missedOrdersCount: missedOrdersCount,
      missedOrdersTotal: missedOrdersTotal
    };
    
    return JSON.stringify(result);
  } catch (error) {
    console.error("Error getting sales summary:", error);
    return emptySalesSummary(dateString);
  }
}

function emptySalesSummary(dateString) {
  const result = {
    date: dateString,
    summary: {},
    totalSales: 0,
    totalSalesByCash: 0,
    totalSalesByNonCash: 0,
    salesByPaymentMethod: {},
    ordersCount: 0,
    totalDeposits: 0,
    totalDepositsByQris: 0,
    missedOrdersCount: 0,
    missedOrdersTotal: 0
  };
  
  return JSON.stringify(result);
}

/**
 * Sales Summary payment breakdown:
 * - Auto Deposit methods from Deposits sheet (excluding Asset destination accounts)
 * - Other non-cash methods from Orders (Auto Deposit = No)
 * - Cash = Total Sales − other Orders methods − auto-deposit totals
 */
function accumulateDepositForSalesBreakdown_(row, indices, accountTypeMap, paymentConfigMap, accumulators) {
  accumulators = accumulators || {};
  var amountIdx = indices.amountIdx;
  var accountIdx = indices.accountIdx;
  var payMethodIdx = indices.payMethodIdx;

  var amount = amountIdx >= 0 ? (parseFloat(row[amountIdx]) || 0) : 0;
  if (amount <= 0) return;

  var depAccount = accountIdx >= 0 ? (row[accountIdx] || '').toString().trim() : '';
  if (typeof isDepositAccountExcludedFromSalesCalc_ === 'function' &&
      isDepositAccountExcludedFromSalesCalc_(depAccount, accountTypeMap)) {
    return;
  }

  if (accumulators.totalDeposits !== undefined) {
    accumulators.totalDeposits += amount;
  }

  var depPayRaw = payMethodIdx >= 0 ? (row[payMethodIdx] || '').toString().trim() : '';
  var depMethod = depPayRaw || depAccount;
  if (typeof paymentMethodHasAutoDeposit_ === 'function' &&
      paymentMethodHasAutoDeposit_(depMethod, paymentConfigMap)) {
    var depKey = depPayRaw ? depPayRaw.toLowerCase() : depMethod.toLowerCase();
    if (!accumulators.depositsPaymentTotals) accumulators.depositsPaymentTotals = {};
    accumulators.depositsPaymentTotals[depKey] =
      (accumulators.depositsPaymentTotals[depKey] || 0) + amount;
    if (accumulators.totalDepositsByQris !== undefined) {
      accumulators.totalDepositsByQris += amount;
    }
  }
}

function buildSalesByPaymentMethodBreakdown_(options) {
  options = options || {};
  var configMap = options.configMap || (typeof loadPaymentMethodConfigMap_ === 'function' ? loadPaymentMethodConfigMap_() : {});
  var totalSales = Number(options.totalSales) || 0;
  var ordersPaymentTotals = options.ordersPaymentTotals || {};
  var depositsPaymentTotals = options.depositsPaymentTotals || {};
  var salesByPaymentMethod = {};

  Object.keys(depositsPaymentTotals).forEach(function(methodKey) {
    var amt = depositsPaymentTotals[methodKey] || 0;
    if (amt <= 0) return;
    var label = typeof canonicalPaymentMethodName_ === 'function'
      ? canonicalPaymentMethodName_(methodKey, configMap)
      : methodKey;
    salesByPaymentMethod[label] = (salesByPaymentMethod[label] || 0) + amt;
  });

  Object.keys(ordersPaymentTotals).forEach(function(methodKey) {
    var amt = ordersPaymentTotals[methodKey] || 0;
    if (amt <= 0) return;
    var label = typeof canonicalPaymentMethodName_ === 'function'
      ? canonicalPaymentMethodName_(methodKey, configMap)
      : methodKey;
    salesByPaymentMethod[label] = (salesByPaymentMethod[label] || 0) + amt;
  });

  var totalDepositsAuto = Object.keys(depositsPaymentTotals).reduce(function(sum, key) {
    return sum + (depositsPaymentTotals[key] || 0);
  }, 0);
  var totalOrdersOther = Object.keys(ordersPaymentTotals).reduce(function(sum, key) {
    return sum + (ordersPaymentTotals[key] || 0);
  }, 0);
  var cashResidual = Math.round((totalSales - totalOrdersOther - totalDepositsAuto) * 100) / 100;
  if (cashResidual !== 0) {
    salesByPaymentMethod['Cash'] = (salesByPaymentMethod['Cash'] || 0) + cashResidual;
  }

  var totalSalesByCash = salesByPaymentMethod['Cash'] || 0;
  var totalSalesByNonCash = totalSales - totalSalesByCash;
  return {
    salesByPaymentMethod: salesByPaymentMethod,
    totalSalesByCash: totalSalesByCash,
    totalSalesByNonCash: totalSalesByNonCash,
    totalDepositsAuto: totalDepositsAuto,
    totalOrdersOther: totalOrdersOther
  };
}

/** All-time/outlet-filtered sheet data → payment breakdown (Cash Management, etc.). */
function computeSalesPaymentBreakdownFromSheetData_(salesData, depositData) {
  var configMap = typeof loadPaymentMethodConfigMap_ === 'function' ? loadPaymentMethodConfigMap_() : {};
  var totalSales = 0;
  var ordersPaymentTotals = {};
  var depositsPaymentTotals = {};

  if (salesData && salesData.length > 1) {
    var headers = salesData[0];
    var totalIdx = headers.findIndex(function(h) { return /total|amount/i.test(h); });
    var paymentIdx = headers.findIndex(function(h) { return /payment/i.test(h); });
    var qtyIdx = headers.findIndex(function(h) { return /qty|quantity/i.test(h); });
    var priceIdx = headers.findIndex(function(h) { return /price/i.test(h); });

    salesData.slice(1).forEach(function(row) {
      var qty = qtyIdx >= 0 ? (Number(row[qtyIdx]) || 0) : 0;
      var price = priceIdx >= 0 ? (Number(row[priceIdx]) || 0) : 0;
      var total = totalIdx >= 0 ? (Number(row[totalIdx]) || 0) : 0;
      var salesAmount = qty * price > 0 ? qty * price : total;
      if (salesAmount <= 0) return;
      totalSales += salesAmount;

      var payment = paymentIdx >= 0 ? (row[paymentIdx] || '').toString().trim() : '';
      var paymentLower = payment.toLowerCase();
      if (paymentLower && paymentLower !== 'none' &&
          !isCashLikePaymentMethodName_(paymentLower) &&
          !isPartialPaymentLabel_(payment) &&
          typeof paymentMethodHasAutoDeposit_ === 'function' &&
          !paymentMethodHasAutoDeposit_(payment, configMap)) {
        ordersPaymentTotals[paymentLower] = (ordersPaymentTotals[paymentLower] || 0) + salesAmount;
      }
    });
  }

  if (depositData && depositData.length > 1) {
    var dHeaders = depositData[0];
    var amountIdx = dHeaders.findIndex(function(h) { return /amount|total/i.test(h); });
    var accountIdx = dHeaders.findIndex(function(h) { return /account/i.test(h); });
    var payIdx = dHeaders.findIndex(function(h) { return /paymentmethod|payment.method/i.test(h); });
    var accountTypeMap = typeof loadAccountTypeByNameMap_ === 'function'
      ? loadAccountTypeByNameMap_()
      : {};
    var depositAccum = { depositsPaymentTotals: depositsPaymentTotals };

    depositData.slice(1).forEach(function(row) {
      accumulateDepositForSalesBreakdown_(row, {
        amountIdx: amountIdx,
        accountIdx: accountIdx,
        payMethodIdx: payIdx
      }, accountTypeMap, configMap, depositAccum);
    });
  }

  return buildSalesByPaymentMethodBreakdown_({
    totalSales: totalSales,
    ordersPaymentTotals: ordersPaymentTotals,
    depositsPaymentTotals: depositsPaymentTotals,
    configMap: configMap
  });
}

/**
 * All order line items for a business date (outlet-scoped, read-only).
 * Used by cashier/supervisor daily sales — excel-style line listing.
 */
function getOrderLinesByDate(params) {
  try {
    var dateString, deviceId, username;
    if (params && typeof params === 'object' && params._authContext) {
      dateString = params.dateString || params.date;
      deviceId = params._authContext.deviceId;
      username = params._authContext.username;
    } else if (arguments.length >= 3) {
      dateString = arguments[0];
      deviceId = arguments[1];
      username = arguments[2];
    } else if (typeof params === 'string') {
      return JSON.stringify({ date: params, lines: [], count: 0 });
    } else {
      return JSON.stringify({ date: '', lines: [], count: 0 });
    }

    var fixedDateString = dateString;
    if (typeof dateString === 'object' && dateString !== null) {
      if (dateString.hasOwnProperty('0')) {
        var chars = [];
        for (var ci = 0; dateString.hasOwnProperty(ci.toString()); ci++) {
          chars.push(dateString[ci.toString()]);
        }
        fixedDateString = chars.join('');
      } else {
        fixedDateString = new Date().toISOString().split('T')[0];
      }
    } else if (typeof dateString !== 'string') {
      fixedDateString = String(dateString);
    }

    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return JSON.stringify({ date: fixedDateString, lines: [], count: 0 });
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var session = getUserSessionEnhanced(deviceId, username) || {};
    var userOutlet = session.outlet;
    var actualUserOutlet = userOutlet;
    if (!userOutlet || userOutlet === 'undefined' || userOutlet === undefined) {
      if (username && username.toLowerCase() === 'cineam') {
        actualUserOutlet = 'Cineam';
      } else {
        return JSON.stringify({ date: fixedDateString, lines: [], count: 0 });
      }
    }

    var outletIdx = headers.findIndex(function(h) { return h && /outlet/i.test(h); });
    var dateIdx = headers.findIndex(function(h) { return h && /timestamp|date|tanggal|waktu|time/i.test(h); });
    var productIdx = headers.findIndex(function(h) { return h && /product|item|name/i.test(h); });
    var qtyIdx = headers.findIndex(function(h) { return h && /qty|quantity/i.test(h); });
    var priceIdx = headers.findIndex(function(h) { return h && /price/i.test(h); });
    var totalIdx = headers.findIndex(function(h) { return h && /total|amount/i.test(h); });
    var categoryIdx = headers.findIndex(function(h) { return h && /category/i.test(h); });
    var orderNumberIdx = headers.findIndex(function(h) { return h && /order\s*number/i.test(h); });
    var paymentIdx = headers.findIndex(function(h) { return h && /payment/i.test(h); });
    var usernameIdx = headers.findIndex(function(h) { return h && /user/i.test(h); });

    var targetDate;
    try {
      if (fixedDateString.indexOf('-') >= 0) {
        var dateParts = fixedDateString.split('-');
        if (dateParts.length === 3) {
          targetDate = new Date(parseInt(dateParts[0], 10), parseInt(dateParts[1], 10) - 1, parseInt(dateParts[2], 10));
        }
      }
      if (!targetDate || isNaN(targetDate.getTime())) {
        targetDate = new Date(fixedDateString);
      }
    } catch (parseErr) {
      targetDate = new Date(fixedDateString);
    }
    if (!targetDate || isNaN(targetDate.getTime())) {
      return JSON.stringify({ date: fixedDateString, lines: [], count: 0 });
    }

    var targetStartOfDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
    var targetEndOfDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 23, 59, 59, 999);
    var lines = [];

    var dayOrders = readRowsForDay_(sheet, /timestamp|date|tanggal|waktu|time/i, targetStartOfDay, targetEndOfDay);
    var dayRows = dayOrders.rows;

    for (var k = 0; k < dayRows.length; k++) {
      var row = dayRows[k];
      var i = dayOrders.firstRow - 1 + k;
      if (outletIdx >= 0) {
        if (!actualUserOutlet) continue;
        if (typeof outletRowMatchesUser_ === 'function') {
          if (!outletRowMatchesUser_(row[outletIdx], actualUserOutlet)) continue;
        } else if (String(row[outletIdx] || '').trim() !== String(actualUserOutlet).trim()) continue;
      }

      var orderDateRaw = row[dateIdx];
      if (!orderDateRaw) continue;
      var orderDate = (orderDateRaw instanceof Date) ? orderDateRaw : new Date(orderDateRaw);
      if (isNaN(orderDate.getTime())) continue;
      if (orderDate < targetStartOfDay || orderDate > targetEndOfDay) continue;

      var qty = Number(row[qtyIdx]) || 0;
      var price = Number(row[priceIdx]) || 0;
      var total = Number(row[totalIdx]) || 0;
      var itemName = row[productIdx];
      if (!itemName || !qty) continue;

      var lineTotal = (qty * price) > 0 ? qty * price : total;

      lines.push({
        rowIndex: i + 1,
        timestamp: orderDate.toISOString(),
        orderNumber: orderNumberIdx >= 0 ? String(row[orderNumberIdx] || '') : '',
        category: categoryIdx >= 0 ? String(row[categoryIdx] || '') : '',
        productName: String(itemName),
        qty: qty,
        price: price,
        lineTotal: lineTotal,
        payment: paymentIdx >= 0 ? String(row[paymentIdx] || '') : '',
        username: usernameIdx >= 0 ? String(row[usernameIdx] || '') : '',
        outlet: outletIdx >= 0 ? String(row[outletIdx] || '') : ''
      });
    }

    lines.sort(function(a, b) {
      var ta = new Date(a.timestamp).getTime() || 0;
      var tb = new Date(b.timestamp).getTime() || 0;
      if (ta !== tb) return ta - tb;
      return String(a.orderNumber).localeCompare(String(b.orderNumber));
    });

    return JSON.stringify({ date: fixedDateString, lines: lines, count: lines.length });
  } catch (error) {
    Logger.log('getOrderLinesByDate error: ' + error.toString());
    return JSON.stringify({ date: '', lines: [], count: 0, error: error.message });
  }
}

function getMonthlyUsageReport(deviceId, username) {
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
  if (!sheet) return [];

  var session = getUserSessionEnhanced(deviceId, username) || {};
  var userOutlet = session.outlet;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data.shift();

  var dateIdx = headers.findIndex(function(h) { return h && /timestamp|date|tanggal|waktu|time/i.test(h); });
  var catIdx = headers.findIndex(function(h) { return h && /category|kategori/i.test(h); });
  var nameIdx = headers.findIndex(function(h) { return h && /name|product|item/i.test(h); });
  var qtyIdx = headers.findIndex(function(h) { return h && /qty|quantity|jumlah/i.test(h); });
  var outletIdx = headers.findIndex(function(h) { return h && /outlet/i.test(h); });

  // Aggregate: month||category||name → qty
  var agg = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) continue;
    var d = row[dateIdx];
    var category = row[catIdx];
    var name = row[nameIdx];
    var qty = Number(row[qtyIdx]) || 0;
    if (!d || !category || !name) continue;
    var dateObj = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dateObj.getTime())) continue;
    var month = dateObj.getFullYear() + '-' + ('0' + (dateObj.getMonth() + 1)).slice(-2);
    var key = month + '||' + category + '||' + name;
    if (!agg[key]) {
      agg[key] = { month: month, category: category, name: name, qty: 0 };
    }
    agg[key].qty += qty;
  }

  var result = Object.keys(agg).map(function(k) { return agg[k]; });
  result.sort(function(a, b) {
    if (a.month !== b.month) return b.month.localeCompare(a.month);
    if (a.category !== b.category) return a.category.localeCompare(b.category);
    return a.name.localeCompare(b.name);
  });

  return result;
}


/**
 * Helper function to format date for comparison (YYYY-MM-DD)
 * Handles timezone issues by using local date components
 */
function formatDateForComparison(date) {
  if (!date || isNaN(date.getTime())) {
    return null;
  }
  
  // Use local date components to avoid timezone issues
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const formatted = year + '-' + month + '-' + day;
  
  return formatted;
}

/**
 * Helper function to generate a product code based on category and name
 * This mirrors the logic from orders.js to ensure consistency
 * @param {Object} item - Object containing category and name properties
 * @returns {string} - Generated product code in format CATEGORY_PRODUCT_NAME
 */
function generateProductCode(item) {
  // Return existing productCode if available
  if (item.productCode && item.productCode.trim() !== '') {
    return item.productCode.trim();
  }
  
  // Generate product code from category and name
  var category = (item.category || '').toString().trim();
  var name = (item.name || '').toString().trim();
  
  // Handle cases where category or name might be empty
  if (!category && !name) {
    return 'UNKNOWN_PRODUCT';
  } else if (!category) {
    return name.replace(/\s+/g, '_').toUpperCase();
  } else if (!name) {
    return category.replace(/\s+/g, '_').toUpperCase();
  }
  
  // Standard format: CATEGORY_NAME with spaces replaced by underscores and converted to uppercase
  var productCode = (category + '_' + name).replace(/\s+/g, '_').toUpperCase();
  
  // Remove any special characters except underscores and alphanumeric
  productCode = productCode.replace(/[^A-Z0-9_]/g, '');
  
  // Ensure it doesn't start or end with underscore
  productCode = productCode.replace(/^_+|_+$/g, '');
  
  // Replace multiple consecutive underscores with single underscore
  productCode = productCode.replace(/_+/g, '_');
  
  return productCode || 'UNKNOWN_PRODUCT';
}

// Test function to add sample order data for testing
function addTestOrderData(deviceId, username) {
  try {
    Logger.log("=== ADDING TEST ORDER DATA ===");
    
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (!sheet) {
      Logger.log("ERROR: Orders sheet not found");
      return "Orders sheet not found";
    }
    
    // Check if we already have headers
    if (sheet.getLastRow() === 0) {
      // Add headers
      const headers = [
        'Timestamp', 'Order Number', 'Product Name', 'Quantity', 
        'Price', 'Payment Method', 'Total', 'Category', 'Outlet', 'Username'
      ];
      sheet.appendRow(headers);
      Logger.log("Added headers");
    }
    
    // Add a few test orders for today
    const today = new Date();
    const testOrders = [
      {
        orderNumber: 'TEST-001',
        items: [
          { name: 'Coffee', qty: 2, price: 15000, category: 'Beverages' },
          { name: 'Sandwich', qty: 1, price: 25000, category: 'Food' }
        ],
        payment: 'Cash'
      },
      {
        orderNumber: 'TEST-002', 
        items: [
          { name: 'Tea', qty: 1, price: 10000, category: 'Beverages' }
        ],
        payment: 'Dana'
      }
    ];
    
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const outlet = session.outlet || 'Test Outlet';
    const username = session.username || 'Test User';
    
    testOrders.forEach(order => {
      order.items.forEach(item => {
        const itemTotal = item.price * item.qty;
        const row = [
          today,
          order.orderNumber,
          item.name,
          item.qty,
          item.price,
          order.payment,
          itemTotal,
          item.category,
          outlet,
          username
        ];
        sheet.appendRow(row);
        Logger.log("Added test row:", JSON.stringify(row));
      });
    });
    
    Logger.log("=== TEST DATA ADDED ===");
    return "Test data added successfully";
    
  } catch (error) {
    Logger.log("Error adding test data:", error.toString());
    return "Error: " + error.toString();
  }
}

// Function to add Product Code header to existing Orders sheet
function addProductCodeHeader() {
  try {
    Logger.log("=== ADDING PRODUCT CODE HEADER ===");
    
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (!sheet) {
      Logger.log("ERROR: Orders sheet not found");
      return "Orders sheet not found";
    }
    
    const lastCol = sheet.getLastColumn();
    Logger.log("Current last column: " + lastCol);
    
    // Check if we already have the Product Code header
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log("Current headers: " + JSON.stringify(headers));
    
    const productCodeExists = headers.some(header => 
      header && header.toString().toLowerCase().includes('product') && 
      header.toString().toLowerCase().includes('code')
    );
    
    if (productCodeExists) {
      Logger.log("Product Code header already exists");
      return "Product Code header already exists";
    }
    
    // Add Product Code header in the next column
    const newCol = lastCol + 1;
    sheet.getRange(1, newCol).setValue("Product Code");
    Logger.log("Added Product Code header in column " + newCol);
    
    // Update existing rows with product codes based on category + name
    const dataRows = sheet.getLastRow() - 1; // excluding header
    if (dataRows > 0) {
      Logger.log("Updating " + dataRows + " existing rows with product codes");
      
      // Find category and product name column indices
      const categoryIdx = headers.findIndex(h => h && /category/i.test(h.toString()));
      const nameIdx = headers.findIndex(h => h && /product|item|name/i.test(h.toString()));
      
      Logger.log("Category column: " + categoryIdx + ", Name column: " + nameIdx);
      
      if (categoryIdx >= 0 && nameIdx >= 0) {
        for (let row = 2; row <= sheet.getLastRow(); row++) {
          const category = sheet.getRange(row, categoryIdx + 1).getValue() || 'OTHER';
          const productName = sheet.getRange(row, nameIdx + 1).getValue() || 'UNKNOWN';
          const productCode = (category + '_' + productName).replace(/\s+/g, '_').toUpperCase();
          
          sheet.getRange(row, newCol).setValue(productCode);
          
          if (row <= 5) { // Log first few rows
            Logger.log("Row " + row + ": " + category + " + " + productName + " = " + productCode);
          }
        }
      }
    }
    
    Logger.log("=== PRODUCT CODE HEADER ADDED SUCCESSFULLY ===");
    return "Product Code header added successfully";
    
  } catch (error) {
    Logger.log("Error adding product code header: " + error.toString());
    return "Error: " + error.toString();
  }
}



// Test function to populate product codes in Products sheet
function testPopulateProductCodes() {
  Logger.log("=== Testing Product Code Population ===");
  
  try {
    var result = populateProductCodes();
    Logger.log("Product code population result: " + JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log("Error in product code population test: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Populate product codes in Products sheet using category and name
 * This function adds a Product Code column if it doesn't exist and populates it
 * @returns {Object} - Result with success status and update count
 */
function populateProductCodes() {
  try {
    Logger.log("=== POPULATING PRODUCT CODES IN PRODUCTS SHEET ===");
    
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) {
      throw new Error('Products sheet not found');
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      Logger.log("No data rows found in Products sheet");
      return { success: true, updatedCount: 0, message: "No data rows to update" };
    }
    
    // Get headers
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log("Current headers: " + JSON.stringify(headers));
    
    // Check if Product Code column exists
    let productCodeColIdx = headers.findIndex(h => h && /product.*code|code/i.test(h.toString()));
    
    if (productCodeColIdx === -1) {
      // Add Product Code header
      const newCol = lastCol + 1;
      sheet.getRange(1, newCol).setValue("Product Code");
      productCodeColIdx = newCol - 1; // Convert to 0-based index
      Logger.log("Added Product Code header in column " + newCol);
    } else {
      Logger.log("Product Code column already exists at index " + productCodeColIdx);
    }
    
    // Find necessary columns
    const nameIdx = headers.findIndex(h => h && /name/i.test(h.toString()));
    const categoryIdx = headers.findIndex(h => h && /category/i.test(h.toString()));
    
    Logger.log("Column indices - Name: " + nameIdx + ", Category: " + categoryIdx + ", Product Code: " + productCodeColIdx);
    
    if (nameIdx === -1) {
      throw new Error("Name column not found in Products sheet");
    }
    
    let updatedCount = 0;
    
    // Process each product row
    for (let row = 2; row <= lastRow; row++) {
      const name = sheet.getRange(row, nameIdx + 1).getValue() || '';
      const category = categoryIdx >= 0 ? sheet.getRange(row, categoryIdx + 1).getValue() || '' : '';
      
      if (!name.toString().trim()) {
        Logger.log("Skipping row " + row + " - no product name");
        continue;
      }
      
      // Generate product code using the same logic as orders
      const productCode = generateProductCode({
        name: name.toString().trim(),
        category: category.toString().trim()
      });
      
      // Update the product code column
      sheet.getRange(row, productCodeColIdx + 1).setValue(productCode);
      updatedCount++;
      
      if (row <= 5) { // Log first few rows
        Logger.log("Row " + row + ": " + category + " + " + name + " = " + productCode);
      }
    }
    
    Logger.log("=== PRODUCT CODES POPULATED SUCCESSFULLY ===");
    Logger.log("Updated " + updatedCount + " product rows");
    
    return { 
      success: true, 
      updatedCount: updatedCount, 
      message: "Product codes populated successfully" 
    };
    
  } catch (error) {
    Logger.log("Error populating product codes: " + error.toString());
    return { 
      success: false, 
      error: error.toString(),
      message: "Failed to populate product codes" 
    };
  }
}

// Test function to verify product sorting works correctly for all categories
function testProductSorting() {
  // Create sample data that mimics the actual sales summary structure
  const testData = {
    summary: {
      "Ayam": {
        "AYAM_AYAM_HOT_WINGS": { qty: 5, sales: 50000 },
        "AYAM_AYAM_ORI_CRISPY": { qty: 8, sales: 80000 },
        "AYAM_AYAM_HOT_SPICY": { qty: 3, sales: 30000 },
        "AYAM_GRILLED": { qty: 12, sales: 120000 },
        "AYAM_AYAM_ORI_CLASSIC": { qty: 6, sales: 60000 },
        "AYAM_FRIED": { qty: 15, sales: 150000 }
      }
    }
  };
  
  console.log("=== Testing Product Sorting Logic (All Categories) ===");
  console.log("Original items:", Object.keys(testData.summary.Ayam));
  
  // Apply the same sorting logic as used in the UI
  const category = "Ayam";
  
  // Sort all categories by product name alphabetically (same logic for all categories now)
  const itemsArray = Object.entries(testData.summary[category])
    .sort((a, b) => {
      // Get the display names (what will actually be shown in the report)
      const nameA = (a[1].itemName || a[0].split('_').slice(1).join('_').replace(/_/g, ' ')).toLowerCase();
      const nameB = (b[1].itemName || b[0].split('_').slice(1).join('_').replace(/_/g, ' ')).toLowerCase();
      
      // Simple alphabetical sort by name
      return nameA.localeCompare(nameB);
    });
  
  console.log("=== Sorted Results ===");
  itemsArray.forEach(([key, data], index) => {
    const productName = key.includes('_') ? 
      key.split('_').slice(1).join('_').replace(/_/g, ' ') : 
      key;
    console.log(`${index + 1}. ${productName} (${key}) - Qty: ${data.qty}, Sales: ${data.sales}`);
  });
  
  console.log("=== Expected Order (All Items Alphabetically by Name) ===");
  console.log("1. AYAM HOT SPICY");
  console.log("2. AYAM HOT WINGS"); 
  console.log("3. AYAM ORI CLASSIC");
  console.log("4. AYAM ORI CRISPY");
  console.log("5. FRIED");
  console.log("6. GRILLED");
  
  return itemsArray;
}

// Backward compatibility alias
function testAyamSorting() {
  return testProductSorting();
}

// Debug function to analyze categories and products in Orders sheet
function debugCategoriesAndProducts(deviceId, username) {
  try {
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (!sheet) {
      Logger.log("ERROR: Orders sheet not found");
      return "ERROR: Orders sheet not found";
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("No data in Orders sheet");
      return "No data in Orders sheet";
    }

    const headers = data[0];
    const categoryIdx = headers.findIndex(h => /category/i.test(h));
    const productIdx = headers.findIndex(h => /product|item|name/i.test(h));
    const dateIdx = headers.findIndex(h => /timestamp|date|tanggal|waktu|time/i.test(h));
    const outletIdx = headers.findIndex(h => /outlet/i.test(h));

    Logger.log("Column indices - Category: " + categoryIdx + ", Product: " + productIdx + ", Date: " + dateIdx + ", Outlet: " + outletIdx);

    const categories = new Set();
    const products = new Set();
    const categoryProducts = {};
    let ayamSisaCount = 0;
    let totalOrders = 0;

    // Get current user's outlet
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;

    data.slice(1).forEach((row, index) => {
      // Filter by outlet if specified
      if (userOutlet && outletIdx >= 0 && row[outletIdx] !== userOutlet) return;

      totalOrders++;
      const category = row[categoryIdx] || 'Unknown';
      const product = row[productIdx] || 'Unknown';
      const date = row[dateIdx];

      categories.add(category);
      products.add(product);

      if (!categoryProducts[category]) {
        categoryProducts[category] = new Set();
      }
      categoryProducts[category].add(product);

      // Check for AYAM_SISA variations
      const categoryStr = category.toString().toUpperCase();
      const productStr = product.toString().toUpperCase();
      if (categoryStr.includes('AYAM') && categoryStr.includes('SISA')) {
        ayamSisaCount++;
        Logger.log("Found AYAM_SISA entry at row " + (index + 2) + ": Category=" + category + ", Product=" + product + ", Date=" + date);
      }
      if (productStr.includes('AYAM') && productStr.includes('SISA')) {
        Logger.log("Found AYAM SISA in product name at row " + (index + 2) + ": Category=" + category + ", Product=" + product + ", Date=" + date);
      }
      
      // Show detailed outlet comparison for debugging
      if (index < 5) { // First 5 rows for today
        const rowOutletStr = outletIdx >= 0 ? String(row[outletIdx]).trim() : 'N/A';
        const userOutletStr = String(userOutlet).trim();
        Logger.log("Row " + (index + 2) + " outlet debug:");
        Logger.log("  - Row outlet: '" + rowOutletStr + "' (length: " + rowOutletStr.length + ")");
        Logger.log("  - User outlet: '" + userOutletStr + "' (length: " + userOutletStr.length + ")");
        Logger.log("  - Match: " + (rowOutletStr === userOutletStr));
        Logger.log("  - Row data: " + JSON.stringify(row.slice(0, 5)));
      }
    });

    const result = {
      totalOrders: totalOrders,
      userOutlet: userOutlet,
      categoriesCount: categories.size,
      productsCount: products.size,
      ayamSisaCount: ayamSisaCount,
      categories: Array.from(categories).sort(),
      categoryProducts: {}
    };

    // Convert Set to Array for each category
    Object.keys(categoryProducts).forEach(cat => {
      result.categoryProducts[cat] = Array.from(categoryProducts[cat]).sort();
    });

    Logger.log("Debug results: " + JSON.stringify(result, null, 2));
    return JSON.stringify(result, null, 2);

  } catch (error) {
    Logger.log("Error in debugCategoriesAndProducts: " + error.toString());
    return "Error: " + error.toString();
  }
}

// Test function for Monthly Sales Report
function testMonthlySalesReportIntegration() {
  try {
    Logger.log("=== Testing Monthly Sales Report Integration ===");
    
    const currentDate = new Date();
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth() + 1;
    
    // Test the new monthly sales report functions
    const monthlySalesResult = getMonthlySalesReport(year, month);
    Logger.log("Monthly Sales Report result:", monthlySalesResult);
    
    const summaryResult = getMonthlySalesReportSummary(year, month, true);
    Logger.log("Monthly Sales Report Summary result:", summaryResult);
    
    const availableMonthsResult = getAvailableMonths();
    Logger.log("Available Months result:", availableMonthsResult);
    
    return "Monthly Sales Report integration test completed - check logs";
    
  } catch (error) {
    Logger.log("Error in monthly sales report integration test:", error.toString());
    return "Error: " + error.toString();
  }
}

// ============================
// Sales analytics & period reports (Snowbee A1.2)
// ============================

function _resolveSalesAuthFromParams(params) {
  var deviceId = null;
  var username = null;
  if (params && params._authContext) {
    deviceId = params._authContext.deviceId;
    username = params._authContext.username;
  }
  var session = getUserSessionEnhanced(deviceId, username) || {};
  return { deviceId: deviceId, username: username, outlet: session.outlet };
}

function _loadFilteredOrderMetrics(deviceId, username) {
  var auth = _resolveSalesAuthFromParams({ _authContext: { deviceId: deviceId, username: username } });
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { byDate: {}, productQty: {} };
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var outletIdx = headers.findIndex(function(h) { return h && /outlet/i.test(h); });
  var dateIdx = headers.findIndex(function(h) { return h && /timestamp|date|tanggal|waktu|time/i.test(h); });
  var orderIdx = headers.findIndex(function(h) { return h && /order.*number|ordernumber/i.test(h); });
  var productIdx = headers.findIndex(function(h) { return h && /product|item|name/i.test(h); });
  var qtyIdx = headers.findIndex(function(h) { return h && /qty|quantity/i.test(h); });
  var totalIdx = headers.findIndex(function(h) { return h && /total|amount/i.test(h); });

  var byDate = {};
  var productQty = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (outletIdx !== -1 && auth.outlet && row[outletIdx] !== auth.outlet) continue;
    var d = row[dateIdx];
    if (!d) continue;
    var dateObj = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dateObj.getTime())) continue;
    var dayKey = formatDateForComparison(dateObj);
    var total = Number(row[totalIdx]) || 0;
    var orderNum = orderIdx >= 0 ? String(row[orderIdx] || '') : ('row-' + i);
    var product = productIdx >= 0 ? String(row[productIdx] || '') : '';
    var qty = Number(row[qtyIdx]) || 0;

    if (!byDate[dayKey]) byDate[dayKey] = { totalSales: 0, orders: {} };
    byDate[dayKey].totalSales += total;
    byDate[dayKey].orders[orderNum] = true;

    if (product) {
      if (!productQty[product]) productQty[product] = 0;
      productQty[product] += qty;
    }
  }

  var byDateSummary = {};
  Object.keys(byDate).forEach(function(k) {
    byDateSummary[k] = {
      totalSales: byDate[k].totalSales,
      ordersCount: Object.keys(byDate[k].orders).length
    };
  });

  return { byDate: byDateSummary, productQty: productQty };
}

function _pctChange(current, previous) {
  if (!previous) return current ? 100 : 0;
  return Math.round(((current - previous) / previous) * 1000) / 10;
}

function _sumRange(byDate, startDate, endDate) {
  var total = 0;
  var d = new Date(startDate);
  while (d <= endDate) {
    var key = formatDateForComparison(d);
    if (byDate[key]) total += byDate[key].totalSales;
    d.setDate(d.getDate() + 1);
  }
  return total;
}

function getSalesAnalytics(params) {
  try {
    var auth = _resolveSalesAuthFromParams(params || {});
    var metrics = _loadFilteredOrderMetrics(auth.deviceId, auth.username);
    var byDate = metrics.byDate;
    var today = new Date();
    var todayKey = formatDateForComparison(today);
    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    var yesterdayKey = formatDateForComparison(yesterday);

    var weekStart = new Date(today);
    weekStart.setDate(weekStart.getDate() - 6);
    var prevWeekEnd = new Date(weekStart);
    prevWeekEnd.setDate(prevWeekEnd.getDate() - 1);
    var prevWeekStart = new Date(prevWeekEnd);
    prevWeekStart.setDate(prevWeekStart.getDate() - 6);

    var monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
    var prevMonthEnd = new Date(monthStart);
    prevMonthEnd.setDate(prevMonthEnd.getDate() - 1);
    var prevMonthStart = new Date(prevMonthEnd.getFullYear(), prevMonthEnd.getMonth(), 1);

    var todaySales = (byDate[todayKey] && byDate[todayKey].totalSales) || 0;
    var yesterdaySales = (byDate[yesterdayKey] && byDate[yesterdayKey].totalSales) || 0;

    var topProduct = 'N/A';
    var topProductSales = 0;
    Object.keys(metrics.productQty).forEach(function(p) {
      if (metrics.productQty[p] > topProductSales) {
        topProduct = p;
        topProductSales = metrics.productQty[p];
      }
    });

    return {
      todaySales: todaySales,
      todayChange: _pctChange(todaySales, yesterdaySales),
      weekSales: _sumRange(byDate, weekStart, today),
      weekChange: _pctChange(_sumRange(byDate, weekStart, today), _sumRange(byDate, prevWeekStart, prevWeekEnd)),
      monthSales: _sumRange(byDate, monthStart, today),
      monthChange: _pctChange(_sumRange(byDate, monthStart, today), _sumRange(byDate, prevMonthStart, prevMonthEnd)),
      topProduct: topProduct,
      topProductSales: topProductSales
    };
  } catch (e) {
    Logger.log('getSalesAnalytics error: ' + e.toString());
    return {
      todaySales: 0, todayChange: 0, weekSales: 0, weekChange: 0,
      monthSales: 0, monthChange: 0, topProduct: 'N/A', topProductSales: 0
    };
  }
}

function _buildPeriodReport(params, groupFn) {
  var auth = _resolveSalesAuthFromParams(params || {});
  var metrics = _loadFilteredOrderMetrics(auth.deviceId, auth.username);
  var groups = {};
  Object.keys(metrics.byDate).forEach(function(dayKey) {
    var gKey = groupFn(dayKey);
    if (!groups[gKey]) groups[gKey] = { date: gKey, totalSales: 0, ordersCount: 0 };
    groups[gKey].totalSales += metrics.byDate[dayKey].totalSales;
    groups[gKey].ordersCount += metrics.byDate[dayKey].ordersCount;
  });
  var rows = Object.keys(groups).map(function(k) { return groups[k]; });
  rows.sort(function(a, b) { return b.date.localeCompare(a.date); });
  return JSON.stringify(rows);
}

function getSalesDailyReport(params) {
  try {
    var auth = _resolveSalesAuthFromParams(params || {});
    var metrics = _loadFilteredOrderMetrics(auth.deviceId, auth.username);
    var rows = Object.keys(metrics.byDate).map(function(k) {
      return { date: k, totalSales: metrics.byDate[k].totalSales, ordersCount: metrics.byDate[k].ordersCount };
    });
    rows.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return JSON.stringify(rows.slice(0, 30));
  } catch (e) {
    Logger.log('getSalesDailyReport error: ' + e.toString());
    return JSON.stringify([]);
  }
}

function getSalesWeeklyReport(params) {
  return _buildPeriodReport(params, function(dayKey) {
    var p = dayKey.split('-');
    var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    var day = d.getDay();
    var diff = d.getDate() - day + (day === 0 ? -6 : 1);
    d.setDate(diff);
    return formatDateForComparison(d);
  });
}

function getSalesMonthlyReport(params) {
  return _buildPeriodReport(params, function(dayKey) {
    return dayKey.substring(0, 7);
  });
}
