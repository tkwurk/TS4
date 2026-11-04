function getCashSummary(deviceId, username) {
  try {
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;
    
    Logger.log("getCashSummary - Session: " + JSON.stringify(session));
    Logger.log("getCashSummary - User outlet: " + userOutlet);

    // Filtered data for the user's outlet
    const salesData = getSheetData(SHEETS.ORDERS).filter((row, idx, arr) => {
      if (idx === 0) return true; // keep header
      const outletIdx = arr[0].findIndex(h => /outlet/i.test(h));
      return outletIdx === -1 || row[outletIdx] === userOutlet;
    });
    
    const depositData = getSheetData(SHEETS.DEPOSITS).filter((row, idx, arr) => {
      if (idx === 0) return true;
      const outletIdx = arr[0].findIndex(h => /outlet/i.test(h));
      return outletIdx === -1 || row[outletIdx] === userOutlet;
    });
    const expenseData = getSheetData(SHEETS.EXPENSES).filter((row, idx, arr) => {
      if (idx === 0) return true;
      const outletIdx = arr[0].findIndex(h => /outlet/i.test(h));
   
      return outletIdx === -1 || row[outletIdx] === userOutlet;
    });
    
    const purchaseData = getSheetData(SHEETS.PURCHASES).filter((row, idx, arr) => {
      if (idx === 0) return true;
      const outletIdx = arr[0].findIndex(h => /outlet/i.test(h));
      return outletIdx === -1 || row[outletIdx] === userOutlet;
    });


    const sales = analyzeSales(salesData);
    const deposits = analyzeDeposits(depositData);
    const expenses = analyzeExpenses(expenseData);
    const purchases = analyzePurchases(purchaseData);
    const salesPaymentBreakdown = typeof computeSalesPaymentBreakdownFromSheetData_ === 'function'
      ? computeSalesPaymentBreakdownFromSheetData_(salesData, depositData)
      : {
          salesByPaymentMethod: sales.byPaymentMethod,
          totalSalesByCash: sales.totalCash,
          totalSalesByNonCash: sales.totalNonCash
        };
    const cashInHand = sales.total - expenses.totalDeductible - purchases.totalCash - deposits.total;

    // Current month cash in hand
    const _now = new Date();
    const _curYear = _now.getFullYear();
    const _curMonth = _now.getMonth();
    function filterCurrentMonth(data) {
      if (data.length <= 1) return data;
      const h = data[0];
      const di = h.findIndex(function(col) { return col && /timestamp|date/i.test(col); });
      if (di === -1) return data;
      return [h].concat(data.slice(1).filter(function(row) {
        var d = (row[di] instanceof Date) ? row[di] : new Date(row[di]);
        return d && !isNaN(d.getTime()) && d.getFullYear() === _curYear && d.getMonth() === _curMonth;
      }));
    }
    const curMonthSales = analyzeSales(filterCurrentMonth(salesData));
    const curMonthDeposits = analyzeDeposits(filterCurrentMonth(depositData));
    const curMonthExpenses = analyzeExpenses(filterCurrentMonth(expenseData));
    const curMonthPurchases = analyzePurchases(filterCurrentMonth(purchaseData));
    const currentMonthCashInHand = curMonthSales.total - curMonthExpenses.totalDeductible - curMonthPurchases.totalCash - curMonthDeposits.total;

    Logger.log(`Cash In Hand calculation: ${sales.total} - ${expenses.totalCash} - ${purchases.totalCash} - ${deposits.total} = ${cashInHand}`);
    Logger.log(`Sales: ${JSON.stringify(sales)}`);
    Logger.log(`Deposits: ${JSON.stringify(deposits)}`);
    Logger.log(`Expenses: ${JSON.stringify(expenses)}`);
    Logger.log(`Purchases: ${JSON.stringify(purchases)}`);

Logger.log(expenses)

    return JSON.stringify({
      totalSales: sales.total,
      totalSalesByCash: salesPaymentBreakdown.totalSalesByCash,
      totalSalesByNonCash: salesPaymentBreakdown.totalSalesByNonCash,
      salesByPaymentMethod: salesPaymentBreakdown.salesByPaymentMethod,
      totalDeposits: deposits.total,
      totalDepositsQris: deposits.totalQris,
      totalDepositsTransfer: deposits.totalTransfer,
      depositsByPaymentMethod: deposits.byPaymentMethod,
      totalCashExpenses: expenses.totalCash,
      totalDeductibleExpenses: expenses.totalDeductible,
      expensesByPaymentMethod: expenses.byPaymentMethod,
      totalCashPurchases: purchases.totalCash,
      purchasesByPaymentMethod: purchases.byPaymentMethod,
      cashInHand: cashInHand,
      currentMonthCashInHand: currentMonthCashInHand,
      deposits: deposits.records,
      cashExpenses: expenses.cashRecords,
      allExpenses: expenses.allDeductibleRecords,
      cashPurchases: purchases.cashRecords,
      lastUpdated: formatDateForJSON(new Date())
    });
  } catch (error) {
    console.error("Error getting cash summary:", error);
    throw new Error("Failed to get cash summary");
  }
}

function analyzeSales(data) {
  if (data.length <= 1) return { total: 0, totalCash: 0, totalNonCash: 0, byPaymentMethod: {} };
  
  const headers = data[0];
  const totalIdx = headers.findIndex(h => /total|amount/i.test(h));
  const paymentIdx = headers.findIndex(h => /payment/i.test(h));
  
  let total = 0, totalCash = 0, totalNonCash = 0;
  const byPaymentMethod = {};
  const CASH_KEYS = ['cash', 'tunai'];
  const SKIP_KEYS = ['none', ''];
  
  data.slice(1).forEach(row => {
    const amount = Number(row[totalIdx]) || 0;
    const method = paymentIdx > -1 ? (row[paymentIdx] || '').toString().trim() : 'cash';
    const methodLower = method.toLowerCase();
    
    total += amount;
    
    if (CASH_KEYS.includes(methodLower) || isPartialPaymentLabel_(method)) {
      totalCash += amount;
      byPaymentMethod['Cash'] = (byPaymentMethod['Cash'] || 0) + amount;
    } else if (!SKIP_KEYS.includes(methodLower)) {
      totalNonCash += amount;
      if (method) {
        byPaymentMethod[method] = (byPaymentMethod[method] || 0) + amount;
      }
    }
  });

  Logger.log(`Sales analysis - Total: ${total}, Cash: ${totalCash}, NonCash: ${totalNonCash}`);
  return { total, totalCash, totalNonCash, byPaymentMethod };
}

function analyzeDeposits(data) {
  if (data.length <= 1) return { total: 0, totalQris: 0, totalTransfer: 0, byPaymentMethod: {}, records: [] };
  
  const headers = data[0];
  const dateIdx = headers.findIndex(h => /timestamp|date/i.test(h));
  const amountIdx = headers.findIndex(h => /amount|total/i.test(h));
  const accountIdx = headers.findIndex(h => /account/i.test(h));
  const noteIdx = headers.findIndex(h => /note/i.test(h));
  
  const paymentMethodIdx = headers.findIndex(h => /paymentmethod|payment.method/i.test(h));
  const configMap = typeof loadPaymentMethodConfigMap_ === 'function' ? loadPaymentMethodConfigMap_() : {};
  const records = [];
  let total = 0, totalQris = 0, totalTransfer = 0;
  const byPaymentMethod = {};
  
  data.slice(1).forEach(row => {
    const amount = Number(row[amountIdx]) || 0;
    if (amount > 0) {
      const account = accountIdx > -1 ? (row[accountIdx] || '').toString().toUpperCase().trim() : '';
      const payKey = paymentMethodIdx > -1 ? (row[paymentMethodIdx] || '').toString().trim() : '';
      const depMethod = payKey || account;
      const isAutoDeposit = typeof paymentMethodHasAutoDeposit_ === 'function'
        ? paymentMethodHasAutoDeposit_(depMethod, configMap)
        : (payKey ? ['qris', 'dana'].some(function(k) { return k === payKey.toLowerCase(); }) : account === 'DANA');
      const methodKey = payKey || (isAutoDeposit ? 'qris' : 'transfer');
      total += amount;
      if (isAutoDeposit) totalQris += amount; else totalTransfer += amount;
      byPaymentMethod[methodKey] = (byPaymentMethod[methodKey] || 0) + amount;
      records.push({
        timestamp: formatDateForJSON(row[dateIdx]),
        amount: amount,
        note: row[noteIdx] || ""
      });
    }
  });

  Logger.log(`Deposits analysis - Total: ${total}, QRIS: ${totalQris}, Transfer: ${totalTransfer}, Records: ${records.length}`);
  return { total, totalQris, totalTransfer, byPaymentMethod, records };
}

function analyzeExpenses(data) {
  if (data.length <= 1) return { totalCash: 0, cashRecords: [], totalOthers: 0 };
  
  const headers = data[0];
  const dateIdx = headers.findIndex(h => /timestamp|date/i.test(h));
  const amountIdx = headers.findIndex(h => /amount|total/i.test(h));
  const accountIdx = headers.findIndex(h => /account/i.test(h));
  const paymentIdx = headers.findIndex(h => /payment/i.test(h));
  const noteIdx = headers.findIndex(h => /note/i.test(h));
  
  Logger.log(`Expenses column indices - Date: ${dateIdx}, Amount: ${amountIdx}, Account: ${accountIdx}, Payment: ${paymentIdx}, Note: ${noteIdx}`);
  
  const EXPENSE_DEDUCTIBLE_KEYS = ['cash', 'partial', 'qris', 'dana', 'free'];
  const cashRecords = [];
  const allDeductibleRecords = [];
  let totalCash = 0;
  let totalDeductible = 0;
  let totalOthers = 0;
  const byPaymentMethod = {};
  
  data.slice(1).forEach(row => {
    const amount = Number(row[amountIdx]) || 0;
    const methodRaw = paymentIdx > -1 ? (row[paymentIdx] || '').toString().trim() : 'cash';
    const source = methodRaw.toLowerCase();
    const methodKey = methodRaw || 'cash';
    
    Logger.log(`Expense row - Amount: ${amount}, Payment: ${source}`);
    
    byPaymentMethod[methodKey] = (byPaymentMethod[methodKey] || 0) + amount;
    if (EXPENSE_DEDUCTIBLE_KEYS.includes(source)) {
      totalDeductible += amount;
      allDeductibleRecords.push({
        timestamp: formatDateForJSON(row[dateIdx]),
        amount: amount,
        account: row[accountIdx] || "",
        note: row[noteIdx] || "",
        payment: methodKey
      });
    }
    if (source === 'cash') {
      totalCash += amount;
      cashRecords.push({
        timestamp: formatDateForJSON(row[dateIdx]),
        amount: amount,
        account: row[accountIdx] || "",
        note: row[noteIdx] || ""
      });
    } else {
      totalOthers += amount;
    }
  });

  Logger.log(`Expenses analysis - Cash: ${totalCash}, Deductible: ${totalDeductible}, Others: ${totalOthers}, Cash Records: ${cashRecords.length}, All Deductible Records: ${allDeductibleRecords.length}`);
  return { totalCash, totalDeductible, cashRecords, allDeductibleRecords, totalOthers, byPaymentMethod };
}

function analyzePurchases(data) {
  if (data.length <= 1) return { totalCash: 0, cashRecords: [], totalOthers: 0 };
  
  const headers = data[0];
  const dateIdx = headers.findIndex(h => /timestamp|date/i.test(h));
  const amountIdx = headers.findIndex(h => /total|amount/i.test(h));
  const productIdx = headers.findIndex(h => /product|item/i.test(h));
  const paymentIdx = headers.findIndex(h => /_selectingPayment|payment/i.test(h));
  const supplierIdx = headers.findIndex(h => /supplier/i.test(h));
  const invIdx = purchaseInvestmentColIndex_(headers);
  
  Logger.log(`Purchases column indices - Date: ${dateIdx}, Amount: ${amountIdx}, Product: ${productIdx}, Payment: ${paymentIdx}, Supplier: ${supplierIdx}`);
  
  const cashRecords = [];
  let totalCash = 0;
  let totalOthers = 0;
  const byPaymentMethod = {};
  
  data.slice(1).forEach(row => {
    if (isInvestmentPurchaseRow_(row, invIdx)) return; // investment spend not part of cash flow
    const amount = Number(row[amountIdx]) || 0;
    const methodRaw = paymentIdx > -1 ? (row[paymentIdx] || '').toString().trim() : 'cash';
    const payment = methodRaw.toLowerCase();
    const methodKey = methodRaw || 'cash';
    
    Logger.log(`Purchase row - Amount: ${amount}, Payment: ${payment}`);
    
    byPaymentMethod[methodKey] = (byPaymentMethod[methodKey] || 0) + amount;
    if (payment === 'cash') {
      totalCash += amount;
      // Escape special characters for JSON transfer
      let productName = row[productIdx] || "";
      try {
        productName = JSON.stringify(productName).slice(1, -1); // Ensures proper escaping
      } catch (e) {
        // fallback: replace problematic characters
        productName = String(productName).replace(/[\u0000-\u001F\u007F-\u009F"'\\]/g, '?');
      }
      cashRecords.push({
        timestamp: formatDateForJSON(row[dateIdx]),
        amount: amount,
        product: productName,
        supplier: row[supplierIdx] || ""
      });
    } else {
      totalOthers += amount;
    }
  });

  Logger.log(`Purchases analysis - Cash: ${totalCash}, Others: ${totalOthers}, Cash Records: ${cashRecords.length}`);
  return { totalCash, cashRecords, totalOthers, byPaymentMethod };
}

function getMonthlyCashSummary(deviceId, username) {
  try {
    var session = getUserSessionEnhanced(deviceId, username) || {};
    var userOutlet = session.outlet;
    var ss = getGlobalSpreadsheet();
    var months = {};

    function getMonth(val) {
      var d = (val instanceof Date) ? val : new Date(val);
      if (!d || isNaN(d.getTime())) return null;
      return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
    }
    function ensure(m) {
      if (!months[m]) months[m] = { totalSales: 0, cashSales: 0, cashExpenses: 0, deductibleExpenses: 0, cashPurchases: 0, deposits: 0, depositsQris: 0, depositsTransfer: 0, depositsByPaymentMethod: {}, salesByPaymentMethod: {} };
    }

    // Orders - cash sales
    var sheet = ss.getSheetByName(SHEETS.ORDERS);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var h = data[0];
        var di = h.findIndex(function(x) { return x && /timestamp|date/i.test(x); });
        var ti = h.findIndex(function(x) { return x && /total|amount/i.test(x); });
        var pi = h.findIndex(function(x) { return x && /payment/i.test(x); });
        var oi = h.findIndex(function(x) { return x && /outlet/i.test(x); });
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (oi !== -1 && userOutlet && row[oi] !== userOutlet) continue;
          var m = getMonth(row[di]);
          if (!m) continue;
          var pay = pi > -1 ? (row[pi] || '').toString().trim() : 'cash';
          var payLower = pay.toLowerCase();
          var amt = Number(row[ti]) || 0;
          ensure(m);
          months[m].totalSales += amt;
          if (payLower === 'cash') { months[m].cashSales += amt; }
          if (pay) {
            months[m].salesByPaymentMethod[pay] = (months[m].salesByPaymentMethod[pay] || 0) + amt;
          }
        }
      }
    }

    // Expenses - cash expenses
    sheet = ss.getSheetByName(SHEETS.EXPENSES);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var h = data[0];
        var di = h.findIndex(function(x) { return x && /timestamp|date/i.test(x); });
        var ai = h.findIndex(function(x) { return x && /amount|total/i.test(x); });
        var pi = h.findIndex(function(x) { return x && /payment/i.test(x); });
        var oi = h.findIndex(function(x) { return x && /outlet/i.test(x); });
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (oi !== -1 && userOutlet && row[oi] !== userOutlet) continue;
          var m = getMonth(row[di]);
          if (!m) continue;
          var pay = pi > -1 ? (row[pi] || '').toString().toLowerCase().trim() : 'cash';
          var expAmt = Number(row[ai]) || 0;
          if (pay === 'cash') { ensure(m); months[m].cashExpenses += expAmt; }
          if (['cash', 'partial', 'qris', 'dana', 'free'].includes(pay)) { ensure(m); months[m].deductibleExpenses += expAmt; }
        }
      }
    }

    // Purchases - cash purchases
    sheet = ss.getSheetByName(SHEETS.PURCHASES);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var h = data[0];
        var di = h.findIndex(function(x) { return x && /timestamp|date/i.test(x); });
        var ai = h.findIndex(function(x) { return x && /total|amount/i.test(x); });
        var pi = h.findIndex(function(x) { return x && /payment/i.test(x); });
        var oi = h.findIndex(function(x) { return x && /outlet/i.test(x); });
        var ii = purchaseInvestmentColIndex_(h);
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (isInvestmentPurchaseRow_(row, ii)) continue;
          if (oi !== -1 && userOutlet && row[oi] !== userOutlet) continue;
          var m = getMonth(row[di]);
          if (!m) continue;
          var pay = pi > -1 ? (row[pi] || '').toString().toLowerCase().trim() : 'cash';
          if (pay === 'cash') { ensure(m); months[m].cashPurchases += Number(row[ai]) || 0; }
        }
      }
    }

    // Deposits
    sheet = ss.getSheetByName(SHEETS.DEPOSITS);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var h = data[0];
        var di = h.findIndex(function(x) { return x && /timestamp|date/i.test(x); });
        var ai = h.findIndex(function(x) { return x && /amount|total/i.test(x); });
        var aci = h.findIndex(function(x) { return x && /account/i.test(x); });
        var oi = h.findIndex(function(x) { return x && /outlet/i.test(x); });
        var pmi = h.findIndex(function(x) { return x && /paymentmethod|payment.method/i.test(x); });
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (oi !== -1 && userOutlet && row[oi] !== userOutlet) continue;
          var m = getMonth(row[di]);
          if (!m) continue;
          var amt = Number(row[ai]) || 0;
          if (amt > 0) {
            ensure(m);
            months[m].deposits += amt;
            var acct = aci > -1 ? (row[aci] || '').toString().toUpperCase().trim() : '';
            var payKey = pmi > -1 ? (row[pmi] || '').toString().trim() : '';
            var isQris = payKey ? (payKey.toLowerCase() === 'qris' || payKey.toLowerCase() === 'dana') : acct === 'DANA';
            if (isQris) months[m].depositsQris += amt;
            else months[m].depositsTransfer += amt;
            var methodKey = payKey || (isQris ? 'qris' : 'transfer');
            months[m].depositsByPaymentMethod[methodKey] = (months[m].depositsByPaymentMethod[methodKey] || 0) + amt;
          }
        }
      }
    }

    // Build result sorted newest first
    var result = Object.keys(months).sort(function(a, b) { return b.localeCompare(a); }).map(function(key) {
      var d = months[key];
      var parts = key.split('-');
      return {
        month: key,
        displayName: getMonthName(parseInt(parts[1])) + ' ' + parts[0],
        totalSales: d.totalSales,
        cashSales: d.cashSales,
        cashExpenses: d.cashExpenses,
        deductibleExpenses: d.deductibleExpenses,
        cashPurchases: d.cashPurchases,
        deposits: d.deposits,
        depositsQris: d.depositsQris,
        depositsTransfer: d.depositsTransfer,
        depositsByPaymentMethod: d.depositsByPaymentMethod,
        salesByPaymentMethod: d.salesByPaymentMethod,
        cashInHand: 0
      };
    });

    // Cumulative cash in hand from oldest to newest
    // CIH = Previous CIH + Total Sales - Deductible Expenses (Cash/Partial/QRIS/Dana/Free) - Cash Purchases - Deposits
    var running = 0;
    for (var i = result.length - 1; i >= 0; i--) {
      var r = result[i];
      running += r.totalSales - r.deductibleExpenses - r.cashPurchases - r.deposits;
      r.cashInHand = running;
    }

    return result;
  } catch (error) {
    Logger.log('Error in getMonthlyCashSummary: ' + error.toString());
    return { error: error.toString() };
  }
}

function getDailyCashSummary(deviceId, username) {
  try {
    var session = getUserSessionEnhanced(deviceId, username) || {};
    var userOutlet = session.outlet;
    var ss = getGlobalSpreadsheet();
    var days = {};

    function getDay(val) {
      var d = (val instanceof Date) ? val : new Date(val);
      if (!d || isNaN(d.getTime())) return null;
      return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
    }
    function ensure(dk) {
      if (!days[dk]) days[dk] = { totalSales: 0, cashSales: 0, cashExpenses: 0, deductibleExpenses: 0, cashPurchases: 0, deposits: 0, depositsQris: 0, depositsTransfer: 0, depositsByPaymentMethod: {}, salesByPaymentMethod: {} };
    }

    // Orders
    var sheet = ss.getSheetByName(SHEETS.ORDERS);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var h = data[0];
        var di = h.findIndex(function(x) { return x && /timestamp|date/i.test(x); });
        var ti = h.findIndex(function(x) { return x && /total|amount/i.test(x); });
        var pi = h.findIndex(function(x) { return x && /payment/i.test(x); });
        var oi = h.findIndex(function(x) { return x && /outlet/i.test(x); });
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (oi !== -1 && userOutlet && row[oi] !== userOutlet) continue;
          var dk = getDay(row[di]);
          if (!dk) continue;
          var pay = pi > -1 ? (row[pi] || '').toString().trim() : 'cash';
          var payLower = pay.toLowerCase();
          var amt = Number(row[ti]) || 0;
          ensure(dk);
          days[dk].totalSales += amt;
          if (payLower === 'cash') { days[dk].cashSales += amt; }
          if (pay) {
            days[dk].salesByPaymentMethod[pay] = (days[dk].salesByPaymentMethod[pay] || 0) + amt;
          }
        }
      }
    }

    // Expenses
    sheet = ss.getSheetByName(SHEETS.EXPENSES);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var h = data[0];
        var di = h.findIndex(function(x) { return x && /timestamp|date/i.test(x); });
        var ai = h.findIndex(function(x) { return x && /amount|total/i.test(x); });
        var pi = h.findIndex(function(x) { return x && /payment/i.test(x); });
        var oi = h.findIndex(function(x) { return x && /outlet/i.test(x); });
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (oi !== -1 && userOutlet && row[oi] !== userOutlet) continue;
          var dk = getDay(row[di]);
          if (!dk) continue;
          var pay = pi > -1 ? (row[pi] || '').toString().toLowerCase().trim() : 'cash';
          var expAmt = Number(row[ai]) || 0;
          if (pay === 'cash') { ensure(dk); days[dk].cashExpenses += expAmt; }
          if (['cash', 'partial', 'qris', 'dana', 'free'].includes(pay)) { ensure(dk); days[dk].deductibleExpenses += expAmt; }
        }
      }
    }

    // Purchases
    sheet = ss.getSheetByName(SHEETS.PURCHASES);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var h = data[0];
        var di = h.findIndex(function(x) { return x && /timestamp|date/i.test(x); });
        var ai = h.findIndex(function(x) { return x && /total|amount/i.test(x); });
        var pi = h.findIndex(function(x) { return x && /payment/i.test(x); });
        var oi = h.findIndex(function(x) { return x && /outlet/i.test(x); });
        var ii = purchaseInvestmentColIndex_(h);
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (isInvestmentPurchaseRow_(row, ii)) continue;
          if (oi !== -1 && userOutlet && row[oi] !== userOutlet) continue;
          var dk = getDay(row[di]);
          if (!dk) continue;
          var pay = pi > -1 ? (row[pi] || '').toString().toLowerCase().trim() : 'cash';
          if (pay === 'cash') { ensure(dk); days[dk].cashPurchases += Number(row[ai]) || 0; }
        }
      }
    }

    // Deposits
    sheet = ss.getSheetByName(SHEETS.DEPOSITS);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var h = data[0];
        var di = h.findIndex(function(x) { return x && /timestamp|date/i.test(x); });
        var ai = h.findIndex(function(x) { return x && /amount|total/i.test(x); });
        var aci = h.findIndex(function(x) { return x && /account/i.test(x); });
        var oi = h.findIndex(function(x) { return x && /outlet/i.test(x); });
        var pmi = h.findIndex(function(x) { return x && /paymentmethod|payment.method/i.test(x); });
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (oi !== -1 && userOutlet && row[oi] !== userOutlet) continue;
          var dk = getDay(row[di]);
          if (!dk) continue;
          var amt = Number(row[ai]) || 0;
          if (amt > 0) {
            ensure(dk);
            days[dk].deposits += amt;
            var acct = aci > -1 ? (row[aci] || '').toString().toUpperCase().trim() : '';
            var payKey = pmi > -1 ? (row[pmi] || '').toString().trim() : '';
            var isQris = payKey ? (payKey.toLowerCase() === 'qris' || payKey.toLowerCase() === 'dana') : acct === 'DANA';
            if (isQris) days[dk].depositsQris += amt;
            else days[dk].depositsTransfer += amt;
            var methodKey = payKey || (isQris ? 'qris' : 'transfer');
            days[dk].depositsByPaymentMethod[methodKey] = (days[dk].depositsByPaymentMethod[methodKey] || 0) + amt;
          }
        }
      }
    }

    // Build result sorted newest first
    var DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    var MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var result = Object.keys(days).sort(function(a, b) { return b.localeCompare(a); }).map(function(key) {
      var d = days[key];
      var parts = key.split('-');
      var yr = parseInt(parts[0]), mo = parseInt(parts[1]), dy = parseInt(parts[2]);
      var dateObj = new Date(yr, mo - 1, dy);
      var displayDate = DAY_NAMES[dateObj.getDay()] + ', ' + ('0' + dy).slice(-2) + ' ' + MONTH_NAMES[mo - 1].slice(0, 3) + ' ' + yr;
      var monthKey = key.substring(0, 7);
      return {
        date: key,
        month: monthKey,
        displayDate: displayDate,
        monthDisplayName: MONTH_NAMES[mo - 1] + ' ' + yr,
        totalSales: d.totalSales,
        cashSales: d.cashSales,
        cashExpenses: d.cashExpenses,
        deductibleExpenses: d.deductibleExpenses,
        cashPurchases: d.cashPurchases,
        deposits: d.deposits,
        depositsQris: d.depositsQris,
        depositsTransfer: d.depositsTransfer,
        depositsByPaymentMethod: d.depositsByPaymentMethod,
        salesByPaymentMethod: d.salesByPaymentMethod,
        cashInHand: 0
      };
    });

    // Cumulative CIH from oldest to newest
    var running = 0;
    for (var i = result.length - 1; i >= 0; i--) {
      var r = result[i];
      running += r.totalSales - r.deductibleExpenses - r.cashPurchases - r.deposits;
      r.cashInHand = running;
    }

    return result;
  } catch (error) {
    Logger.log('Error in getDailyCashSummary: ' + error.toString());
    return { error: error.toString() };
  }
}