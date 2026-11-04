// Monthly Sales Report with Daily Breakdown and Product Details

/**
 * Get comprehensive monthly sales report with daily breakdown and product details
 * @param {number} year - The year (e.g., 2025)
 * @param {number} month - The month (1-12)
 * @param {string} deviceId - Device ID for session authentication
 * @param {string} username - Username for session authentication
 * @returns {string} JSON string containing the monthly sales report
 */
function getMonthlySalesReport(year, month, deviceId, username) {
  try {
    Logger.log(`getMonthlySalesReport - Year: ${year}, Month: ${month}, DeviceId: ${deviceId}, Username: ${username}`);
    
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;
    
    Logger.log(`Session info - UserOutlet: ${userOutlet}, Session: ${JSON.stringify(session)}`);
    
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (!sheet) {
      Logger.log("ERROR: Orders sheet not found");
      return JSON.stringify({
        success: false,
        error: "Orders sheet not found",
        data: [],
        summary: {}
      });
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("No data in Orders sheet");
      return JSON.stringify({
        success: true,
        data: [],
        summary: { totalDays: 0, totalSales: 0, totalQuantity: 0, totalOrders: 0 },
        month: `${year}-${String(month).padStart(2, '0')}`
      });
    }
    
    const headers = data[0];
    const outletIdx = headers.findIndex(h => /outlet/i.test(h));
    const dateIdx = headers.findIndex(h => /timestamp|date/i.test(h));
    const productIdx = headers.findIndex(h => /product|item|name/i.test(h));
    const qtyIdx = headers.findIndex(h => /qty|quantity/i.test(h));
    const priceIdx = headers.findIndex(h => /price/i.test(h));
    const totalIdx = headers.findIndex(h => /total|amount/i.test(h));
    const categoryIdx = headers.findIndex(h => /category/i.test(h));
    const orderNumberIdx = headers.findIndex(h => /order.*number/i.test(h));
    const paymentIdx = headers.findIndex(h => /payment/i.test(h));
    const productCodeIdx = headers.findIndex(h => /product.*code|code/i.test(h));
    
    Logger.log("Column indices - Date: " + dateIdx + ", Product: " + productIdx + ", Qty: " + qtyIdx + ", Price: " + priceIdx + ", Total: " + totalIdx + ", Outlet: " + outletIdx);
    Logger.log("Headers: " + JSON.stringify(headers));
    
    // Get month boundaries
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);
    endDate.setHours(23, 59, 59, 999); // Set to end of day

    Logger.log(`Date range: ${startDate.toISOString()} to ${endDate.toISOString()}`);

    // 1. Process Deposits first
    const dailyDeposits = new Map();
    const depositsSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.DEPOSITS);
    if (depositsSheet) {
      const depositsData = depositsSheet.getDataRange().getValues();
      if (depositsData.length > 1) {
        const depHeaders = depositsData[0];
        const depDateIdx = depHeaders.findIndex(h => /timestamp|date/i.test(h));
        const depAmountIdx = depHeaders.findIndex(h => /amount|total/i.test(h));
        const depOutletIdx = depHeaders.findIndex(h => /outlet/i.test(h));

        if (depDateIdx > -1 && depAmountIdx > -1) {
          Logger.log(`Found Deposits columns - Date: ${depDateIdx}, Amount: ${depAmountIdx}, Outlet: ${depOutletIdx}`);
          depositsData.slice(1).forEach((row, i) => {
            // Filter by outlet - include if no outlet column or outlet matches
            if (depOutletIdx !== -1 && row[depOutletIdx] !== userOutlet) return;
            try {
              const amount = Number(row[depAmountIdx]) || 0;
              if (amount === 0) return;

              const dateRaw = row[depDateIdx];
              if (!dateRaw) return;

              // --- Robust Date Parsing ---
              let depositDate;
              if (dateRaw instanceof Date) {
                depositDate = dateRaw;
              } else if (typeof dateRaw === 'string') {
                try {
                  depositDate = new Date(dateRaw);
                  if (isNaN(depositDate.getTime())) {
                     depositDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
                  }
                } catch(e) {
                   Logger.log(`Could not parse date string '${dateRaw}' for deposit row ${i+1}. Error: ${e}`);
                   return; // Skip this row
                }
              } else {
                 depositDate = new Date(dateRaw);
              }
              
              if (!depositDate || isNaN(depositDate.getTime())) {
                Logger.log(`Invalid or unparseable date for deposit row ${i+1}: ${dateRaw}`);
                return;
              }
              // --- End Robust Date Parsing ---

              if (depositDate >= startDate && depositDate <= endDate) {
                const dateKey = formatDateKey(depositDate);
                dailyDeposits.set(dateKey, (dailyDeposits.get(dateKey) || 0) + amount);
              }
            } catch (e) {
              Logger.log(`Skipping deposit row ${i+1} due to error: ${e}`)
            }
          });
        } else {
          Logger.log("Could not find required 'Date' and 'Amount'/'Total' columns in Deposits sheet.");
        }
      }
    } else {
      Logger.log("Deposits sheet with name '" + SHEETS.DEPOSITS + "' not found.");
    }
    Logger.log(`Processed ${dailyDeposits.size} days of deposits.`);

    // 2. Process Expenses - split by payment method
    const dailyCashExpenses = new Map();
    const dailyOtherExpenses = new Map();
    const expensesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.EXPENSES);
    if (expensesSheet) {
      const expensesData = expensesSheet.getDataRange().getValues();
      if (expensesData.length > 1) {
        const expHeaders = expensesData[0];
        const expDateIdx = expHeaders.findIndex(h => /timestamp|date/i.test(h));
        const expAmountIdx = expHeaders.findIndex(h => /amount|total/i.test(h));
        const expPaymentIdx = expHeaders.findIndex(h => /payment/i.test(h));
        const expOutletIdx = expHeaders.findIndex(h => /outlet/i.test(h));

        if (expDateIdx > -1 && expAmountIdx > -1) {
          Logger.log(`Found Expenses columns - Date: ${expDateIdx}, Amount: ${expAmountIdx}, Payment: ${expPaymentIdx}, Outlet: ${expOutletIdx}`);
          expensesData.slice(1).forEach((row, i) => {
            // Filter by outlet - include if no outlet column or outlet matches
            if (expOutletIdx !== -1 && row[expOutletIdx] !== userOutlet) return;
            try {
              const amount = Number(row[expAmountIdx]) || 0;
              if (amount === 0) return;

              const dateRaw = row[expDateIdx];
              if (!dateRaw) return;

              const paymentMethod = expPaymentIdx > -1 ? (row[expPaymentIdx] || '').toLowerCase().trim() : 'cash';

              // --- Robust Date Parsing ---
              let expenseDate;
              if (dateRaw instanceof Date) {
                expenseDate = dateRaw;
              } else if (typeof dateRaw === 'string') {
                try {
                  expenseDate = new Date(dateRaw);
                  if (isNaN(expenseDate.getTime())) {
                     // Fallback for different formats if needed, similar to sales
                     expenseDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
                  }
                } catch(e) {
                   Logger.log(`Could not parse date string '${dateRaw}' for expense row ${i+1}. Error: ${e}`);
                   return; // Skip this row
                }
              } else {
                 expenseDate = new Date(dateRaw);
              }
              
              if (!expenseDate || isNaN(expenseDate.getTime())) {
                Logger.log(`Invalid or unparseable date for expense row ${i+1}: ${dateRaw}`);
                return;
              }
              // --- End Robust Date Parsing ---

              if (expenseDate >= startDate && expenseDate <= endDate) {
                const dateKey = formatDateKey(expenseDate);
                
                // Split expenses by payment method
                if (['cash', 'partial', 'qris', 'dana'].includes(paymentMethod)) {
                  dailyCashExpenses.set(dateKey, (dailyCashExpenses.get(dateKey) || 0) + amount);
                } else {
                  // Transfer, etc. are considered "other" expenses (non-deductible from cash in hand)
                  dailyOtherExpenses.set(dateKey, (dailyOtherExpenses.get(dateKey) || 0) + amount);
                }
              }
            } catch (e) {
              Logger.log(`Skipping expense row ${i+1} due to error: ${e}`)
            }
          });
        } else {
          Logger.log("Could not find required 'Date' and 'Amount'/'Total' columns in Expenses sheet.");
        }
      }
    } else {
      Logger.log("Expenses sheet with name '" + SHEETS.EXPENSES + "' not found.");
    }
    Logger.log(`Processed ${dailyCashExpenses.size} days of cash expenses and ${dailyOtherExpenses.size} days of other expenses.`);

    // 3. Process Purchases - split by payment method
    const dailyCashPurchases = new Map();
    const dailyOtherPurchases = new Map();
    const purchasesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
    if (purchasesSheet) {
      const purchasesData = purchasesSheet.getDataRange().getValues();
      if (purchasesData.length > 1) {
        const purHeaders = purchasesData[0];
        const purDateIdx = purHeaders.findIndex(h => /timestamp|date/i.test(h));
        const purAmountIdx = purHeaders.findIndex(h => /amount|total/i.test(h));
        const purPaymentIdx = purHeaders.findIndex(h => /_selectingPayment|payment/i.test(h));
        const purOutletIdx = purHeaders.findIndex(h => /outlet/i.test(h));
        const purInvIdx = purchaseInvestmentColIndex_(purHeaders);

        if (purDateIdx > -1 && purAmountIdx > -1) {
          Logger.log(`Found Purchases columns - Date: ${purDateIdx}, Amount: ${purAmountIdx}, Payment: ${purPaymentIdx}, Outlet: ${purOutletIdx}`);
          purchasesData.slice(1).forEach((row, i) => {
            // Exclude investment (capital) purchases from cash flow
            if (isInvestmentPurchaseRow_(row, purInvIdx)) return;
            // Filter by outlet - include if no outlet column or outlet matches
            if (purOutletIdx !== -1 && row[purOutletIdx] !== userOutlet) return;
            try {
              const amount = Number(row[purAmountIdx]) || 0;
              if (amount === 0) return;

              const dateRaw = row[purDateIdx];
              if (!dateRaw) return;

              const paymentMethod = purPaymentIdx > -1 ? (row[purPaymentIdx] || '').toLowerCase().trim() : 'cash';

              // --- Robust Date Parsing ---
              let purchaseDate;
              if (dateRaw instanceof Date) {
                purchaseDate = dateRaw;
              } else if (typeof dateRaw === 'string') {
                try {
                  purchaseDate = new Date(dateRaw);
                  if (isNaN(purchaseDate.getTime())) {
                     purchaseDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
                  }
                } catch(e) {
                   Logger.log(`Could not parse date string '${dateRaw}' for purchase row ${i+1}. Error: ${e}`);
                   return; // Skip this row
                }
              } else {
                 purchaseDate = new Date(dateRaw);
              }
              
              if (!purchaseDate || isNaN(purchaseDate.getTime())) {
                Logger.log(`Invalid or unparseable date for purchase row ${i+1}: ${dateRaw}`);
                return;
              }
              // --- End Robust Date Parsing ---

              if (purchaseDate >= startDate && purchaseDate <= endDate) {
                const dateKey = formatDateKey(purchaseDate);
                
                // Split purchases by payment method
                if (paymentMethod === 'cash') {
                  dailyCashPurchases.set(dateKey, (dailyCashPurchases.get(dateKey) || 0) + amount);
                } else {
                  // Dana, transfer, etc. are considered "other" purchases
                  dailyOtherPurchases.set(dateKey, (dailyOtherPurchases.get(dateKey) || 0) + amount);
                }
              }
            } catch (e) {
              Logger.log(`Skipping purchase row ${i+1} due to error: ${e}`)
            }
          });
        } else {
          Logger.log("Could not find required 'Date' and 'Amount'/'Total' columns in Purchases sheet.");
        }
      }
    } else {
      Logger.log("Purchases sheet with name '" + SHEETS.PURCHASES + "' not found.");
    }
    Logger.log(`Processed ${dailyCashPurchases.size} days of cash purchases and ${dailyOtherPurchases.size} days of other purchases.`);

    // 4. Process Sales and combine with Expenses, Purchases and Deposits
    const dailyData = {};
    const monthSummary = {
      totalSales: 0,
      cashSales: 0,
      otherSales: 0,
      totalQuantity: 0,
      totalOrders: 0,
      uniqueOrderNumbers: new Set(),
      totalDays: 0,
      operationalExpenses: 0,
      cashExpenses: 0,
      otherExpenses: 0,
      cashPurchases: 0,
      otherPurchases: 0,
      totalDeposits: 0,
      cashInHand: 0
    };
    
    // Process each row
    let totalRowsProcessed = 0;
    let rowsFilteredByOutlet = 0;
    let rowsFilteredByDate = 0;
    let rowsWithValidData = 0;
    
    data.slice(1).forEach((row, index) => {
      totalRowsProcessed++;
      
      // Filter by outlet - include if no outlet column or outlet matches
      if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) {
        rowsFilteredByOutlet++;
        return;
      }
      
      // Parse and validate date
      let rowDate;
      try {
        const dateRaw = row[dateIdx];
        
        if (dateRaw instanceof Date) {
          rowDate = dateRaw;
        } else if (typeof dateRaw === 'string') {
          // Try Utilities.parseDate first for Google Sheets compatibility
          try {
            rowDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!rowDate || isNaN(rowDate.getTime())) {
              rowDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            }
            if (!rowDate || isNaN(rowDate.getTime())) {
              rowDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), 'dd-MM-yyyy');
            }
          } catch (e) {
            Logger.log('Utilities.parseDate failed, trying standard Date: ' + e);
          }
          
          // Fallback to standard Date parsing if Utilities.parseDate failed
          if (!rowDate || isNaN(rowDate.getTime())) {
            rowDate = new Date(dateRaw);
          }
        } else {
          rowDate = new Date(dateRaw);
        }
        
        if (!rowDate || isNaN(rowDate.getTime())) {
          Logger.log('Failed to parse date in monthly sales report: ' + JSON.stringify(dateRaw));
          return;
        }
      } catch (e) {
        Logger.log('Error parsing date in monthly sales report: ' + e + ', dateRaw: ' + JSON.stringify(row[dateIdx]));
        return;
      }
      
      // Check if date is within the target month
      if (rowDate < startDate || rowDate > endDate) {
        rowsFilteredByDate++;
        return;
      }
      
      rowsWithValidData++;
      
      // Extract row data
      const orderNumber = row[orderNumberIdx] || '';
      const productName = row[productIdx] || '';
      const category = row[categoryIdx] || 'Other';
      const qty = Number(row[qtyIdx]) || 0;
      const unitPrice = Number(row[priceIdx]) || 0;
      const total = Number(row[totalIdx]) || 0;
      const payment = row[paymentIdx] || '';
      const productCode = row[productCodeIdx] || (category + '_' + productName).replace(/\s+/g, '_').toUpperCase();
      
      // Skip if missing essential data
      if (!productName || qty === 0) {
        Logger.log("Skipping row due to missing essential data - Product: " + productName + ", Qty: " + qty);
        return;
      }
      
      // Calculate actual amount (prefer calculated over stored)
      const calculatedAmount = qty * unitPrice;
      const actualAmount = calculatedAmount > 0 ? calculatedAmount : total;
      
      // Format date key
      const dateKey = formatDateKey(rowDate);
      
      // Initialize day data if needed
      if (!dailyData[dateKey]) {
        dailyData[dateKey] = {
          date: dateKey,
          displayDate: formatDisplayDate(rowDate),
          products: {},
          dailySummary: {
            totalSales: 0,
            cashSales: 0,
            otherSales: 0,
            totalQuantity: 0,
            totalOrders: 0,
            orderNumbers: new Set(),
            operationalExpenses: 0,
            cashExpenses: dailyCashExpenses.get(dateKey) || 0,
            otherExpenses: dailyOtherExpenses.get(dateKey) || 0,
            cashPurchases: dailyCashPurchases.get(dateKey) || 0,
            otherPurchases: dailyOtherPurchases.get(dateKey) || 0,
            totalDeposits: dailyDeposits.get(dateKey) || 0,
            cashInHand: 0
          }
        };
      }
      
      const dayData = dailyData[dateKey];
      
      // Group by product code for detailed view
      if (!dayData.products[productCode]) {
        dayData.products[productCode] = {
          productCode: productCode,
          productName: productName,
          category: category,
          totalQuantity: 0,
          unitPrice: unitPrice,
          totalAmount: 0,
          orders: []
        };
      }
      
      const productData = dayData.products[productCode];
      productData.totalQuantity += qty;
      productData.totalAmount += actualAmount;
      
      // Add order details
      productData.orders.push({
        orderNumber: orderNumber,
        quantity: qty,
        unitPrice: unitPrice,
        amount: actualAmount,
        payment: payment
      });
      
      // Update daily summary
      dayData.dailySummary.totalSales += actualAmount;
      dayData.dailySummary.totalQuantity += qty;
      dayData.dailySummary.orderNumbers.add(orderNumber);
      
      // Track payment method - split into cash and other sales
      const paymentLower = payment.toLowerCase().trim();
      if (paymentLower === 'cash' || isPartialPaymentLabel_(payment)) {
        dayData.dailySummary.cashSales += actualAmount;
      } else {
        dayData.dailySummary.otherSales += actualAmount;
      }
      
      // Update month summary
      monthSummary.totalSales += actualAmount;
      monthSummary.totalQuantity += qty;
      monthSummary.uniqueOrderNumbers.add(orderNumber);
      
      if (paymentLower === 'cash' || isPartialPaymentLabel_(payment)) {
        monthSummary.cashSales += actualAmount;
      } else {
        monthSummary.otherSales += actualAmount;
      }
    });
    
    Logger.log(`Row processing summary: Total=${totalRowsProcessed}, FilteredByOutlet=${rowsFilteredByOutlet}, FilteredByDate=${rowsFilteredByDate}, ValidData=${rowsWithValidData}`);

    // 4. Add days with expenses, purchases or deposits but no sales
    const allExpenseDates = new Set([...dailyCashExpenses.keys(), ...dailyOtherExpenses.keys()]);
    const allPurchaseDates = new Set([...dailyCashPurchases.keys(), ...dailyOtherPurchases.keys()]);
    const allDepositDates = new Set([...dailyDeposits.keys()]);
    const allNonSalesDates = new Set([...allExpenseDates, ...allPurchaseDates, ...allDepositDates]);
    
    allNonSalesDates.forEach(dateKey => {
      if (!dailyData[dateKey]) {
        const nonSalesDate = new Date(dateKey + 'T00:00:00'); // Avoid timezone shifts
        dailyData[dateKey] = {
          date: dateKey,
          displayDate: formatDisplayDate(nonSalesDate),
          products: {},
          dailySummary: {
            totalSales: 0,
            cashSales: 0,
            otherSales: 0,
            totalQuantity: 0,
            totalOrders: 0,
            operationalExpenses: 0,
            cashExpenses: dailyCashExpenses.get(dateKey) || 0,
            otherExpenses: dailyOtherExpenses.get(dateKey) || 0,
            cashPurchases: dailyCashPurchases.get(dateKey) || 0,
            otherPurchases: dailyOtherPurchases.get(dateKey) || 0,
            totalDeposits: dailyDeposits.get(dateKey) || 0,
            cashInHand: 0
          }
        };
      }
    });

    // 5. Final calculations
    let totalMonthCashExpenses = 0;
    let totalMonthOtherExpenses = 0;
    let totalMonthCashPurchases = 0;
    let totalMonthOtherPurchases = 0;
    let totalMonthDeposits = 0;
    Object.values(dailyData).forEach(day => {
      // Calculate total operational expenses (including purchases)
      day.dailySummary.operationalExpenses = day.dailySummary.cashExpenses + day.dailySummary.otherExpenses + day.dailySummary.cashPurchases + day.dailySummary.otherPurchases;
      
      // Calculate cash in hand (total sales - cash expenses - cash purchases - deposits)
      day.dailySummary.cashInHand = day.dailySummary.totalSales - day.dailySummary.cashExpenses - day.dailySummary.cashPurchases - day.dailySummary.totalDeposits;
      
      // Add to monthly totals
      totalMonthCashExpenses += day.dailySummary.cashExpenses;
      totalMonthOtherExpenses += day.dailySummary.otherExpenses;
      totalMonthCashPurchases += day.dailySummary.cashPurchases;
      totalMonthOtherPurchases += day.dailySummary.otherPurchases;
      totalMonthDeposits += day.dailySummary.totalDeposits;
    });

    // Convert daily data to array and sort by date
    const dailyArray = Object.values(dailyData).sort((a, b) => a.date.localeCompare(b.date));
    
    // Update final counts and calculations
    monthSummary.totalOrders = monthSummary.uniqueOrderNumbers.size;
    monthSummary.totalDays = dailyArray.length;
    monthSummary.cashExpenses = totalMonthCashExpenses;
    monthSummary.otherExpenses = totalMonthOtherExpenses;
    monthSummary.cashPurchases = totalMonthCashPurchases;
    monthSummary.otherPurchases = totalMonthOtherPurchases;
    monthSummary.operationalExpenses = totalMonthCashExpenses + totalMonthOtherExpenses + totalMonthCashPurchases + totalMonthOtherPurchases;
    monthSummary.totalDeposits = totalMonthDeposits;
    monthSummary.cashInHand = monthSummary.totalSales - monthSummary.cashExpenses - monthSummary.cashPurchases - monthSummary.totalDeposits;
    
    // Convert sets to counts for JSON serialization
    dailyArray.forEach(day => {
      day.dailySummary.totalOrders = day.dailySummary.orderNumbers ? day.dailySummary.orderNumbers.size : 0;
      delete day.dailySummary.orderNumbers; // Remove Set for JSON
    });
    
    delete monthSummary.uniqueOrderNumbers; // Remove Set for JSON

    // Calculate opening cash in hand (all transactions before this month)
    let openingCashInHand = 0;
    try {
      const ss = getGlobalSpreadsheet();
      function sumBeforeMonth(sheetName, amountCol, dateCol, outletCol, paymentCol, paymentFilter) {
        const s = ss.getSheetByName(sheetName);
        if (!s) return 0;
        const rows = s.getDataRange().getValues();
        if (rows.length <= 1) return 0;
        const h = rows[0];
        const di = dateCol != null ? dateCol : h.findIndex(function(x){return x&&/timestamp|date/i.test(x);});
        const ai = amountCol != null ? amountCol : h.findIndex(function(x){return x&&/amount|total/i.test(x);});
        const oi = outletCol != null ? outletCol : h.findIndex(function(x){return x&&/outlet/i.test(x);});
        const pi = paymentCol != null ? paymentCol : h.findIndex(function(x){return x&&/payment/i.test(x);});
        const invI = sheetName === SHEETS.PURCHASES ? purchaseInvestmentColIndex_(h) : -1;
        let total = 0;
        rows.slice(1).forEach(function(row) {
          if (invI !== -1 && isInvestmentPurchaseRow_(row, invI)) return;
          if (oi !== -1 && row[oi] !== userOutlet) return;
          if (paymentFilter && pi !== -1) {
            const pay = (row[pi] || '').toString().toLowerCase().trim();
            const filters = Array.isArray(paymentFilter) ? paymentFilter : [paymentFilter];
            if (!filters.includes(pay)) return;
          }
          const d = row[di] instanceof Date ? row[di] : new Date(row[di]);
          if (!d || isNaN(d.getTime())) return;
          if (d >= startDate) return; // only before this month
          total += Number(row[ai]) || 0;
        });
        return total;
      }
      const prevTotalSales = sumBeforeMonth(SHEETS.ORDERS, null, null, null, null, null);
      const prevCashExpenses = sumBeforeMonth(SHEETS.EXPENSES, null, null, null, null, ['cash', 'partial', 'qris', 'dana']);
      const prevCashPurchases = sumBeforeMonth(SHEETS.PURCHASES, null, null, null, null, 'cash');
      // Deposits have no payment filter
      const prevDeposits = sumBeforeMonth(SHEETS.DEPOSITS, null, null, null, -1, null);
      openingCashInHand = prevTotalSales - prevCashExpenses - prevCashPurchases - prevDeposits;
    } catch(e) {
      Logger.log('Error calculating openingCashInHand: ' + e.toString());
    }

    Logger.log(`Monthly report generated - ${dailyArray.length} days, Total Sales: ${monthSummary.totalSales}`);

    return JSON.stringify({
      success: true,
      data: dailyArray,
      summary: monthSummary,
      openingCashInHand: openingCashInHand,
      month: `${year}-${String(month).padStart(2, '0')}`,
      monthName: getMonthName(month),
      year: year
    });
    
  } catch (error) {
    Logger.log("Error getting monthly sales report:", error.toString());
    return JSON.stringify({
      success: false,
      error: error.toString(),
      data: [],
      summary: {}
    });
  }
}

/**
 * Get monthly sales report summary with option to show/hide product details
 * @param {number} year - The year (e.g., 2025)
 * @param {number} month - The month (1-12)
 * @param {boolean} showProducts - Whether to include product details
 * @param {string} deviceId - Device ID for session
 * @param {string} username - Username for session
 * @returns {string} JSON string containing the report data
 */
function getMonthlySalesReportSummary(year, month, showProducts = false, deviceId, username) {
  try {
    Logger.log(`getMonthlySalesReportSummary - Year: ${year}, Month: ${month}, ShowProducts: ${showProducts}`);
    
    const fullReportJson = getMonthlySalesReport(year, month, deviceId, username);
    const fullReport = JSON.parse(fullReportJson);
    
    if (!fullReport.success) {
      return fullReportJson; // Return error as-is
    }
    
    // If showProducts is false, create summarized version
    if (!showProducts) {
      const summarizedData = fullReport.data.map(day => ({
        date: day.date,
        displayDate: day.displayDate,
        dailySummary: day.dailySummary
      }));
      
      return JSON.stringify({
        ...fullReport,
        data: summarizedData,
        showProducts: false
      });
    }
    
    // Return full report with products
    return JSON.stringify({
      ...fullReport,
      showProducts: true
    });
    
  } catch (error) {
    Logger.log("Error getting monthly sales summary:", error.toString());
    return JSON.stringify({
      success: false,
      error: error.toString(),
      data: [],
      summary: {}
    });
  }
}

// Helper function to format date as YYYY-MM-DD
function formatDateKey(date) {
  return date.getFullYear() + '-' + 
         String(date.getMonth() + 1).padStart(2, '0') + '-' + 
         String(date.getDate()).padStart(2, '0');
}

// Helper function to format date for display
function formatDisplayDate(date) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  
  return `${days[date.getDay()]}, ${months[date.getMonth()]} ${date.getDate()}`;
}

// Helper function to get month name
function getMonthName(month) {
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                 'July', 'August', 'September', 'October', 'November', 'December'];
  return months[month - 1] || 'Unknown';
}

/**
 * Get list of available months that have sales data
 * @returns {string} JSON string containing array of available months
 */
function getAvailableMonths(deviceId, username) {
  try {
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;
    
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (!sheet) {
      return JSON.stringify([]);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return JSON.stringify([]);
    }
    
    const headers = data[0];
    const outletIdx = headers.findIndex(h => /outlet/i.test(h));
    const dateIdx = headers.findIndex(h => /timestamp|date/i.test(h));
    
    const monthsSet = new Set();
    
    data.slice(1).forEach(row => {
      // Filter by outlet - include if no outlet column or outlet matches
      if (outletIdx !== -1 && row[outletIdx] !== userOutlet) return;
      
      // Parse date
      let rowDate;
      try {
        const dateRaw = row[dateIdx];
        
        if (dateRaw instanceof Date) {
          rowDate = dateRaw;
        } else if (typeof dateRaw === 'string') {
          try {
            rowDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!rowDate || isNaN(rowDate.getTime())) {
              rowDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            }
            if (!rowDate || isNaN(rowDate.getTime())) {
              rowDate = Utilities.parseDate(dateRaw, Session.getScriptTimeZone(), 'dd-MM-yyyy');
            }
          } catch (e) {
            // Fallback to standard Date parsing
            rowDate = new Date(dateRaw);
          }
          
          if (!rowDate || isNaN(rowDate.getTime())) {
            rowDate = new Date(dateRaw);
          }
        } else {
          rowDate = new Date(dateRaw);
        }
        
        if (rowDate && !isNaN(rowDate.getTime())) {
          const monthKey = rowDate.getFullYear() + '-' + String(rowDate.getMonth() + 1).padStart(2, '0');
          monthsSet.add(monthKey);
        }
      } catch (e) {
        // Skip invalid dates
      }
    });
    
    // Convert to array and sort (newest first)
    const monthsArray = Array.from(monthsSet)
      .sort((a, b) => b.localeCompare(a))
      .map(monthKey => {
        const [year, month] = monthKey.split('-');
        return {
          key: monthKey,
          year: parseInt(year),
          month: parseInt(month),
          displayName: getMonthName(parseInt(month)) + ' ' + year
        };
      });
    
    return JSON.stringify(monthsArray);
    
  } catch (error) {
    Logger.log("Error getting available months:", error.toString());
    return JSON.stringify([]);
  }
}

// Test function for monthly sales report
function testMonthlySalesReport() {
  try {
    Logger.log("=== Testing Monthly Sales Report ===");
    
    const currentDate = new Date();
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth() + 1;
    
    Logger.log(`Testing with current month: ${year}-${month}`);
    
    // Test summary version with null session (for testing)
    const summaryResult = getMonthlySalesReportSummary(year, month, false, null, null);
    Logger.log("Summary result:", summaryResult);
    
    // Test detailed version
    const detailedResult = getMonthlySalesReportSummary(year, month, true, null, null);
    Logger.log("Detailed result:", detailedResult);
    
    // Test available months
    const availableMonths = getAvailableMonths(null, null);
    Logger.log("Available months:", availableMonths);
    
    return "Test completed successfully";
    
  } catch (error) {
    Logger.log("Error in monthly sales report test:", error.toString());
    return "Error: " + error.toString();
  }
}

// Enhanced debug function to diagnose monthly sales issues
function debugMonthlySalesData() {
  try {
    Logger.log("=== MONTHLY SALES DEBUG ===");
    
    // 1. Check spreadsheet access
    const spreadsheet = getGlobalSpreadsheet();
    Logger.log("Spreadsheet ID: " + spreadsheet.getId());
    
    // 2. Check Orders sheet
    const ordersSheet = spreadsheet.getSheetByName(SHEETS.ORDERS);
    Logger.log("Orders sheet found: " + !!ordersSheet);
    
    if (!ordersSheet) {
      Logger.log("ERROR: Orders sheet not found with name: " + SHEETS.ORDERS);
      return "Orders sheet not found";
    }
    
    // 3. Check data
    const data = ordersSheet.getDataRange().getValues();
    Logger.log("Total rows: " + data.length);
    
    if (data.length <= 1) {
      Logger.log("No data in Orders sheet");
      return "No data found";
    }
    
    // 4. Analyze headers
    const headers = data[0];
    Logger.log("Headers: " + JSON.stringify(headers));
    
    const dateIdx = headers.findIndex(h => /timestamp|date/i.test(h));
    const productIdx = headers.findIndex(h => /product|item|name/i.test(h));
    const qtyIdx = headers.findIndex(h => /qty|quantity/i.test(h));
    const priceIdx = headers.findIndex(h => /price/i.test(h));
    const totalIdx = headers.findIndex(h => /total|amount/i.test(h));
    
    Logger.log(`Column indices - Date: ${dateIdx}, Product: ${productIdx}, Qty: ${qtyIdx}, Price: ${priceIdx}, Total: ${totalIdx}`);
    
    // 5. Sample data analysis
    let validRowCount = 0;
    let sampleCount = Math.min(10, data.length - 1);
    
    for (let i = 1; i <= sampleCount; i++) {
      const row = data[i];
      const dateRaw = row[dateIdx];
      const product = row[productIdx];
      const qty = Number(row[qtyIdx]) || 0;
      const price = Number(row[priceIdx]) || 0;
      const total = Number(row[totalIdx]) || 0;
      
      Logger.log(`Row ${i}: Date="${dateRaw}", Product="${product}", Qty=${qty}, Price=${price}, Total=${total}`);
      
      if (product && qty > 0) {
        validRowCount++;
      }
    }
    
    Logger.log(`Valid rows in sample: ${validRowCount}/${sampleCount}`);
    
    // 6. Test current month function
    const currentDate = new Date();
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth() + 1;
    
    Logger.log(`Testing current month: ${year}-${month}`);
    const result = getMonthlySalesReportSummary(year, month, false, null, null);
    Logger.log("Function result: " + result);
    
    try {
      const parsed = JSON.parse(result);
      Logger.log("Parsed success: " + parsed.success);
      Logger.log("Parsed error: " + parsed.error);
      Logger.log("Data length: " + (parsed.data ? parsed.data.length : 'N/A'));
      Logger.log("Summary: " + JSON.stringify(parsed.summary));
    } catch (parseError) {
      Logger.log("Parse error: " + parseError.toString());
    }
    
    return "Debug completed";
    
  } catch (error) {
    Logger.log("Debug error: " + error.toString());
    return "Error: " + error.toString();
  }
}

// Test with sample data insertion
function testWithSampleData() {
  try {
    Logger.log("=== TESTING WITH SAMPLE DATA ===");
    
    const spreadsheet = getGlobalSpreadsheet();
    const ordersSheet = spreadsheet.getSheetByName(SHEETS.ORDERS);
    
    if (!ordersSheet) {
      Logger.log("Orders sheet not found");
      return "Orders sheet not found";
    }
    
    // Check if there's already data
    const existingData = ordersSheet.getDataRange().getValues();
    Logger.log("Existing rows: " + existingData.length);
    
    if (existingData.length <= 1) {
      Logger.log("No existing data found. Need to check sheet structure or add sample data.");
      Logger.log("Headers should be something like: Timestamp, Product, Quantity, Price, Total, etc.");
      
      // If there are no headers, we can't proceed
      if (existingData.length === 0) {
        Logger.log("ERROR: Sheet is completely empty - no headers found");
        return "Sheet is empty";
      }
      
      Logger.log("Headers found: " + JSON.stringify(existingData[0]));
      return "Empty sheet but headers exist";
    }
    
    // Test the function with current month
    const currentDate = new Date();
    const result = getMonthlySalesReportSummary(currentDate.getFullYear(), currentDate.getMonth() + 1, false, null, null);
    Logger.log("Test result: " + result);
    
    return "Test completed";
    
  } catch (error) {
    Logger.log("Test error: " + error.toString());
    return "Error: " + error.toString();
  }
}

// Test function to verify outlet filtering for expenses and deposits
function testOutletFiltering(deviceId, username) {
  try {
    Logger.log("=== Testing Outlet Filtering ===");
    
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;
    Logger.log("Current user outlet: " + (userOutlet || 'NOT SET'));
    
    // Test expenses filtering
    const expensesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.EXPENSES);
    if (expensesSheet) {
      const expensesData = expensesSheet.getDataRange().getValues();
      if (expensesData.length > 1) {
        const expHeaders = expensesData[0];
        const expOutletIdx = expHeaders.findIndex(h => /outlet/i.test(h));
        Logger.log("Expenses sheet - outlet column index: " + expOutletIdx);
        
        if (expOutletIdx >= 0) {
          const totalExpenseRows = expensesData.length - 1;
          const filteredExpenseRows = expensesData.slice(1).filter(row => 
            !userOutlet || row[expOutletIdx] === userOutlet
          ).length;
          Logger.log(`Expenses: ${filteredExpenseRows}/${totalExpenseRows} rows match outlet filter`);
        }
      }
    }
    
    // Test deposits filtering
    const depositsSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.DEPOSITS);
    if (depositsSheet) {
      const depositsData = depositsSheet.getDataRange().getValues();
      if (depositsData.length > 1) {
        const depHeaders = depositsData[0];
        const depOutletIdx = depHeaders.findIndex(h => /outlet/i.test(h));
        Logger.log("Deposits sheet - outlet column index: " + depOutletIdx);
        
        if (depOutletIdx >= 0) {
          const totalDepositRows = depositsData.length - 1;
          const filteredDepositRows = depositsData.slice(1).filter(row => 
            !userOutlet || row[depOutletIdx] === userOutlet
          ).length;
          Logger.log(`Deposits: ${filteredDepositRows}/${totalDepositRows} rows match outlet filter`);
        }
      }
    }
    
    // Test orders filtering
    const ordersSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (ordersSheet) {
      const ordersData = ordersSheet.getDataRange().getValues();
      if (ordersData.length > 1) {
        const orderHeaders = ordersData[0];
        const orderOutletIdx = orderHeaders.findIndex(h => /outlet/i.test(h));
        Logger.log("Orders sheet - outlet column index: " + orderOutletIdx);
        
        if (orderOutletIdx >= 0) {
          const totalOrderRows = ordersData.length - 1;
          const filteredOrderRows = ordersData.slice(1).filter(row => 
            !userOutlet || row[orderOutletIdx] === userOutlet
          ).length;
          Logger.log(`Orders: ${filteredOrderRows}/${totalOrderRows} rows match outlet filter`);
        }
      }
    }
    
    Logger.log("=== Outlet Filtering Test Complete ===");
    return "Outlet filtering test completed - check logs for details";
    
  } catch (error) {
    Logger.log("Outlet filtering test error: " + error.toString());
    return "Error: " + error.toString();
  }
}

/**
 * Get total cash in hand across all time (All Cash Sales - All Cash Expenses - All Deposits)
 * @returns {string} JSON string containing the total cash in hand
 */
function getTotalCashInHand(deviceId, username) {
  try {
    Logger.log("getTotalCashInHand - Calculating total cash in hand across all time");
    
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;
    
    let totalCashSales = 0;
    let totalCashExpenses = 0;
    let totalCashPurchases = 0;
    let totalDeposits = 0;
    
    // 1. Calculate total sales (all payment methods)
    const ordersSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (ordersSheet) {
      const ordersData = ordersSheet.getDataRange().getValues();
      if (ordersData.length > 1) {
        const headers = ordersData[0];
        const outletIdx = headers.findIndex(h => /outlet/i.test(h));
        const totalIdx = headers.findIndex(h => /total|amount/i.test(h));
        
        if (totalIdx > -1) {
          ordersData.slice(1).forEach(row => {
            // Filter by outlet - include if no outlet column or outlet matches
            if (outletIdx !== -1 && row[outletIdx] !== userOutlet) return;
            
            const amount = Number(row[totalIdx]) || 0;
            totalCashSales += amount;
          });
        }
      }
    }
    
    // 2. Calculate total cash expenses
    const expensesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.EXPENSES);
    if (expensesSheet) {
      const expensesData = expensesSheet.getDataRange().getValues();
      if (expensesData.length > 1) {
        const expHeaders = expensesData[0];
        const expOutletIdx = expHeaders.findIndex(h => /outlet/i.test(h));
        const expAmountIdx = expHeaders.findIndex(h => /amount|total/i.test(h));
        const expPaymentIdx = expHeaders.findIndex(h => /payment/i.test(h));
        
        if (expAmountIdx > -1) {
          expensesData.slice(1).forEach(row => {
            // Filter by outlet - include if no outlet column or outlet matches
            if (expOutletIdx !== -1 && row[expOutletIdx] !== userOutlet) return;
            
            const amount = Number(row[expAmountIdx]) || 0;
            const payment = expPaymentIdx > -1 ? (row[expPaymentIdx] || '').toLowerCase().trim() : 'cash';
            
            if (['cash', 'partial', 'qris', 'dana'].includes(payment)) {
              totalCashExpenses += amount;
              if (amount < 0) {
                Logger.log(`Negative cash expense included: ${amount}`);
              }
            }
          });
        }
      }
    }
    
    // 3. Calculate total cash purchases
    const purchasesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
    if (purchasesSheet) {
      const purchasesData = purchasesSheet.getDataRange().getValues();
      if (purchasesData.length > 1) {
        const purHeaders = purchasesData[0];
        const purOutletIdx = purHeaders.findIndex(h => /outlet/i.test(h));
        const purAmountIdx = purHeaders.findIndex(h => /total|amount/i.test(h));
        const purPaymentIdx = purHeaders.findIndex(h => /_selectingPayment|payment/i.test(h));
        const purInvIdx = purchaseInvestmentColIndex_(purHeaders);
        
        if (purAmountIdx > -1) {
          purchasesData.slice(1).forEach(row => {
            // Exclude investment (capital) purchases from cash flow
            if (isInvestmentPurchaseRow_(row, purInvIdx)) return;
            // Filter by outlet - include if no outlet column or outlet matches
            if (purOutletIdx !== -1 && row[purOutletIdx] !== userOutlet) return;
            
            const amount = Number(row[purAmountIdx]) || 0;
            const payment = purPaymentIdx > -1 ? (row[purPaymentIdx] || '').toLowerCase().trim() : 'cash';
            
            if (payment === 'cash') {
              totalCashPurchases += amount;
              if (amount < 0) {
                Logger.log(`Negative cash purchase included: ${amount}`);
              }
            }
          });
        }
      }
    }
    
    // 4. Calculate total deposits
    const depositsSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.DEPOSITS);
    if (depositsSheet) {
      const depositsData = depositsSheet.getDataRange().getValues();
      if (depositsData.length > 1) {
        const depHeaders = depositsData[0];
        const depOutletIdx = depHeaders.findIndex(h => /outlet/i.test(h));
        const depAmountIdx = depHeaders.findIndex(h => /amount|total/i.test(h));
        
        if (depAmountIdx > -1) {
          depositsData.slice(1).forEach(row => {
            // Filter by outlet - include if no outlet column or outlet matches
            if (depOutletIdx !== -1 && row[depOutletIdx] !== userOutlet) return;
            
            const amount = Number(row[depAmountIdx]) || 0;
            if (amount > 0) {
              totalDeposits += amount;
            }
          });
        }
      }
    }
    
    // Calculate total cash in hand (all sales - cash expenses - cash purchases - all deposits)
    const totalCashInHand = totalCashSales - totalCashExpenses - totalCashPurchases - totalDeposits;
    
    Logger.log(`Total Cash In Hand calculation: ${totalCashSales} - ${totalCashExpenses} - ${totalCashPurchases} - ${totalDeposits} = ${totalCashInHand}`);
    Logger.log(`Breakdown - Total Sales: ${totalCashSales}, Cash Expenses: ${totalCashExpenses}, Cash Purchases: ${totalCashPurchases}, Deposits: ${totalDeposits}`);
    
    return JSON.stringify({
      success: true,
      totalCashInHand: totalCashInHand,
      breakdown: {
        totalCashSales: totalCashSales,
        totalCashExpenses: totalCashExpenses,
        totalCashPurchases: totalCashPurchases,
        totalDeposits: totalDeposits
      }
    });
    
  } catch (error) {
    Logger.log("Error getting total cash in hand:", error.toString());
    return JSON.stringify({
      success: false,
      error: error.toString(),
      totalCashInHand: 0
    });
  }
}
