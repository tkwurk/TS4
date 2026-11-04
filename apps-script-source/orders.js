function validateOrderCart_(order) {
  if (!order.items || !Array.isArray(order.items) || order.items.length === 0) {
    throw new Error('No items in order');
  }
  order.items.forEach(function(item, index) {
    if (!item || !item.name || !String(item.name).trim()) {
      throw new Error('Item ' + (index + 1) + ' is missing a product name');
    }
    var qty = Number(item.qty);
    if (!qty || qty <= 0 || isNaN(qty)) {
      throw new Error('Item ' + (index + 1) + ' has invalid quantity');
    }
    var price = Number(item.price);
    if (isNaN(price) || price < 0) {
      throw new Error('Item ' + (index + 1) + ' has invalid price');
    }
  });
}

function buildOrderItemRow_(timestamp, uniqueOrderNumber, item, order, session, resolvedOutlet) {
  var itemTotal = Number(item.price) * Number(item.qty);
  var productCode = generateProductCode(item);
  var outlet = resolvedOutlet || session.outlet || '';
  var payment = order.payment;
  // Split payments: record full sale as cash; auto-deposit legs handled separately.
  if (typeof isSplitPaymentOrder_ === 'function' && isSplitPaymentOrder_(order)) {
    payment = 'cash';
  }
  return [
    formatTimestampForSheet(timestamp),
    uniqueOrderNumber,
    item.name,
    item.qty,
    item.price,
    payment,
    itemTotal,
    item.category || '',
    outlet,
    session.username || '',
    productCode,
    item.navCategory || item.category || ''
  ];
}

function buildOrderRows_(order, session, timestamp, baseOrderNumber) {
  var rows = [];
  var generatedOrderNumbers = [];
  var outlet = (session.outlet || '').toString().trim();
  order.items.forEach(function(item, index) {
    var itemNumber = String(index + 1).padStart(3, '0');
    var uniqueOrderNumber = baseOrderNumber + '-' + itemNumber;
    generatedOrderNumbers.push(uniqueOrderNumber);
    rows.push(buildOrderItemRow_(timestamp, uniqueOrderNumber, item, order, session, outlet));
  });
  return { rows: rows, generatedOrderNumbers: generatedOrderNumbers };
}

function shouldAutoDepositForMethod_(methodName, autoDepositFlag, configMap) {
  if (autoDepositFlag === true) return true;
  if (autoDepositFlag === false) return false;
  if (typeof paymentMethodHasAutoDeposit_ === 'function') {
    return paymentMethodHasAutoDeposit_(methodName, configMap || {});
  }
  return true;
}

function depositAccountForMethod_(methodName, destAccount, configMap) {
  if (destAccount) return destAccount;
  var key = (methodName || '').toString().trim().toLowerCase();
  var meta = configMap && configMap[key];
  if (meta && meta.destinationAccount) return meta.destinationAccount;
  return methodName;
}

function depositLabelForMethod_(methodName, configMap) {
  if (configMap && Object.keys(configMap).length && typeof canonicalPaymentMethodName_ === 'function') {
    return canonicalPaymentMethodName_(methodName, configMap);
  }
  return methodName;
}

function orderIncludesClientPaymentMeta_(order) {
  if (!order) return false;
  if (order.autoDeposit !== undefined) return true;
  if (order.paymentSplits && order.paymentSplits.length) {
    return order.paymentSplits.some(function(s) { return s.autoDeposit !== undefined; });
  }
  return false;
}

function buildAutoDepositRows_(order, session, timestamp, generatedOrderNumbers) {
  var useSheetConfig = !orderIncludesClientPaymentMeta_(order);
  var configMap = useSheetConfig && typeof loadPaymentMethodConfigMap_ === 'function'
    ? loadPaymentMethodConfigMap_()
    : {};
  var note = generatedOrderNumbers.join(', ');
  var rows = [];

  function pushDeposit(methodName, amount, destAccount, autoDeposit) {
    var amt = Number(amount) || 0;
    if (amt <= 0) return;
    if (!shouldAutoDepositForMethod_(methodName, autoDeposit, configMap)) {
      return;
    }
    rows.push([
      formatTimestampForSheet(timestamp),
      amt,
      depositAccountForMethod_(methodName, destAccount, configMap),
      note,
      session.username || '',
      session.outlet || '',
      depositLabelForMethod_(methodName, configMap)
    ]);
  }

  if (typeof getOrderPaymentSplits_ === 'function' && isSplitPaymentOrder_(order)) {
    getOrderPaymentSplits_(order).forEach(function(split) {
      pushDeposit(split.method, split.amount, split.destinationAccount, split.autoDeposit);
    });
    return rows;
  }

  var orderPaymentLower = (order.payment || '').toLowerCase();
  if (order.payment && !isCashLikePaymentMethodName_(orderPaymentLower) && !isCheckoutOnlyPaymentMethodName_(orderPaymentLower)) {
    pushDeposit(order.payment, Number(order.total) || 0, order.destinationAccount, order.autoDeposit);
  }
  return rows;
}

function buildAutoDepositRow_(order, session, timestamp, generatedOrderNumbers) {
  var rows = buildAutoDepositRows_(order, session, timestamp, generatedOrderNumbers);
  return rows.length ? rows[0] : null;
}

function saveOrder(orderJson) {
  try {
    var order = typeof orderJson === 'string' ? JSON.parse(orderJson) : orderJson;
    var deviceId = order.deviceId;
    var username = order.username;

    if (!deviceId || !username) {
      return apiFail('Device ID and username are required to save an order', 'ORDER_AUTH_REQUIRED');
    }

    var session = getUserSessionEnhanced(deviceId, username, {
      skipPeopleRefresh: true,
      touchMs: 120000
    });
    if (!session || !session.username) {
      return apiFail('No valid session found for the provided device and user', 'ORDER_SESSION_INVALID');
    }
    if (session.username !== username) {
      return apiFail('Session username does not match order username', 'ORDER_SESSION_MISMATCH');
    }

    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    if (!sheet) {
      return apiFail('Orders sheet not found', 'ORDER_SHEET_MISSING');
    }

    validateOrderCart_(order);

    var timestamp = new Date();
    var baseOrderNumber = order.orderNumber
      ? order.orderNumber.replace(/-\d{3}$/, '')
      : generateOrderNumber();

    var built = buildOrderRows_(order, session, timestamp, baseOrderNumber);
    appendRowsAtomic_(sheet, built.rows);

    var depositRows = buildAutoDepositRows_(order, session, timestamp, built.generatedOrderNumbers);
    if (depositRows.length) {
      try {
        var depositsSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.DEPOSITS);
        if (depositsSheet) {
          appendRowsAtomic_(depositsSheet, depositRows);
        } else {
          Logger.log('WARNING: Deposits sheet not found, could not auto-deposit');
        }
      } catch (depositError) {
        Logger.log('WARNING: Failed to auto-deposit: ' + depositError.toString());
      }
    }

    invalidatePnLCache_(session.outlet);
    if (typeof invalidateCookingStockCache_ === 'function') {
      try { invalidateCookingStockCache_(session.outlet); } catch (kcErr) { Logger.log('cooking cache invalidation failed: ' + kcErr); }
    }
    return apiOk({
      orderNumber: baseOrderNumber,
      orderNumbers: built.generatedOrderNumbers,
      itemCount: order.items.length
    }, 'ORDER_SAVED');

  } catch (error) {
    Logger.log("ERROR saving order: " + error.toString());
    return apiFail(error.message || error, 'ORDER_SAVE_FAILED');
  }
}

function generateOrderNumber() {
  var now = new Date();
  var year = now.getFullYear();
  var month = String(now.getMonth() + 1).padStart(2, '0');
  var day = String(now.getDate()).padStart(2, '0');
  var hours = String(now.getHours()).padStart(2, '0');
  var minutes = String(now.getMinutes()).padStart(2, '0');
  var seconds = String(now.getSeconds()).padStart(2, '0');
  var random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return 'ORD-' + year + month + day + '-' + hours + minutes + seconds + '-' + random;
}

function saveOrderSecure(params) {
  return withAgentGuard_('checkout.save', params, function () {
    var orderData = params.orderData;
    if (!orderData) {
      return apiFail('Order data is required', 'ORDER_DATA_MISSING');
    }
    var authContext = params._authContext;
    if (!authContext || !authContext.deviceId || !authContext.username) {
      return apiFail('Invalid authentication context', 'ORDER_AUTH_INVALID');
    }
    var order = typeof orderData === 'string' ? JSON.parse(orderData) : orderData;
    order.deviceId = authContext.deviceId;
    order.username = authContext.username;
    return saveOrder(JSON.stringify(order));
  });
}

function saveMissedOrdersSecure(params) {
  try {
    var missedData = params.missedData;
    if (!missedData) {
      return apiFail('Missed order data is required', 'MISSED_DATA_MISSING');
    }
    var authContext = params._authContext;
    if (!authContext || !authContext.deviceId || !authContext.username) {
      return apiFail('Invalid authentication context', 'MISSED_AUTH_INVALID');
    }

    var data = typeof missedData === 'string' ? JSON.parse(missedData) : missedData;
    var session = getUserSessionEnhanced(authContext.deviceId, authContext.username, {
      skipPeopleRefresh: true,
      touchMs: 120000
    });
    if (!session || !session.username) {
      return apiFail('No valid session found', 'MISSED_SESSION_INVALID');
    }

    var items = data.items || [];
    if (!items.length) {
      return apiOk({ itemCount: 0 }, 'MISSED_SAVED');
    }

    var ss = getGlobalSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MISSED_ORDERS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.MISSED_ORDERS);
      sheet.appendRow(['Timestamp', 'OrderNumber', 'ProductName', 'Quantity', 'Price', 'Total', 'Category', 'Outlet', 'Username', 'ProductCode']);
    }

    var timestamp = new Date();
    var baseOrderNumber = (data.orderNumber || 'MISS-' + timestamp.getTime()).replace(/-\d{3}$/, '');
    var rows = [];
    items.forEach(function(item, index) {
      var itemTotal = (item.price || 0) * (item.qty || 0);
      var itemNumber = String(index + 1).padStart(3, '0');
      var uniqueNumber = 'MISS-' + baseOrderNumber + '-' + itemNumber;
      var productCode = typeof generateProductCode === 'function' ? generateProductCode(item) : (item.name || '');
      rows.push([
        formatTimestampForSheet(timestamp),
        uniqueNumber,
        item.name || '',
        item.qty || 0,
        item.price || 0,
        itemTotal,
        item.category || '',
        session.outlet || '',
        session.username || '',
        productCode
      ]);
    });
    appendRowsAtomic_(sheet, rows);

    Logger.log("=== MISSED ORDERS SAVED: " + items.length + " items ===");
    return apiOk({ itemCount: items.length }, 'MISSED_SAVED');
  } catch (error) {
    Logger.log("ERROR in saveMissedOrdersSecure: " + error.toString());
    return apiFail(error.message || error, 'MISSED_SAVE_FAILED');
  }
}
