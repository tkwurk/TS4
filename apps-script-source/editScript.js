function getRows(date, outlet) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];

    // Prepare input date as Date object, at midnight
    let inputDateObj = null;
    if (date) {
      inputDateObj = new Date(date);
      inputDateObj.setHours(0, 0, 0, 0);
    }

    for (let i = 1; i < data.length; i++) {
      let cell = data[i][0]; // First column for date
      let cellDateObj = null;
      if (cell instanceof Date) {
        cellDateObj = new Date(cell.getFullYear(), cell.getMonth(), cell.getDate());
      } else if (typeof cell === 'string' && cell.length >= 6) {
        let parsed = new Date(cell);
        if (!isNaN(parsed)) {
          cellDateObj = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
        }
      }
      if (cellDateObj) cellDateObj.setHours(0,0,0,0);

      let dateMatch = true;
      if (inputDateObj) {
        dateMatch = cellDateObj && (cellDateObj.valueOf() === inputDateObj.valueOf());
      }

      // Outlet in column 9
      let outletMatch = true;
      if (outlet) {
        let outletVal = (data[i][8] || "").toString().trim();
        outletMatch = outletVal.toLowerCase() === outlet.toLowerCase();
      }

      if (dateMatch && outletMatch) {
        result.push([i + 1, ...data[i]]);
      }
    }

    // Return as JSON string
    return JSON.stringify({ headers: ["Row#"].concat(headers), rows: result });

  } catch (e) {
    Logger.log("Error in getRows: " + e);
    return JSON.stringify({ headers: [], rows: [] });
  }
}


function updateEditableFields(rowNumber, qty, price, paymentMethod) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  sheet.getRange(rowNumber, 4).setValue(qty);          // Qty: column 4 (index 3), so column 4+1=5
  sheet.getRange(rowNumber, 5).setValue(price);         // Price: column 5 (index 4), so 5+1=6
  sheet.getRange(rowNumber, 6).setValue(paymentMethod); // PaymentMethod: column 6 (index 5), so 6+1=7
   sheet.getRange(rowNumber, 7).setValue(Number(qty) * Number(price));
  return true;
}

function deleteRow(rowNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  sheet.deleteRow(rowNumber);
  return true;
}
