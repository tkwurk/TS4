function sendDailyPurchaseSummary() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var yesterday = new Date(now.getTime() - 24*60*60*1000);
  var yesterdayStr = Utilities.formatDate(yesterday, tz, "dd-MM-yyyy");
 var tOutlet = SpreadsheetApp.getActive().getName().substring(0, 3);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Purchases"); // adjust name
  var data = sheet.getDataRange().getValues();

  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ordersSheet = ss.getSheetByName("Orders");
  var employeesSheet = ss.getSheetByName("Employees");

  // --- Get emails for all positions except Cashier ---
  var empData = employeesSheet.getDataRange().getValues();
  var recipients = empData.slice(1) // skip header
    .filter(function(row) { return row[3] !== "Cashier" && row[7]; })
    .map(function(row) { return row[7]; });

  if (recipients.length === 0) {
    Logger.log("No eligible emails found.");
    return;
  }

  // Header indices
  var TIMESTAMP = 0;
  var PRODUCT_NAME = 2;
  var PRODUCT_CODE = 3;
  var QTY = 4;
  var PRICE = 5;
  var CATEGORY = 6;
  var AMOUNT = 9;
  var MODE = 15; // "Mode" column index based on your header list

  var categorySummary = {};
  var lastRow = data.length - 1;

  // Scan bottom → top
  for (var i = lastRow; i >= 1; i--) {
    var row = data[i];
    var ts = row[TIMESTAMP];
    if (!(ts instanceof Date)) continue;

    var tsDateStr = Utilities.formatDate(ts, tz, "dd-MM-yyyy");
    if (tsDateStr !== yesterdayStr) {
      // Early stop: once we pass yesterday’s block, break
      if (ts < yesterday) break;
      continue;
    }

    // Skip if Mode == "N"
  const mode = String(row[MODE]).trim();
if (mode === "N" || mode === "") continue;

    var cat = row[CATEGORY] || "Uncategorized";
    if (!categorySummary[cat]) {
      categorySummary[cat] = {
        items: [],
        totalQty: 0,
        totalAmount: 0
      };
    }

    var qty = Number(row[QTY]);
    var amount = Number(row[AMOUNT]);
    categorySummary[cat].items.push({
      name: row[PRODUCT_NAME],
      code: row[PRODUCT_CODE],
      qty: qty,
      price: Number(row[PRICE]),
      amount: amount
    });
    categorySummary[cat].totalQty += qty;
    categorySummary[cat].totalAmount += amount;
  }

  if (Object.keys(categorySummary).length === 0) {
    Logger.log("No purchases found for " + yesterdayStr);
    return;
  }

  // Build HTML email body
  var body = "<p>Hello Team,</p>";
  body += "<p>TS3 - Purchase summary for <b>" + yesterdayStr + "</b>:</p>";

  var grandTotal = 0;
  for (var cat in categorySummary) {
    var catData = categorySummary[cat];
    body += "<h3>Category: " + cat + "</h3>";
    body += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;'>"
         + "<tr style='background-color:#f2f2f2;'><th>Product Name</th><th>Code</th><th>Qty</th><th>Price</th><th>Amount</th></tr>";
    for (var j = 0; j < catData.items.length; j++) {
      var item = catData.items[j];
      body += "<tr><td>" + item.name + "</td><td>" + item.code + "</td><td>" + item.qty + "</td><td>" + item.price + "</td><td>" + item.amount + "</td></tr>";
    }
    body += "</table>";
    body += "<p><b>Total for " + cat + ":</b> Qty: " + catData.totalQty + " | Amount: " + catData.totalAmount + "</p>";
    grandTotal += catData.totalAmount;
  }

  body += "<p><b>Grand Total:</b> " + grandTotal + "</p>";
  body += "<p>Best regards,<br>Purchase Bot</p>";

Logger.log( recipients.join(","));
Logger.log(MailApp.getRemainingDailyQuota() );


  MailApp.sendEmail({
    to: recipients.join(","), // adjust recipient
    subject: tOutlet+" Daily Purchase Summary – " + yesterdayStr,
    htmlBody: body
  });
}

// Trigger setup: run daily at 10:00 AM
function createDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendDailyPurchaseSummary") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("sendDailyPurchaseSummary")
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();
}
