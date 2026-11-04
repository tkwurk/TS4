function sendSalesReportHTML() {
  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var hour = now.getHours();
  var minute = now.getMinutes();
  var tOutlet = SpreadsheetApp.getActive().getName().substring(0, 3);

  // Guard: skip if between 21:30 and 08:00
  if ((hour > 21) || (hour === 21 && minute >= 30) || (hour < 7)) {
    Logger.log("Report skipped due to off-hours.");
    return;
  }

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

  var lastRow = ordersSheet.getLastRow();
  var today = now;
  var startRow = null;

  // Scan timestamps once
  var tsData = ordersSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = tsData.length - 1; i >= 0; i--) {
    var ts = tsData[i][0];
    if (ts instanceof Date) {
      if (ts.getDate() === today.getDate() &&
          ts.getMonth() === today.getMonth() &&
          ts.getFullYear() === today.getFullYear()) {
        startRow = i + 2;
      } else if (startRow) {
        break;
      }
    }
  }

  if (!startRow) {
    Logger.log("No sales found for today.");
     MailApp.sendEmail({
      to: recipients.join(","),
      subject: tOutlet + " Hourly & Daily Sales Report for " +" – " + todayStr,
      htmlBody: "<p>No sales found for today.</p>"
    });

  } else {
    // Pull all needed columns in one go
    var data = ordersSheet.getRange(startRow, 1, lastRow - startRow + 1, 12).getValues();
    var dailySummaryMap = {};
    var lastHourSummaryMap = {};
    var currentHour = now.getHours();

    data.forEach(function(row) {
      var ts = row[0];
      if (!(ts instanceof Date)) return;

      if (ts.getDate() === today.getDate() &&
          ts.getMonth() === today.getMonth() &&
          ts.getFullYear() === today.getFullYear()) {

        var item = row[2];
        var qty = Number(row[3]);
        var total = Number(row[6]);
        var navcat = row[11];
        var key = navcat + "|" + item;

        if (!dailySummaryMap[key]) dailySummaryMap[key] = {item:item, navcat:navcat, qty:0, total:0};
        dailySummaryMap[key].qty += qty;
        dailySummaryMap[key].total += total;

        if (currentHour - ts.getHours() < 1) {
          if (!lastHourSummaryMap[key]) lastHourSummaryMap[key] = {item:item, navcat:navcat, qty:0, total:0};
          lastHourSummaryMap[key].qty += qty;
          lastHourSummaryMap[key].total += total;
        }
      }
    });

    // Convert maps to arrays and sort
    var dailySummary = Object.values(dailySummaryMap).sort(function(a, b) {
      var catCompare = a.navcat.localeCompare(b.navcat);
      if (catCompare !== 0) return catCompare;
      return a.item.localeCompare(b.item);
    });

    var lastHourSummary = Object.values(lastHourSummaryMap).sort(function(a, b) {
      var catCompare = a.navcat.localeCompare(b.navcat);
      if (catCompare !== 0) return catCompare;
      return a.item.localeCompare(b.item);
    });

    // Build HTML body
    var bodyParts = [];
    var todayStr = Utilities.formatDate(now, tz, "dd-MM-yyyy");

    bodyParts.push("<p>Hello Team,</p>");
    bodyParts.push("<p>Here’s the sales summary for " + tOutlet +" <b>" + todayStr + "</b>:</p>");

    // Last hour section
    bodyParts.push("<h3>Last 2 Hours Item Sales (" + currentHour + ":00)</h3>");
    bodyParts.push("<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;'>"
         + "<tr style='background-color:#f2f2f2;'><th>NAVCATEGORY</th><th>Item Name</th><th>Qty Sold</th><th>Sales Amount</th></tr>");
    var totalHourQty = 0, totalHourSales = 0;
    var lastHourRows = [];
    lastHourSummary.forEach(function(e) {
      lastHourRows.push("<tr><td>" + e.navcat + "</td><td>" + e.item + "</td><td>" + e.qty + "</td><td>" + e.total + "</td></tr>");
      totalHourQty += e.qty;
      totalHourSales += e.total;
    });
    bodyParts.push(lastHourRows.join(""));
    bodyParts.push("</table>");
    bodyParts.push("<p><b>Total for Last Hour:</b> Qty: " + totalHourQty + " | Sales: " + totalHourSales + "</p>");

    // Daily section
    bodyParts.push("<h3>Daily Item Sales (All Sales Today)</h3>");
    bodyParts.push("<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;'>"
         + "<tr style='background-color:#f2f2f2;'><th>NAVCATEGORY</th><th>Item Name</th><th>Total Qty</th><th>Total Sales</th></tr>");
    var totalDayQty = 0, totalDaySales = 0;
    var dailyRows = [];
    dailySummary.forEach(function(e) {
      dailyRows.push("<tr><td>" + e.navcat + "</td><td>" + e.item + "</td><td>" + e.qty + "</td><td>" + e.total + "</td></tr>");
      totalDayQty += e.qty;
      totalDaySales += e.total;
    });
    bodyParts.push(dailyRows.join(""));
    bodyParts.push("</table>");
    bodyParts.push("<p><b>Daily Totals:</b> Qty: " + totalDayQty + " | Sales: " + totalDaySales + "</p>");

    bodyParts.push("<p>Best regards,<br>Sales Bot</p>");

    var body = bodyParts.join("");

    MailApp.sendEmail({
      to: recipients.join(","),
      subject: "Hourly & Daily Sales Report for " + tOutlet + " - " + todayStr,
      htmlBody: body
    });
  }

  checkAttendanceAndNotify(recipients,tOutlet);
  rescheduleSendSalesReportHTML();
}


// Attendance check optimized
function checkAttendanceAndNotify(recipients,tOutlet) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    MailApp.sendEmail(recipients.join(","), "Attendance Missing", "No attendance data entered today.");
    return;
  }

  var now = new Date();
  var todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy");

  // Batch read column 3 once
  var dates = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  var found = false;
  for (var i = dates.length - 1; i >= 0; i--) {
    var dateVal = dates[i][0];
    if (dateVal instanceof Date) {
      var dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd-MM-yyyy");
      if (dateStr === todayStr) {
        found = true;
        break;
      } else if (dateStr < todayStr) {
        break;
      }
    }
  }

  if (!found) {
    MailApp.sendEmail(recipients.join(","), "Attendance Missing - " + tOutlet, "No attendance data entered for " + todayStr + ".");
  }
}


// Trigger rescheduler
function rescheduleSendSalesReportHTML() {
  // Delete existing triggers for sendSalesReportHTML
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendSalesReportHTML") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  var now = new Date();
  var nextTrigger = new Date(now.getTime());

  // If after 8:55 PM → tomorrow 7:55 AM
  if (now.getHours() > 20 || (now.getHours() === 20 && now.getMinutes() > 55)) {
    nextTrigger.setDate(now.getDate() + 1);
    nextTrigger.setHours(7);
    nextTrigger.setMinutes(55);
    nextTrigger.setSeconds(0);
  }
  // If before 7 AM → today 7:55 AM
  else if (now.getHours() < 7) {
    nextTrigger.setHours(7);
    nextTrigger.setMinutes(55);
    nextTrigger.setSeconds(0);
  }
  // Otherwise → next hour’s :55
  else {
    nextTrigger.setHours(now.getHours() + 2);
    nextTrigger.setMinutes(55);
    nextTrigger.setSeconds(0);
  }

  // Create the next one-time trigger
  ScriptApp.newTrigger("sendSalesReportHTML")
    .timeBased()
    .at(nextTrigger)
    .create();

  Logger.log("Next trigger scheduled for: " + nextTrigger);
}
