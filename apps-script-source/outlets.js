function getOutlets() {
  var sheet = getGlobalSpreadsheet().getSheetByName('Outlets'); // Adjust sheet name if needed
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data.shift();
  var idIdx = headers.findIndex(h => /id/i.test(h));
  var nameIdx = headers.findIndex(h => /name/i.test(h));
  return data.map(function(row) {
    return {
      id: idIdx >= 0 ? row[idIdx] : row[0],
      name: nameIdx >= 0 ? row[nameIdx] : row[1]
    };
  }).filter(out => out.id && out.name);
}

// Secure version for SecurityWrapper compatibility
function getOutletsSecure(params) {
  try {
    Logger.log("getOutletsSecure called with params: " + JSON.stringify(params));
    
    // Extract auth context if using SecurityWrapper
    const authContext = params._authContext;
    if (authContext) {
      Logger.log("Auth context: " + JSON.stringify(authContext));
    }
    
    const outlets = getOutlets();
    
    Logger.log("Retrieved " + outlets.length + " outlets");
    
    return {
      success: true,
      data: outlets,
      outlets: outlets // Provide both formats for compatibility
    };
    
  } catch (error) {
    Logger.log("getOutletsSecure error: " + error.toString());
    return {
      success: false,
      error: error.message,
      data: []
    };
  }
}