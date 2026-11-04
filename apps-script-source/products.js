function isProductActive_(activeVal) {
  if (activeVal === undefined || activeVal === null || activeVal === '') return true;
  var v = activeVal.toString().trim().toUpperCase();
  return v !== 'N' && v !== 'NO' && v !== 'FALSE' && v !== '0';
}

function ensureProductActiveColumn() {
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
  if (!sheet || sheet.getLastRow() === 0) return;
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < h.length; i++) {
    if (h[i] && h[i].toString().trim().toUpperCase() === 'ACTIVE') return;
  }
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue('ACTIVE');
}

function getProducts(typeFilter, userOutlet) {
  try {
    Logger.log('getProducts called with typeFilter: ' + typeFilter + ', userOutlet: ' + userOutlet);
    
    const spreadsheet = getGlobalSpreadsheet();
    Logger.log('Got spreadsheet: ' + (spreadsheet ? 'yes' : 'no'));
    
    const sheet = spreadsheet.getSheetByName(SHEETS.PRODUCTS);
    Logger.log('Got sheet "' + SHEETS.PRODUCTS + '": ' + (sheet ? 'yes' : 'no'));
    
    if (!sheet) {
      Logger.log('ERROR: Products sheet not found. Available sheets: ' + spreadsheet.getSheets().map(s => s.getName()).join(', '));
      return JSON.stringify([]);
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log('Data rows: ' + data.length);
    
    if (data.length === 0) {
      Logger.log('ERROR: No data in Products sheet');
      return JSON.stringify([]);
    }
    
    const headers = data[0].map(function(h) {
      return (h !== null && h !== undefined) ? h.toString().trim() : '';
    });
    const activeIdx = headers.findIndex(function(h) { return h.toUpperCase() === 'ACTIVE'; });
    Logger.log('Headers: ' + JSON.stringify(headers));
    
    const products = [];
    
    Logger.log('Starting to process ' + (data.length - 1) + ' product rows...');
    
    for (let i = 1; i < data.length; i++) {
      try {
        const row = data[i];

        if (!isProductActive_(activeIdx >= 0 ? row[activeIdx] : 'Y')) {
          continue;
        }

        const product = {};
        headers.forEach((header, index) => {
          product[header] = row[index];
        });
        
        // Log first few products to see structure
        if (i <= 3) {
          Logger.log('Product ' + i + ' raw data: ' + JSON.stringify(product));
        }

        // Clean and set default type to 'Sell' if not specified
        const originalType = product.TYPE;
        product.TYPE = (product.TYPE || 'Sell').toString().trim();
        
        // Set image property
        product.image = product.IMAGEFILEID;
        
        // Ensure numeric fields are properly handled
      const numericFields = ['PRICE', 'BUYPRICE', 'Cost', 'UnitPrice', 'Unit Price'];
      numericFields.forEach(field => {
        if (product.hasOwnProperty(field)) {
          const value = product[field];
          if (value !== null && value !== undefined && value !== '') {
            const numericValue = Number(value);
            product[field] = isNaN(numericValue) ? 0 : numericValue;
          } else {
            product[field] = 0;
          }
        }
      });
      
      // Also check for common buy price field variations and standardize
      const buyPriceFields = ['BUYPRICE', 'BuyPrice', 'Buy Price', 'buyPrice', 'Cost', 'cost'];
      let buyPrice = 0;
      for (const field of buyPriceFields) {
        if (product.hasOwnProperty(field) && product[field] !== null && product[field] !== undefined && product[field] !== '') {
          buyPrice = Number(product[field]) || 0;
          break;
        }
      }
      product.BUYPRICE = buyPrice; // Standardize to BuyPrice
      
      // Log product for debugging (especially products with buy price issues)
      if (!product.NAME || product.NAME.toLowerCase().includes('ayam') || buyPrice > 0) {
        Logger.log('Product: ' + product.NAME + ', Original Type: "' + originalType + '", Final Type: "' + product.TYPE + '", BuyPrice: ' + buyPrice);
      }
      
      products.push(product);
      } catch (rowError) {
        Logger.log('Error processing product row ' + i + ': ' + rowError.toString());
        // Continue with next row
      }
    }
    
    Logger.log('Successfully processed ' + products.length + ' products out of ' + (data.length - 1) + ' rows');
    
    // Filter products based on type if specified
    let filteredProducts = products;
    if (typeFilter) {
      const cleanTypeFilter = typeFilter.toString().trim();
      filteredProducts = products.filter(p => {
        const productType = (p.TYPE || '').toString().trim();
        const matches = productType === cleanTypeFilter;
        
        // Log filtering for AYAM products specifically
        if (p.NAME && p.NAME.toLowerCase().includes('ayam')) {
          Logger.log('AYAM Filtering - Product: ' + p.NAME + ', Type: "' + productType + '", Filter: "' + cleanTypeFilter + '", Matches: ' + matches);
        }
        
        return matches;
      });
    }
    
    // Filter products based on outlet if specified
    if (userOutlet && userOutlet.trim() !== '') {
      const cleanUserOutlet = userOutlet.toString().trim();
      Logger.log('Filtering products by outlet: ' + cleanUserOutlet);
      
      filteredProducts = filteredProducts.filter(p => {
        const productOutlet = (p.OUTLET || '').toString().trim();
        
        // If outlet column is empty, include the product (default behavior)
        if (!productOutlet) {
          return true;
        }
        
        // If outlet contains "ALL" (case insensitive), include the product
        if (productOutlet.toLowerCase().includes('all')) {
          return true;
        }
        
        // If outlet contains comma-separated values, check if user's outlet is in the list
        const outletList = productOutlet.split(',').map(outlet => outlet.trim().toLowerCase());
        const userOutletLower = cleanUserOutlet.toLowerCase();
        const isIncluded = outletList.includes(userOutletLower);
        
        // Log filtering for debugging
        if (p.NAME && (p.NAME.toLowerCase().includes('ayam') || !isIncluded)) {
          Logger.log('Outlet Filtering - Product: ' + p.NAME + ', Product Outlet: "' + productOutlet + '", User Outlet: "' + cleanUserOutlet + '", Included: ' + isIncluded);
        }
        
        return isIncluded;
      });
      
      Logger.log('Products after outlet filtering: ' + filteredProducts.length);
    }
    
    Logger.log('Total products: ' + products.length + ', Filtered products: ' + filteredProducts.length + ' (filter: "' + (typeFilter || 'none') + '", outlet: "' + (userOutlet || 'none') + '")');
    
    // Log type distribution for debugging
    const typeDistribution = {};
    products.forEach(p => {
      const type = p.TYPE || 'undefined';
      typeDistribution[type] = (typeDistribution[type] || 0) + 1;
    });
    Logger.log('Type distribution: ' + JSON.stringify(typeDistribution));
    
    Logger.log('Total products: ' + products.length + ', Filtered products: ' + filteredProducts.length);
    
    // Log specific AYAM products in final result
    const ayamInResult = filteredProducts.filter(p => p.NAME && p.NAME.toLowerCase().includes('ayam'));
    if (ayamInResult.length > 0) {
      Logger.log('AYAM products in final result: ' + JSON.stringify(ayamInResult.map(p => ({
        name: p.NAME, 
        type: p.TYPE, 
        buyPrice: p.BUYPRICE,
        price: p.PRICE
      }))));
    }
    
    // Log all products with buy prices for debugging
    const productsWithBuyPrice = filteredProducts.filter(p => p.BUYPRICE && p.BUYPRICE > 0);
    if (productsWithBuyPrice.length > 0) {
      Logger.log('Products with buy prices: ' + JSON.stringify(productsWithBuyPrice.map(p => ({
        name: p.NAME,
        buyPrice: p.BUYPRICE,
        type: p.TYPE
      }))));
    } else {
      Logger.log('WARNING: No products found with buy prices > 0');
    }
    
    return JSON.stringify(filteredProducts);
  } catch (error) {
    Logger.log('ERROR in getProducts: ' + error.toString());
    Logger.log('ERROR stack: ' + error.stack);
    Logger.log('ERROR occurred with typeFilter: ' + typeFilter);
    console.error('Error in getProducts:', error);
    return JSON.stringify([]);
  }
}

/**
 * Get products with outlet filtering and category filtering for purchasing interface
 * @param {string} typeFilter - Type of products to filter (e.g., 'Buy', 'Sell')
 * @param {string} categoryPrefix - Category prefix to filter by (e.g., 'Local')
 * @param {string} deviceId - Device ID for authentication
 * @param {string} username - Username for authentication
 * @returns {string} - JSON string of filtered products
 */
function getProductsWithOutletAndCategoryFilter(typeFilter, categoryPrefix, deviceId, username) {
  try {
    Logger.log('getProductsWithOutletAndCategoryFilter called with typeFilter: ' + typeFilter + ', categoryPrefix: ' + categoryPrefix + ', deviceId: ' + deviceId + ', username: ' + username);
    
    // Get user session to determine outlet
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;
    
    Logger.log('User session outlet: ' + userOutlet + ', category prefix: ' + categoryPrefix);
    
    // Get all products with outlet filtering
    const productsJson = getProducts(typeFilter, userOutlet);
    const products = JSON.parse(productsJson);
    
    // Filter by category prefix if specified
    let filteredProducts = products;
    if (categoryPrefix && categoryPrefix.trim() !== '') {
      const cleanCategoryPrefix = categoryPrefix.toString().trim().toLowerCase();
      filteredProducts = products.filter(p => {
        const productCategory = (p.CATEGORY || '').toString().trim().toLowerCase();
        const matches = productCategory.startsWith(cleanCategoryPrefix);
        
        // Log filtering for debugging
        if (p.NAME && (p.NAME.toLowerCase().includes('local') || matches)) {
          Logger.log('Category Filtering - Product: ' + p.NAME + ', Category: "' + productCategory + '", Prefix: "' + cleanCategoryPrefix + '", Matches: ' + matches);
        }
        
        return matches;
      });
      
      Logger.log('Products after category prefix filtering: ' + filteredProducts.length + ' out of ' + products.length);
    }
    
    return JSON.stringify(filteredProducts);
  } catch (error) {
    Logger.log('ERROR in getProductsWithOutletAndCategoryFilter: ' + error.toString());
    Logger.log('ERROR stack: ' + error.stack);
    console.error('Error in getProductsWithOutletAndCategoryFilter:', error);
    return JSON.stringify([]);
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

/**
 * Debug function to inspect the Products sheet structure and buy price data
 * Call this function to troubleshoot buy price issues
 * @returns {Object} - Debug information about the sheet structure and buy prices
 */
function debugProductsBuyPrice() {
  try {
    Logger.log("=== DEBUGGING PRODUCTS BUY PRICE ===");
    
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) {
      return { error: "Products sheet not found" };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    Logger.log("Headers found: " + JSON.stringify(headers));
    
    // Find buy price related columns
    const buyPriceColumns = [];
    headers.forEach((header, index) => {
      const headerStr = header.toString().toLowerCase();
      if (headerStr.includes('buy') || headerStr.includes('cost') || headerStr.includes('purchase')) {
        buyPriceColumns.push({ header: header, index: index });
      }
    });
    
    Logger.log("Buy price related columns: " + JSON.stringify(buyPriceColumns));
    
    // Sample first 5 rows of data
    const sampleData = [];
    for (let i = 1; i <= Math.min(5, data.length - 1); i++) {
      const row = data[i];
      const rowData = {
        rowNumber: i + 1,
        name: row[headers.indexOf('Name')] || 'N/A'
      };
      
      // Add buy price column values
      buyPriceColumns.forEach(col => {
        rowData[col.header] = row[col.index];
      });
      
      sampleData.push(rowData);
    }
    
    Logger.log("Sample data: " + JSON.stringify(sampleData));
    
    // Check for products with valid buy prices
    let productsWithBuyPrice = 0;
    let totalProducts = 0;
    
    for (let i = 1; i < data.length; i++) {
      totalProducts++;
      const row = data[i];
      
      // Check all buy price columns for non-zero values
      let hasBuyPrice = false;
      buyPriceColumns.forEach(col => {
        const value = row[col.index];
        if (value !== null && value !== undefined && value !== '' && Number(value) > 0) {
          hasBuyPrice = true;
        }
      });
      
      if (hasBuyPrice) {
        productsWithBuyPrice++;
      }
    }
    
    const result = {
      success: true,
      totalProducts: totalProducts,
      productsWithBuyPrice: productsWithBuyPrice,
      headers: headers,
      buyPriceColumns: buyPriceColumns,
      sampleData: sampleData
    };
    
    Logger.log("Debug result: " + JSON.stringify(result));
    return result;
    
  } catch (error) {
    Logger.log("Error in debugProductsBuyPrice: " + error.toString());
    return { error: error.toString() };
  }
}

/**
 * Get products filtered by Mode column value with outlet filtering.
 * @param {string} typeFilter - Type of products (e.g., 'Buy', 'Sell')
 * @param {string} modeFilter - Mode value to filter by (e.g., 'N', 'L', 'LO') or comma-separated for multiple
 * @param {string} deviceId - Device ID for authentication
 * @param {string} username - Username for authentication
 * @returns {string} - JSON string of filtered products
 */
function getProductsByMode(typeFilter, modeFilter, deviceId, username) {
  try {
    Logger.log('getProductsByMode called with typeFilter: ' + typeFilter + ', modeFilter: ' + modeFilter);
    
    var session = getUserSessionEnhanced(deviceId, username) || {};
    var userOutlet = session.outlet;
    
    var productsJson = getProducts(typeFilter, userOutlet);
    var products = JSON.parse(productsJson);
    
    if (modeFilter && modeFilter.trim() !== '') {
      var modes = modeFilter.split(',').map(function(m) { return m.trim().toUpperCase(); });
      products = products.filter(function(p) {
        var productMode = (p.MODE || '').toString().trim().toUpperCase();
        return modes.indexOf(productMode) >= 0;
      });
      Logger.log('Products after Mode filtering (' + modeFilter + '): ' + products.length);
    }
    
    return JSON.stringify(products);
  } catch (error) {
    Logger.log('ERROR in getProductsByMode: ' + error.toString());
    return JSON.stringify([]);
  }
}

/**
 * Returns all unique categories for Buy-type products.
 */
function getBuyProductCategories() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    var headers = data[0].map(function(h) { return h.toString().trim().toUpperCase(); });
    var categoryCol = headers.indexOf('CATEGORY');
    var typeCol = headers.indexOf('TYPE');
    var activeCol = headers.indexOf('ACTIVE');
    
    if (categoryCol === -1) return [];
    
    var categories = {};
    for (var i = 1; i < data.length; i++) {
      if (!isProductActive_(activeCol >= 0 ? data[i][activeCol] : 'Y')) continue;
      var type = typeCol >= 0 ? (data[i][typeCol] || '').toString().trim() : '';
      var category = (data[i][categoryCol] || '').toString().trim();
      if (category && (type.toLowerCase() === 'buy' || type === '')) {
        categories[category] = true;
      }
    }
    
    return Object.keys(categories).sort();
  } catch (error) {
    Logger.log('Error in getBuyProductCategories: ' + error.toString());
    return [];
  }
}