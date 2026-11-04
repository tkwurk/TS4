/**
 * UNIFIED AUTHENTICATION AND DATA ACCESS SYSTEM
 *
 * @deprecated Project Snowbee A1 — use cleanLoginBackend.js + login.js instead.
 * Do not call from new UI code. Scheduled for removal in Phase G3.
 * Session keys here (auth_session_*) conflict with canonical session_* keys.
 */

// ============================
// SESSION MANAGEMENT
// ============================

/**
 * Creates a new user session after successful login
 * Called by the login system when authentication succeeds
 */
function createUserSession(username, outlet, allowedTabs, deviceId) {
  try {
    if (!deviceId || !username) {
      throw new Error('Missing required parameters for session creation');
    }
    
    // Validate device ID format for security
    const expectedUserSuffix = '_user_' + username.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    if (!deviceId.endsWith(expectedUserSuffix)) {
      throw new Error('Invalid device ID format for user: ' + username);
    }
    
    // Clear any existing sessions for this user/device first
    clearUserSession(username, deviceId);
    
    // Create session data
    const sessionData = {
      username: username,
      deviceId: deviceId,
      outlet: outlet,
      allowedTabs: allowedTabs || [],
      loginTime: new Date().getTime(),
      lastActivity: new Date().getTime(),
      sessionId: Utilities.getUuid() // Add unique session identifier
    };
    
    // Store session in PropertiesService
    const scriptProps = PropertiesService.getScriptProperties();
    const sessionKey = 'auth_session_' + username + '_' + deviceId;
    
    scriptProps.setProperty(sessionKey, JSON.stringify(sessionData));
    
    Logger.log("SESSION CREATED: User " + username + " logged into outlet " + outlet + " with device " + deviceId);
    
    return {
      success: true,
      sessionId: sessionData.sessionId,
      message: "Session created successfully"
    };
    
  } catch (error) {
    Logger.log("ERROR creating session: " + error.toString());
    throw error;
  }
}

/**
 * Validates and retrieves a user session
 * Returns null if session is invalid, expired, or doesn't exist
 */
function validateUserSession(username, deviceId) {
  try {
    if (!username || !deviceId) {
      Logger.log("SESSION VALIDATION: Missing username or deviceId");
      return null;
    }
    
    // Validate device ID format
    const expectedUserSuffix = '_user_' + username.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    if (!deviceId.endsWith(expectedUserSuffix)) {
      Logger.log("SESSION VALIDATION: Invalid device ID format. Expected suffix: " + expectedUserSuffix);
      return null;
    }
    
    const scriptProps = PropertiesService.getScriptProperties();
    const sessionKey = 'auth_session_' + username + '_' + deviceId;
    const sessionJson = scriptProps.getProperty(sessionKey);
    
    if (!sessionJson) {
      Logger.log("SESSION VALIDATION: No session found for key: " + sessionKey);
      return null;
    }
    
    const sessionData = JSON.parse(sessionJson);
    
    // Validate session data integrity
    if (sessionData.username !== username || sessionData.deviceId !== deviceId) {
      Logger.log("SESSION VALIDATION: Session data mismatch");
      clearUserSession(username, deviceId);
      return null;
    }
    
    // Check session expiry (24 hours)
    const now = new Date().getTime();
    const sessionAge = now - sessionData.lastActivity;
    const maxAge = 24 * 60 * 60 * 1000; // 24 hours
    
    if (sessionAge > maxAge) {
      Logger.log("SESSION VALIDATION: Session expired (age: " + sessionAge + "ms)");
      clearUserSession(username, deviceId);
      return null;
    }
    
    // Update last activity
    sessionData.lastActivity = now;
    scriptProps.setProperty(sessionKey, JSON.stringify(sessionData));
    
    Logger.log("SESSION VALIDATION: Valid session for user " + username);
    
    return {
      username: sessionData.username,
      outlet: sessionData.outlet,
      allowedTabs: sessionData.allowedTabs,
      deviceId: sessionData.deviceId,
      sessionId: sessionData.sessionId,
      loginTime: sessionData.loginTime
    };
    
  } catch (error) {
    Logger.log("ERROR validating session: " + error.toString());
    return null;
  }
}

/**
 * Clears a user session
 */
function clearUserSession(username, deviceId) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const sessionKey = 'auth_session_' + username + '_' + deviceId;
    scriptProps.deleteProperty(sessionKey);
    Logger.log("SESSION CLEARED: " + sessionKey);
    return true;
  } catch (error) {
    Logger.log("ERROR clearing session: " + error.toString());
    return false;
  }
}

// ============================
// AUTHENTICATION CONTEXT
// ============================

/**
 * Validates authentication context from client calls
 * This is the single point of authentication for all server functions
 */
function authenticateRequest(authContext) {
  try {
    // Basic parameter validation
    if (!authContext) {
      throw new Error('AUTHENTICATION FAILED: No authentication context provided');
    }
    
    if (!authContext.username || !authContext.deviceId || !authContext.timestamp) {
      throw new Error('AUTHENTICATION FAILED: Incomplete authentication context - missing username, deviceId, or timestamp');
    }
    
    // Check request timestamp to prevent replay attacks
    const now = new Date().getTime();
    const requestAge = now - authContext.timestamp;
    const maxRequestAge = 5 * 60 * 1000; // 5 minutes
    
    if (requestAge > maxRequestAge) {
      throw new Error('AUTHENTICATION FAILED: Request too old (age: ' + requestAge + 'ms)');
    }
    
    if (requestAge < -30000) { // Reject requests from future (clock skew tolerance: 30 seconds)
      throw new Error('AUTHENTICATION FAILED: Request timestamp from future');
    }
    
    // Validate session
    const session = validateUserSession(authContext.username, authContext.deviceId);
    if (!session) {
      throw new Error('AUTHENTICATION FAILED: Invalid or expired session');
    }
    
    Logger.log("AUTHENTICATION SUCCESS: User " + session.username + " authenticated for " + (authContext.functionName || 'unknown function'));
    
    return {
      isAuthenticated: true,
      username: session.username,
      outlet: session.outlet,
      allowedTabs: session.allowedTabs,
      deviceId: session.deviceId,
      sessionId: session.sessionId
    };
    
  } catch (error) {
    Logger.log("AUTHENTICATION ERROR: " + error.toString());
    throw error;
  }
}

// ============================
// SECURE DATA ACCESS
// ============================

/**
 * Secure sheet access wrapper
 * Only allows access to sheets after authentication
 */
function getSecureSheetData(sheetName, authContext) {
  // Authenticate first
  const authResult = authenticateRequest(authContext);
  
  try {
    Logger.log("SECURE SHEET ACCESS: User " + authResult.username + " accessing " + sheetName);
    
    // Get sheet data using existing sheet cache system
    const sheet = getCachedSheet(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log("SECURE SHEET ACCESS: Retrieved " + data.length + " rows from " + sheetName);
    
    return {
      data: data,
      userContext: authResult
    };
    
  } catch (error) {
    Logger.log("ERROR in secure sheet access: " + error.toString());
    throw error;
  }
}

// ============================
// PRODUCT MANAGEMENT
// ============================

/**
 * NEW: Secure product loading function
 * Completely rewritten to ensure authentication before any sheet access
 */
function loadProducts(params) {
  try {
    Logger.log("=== LOADING PRODUCTS ===");
    
    // Validate authentication context
    if (!params || !params._authContext) {
      throw new Error('Authentication required: No authentication context provided');
    }
    
    // Authenticate the request
    const authResult = authenticateRequest(params._authContext);
    
    // Get type filter (default to 'Sell')
    const typeFilter = params.type || 'Sell';
    
    Logger.log("PRODUCTS: Loading " + typeFilter + " products for user " + authResult.username + " (outlet: " + authResult.outlet + ")");
    
    // Access products sheet securely
    const sheetResult = getSecureSheetData('Products', params._authContext);
    const data = sheetResult.data;
    
    if (data.length < 2) {
      Logger.log("PRODUCTS: No product data found");
      return JSON.stringify([]);
    }
    
    const headers = data[0].map(function(h) {
      return (h !== null && h !== undefined) ? h.toString().trim() : '';
    });
    const activeIdx = headers.findIndex(function(h) { return h.toUpperCase() === 'ACTIVE'; });
    const rows = data.slice(1);
    
    // Find relevant column indices
    const outletColumnIndex = headers.findIndex(h => 
      h && h.toString().toLowerCase().includes('outlet')
    );
    const typeColumnIndex = headers.findIndex(h => 
      h && (h.toString().toLowerCase() === 'type' || h.toString().toLowerCase().includes('type'))
    );
    
    Logger.log("PRODUCTS: Processing " + rows.length + " rows. Outlet column: " + outletColumnIndex + ", Type column: " + typeColumnIndex);
    
    // Filter products by outlet and type
    const filteredProducts = rows
      .filter(row => {
        // Skip empty rows
        if (!row || row.length === 0 || !row[0]) {
          return false;
        }

        if (typeof isProductActive_ === 'function' && !isProductActive_(activeIdx >= 0 ? row[activeIdx] : 'Y')) {
          return false;
        }
        
        // Type filtering
        const productType = (typeColumnIndex >= 0 && row[typeColumnIndex]) ? 
          row[typeColumnIndex].toString().trim() : 'Sell';
          
        if (productType !== typeFilter) {
          return false;
        }
        
        // Outlet filtering - include if:
        // 1. No outlet column exists (universal products)
        // 2. Product outlet is empty (universal)
        // 3. Product outlet is 'ALL' (universal)
        // 4. Product outlet matches user's outlet
        if (outletColumnIndex >= 0) {
          const productOutlet = (row[outletColumnIndex] || '').toString().trim();
          const isUniversal = !productOutlet || productOutlet.toUpperCase() === 'ALL';
          const isOutletMatch = productOutlet === authResult.outlet;
          
          if (!isUniversal && !isOutletMatch) {
            return false;
          }
        }
        
        return true;
      })
      .map(row => {
        // Convert row to product object
        const product = {};
        headers.forEach((header, index) => {
          product[header] = row[index];
        });
        
        // Ensure type is correctly set
        product.Type = product.Type || typeFilter;
        
        // Set image property if exists
        if (product.ImageFileId) {
          product.image = product.ImageFileId;
        }
        
        // Convert numeric fields
        const numericFields = ['Price', 'BuyPrice', 'Buy Price', 'Cost', 'UnitPrice', 'Unit Price'];
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
        
        return product;
      });
    
    Logger.log("PRODUCTS: Filtered to " + filteredProducts.length + " products for outlet " + authResult.outlet);
    
    // Log sample for debugging
    if (filteredProducts.length > 0) {
      const sample = filteredProducts[0];
      Logger.log("PRODUCTS: Sample - Name: " + sample.Name + ", Type: " + sample.Type + ", Price: " + sample.Price);
    }
    
    return JSON.stringify(filteredProducts);
    
  } catch (error) {
    Logger.log("ERROR loading products: " + error.toString());
    throw error;
  }
}

/**
 * Load expense accounts securely
 */
function loadExpenseAccounts(params) {
  try {
    Logger.log("=== LOADING EXPENSE ACCOUNTS ===");
    
    const authResult = authenticateRequest(params._authContext);
    const sheetResult = getSecureSheetData('Accounts', params._authContext);
    const data = sheetResult.data;
    
    if (data.length < 2) {
      return [];
    }
    
    const accounts = data.slice(1)
      .filter(row => row[1] && row[1].toString().toLowerCase() === 'expense')
      .map(row => row[0])
      .filter(Boolean);
    
    Logger.log("EXPENSE ACCOUNTS: Returning " + accounts.length + " accounts for user " + authResult.username);
    return accounts;
    
  } catch (error) {
    Logger.log("ERROR loading expense accounts: " + error.toString());
    throw error;
  }
}

/**
 * Load payment sources securely
 */
function loadPaymentSources(params) {
  try {
    Logger.log("=== LOADING PAYMENT SOURCES ===");
    
    const authResult = authenticateRequest(params._authContext);
    const sheetResult = getSecureSheetData('Payment Sources', params._authContext);
    const data = sheetResult.data;
    
    if (data.length < 2) {
      return [];
    }
    
    const sources = data.slice(1)
      .map(row => row[0])
      .filter(Boolean);
    
    Logger.log("PAYMENT SOURCES: Returning " + sources.length + " sources for user " + authResult.username);
    return sources;
    
  } catch (error) {
    Logger.log("ERROR loading payment sources: " + error.toString());
    throw error;
  }
}

// ============================
// LOGIN INTEGRATION
// ============================

/**
 * Enhanced login function that creates session and validates authentication
 * This replaces the existing login function to ensure consistency
 */
function authenticatedLogin(username, code, deviceId) {
  try {
    Logger.log("=== AUTHENTICATED LOGIN ===");
    Logger.log("LOGIN ATTEMPT: User " + username + " with device " + deviceId);
    
    // First validate the user credentials using existing validation
    const loginResult = validateLoginWithAllowedTabs(username, code);
    
    if (!loginResult.success) {
      throw new Error('Invalid credentials: ' + loginResult.error);
    }
    
    // Create session with validated data
    const sessionResult = createUserSession(
      username, 
      loginResult.outlet, 
      loginResult.allowedTabs, 
      deviceId
    );
    
    if (!sessionResult.success) {
      throw new Error('Failed to create session: ' + sessionResult.message);
    }
    
    Logger.log("LOGIN SUCCESS: User " + username + " authenticated, session created");
    
    return {
      success: true,
      username: username,
      outlet: loginResult.outlet,
      allowedTabs: loginResult.allowedTabs,
      deviceId: deviceId,
      sessionId: sessionResult.sessionId,
      message: "Login successful"
    };
    
  } catch (error) {
    Logger.log("LOGIN ERROR: " + error.toString());
    throw error;
  }
}

// ============================
// DEBUGGING AND TESTING
// ============================

/**
 * Test the complete authentication flow
 */
function testAuthenticationFlow(username, deviceId) {
  try {
    Logger.log("=== TESTING AUTHENTICATION FLOW ===");
    
    // Test session validation
    const session = validateUserSession(username, deviceId);
    if (!session) {
      return { success: false, error: "No valid session found" };
    }
    
    // Test authentication context
    const authContext = {
      username: username,
      deviceId: deviceId,
      timestamp: new Date().getTime(),
      functionName: 'testAuthenticationFlow'
    };
    
    const authResult = authenticateRequest(authContext);
    
    // Test product loading
    const productParams = {
      type: 'Sell',
      _authContext: authContext
    };
    
    const products = loadProducts(productParams);
    const productCount = JSON.parse(products).length;
    
    Logger.log("TEST SUCCESS: Authentication flow working, " + productCount + " products loaded");
    
    return {
      success: true,
      session: session,
      authResult: authResult,
      productCount: productCount
    };
    
  } catch (error) {
    Logger.log("TEST ERROR: " + error.toString());
    return { success: false, error: error.toString() };
  }
}
