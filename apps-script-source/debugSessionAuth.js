/**
 * Debug authentication issue by testing session creation and retrieval
 * Add this to serverSecurity.js to help diagnose the session timing issue
 */

/**
 * Debug function to test session creation and immediate retrieval
 * Call this to understand what's happening with the session storage
 */
function debugSessionAuthentication(username, deviceId) {
  Logger.log("=== DEBUG SESSION AUTHENTICATION ===");
  Logger.log("Testing with username: " + username + ", deviceId: " + deviceId);
  
  try {
    // 1. Test if getUserSessionEnhanced exists
    if (typeof getUserSessionEnhanced !== 'function') {
      Logger.log("ERROR: getUserSessionEnhanced function not found!");
      return { error: "getUserSessionEnhanced function not found" };
    }
    
    // 2. Check current session
    Logger.log("Checking for existing session...");
    const existingSession = getUserSessionEnhanced(deviceId, username);
    Logger.log("Existing session result: " + JSON.stringify(existingSession));
    
    // 3. Check PropertiesService directly
    const scriptProps = PropertiesService.getScriptProperties();
    const sessionKey = 'session_' + username + '_' + deviceId;
    Logger.log("Looking for session key: " + sessionKey);
    
    const rawSessionData = scriptProps.getProperty(sessionKey);
    Logger.log("Raw session data from PropertiesService: " + rawSessionData);
    
    if (rawSessionData) {
      try {
        const parsedSession = JSON.parse(rawSessionData);
        Logger.log("Parsed session data: " + JSON.stringify(parsedSession));
      } catch (parseError) {
        Logger.log("Error parsing session data: " + parseError.toString());
      }
    }
    
    // 4. List all session-related properties
    const allProperties = scriptProps.getProperties();
    const sessionProperties = {};
    for (const key in allProperties) {
      if (key.includes('session') || key.includes(username)) {
        sessionProperties[key] = allProperties[key];
      }
    }
    Logger.log("All session-related properties: " + JSON.stringify(sessionProperties));
    
    // 5. Test device ID validation
    const expectedUserSuffix = '_user_' + username.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    const deviceIdValid = deviceId.endsWith(expectedUserSuffix);
    Logger.log("Device ID validation - Expected suffix: " + expectedUserSuffix + ", Valid: " + deviceIdValid);
    
    return {
      username: username,
      deviceId: deviceId,
      sessionKey: sessionKey,
      existingSession: existingSession,
      rawSessionData: rawSessionData,
      deviceIdValid: deviceIdValid,
      expectedSuffix: expectedUserSuffix,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log("ERROR in debugSessionAuthentication: " + error.toString());
    return { error: error.toString() };
  }
}

/**
 * Enhanced version of validateAuthenticationContext with detailed logging
 */
function validateAuthenticationContextDebug(authContext) {
  Logger.log("=== DEBUG AUTHENTICATION CONTEXT ===");
  
  try {
    if (!authContext) {
      Logger.log("SECURITY: No authentication context provided");
      throw new Error('Authentication required - no context provided');
    }

    Logger.log("Auth context received: " + JSON.stringify(authContext));

    if (!authContext.deviceId || !authContext.username || !authContext.timestamp) {
      Logger.log("SECURITY: Incomplete authentication context: " + JSON.stringify(authContext));
      throw new Error('Authentication required - incomplete context');
    }

    // Check timestamp to prevent replay attacks (requests older than 5 minutes rejected)
    const now = new Date().getTime();
    const requestAge = now - authContext.timestamp;
    const maxAge = 5 * 60 * 1000; // 5 minutes

    Logger.log("Request timestamp validation - Age: " + requestAge + "ms, Max: " + maxAge + "ms");

    if (requestAge > maxAge) {
      Logger.log("SECURITY: Authentication context too old: " + requestAge + "ms");
      throw new Error('Authentication expired - request too old');
    }

    // Detailed session validation
    Logger.log("Calling getUserSessionEnhanced with deviceId: " + authContext.deviceId + ", username: " + authContext.username);
    
    const session = getUserSessionEnhanced(authContext.deviceId, authContext.username);
    
    Logger.log("getUserSessionEnhanced returned: " + JSON.stringify(session));
    
    if (!session) {
      Logger.log("SECURITY: No valid session found for authentication context");
      
      // Additional debugging
      const debugResult = debugSessionAuthentication(authContext.username, authContext.deviceId);
      Logger.log("Debug session result: " + JSON.stringify(debugResult));
      
      throw new Error('Authentication required - no valid session');
    }

    if (session.username !== authContext.username) {
      Logger.log("SECURITY: Session username mismatch in authentication context - Session: " + session.username + ", Request: " + authContext.username);
      throw new Error('Authentication required - session mismatch');
    }

    Logger.log("SECURITY: Authentication context validated successfully for " + authContext.username + " (" + authContext.functionName + ")");
    return {
      valid: true,
      username: session.username,
      outlet: session.outlet,
      allowedTabs: session.allowedTabs || [],
      deviceId: authContext.deviceId
    };

  } catch (error) {
    Logger.log("SECURITY: Authentication validation failed: " + error.toString());
    throw error;
  }
}

/**
 * Test function to manually test the getProductsSecure flow
 */
function testGetProductsSecureFlow(username, deviceId) {
  Logger.log("=== TESTING getProductsSecure FLOW ===");
  
  try {
    // Create test auth context
    const authContext = {
      deviceId: deviceId,
      username: username,
      timestamp: Date.now(),
      functionName: 'testGetProductsSecure'
    };
    
    // Test the authentication
    const authResult = validateAuthenticationContextDebug(authContext);
    Logger.log("Authentication result: " + JSON.stringify(authResult));
    
    // If authentication passed, test the actual getProductsSecure call
    const testParams = {
      type: 'Sell',
      _authContext: authContext
    };
    
    Logger.log("Calling getProductsSecure with params: " + JSON.stringify(testParams));
    const result = getProductsSecure(testParams);
    Logger.log("getProductsSecure result length: " + (result ? result.length : 'null'));
    
    return {
      success: true,
      authResult: authResult,
      productsCount: result ? JSON.parse(result).length : 0
    };
    
  } catch (error) {
    Logger.log("ERROR in testGetProductsSecureFlow: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
