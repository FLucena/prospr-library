/**
 * PROSPR Authentication Module
 * Handles admin authentication and session management
 */

/**
 * Check if admin is authenticated
 */
function isAdminAuthenticated() {
  const properties = PropertiesService.getScriptProperties();
  const authStatus = properties.getProperty('admin_authenticated');
  const authTime = properties.getProperty('admin_auth_time');
  
  if (authStatus !== 'true') {
    return false;
  }
  
  // Check session timeout
  if (authTime) {
    const authTimestamp = parseInt(authTime);
    const currentTime = Date.now();
    const sessionTimeout = getConfig('SESSION_TIMEOUT', 24 * 60 * 60 * 1000);
    
    if (currentTime - authTimestamp > sessionTimeout) {
      // Session expired
      setAdminAuthenticated(false);
      return false;
    }
  }
  
  return true;
}

/**
 * Set admin authentication status
 */
function setAdminAuthenticated(authenticated) {
  const properties = PropertiesService.getScriptProperties();
  
  if (authenticated) {
    properties.setProperty('admin_authenticated', 'true');
    properties.setProperty('admin_auth_time', Date.now().toString());
  } else {
    properties.deleteProperty('admin_authenticated');
    properties.deleteProperty('admin_auth_time');
  }
}

/**
 * Show admin authentication dialog
 */
function showAdminAuthDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Admin Authentication',
    'Please enter the admin code:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const enteredCode = result.getResponseText();
    
    if (validateAdminCode(enteredCode)) {
      setAdminAuthenticated(true);
      ui.alert('Success', getConfig('UI.MESSAGES.AUTH_SUCCESS', 'Admin access granted! Please try again.'), ui.ButtonSet.OK);
      // Refresh the UI to show admin menu
      onOpen();
    } else {
      ui.alert('Error', getConfig('UI.MESSAGES.AUTH_ERROR', 'Invalid admin code. Please try again.'), ui.ButtonSet.OK);
    }
  }
}

/**
 * Validate admin code
 */
function validateAdminCode(enteredCode) {
  const expectedCode = ADMIN_CODE;
  return enteredCode === expectedCode;
}

/**
 * Reset admin access
 */
function resetAdminAccess() {
  setAdminAuthenticated(false);
  const ui = SpreadsheetApp.getUi();
  ui.alert('Success', 'Admin access has been reset.', ui.ButtonSet.OK);
  onOpen();
}

/**
 * Get admin session info
 */
function getAdminSessionInfo() {
  const properties = PropertiesService.getScriptProperties();
  const authTime = properties.getProperty('admin_auth_time');
  
  if (!authTime) {
    return null;
  }
  
  const authTimestamp = parseInt(authTime);
  const currentTime = Date.now();
  const sessionTimeout = getConfig('SESSION_TIMEOUT', 24 * 60 * 60 * 1000);
  const remainingTime = sessionTimeout - (currentTime - authTimestamp);
  
  return {
    isAuthenticated: isAdminAuthenticated(),
    authTime: new Date(authTimestamp),
    remainingTime: Math.max(0, remainingTime),
    sessionTimeout: sessionTimeout
  };
}