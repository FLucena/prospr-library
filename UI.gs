function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const adminMenu = ui.createMenu(getConfig('UI.MENU_NAMES.ADMIN', 'Admin'))
    .addItem('Monthly Comparative Report', 'generateComparativeReport')
    .addItem('Create Email Draft', 'sendReportEmail')
    .addSeparator()
    .addItem('Reset Admin Access', 'resetAdminAccess')
    .addItem('Session Info', 'showSessionInfo')
    .addSeparator()
    .addItem('Help', 'showHelpDialog');
  
  adminMenu.addToUi();
}

/**
 * Show session information dialog
 */
function showSessionInfo() {
  const sessionInfo = getAdminSessionInfo();
  const ui = SpreadsheetApp.getUi();
  
  if (!sessionInfo) {
    ui.alert('Session Info', 'No active admin session found.', ui.ButtonSet.OK);
    return;
  }
  
  const remainingHours = Math.floor(sessionInfo.remainingTime / (1000 * 60 * 60));
  const remainingMinutes = Math.floor((sessionInfo.remainingTime % (1000 * 60 * 60)) / (1000 * 60));
  
  const message = 'Session Information:\n\n' +
    'Status: ' + (sessionInfo.isAuthenticated ? 'Active' : 'Inactive') + '\n' +
    'Login Time: ' + sessionInfo.authTime.toLocaleString() + '\n' +
    'Remaining Time: ' + remainingHours + 'h ' + remainingMinutes + 'm\n' +
    'Session Timeout: ' + Math.floor(sessionInfo.sessionTimeout / (1000 * 60 * 60)) + ' hours';
  
  ui.alert('Admin Session Info', message, ui.ButtonSet.OK);
}

/**
 * Show success message
 */
function showSuccessMessage(message) {
  const ui = SpreadsheetApp.getUi();
  var displayMessage = message;
  if (!displayMessage) {
    try {
      displayMessage = getConfig('UI.MESSAGES.SUCCESS', 'Operation completed successfully!');
    } catch (e) {
      displayMessage = 'Operation completed successfully!';
    }
  }
  ui.alert('Success', displayMessage, ui.ButtonSet.OK);
}

/**
 * Show error message
 */
function showErrorMessage(message) {
  const ui = SpreadsheetApp.getUi();
  var displayMessage = message;
  if (!displayMessage) {
    try {
      displayMessage = getConfig('UI.MESSAGES.ERROR', 'An error occurred. Please try again.');
    } catch (e) {
      displayMessage = 'An error occurred. Please try again.';
    }
  }
  ui.alert('Error', displayMessage, ui.ButtonSet.OK);
}

/**
 * Show confirmation dialog
 */
function showConfirmationDialog(title, message) {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(title, message, ui.ButtonSet.YES_NO);
  return result === ui.Button.YES;
}

/**
 * Show detailed error dialog
 */
function showDetailedErrorDialog(error, context) {
  if (context === undefined) {
    context = '';
  }
  const ui = SpreadsheetApp.getUi();
  const message = 'Error Details:\n\n' +
    'Context: ' + context + '\n' +
    'Error: ' + error.message + '\n\n' +
    'Please check the Apps Script logs for more details.';
  
  ui.alert('Detailed Error', message, ui.ButtonSet.OK);
}

/**
 * Show help dialog
 */
function showHelpDialog() {
  const ui = SpreadsheetApp.getUi();
  const helpText =
    '• Monthly Comparative Report: Generate a detailed report comparing actual vs. planned totals for each main budget category. Only available to authenticated admins.\n' +
    '• Create Email Draft: Create a Gmail draft with the latest comparative report, ready to send to a client. Only available to authenticated admins.\n' +
    '• Reset Admin Access: Log out of admin mode and require re-authentication for admin features.\n' +
    '• Session Info: View your current admin session status and remaining time.\n' +
    'For admin access, use the admin code configured in the system.';
  
  ui.alert('PROSPR Admin Menu Help', helpText, ui.ButtonSet.OK);
}