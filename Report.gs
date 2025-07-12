/**
 * PROSPR Reports Module
 * Handles report generation, formatting, and email functionality
 */

/**
 * Generate monthly comparative report
 */
function generateComparativeReport() {
  // Check admin authentication first
  if (!isAdminAuthenticated()) {
    showAdminAuthDialog();
    return;
  }
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const budgetSheet = spreadsheet.getSheetByName(getConfig('SHEET_NAMES.MONTHLY_BUDGET'));
    
    if (!budgetSheet) {
      throw new Error('Monthly Budget sheet not found');
    }
    
    // Get budget data
    const budgetData = getBudgetData(budgetSheet);
    
    // Validate data
    const validation = validateBudgetData(budgetData);
    if (!validation.isValid) {
      throw new Error('Invalid budget data: ' + validation.errors.join(', '));
    }
    
    // Generate report
    const report = generateReportContent(budgetData);
    
    // Create or update report sheet
    createReportSheet(report);
    
    // Show success message
    showSuccessMessage('Monthly comparative report generated successfully!');
    
  } catch (error) {
    console.error('Error generating report:', error);
    showDetailedErrorDialog(error, 'Report Generation');
  }
}

/**
 * Create or update the report sheet
 */
function createReportSheet(reportData) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = spreadsheet.getSheetByName(getConfig('SHEET_NAMES.COMPARATIVE_REPORT'));
  
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(getConfig('SHEET_NAMES.COMPARATIVE_REPORT'));
  } else {
    reportSheet.clear();
  }
  
  // Validate report data
  if (!reportData || reportData.length === 0) {
    throw new Error('No report data to write');
  }
  
  // Ensure all rows have exactly one column (clean single-column format)
  for (var i = 0; i < reportData.length; i++) {
    if (!reportData[i]) {
      reportData[i] = [''];
    } else if (reportData[i].length > 1) {
      // If there are multiple columns, join them into a single column
      reportData[i] = [reportData[i].join(' ')];
    } else if (reportData[i].length === 0) {
      reportData[i] = [''];
    }
  }
  
  // Write report data as single column
  reportSheet.getRange(1, 1, reportData.length, 1).setValues(reportData);
  
  // Remove any extra columns that might exist
  var lastColumn = reportSheet.getLastColumn();
  if (lastColumn > 1) {
    reportSheet.deleteColumns(2, lastColumn - 1);
  }
  
  // Format the sheet
  formatReportSheet(reportSheet, reportData.length);
}

/**
 * Format the report sheet
 */
function formatReportSheet(sheet, dataLength) {
  // Set column width for single column report
  sheet.setColumnWidth(1, 600);
  
  // Format the title
  const titleRange = sheet.getRange(1, 1);
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(16);
  titleRange.setFontColor('#1a73e8'); // Google Blue
  
  // Format the timestamp
  const timestampRange = sheet.getRange(2, 1);
  timestampRange.setFontSize(10);
  timestampRange.setFontColor('#5f6368'); // Gray
  
  // Format category summaries (bold, larger font)
  for (var i = 1; i <= dataLength; i++) {
    var cellValue = sheet.getRange(i, 1).getValue();
    if (cellValue && typeof cellValue === 'string') {
      // Check if this is a category summary (contains "is over budget" or "is under budget")
      if (cellValue.indexOf(' is ') > 0 && (cellValue.indexOf('over budget') > 0 || cellValue.indexOf('under budget') > 0)) {
        var categoryRange = sheet.getRange(i, 1);
        categoryRange.setFontWeight('bold');
        categoryRange.setFontSize(12);
        
        // Color code based on over/under budget
        if (cellValue.indexOf('over budget') > 0) {
          categoryRange.setFontColor('#d93025'); // Red for over budget
        } else if (cellValue.indexOf('under budget') > 0) {
          categoryRange.setFontColor('#137333'); // Green for under budget
        }
      }
      
      // Check if this is a line item (starts with bullet point)
      if (cellValue.indexOf('  • ') === 0) {
        var itemRange = sheet.getRange(i, 1);
        itemRange.setFontSize(10);
        itemRange.setFontColor('#5f6368'); // Gray for line items
      }
      
      // Check if this is the "no deviations" message
      if (cellValue.indexOf('No significant deviations found') === 0) {
        var messageRange = sheet.getRange(i, 1);
        messageRange.setFontStyle('italic');
        messageRange.setFontColor('#137333'); // Green for good news
      }
    }
  }
  
  // Add subtle borders only to the data area
  if (dataLength > 0) {
    const dataRange = sheet.getRange(1, 1, dataLength, 1);
    dataRange.setBorder(false, false, false, false, false, false);
    
    // Add bottom border to title
    sheet.getRange(1, 1).setBorder(false, false, true, false, false, false);
  }
  
  // Set background color for the title row
  sheet.getRange(1, 1).setBackground('#f8f9fa');
  
  // Freeze the title row
  sheet.setFrozenRows(1);
}

/**
 * Create email draft for report
 */
function sendReportEmail() {
  // Check admin authentication first
  if (!isAdminAuthenticated()) {
    showAdminAuthDialog();
    return;
  }
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheet = spreadsheet.getSheetByName(getConfig('SHEET_NAMES.COMPARATIVE_REPORT'));
    
    if (!reportSheet) {
      throw new Error('Comparative Report sheet not found. Please generate a report first.');
    }
    
    // Get report content
    const reportData = reportSheet.getDataRange().getValues();
    const emailBody = formatReportForEmail(reportData);
    
    // Create email draft with HTML
    const subject = getConfig('EMAIL.SUBJECT_PREFIX', 'PROSPR Monthly Comparative Report') + ' - ' + new Date().toLocaleDateString();
    const htmlBody = formatEmailBody(emailBody);
    
    GmailApp.createDraft('', subject, '', { htmlBody: htmlBody });
    
    showSuccessMessage('Email draft created successfully! Check your Gmail drafts folder.');
    
  } catch (error) {
    console.error('Error creating email draft:', error);
    showDetailedErrorDialog(error, 'Email Draft Creation');
  }
}

/**
 * Format report data for email with HTML styling
 */
function formatReportForEmail(reportData) {
  var emailBody = '';
  
  for (var i = 0; i < reportData.length; i++) {
    var row = reportData[i];
    var cellValue = row[0] || '';
    
    if (cellValue && typeof cellValue === 'string') {
      // Format title
      if (cellValue.indexOf('MONTHLY COMPARATIVE REPORT') === 0) {
        emailBody += '<h1 style="color: #1a73e8; font-size: 24px; margin: 20px 0; text-align: center; font-family: Arial, sans-serif;">' + cellValue + '</h1>';
      }
      // Format timestamp
      else if (cellValue.indexOf('Generated:') === 0) {
        emailBody += '<p style="color: #5f6368; font-size: 12px; text-align: center; margin: 10px 0 30px 0; font-family: Arial, sans-serif;">' + cellValue + '</p>';
      }
      // Format category summaries (over/under budget)
      else if (cellValue.indexOf(' is ') > 0 && (cellValue.indexOf('over budget') > 0 || cellValue.indexOf('under budget') > 0)) {
        var color = cellValue.indexOf('over budget') > 0 ? '#d93025' : '#137333';
        emailBody += '<h3 style="color: ' + color + '; font-size: 16px; margin: 20px 0 10px 0; font-family: Arial, sans-serif; font-weight: bold;">' + cellValue + '</h3>';
      }
      // Format line items (bullet points)
      else if (cellValue.indexOf('  • ') === 0) {
        emailBody += '<p style="color: #5f6368; font-size: 14px; margin: 5px 0 5px 20px; font-family: Arial, sans-serif;">' + cellValue + '</p>';
      }
      // Format "no deviations" message
      else if (cellValue.indexOf('No significant deviations found') === 0) {
        emailBody += '<p style="color: #137333; font-size: 16px; margin: 20px 0; font-family: Arial, sans-serif; font-style: italic; text-align: center;">' + cellValue + '</p>';
      }
      // Format empty lines
      else if (cellValue.trim() === '') {
        emailBody += '<br>';
      }
      // Default formatting for other content
      else {
        emailBody += '<p style="color: #202124; font-size: 14px; margin: 5px 0; font-family: Arial, sans-serif;">' + cellValue + '</p>';
      }
    }
  }
  
  return emailBody;
}

/**
 * Format email body with professional HTML template
 */
function formatEmailBody(reportContent) {
  const template = getConfig('EMAIL.TEMPLATE', {});
  
  var greeting = template.GREETING || 'Dear Client,';
  var bodyText = template.BODY || 'Please find below the monthly comparative report for your financial plan. This report highlights any significant deviations from your budget and provides detailed insights into your spending patterns.';
  var signature = template.SIGNATURE || 'Best regards,<br>PROSPR Team';
  
  var htmlBody = 
    '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #ffffff;">' +
      '<!-- Header -->' +
      '<div style="background: linear-gradient(135deg, #1a73e8 0%, #4285f4 100%); padding: 30px; text-align: center; border-radius: 8px 8px 0 0;">' +
        '<h1 style="color: white; margin: 0; font-size: 28px; font-weight: 300;">PROSPR Financial Planning</h1>' +
        '<p style="color: #e8f0fe; margin: 10px 0 0 0; font-size: 16px;">Monthly Comparative Report</p>' +
      '</div>' +
      
      '<!-- Content -->' +
      '<div style="padding: 30px; background-color: #ffffff; border: 1px solid #e8eaed; border-top: none;">' +
        '<!-- Greeting -->' +
        '<p style="color: #202124; font-size: 16px; margin: 0 0 20px 0; line-height: 1.5;">' +
          greeting +
        '</p>' +
        
        '<!-- Introduction -->' +
        '<p style="color: #5f6368; font-size: 14px; margin: 0 0 30px 0; line-height: 1.6;">' +
          bodyText +
        '</p>' +
        
        '<!-- Report Content -->' +
        '<div style="background-color: #f8f9fa; border-radius: 8px; padding: 25px; margin: 20px 0; border-left: 4px solid #1a73e8;">' +
          reportContent +
        '</div>' +
        
        '<!-- Summary Box -->' +
        '<div style="background-color: #e8f5e8; border-radius: 8px; padding: 20px; margin: 25px 0; border-left: 4px solid #137333;">' +
          '<h3 style="color: #137333; margin: 0 0 10px 0; font-size: 16px;">Report Summary</h3>' +
          '<p style="color: #5f6368; margin: 0; font-size: 14px; line-height: 1.5;">' +
            'This report identifies categories with significant budget deviations (15% or more). ' +
            'Categories in <span style="color: #d93025; font-weight: bold;">red</span> indicate overspending, ' +
            'while those in <span style="color: #137333; font-weight: bold;">green</span> show underspending.' +
          '</p>' +
        '</div>' +
        
        '<!-- Signature -->' +
        '<div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e8eaed;">' +
          '<p style="color: #202124; font-size: 14px; margin: 0; line-height: 1.5;">' +
            signature +
          '</p>' +
        '</div>' +
      '</div>' +
      
      '<!-- Footer -->' +
      '<div style="background-color: #f8f9fa; padding: 20px; text-align: center; border-radius: 0 0 8px 8px; border: 1px solid #e8eaed; border-top: none;">' +
        '<p style="color: #5f6368; font-size: 12px; margin: 0;">' +
          'This report was generated automatically by PROSPR Financial Planning System.<br>' +
          'For questions or support, please contact your financial advisor.' +
        '</p>' +
      '</div>' +
    '</div>';
  
  return htmlBody;
}