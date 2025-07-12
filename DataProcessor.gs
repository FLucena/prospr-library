/**
 * PROSPR Data Processing Module
 * Handles budget data extraction, analysis, and report generation
 */

/**
 * Get budget data from the Monthly Budget sheet
 */
function getBudgetData(sheet) {
  const dataStartRow = getConfig('DATA_PROCESSING.DATA_START_ROW', 4);
  const data = sheet.getDataRange().getValues().slice(dataStartRow);

  const headers = data[0];
  const outputHeaders = ['Category', 'Budget', 'Actual'];
  
  // Find column indices dynamically based on headers
  const columnIndices = findColumnIndices(headers, outputHeaders);
  
  // Initialize organized data structures
  const categoryTotals = {};
  const categoryLineItems = {};
  
    // Process data from bottom to top to find categories for line items
  var currentCategory = '';
  
  for (var i = data.length - 1; i >= 2; i--) {
    var row = data[i];
    var category = row[1] ? row[1].toString().trim() : '';
    var subcategory = row[2] ? row[2].toString().trim() : '';
    
    // If this is a total row (category starts with 'Total')
    if (
      category !== '' &&
      category.toLowerCase().indexOf('total') === 0
    ) {
      
      // Try different approaches to extract category name
      var categoryName = '';
      if (category.toLowerCase().indexOf('total ') === 0) {
        categoryName = category.substring(6).trim(); // Remove "Total " (6 characters)
      } else if (category.toLowerCase().indexOf('total') === 0) {
        categoryName = category.substring(5).trim(); // Remove "Total" (5 characters)
      } else {
        categoryName = category.replace(/^Total\s+/i, '').trim(); // Fallback to regex
      }
      const totalBudget = parseFloat(row[columnIndices.budget]) || 0;
      const totalActual = parseFloat(row[columnIndices.actual]) || 0;
      
      // Store category totals
      categoryTotals[categoryName] = {
        budget: totalBudget,
        actual: totalActual
      };
      
      // Initialize line items array if it doesn't exist
      if (!categoryLineItems[categoryName]) {
        categoryLineItems[categoryName] = [];
      }
      
      // Set current category for line items above this total
      currentCategory = categoryName;
      continue;
    }
    
    // If this is a line item (category empty, subcategory non-empty)
    if (
      category === '' &&
      subcategory !== '' &&
      currentCategory !== ''
    ) {
      const itemBudget = parseFloat(row[columnIndices.budget]) || 0;
      const itemActual = parseFloat(row[columnIndices.actual]) || 0;
      if (itemBudget !== 0 || itemActual !== 0) {
        const lineItem = {
          name: subcategory,
          budget: itemBudget,
          actual: itemActual
        };
        categoryLineItems[currentCategory].push(lineItem);
      }
      continue;
    }
    
    // If this is a category header (category non-empty, subcategory empty, not a total)
    if (
      category !== '' &&
      subcategory === '' &&
      category.toLowerCase().indexOf('total') !== 0 &&
      category !== 'Monthly Budget' &&
      category !== 'Detail View'
    ) {
      // This is a category header, but we're processing from bottom up
      // So we don't need to set currentCategory here
      continue;
    }
  }
  
  // Return structured data with organized objects
  return {
    headers: outputHeaders,
    categoryTotals: categoryTotals,
    categoryLineItems: categoryLineItems,
    rawData: data,
    columnIndices: columnIndices
  };
}

/**
 * Find column indices based on header names
 */
function findColumnIndices(headers, outputHeaders) {
  const categoryColumn = getConfig('DATA_PROCESSING.CATEGORY_COLUMN', 1);
  const nameColumn = getConfig('DATA_PROCESSING.NAME_COLUMN', 2); // Default to column C for names
  const indices = {
    category: categoryColumn, // Default to column B since category is not in headers
    name: nameColumn,         // Default to column C for line item/category/total names
    budget: -1,
    actual: -1,
    ytdBudget: -1,
    ytdActual: -1
  };
  
      // Map outputHeaders to their corresponding indices
    const headerMap = {
      'Budget': 'budget',
      'Actual': 'actual'
    };
  
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i].toString().trim();
    
    // Check if this header matches any of our outputHeaders
    for (var j = 0; j < outputHeaders.length; j++) {
      var newHeader = outputHeaders[j];
      if (header.toLowerCase() === newHeader.toLowerCase()) {
        var key = headerMap[newHeader];
        if (key) {
          indices[key] = i;
        }
      }
    }
  }
  
  // Validate that we found the required columns
  if (indices.budget === -1 || indices.actual === -1) {
    throw new Error('Required columns (Budget, Actual) not found in headers');
  }
  
  return indices;
}

/**
 * Generate report content using the new data structure
 */
function generateReportContent(budgetData) {
  var report = [];
  
  // Add header with proper formatting
  report.push(['MONTHLY COMPARATIVE REPORT']);
  report.push(['Generated: ' + new Date().toLocaleString()]);
  report.push([]);
  
  var hasSignificantDeviations = false;
  const deviationThreshold = getConfig('DEVIATION_THRESHOLD', 0.15);
  
  // Process category totals data in reverse order
  if (budgetData.categoryTotals) {
    // Get category names and reverse the order
    var categoryNames = Object.keys(budgetData.categoryTotals);
    categoryNames.reverse();
    
    for (var i = 0; i < categoryNames.length; i++) {
      var categoryName = categoryNames[i];
      var categoryTotal = budgetData.categoryTotals[categoryName];
      
      const deviation = categoryTotal.actual - categoryTotal.budget;
      const deviationPercent = categoryTotal.budget > 0 ? (deviation / categoryTotal.budget) * 100 : (categoryTotal.actual !== 0 ? (categoryTotal.actual > 0 ? 100 : -100) : 0);
      const isSignificant = Math.abs(deviationPercent) >= (deviationThreshold * 100);
      
      if (isSignificant) {
        hasSignificantDeviations = true;
        
        // Format category summary with proper currency formatting
        var categorySummary = categoryName + ' is ' + (deviationPercent > 0 ? 'over' : 'under') + ' budget by ' + Math.abs(deviationPercent).toFixed(1) + '%';
        if (Math.abs(deviation) > 0) {
          categorySummary += ' ($' + Math.abs(deviation).toFixed(2) + ')';
        }
        categorySummary += '.';
        report.push([categorySummary]);
        
        // Find and show specific line items causing the deviation
        if (budgetData.categoryLineItems && budgetData.categoryLineItems[categoryName]) {
          var categoryItems = budgetData.categoryLineItems[categoryName];
          
          // Show significant line items within this category
          for (var j = 0; j < categoryItems.length; j++) {
            var item = categoryItems[j];
            var itemDeviation = item.actual - item.budget;
            var itemDeviationPercent = item.budget > 0 ? (itemDeviation / item.budget) * 100 : 0;
            
            // Show items with significant deviations or if they contribute to the overall deviation
            if (Math.abs(itemDeviationPercent) >= (deviationThreshold * 100) || Math.abs(itemDeviation) > 0) {
              var itemDetail = '  • ' + item.name + ': $' + item.actual.toFixed(2) + ' (Actual) vs $' + item.budget.toFixed(2) + ' (Planned)';
              if (Math.abs(itemDeviationPercent) > 0) {
                itemDetail += ' (' + (itemDeviationPercent > 0 ? '+' : '') + itemDeviationPercent.toFixed(1) + '%)';
              }
              report.push([itemDetail]);
            }
          }
        }
        
        report.push([]);
      }
    }
  }
  
  if (!hasSignificantDeviations) {
    report.push(['No significant deviations found. All categories are within ' + (deviationThreshold * 100) + '% of budget.']);
  }
  
  return report;
}



/**
 * Validate budget data structure
 */
function validateBudgetData(budgetData) {
  const errors = [];
  const requiredColumns = getConfig('DATA_PROCESSING.REQUIRED_COLUMNS', 5);
  
  if (!budgetData || !budgetData.categoryTotals) {
    errors.push('No budget data found');
    return { isValid: false, errors: errors };
  }
  
  if (!budgetData.headers || budgetData.headers.length === 0) {
    errors.push('No headers found in budget data');
  }
  
  // Validate category totals
  for (var categoryName in budgetData.categoryTotals) {
    var categoryTotal = budgetData.categoryTotals[categoryName];
    
    var budget = parseFloat(categoryTotal.budget);
    var actual = parseFloat(categoryTotal.actual);
    
    if (isNaN(budget) || isNaN(actual)) {
      errors.push('Invalid numeric values in category totals: ' + categoryName);
    }
  }
  
  // Validate line items
  if (budgetData.categoryLineItems) {
    for (var categoryName in budgetData.categoryLineItems) {
      var lineItems = budgetData.categoryLineItems[categoryName];
      
      for (var i = 0; i < lineItems.length; i++) {
        var item = lineItems[i];
        
        var itemBudget = parseFloat(item.budget);
        var itemActual = parseFloat(item.actual);
        
        if (isNaN(itemBudget) || isNaN(itemActual)) {
          errors.push('Invalid numeric values in line item: ' + item.name + ' for category: ' + categoryName);
        }
      }
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}