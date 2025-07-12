/**
 * PROSPR Configuration Module
 * Centralized configuration for all PROSPR functionality
 */

// Global configuration object
var PROSPR_CONFIG = {
  // Security settings
  SESSION_TIMEOUT: 24 * 60 * 60 * 1000, // 24 hours in milliseconds
  
  // Analysis settings
  DEVIATION_THRESHOLD: 0.15, // 15% threshold for significant differences
  
  // Data processing settings
  DATA_PROCESSING: {
    DATA_START_ROW: 4, // Row where data starts (0-indexed)
    CATEGORY_COLUMN: 1, // Default category column (column B)
    INSIGHT_THRESHOLD: 20, // Percentage threshold for generating insights
    REQUIRED_COLUMNS: 5, // Number of required columns in totals data
    DEFAULT_VALUES: {
      STRING: '',
      NUMBER: 0
    }
  },
  
  // Sheet names
  SHEET_NAMES: {
    MONTHLY_BUDGET: 'Monthly Budget',
    COMPARATIVE_REPORT: 'Comparative Report',
    DEPLOYMENT_LOGS: 'Deployment Logs',
    CLIENT_CONFIG: 'Client Configurations'
  },
  
  // Budget categories
  CATEGORIES: [
    'Shelter', 
    'Food', 
    'Transportation', 
    'Utilities', 
    'Entertainment', 
    'Healthcare', 
    'Other'
  ],
  
  // Email settings
  EMAIL: {
    SUBJECT_PREFIX: 'PROSPR Monthly Comparative Report',
    TEMPLATE: {
      GREETING: 'Dear Client,',
      BODY: 'Please find attached the monthly comparative report for your financial plan.',
      SIGNATURE: 'Best regards,\nPROSPR Team'
    }
  },
  
  // UI settings
  UI: {
    MENU_NAMES: {
      MAIN: 'PROSPR Tools',
      ADMIN: 'Admin'
    },
    MESSAGES: {
      SUCCESS: 'Operation completed successfully!',
      ERROR: 'An error occurred. Please try again.',
      AUTH_SUCCESS: 'Admin access granted! Please try again.',
      AUTH_ERROR: 'Invalid admin code. Please try again.'
    }
  },
  
  // Deployment settings
  DEPLOYMENT: {
    BATCH_SIZE: 10, // Number of clients to process in one batch
    RETRY_ATTEMPTS: 3,
    TIMEOUT: 30000 // 30 seconds timeout per client
  }
};

/**
 * Get configuration value with fallback
 */
function getConfig(key, defaultValue) {
  // Set default value if not provided
  if (defaultValue === undefined) {
    defaultValue = null;
  }
  
  const keys = key.split('.');
  var value = PROSPR_CONFIG;
  
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (value && typeof value === 'object' && k in value) {
      value = value[k];
    } else {
      return defaultValue;
    }
  }
  
  return value;
}

/**
 * Set configuration value
 */
function setConfig(key, value) {
  const keys = key.split('.');
  const lastKey = keys.pop();
  var current = PROSPR_CONFIG;
  
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (!(k in current)) {
      current[k] = {};
    }
    current = current[k];
  }
  
  current[lastKey] = value;
}

/**
 * Validate configuration
 */
function validateConfig() {
  const errors = [];
  
  // Check required settings
  if (!PROSPR_CONFIG.ADMIN_CODE || PROSPR_CONFIG.ADMIN_CODE.length < 4) {
    errors.push('ADMIN_CODE must be at least 4 characters long');
  }
  
  if (PROSPR_CONFIG.DEVIATION_THRESHOLD <= 0 || PROSPR_CONFIG.DEVIATION_THRESHOLD > 1) {
    errors.push('DEVIATION_THRESHOLD must be between 0 and 1');
  }
  
  if (!PROSPR_CONFIG.SHEET_NAMES.MONTHLY_BUDGET) {
    errors.push('MONTHLY_BUDGET sheet name is required');
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}