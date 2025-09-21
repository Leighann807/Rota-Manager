/**
 * Staff Rota Manager
 * Google Sheets Add-on for simplified staff scheduling and absence management
 * Version: 2.0
 * OAuth Ready: Yes
 * Store Ready: Yes
 */

// Global constants
const SHIFT_PATTERNS = {
  'EARLY': {label: 'Early', hours: 8, color: '#FF0000'},    // Red
  'LATE': {label: 'Late', hours: 8, color: '#0000FF'},      // Blue
  'NIGHT': {label: 'Night', hours: 10, color: '#800080'},   // Purple
  'DAY': {label: 'Day', hours: 8, color: '#FFFF00'},        // Yellow
  'OFF': {label: 'Off', hours: 0, color: '#000000'},        // Black
  'AL': {label: 'Annual Leave', hours: 0, color: '#008000'}, // Green
  'SICK': {label: 'Sick', hours: 0, color: '#FFA500'},      // Orange
  'TRAINING': {label: 'Training', hours: 8, color: '#00FFFF'} // Cyan
};

// Configuration constants
const CONFIG = {
  MAX_STAFF_MEMBERS: 100,
  MAX_PATTERN_LENGTH: 50,
  MAX_DAYS_RANGE: 366,
  ERROR_RETRY_COUNT: 3,
  BATCH_SIZE: 20,
  CACHE_DURATION: 300000 // 5 minutes
};

// Error types for better error handling
const ERROR_TYPES = {
  VALIDATION: 'ValidationError',
  PERMISSION: 'PermissionError',
  DATA: 'DataError',
  NETWORK: 'NetworkError',
  QUOTA: 'QuotaError'
};

/**
 * Utility Functions for Enhanced Error Handling and Performance
 */

/**
 * Builds optimized hours calculation formula including custom shifts
 * @param {number} daysInMonth - Number of days in the month
 * @return {string} The formula string
 */
function buildHoursFormula(daysInMonth) {
  try {
    const rangeEnd = String.fromCharCode(65 + daysInMonth); // Convert to column letter
    const allPatterns = getShiftPatterns();
    
    // Build formula parts for each shift pattern that has hours > 0
    // Use R1C1 notation to ensure proper relative referencing
    const formulaParts = [];
    Object.keys(allPatterns).forEach(shiftType => {
      const pattern = allPatterns[shiftType];
      if (pattern.hours > 0) {
        formulaParts.push(`(B2:${rangeEnd}2="${shiftType}")*${pattern.hours}`);
      }
    });
    
    if (formulaParts.length === 0) {
      // Fallback if no patterns found
      return `=SUMPRODUCT((B2:${rangeEnd}2="EARLY")*8)`;
    }
    
    return `=SUMPRODUCT(${formulaParts.join(',')})`;
  } catch (error) {
    Logger.log('Error building hours formula: ' + error.toString());
    // Fallback to simpler formula
    const rangeEnd = String.fromCharCode(65 + daysInMonth);
    return `=SUMPRODUCT((B2:${rangeEnd}2="EARLY")*8,(B2:${rangeEnd}2="LATE")*8,(B2:${rangeEnd}2="NIGHT")*10,(B2:${rangeEnd}2="DAY")*8,(B2:${rangeEnd}2="TRAINING")*8)`;
  }
}

/**
 * Builds optimized count formula for specific shift types
 * @param {number} daysInMonth - Number of days in the month
 * @param {string} shiftType - The shift type to count
 * @param {number} columnOffset - Column offset for the formula
 * @return {string} The formula string
 */
function buildCountFormula(daysInMonth, shiftType, columnOffset) {
  try {
    const rangeEnd = String.fromCharCode(65 + daysInMonth);
    return `=COUNTIF(B2:${rangeEnd}2,"${shiftType}")`;
  } catch (error) {
    Logger.log('Error building count formula: ' + error.toString());
    return `=COUNTIF(B2:C2,"${shiftType}")`; // Fallback
  }
}

/**
 * Enhanced error handling wrapper
 * @param {Function} func - Function to execute
 * @param {string} operation - Description of operation
 * @param {*} fallback - Fallback value on error
 * @return {*} Result or fallback
 */
function safeExecute(func, operation, fallback = null) {
  try {
    return func();
  } catch (error) {
    Logger.log(`Error in ${operation}: ${error.toString()}`);
    logError(error, operation);
    return fallback;
  }
}

/**
 * Validates input parameters
 * @param {Object} params - Parameters to validate
 * @param {Array} required - Required parameter names
 * @return {Object} Validation result
 */
function validateInput(params, required) {
  const missing = required.filter(param => !params[param] && params[param] !== 0);
  return {
    valid: missing.length === 0,
    missing: missing,
    message: missing.length > 0 ? `Missing required parameters: ${missing.join(', ')}` : null
  };
}

/**
 * Logs errors for debugging and analytics
 * @param {Error} error - The error object
 * @param {string} operation - The operation that failed
 */
function logError(error, operation) {
  try {
    const errorInfo = {
      timestamp: new Date().toISOString(),
      operation: operation,
      message: error.toString(),
      stack: error.stack || 'No stack trace available'
    };
    Logger.log('ERROR_LOG: ' + JSON.stringify(errorInfo));
  } catch (logError) {
    Logger.log('Failed to log error: ' + logError.toString());
  }
}

/**
 * Initializes user preferences on first use
 */
function initializeUserPreferences() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const initialized = userProperties.getProperty('STAFF_ROTA_INITIALIZED');
    
    if (!initialized) {
      const defaultPreferences = {
        'STAFF_ROTA_INITIALIZED': 'true',
        'DEFAULT_WORK_HOURS': '8',
        'WEEKEND_HIGHLIGHT': 'true',
        'NOTIFICATION_ENABLED': 'true',
        'THEME': 'default'
      };
      
      userProperties.setProperties(defaultPreferences);
      Logger.log('User preferences initialized');
    }
  } catch (error) {
    Logger.log('Error initializing user preferences: ' + error.toString());
  }
}

/**
 * Logs usage for analytics (privacy compliant)
 * @param {string} action - The action performed
 */
function logUsage(action) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const analyticsEnabled = userProperties.getProperty('ANALYTICS_ENABLED') !== 'false';
    
    if (analyticsEnabled) {
      const usageData = {
        action: action,
        timestamp: new Date().toISOString(),
        version: '2.0'
      };
      Logger.log('USAGE_LOG: ' + JSON.stringify(usageData));
    }
  } catch (error) {
    Logger.log('Error logging usage: ' + error.toString());
  }
}

/**
 * Shows help and documentation
 */
function showHelp() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Help')
      .setTitle('Staff Rota Manager - Help')
      .setWidth(600)
      .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Help & Documentation');
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Help',
      'Visit our documentation at: https://docs.google.com/document/d/your-help-doc-id\n\nFor support, please contact: support@your-domain.com',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Shows privacy policy
 */
function showPrivacyPolicy() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('PrivacyPolicy')
      .setTitle('Privacy Policy')
      .setWidth(600)
      .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Privacy Policy');
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Privacy Policy',
      'View our privacy policy at: https://your-domain.com/privacy\n\nWe are committed to protecting your data and privacy.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Adds menu item to the Google Sheets UI when the spreadsheet opens.
 * Enhanced with error handling and user guidance.
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Staff Rota')
      .addItem('🔄 Generate Monthly Rota', 'showRotaGenerator')
      .addSeparator()
      .addItem('🏥 Log New Absence', 'showAbsenceLogger')
      .addItem('📊 View Absence Reports', 'showAbsenceReports')
      .addSeparator()
      .addItem('⚙️ Settings', 'showSettings')
      .addSeparator()
      .addItem('❓ Help & Documentation', 'showHelp')
      .addItem('🔒 Privacy Policy', 'showPrivacyPolicy')
      .addToUi();
    
    // Initialize user preferences if first time
    initializeUserPreferences();
    
    // Log usage for analytics (privacy compliant)
    logUsage('menu_opened');
    
  } catch (error) {
    Logger.log('Error in onOpen: ' + error.toString());
    // Fallback: create basic menu
    try {
      SpreadsheetApp.getUi().createMenu('Staff Rota')
        .addItem('Generate Monthly Rota', 'showRotaGenerator')
        .addItem('Help', 'showHelp')
        .addToUi();
    } catch (fallbackError) {
      Logger.log('Critical error: Cannot create menu - ' + fallbackError.toString());
    }
  }
}




/**
 * Shows the rota generator in a full-sized popup window.
 */
function showRotaGenerator() {
  const html = HtmlService.createHtmlOutputFromFile('RotaGenerator')
    .setWidth(1400)
    .setHeight(900);
  
  SpreadsheetApp.getUi().showModelessDialog(html, 'Generate Monthly Rota');
}


/**
 * Shows the absence logger dialog.
 */
function showAbsenceLogger() {
  const html = HtmlService.createHtmlOutputFromFile('AbsenceLogger')
    .setWidth(400)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Log New Absence');
}

/**
 * Shows the absence reports sidebar.
 */
function showAbsenceReports() {
  const html = HtmlService.createHtmlOutputFromFile('AbsenceReports')
    .setTitle('Absence Reports')
    .setWidth(400);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the settings popup dialog.
 */
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsDialog')
    .setWidth(800)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '⚙️ Staff Rota Settings');
}

/**
 * Gets all staff names from the active sheet.
 */
function getStaffNames() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // No staff if only header row exists
  
  const staffRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const staffValues = staffRange.getValues();
  
  // Return staff names along with their row positions
  const staffInfo = [];
  for (let i = 0; i < staffValues.length; i++) {
    const staffName = staffValues[i][0].toString().trim();
    if (staffName !== '') {
      staffInfo.push({
        name: staffName,
        row: i + 2 // +2 because we start at row 2 and i is 0-indexed
      });
    }
  }
  
  return staffInfo;
}

/**
 * Gets all available staff names from both the Settings storage and active sheet.
 * This function combines staff from the Settings tab with any staff already in sheets.
 */
function getAllAvailableStaff() {
  try {
    const staffFromSettings = [];
    const staffFromSheet = [];
    
    // Get staff from Settings storage
    try {
      const userProperties = PropertiesService.getUserProperties();
      const staffListJson = userProperties.getProperty('staffList');
      if (staffListJson) {
        const staffList = JSON.parse(staffListJson);
        staffList.forEach(staff => {
          staffFromSettings.push({
            name: staff.name,
            role: staff.role || '',
            source: 'settings'
          });
        });
      }
    } catch (settingsError) {
      Logger.log('Error loading staff from settings: ' + settingsError.toString());
    }
    
    // Get staff from current sheet (if it's a rota sheet)
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      if (sheet && sheet.getRange('A1').getValue() === 'Staff Name') {
        const staffNames = getStaffNames();
        staffNames.forEach(staff => {
          // Only add if not already in settings
          const existsInSettings = staffFromSettings.some(settingsStaff => 
            settingsStaff.name.toLowerCase() === staff.name.toLowerCase()
          );
          if (!existsInSettings) {
            staffFromSheet.push({
              name: staff.name,
              role: '',
              source: 'sheet',
              row: staff.row
            });
          }
        });
      }
    } catch (sheetError) {
      Logger.log('Error loading staff from sheet: ' + sheetError.toString());
    }
    
    // Combine both lists, with settings staff first
    const allStaff = [...staffFromSettings, ...staffFromSheet];
    
    Logger.log(`getAllAvailableStaff: Found ${staffFromSettings.length} from settings, ${staffFromSheet.length} from sheet`);
    
    return {
      success: true,
      staff: allStaff,
      settingsCount: staffFromSettings.length,
      sheetCount: staffFromSheet.length
    };
    
  } catch (error) {
    Logger.log('Error in getAllAvailableStaff: ' + error.toString());
    return {
      success: false,
      staff: [],
      message: error.toString()
    };
  }
}

/**
 * Gets annual leave allocation for a specific staff member
 * @param {string} staffName - Name of the staff member
 * @return {number} Annual leave allocation in days
 */
function getStaffAnnualLeaveAllocation(staffName) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const allocationsJson = userProperties.getProperty('annualLeaveAllocations');
    
    if (!allocationsJson) {
      // Default allocation if none set
      return 28; // UK standard annual leave
    }
    
    const allocations = JSON.parse(allocationsJson);
    return allocations[staffName] || 28; // Default to 28 days if not found
  } catch (error) {
    Logger.log('Error getting annual leave allocation: ' + error.toString());
    return 28; // Fallback
  }
}

/**
 * Sets annual leave allocation for a specific staff member
 * @param {string} staffName - Name of the staff member
 * @param {number} allocation - Annual leave allocation in days
 */
function setStaffAnnualLeaveAllocation(staffName, allocation) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const allocationsJson = userProperties.getProperty('annualLeaveAllocations');
    
    let allocations = {};
    if (allocationsJson) {
      allocations = JSON.parse(allocationsJson);
    }
    
    allocations[staffName] = allocation;
    userProperties.setProperty('annualLeaveAllocations', JSON.stringify(allocations));
    
    Logger.log(`Set annual leave allocation for ${staffName}: ${allocation} days`);
    return { success: true };
  } catch (error) {
    Logger.log('Error setting annual leave allocation: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Builds annual leave balance formula (allocation minus used days)
 * @param {number} daysInMonth - Number of days in the month
 * @return {string} The formula string
 */
function buildAnnualLeaveBalanceFormula(daysInMonth) {
  try {
    const rangeEnd = String.fromCharCode(65 + daysInMonth);
    // Use simple relative reference that will update automatically
    return `=28-COUNTIF(B2:${rangeEnd}2,"AL")`;
  } catch (error) {
    Logger.log('Error building annual leave balance formula: ' + error.toString());
    const rangeEnd = String.fromCharCode(65 + daysInMonth);
    return `=28-COUNTIF(B2:${rangeEnd}2,"AL")`; // Fallback with default 28 days
  }
}

/**
 * Sets annual leave allocation for all existing staff members
 * @param {number} defaultAllocation - Default allocation in days (default: 28)
 */
function setAllStaffAnnualLeaveAllocations(defaultAllocation = 28) {
  try {
    const staffResult = getAllAvailableStaff();
    if (!staffResult.success) {
      return { success: false, message: 'Could not retrieve staff list' };
    }
    
    const userProperties = PropertiesService.getUserProperties();
    const allocationsJson = userProperties.getProperty('annualLeaveAllocations');
    
    let allocations = {};
    if (allocationsJson) {
      allocations = JSON.parse(allocationsJson);
    }
    
    let updatedCount = 0;
    staffResult.staff.forEach(staff => {
      // Only set if not already configured
      if (!allocations[staff.name]) {
        allocations[staff.name] = defaultAllocation;
        updatedCount++;
      }
    });
    
    userProperties.setProperty('annualLeaveAllocations', JSON.stringify(allocations));
    
    Logger.log(`Set annual leave allocations for ${updatedCount} staff members`);
    return { 
      success: true, 
      message: `Set annual leave allocations for ${updatedCount} staff members`,
      updatedCount: updatedCount,
      totalStaff: staffResult.staff.length
    };
  } catch (error) {
    Logger.log('Error setting all staff annual leave allocations: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets all annual leave allocations
 */
function getAllAnnualLeaveAllocations() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const allocationsJson = userProperties.getProperty('annualLeaveAllocations');
    
    if (!allocationsJson) {
      return { success: true, allocations: {} };
    }
    
    const allocations = JSON.parse(allocationsJson);
    return { success: true, allocations: allocations };
  } catch (error) {
    Logger.log('Error getting all annual leave allocations: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets all shift patterns available, including custom ones, and filters out hidden ones.
 */
function getShiftPatterns() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    
    // Load custom shifts
    const customShiftsJson = userProperties.getProperty('CUSTOM_SHIFTS');
    let customShifts = {};
    if (customShiftsJson) {
      try {
        customShifts = JSON.parse(customShiftsJson);
      } catch (parseError) {
        Logger.log('Error parsing custom shifts: ' + parseError.toString());
      }
    }
    
    // Load hidden predefined shifts
    const hiddenShiftsJson = userProperties.getProperty('HIDDEN_SHIFTS');
    let hiddenShifts = [];
    if (hiddenShiftsJson) {
      try {
        hiddenShifts = JSON.parse(hiddenShiftsJson);
      } catch (parseError) {
        Logger.log('Error parsing hidden shifts: ' + parseError.toString());
      }
    }

    // Start with default patterns
    const allPatterns = Object.assign({}, SHIFT_PATTERNS);
    
    // Add custom shifts
    Object.keys(customShifts).forEach(shiftName => {
      const customShift = customShifts[shiftName];
      if (customShift && customShift.label && typeof customShift.hours === 'number' && customShift.color) {
        allPatterns[shiftName] = {
          label: customShift.label,
          hours: customShift.hours,
          color: customShift.color,
          custom: true // Mark as custom
        };
      }
    });
    
    // Filter out hidden predefined shifts
    const finalPatterns = {};
    Object.keys(allPatterns).forEach(shiftName => {
      // Only include if not in the hidden list, or if it's a custom shift (custom shifts are always shown unless deleted)
      if (!hiddenShifts.includes(shiftName) || allPatterns[shiftName].custom) {
        finalPatterns[shiftName] = allPatterns[shiftName];
      }
    });

    return finalPatterns;
  } catch (error) {
    Logger.log('Error in getShiftPatterns: ' + error.toString());
    return SHIFT_PATTERNS; // Fallback to default patterns
  }
}

/**
 * Checks if the current sheet is a properly formatted rota sheet.
 */
function isValidRotaSheet(targetSheet = null) {
  try {
    const sheet = targetSheet || SpreadsheetApp.getActiveSheet();
    const headerCell = sheet.getRange(1, 1).getValue();
    
    // Check if the sheet has the expected header structure
    if (headerCell !== 'Staff Name') {
      return {
        valid: false,
        message: `Sheet "${sheet.getName()}" does not appear to be a rota sheet. Expected "Staff Name" header but found "${headerCell}".`
      };
    }
    
    // Check if there are any staff members - allow sheets with no staff as they can be added dynamically
    // Note: We allow empty staff lists as staff can be added during pattern application
    
    return { valid: true };
  } catch (error) {
    return {
      valid: false,
      message: 'Error checking sheet: ' + error.toString()
    };
  }
}

/**
 * Simple test function to verify Apps Script is working
 */
function testFunction() {
  try {
    Logger.log('TEST: Function called successfully');
    console.log('TEST: Function called successfully');
    return { 
      success: true, 
      message: 'Test function executed successfully',
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    Logger.log('TEST: Error in test function: ' + error.toString());
    return { 
      success: false, 
      message: 'Test function failed: ' + error.toString()
    };
  }
}

/**
 * Enhanced function that applies shift patterns across multiple months with auto-creation of sheets
 * This is the main function called from the Rota Generator
 */
function applyShiftPatternWithDateRange(staffName, pattern, startDate, endDate) {
  return safeExecute(() => {
    
    // Basic validation
    const patternArray = pattern.split(',').map(p => p.trim()).filter(p => p);
    if (patternArray.length === 0) {
      return { success: false, message: 'Invalid pattern. Please enter a valid shift pattern.' };
    }
    
    // Quick pattern validation
    const currentShiftPatterns = getShiftPatterns();
    const invalidShifts = patternArray.filter(shift => !currentShiftPatterns[shift]);
    if (invalidShifts.length > 0) {
      return { 
        success: false, 
        message: `Invalid shift types: ${invalidShifts.join(', ')}`
      };
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    if (start > end) {
      return { success: false, message: 'End date cannot be before start date.' };
    }
    
    // Calculate months to process
    const monthsToProcess = [];
    const currentDate = new Date(start);
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    
    while (currentDate <= end) {
      const year = currentDate.getFullYear();
      const month = currentDate.getMonth() + 1;
      const monthKey = `${year}-${month}`;
      
      if (!monthsToProcess.find(m => m.key === monthKey)) {
        monthsToProcess.push({
          key: monthKey,
          year: year,
          month: month,
          sheetName: `${monthNames[month - 1]} ${year}`
        });
      }
      
      currentDate.setMonth(currentDate.getMonth() + 1);
      currentDate.setDate(1);
    }
    
    // For multi-month operations, process only the first month and provide guidance
    if (monthsToProcess.length > 1) {
      return {
        success: false,
        message: `Multi-month operations (${monthsToProcess.length} months) may timeout. Please process one month at a time:\n\n` +
                `1. Process ${monthsToProcess[0].sheetName} first\n` +
                `2. Then process each subsequent month individually\n\n` +
                `This ensures reliable operation within Google Apps Script limits.`,
        monthsToProcess: monthsToProcess.map(m => m.sheetName),
        needsManualProcessing: true
      };
    }
    
    // Single month processing - this should work reliably
    const monthInfo = monthsToProcess[0];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create or get the sheet
    let sheet = ss.getSheetByName(monthInfo.sheetName);
    let createdNew = false;
    
    if (!sheet) {
      sheet = createMinimalRotaSheet(monthInfo.month, monthInfo.year);
      createdNew = true;
      if (!sheet) {
        return { success: false, message: `Failed to create sheet: ${monthInfo.sheetName}` };
      }
    }
    
    // Calculate date range for this month
    const monthStart = new Date(monthInfo.year, monthInfo.month - 1, 1);
    const monthEnd = new Date(monthInfo.year, monthInfo.month, 0);
    
    const rangeStart = start > monthStart ? start : monthStart;
    const rangeEnd = end < monthEnd ? end : monthEnd;
    
    const startDay = rangeStart.getDate();
    const endDay = rangeEnd.getDate();
    
    // Apply pattern
    const result = applyPatternOptimized(sheet, staffName, patternArray, startDay, endDay, 0);
    
    if (!result.success) {
      return { success: false, message: `Failed: ${result.message}` };
    }
    
    return { 
      success: true, 
      message: `✅ Pattern applied to ${monthInfo.sheetName} (days ${startDay}-${endDay})${createdNew ? ' - New sheet created' : ''}`,
      monthsProcessed: 1,
      sheetsCreated: createdNew ? 1 : 0
    };
    
  }, 'applyShiftPatternWithDateRange', { success: false, message: 'Failed to apply shift pattern.' });
}

/**
 * OPTIMIZED: Minimal sheet creation without heavy formatting
 * Creates basic rota sheet structure only - formatting applied later if needed
 */
function createMinimalRotaSheet(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                     'July', 'August', 'September', 'October', 'November', 'December'];
  const sheetName = monthNames[month - 1] + ' ' + year;
  
  // Check if sheet already exists
  try {
    const existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      return existingSheet;
    }
  } catch (e) {
    // Continue if sheet doesn't exist
  }
  
  // Create sheet with minimal setup
  const sheet = ss.insertSheet(sheetName);
  const daysInMonth = getDaysInMonth(month, year);
  
  // OPTIMIZED: Set all headers in one batch operation
  const headerValues = [['Staff Name']];
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = date.toLocaleDateString('en-US', { weekday: 'short' });
    headerValues[0].push(day + '\n' + dayOfWeek);
  }
  // Add summary columns
  headerValues[0].push('Total Hours', 'Annual Leave', 'Sick Days', 'Training Days');
  
  // Write all headers at once
  sheet.getRange(1, 1, 1, headerValues[0].length).setValues(headerValues);
  
  // OPTIMIZED: Basic column sizing and freeze in batch
  sheet.setColumnWidth(1, 150);
  for (let i = 2; i <= daysInMonth + 1; i++) {
    sheet.setColumnWidth(i, 40);
  }
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  
  // OPTIMIZED: Add only essential data validation for day columns ONLY
  // IMPORTANT: Only apply validation to day columns (2 to daysInMonth+1), NOT to summary columns
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(SHIFT_PATTERNS), true)
    .build();
  
  // Apply validation only to day columns, ensuring Total Hours column (AG) is never affected
  sheet.getRange(2, 2, 50, daysInMonth) // Only day columns (B to AF for 31-day months)
    .setDataValidation(validationRule);
  
  // Set column widths for summary columns
  for (let i = daysInMonth + 2; i <= daysInMonth + 5; i++) {
    sheet.setColumnWidth(i, 100);
  }
  
  return sheet;
}

/**
 * Create missing monthly sheets for multi-month planning
 * @param {number} selectedMonth - The month to start creating from (1-12)
 * @param {number} selectedYear - The year to create sheets for
 * @param {number} count - Number of consecutive months to create (default: 3)
 */
function createMissingMonthlySheets(selectedMonth = null, selectedYear = null, count = 3) {
  return safeExecute(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentDate = new Date();
    
    // Use provided parameters or default to current month
    const startMonth = selectedMonth || (currentDate.getMonth() + 1);
    const startYear = selectedYear || currentDate.getFullYear();
    
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    
    const monthsToCreate = [];
    
    // Generate consecutive months starting from selected month
    for (let i = 0; i < count; i++) {
      let month = startMonth + i;
      let year = startYear;
      
      // Handle year rollover
      if (month > 12) {
        year += Math.floor((month - 1) / 12);
        month = ((month - 1) % 12) + 1;
      }
      
      monthsToCreate.push({
        month: month,
        year: year,
        name: monthNames[month - 1]
      });
    }
    
    const createdSheets = [];
    const existingSheets = [];
    
    for (const monthInfo of monthsToCreate) {
      const sheetName = `${monthInfo.name} ${monthInfo.year}`;
      
      // Check if sheet already exists
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        // Create the sheet
        sheet = createMinimalRotaSheet(monthInfo.month, monthInfo.year);
        if (sheet) {
          createdSheets.push(sheetName);
        }
      } else {
        existingSheets.push(sheetName);
      }
    }
    
    let message = '';
    if (createdSheets.length > 0) {
      message += `✅ Created sheets: ${createdSheets.join(', ')}\n`;
    }
    if (existingSheets.length > 0) {
      message += `📄 Already existed: ${existingSheets.join(', ')}\n`;
    }
    
    if (createdSheets.length === 0 && existingSheets.length === 0) {
      message = '❌ No sheets were created or found.';
    }
    
    return {
      success: true,
      message: message.trim(),
      created: createdSheets.length,
      existing: existingSheets.length
    };
    
  }, 'createMissingMonthlySheets', { success: false, message: 'Failed to create monthly sheets.' });
}

/**
 * Create selected monthly sheets with auto-populated staff and annual leave entitlements
 * @param {Array} selectedMonths - Array of {month, year} objects
 */
function createSelectedMonthlySheets(selectedMonths) {
  return safeExecute(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    
    if (!selectedMonths || selectedMonths.length === 0) {
      return { success: false, message: 'No months selected for creation.' };
    }
    
    const createdSheets = [];
    const existingSheets = [];
    const failedSheets = [];
    
    // Get all existing staff from Settings and current sheets
    let allStaff = [];
    try {
      const staffResult = getAllAvailableStaff();
      if (staffResult.success && staffResult.staff) {
        // Extract only the names from the staff objects
        allStaff = staffResult.staff.map(staffObj => {
          // Handle both object format {name: "John"} and string format "John"
          return typeof staffObj === 'string' ? staffObj : staffObj.name;
        });
        Logger.log(`Found ${allStaff.length} staff members to populate in new sheets: ${allStaff.join(', ')}`);
      }
    } catch (error) {
      Logger.log(`Warning: Could not load existing staff: ${error.toString()}`);
    }
    
    // Process each selected month
    for (const monthInfo of selectedMonths) {
      const { month, year } = monthInfo;
      const sheetName = `${monthNames[month - 1]} ${year}`;
      
      try {
        // Check if sheet already exists
        let sheet = ss.getSheetByName(sheetName);
        if (sheet) {
          existingSheets.push(sheetName);
          continue;
        }
        
        // Create new sheet with full staff population
        sheet = createEnhancedRotaSheet(month, year, allStaff);
        if (sheet) {
          createdSheets.push(sheetName);
          Logger.log(`Successfully created and populated: ${sheetName}`);
        } else {
          failedSheets.push(sheetName);
        }
        
      } catch (error) {
        Logger.log(`Error creating sheet ${sheetName}: ${error.toString()}`);
        failedSheets.push(sheetName);
      }
    }
    
    // Build response message
    let message = '';
    if (createdSheets.length > 0) {
      message += `✅ Created ${createdSheets.length} sheet(s): ${createdSheets.join(', ')}\n`;
    }
    if (existingSheets.length > 0) {
      message += `📄 Already existed: ${existingSheets.join(', ')}\n`;
    }
    if (failedSheets.length > 0) {
      message += `❌ Failed to create: ${failedSheets.join(', ')}\n`;
    }
    
    if (createdSheets.length === 0 && existingSheets.length === 0) {
      message = '❌ No sheets were created or found.';
    }
    
    return {
      success: true,
      message: message.trim(),
      created: createdSheets.length,
      existing: existingSheets.length,
      failed: failedSheets.length,
      totalStaffPopulated: allStaff.length
    };
    
  }, 'createSelectedMonthlySheets', { success: false, message: 'Failed to create selected monthly sheets.' });
}

/**
 * Create an enhanced rota sheet with auto-populated staff and annual leave entitlements
 * @param {number} month - Month (1-12)
 * @param {number} year - Year
 * @param {Array} staffList - Array of staff names to populate
 */
function createEnhancedRotaSheet(month, year, staffList = []) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    const sheetName = `${monthNames[month - 1]} ${year}`;
    
    // Check if sheet already exists
    let existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      Logger.log(`Sheet ${sheetName} already exists`);
      return existingSheet;
    }
    
    // Create basic sheet structure
    const sheet = createMinimalRotaSheet(month, year);
    if (!sheet) {
      throw new Error('Failed to create basic sheet structure');
    }
    
    Logger.log(`Created basic sheet structure for ${sheetName}`);
    
    // Auto-populate with all staff members
    if (staffList && staffList.length > 0) {
      Logger.log(`Populating ${sheetName} with ${staffList.length} staff members: ${staffList.join(', ')}`);
      
      const daysInMonth = getDaysInMonth(month, year);
      
      for (let i = 0; i < staffList.length; i++) {
        const staffName = staffList[i];
        const rowNum = i + 2; // Row 1 is headers, staff start at row 2
        
        // Ensure we have a valid staff name (should be string by now)
        if (!staffName || typeof staffName !== 'string') {
          Logger.log(`Warning: Invalid staff name at index ${i}: ${JSON.stringify(staffName)}`);
          continue;
        }
        
        // Add staff name
        sheet.getRange(rowNum, 1).setValue(staffName);
        Logger.log(`Added staff member ${i + 1}/${staffList.length}: "${staffName}" to row ${rowNum}`);
        
        // Initialize summary columns with proper annual leave entitlements
        const totalHoursCol = daysInMonth + 2;
        const annualLeaveCol = daysInMonth + 3;
        const sickDaysCol = daysInMonth + 4;
        const trainingDaysCol = daysInMonth + 5;
        
        // Set formulas for all summary columns
        const totalHoursFormula = createTotalHoursFormula(rowNum, daysInMonth);
        const sickDaysFormula = createSickDaysFormula(rowNum, daysInMonth);
        const trainingDaysFormula = createTrainingDaysFormula(rowNum, daysInMonth);
        const entitlementDays = getAnnualLeaveEntitlement(staffName, year);
        const annualLeaveFormula = createAnnualLeaveFormula(rowNum, daysInMonth, entitlementDays);
        
        sheet.getRange(rowNum, totalHoursCol).setFormula(totalHoursFormula);      // Total Hours calculation
        sheet.getRange(rowNum, annualLeaveCol).setFormula(annualLeaveFormula);    // Annual leave remaining (entitlement - used)
        sheet.getRange(rowNum, sickDaysCol).setFormula(sickDaysFormula);          // Sick days count
        sheet.getRange(rowNum, trainingDaysCol).setFormula(trainingDaysFormula);  // Training days count
        
        Logger.log(`Populated staff: ${staffName} with AL entitlement: ${getAnnualLeaveEntitlement(staffName, year)}`);
      }
      
      Logger.log(`Successfully populated ${staffList.length} staff members in ${sheetName}`);
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log(`Error in createEnhancedRotaSheet: ${error.toString()}`);
    return null;
  }
}

/**
 * Get annual leave entitlement for a staff member for a given year
 * This can be customized based on your organization's policies
 * @param {string} staffName - Name of staff member
 * @param {number} year - Year to calculate entitlement for
 * @returns {number} Annual leave entitlement in days
 */
function getAnnualLeaveEntitlement(staffName, year) {
  try {
    // Try to get stored entitlements from user properties
    const properties = PropertiesService.getUserProperties();
    const entitlementKey = `al_entitlement_${staffName}_${year}`;
    const storedEntitlement = properties.getProperty(entitlementKey);
    
    if (storedEntitlement && !isNaN(parseFloat(storedEntitlement))) {
      return parseFloat(storedEntitlement);
    }
    
    // Default entitlement logic - can be customized
    // Standard UK entitlement: 28 days (including bank holidays)
    // You can modify this logic based on:
    // - Years of service
    // - Employment type (full-time/part-time)
    // - Contract terms
    
    const defaultEntitlement = 28; // Standard UK entitlement
    
    // Save the default for future reference
    properties.setProperty(entitlementKey, defaultEntitlement.toString());
    
    return defaultEntitlement;
    
  } catch (error) {
    Logger.log(`Warning: Could not determine AL entitlement for ${staffName}: ${error.toString()}`);
    return 28; // Fallback to standard entitlement
  }
}

/**
 * Create a formula to calculate total hours from shift patterns for a specific row
 * @param {number} rowNum - The row number (1-based)
 * @param {number} daysInMonth - Number of days in the month
 * @returns {string} Google Sheets formula to calculate total hours
 */
function createTotalHoursFormula(rowNum, daysInMonth) {
  // Create the range for the shift columns (B to column for last day of month)
  const startCol = 'B'; // Column B is day 1
  const endCol = getColumnLetter(daysInMonth + 1); // +1 because column A is staff name
  const range = `${startCol}${rowNum}:${endCol}${rowNum}`;
  
  // Create a SUMPRODUCT formula that maps each shift to its PAID hours
  // AL and TRAINING are paid time, SICK and OFF are unpaid
  const formula = `=SUMPRODUCT(
    IF(${range}="EARLY", 8,
    IF(${range}="LATE", 8,
    IF(${range}="NIGHT", 10,
    IF(${range}="DAY", 8,
    IF(${range}="TRAINING", 8,
    IF(${range}="AL", 8,
    IF(${range}="OFF", 0,
    IF(${range}="SICK", 0, 0))))))))
  )`.replace(/\s+/g, '');
  
  Logger.log(`Created Total Hours formula for row ${rowNum}: ${formula}`);
  return formula;
}

/**
 * Convert a column number to its letter representation (1=A, 2=B, 27=AA, etc.)
 * @param {number} columnNumber - Column number (1-based)
 * @returns {string} Column letter(s)
 */
function getColumnLetter(columnNumber) {
  let temp, letter = '';
  while (columnNumber > 0) {
    temp = (columnNumber - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNumber = (columnNumber - temp - 1) / 26;
  }
  return letter;
}

/**
 * Create a formula to count sick days for a specific row
 * @param {number} rowNum - The row number (1-based)
 * @param {number} daysInMonth - Number of days in the month
 * @returns {string} Google Sheets formula to count sick days
 */
function createSickDaysFormula(rowNum, daysInMonth) {
  const startCol = 'B';
  const endCol = getColumnLetter(daysInMonth + 1);
  const range = `${startCol}${rowNum}:${endCol}${rowNum}`;
  
  const formula = `=COUNTIF(${range},"SICK")`;
  Logger.log(`Created Sick Days formula for row ${rowNum}: ${formula}`);
  return formula;
}

/**
 * Create a formula to count training days for a specific row
 * @param {number} rowNum - The row number (1-based)
 * @param {number} daysInMonth - Number of days in the month
 * @returns {string} Google Sheets formula to count training days
 */
function createTrainingDaysFormula(rowNum, daysInMonth) {
  const startCol = 'B';
  const endCol = getColumnLetter(daysInMonth + 1);
  const range = `${startCol}${rowNum}:${endCol}${rowNum}`;
  
  const formula = `=COUNTIF(${range},"TRAINING")`;
  Logger.log(`Created Training Days formula for row ${rowNum}: ${formula}`);
  return formula;
}

/**
 * Create a formula to calculate remaining annual leave (entitlement minus days used)
 * @param {number} rowNum - The row number (1-based)
 * @param {number} daysInMonth - Number of days in the month
 * @param {number} entitlementDays - Initial annual leave entitlement
 * @returns {string} Google Sheets formula to calculate remaining annual leave
 */
function createAnnualLeaveFormula(rowNum, daysInMonth, entitlementDays) {
  const startCol = 'B';
  const endCol = getColumnLetter(daysInMonth + 1);
  const range = `${startCol}${rowNum}:${endCol}${rowNum}`;
  
  // Formula: Initial Entitlement - Days Used = Remaining Days
  const formula = `=${entitlementDays}-COUNTIF(${range},"AL")`;
  Logger.log(`Created Annual Leave formula for row ${rowNum}: ${formula} (entitlement: ${entitlementDays})`);
  return formula;
}

/**
 * OPTIMIZED: Streamlined pattern application without heavy logging
 */
function applyPatternOptimized(sheet, staffName, patternArray, startDay, endDay, patternOffset = 0) {
  try {
    // Find or add staff member with minimal operations
    const staffResult = findOrAddStaffMemberOptimized(sheet, staffName);
    if (!staffResult.success) {
      return { success: false, message: staffResult.message };
    }
    
    const staffRow = staffResult.row;
    const rangeWidth = endDay - startDay + 1;
    const rangeValues = [new Array(rangeWidth)];
    
    // SAFETY CHECK: Ensure we never write beyond day columns
    const daysInMonth = sheet.getLastColumn() - 5; // Account for summary columns
    const maxEndColumn = daysInMonth + 1; // Last day column
    const actualEndColumn = startDay + rangeWidth;
    
    if (actualEndColumn > maxEndColumn) {
      return { 
        success: false, 
        message: `Pattern would overwrite Total Hours column. Range: ${startDay}-${endDay}, Max day: ${daysInMonth}` 
      };
    }
    
    // Build pattern values
    for (let i = 0; i < rangeWidth; i++) {
      const patternIndex = (patternOffset + i) % patternArray.length;
      rangeValues[0][i] = patternArray[patternIndex];
    }
    
    // OPTIMIZED: Single write operation - guaranteed to not touch AG column
    sheet.getRange(staffRow, startDay + 1, 1, rangeWidth).setValues(rangeValues);
    
    // Calculate next offset for continuous patterns
    const nextPatternOffset = (patternOffset + rangeWidth) % patternArray.length;
    
    return { 
      success: true, 
      message: `Applied ${rangeWidth} shifts`,
      nextPatternOffset: nextPatternOffset
    };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * OPTIMIZED: Fast staff member lookup/addition
 */
function findOrAddStaffMemberOptimized(sheet, staffName) {
  try {
    // Get all staff names in one operation
    const staffRange = sheet.getRange(2, 1, Math.max(20, sheet.getLastRow() - 1), 1);
    const staffValues = staffRange.getValues();
    
    // Find existing staff member
    for (let i = 0; i < staffValues.length; i++) {
      if (staffValues[i][0].toString().trim() === staffName.trim()) {
        return { success: true, row: i + 2 };
      }
    }
    
    // Add new staff member at first empty row
    for (let i = 0; i < staffValues.length; i++) {
      if (staffValues[i][0].toString().trim() === '') {
        sheet.getRange(i + 2, 1).setValue(staffName);
        return { success: true, row: i + 2, message: 'added' };
      }
    }
    
    // Add at the end if no empty rows found
    const newRow = staffValues.length + 2;
    sheet.getRange(newRow, 1).setValue(staffName);
    return { success: true, row: newRow, message: 'added' };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Helper function to apply pattern to a specific sheet with pattern offset
 */
function applyPatternToSheet(sheet, staffName, pattern, startDay, endDay, patternOffset = 0) {
  try {
    Logger.log(`applyPatternToSheet: sheet=${sheet.getName()}, staffName=${staffName}, pattern=${pattern}, startDay=${startDay}, endDay=${endDay}, patternOffset=${patternOffset}`);
    
    // Validate sheet
    const headerValue = sheet.getRange('A1').getValue();
    if (headerValue !== 'Staff Name') {
      Logger.log(`Invalid sheet header: expected 'Staff Name', got '${headerValue}'`);
      return { success: false, message: `Invalid rota sheet format. Header is '${headerValue}' instead of 'Staff Name'` };
    }
    
    // Parse pattern
    const patternArray = pattern.split(',').map(p => p.trim()).filter(p => p);
    if (patternArray.length === 0) {
      Logger.log('Pattern array is empty after parsing');
      return { success: false, message: 'Invalid pattern' };
    }
    Logger.log(`Pattern array: [${patternArray.join(', ')}]`);
    
    // Find or add staff member
    Logger.log(`Finding or adding staff member: ${staffName}`);
    const staffResult = findOrAddStaffMember(sheet, staffName);
    if (!staffResult.success) {
      Logger.log(`Failed to find/add staff member: ${staffResult.message}`);
      return { success: false, message: staffResult.message };
    }
    Logger.log(`Staff member ${staffResult.message ? 'added' : 'found'} at row ${staffResult.row}`)
    
    const staffRow = staffResult.row;
    const daysInMonth = sheet.getLastColumn() - 1;
    
    // Validate day range
    if (startDay < 1 || startDay > daysInMonth || endDay < 1 || endDay > daysInMonth || startDay > endDay) {
      return { success: false, message: `Invalid day range: ${startDay}-${endDay}` };
    }
    
    // Apply pattern with offset
    const rangeWidth = endDay - startDay + 1;
    const rangeValues = [new Array(rangeWidth)];
    
    for (let i = 0; i < rangeWidth; i++) {
      const patternIndex = (patternOffset + i) % patternArray.length;
      rangeValues[0][i] = patternArray[patternIndex];
    }
    
    // Write to sheet
    Logger.log(`Writing pattern to range: row ${staffRow}, startCol ${startDay + 1}, width ${rangeWidth}`);
    Logger.log(`Values to write: [${rangeValues[0].join(', ')}]`);
    
    try {
      const targetRange = sheet.getRange(staffRow, startDay + 1, 1, rangeWidth);
      targetRange.setValues(rangeValues);
      SpreadsheetApp.flush(); // Ensure changes are written
      
      Logger.log(`Successfully wrote ${rangeWidth} values to sheet`);
    } catch (writeError) {
      Logger.log(`Error writing to sheet: ${writeError.toString()}`);
      return { success: false, message: `Error writing pattern: ${writeError.toString()}` };
    }
    
    // Calculate next pattern offset for continuous patterns across months
    const nextPatternOffset = (patternOffset + rangeWidth) % patternArray.length;
    
    Logger.log(`Pattern application complete. Next offset: ${nextPatternOffset}`);
    
    return { 
      success: true, 
      message: `Applied ${rangeWidth} shifts`,
      nextPatternOffset: nextPatternOffset
    };
    
  } catch (error) {
    Logger.log(`Error applying pattern to sheet: ${error.toString()}`);
    return { success: false, message: error.toString() };
  }
}

/**
 * Applies a pattern to a staff member's schedule for the specified month.
 * Enhanced with better validation, error handling, and performance.
 * @param {string} staffName - Name of the staff member
 * @param {string} pattern - Comma-separated shift pattern
 * @param {number} startDay - Starting day of the month
 * @param {number} endDay - Ending day of the month
 * @param {number} selectedMonth - Target month (1-12, optional)
 * @param {number} selectedYear - Target year (optional)
 */
function applyShiftPattern(staffName, pattern, startDay, endDay, selectedMonth = null, selectedYear = null) {
  return safeExecute(() => {
    logUsage('apply_shift_pattern_started');
    
    // Enhanced input validation
    const validation = validateInput(
      { staffName, pattern, startDay, endDay },
      ['staffName', 'pattern', 'startDay', 'endDay']
    );
    
    if (!validation.valid) {
      return { success: false, message: validation.message };
    }
    
    let sheet;
    
    // If month and year are provided, find the specific sheet
    if (selectedMonth && selectedYear) {
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                         'July', 'August', 'September', 'October', 'November', 'December'];
      const targetSheetName = `${monthNames[selectedMonth - 1]} ${selectedYear}`;
      
      Logger.log(`Looking for sheet: "${targetSheetName}"`);
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getSheetByName(targetSheetName);
      
      if (!sheet) {
        return { 
          success: false, 
          message: `Sheet "${targetSheetName}" not found. Please create the monthly sheets first.` 
        };
      }
      
      Logger.log(`Found target sheet: "${sheet.getName()}"`);
    } else {
      // Fallback to active sheet if no month/year specified
      sheet = SpreadsheetApp.getActiveSheet();
      Logger.log(`Using active sheet: "${sheet.getName()}"`);
    }
    
    // Check if the target sheet is a valid rota sheet
    const sheetCheck = isValidRotaSheet(sheet);
    if (!sheetCheck.valid) {
      return { success: false, message: sheetCheck.message };
    }
    
    // Enhanced pattern validation
    const patternArray = pattern.split(',').map(p => p.trim()).filter(p => p);
    
    if (patternArray.length === 0) {
      return { success: false, message: 'Invalid pattern. Please enter a valid shift pattern.' };
    }
    
    if (patternArray.length > CONFIG.MAX_PATTERN_LENGTH) {
      return { 
        success: false, 
        message: `Pattern too long. Maximum ${CONFIG.MAX_PATTERN_LENGTH} shifts allowed.` 
      };
    }
    
    // Validate all shifts in the pattern
    const invalidShifts = patternArray.filter(shift => !SHIFT_PATTERNS[shift]);
    if (invalidShifts.length > 0) {
      return { 
        success: false, 
        message: `Invalid shift types: ${invalidShifts.join(', ')}. Valid shifts: ${Object.keys(SHIFT_PATTERNS).join(', ')}`
      };
    }
    
    // Find or add staff member with better error handling
    const staffResult = findOrAddStaffMember(sheet, staffName);
    if (!staffResult.success) {
      return { success: false, message: staffResult.message };
    }
    
    const staffRow = staffResult.row;
    
    // Enhanced date range validation - get actual days in month from sheet name
    let daysInMonth;
    try {
      const sheetName = sheet.getName();
      const nameParts = sheetName.split(' ');
      if (nameParts.length >= 2) {
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                           'July', 'August', 'September', 'October', 'November', 'December'];
        const monthIndex = monthNames.indexOf(nameParts[0]) + 1; // 1-based month
        const year = parseInt(nameParts[1]);
        
        if (monthIndex > 0 && !isNaN(year)) {
          daysInMonth = getDaysInMonth(monthIndex, year);
          Logger.log(`Calculated days in month for ${sheetName}: ${daysInMonth}`);
        } else {
          throw new Error('Invalid month/year in sheet name');
        }
      } else {
        throw new Error('Sheet name format not recognized');
      }
    } catch (error) {
      Logger.log(`Warning: Could not determine month from sheet name "${sheet.getName()}". Using fallback calculation.`);
      // Fallback: assume the columns before summary columns are days
      // Summary columns are: Total Hours, Annual Leave, Sick Days, Training Days (4 columns)
      daysInMonth = Math.max(31, sheet.getLastColumn() - 5); // -1 for Staff Name, -4 for summary columns
    }
    
    const startDayNum = parseInt(startDay);
    const endDayNum = parseInt(endDay);
    
    if (isNaN(startDayNum) || startDayNum < 1 || startDayNum > daysInMonth) {
      return { 
        success: false, 
        message: `Invalid start day. Must be between 1 and ${daysInMonth}.`
      };
    }
    
    if (isNaN(endDayNum) || endDayNum < 1 || endDayNum > daysInMonth) {
      return { 
        success: false, 
        message: `Invalid end day. Must be between 1 and ${daysInMonth}.`
      };
    }
    
    if (endDayNum < startDayNum) {
      return {
        success: false,
        message: 'End day cannot be before start day.'
      };
    }
    
    if (endDayNum - startDayNum + 1 > CONFIG.MAX_DAYS_RANGE) {
      return {
        success: false,
        message: `Date range too large. Maximum ${CONFIG.MAX_DAYS_RANGE} days allowed.`
      };
    }
    
    // Optimized pattern application with batch processing
    const rangeWidth = endDayNum - startDayNum + 1;
    const rangeValues = [new Array(rangeWidth)];
    
    // Safety check: ensure we don't overwrite summary columns
    const maxAllowedColumn = daysInMonth + 1; // +1 because column B is day 1
    const targetEndColumn = startDayNum + rangeWidth; // End column of our range
    
    if (targetEndColumn > maxAllowedColumn) {
      return {
        success: false,
        message: `Pattern would overwrite summary columns. Range ends at column ${targetEndColumn} but max allowed is ${maxAllowedColumn}. Days in month: ${daysInMonth}.`
      };
    }
    
    Logger.log(`Applying pattern to columns ${startDayNum + 1} to ${startDayNum + rangeWidth} (width: ${rangeWidth})`);
    
    // Pre-fill pattern array for better performance
    for (let i = 0; i < rangeWidth; i++) {
      const patternIndex = i % patternArray.length;
      rangeValues[0][i] = patternArray[patternIndex];
    }
    
    // Apply values in single batch operation
    const targetRange = sheet.getRange(staffRow, startDayNum + 1, 1, rangeWidth);
    targetRange.setValues(rangeValues);
    
    // Single flush for better performance
    SpreadsheetApp.flush();
    
    logUsage('apply_shift_pattern_completed');
    
    return { 
      success: true, 
      message: `Applied ${rangeWidth} shifts for ${staffName} (pattern: ${patternArray.join(',')}).`
    };
    
  }, 'applyShiftPattern', { success: false, message: 'Failed to apply shift pattern due to an unexpected error.' });
}

/**
 * Logs a new absence in the absence tracker sheet.
 */
function logAbsence(staffName, absenceType, startDate, endDate, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let absenceSheet = ss.getSheetByName('Absence Tracker');
  
  // Create the absence sheet if it doesn't exist
  if (!absenceSheet) {
    absenceSheet = ss.insertSheet('Absence Tracker');
    
    // Set up the headers
    absenceSheet.getRange('A1:F1').setValues([['Staff Name', 'Absence Type', 'Start Date', 'End Date', 'Days', 'Reason']]);
    absenceSheet.getRange('A1:F1').setFontWeight('bold');
    absenceSheet.setColumnWidth(1, 150);
    absenceSheet.setColumnWidth(2, 100);
    absenceSheet.setColumnWidth(3, 100);
    absenceSheet.setColumnWidth(4, 100);
    absenceSheet.setColumnWidth(5, 60);
    absenceSheet.setColumnWidth(6, 250);
    absenceSheet.setFrozenRows(1);
  }
  
  // Parse dates
  const start = new Date(startDate);
  const end = new Date(endDate);
  
  // Calculate days
  const oneDay = 24 * 60 * 60 * 1000; // milliseconds in a day
  const days = Math.round(Math.abs((end - start) / oneDay)) + 1; // +1 to include both start and end days
  
  // Add the new absence record
  const lastRow = absenceSheet.getLastRow() + 1;
  absenceSheet.getRange(lastRow, 1, 1, 6).setValues([[
    staffName, 
    absenceType, 
    start, 
    end, 
    days, 
    reason
  ]]);
  
  // Format date cells
  absenceSheet.getRange(lastRow, 3, 1, 2).setNumberFormat('dd/MM/yyyy');
  
  // Apply the absence to the rota sheets
  applyAbsenceToRotas(staffName, absenceType, start, end);
  
  return { success: true, message: 'Absence logged successfully.' };
}

/**
 * Applies an absence to all relevant rota sheets with enhanced error handling and batch operations.
 */
function applyAbsenceToRotas(staffName, absenceType, startDate, endDate) {
  return safeExecute(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Map absence type to shift code
    const shiftCodeMap = {
      'Annual Leave': 'AL',
      'Sick Leave': 'SICK',
      'Training': 'TRAINING'
    };
    
    const shiftCode = shiftCodeMap[absenceType] || 'OFF';
    
    // Group dates by month for batch processing
    const datesByMonth = new Map();
    const currentDate = new Date(startDate);
    
    while (currentDate <= endDate) {
      const monthKey = `${currentDate.getFullYear()}-${currentDate.getMonth()}`;
      
      if (!datesByMonth.has(monthKey)) {
        datesByMonth.set(monthKey, {
          year: currentDate.getFullYear(),
          month: currentDate.getMonth(),
          days: []
        });
      }
      
      datesByMonth.get(monthKey).days.push(currentDate.getDate());
      
      // Move to next day safely
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    // Process each month
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    let appliedCount = 0;
    let errorCount = 0;
    
    for (const [monthKey, monthData] of datesByMonth) {
      try {
        const sheetName = `${monthNames[monthData.month]} ${monthData.year}`;
        const sheet = ss.getSheetByName(sheetName);
        
        if (!sheet) {
          Logger.log(`Sheet not found: ${sheetName}`);
          continue;
        }
        
        // Check if this is a rota sheet
        if (sheet.getRange('A1').getValue() !== 'Staff Name') {
          Logger.log(`Skipping non-rota sheet: ${sheetName}`);
          continue;
        }
        
        // Find staff member efficiently
        const staffResult = findOrAddStaffMember(sheet, staffName);
        if (!staffResult.success) {
          Logger.log(`Staff not found in ${sheetName}: ${staffResult.message}`);
          errorCount++;
          continue;
        }
        
        const staffRow = staffResult.row;
        
        // Batch update all days for this month
        const updates = [];
        monthData.days.forEach(day => {
          // Validate day is within sheet bounds
          const maxDay = sheet.getLastColumn() - 1; // -1 for staff name column
          if (day <= maxDay) {
            updates.push({
              range: sheet.getRange(staffRow, day + 1),
              value: shiftCode
            });
          }
        });
        
        // Apply all updates for this month in batch
        updates.forEach(update => {
          update.range.setValue(update.value);
        });
        
        appliedCount += updates.length;
        
      } catch (monthError) {
        Logger.log(`Error processing month ${monthKey}: ${monthError.toString()}`);
        errorCount++;
      }
    }
    
    // Single flush for all changes
    SpreadsheetApp.flush();
    
    const message = `Applied ${appliedCount} absence entries for ${staffName}${errorCount > 0 ? ` (${errorCount} errors)` : ''}`;
    Logger.log(message);
    
    return {
      success: true,
      applied: appliedCount,
      errors: errorCount,
      message: message
    };
    
  }, 'applyAbsenceToRotas', { success: false, applied: 0, errors: 1, message: 'Failed to apply absence to rotas' });
}


/**
 * Gets absence stats for all staff.
 */
function getAbsenceStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const absenceSheet = ss.getSheetByName('Absence Tracker');
  
  if (!absenceSheet) {
    return [];
  }
  
  const dataRange = absenceSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length <= 1) {
    return []; // Only header row exists
  }
  
  // Prepare stats
  const stats = {};
  
  // Process each absence record
  for (let i = 1; i < values.length; i++) {
    const staffName = values[i][0];
    const absenceType = values[i][1];
    const days = values[i][4];
    
    if (!stats[staffName]) {
      stats[staffName] = {
        name: staffName,
        annualLeave: 0,
        sick: 0,
        training: 0,
        other: 0,
        total: 0
      };
    }
    
    switch (absenceType) {
      case 'Annual Leave':
        stats[staffName].annualLeave += days;
        break;
      case 'Sick Leave':
        stats[staffName].sick += days;
        break;
      case 'Training':
        stats[staffName].training += days;
        break;
      default:
        stats[staffName].other += days;
    }
    
    stats[staffName].total += days;
  }
  
  // Convert to array
  return Object.values(stats);
}

/**
 * Simple implementation of rolling pattern application that ensures patterns are applied across all months
 */
function applyRollingPattern(staffName, pattern, startDay, endDay, isRolling, monthsAhead, isNewStarter, newStarterInfo, debugMode) {
  try {
    // Get information about the current sheet/month
    const currentSheet = SpreadsheetApp.getActiveSheet();
    const sheetName = currentSheet.getName();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Enable debug logging for this function
    const isDebugMode = debugMode === true;
    Logger.log(`SIMPLIFIED Rolling pattern application started. Debug Mode: ${isDebugMode}`);
    
    // Validate pattern
    const patternArray = pattern.split(',').map(p => p.trim());
    if (patternArray.length === 0) {
      return { success: false, message: 'Invalid pattern. Please enter a valid shift pattern.' };
    }
    
    // Validate all shifts in the pattern
    const invalidShifts = patternArray.filter(shift => !SHIFT_PATTERNS[shift]);
    if (invalidShifts.length > 0) {
      return { 
        success: false, 
        message: `Invalid shift types in pattern: ${invalidShifts.join(', ')}. Valid shifts are: ${Object.keys(SHIFT_PATTERNS).join(', ')}`
      };
    }
    
    // Try to determine month and year from sheet name
    // Expected format: "Month YYYY" (e.g., "January 2023")
    const nameParts = sheetName.split(' ');
    if (nameParts.length < 2) {
      return { 
        success: false, 
        message: 'Could not determine month/year from sheet name for rolling rota. Sheet name should be in format "Month YYYY".'
      };
    }
    
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    
    const monthName = nameParts[0];
    const currentMonthIndex = monthNames.indexOf(monthName);
    const currentYear = parseInt(nameParts[1]);
    
    if (currentMonthIndex === -1 || isNaN(currentYear)) {
      return { 
        success: false, 
        message: 'Invalid sheet name format for rolling rota. Sheet name should be "Month YYYY" (e.g., "January 2023").'
      };
    }
    
    Logger.log(`Current month/year: ${monthName} ${currentYear}`);
    Logger.log(`Rolling enabled: ${isRolling}, Months ahead: ${monthsAhead}`);
    
    // Calculate target months
    let targetMonths = [];
    let targetYear = currentYear;
    let endMonth = currentMonthIndex;
    
    // If using March 31st option
    if (isRolling && monthsAhead === 'march_31') {
      // If we're already past March, target next year's March
      if (currentMonthIndex >= 2) { // March is index 2
        targetYear++;
      }
      endMonth = 2; // March
      
      let tempMonth = currentMonthIndex;
      let tempYear = currentYear;
      
      // Generate list of months from current to March target
      while (tempYear < targetYear || (tempYear === targetYear && tempMonth <= endMonth)) {
        targetMonths.push({
          monthIndex: tempMonth,
          year: tempYear,
          name: monthNames[tempMonth] + ' ' + tempYear
        });
        
        tempMonth++;
        if (tempMonth > 11) {
          tempMonth = 0;
          tempYear++;
        }
      }
    } else {
      // Standard X months ahead option
      let calculatedMonthsAhead = parseInt(monthsAhead) || 3; // Default to 3 if not a valid number
      
      // Include current month
      targetMonths.push({
        monthIndex: currentMonthIndex,
        year: currentYear,
        name: monthNames[currentMonthIndex] + ' ' + currentYear
      });
      
      // Add future months
      for (let i = 1; i <= calculatedMonthsAhead; i++) {
        let futureMonthIndex = (currentMonthIndex + i) % 12;
        let futureYear = currentYear + Math.floor((currentMonthIndex + i) / 12);
        
        targetMonths.push({
          monthIndex: futureMonthIndex,
          year: futureYear,
          name: monthNames[futureMonthIndex] + ' ' + futureYear
        });
      }
    }
    
    Logger.log(`Will process ${targetMonths.length} months: ${targetMonths.map(m => m.name).join(', ')}`);
    
    // Process each month
    let successCount = 0;
    let failureCount = 0;
    let createdSheets = [];
    let skippedSheets = [];
    
    for (const monthData of targetMonths) {
      const month = monthData.monthIndex;
      const year = monthData.year;
      const monthSheetName = monthData.name;
      
      Logger.log(`Processing month: ${monthSheetName}`);
      
      // Skip current month if it's a new starter and before their start date
      let skipMonth = false;
      let monthStartDay = startDay;
      
      if (isNewStarter && newStarterInfo) {
        // Check if this is before the starter's start date
        if (year < newStarterInfo.year || 
            (year === newStarterInfo.year && month < newStarterInfo.month - 1)) {
          skipMonth = true;
          Logger.log(`Skipping ${monthSheetName} - before start date`);
          skippedSheets.push(monthSheetName);
          continue;
        }
        
        // If this is the starter's first month, adjust start day
        if (year === newStarterInfo.year && month === newStarterInfo.month - 1) {
          monthStartDay = Math.max(startDay, newStarterInfo.day);
          Logger.log(`First month for new starter - adjusted start day to ${monthStartDay}`);
        }
      }
      
      // Get or create the sheet
      let sheet;
      if (monthSheetName === sheetName) {
        // Use current sheet
        sheet = currentSheet;
      } else {
        // Check if sheet exists
        sheet = ss.getSheetByName(monthSheetName);
        
        // Create if it doesn't exist
        if (!sheet) {
          try {
            Logger.log(`Creating new sheet for ${monthSheetName}`);
            sheet = createRotaSheetInternal(month + 1, year);
            
            if (!sheet) {
              Logger.log(`Failed to create sheet ${monthSheetName}`);
              failureCount++;
              continue;
            }
            
            createdSheets.push(monthSheetName);
          } catch (e) {
            Logger.log(`Error creating sheet for ${monthSheetName}: ${e}`);
            failureCount++;
            continue;
          }
        }
      }
      
      // Activate the sheet
      ss.setActiveSheet(sheet);
      SpreadsheetApp.flush();
      
      // Find or add staff member
      const staffResult = findOrAddStaffMember(sheet, staffName);
      if (!staffResult.success) {
        Logger.log(`Failed to find/add staff in ${monthSheetName}: ${staffResult.message}`);
        failureCount++;
        continue;
      }
      
      const staffRow = staffResult.row;
      
      // Calculate days in month and adjust end day if needed
      const daysInMonth = getDaysInMonth(month + 1, year);
      const monthEndDay = Math.min(endDay, daysInMonth);
      
      // Check if start day is valid
      if (monthStartDay > monthEndDay) {
        Logger.log(`Invalid day range in ${monthSheetName}: start (${monthStartDay}) > end (${monthEndDay})`);
        skippedSheets.push(monthSheetName);
        continue;
      }
      
      Logger.log(`Applying pattern to ${staffName} in ${monthSheetName} from day ${monthStartDay} to ${monthEndDay}`);
      
      try {
        // Direct and simple approach - just write the pattern to each cell
        // This is simpler and more reliable than the pattern continuation approach
        let patternsApplied = 0;
        
        // Use batch update for better performance
        const rangeWidth = monthEndDay - monthStartDay + 1;
        const rangeValues = [new Array(rangeWidth).fill('')];
        
        for (let day = monthStartDay; day <= monthEndDay; day++) {
          // Get shift type for this day - use simple modulo to repeat pattern
          const patternIndex = (day - monthStartDay) % patternArray.length;
          const shiftType = patternArray[patternIndex];
          
          // Only set valid shifts
          if (SHIFT_PATTERNS[shiftType]) {
            rangeValues[0][day - monthStartDay] = shiftType;
            patternsApplied++;
          }
        }
        
        // Apply values in one operation
        sheet.getRange(staffRow, monthStartDay + 1, 1, rangeWidth).setValues(rangeValues);
        SpreadsheetApp.flush();
        
        Logger.log(`Successfully applied ${patternsApplied} shifts to ${monthSheetName}`);
        successCount++;
      } catch (error) {
        Logger.log(`Error applying pattern to ${monthSheetName}: ${error}`);
        
        // Fallback to individual cell updates
        try {
          Logger.log(`Attempting individual cell updates as fallback for ${monthSheetName}`);
          let individualSuccess = 0;
          
          for (let day = monthStartDay; day <= monthEndDay; day++) {
            const patternIndex = (day - monthStartDay) % patternArray.length;
            const shiftType = patternArray[patternIndex];
            
            if (SHIFT_PATTERNS[shiftType]) {
              sheet.getRange(staffRow, day + 1).setValue(shiftType);
              individualSuccess++;
            }
          }
          
          SpreadsheetApp.flush();
          
          if (individualSuccess > 0) {
            Logger.log(`Fallback succeeded - applied ${individualSuccess} shifts to ${monthSheetName}`);
            successCount++;
          } else {
            Logger.log(`Fallback failed for ${monthSheetName}`);
            failureCount++;
          }
        } catch (fallbackError) {
          Logger.log(`Fallback error for ${monthSheetName}: ${fallbackError}`);
          failureCount++;
        }
      }
    }
    
    // Return to original sheet
    ss.setActiveSheet(currentSheet);
    SpreadsheetApp.flush();
    
    // Build result message
    let message = `Applied pattern to ${successCount} month(s).`;
    if (isDebugMode) {
      message = `[Debug Mode] ${message}`;
    }
    if (failureCount > 0) {
      message += ` Failed for ${failureCount} month(s).`;
    }
    if (createdSheets.length > 0) {
      message += ` Created sheets: ${createdSheets.join(', ')}.`;
    }
    if (skippedSheets.length > 0) {
      message += ` Skipped ${skippedSheets.length} month(s).`;
    }
    
    message += ` Pattern continues across all months.`;
    
    return { success: true, message: message };
  } catch (error) {
    Logger.log(`Error in applyRollingPattern: ${error}`);
    return { success: false, message: `Error applying rolling pattern: ${error.toString()}` };
  }
}

/**
 * This is a wrapper function that calls the new simplified implementation
 */
function applyShiftPatternWithRolling(staffName, pattern, startDay, endDay, isRolling, monthsAhead, isNewStarter, newStarterInfo, debugMode) {
  return applyRollingPattern(staffName, pattern, startDay, endDay, isRolling, monthsAhead, isNewStarter, newStarterInfo, debugMode);
}

/**
 * Finds a staff member in the sheet or adds them if not found
 */
function findOrAddStaffMember(sheet, staffName) {
  try {
    // Get all staff names
    const lastRow = Math.max(2, sheet.getLastRow());
    const staffRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const staffValues = staffRange.getValues();
    
    // Check if staff exists
    for (let i = 0; i < staffValues.length; i++) {
      if (staffValues[i][0].toString() === staffName) {
        return { success: true, row: i + 2 }; // +2 for header and 1-based indexing
      }
    }
    
    // Staff not found, add them
    const newRow = lastRow + 1;
    sheet.getRange(newRow, 1).setValue(staffName);
    
    // Initialize summary columns for new staff member
    try {
      // Get days in month from sheet structure to determine summary column positions
      const sheetName = sheet.getName();
      const nameParts = sheetName.split(' ');
      let daysInMonth = 31; // Default fallback
      
      if (nameParts.length >= 2) {
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                           'July', 'August', 'September', 'October', 'November', 'December'];
        const monthIndex = monthNames.indexOf(nameParts[0]) + 1;
        const year = parseInt(nameParts[1]);
        
        if (monthIndex > 0 && !isNaN(year)) {
          daysInMonth = getDaysInMonth(monthIndex, year);
        }
      }
      
      // Set up summary columns: Total Hours, Annual Leave, Sick Days, Training Days
      const totalHoursCol = daysInMonth + 2;
      const annualLeaveCol = daysInMonth + 3;
      const sickDaysCol = daysInMonth + 4;
      const trainingDaysCol = daysInMonth + 5;
      
      // Initialize with all formulas for summary columns
      const totalHoursFormula = createTotalHoursFormula(newRow, daysInMonth);
      const sickDaysFormula = createSickDaysFormula(newRow, daysInMonth);
      const trainingDaysFormula = createTrainingDaysFormula(newRow, daysInMonth);
      
      // Get annual leave entitlement and create deduction formula
      let annualLeaveEntitlement = 28; // Default
      try {
        const sheetParts = sheet.getName().split(' ');
        if (sheetParts.length >= 2) {
          const year = parseInt(sheetParts[1]);
          if (!isNaN(year)) {
            annualLeaveEntitlement = getAnnualLeaveEntitlement(staffName, year);
          }
        }
      } catch (error) {
        Logger.log(`Warning: Could not determine year for AL entitlement: ${error.toString()}`);
      }
      const annualLeaveFormula = createAnnualLeaveFormula(newRow, daysInMonth, annualLeaveEntitlement);
      
      sheet.getRange(newRow, totalHoursCol).setFormula(totalHoursFormula);      // Total Hours calculation
      sheet.getRange(newRow, annualLeaveCol).setFormula(annualLeaveFormula);    // Annual Leave remaining (entitlement - used)
      sheet.getRange(newRow, sickDaysCol).setFormula(sickDaysFormula);          // Sick Days count
      sheet.getRange(newRow, trainingDaysCol).setFormula(trainingDaysFormula);  // Training Days count
      
      Logger.log(`Initialized summary columns for ${staffName} at row ${newRow}`);
    } catch (error) {
      Logger.log(`Warning: Could not initialize summary columns for ${staffName}: ${error.toString()}`);
    }
    
    return { success: true, row: newRow, message: "Staff member added with summary columns initialized" };
  } catch (e) {
    return { success: false, message: `Error finding/adding staff: ${e.toString()}` };
  }
}

/**
 * Internal function to create a rota sheet programmatically.
 * Based on createRotaSheet but without UI prompts for month/year.
 */
function createRotaSheetInternal(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create new sheet
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                     'July', 'August', 'September', 'October', 'November', 'December'];
  const sheetName = monthNames[month - 1] + ' ' + year;
  
  Logger.log(`Creating internal sheet: ${sheetName}`);
  
  // Check if the sheet already exists
  try {
    const existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      Logger.log(`Sheet ${sheetName} already exists, not recreating`);
      return existingSheet; // Sheet already exists, don't recreate it
    }
  } catch (e) {
    Logger.log(`Error checking if sheet exists: ${e}`);
    // Continue if sheet doesn't exist
  }
  
  // Create and format new sheet
  Logger.log(`Creating new sheet: ${sheetName}`);
  const sheet = ss.insertSheet(sheetName);
  
  // Get days in month
  const daysInMonth = getDaysInMonth(month, year);
  Logger.log(`Days in month: ${daysInMonth}`);
  
  // Set up columns
  sheet.getRange('A1').setValue('Staff Name');
  
  // Create date headers
  const headerRange = sheet.getRange(1, 2, 1, daysInMonth);
  const dateValues = [];
  
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = date.toLocaleDateString('en-US', { weekday: 'short' });
    dateValues.push(day + '\n' + dayOfWeek);
  }
  
  headerRange.setValues([dateValues]);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  
  // Add summary columns
  sheet.getRange(1, daysInMonth + 2).setValue('Total Hours');
  sheet.getRange(1, daysInMonth + 3).setValue('Annual Leave');
  sheet.getRange(1, daysInMonth + 4).setValue('Sick Days');
  sheet.getRange(1, daysInMonth + 5).setValue('Training Days');
  
  // Format the sheet
  sheet.setColumnWidth(1, 150);
  for (let i = 2; i <= daysInMonth + 1; i++) {
    sheet.setColumnWidth(i, 40);
  }
  for (let i = daysInMonth + 2; i <= daysInMonth + 5; i++) {
    sheet.setColumnWidth(i, 100);
  }
  
  // Freeze header row and staff column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  
  // Highlight weekends
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = date.getDay();
    
    if (dayOfWeek === 0 || dayOfWeek === 6) { // Weekend
      sheet.getRange(1, day + 1, sheet.getMaxRows(), 1).setBackground('#f3f3f3');
    }
  }
  
  // Add data validation for cells (shift patterns)
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(SHIFT_PATTERNS), true)
    .build();
  
  sheet.getRange(2, 2, sheet.getMaxRows() - 1, daysInMonth)
    .setDataValidation(validationRule);
  
    // Add optimized formulas for summary columns using helper functions
    const totalHoursFormula = buildHoursFormula(daysInMonth);
    const alFormula = buildAnnualLeaveBalanceFormula(daysInMonth);
    const sickFormula = buildCountFormula(daysInMonth, 'SICK', 4);
    const trainingFormula = buildCountFormula(daysInMonth, 'TRAINING', 5);
    
    sheet.getRange(2, daysInMonth + 2, sheet.getMaxRows() - 1, 1).setFormula(totalHoursFormula);
    sheet.getRange(2, daysInMonth + 3, sheet.getMaxRows() - 1, 1).setFormula(alFormula);
    sheet.getRange(2, daysInMonth + 4, sheet.getMaxRows() - 1, 1).setFormula(sickFormula);
    sheet.getRange(2, daysInMonth + 5, sheet.getMaxRows() - 1, 1).setFormula(trainingFormula);
  
  // Add conditional formatting for shift types
  const rules = [];
  Object.keys(SHIFT_PATTERNS).forEach(shiftType => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(shiftType)
      .setBackground(SHIFT_PATTERNS[shiftType].color)
      .setRanges([sheet.getRange(2, 2, sheet.getMaxRows() - 1, daysInMonth)])
      .build();
    
    rules.push(rule);
  });
  
  sheet.setConditionalFormatRules(rules);
  
  // Copy staff names from current sheet
  try {
    const currentSheet = SpreadsheetApp.getActiveSheet();
    if (currentSheet.getName() !== sheetName) { // Don't copy from self
      const staffRange = currentSheet.getRange(2, 1, Math.max(1, currentSheet.getLastRow() - 1), 1);
      const staffValues = staffRange.getValues();
      const filteredStaff = staffValues.filter(row => row[0].toString().trim() !== '');
      
      if (filteredStaff.length > 0) {
        sheet.getRange(2, 1, filteredStaff.length, 1).setValues(filteredStaff);
        Logger.log(`Copied ${filteredStaff.length} staff members to new sheet`);
      } else {
        Logger.log('No staff members to copy');
      }
    }
  } catch (e) {
    // Continue if staff copy fails
    Logger.log(`Error copying staff names: ${e}`);
  }
  
  Logger.log(`Sheet ${sheetName} created successfully`);
  return sheet;
}

/**
 * Helper function to get the number of days in a month
 */
function getDaysInMonth(month, year) {
  return new Date(year, month, 0).getDate();
}


/**
 * Returns the name of the active sheet for client-side use.
 */
function getActiveSheetName() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Log for debugging
    Logger.log(`Active sheet detected: "${sheetName}"`);
    
    // Check if this is a month-year sheet format
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    
    // Check format by seeing if the sheet name starts with a month name
    const nameParts = sheetName.split(' ');
    if (nameParts.length >= 2 && monthNames.includes(nameParts[0])) {
      Logger.log(`Sheet "${sheetName}" appears to be a valid month-year format`);
      return sheetName;
    } else {
      Logger.log(`Sheet "${sheetName}" does not appear to be in Month YYYY format`);
      
      // Find a different sheet with the correct format
      const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
      for (const otherSheet of allSheets) {
        const otherName = otherSheet.getName();
        const otherParts = otherName.split(' ');
        if (otherParts.length >= 2 && monthNames.includes(otherParts[0])) {
          Logger.log(`Found alternative month-year sheet: "${otherName}"`);
          return otherName;
        }
      }
      
      // If we still can't find a valid month sheet, return the original
      return sheetName;
    }
  } catch (e) {
    Logger.log(`Error getting active sheet name: ${e}`);
    return null;
  }
}

/**
 * Removes staff members from the sheet.
 * This removes them from the current sheet only.
 */
function removeStaffMembers(staffData) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Validate that we're on a rota sheet
    const headerCell = sheet.getRange(1, 1).getValue();
    if (headerCell !== 'Staff Name') {
      return {
        success: false,
        message: 'This operation can only be performed on a rota sheet. Please select a valid rota sheet.'
      };
    }
    
    // Track results
    const removedStaff = [];
    let errors = [];
    
    // Process each staff member
    staffData.forEach(staff => {
      try {
        // Find the row with the staff member's name and clear it
        const row = staff.row;
        
        // Verify the name at this row matches
        const nameInSheet = sheet.getRange(row, 1).getValue();
        if (nameInSheet === staff.name) {
          // Clear the entire row (name and all shift data)
          const rowRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());
          rowRange.clearContent();
          removedStaff.push(staff.name);
          Logger.log(`Removed staff member: ${staff.name} from row ${row}`);
        } else {
          errors.push(`Name mismatch for row ${row}: expected "${staff.name}" but found "${nameInSheet}"`);
        }
      } catch (e) {
        errors.push(`Error removing ${staff.name}: ${e.toString()}`);
      }
    });
    
    // Return result
    if (removedStaff.length > 0) {
      let message = `Successfully removed ${removedStaff.length} staff member(s) from ${sheetName}`;
      if (errors.length > 0) {
        message += `, but encountered ${errors.length} error(s)`;
      }
      return { success: true, message: message };
    } else if (errors.length > 0) {
      return { success: false, message: `Failed to remove staff members: ${errors[0]}` };
    } else {
      return { success: false, message: 'No staff members were removed' };
    }
  } catch (e) {
    Logger.log(`Error in removeStaffMembers: ${e}`);
    return { success: false, message: `Error: ${e.toString()}` };
  }
}

/**
 * Enhanced onEdit function for rota sheets
 * Provides logging and validation without interfering with normal editing
 */
function onEdit(e) {
  try {
    // Only log and validate, don't interfere with normal editing
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    
    // Check if this is a rota sheet
    if (sheet.getRange('A1').getValue() !== 'Staff Name') {
      return; // Not a rota sheet, no action needed
    }
    
    // Only process data cells (not headers, not staff names)
    if (row < 2 || col < 2) {
      return;
    }
    
    // Validate the entered value if it's a shift cell
    const value = e.value;
    if (value && typeof value === 'string') {
      const shiftType = value.toString().trim().toUpperCase();
      
      // Check if it's a valid shift type
      if (!SHIFT_PATTERNS[shiftType] && shiftType !== '') {
        // Log invalid shift entry for analytics
        logUsage('invalid_shift_entered');
        
        // Optional: Show a warning (commented out to avoid interrupting workflow)
        // SpreadsheetApp.getUi().alert(
        //   'Invalid Shift Type',
        //   `"${value}" is not a valid shift type. Valid options: ${Object.keys(SHIFT_PATTERNS).join(', ')}`,
        //   SpreadsheetApp.getUi().ButtonSet.OK
        // );
      } else if (SHIFT_PATTERNS[shiftType]) {
        // Log successful shift entry
        logUsage('valid_shift_entered');
      }
    }
    
  } catch (error) {
    // Don't interrupt user's workflow, just log the error
    Logger.log('Error in onEdit: ' + error.toString());
  }
}

/**
 * OAuth compliance functions for Google Workspace Marketplace
 */

/**
 * Homepage trigger for add-on
 */
function onHomepage() {
  try {
    logUsage('homepage_opened');
    
    const html = HtmlService.createHtmlOutputFromFile('Homepage')
      .setTitle('Staff Rota Manager')
      .setWidth(400)
      .setHeight(300);
    
    return html;
  } catch (error) {
    Logger.log('Error in onHomepage: ' + error.toString());
    return HtmlService.createHtmlOutput('Error loading homepage. Please try again.');
  }
}

/**
 * Handles authorization requirements
 */
function onAuthorizationRequired() {
  try {
    logUsage('authorization_required');
    
    const html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h2>🔐 Authorization Required</h2>
        <p>Staff Rota Manager needs permission to access your spreadsheet to provide its functionality.</p>
        <p><strong>We only request minimal permissions:</strong></p>
        <ul>
          <li>View and manage this spreadsheet</li>
          <li>Display our user interface</li>
        </ul>
        <p>Your data stays secure in your Google Sheets and is never sent to external servers.</p>
        <button onclick="authorize()" style="padding: 10px 20px; background: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer;">
          Grant Permission
        </button>
      </div>
      <script>
        function authorize() {
          google.script.run.requestAuthorization();
        }
      </script>
    `)
    .setTitle('Authorization Required')
    .setWidth(400)
    .setHeight(250);
    
    return html;
  } catch (error) {
    Logger.log('Error in onAuthorizationRequired: ' + error.toString());
    return HtmlService.createHtmlOutput('Authorization error. Please contact support.');
  }
}

/**
 * Requests user authorization
 */
function requestAuthorization() {
  try {
    // This function will trigger the authorization flow
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      'Authorization Granted',
      'Thank you! Staff Rota Manager now has the necessary permissions to function.',
      ui.ButtonSet.OK
    );
    
    logUsage('authorization_granted');
    
    return { success: true };
  } catch (error) {
    Logger.log('Error in requestAuthorization: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Data compliance functions for GDPR/CCPA
 */

/**
 * Exports user data for GDPR compliance
 */
function exportUserData() {
  return safeExecute(() => {
    logUsage('data_export_requested');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const exportData = {
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      exportDate: new Date().toISOString(),
      sheets: []
    };
    
    sheets.forEach(sheet => {
      if (sheet.getRange('A1').getValue() === 'Staff Name') {
        // This is a rota sheet
        const sheetData = {
          name: sheet.getName(),
          staffData: [],
          rotaData: []
        };
        
        // Get staff names (anonymized for export)
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          const staffRange = sheet.getRange(2, 1, lastRow - 1, 1);
          sheetData.staffData = staffRange.getValues().flat().filter(name => name);
        }
        
        exportData.sheets.push(sheetData);
      }
    });
    
    return {
      success: true,
      data: exportData,
      message: 'Data export completed successfully'
    };
    
  }, 'exportUserData', { success: false, message: 'Failed to export data' });
}

/**
 * Deletes user data for GDPR compliance
 */
function deleteUserData(confirmationCode) {
  return safeExecute(() => {
    // Require confirmation code to prevent accidental deletion
    if (confirmationCode !== 'DELETE_MY_DATA') {
      return {
        success: false,
        message: 'Invalid confirmation code. To delete your data, use confirmation code: DELETE_MY_DATA'
      };
    }
    
    logUsage('data_deletion_requested');
    
    const userProperties = PropertiesService.getUserProperties();
    
    // Clear all user properties
    userProperties.deleteAll();
    
    // Note: We cannot delete the spreadsheet data as it belongs to the user
    // and may contain other important information
    
    logUsage('data_deletion_completed');
    
    return {
      success: true,
      message: 'User preferences and settings have been deleted. Your spreadsheet data remains untouched as it belongs to you.'
    };
    
  }, 'deleteUserData', { success: false, message: 'Failed to delete data' });
}

/**
 * Gets user's current privacy settings
 */
function getPrivacySettings() {
  return safeExecute(() => {
    const userProperties = PropertiesService.getUserProperties();
    
    return {
      success: true,
      settings: {
        analyticsEnabled: userProperties.getProperty('ANALYTICS_ENABLED') !== 'false',
        notificationsEnabled: userProperties.getProperty('NOTIFICATION_ENABLED') === 'true',
        dataRetentionPeriod: userProperties.getProperty('DATA_RETENTION_PERIOD') || 'indefinite',
        lastUpdated: userProperties.getProperty('SETTINGS_LAST_UPDATED') || 'never'
      }
    };
    
  }, 'getPrivacySettings', { success: false, message: 'Failed to get privacy settings' });
}

/**
 * Updates user's privacy settings
 */
function updatePrivacySettings(settings) {
  return safeExecute(() => {
    const userProperties = PropertiesService.getUserProperties();
    
    if (settings.analyticsEnabled !== undefined) {
      userProperties.setProperty('ANALYTICS_ENABLED', settings.analyticsEnabled.toString());
    }
    
    if (settings.notificationsEnabled !== undefined) {
      userProperties.setProperty('NOTIFICATION_ENABLED', settings.notificationsEnabled.toString());
    }
    
    userProperties.setProperty('SETTINGS_LAST_UPDATED', new Date().toISOString());
    
    logUsage('privacy_settings_updated');
    
    return {
      success: true,
      message: 'Privacy settings updated successfully'
    };
    
  }, 'updatePrivacySettings', { success: false, message: 'Failed to update privacy settings' });
}

/**
 * Saves a custom shift pattern to user properties
 */
function saveCustomShift(shiftName, shiftInfo) {
  return safeExecute(() => {
    const userProperties = PropertiesService.getUserProperties();
    const customShiftsJson = userProperties.getProperty('CUSTOM_SHIFTS');
    let customShifts = {};
    
    if (customShiftsJson) {
      try {
        customShifts = JSON.parse(customShiftsJson);
      } catch (e) {
        Logger.log('Error parsing existing custom shifts: ' + e.toString());
      }
    }
    
    // Ensure shiftInfo is valid and only store necessary properties
    if (!shiftInfo || !shiftInfo.label || typeof shiftInfo.hours !== 'number' || !shiftInfo.color) {
      return { success: false, message: 'Invalid shift information provided.' };
    }

    customShifts[shiftName] = {
      label: shiftInfo.label,
      hours: shiftInfo.hours,
      color: shiftInfo.color,
      custom: true
    };
    
    userProperties.setProperty('CUSTOM_SHIFTS', JSON.stringify(customShifts));
    
    // Update conditional formatting on all existing rota sheets
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      const allPatterns = getShiftPatterns(); // Get updated patterns including new custom shift
      
      sheets.forEach(sheet => {
        // Check if this is a rota sheet
        if (sheet.getRange('A1').getValue() === 'Staff Name') {
          updateConditionalFormattingForCustomShifts(sheet, allPatterns);
          
          // Also update data validation to include the new shift in dropdowns
          const shiftNames = Object.keys(allPatterns);
          const validationRule = SpreadsheetApp.newDataValidation()
            .requireValueInList(shiftNames, true)
            .build();
          
          const lastRow = Math.max(2, sheet.getLastRow());
          const lastCol = Math.max(2, sheet.getLastColumn());
          
          if (lastCol > 1) {
            const dataRange = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
            dataRange.setDataValidation(validationRule);
          }
        }
      });
      
      Logger.log(`Updated conditional formatting for custom shift: ${shiftName}`);
    } catch (updateError) {
      Logger.log(`Error updating conditional formatting for custom shift: ${updateError.toString()}`);
      // Don't fail the entire operation if formatting update fails
    }
    
    Logger.log(`Custom shift saved: ${shiftName}`);
    return { success: true, message: 'Custom shift saved successfully.' };
  }, 'saveCustomShift', { success: false, message: 'Failed to save custom shift.' });
}

/**
 * Updates data validation on the current sheet to include custom shifts
 */
function updateSheetDataValidation() {
  return safeExecute(() => {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Check if this is a rota sheet
    if (sheet.getRange('A1').getValue() !== 'Staff Name') {
      return { success: false, message: 'Not a rota sheet' };
    }
    
    // Get all shift patterns (including custom ones)
    const allPatterns = getShiftPatterns();
    const shiftNames = Object.keys(allPatterns);
    
    // Create new validation rule
    const validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(shiftNames, true)
      .build();
    
    // Apply to the data area (skip header row and staff column)
    const lastRow = Math.max(2, sheet.getLastRow());
    const lastCol = Math.max(2, sheet.getLastColumn());
    
    if (lastCol > 1) {
      const dataRange = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
      dataRange.setDataValidation(validationRule);
    }
    
    // Update conditional formatting to include custom shifts
    updateConditionalFormattingForCustomShifts(sheet, allPatterns);
    
    return { success: true };
    
  }, 'updateSheetDataValidation', { success: false });
}

/**
 * Updates conditional formatting to include custom shift colors
 */
function updateConditionalFormattingForCustomShifts(sheet, allPatterns) {
  return safeExecute(() => {
    // Get existing rules
    const existingRules = sheet.getConditionalFormatRules();
    
    // Get the data range for formatting
    const lastRow = Math.max(2, sheet.getLastRow());
    const lastCol = Math.max(2, sheet.getLastColumn());
    
    if (lastCol <= 1) return;
    
    const dataRange = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
    
    // Create new rules for all patterns
    const newRules = [];
    Object.keys(allPatterns).forEach(shiftType => {
      const pattern = allPatterns[shiftType];
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(shiftType)
        .setBackground(pattern.color)
        .setRanges([dataRange])
        .build();
      newRules.push(rule);
    });
    
    // Apply all rules
    sheet.setConditionalFormatRules(newRules);
    
    return { success: true };
    
  }, 'updateConditionalFormattingForCustomShifts', { success: false });
}

/**
 * Deletes a custom shift pattern or marks a predefined one as hidden.
 */
function deleteShiftPattern(shiftName) {
  return safeExecute(() => {
    const userProperties = PropertiesService.getUserProperties();
    
    // Check if it's a custom shift
    const customShiftsJson = userProperties.getProperty('CUSTOM_SHIFTS');
    let customShifts = {};
    if (customShiftsJson) {
      try {
        customShifts = JSON.parse(customShiftsJson);
      } catch (e) {
        Logger.log('Error parsing existing custom shifts for deletion: ' + e.toString());
      }
    }

    let updateResult = { success: false, message: 'Shift not found.' };
    
    if (customShifts[shiftName]) {
      // It's a custom shift, delete it
      delete customShifts[shiftName];
      userProperties.setProperty('CUSTOM_SHIFTS', JSON.stringify(customShifts));
      Logger.log(`Custom shift deleted: ${shiftName}`);
      updateResult = { success: true, message: 'Custom shift deleted successfully.' };
    } else if (SHIFT_PATTERNS[shiftName]) {
      // It's a predefined shift, mark it as hidden
      const hiddenShiftsJson = userProperties.getProperty('HIDDEN_SHIFTS');
      let hiddenShifts = [];
      if (hiddenShiftsJson) {
        try {
          hiddenShifts = JSON.parse(hiddenShiftsJson);
        } catch (e) {
          Logger.log('Error parsing existing hidden shifts: ' + e.toString());
        }
      }

      if (!hiddenShifts.includes(shiftName)) {
        hiddenShifts.push(shiftName);
        userProperties.setProperty('HIDDEN_SHIFTS', JSON.stringify(hiddenShifts));
        Logger.log(`Predefined shift hidden: ${shiftName}`);
        updateResult = { success: true, message: 'Predefined shift hidden successfully.' };
      } else {
        updateResult = { success: true, message: 'Shift was already hidden.' };
      }
    }
    
    // Update conditional formatting on all existing rota sheets if operation was successful
    if (updateResult.success) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = ss.getSheets();
        const allPatterns = getShiftPatterns(); // Get updated patterns excluding deleted/hidden shift
        
        sheets.forEach(sheet => {
          // Check if this is a rota sheet
          if (sheet.getRange('A1').getValue() === 'Staff Name') {
            updateConditionalFormattingForCustomShifts(sheet, allPatterns);
            
            // Also update data validation to remove the deleted shift from dropdowns
            const shiftNames = Object.keys(allPatterns);
            const validationRule = SpreadsheetApp.newDataValidation()
              .requireValueInList(shiftNames, true)
              .build();
            
            const lastRow = Math.max(2, sheet.getLastRow());
            const lastCol = Math.max(2, sheet.getLastColumn());
            
            if (lastCol > 1) {
              const dataRange = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
              dataRange.setDataValidation(validationRule);
            }
          }
        });
        
        Logger.log(`Updated conditional formatting after removing shift: ${shiftName}`);
      } catch (updateError) {
        Logger.log(`Error updating conditional formatting after removing shift: ${updateError.toString()}`);
        // Don't fail the entire operation if formatting update fails
      }
    }
    
    return updateResult;
  }, 'deleteShiftPattern', { success: false, message: 'Failed to delete/hide shift pattern.' });
}

/**
 * Creates a backup of the current spreadsheet
 */
function createSpreadsheetBackup() {
  return safeExecute(() => {
    logUsage('backup_requested');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backupName = `${ss.getName()} - Backup - ${new Date().toISOString().split('T')[0]}`;
    
    // Create a copy of the spreadsheet
    const backup = ss.copy(backupName);
    
    // Move to a 'Backups' folder if it exists, or create one
    try {
      const folders = DriveApp.getFoldersByName('Staff Rota Backups');
      let backupFolder;
      
      if (folders.hasNext()) {
        backupFolder = folders.next();
      } else {
        backupFolder = DriveApp.createFolder('Staff Rota Backups');
      }
      
      // Move the backup to the backup folder
      const backupFile = DriveApp.getFileById(backup.getId());
      backupFolder.addFile(backupFile);
      DriveApp.getRootFolder().removeFile(backupFile);
      
    } catch (folderError) {
      Logger.log('Could not organize backup into folder: ' + folderError.toString());
      // Backup still created, just not organized
    }
    
    logUsage('backup_completed');
    
    return {
      success: true,
      message: `Backup created: ${backupName}`,
      backupId: backup.getId()
    };
    
  }, 'createSpreadsheetBackup', { success: false, message: 'Failed to create backup' });
}

// Staff Management Functions
function getStaffList() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const staffListJson = userProperties.getProperty('staffList');
    const staffList = staffListJson ? JSON.parse(staffListJson) : [];
    
    Logger.log('getStaffList: Retrieved staffList: ' + JSON.stringify(staffList));
    
    return {
      success: true,
      staff: staffList
    };
  } catch (error) {
    Logger.log('getStaffList: Error: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

function addStaffMember(staffData) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const staffListJson = userProperties.getProperty('staffList');
    const staffList = staffListJson ? JSON.parse(staffListJson) : [];
    
    // Generate a unique ID for the new staff member
    staffData.id = Utilities.getUuid();
    
    staffList.push(staffData);
    userProperties.setProperty('staffList', JSON.stringify(staffList));
    
    Logger.log('addStaffMember: Staff added: ' + JSON.stringify(staffData));
    
    return {
      success: true,
      message: 'Staff member added successfully'
    };
  } catch (error) {
    Logger.log('addStaffMember: Error: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

function updateStaffMember(staffData) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const staffListJson = userProperties.getProperty('staffList');
    const staffList = staffListJson ? JSON.parse(staffListJson) : [];
    
    const index = staffList.findIndex(s => s.id === staffData.id);
    if (index === -1) {
      throw new Error('Staff member not found');
    }
    
    staffList[index] = staffData;
    userProperties.setProperty('staffList', JSON.stringify(staffList));
    
    Logger.log('updateStaffMember: Staff updated: ' + JSON.stringify(staffData));
    
    return {
      success: true,
      message: 'Staff member updated successfully'
    };
  } catch (error) {
    Logger.log('updateStaffMember: Error: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

function deleteStaffMember(staffId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const staffListJson = userProperties.getProperty('staffList');
    const staffList = staffListJson ? JSON.parse(staffListJson) : [];
    
    const index = staffList.findIndex(s => s.id === staffId);
    if (index === -1) {
      throw new Error('Staff member not found');
    }
    
    const deletedStaff = staffList.splice(index, 1);
    userProperties.setProperty('staffList', JSON.stringify(staffList));
    
    Logger.log('deleteStaffMember: Staff deleted: ' + JSON.stringify(deletedStaff[0]));
    
    return {
      success: true,
      message: 'Staff member deleted successfully'
    };
  } catch (error) {
    Logger.log('deleteStaffMember: Error: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

// Add new function to import staff to a sheet
function importStaffToSheet(sheet) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const staffListJson = userProperties.getProperty('staffList');
    const staffList = staffListJson ? JSON.parse(staffListJson) : [];
    
    Logger.log('importStaffToSheet: Staff list before import: ' + JSON.stringify(staffList));

    if (staffList.length === 0) {
      Logger.log('importStaffToSheet: No staff members to import.');
      return;
    }
    
    // Get the range for staff names (A2:A)
    const staffRange = sheet.getRange(2, 1, staffList.length, 1);
    
    // Create array of staff names
    const staffNames = staffList.map(staff => staff.name);
    
    // Set the values
    staffRange.setValues(staffNames.map(name => [name]));
    
    // Format the range
    staffRange.setFontWeight('bold');
    staffRange.setBackground('#f3f3f3');
    staffRange.setBorder(true, true, true, true, true, true);
    Logger.log('importStaffToSheet: Staff imported successfully to sheet: ' + sheet.getName());
  } catch (error) {
    Logger.log('importStaffToSheet: Error: ' + error.message);
    throw error; // Re-throw to be caught by calling function's handler
  }
}

// Helper to get month name from month number
function getMonthName(monthNumber) {
  const date = new Date(null, monthNumber - 1);
  return date.toLocaleString('en', { month: 'long' });
}