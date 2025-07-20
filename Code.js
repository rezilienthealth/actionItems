// ActionItems - Unified Task and Order Management System
// Created: July 18, 2025

// Constants
// Ensure this ID matches your actual ActionItems spreadsheet
const SPREADSHEET_ID = '1Kq1v6W6zjOU84gqy1UQSLew7-KJu4a7rF2CrF9O7sLU';
Logger.log('Using spreadsheet ID: ' + SPREADSHEET_ID); // ActionItems spreadsheet

// Test spreadsheet access on initialization
try {
  const testAccess = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('✅ Spreadsheet access verified on initialization: ' + testAccess.getName());
} catch (e) {
  Logger.log('❌ CRITICAL ERROR: Cannot access spreadsheet on initialization: ' + e.toString());
}
const USER_SPREADSHEET_ID = "1-2sMd80EC1v3KX7oItxe6koqGCvWDCxMCUr8_mV0GkQ"; // Shared user sheet
const ALLOWED_DOMAINS = ['rezilienthealth.com', 'dynamicsurgical.com'];
Logger.log('Allowed domains: ' + ALLOWED_DOMAINS);
const VALID_ROLES = ['admin', 'provider', 'staff', 'user'];
Logger.log('Valid roles: ' + VALID_ROLES);

// Sheet names
const SHEET_NAMES = {
  ACTION_ITEMS: 'actionItems',
  ACTION_ITEM_OPTIONS: 'actionItemOptions',
  ACTION_ITEM_AUDIT: 'actionItemsAudit',
  ACTION_ITEM_COMMENTS: 'actionItemComments',
  USERS: 'authorizedUsers',
  NOTIFICATION_GROUPS: 'notificationGroups',
  GROUP_MEMBERSHIPS: 'groupMemberships'
};

// Cache for spreadsheet data to improve performance
const CACHE_EXPIRATION_SECONDS = 300; // 5 minutes
let cachedItems = null;
let cachedOptions = null;
let lastItemsCacheTime = 0;
let lastOptionsCacheTime = 0;

/**
 * Serves HTML content for the web app
 * @param {Object} e The event parameter for a web app request
 * @returns {HtmlOutput} HTML output
 */
function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();

  // Check if the user is authenticated and belongs to the allowed domains
  if (!userEmail || !ALLOWED_DOMAINS.some(domain => userEmail.endsWith("@" + domain))) {
    return HtmlService.createHtmlOutput(`
      <h2>Access Denied</h2>
      <p>You must log in with an authorized account to access this page.</p>
    `);
  }

  // Log the access
  logAuditEvent("Page Access");

  // Render the main UI
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ActionItems')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Includes an HTML file in the output
 * @param {string} filename - The name of the file to include
 * @returns {string} The contents of the file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets server logs for debugging
 * @returns {string} Server logs
 */
function getServerLogs() {
  try {
    Logger.log('getServerLogs called');
    
    // Test spreadsheet access
    Logger.log('Testing spreadsheet access...');
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      Logger.log('✅ Spreadsheet access successful: ' + ss.getName());
      Logger.log('Available sheets: ' + ss.getSheets().map(s => s.getName()).join(', '));
      
      // Check actionItemOptions sheet
      const optionsSheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_OPTIONS);
      if (optionsSheet) {
        Logger.log('ActionItemOptions sheet found');
        Logger.log('Dimensions: ' + optionsSheet.getLastRow() + ' rows, ' + optionsSheet.getLastColumn() + ' columns');
        
        if (optionsSheet.getLastRow() > 1) {
          const headers = optionsSheet.getRange(1, 1, 1, optionsSheet.getLastColumn()).getValues()[0];
          Logger.log('Headers: ' + JSON.stringify(headers));
          
          // Check for data rows
          if (optionsSheet.getLastRow() > 1) {
            const firstDataRow = optionsSheet.getRange(2, 1, 1, optionsSheet.getLastColumn()).getValues()[0];
            Logger.log('First data row: ' + JSON.stringify(firstDataRow));
          } else {
            Logger.log('No data rows found in ActionItemOptions sheet');
          }
        } else {
          Logger.log('ActionItemOptions sheet is empty');
        }
      } else {
        Logger.log('ActionItemOptions sheet not found');
      }
      
      // Check actionItems sheet
      const itemsSheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEMS);
      if (itemsSheet) {
        Logger.log('ActionItems sheet found');
        Logger.log('Dimensions: ' + itemsSheet.getLastRow() + ' rows, ' + itemsSheet.getLastColumn() + ' columns');
        
        if (itemsSheet.getLastRow() > 1) {
          const headers = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
          Logger.log('Headers: ' + JSON.stringify(headers));
          
          // Check for data rows
          if (itemsSheet.getLastRow() > 1) {
            const firstDataRow = itemsSheet.getRange(2, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
            Logger.log('First data row: ' + JSON.stringify(firstDataRow));
          } else {
            Logger.log('No data rows found in ActionItems sheet');
          }
        } else {
          Logger.log('ActionItems sheet is empty');
        }
      } else {
        Logger.log('ActionItems sheet not found');
      }
      
    } catch (e) {
      Logger.log('❌ CRITICAL ERROR accessing spreadsheet: ' + e.toString());
      Logger.log('Error stack: ' + e.stack);
    }
    
    // Get logs
    const logs = Logger.getLog();
    return logs;
  } catch (error) {
    console.error('Error getting server logs:', error);
    return 'Error getting server logs: ' + error.toString();
  }
}

/**
 * Logs an audit event
 * @param {string} action - The action being performed
 */
function logAuditEvent(action) {
  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("auditTrail");
  if (sheet) {
    sheet.appendRow([timestamp, userEmail, action]);
  }
}

/**
 * Gets the current user's email address
 * @returns {string} The user's email address
 */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Gets the current user's role
 * @returns {string} The user's role
 */
function getUserRole() {
  const email = getUserEmail();
  const ss = SpreadsheetApp.openById(USER_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("authorizedUsers");
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf("email");
  const roleCol = headers.indexOf("role");
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] === email) {
      return data[i][roleCol];
    }
  }
  
  return "user"; // Default role
}

/**
 * Checks if the current user is a provider
 * @returns {boolean} True if the user is a provider
 */
function isProvider() {
  const role = getUserRole();
  return role === "provider" || role === "admin";
}

/**
 * Checks if the current user is an admin
 * @returns {boolean} True if the user is an admin
 */
function isAdmin() {
  return getUserRole() === "admin";
}

/**
 * Final, foolproof wrapper to be called by the client. This function retrieves action items,
 * explicitly converts any Date objects into ISO strings to prevent serialization timeouts,
 * and returns a clean, safe array to the client.
 */
function getAndProcessClientActionItems() {
  let items = getActionItems();
  if (!items || items.length === 0) {
    Logger.log('getAndProcessClientActionItems: No items received from getActionItems. Returning empty array.');
    return [];
  }

  try {
    Logger.log('getAndProcessClientActionItems: Received ' + items.length + ' items. Starting date conversion.');
    // The root cause of timeouts is often non-serializable data types like Date objects.
    // We must iterate through the items and convert any Date objects to a string format (ISO).
    items.forEach((item, index) => {
      for (const key in item) {
        if (item[key] instanceof Date) {
          item[key] = item[key].toISOString();
        }
      }
    });
    Logger.log('Successfully converted all Date objects to ISO strings for serialization.');
    return items;
  } catch (e) {
    Logger.log('FATAL ERROR during date serialization in getAndProcessClientActionItems: ' + e.toString() + ' Stack: ' + e.stack);
    // If serialization itself fails, we can't safely return the data.
    return []; 
  }
}

function getClientActionItems() {
  const items = getActionItems();
  return items || [];
}

/**
 * Gets all action items
 * @returns {Array} Array of action items
 */
function getActionItems() {
    try {
        Logger.log('getActionItems called - Reading from Google Sheet');
        
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        Logger.log('Successfully opened spreadsheet for items: ' + ss.getName());
        
        const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEMS);
        if (!sheet) {
            Logger.log('CRITICAL ERROR: "actionItems" sheet not found.');
            return [];
        }
        Logger.log('Successfully found sheet: ' + sheet.getName());
        Logger.log('Sheet dimensions: ' + sheet.getLastRow() + ' rows, ' + sheet.getLastColumn() + ' columns');

        const data = sheet.getDataRange().getValues();
        Logger.log('Retrieved data from sheet: ' + data.length + ' rows');

        if (data.length < 2) {
            Logger.log('No data rows found in actionItems sheet.');
            return [];
        }

        const headers = data[0].map(header => header.toString().trim());
        Logger.log('Headers processed: ' + JSON.stringify(headers));

        const items = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const isEmpty = row.every(cell => cell === '');
            if (isEmpty) {
                Logger.log('Skipping empty row at sheet row ' + (i + 1));
                continue;
            }

            let item = {};
            row.forEach((cell, j) => {
                const header = headers[j];
                if (header) {
                    item[header] = cell;
                }
            });
            item.row = i + 1; // Use sheet row number for clarity
            items.push(item);
        }
        Logger.log('Successfully processed ' + items.length + ' items from sheet.');

        // Update cache
        cachedItems = items;
        lastItemsCacheTime = new Date().getTime();
        Logger.log('Cache updated with ' + items.length + ' items.');

        return items;
    } catch (e) {
        Logger.log('FATAL ERROR in getActionItems: ' + e.toString() + ' Stack: ' + e.stack);
        return []; // Ensure an array is always returned on error
    }
}

/**
 * Gets action item options (wrapper for the version in ActionItemOptions.js)
 * @returns {Object} Hierarchical object of action item options
 */
function getActionItemOptions() {
  try {
    Logger.log('SIMPLE LOG: getActionItemOptions in Code.js called');
    const options = buildActionItemOptions(true);
    if (!options || !options.actionItems) {
      Logger.log('ERROR: buildActionItemOptions did not return a valid options object.');
      return { actionItems: { categories: {}, selectionTypes: {} } };
    }
    return options;
  } catch (error) {
    Logger.log('ERROR in getActionItemOptions wrapper: ' + error.toString() + ' Stack: ' + error.stack);
    console.error('Error getting action item options:', error);
    return null; // Return null or a safe default on error
  }
}

/**
 * Exposes a function to the client to clear the server-side cache.
 */
function clearActionItemOptionsCache() {
  Logger.log('SIMPLE LOG: Manually clearing action item options cache.');
  clearCachedOptions();
  return 'Cache cleared successfully.';
}

/**
 * Client-callable function to get action item options
 * This function is specifically designed to be called from the client
 * @returns {Object} Hierarchical object of action item options
 */
function getClientActionItemOptions() {
  Logger.log('SIMPLE LOG: getClientActionItemOptions called from client');
  return getActionItemOptions();
}

/**
 * Simple test function to verify client-server communication
 * @returns {string} A simple test message
 */
function testClientServerCommunication() {
  Logger.log('SIMPLE LOG: testClientServerCommunication called');
  return 'Communication successful at ' + new Date().toString();
}

/**
 * Gets action item templates
 * @returns {Array} Array of action item templates
 */
function getActionItemTemplates() {
  try {
    const items = getActionItems();
    return items.filter(item => item.isTemplate === true);
  } catch (error) {
    console.error('Error getting action item templates:', error);
    throw error;
  }
}

/**
 * Gets action item options (legacy version)
 * @returns {Object} Hierarchical object of action item options
 * @deprecated Use the version in ActionItemOptions.js instead
 */
function getActionItemOptionsLegacy() {
  try {
    Logger.log('getActionItemOptions called');
    
    // Debug spreadsheet access
    try {
      Logger.log('Attempting to open spreadsheet with ID: ' + SPREADSHEET_ID);
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      Logger.log('Successfully opened spreadsheet: ' + ss.getName());
      
      const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_OPTIONS);
      if (sheet) {
        Logger.log('Successfully found sheet: ' + sheet.getName());
        Logger.log('Sheet dimensions: ' + sheet.getLastRow() + ' rows, ' + sheet.getLastColumn() + ' columns');
      } else {
        Logger.log('ERROR: Could not find sheet named: ' + SHEET_NAMES.ACTION_ITEM_OPTIONS);
        Logger.log('Available sheets: ' + ss.getSheets().map(s => s.getName()).join(', '));
      }
    } catch (e) {
      Logger.log('CRITICAL ERROR accessing spreadsheet: ' + e.toString());
      Logger.log('Error stack: ' + e.stack);
      Logger.log('ERROR accessing spreadsheet: ' + e.toString());
    }
    // Check cache first
    const now = new Date().getTime();
    if (cachedOptions && (now - lastOptionsCacheTime < CACHE_EXPIRATION_SECONDS * 1000)) {
      Logger.log('Returning cached options');
      return cachedOptions;
    }
    
    Logger.log('Cache miss or expired, fetching fresh options data');
    
    // Use the correct spreadsheet ID constant
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // This is the correct constant defined at the top
    Logger.log('Using spreadsheet ID for options: ' + SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_OPTIONS);
    Logger.log('Looking for sheet named: ' + SHEET_NAMES.ACTION_ITEM_OPTIONS);
    
    if (!sheet) {
      throw new Error("Action item options sheet not found");
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Create column map for debugging
    const columnMap = {};
    headers.forEach((header, index) => {
      columnMap[header] = index;
    });
    Logger.log('Column map: ' + JSON.stringify(columnMap));
    
    // Check if we have the necessary columns
    const requiredColumns = ['categoryLevel1', 'active'];
    const missingColumns = requiredColumns.filter(col => !(col in columnMap));
    if (missingColumns.length > 0) {
      Logger.log('ERROR: Missing required columns: ' + missingColumns.join(', '));
    }
    
    // Build hierarchical options using unified actionItems structure
    const options = {
      actionItems: {
        categories: {},
        selectionTypes: {}
      }
    };
    
    Logger.log('Data rows count: ' + data.length);
    Logger.log('Headers: ' + JSON.stringify(headers));
    
    for (let i = 1; i < data.length; i++) {
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];
        const value = data[i][j];

        if (value instanceof Date) {
          row[header] = value.toISOString();
        } else if (typeof value === 'boolean') {
          row[header] = value;
        } else if (String(value).toUpperCase() === 'TRUE') {
          row[header] = true;
        } else if (String(value).toUpperCase() === 'FALSE') {
          row[header] = false;
        } else {
          row[header] = value;
        }
      }
      
      // Log the active value for debugging
      Logger.log('Item active value: ' + row.active + ' (type: ' + typeof row.active + ')');
      
      // Skip inactive items
      if (row.active === false || row.active === 'FALSE' || 
          (typeof row.active === 'string' && row.active.toUpperCase() !== 'TRUE')) {
        Logger.log('Skipping inactive item: ' + row.categoryLevel1);
        continue;
      }
      
      // Add debug logging for active items
      Logger.log('Found active item: ' + JSON.stringify(row));
      
      // Get category levels
      const categoryLevel1 = row.categoryLevel1 || '';
      const categoryLevel2 = row.categoryLevel2 || '';
      const categoryLevel3 = row.categoryLevel3 || '';
      const categoryLevel4 = row.categoryLevel4 || '';
      const categoryLevel5 = row.categoryLevel5 || '';
      
      // Skip if no category level 1
      if (!categoryLevel1) continue;
      
      // Build the category path for logging
      const categoryPath = [categoryLevel1, categoryLevel2, categoryLevel3, categoryLevel4, categoryLevel5]
        .filter(Boolean).join('/');
      Logger.log('Processing item with category path: ' + categoryPath);
      
      // Build hierarchical structure
      if (!options.actionItems.categories[categoryLevel1]) {
        options.actionItems.categories[categoryLevel1] = {
          displayName: categoryLevel1,
          options: [],
          subcategories: {}
        };
      }
      
      if (categoryLevel2) {
        if (!options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2]) {
          options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2] = {
            displayName: categoryLevel2,
            options: [],
            subcategories: {}
          };
        }
        
        if (categoryLevel3) {
          if (!options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3]) {
            options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3] = {
              displayName: categoryLevel3,
              options: [],
              subcategories: {}
            };
          }
          
          if (categoryLevel4) {
            if (!options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4]) {
              options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4] = {
                displayName: categoryLevel4,
                options: [],
                subcategories: {}
              };
            }
            
            if (categoryLevel5) {
              if (!options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4].subcategories[categoryLevel5]) {
                options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4].subcategories[categoryLevel5] = {
                  displayName: categoryLevel5,
                  options: []
                };
              }
              
              // Add option at level 5
              options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4].subcategories[categoryLevel5].options.push(row);
            } else {
              // Add option at level 4
              options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4].options.push(row);
            }
          } else {
            // Add option at level 3
            options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].options.push(row);
          }
        } else {
          // Add option at level 2
          options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].options.push(row);
        }
      } else {
        // Add option at level 1
        options.actionItems.categories[categoryLevel1].options.push(row);
      }
      
      // Set selection type
      const selectionType = row.selectionType || 'single';
      if (categoryLevel1) {
        if (!options.actionItems.selectionTypes[categoryLevel1]) {
          options.actionItems.selectionTypes[categoryLevel1] = {};
        }
        if (categoryLevel2) {
          if (!options.actionItems.selectionTypes[categoryLevel1][categoryLevel2]) {
            options.actionItems.selectionTypes[categoryLevel1][categoryLevel2] = {};
          }
          if (categoryLevel3) {
            options.actionItems.selectionTypes[categoryLevel1][categoryLevel2][categoryLevel3] = selectionType;
          }
        }
      }
    }
    
    // Log the final options structure
    Logger.log('Final options structure: ' + JSON.stringify(options));
    Logger.log('Categories count: ' + Object.keys(options.actionItems.categories).length);
    
    // Log the final options structure before returning
    Logger.log('Final options structure keys: ' + JSON.stringify(Object.keys(options)));
    Logger.log('Final actionItems keys: ' + JSON.stringify(Object.keys(options.actionItems)));
    Logger.log('Final categories count: ' + Object.keys(options.actionItems.categories).length);
    
    // Return options
    cachedOptions = options;
    lastOptionsCacheTime = now;
    return options;
  } catch (error) {
    Logger.log('ERROR getting action item options: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    console.error('Error getting action item options:', error);
    return null;
  }
}

/**
 * Helper function to merge objects recursively
 * @param {Object} target - Target object
 * @param {Object} source - Source object
 */
function mergeObjects(target, source) {
  for (const key in source) {
    if (typeof source[key] === 'object' && !Array.isArray(source[key])) {
      if (!target[key]) target[key] = {};
      mergeObjects(target[key], source[key]);
    } else {
      target[key] = source[key];
    }
  }
}

/**
 * Gets an action item by ID
 * @param {string} id - Action item ID
 * @returns {Object} Action item
 */
function getActionItemById(id) {
  try {
    const items = getActionItems();
    return items.find(item => item.actionItemId === id);
  } catch (error) {
    console.error('Error getting action item by ID:', error);
    throw error;
  }
}

/**
 * Gets comments for an action item
 * @param {string} itemId - Action item ID
 * @returns {Array} Array of comments
 */
function getCommentsForActionItem(itemId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_COMMENTS);
    
    if (!sheet) {
      throw new Error("Comments sheet not found");
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Convert to array of objects
    const comments = [];
    for (let i = 1; i < data.length; i++) {
      const comment = {};
      for (let j = 0; j < headers.length; j++) {
        comment[headers[j]] = data[i][j];
      }
      
      if (comment.actionItemId === itemId) {
        comments.push(comment);
      }
    }
    
    // Sort by timestamp
    comments.sort((a, b) => {
      const dateA = new Date(a.timestamp);
      const dateB = new Date(b.timestamp);
      return dateA - dateB;
    });
    
    return comments;
  } catch (error) {
    console.error('Error getting comments:', error);
    throw error;
  }
}

/**
 * Gets history for an action item
 * @param {string} itemId - Action item ID
 * @returns {Array} Array of history events
 */
function getActionItemHistory(itemId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_AUDIT);
    
    if (!sheet) {
      throw new Error("Audit sheet not found");
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Convert to array of objects
    const history = [];
    for (let i = 1; i < data.length; i++) {
      const event = {};
      for (let j = 0; j < headers.length; j++) {
        event[headers[j]] = data[i][j];
      }
      
      if (event.actionItemId === itemId) {
        history.push(event);
      }
    }
    
    // Sort by timestamp
    history.sort((a, b) => {
      const dateA = new Date(a.changedAt);
      const dateB = new Date(b.changedAt);
      return dateB - dateA; // Newest first
    });
    
    return history;
  } catch (error) {
    console.error('Error getting history:', error);
    throw error;
  }
}

/**
 * Gets users for mentions
 * @returns {Array} Array of users
 */
function getUsers() {
  try {
    const ss = SpreadsheetApp.openById(USER_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const groupSheet = ss.getSheetByName(SHEET_NAMES.NOTIFICATION_GROUPS);
    
    if (!userSheet) {
      throw new Error("Users sheet not found");
    }
    
    const userData = userSheet.getDataRange().getValues();
    const userHeaders = userData[0];
    
    // Convert to array of objects
    const users = [];
    for (let i = 1; i < userData.length; i++) {
      const user = {};
      for (let j = 0; j < userHeaders.length; j++) {
        user[userHeaders[j]] = userData[i][j];
      }
      
      users.push({
        name: user.name || user.email,
        email: user.email,
        isGroup: false
      });
    }
    
    // Add groups if available
    if (groupSheet) {
      const groupData = groupSheet.getDataRange().getValues();
      const groupHeaders = groupData[0];
      
      for (let i = 1; i < groupData.length; i++) {
        const group = {};
        for (let j = 0; j < groupHeaders.length; j++) {
          group[groupHeaders[j]] = groupData[i][j];
        }
        
        users.push({
          name: group.groupName,
          email: group.groupId,
          isGroup: true
        });
      }
    }
    
    return users;
  } catch (error) {
    console.error('Error getting users:', error);
    throw error;
  }
}
