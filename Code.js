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
 * Shows the webhook test UI
 * @returns {HtmlOutput} The HTML content for the webhook test UI
 */
// Expose test webhook functions to the client
if (typeof module !== 'undefined') {
  module.exports = {
    testWebhookNotification: testWebhookNotification,
    getUserWebhookDebugInfo: getUserWebhookDebugInfo
  };
}

function showWebhookTestUI() {
  return HtmlService.createHtmlOutputFromFile('WebhookTestUI')
    .setWidth(600)
    .setHeight(800)
    .setTitle('Test Webhook Notifications');
}

/**
 * Tests webhook notification for a user
 * @param {string} email - The email of the user to test
 * @param {string} message - Optional custom test message
 * @returns {Object} Result of the test
 */
function testWebhookNotification(email, message = 'This is a test notification') {
  try {
    const user = getUserByEmail(email);
    if (!user) {
      return { success: false, error: 'User not found' };
    }
    
    const webhookUrl = user.webhookUrl || user.Chatwebhook;
    if (!webhookUrl) {
      return { success: false, error: 'No webhook URL configured for this user' };
    }
    
    const testMessage = {
      text: `🔔 *Test Notification*\n${message}\n\n_This is a test notification from ActionItems._`,
      cards: [{
        header: {
          title: 'Test Notification',
          subtitle: 'ActionItems System',
          imageUrl: 'https://www.gstatic.com/images/branding/product/1x/chat_48dp.png'
        },
        sections: [{
          widgets: [{
            textParagraph: {
              text: message
            }
          }]
        }]
      }]
    };
    
    const response = sendWebhookNotification(webhookUrl, testMessage);
    return {
      success: true,
      message: 'Test notification sent successfully',
      response: response
    };
  } catch (error) {
    console.error('Error testing webhook notification:', error);
    return {
      success: false,
      error: error.message || 'Failed to send test notification'
    };
  }
}

/**
 * Gets debug information about a user's webhook configuration
 * @param {string} email - The email of the user to get debug info for
 * @returns {Object} Debug information about the user's webhook configuration
 */
function getUserWebhookDebugInfo(email) {
  try {
    const user = getUserByEmail(email);
    if (!user) {
      return { 
        success: false, 
        error: 'User not found',
        timestamp: new Date().toISOString()
      };
    }
    
    const webhookUrl = user.webhookUrl || user.Chatwebhook;
    const hasWebhook = !!webhookUrl;
    
    // Check if webhook URL is a valid URL
    let isValidUrl = false;
    try {
      if (webhookUrl) {
        new URL(webhookUrl);
        isValidUrl = true;
      }
    } catch (e) {
      // URL is invalid
      isValidUrl = false;
    }
    
    // Check if user is in any notification groups
    const notificationGroups = [];
    const allGroups = getUsers().filter(u => u.isGroup);
    allGroups.forEach(group => {
      if (group.members && group.members.includes(email)) {
        notificationGroups.push({
          groupName: group.name || group.groupName || group.email,
          groupId: group.groupId || group.email,
          hasWebhook: !!(group.webhookUrl || group.Chatwebhook)
        });
      }
    });
    
    return {
      success: true,
      user: {
        email: user.email,
        name: user.name || `${user.firstName || ''} ${user.lastName || ''}`.trim(),
        hasDirectWebhook: hasWebhook,
        webhookUrl: hasWebhook ? webhookUrl : null,
        webhookUrlValid: isValidUrl,
        isActive: user.active !== false,
        role: user.role || 'user',
        lastUpdated: user.lastUpdated || 'Unknown',
        notificationGroups: notificationGroups
      },
      systemInfo: {
        timestamp: new Date().toISOString(),
        totalUsers: getUsers().filter(u => !u.isGroup).length,
        totalGroups: getUsers().filter(u => u.isGroup).length,
        webhookEndpoint: 'https://chat.googleapis.com/v1/spaces/.../messages'
      }
    };
  } catch (error) {
    console.error('Error getting user webhook debug info:', error);
    return {
      success: false,
      error: error.message || 'Failed to get debug information',
      timestamp: new Date().toISOString()
    };
  }
}

// Make functions available to the client
var clientFunctions = {
  testWebhookNotification: testWebhookNotification,
  getUserWebhookDebugInfo: getUserWebhookDebugInfo,
  showWebhookTestUI: showWebhookTestUI,
  testClientServerCommunication: testClientServerCommunication
};

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
          title: categoryLevel1,
          options: [],
          subcategories: {}
        };
      }
      
      if (categoryLevel2) {
        if (!options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2]) {
          options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2] = {
            title: categoryLevel2,
            options: [],
            subcategories: {}
          };
        }
        
        if (categoryLevel3) {
          if (!options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3]) {
            options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3] = {
              title: categoryLevel3,
              options: [],
              subcategories: {}
            };
          }
          
          if (categoryLevel4) {
            if (!options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4]) {
              options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4] = {
                title: categoryLevel4,
                options: [],
                subcategories: {}
              };
            }
            
            if (categoryLevel5) {
              if (!options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4].subcategories[categoryLevel5]) {
                options.actionItems.categories[categoryLevel1].subcategories[categoryLevel2].subcategories[categoryLevel3].subcategories[categoryLevel4].subcategories[categoryLevel5] = {
                  title: categoryLevel5,
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
    const comments = [];
    
    // Find the index of the comment ID column
    const commentIdIndex = headers.indexOf('commentId');
    const actionItemIdIndex = headers.indexOf('actionItemId');
    const authorIndex = headers.indexOf('author');
    const contentIndex = headers.indexOf('content');
    const timestampIndex = headers.indexOf('timestamp');
    const mentionedUsersIndex = headers.indexOf('mentionedUsers');
    
    // If any required columns are missing, return empty array
    if (commentIdIndex === -1 || actionItemIdIndex === -1 || authorIndex === -1 || 
        contentIndex === -1 || timestampIndex === -1) {
      console.error('Required columns not found in comments sheet');
      return [];
    }
    
    // Filter comments for the specified action item
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[actionItemIdIndex] === itemId) {
        comments.push({
          commentId: row[commentIdIndex],
          actionItemId: row[actionItemIdIndex],
          author: row[authorIndex],
          content: row[contentIndex],
          timestamp: row[timestampIndex],
          mentionedUsers: mentionedUsersIndex !== -1 ? (row[mentionedUsersIndex] || '') : ''
        });
      }
    }
    
    return comments;
  } catch (error) {
    console.error('Error getting comments for action item:', error);
    Logger.log('ERROR in getCommentsForActionItem: ' + error.toString());
    return [];
  }
}

/**
 * Deletes a user or group
 * @param {string} email - The email or group ID to delete
 * @returns {Object} Result of the operation
 */
function deleteUser(email) {
  try {
    if (!email) {
      throw new Error('Email/Group ID is required');
    }
    
    const ss = SpreadsheetApp.openById(USER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    if (!sheet) {
      throw new Error('Users sheet not found');
    }
    
    // Get all data
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    
    // Find email/groupId column
    const emailCol = headers.indexOf('email');
    const groupIdCol = headers.indexOf('groupId');
    
    if (emailCol === -1 && groupIdCol === -1) {
      throw new Error('No identifier columns found in users sheet');
    }
    
    // Find user/group row
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if ((emailCol !== -1 && values[i][emailCol] === email) || 
          (groupIdCol !== -1 && values[i][groupIdCol] === email)) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('User/Group not found: ' + email);
    }
    
    // If this is a group, remove it from all members
    if (values[rowIndex][headers.indexOf('isGroup')] === true) {
      updateGroupMembers(email, []); // Remove all members
    }
    
    // Delete the row
    sheet.deleteRow(rowIndex + 1); // +1 because sheet rows are 1-based
    
    console.log('User/Group deleted successfully');
    return { success: true, message: 'User/Group deleted successfully' };
    
  } catch (error) {
    console.error('Error deleting user/group:', error);
    Logger.log('ERROR in deleteUser: ' + error.toString());
    throw error;
  }
}

/**
 * Saves a user or group
 * @param {Object} user - User or group data to save
 * @returns {Object} Result of the operation
 */
function saveUser(user) {
  console.log('saveUser called with user:', JSON.stringify(user, null, 2));
  
  try {
    const ss = SpreadsheetApp.openById(USER_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    if (!userSheet) {
      const errorMsg = "Users sheet not found";
      console.error(errorMsg);
      throw new Error(errorMsg);
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    let rowIndex = -1;
    
    console.log('Headers:', headers);
    console.log('First data row:', data[1] ? data[1].join(', ') : 'No data rows');
    
    // Determine if this is a group
    const isGroup = user.isGroup || false;
    console.log('Is group:', isGroup);
    const idField = isGroup ? 'groupId' : 'email';
    const idValue = isGroup ? (user.email || user.groupId) : user.email;
    
    // Find existing user/group by email/groupId
    const idIndex = headers.indexOf(isGroup ? 'groupId' : 'email');
    if (idIndex !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] === idValue) {
          rowIndex = i + 1; // +1 because array is 0-based but sheet rows are 1-based
          break;
        }
      }
    }
    
    // Prepare user/group data in the correct order based on headers
    const userData = [];
    console.log('Preparing user data for save. Headers:', headers);
    console.log('User data to save:', JSON.stringify(user, null, 2));
    
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      
      // Special handling for group fields
      if (isGroup) {
        if (header === 'groupId') userData.push(idValue);
        else if (header === 'groupName') userData.push(user.name || '');
        else if (header === 'members' && user.members) userData.push(Array.isArray(user.members) ? user.members.join(',') : user.members);
        else if (header === 'webhookUrl' || header === 'chatwebhook') userData.push(user.webhookUrl || user.chatwebhook || '');
        else if (header === 'role') userData.push('group');
        else if (header === 'isGroup') userData.push(true);
        else if (header === 'active') userData.push(user.active !== undefined ? user.active : true);
        else if (header === 'email') userData.push(''); // Leave email empty for groups
        else userData.push('');
      } 
      // Regular user fields
      else {
        // Generate displayName if not provided
        const displayName = user.displayName || 
                          (user.firstName && user.lastName ? `${user.firstName} ${user.lastName}`.trim() : 
                          user.name || (user.email ? user.email.split('@')[0] : ''));
        
        if (header === 'email') {
          const emailValue = user.email || '';
          console.log(`Setting email header '${header}' to:`, emailValue);
          userData.push(emailValue);
        }
        else if (header === 'firstName') userData.push(user.firstName || user.name?.split(' ')[0] || '');
        else if (header === 'lastName') userData.push(user.lastName || user.name?.split(' ').slice(1).join(' ') || '');
        else if (header === 'displayName') userData.push(displayName);
        else if (header === 'role') userData.push((user.role || 'user').toLowerCase());
        else if (header === 'webhookUrl' || header === 'chatwebhook') {
          const webhookValue = user.webhookUrl || user.chatwebhook || '';
          console.log(`Setting webhook for header '${header}' to:`, webhookValue);
          userData.push(webhookValue);
        }
        else if (header === 'isGroup') userData.push(false);
        else if (header === 'active') {
          // Handle different active field names for backward compatibility
          const activeValue = user.active !== undefined ? user.active : 
                            (user.Isactive !== undefined ? user.Isactive : true);
          userData.push(activeValue);
        }
        else if (header === 'groups') userData.push(Array.isArray(user.groups) ? user.groups.join(',') : user.groups || '');
        // Handle legacy name field for backward compatibility
        else if (header === 'name') userData.push(displayName);
        else userData.push('');
      }
    }
    
    if (rowIndex > 0) {
      // Update existing user/group
      userSheet.getRange(rowIndex, 1, 1, userData.length).setValues([userData]);
    } else {
      // Add new user/group
      userSheet.appendRow(userData);
    }
    
    // If this is a group with members, update the members' group assignments
    if (isGroup && user.members && user.members.length > 0) {
      updateGroupMembers(idValue, user.members);
    }
    
    return { 
      success: true, 
      message: isGroup ? "Group saved successfully" : "User saved successfully" 
    };
  } catch (error) {
    console.error('Error saving user/group:', error);
    throw error;
  }
}

/**
    console.log('Deleting user/group:', email);
    
    if (!email) {
      throw new Error('Email/Group ID is required');
    }
    
    const ss = SpreadsheetApp.openById(USER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    if (!sheet) {
      throw new Error('Users sheet not found');
    }
    
    // Get all data
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    
    // Find email/groupId column
    const emailCol = headers.indexOf('email');
    const groupIdCol = headers.indexOf('groupId');
    
    if (emailCol === -1 && groupIdCol === -1) {
      throw new Error('No identifier columns found in users sheet');
    }
    
    // Find user/group row
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if ((emailCol !== -1 && values[i][emailCol] === email) || 
          (groupIdCol !== -1 && values[i][groupIdCol] === email)) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('User/Group not found: ' + email);
    }
    
    // If this is a group, remove it from all members
    if (values[rowIndex][headers.indexOf('isGroup')] === true) {
      updateGroupMembers(email, []); // Remove all members
    }
    
    // Delete the row
    sheet.deleteRow(rowIndex + 1); // +1 because sheet rows are 1-based
    
    console.log('User/Group deleted successfully');
    return { success: true, message: 'User/Group deleted successfully' };
    
  } catch (error) {
    console.error('Error deleting user/group:', error);
    Logger.log('ERROR in deleteUser: ' + error.toString());
    throw error;
  }
}

/**
 * Gets users and groups for the system
 * @returns {Array} Array of users and groups
 */
/**
 * Gets a single user by email
 * @param {string} email - The email of the user to find
 * @returns {Object} The user object or null if not found
 */
function getUserByEmail(email) {
  try {
    if (!email) {
      throw new Error('Email is required');
    }
    
    console.log('Searching for user with email:', email);
    const users = getUsers();
    console.log('Total users in system:', users.length);
    
    // Case-insensitive search
    const normalizedEmail = email.toLowerCase().trim();
    const user = users.find(u => {
      const userEmail = (u.email || '').toLowerCase().trim();
      return userEmail === normalizedEmail;
    });
    
    if (!user) {
      console.log('User not found with email:', email);
      console.log('Available emails:', users.map(u => u.email).filter(Boolean));
      return null;
    }
    
    // Ensure email is included in the returned user object
    if (!user.email) {
      user.email = email;
    }
    
    console.log('Found user:', JSON.stringify(user, null, 2));
    return {...user, email: user.email}; // Ensure email is in the returned object
  } catch (error) {
    console.error('Error in getUserByEmail:', error);
    throw error;
  }
}

/**
 * Gets all users and groups from the system
 * @returns {Array} Array of user and group objects
 */
function getUsers() {
  try {
    const ss = SpreadsheetApp.openById(USER_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
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
      
      // Check if this is a group (has groupId and groupName)
      const isGroup = !!(user.groupId && user.groupName);
      const email = isGroup ? user.groupId : user.email;
      const name = isGroup ? user.groupName : (user.name || `${user.firstName || ''} ${user.lastName || ''}`.trim() || user.email);
      
      users.push({
        name: name,
        email: email,
        firstName: user.firstName || '',
        lastName: user.lastName || '',
        displayName: user.displayName || name,
        role: user.role || (isGroup ? 'group' : 'user'),
        active: user.active || user.Isactive || true,
        webhookUrl: user.Chatwebhook || user.webhookUrl || '',
        isGroup: isGroup,
        // For groups, store member emails in the members array
        members: user.members ? user.members.split(',').map(m => m.trim()) : []
      });
    }
    
    return users;
  } catch (error) {
    console.error('Error getting users:', error);
    throw error;
  }
}

/**
 * Gets users as a map for webhook lookups with normalized email keys
 * @returns {Object} Map of lowercase email -> user object
 */
function getUserMap() {
  const startTime = new Date();
  const sessionId = Utilities.getUuid().substring(0, 8);
  
  try {
    console.log(`=== GET USER MAP START [${sessionId}] ===`);
    console.log(`[${sessionId}] Loading users...`);
    
    const users = getUsers();
    const userMap = {};
    let validWebhookCount = 0;
    let userCount = 0;
    
    console.log(`[${sessionId}] Loaded ${Array.isArray(users) ? users.length : 'unknown'} users`);
    
    if (!Array.isArray(users)) {
      console.error('getUsers() did not return an array:', users);
      throw new Error('Failed to load users: Invalid data format');
    }
    
    console.log(`Processing ${users.length} users from getUsers()`);
    
    users.forEach((user, index) => {
      userCount++;
      try {
        if (!user || !user.email) {
          console.warn(`[${sessionId}] Skipping user at index ${index}: Missing email`);
          return;
        }
        
        const email = user.email.toLowerCase().trim();
        
        // Log if this is the user we're looking for
        if (email === 'clinicalinnovation@rezilienthealth.com') {
          console.log(`[${sessionId}] FOUND CLINICALINNOVATION USER:`, JSON.stringify(user, null, 2));
        }
        
        // Only include users with a webhook URL
        if (user.webhookUrl && typeof user.webhookUrl === 'string' && user.webhookUrl.trim()) {
          const webhookUrl = user.webhookUrl.trim();
          const isValidUrl = webhookUrl.startsWith('http');
          
          userMap[email] = {
            ...user,
            email: email,
            webhookUrl: isValidUrl ? webhookUrl : '',
            hasValidWebhook: isValidUrl,
            _debug: {
              processedAt: new Date().toISOString(),
              source: 'getUserMap',
              sessionId: sessionId
            }
          };
          
          if (isValidUrl) {
            validWebhookCount++;
            if (email === 'clinicalinnovation@rezilienthealth.com') {
              console.log(`[${sessionId}] CLINICALINNOVATION WEBHOOK CONFIGURED:`, webhookUrl);
            }
          } else if (email === 'clinicalinnovation@rezilienthealth.com') {
            console.warn(`[${sessionId}] CLINICALINNOVATION HAS INVALID WEBHOOK:`, webhookUrl);
          }
        }
        
        // Create user object with normalized data
        userMap[email] = {
          ...user,
          email: email,
          webhookUrl: webhookUrl
        };
        
        console.log(`Added user: ${email}${webhookUrl ? ' (webhook available)' : ''}`);
      } catch (userError) {
        console.error(`Error processing user at index ${index}:`, userError);
      }
    });
    
    const endTime = new Date();
    const durationMs = endTime - startTime;
    
    console.log(`=== GET USER MAP COMPLETE [${sessionId}] ===`);
    console.log(`[${sessionId}] Processed ${userCount} users in ${durationMs}ms`);
    console.log(`[${sessionId}] - Users with valid webhooks: ${validWebhookCount}/${Object.keys(userMap).length}`);
    
    // Log clinicalinnovation user status
    const clinicalUser = userMap['clinicalinnovation@rezilienthealth.com'];
    if (clinicalUser) {
      console.log(`[${sessionId}] CLINICALINNOVATION USER FOUND:`, {
        email: clinicalUser.email,
        hasWebhook: !!clinicalUser.webhookUrl,
        webhookUrl: clinicalUser.webhookUrl ? '***URL_REDACTED***' : 'MISSING',
        hasValidWebhook: clinicalUser.hasValidWebhook || false
      });
    } else {
      console.warn(`[${sessionId}] CLINICALINNOVATION USER NOT FOUND IN USER MAP`);
    }
    
    return userMap;
  } catch (error) {
    const errorMsg = `[${sessionId}] CRITICAL ERROR in getUserMap: ${error.toString()}\n${error.stack || 'No stack trace'}`;
    console.error(errorMsg);
    Logger.log(errorMsg);
    
    // Return empty map instead of throwing to prevent cascading failures
    return {};
  }
}

// =================================================================================
// NOTIFICATION FUNCTIONS
// =================================================================================

/**
 * Extracts mentioned users from text content
 * @param {string} text - The text content to search for mentions
 * @returns {Array} Array of mentioned user emails
 */
function extractMentionedUsers(text) {
  if (!text) return [];
  
  // Match email addresses in the format @user@example.com
  const mentionRegex = /@([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/g;
  const matches = [];
  let match;
  
  while ((match = mentionRegex.exec(text)) !== null) {
    // Remove the @ symbol and add to matches
    matches.push(match[1]);
  }
  
  return [...new Set(matches)]; // Remove duplicates
}

/**
 * Sends a notification to mentioned users
 * @param {string} content - The content containing mentions
 * @param {string} contextUrl - URL to the relevant item
 * @param {string} author - Email of the user who made the mention
 * @param {string} itemTitle - Title of the item being mentioned in
 */
function notifyMentionedUsers(content, contextUrl, author, itemTitle) {
  try {
    const mentionedEmails = extractMentionedUsers(content);
    if (mentionedEmails.length === 0) return;
    
    const userMap = getUserMap();
    const authorUser = userMap[author] || { name: author, email: author };
    
    mentionedEmails.forEach(email => {
      const user = userMap[email];
      if (!user || !user.webhookUrl) return;
      
      const message = {
        text: `You were mentioned by ${authorUser.name} in "${itemTitle}"`,
        cards: [{
          header: {
            title: `New mention in ${itemTitle}`,
            subtitle: `From: ${authorUser.name}`
          },
          sections: [{
            widgets: [{
              textParagraph: {
                text: content.length > 200 ? content.substring(0, 200) + '...' : content
              }
            }, {
              buttons: [{
                text: 'View',
                onClick: {
                  openLink: {
                    url: contextUrl
                  }
                }
              }]
            }]
          }]
        }]
      };
      
      sendWebhookNotification(user.webhookUrl, message);
    });
  } catch (error) {
    console.error('Error notifying mentioned users:', error);
    Logger.log('ERROR in notifyMentionedUsers: ' + error.toString());
  }
}

/**
 * Sends a webhook notification to the specified URL
 * @param {string} webhookUrl - The webhook URL to send the notification to
 * @param {Object} message - The message payload to send
 * @returns {Object} Response from the webhook
 */
function sendWebhookNotification(webhookUrl, message) {
  if (!webhookUrl) {
    console.error('No webhook URL provided');
    return null;
  }
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(message),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode < 200 || responseCode >= 300) {
      console.error(`Webhook notification failed with status ${responseCode}: ${response.getContentText()}`);
    } else {
      console.log(`Webhook notification sent successfully to ${webhookUrl}`);
    }
    
    return {
      status: responseCode,
      content: response.getContentText()
    };
  } catch (error) {
    console.error('Error sending webhook notification:', error);
    Logger.log('ERROR in sendWebhookNotification: ' + error.toString());
    return { error: error.toString() };
  }
}
