/**
 * ActionItemCRUD.js - Core CRUD functions for ActionItems
 */

/**
 * Gets all action items, optionally filtered by patient ID
 * @param {string} patientId - Optional patient ID to filter by
 * @param {boolean} useCache - Whether to use cached data
 * @returns {Array} Array of action items
 */
function getActionItems(patientId = null, useCache = true) {
  // Check cache first if requested
  const currentTime = new Date().getTime();
  if (useCache && cachedItems && (currentTime - lastItemsCacheTime < CACHE_EXPIRATION_SECONDS * 1000)) {
    // Filter cached items by patient if needed
    if (patientId) {
      return cachedItems.filter(item => item.athenaId === patientId);
    }
    return cachedItems;
  }
  
  try {
    // Open the spreadsheet and get the actionItems sheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEMS);
    
    if (!sheet) {
      Logger.log('ERROR: ActionItems sheet not found');
      return [];
    }
    
    // Get all data from the sheet
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      // Only headers or empty sheet
      return [];
    }
    
    const headers = values[0];
    const items = [];
    
    // Map column indices for faster access
    const columnMap = {};
    headers.forEach((header, index) => {
      columnMap[header] = index;
    });
    
    // Process each row
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Create an item object with all properties
      const item = {};
      headers.forEach((header, index) => {
        const value = row[index];

        // Handle special fields first
        if (header === 'tags' || header === 'mentionedUsers' || header === 'selectedOptions' || header === 'relatedIds') {
          item[header] = value ? String(value).split(',').map(s => s.trim()).filter(Boolean) : [];
        } 
        // Generic type handling for everything else
        else if (value instanceof Date) {
          item[header] = value.toISOString();
        } else if (typeof value === 'boolean') {
          item[header] = value;
        } else if (String(value).toUpperCase() === 'TRUE') {
          item[header] = true;
        } else if (String(value).toUpperCase() === 'FALSE') {
          item[header] = false;
        } else {
          item[header] = value;
        }
      });
      
      // Skip if this item doesn't match the patient filter
      if (patientId && item.athenaId !== patientId) {
        continue;
      }
      
      items.push(item);
    }
    
    // Update cache
    if (!patientId) {
      cachedItems = items;
      lastItemsCacheTime = currentTime;
    }
    
    // Final check to ensure we always return an array
    return items || [];
  } catch (error) {
    Logger.log('ERROR in getActionItems: ' + error.toString());
    return [];
  }
}

/**
 * Gets a specific action item by ID
 * @param {string} actionItemId - The ID of the action item to retrieve
 * @returns {Object|null} The action item or null if not found
 */
function getActionItemById(actionItemId) {
  if (!actionItemId) {
    return null;
  }
  
  try {
    // Try to get from cache first
    const items = getActionItems(null, true);
    const item = items.find(item => item.actionItemId === actionItemId);
    
    if (item) {
      return item;
    }
    
    // If not in cache, get directly from sheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEMS);
    
    if (!sheet) {
      return null;
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    
    // Find ID column index
    const idColIndex = headers.indexOf('actionItemId');
    if (idColIndex === -1) {
      return null;
    }
    
    // Find the row with matching ID
    for (let i = 1; i < values.length; i++) {
      if (values[i][idColIndex] === actionItemId) {
        const item = {};
        headers.forEach((header, index) => {
          // Handle special fields
          if (header === 'tags' || header === 'mentionedUsers' || header === 'selectedOptions' || header === 'relatedIds') {
            // Convert comma-separated strings to arrays
            item[header] = values[i][index] ? String(values[i][index]).split(',').map(s => s.trim()).filter(Boolean) : [];
          } else if (header === 'isRecurring' || header === 'isTemplate' || 
                    header === 'faxSent' || header === 'visitInfoAttached' || 
                    header === 'facesheetAttached') {
            // Convert to boolean
            item[header] = values[i][index] === true || String(values[i][index]).toUpperCase() === 'TRUE';
          } else {
            // Standard field
            item[header] = values[i][index];
          }
        });
        return item;
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('ERROR in getActionItemById: ' + error.toString());
    return null;
  }
}

/**
 * Saves an action item (create or update)
 * @param {Object} actionItemData - The action item data to save
 * @returns {Object} The saved action item with ID
 */
function saveActionItem(actionItemData) {
  try {
    // Validate required fields
    if (!actionItemData.title) {
      throw new Error('Title is required');
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEMS);
    
    if (!sheet) {
      throw new Error('ActionItems sheet not found');
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    
    // Map column indices
    const columnMap = {};
    headers.forEach((header, index) => {
      columnMap[header] = index;
    });
    
    // Check if this is a new item or an update
    const isNew = !actionItemData.actionItemId;
    let rowIndex = -1;
    
    if (!isNew) {
      // Find the row with matching ID
      for (let i = 1; i < values.length; i++) {
        if (values[i][columnMap.actionItemId] === actionItemData.actionItemId) {
          rowIndex = i;
          break;
        }
      }
      
      if (rowIndex === -1) {
        throw new Error('Action item with ID ' + actionItemData.actionItemId + ' not found');
      }
    }
    
    // Generate a new ID if needed
    if (isNew) {
      actionItemData.actionItemId = 'AI-' + new Date().getTime();
    }
    
    // Set timestamps
    const now = new Date();
    const userEmail = getUserEmail();
    
    if (isNew) {
      actionItemData.createdBy = userEmail;
      actionItemData.createdAt = now;
    } else {
      // For existing items, preserve the original createdBy and createdAt from the spreadsheet
      const originalRow = values[rowIndex];
      if (!actionItemData.createdBy && originalRow[columnMap.createdBy]) {
        actionItemData.createdBy = originalRow[columnMap.createdBy];
      }
      if (!actionItemData.createdAt && originalRow[columnMap.createdAt]) {
        actionItemData.createdAt = originalRow[columnMap.createdAt];
      }
    }
    
    actionItemData.lastUpdated = now;
    actionItemData.lastUpdatedBy = userEmail;
    
    // Handle approval logic
    if (isNew && isProvider() && !actionItemData.approvedBy) {
      // Auto-approve if created by a provider
      actionItemData.approvedBy = userEmail;
      actionItemData.approvedAt = now;
    }
    
    // Convert arrays to comma-separated strings for storage
    ['tags', 'mentionedUsers', 'selectedOptions', 'relatedIds'].forEach(field => {
      if (Array.isArray(actionItemData[field])) {
        actionItemData[field] = actionItemData[field].join(',');
      }
    });
    
    // Convert booleans to TRUE/FALSE strings
    ['isRecurring', 'isTemplate', 'faxSent', 'visitInfoAttached', 'facesheetAttached'].forEach(field => {
      if (actionItemData[field] !== undefined) {
        actionItemData[field] = actionItemData[field] ? 'TRUE' : 'FALSE';
      }
    });
    
    // Prepare row data
    const rowData = headers.map(header => actionItemData[header] !== undefined ? actionItemData[header] : '');
    
    if (isNew) {
      // Append new row
      sheet.appendRow(rowData);
      
      // Log audit event
      logActionItemAudit(
        actionItemData.actionItemId,
        'created',
        '',
        JSON.stringify(actionItemData)
      );
    } else {
      // Update existing row
      const oldData = {};
      headers.forEach((header, index) => {
        oldData[header] = values[rowIndex][index];
      });
      
      // Update the row
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([rowData]);
      
      // Log changes to audit trail
      headers.forEach(header => {
        if (actionItemData[header] !== oldData[header] && header !== 'lastUpdated' && header !== 'lastUpdatedBy') {
          logActionItemAudit(
            actionItemData.actionItemId,
            header,
            oldData[header],
            actionItemData[header]
          );
        }
      });
    }
    
    // Process mentions
    processActionItemMentions(actionItemData, isNew);
    
    // Clear cache
    cachedItems = null;
    
    // Convert back to arrays for return value
    ['tags', 'mentionedUsers', 'selectedOptions', 'relatedIds'].forEach(field => {
      if (typeof actionItemData[field] === 'string') {
        actionItemData[field] = actionItemData[field] ? actionItemData[field].split(',').map(s => s.trim()).filter(Boolean) : [];
      }
    });
    
    // Convert back to booleans for return value
    ['isRecurring', 'isTemplate', 'faxSent', 'visitInfoAttached', 'facesheetAttached'].forEach(field => {
      if (actionItemData[field] !== undefined) {
        actionItemData[field] = actionItemData[field] === true || String(actionItemData[field]).toUpperCase() === 'TRUE';
      }
    });
    
    return actionItemData;
  } catch (error) {
    Logger.log('ERROR in saveActionItem: ' + error.toString());
    throw error;
  }
}

/**
 * Processes mentions in an action item and sends notifications
 * @param {Object} actionItem - The action item data
 * @param {boolean} isNew - Whether this is a new action item
 */
function processActionItemMentions(actionItem, isNew) {
  try {
    if (!actionItem) return;
    
    // Get the current user's email for the author
    const author = Session.getActiveUser().getEmail();
    
    // Check for mentions in the description
    if (actionItem.description) {
      notifyMentionedUsers(
        actionItem.description, 
        getActionItemUrl(actionItem.actionItemId),
        author,
        actionItem.title || 'Untitled Action Item'
      );
    }
    
    // Check for mentions in the mentionedUsers field
    if (actionItem.mentionedUsers) {
      // Convert string to array if needed
      const mentionedUsers = Array.isArray(actionItem.mentionedUsers) 
        ? actionItem.mentionedUsers 
        : actionItem.mentionedUsers.split(',').map(u => u.trim());
      
      // Only process if there are actually mentioned users
      if (mentionedUsers.length > 0) {
        notifyMentionedUsers(
          mentionedUsers.join(', '), // Convert array back to string for processing
          getActionItemUrl(actionItem.actionItemId),
          author,
          actionItem.title || 'Untitled Action Item',
          true // Explicit mention
        );
      }
    }
    
    // For new action items, also check for mentions in comments if any
    if (isNew && actionItem.comments) {
      // If comments is a string, try to parse it as JSON
      let comments = [];
      try {
        comments = typeof actionItem.comments === 'string' 
          ? JSON.parse(actionItem.comments) 
          : actionItem.comments;
      } catch (e) {
        Logger.log('Error parsing comments: ' + e.toString());
      }
      
      // Process each comment for mentions
      comments.forEach(comment => {
        if (comment.content) {
          notifyMentionedUsers(
            comment.content,
            getActionItemUrl(actionItem.actionItemId) + "#comment-" + (comment.id || ''),
            comment.author || author,
            'Comment on: ' + (actionItem.title || 'Untitled Action Item')
          );
        }
      });
    }
  } catch (error) {
    // Log the error but don't fail the save operation
    console.error('Error processing mentions:', error);
    Logger.log('ERROR in processActionItemMentions: ' + error.toString());
  }
}

/**
 * Gets the URL for an action item
 * @param {string} actionItemId - The action item ID
 * @returns {string} The URL to view the action item
 */
function getActionItemUrl(actionItemId) {
  const scriptUrl = ScriptApp.getService().getUrl();
  return scriptUrl + '?actionItemId=' + encodeURIComponent(actionItemId);
}

/**
 * Updates the status of an action item
 * @param {string} actionItemId - The ID of the action item
 * @param {string} newStatus - The new status
 * @returns {Object} The updated action item
 */
function updateActionItemStatus(actionItemId, newStatus) {
  try {
    // Get the current item
    const item = getActionItemById(actionItemId);
    
    if (!item) {
      throw new Error('Action item not found');
    }
    
    // Update the status
    const oldStatus = item.status;
    item.status = newStatus;
    
    // Handle completion
    if (newStatus === 'Completed' && !item.completedBy) {
      item.completedBy = getUserEmail();
      item.completedAt = new Date();
    }
    
    // Save the updated item
    const updatedItem = saveActionItem(item);
    
    // Send notifications about status change
    sendStatusChangeNotification(updatedItem, oldStatus, newStatus);
    
    return updatedItem;
  } catch (error) {
    Logger.log('ERROR in updateActionItemStatus: ' + error.toString());
    throw error;
  }
}

/**
 * Deletes an action item (marks it as deleted)
 * @param {string} actionItemId - The ID of the action item to delete
 * @returns {boolean} True if successful
 */
function deleteActionItem(actionItemId) {
  try {
    // Get the current item
    const item = getActionItemById(actionItemId);
    
    if (!item) {
      throw new Error('Action item not found');
    }
    
    // Mark as deleted
    item.status = 'Deleted';
    item.lastUpdated = new Date();
    item.lastUpdatedBy = getUserEmail();
    
    // Save the updated item
    saveActionItem(item);
    
    // Log deletion to audit trail
    logActionItemAudit(
      actionItemId,
      'deleted',
      '',
      getUserEmail()
    );
    
    // Clear cache
    cachedItems = null;
    
    return true;
  } catch (error) {
    Logger.log('ERROR in deleteActionItem: ' + error.toString());
    return false;
  }
}

/**
 * Logs an action item audit event
 * @param {string} actionItemId - The ID of the action item
 * @param {string} fieldChanged - The field that changed
 * @param {string} oldValue - The old value
 * @param {string} newValue - The new value
 */
function logActionItemAudit(actionItemId, fieldChanged, oldValue, newValue) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_AUDIT);
    
    if (!sheet) {
      Logger.log('ERROR: ActionItemsAudit sheet not found');
      return;
    }
    
    const timestamp = new Date();
    const userEmail = getUserEmail();
    
    sheet.appendRow([
      actionItemId,
      fieldChanged,
      oldValue,
      newValue,
      userEmail,
      timestamp
    ]);
  } catch (error) {
    Logger.log('ERROR in logActionItemAudit: ' + error.toString());
  }
}

/**
 * Completes an action item (wrapper for updateActionItemStatus)
 * @param {string} actionItemId - The ID of the action item to complete
 * @returns {Object} The updated action item
 */
function completeActionItem(actionItemId) {
  return updateActionItemStatus(actionItemId, 'Completed');
}

/**
 * Assigns an action item to a user
 * @param {string} actionItemId - The ID of the action item
 * @param {string} assigneeEmail - The email of the user to assign to
 * @returns {Object} The updated action item
 */
function assignActionItem(actionItemId, assigneeEmail) {
  try {
    // Get the current item
    const item = getActionItemById(actionItemId);
    
    if (!item) {
      throw new Error('Action item not found');
    }
    
    // Update the assignment
    const oldAssignee = item.assignedTo;
    item.assignedTo = assigneeEmail;
    item.lastUpdated = new Date();
    item.lastUpdatedBy = getUserEmail();
    
    // Save the updated item
    const updatedItem = saveActionItem(item);
    
    // Log assignment change to audit trail
    logActionItemAudit(
      actionItemId,
      'assignedTo',
      oldAssignee || '',
      assigneeEmail
    );
    
    return updatedItem;
  } catch (error) {
    Logger.log('ERROR in assignActionItem: ' + error.toString());
    throw error;
  }
}

/**
 * Updates a specific field of an action item
 * @param {string} actionItemId - The ID of the action item to update
 * @param {string} fieldName - The name of the field to update
 * @param {string} fieldValue - The new value for the field
 * @param {string} [comment] - Optional comment about the update
 * @returns {Object} The updated action item
 */
function updateActionItemField(actionItemId, fieldName, fieldValue, comment) {
  try {
    // Get the current item
    const item = getActionItemById(actionItemId);
    
    if (!item) {
      throw new Error('Action item not found');
    }
    
    // Store old value for audit log
    const oldValue = item[fieldName] || '';
    
    // Update the field
    item[fieldName] = fieldValue;
    
    // Update timestamps if this is an assignment change
    if (fieldName === 'assignedTo') {
      item.lastUpdated = new Date();
      item.lastUpdatedBy = Session.getActiveUser().getEmail();
    }
    
    // Save the updated item
    const updatedItem = saveActionItem(item);
    
    // Log the field update to audit trail
    logActionItemAudit(
      actionItemId,
      fieldName,
      oldValue,
      fieldValue,
      comment
    );
    
    return updatedItem;
  } catch (error) {
    Logger.log('ERROR in updateActionItemField: ' + error.toString());
    throw error;
  }
}

/**
 * Gets the history of an action item from the audit trail
 * @param {string} actionItemId - The ID of the action item
 * @returns {Array} Array of audit events
 */
function getActionItemHistory(actionItemId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_AUDIT);
    
    if (!sheet) {
      return [];
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return [];
    }
    
    const headers = values[0];
    const history = [];
    
    // Find ID column index
    const idColIndex = headers.indexOf('actionItemId');
    
    if (idColIndex === -1) {
      return [];
    }
    
    // Find matching audit events
    for (let i = 1; i < values.length; i++) {
      if (values[i][idColIndex] === actionItemId) {
        const event = {};
        headers.forEach((header, index) => {
          event[header] = values[i][index];
        });
        history.push(event);
      }
    }
    
    return history;
  } catch (error) {
    Logger.log('ERROR in getActionItemHistory: ' + error.toString());
    return [];
  }
}
