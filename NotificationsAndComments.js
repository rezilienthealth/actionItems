/**
 * NotificationsAndComments.js - Functions for handling notifications, comments, and mentions
 */

/**
 * Gets comments for an action item
 * @param {string} actionItemId - The ID of the action item
 * @returns {Array} Array of comments
 */
function getCommentsForActionItem(actionItemId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_COMMENTS);
    
    if (!sheet) {
      return [];
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return [];
    }
    
    const headers = values[0];
    const comments = [];
    
    // Find ID column index
    const idColIndex = headers.indexOf('actionItemId');
    
    if (idColIndex === -1) {
      return [];
    }
    
    // Find matching comments
    for (let i = 1; i < values.length; i++) {
      if (values[i][idColIndex] === actionItemId) {
        const comment = {};
        headers.forEach((header, index) => {
          comment[header] = values[i][index];
        });
        comments.push(comment);
      }
    }
    
    // Sort by timestamp
    comments.sort((a, b) => {
      return new Date(a.timestamp) - new Date(b.timestamp);
    });
    
    return comments;
  } catch (error) {
    Logger.log('ERROR in getCommentsForActionItem: ' + error.toString());
    return [];
  }
}

/**
 * Adds a comment to an action item
 * @param {string} actionItemId - The ID of the action item
 * @param {Object} commentData - The comment data
 * @returns {Object} The saved comment
 */
function addCommentToActionItem(actionItemId, commentData) {
  try {
    // Validate
    if (!actionItemId) {
      throw new Error('Action item ID is required');
    }
    
    if (!commentData.content) {
      throw new Error('Comment content is required');
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_COMMENTS);
    
    if (!sheet) {
      throw new Error('ActionItemComments sheet not found');
    }
    
    // Get headers
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Generate comment ID
    const commentId = 'COM-' + new Date().getTime();
    
    // Prepare comment data
    const comment = {
      commentId: commentId,
      actionItemId: actionItemId,
      content: commentData.content,
      author: getUserEmail(),
      timestamp: new Date()
    };
    
    // Prepare row data
    const rowData = headers.map(header => comment[header] !== undefined ? comment[header] : '');
    
    // Append to sheet
    sheet.appendRow(rowData);
    
    // Process mentions in the comment
    const mentions = extractMentions(commentData.content);
    if (mentions.length > 0) {
      // Get the action item
      const actionItem = getActionItemById(actionItemId);
      
      if (actionItem) {
        // Send notifications
        sendMentionNotifications(mentions, actionItem, comment);
        
        // Update the action item's mentionedUsers field
        if (!actionItem.mentionedUsers) {
          actionItem.mentionedUsers = [];
        }
        
        // Add new mentions
        mentions.forEach(mention => {
          if (!actionItem.mentionedUsers.includes(mention)) {
            actionItem.mentionedUsers.push(mention);
          }
        });
        
        // Save the updated action item
        saveActionItem(actionItem);
      }
    }
    
    return comment;
  } catch (error) {
    Logger.log('ERROR in addCommentToActionItem: ' + error.toString());
    throw error;
  }
}

/**
 * Extracts @mentions from text
 * @param {string} text - The text to extract mentions from
 * @returns {Array} Array of mentioned email addresses
 */
function extractMentions(text) {
  try {
    if (!text) {
      return [];
    }
    
    // Match @username or @user.name patterns
    const mentionRegex = /@([a-zA-Z0-9.]+)/g;
    const matches = text.match(mentionRegex) || [];
    
    // Process matches
    const mentions = matches.map(match => {
      // Remove @ symbol
      const username = match.substring(1);
      
      // If it doesn't contain a domain, add the default domain
      if (!username.includes('@')) {
        return username + '@rezilienthealth.com';
      }
      
      return username;
    });
    
    return [...new Set(mentions)]; // Remove duplicates
  } catch (error) {
    Logger.log('ERROR in extractMentions: ' + error.toString());
    return [];
  }
}

/**
 * Processes mentions in an action item
 * @param {Object} actionItemData - The action item data
 * @param {boolean} isNew - Whether this is a new action item
 */
function processActionItemMentions(actionItemData, isNew) {
  try {
    // Skip if no mentions
    if (!actionItemData.mentionedUsers || actionItemData.mentionedUsers.length === 0) {
      return;
    }
    
    // Convert to array if it's a string
    const mentions = Array.isArray(actionItemData.mentionedUsers) 
      ? actionItemData.mentionedUsers 
      : actionItemData.mentionedUsers.split(',').map(s => s.trim()).filter(Boolean);
    
    // Send notifications
    sendMentionNotifications(mentions, actionItemData);
  } catch (error) {
    Logger.log('ERROR in processActionItemMentions: ' + error.toString());
  }
}

/**
 * Sends notifications for mentions
 * @param {Array} mentions - Array of email addresses
 * @param {Object} actionItem - The action item
 * @param {Object} comment - Optional comment data if the mention is in a comment
 */
function sendMentionNotifications(mentions, actionItem, comment = null) {
  try {
    const userEmail = getUserEmail();
    
    mentions.forEach(email => {
      // Don't notify the current user
      if (email === userEmail) {
        return;
      }
      
      // Check if user exists
      if (!userExists(email)) {
        return;
      }
      
      // Prepare notification
      let subject, body;
      
      if (comment) {
        subject = `You were mentioned in a comment on ActionItem: ${actionItem.title}`;
        body = `
          <p>${userEmail} mentioned you in a comment:</p>
          <p><strong>${comment.content}</strong></p>
          <p>On ActionItem: ${actionItem.title}</p>
          <p><a href="${getScriptUrl()}?view=item&id=${actionItem.actionItemId}">View ActionItem</a></p>
        `;
      } else {
        subject = `You were mentioned in ActionItem: ${actionItem.title}`;
        body = `
          <p>${userEmail} mentioned you in an ActionItem:</p>
          <p><strong>${actionItem.title}</strong></p>
          <p>${actionItem.description || ''}</p>
          <p><a href="${getScriptUrl()}?view=item&id=${actionItem.actionItemId}">View ActionItem</a></p>
        `;
      }
      
      // Send email notification
      GmailApp.sendEmail(
        email,
        subject,
        "This is an HTML email. Please use an HTML-compatible email viewer.",
        { htmlBody: body }
      );
      
      // Also send to Google Chat if webhook is available
      sendToChatWebhook(email, subject, body);
    });
  } catch (error) {
    Logger.log('ERROR in sendMentionNotifications: ' + error.toString());
  }
}

/**
 * Sends a status change notification
 * @param {Object} actionItem - The action item
 * @param {string} oldStatus - The old status
 * @param {string} newStatus - The new status
 */
function sendStatusChangeNotification(actionItem, oldStatus, newStatus) {
  try {
    // Skip if status didn't change
    if (oldStatus === newStatus) {
      return;
    }
    
    const userEmail = getUserEmail();
    
    // Notify the assignee
    if (actionItem.assignedTo && actionItem.assignedTo !== userEmail) {
      const subject = `ActionItem Status Changed: ${actionItem.title}`;
      const body = `
        <p>${userEmail} changed the status of an ActionItem assigned to you:</p>
        <p><strong>${actionItem.title}</strong></p>
        <p>Status changed from <strong>${oldStatus}</strong> to <strong>${newStatus}</strong></p>
        <p><a href="${getScriptUrl()}?view=item&id=${actionItem.actionItemId}">View ActionItem</a></p>
      `;
      
      // Send email notification
      GmailApp.sendEmail(
        actionItem.assignedTo,
        subject,
        "This is an HTML email. Please use an HTML-compatible email viewer.",
        { htmlBody: body }
      );
      
      // Also send to Google Chat if webhook is available
      sendToChatWebhook(actionItem.assignedTo, subject, body);
    }
    
    // Notify mentioned users
    if (actionItem.mentionedUsers && actionItem.mentionedUsers.length > 0) {
      const mentions = Array.isArray(actionItem.mentionedUsers) 
        ? actionItem.mentionedUsers 
        : actionItem.mentionedUsers.split(',').map(s => s.trim()).filter(Boolean);
      
      mentions.forEach(email => {
        // Don't notify the current user or assignee (already notified)
        if (email === userEmail || email === actionItem.assignedTo) {
          return;
        }
        
        const subject = `ActionItem Status Changed: ${actionItem.title}`;
        const body = `
          <p>${userEmail} changed the status of an ActionItem you're mentioned in:</p>
          <p><strong>${actionItem.title}</strong></p>
          <p>Status changed from <strong>${oldStatus}</strong> to <strong>${newStatus}</strong></p>
          <p><a href="${getScriptUrl()}?view=item&id=${actionItem.actionItemId}">View ActionItem</a></p>
        `;
        
        // Send email notification
        GmailApp.sendEmail(
          email,
          subject,
          "This is an HTML email. Please use an HTML-compatible email viewer.",
          { htmlBody: body }
        );
        
        // Also send to Google Chat if webhook is available
        sendToChatWebhook(email, subject, body);
      });
    }
  } catch (error) {
    Logger.log('ERROR in sendStatusChangeNotification: ' + error.toString());
  }
}

/**
 * Sends a notification to a Google Chat webhook
 * @param {string} email - The email of the user to notify
 * @param {string} subject - The notification subject
 * @param {string} htmlBody - The HTML body of the notification
 */
function sendToChatWebhook(email, subject, htmlBody) {
  try {
    // Get user's group memberships
    const groups = getUserGroups(email);
    
    if (!groups || groups.length === 0) {
      return;
    }
    
    // Get webhooks for the groups
    const webhooks = getGroupWebhooks(groups);
    
    if (!webhooks || webhooks.length === 0) {
      return;
    }
    
    // Convert HTML to plain text for Chat
    const textBody = htmlBody
      .replace(/<[^>]*>/g, '') // Remove HTML tags
      .replace(/\s+/g, ' ')    // Normalize whitespace
      .trim();
    
    // Prepare the message
    const message = {
      text: `*${subject}*\n\n${textBody}`
    };
    
    // Send to each webhook
    webhooks.forEach(webhook => {
      try {
        UrlFetchApp.fetch(webhook, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(message)
        });
      } catch (webhookError) {
        Logger.log('ERROR sending to webhook: ' + webhookError.toString());
      }
    });
  } catch (error) {
    Logger.log('ERROR in sendToChatWebhook: ' + error.toString());
  }
}

/**
 * Gets the groups a user belongs to
 * @param {string} email - The user's email
 * @returns {Array} Array of group names
 */
function getUserGroups(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.GROUP_MEMBERSHIPS);
    
    if (!sheet) {
      return [];
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return [];
    }
    
    const headers = values[0];
    const groups = [];
    
    // Find email column index
    const emailColIndex = headers.indexOf('userEmail');
    const groupColIndex = headers.indexOf('groupName');
    
    if (emailColIndex === -1 || groupColIndex === -1) {
      return [];
    }
    
    // Find matching groups
    for (let i = 1; i < values.length; i++) {
      if (values[i][emailColIndex] === email) {
        groups.push(values[i][groupColIndex]);
      }
    }
    
    return groups;
  } catch (error) {
    Logger.log('ERROR in getUserGroups: ' + error.toString());
    return [];
  }
}

/**
 * Gets webhooks for groups
 * @param {Array} groups - Array of group names
 * @returns {Array} Array of webhook URLs
 */
function getGroupWebhooks(groups) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.NOTIFICATION_GROUPS);
    
    if (!sheet) {
      return [];
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return [];
    }
    
    const headers = values[0];
    const webhooks = [];
    
    // Find column indices
    const groupColIndex = headers.indexOf('groupName');
    const webhookColIndex = headers.indexOf('chatSpaceWebhook');
    
    if (groupColIndex === -1 || webhookColIndex === -1) {
      return [];
    }
    
    // Find matching webhooks
    for (let i = 1; i < values.length; i++) {
      if (groups.includes(values[i][groupColIndex]) && values[i][webhookColIndex]) {
        webhooks.push(values[i][webhookColIndex]);
      }
    }
    
    return webhooks;
  } catch (error) {
    Logger.log('ERROR in getGroupWebhooks: ' + error.toString());
    return [];
  }
}

/**
 * Checks if a user exists
 * @param {string} email - The user's email
 * @returns {boolean} True if the user exists
 */
function userExists(email) {
  try {
    const ss = SpreadsheetApp.openById(USER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("authorizedUsers");
    
    if (!sheet) {
      return false;
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return false;
    }
    
    const headers = values[0];
    const emailColIndex = headers.indexOf("email");
    
    if (emailColIndex === -1) {
      return false;
    }
    
    // Check if email exists
    for (let i = 1; i < values.length; i++) {
      if (values[i][emailColIndex] === email) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    Logger.log('ERROR in userExists: ' + error.toString());
    return false;
  }
}

/**
 * Gets the script URL
 * @returns {string} The URL of the web app
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
