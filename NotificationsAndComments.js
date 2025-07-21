/**
 * NotificationsAndComments.js - Functions for handling notifications, comments, and mentions
 */

/**
 * Gets all notification groups for admin management
 * @returns {Array} Array of notification groups
 */
function getNotificationGroups() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
    
    if (!sheet) {
      console.log('Groups sheet not found, returning empty array');
      return [];
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return [];
    }
    
    const headers = values[0];
    const groups = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const group = {};
      
      headers.forEach((header, index) => {
        group[header] = row[index];
      });
      
      groups.push(group);
    }
    
    return groups;
  } catch (error) {
    console.error('Error getting notification groups:', error);
    return [];
  }
}

/**
 * Gets all group memberships for admin management
 * @returns {Array} Array of group memberships
 */
function getGroupMemberships() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.GROUP_MEMBERSHIPS);
    
    if (!sheet) {
      console.log('Group memberships sheet not found, returning empty array');
      return [];
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return [];
    }
    
    const headers = values[0];
    const memberships = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const membership = {};
      
      headers.forEach((header, index) => {
        membership[header] = row[index];
      });
      
      memberships.push(membership);
    }
    
    return memberships;
  } catch (error) {
    console.error('Error getting group memberships:', error);
    return [];
  }
}

/**
 * Simple test function to verify backend execution
 * @returns {Object} Simple test result
 */
function simpleTest() {
  try {
    return {
      success: true,
      message: 'Backend function executed successfully',
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    Logger.log('ERROR in simpleTest: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Test function to debug ActionItemComments sheet contents
 * @returns {Object} Debug information about the sheet
 */
function debugCommentsSheet() {
  console.log('=== debugCommentsSheet called ===');
  
  try {
    // Get spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('Spreadsheet opened successfully');
    
    // Try to get sheet by exact name
    let sheet = ss.getSheetByName('actionItemComments');
    console.log('Sheet lookup result:', sheet ? 'FOUND' : 'NOT FOUND');
    
    if (!sheet) {
      // List all sheets to see what's available
      const allSheets = ss.getSheets().map(s => s.getName());
      console.log('Available sheets:', allSheets);
      
      return {
        error: 'ActionItemComments sheet not found',
        availableSheets: allSheets,
        lookingFor: 'actionItemComments'
      };
    }
    
    // Get sheet info
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    console.log('Sheet dimensions - Last row:', lastRow, 'Last col:', lastCol);
    
    if (lastRow === 0) {
      return {
        sheetExists: true,
        error: 'Sheet is completely empty (no rows)',
        lastRow: lastRow,
        lastCol: lastCol
      };
    }
    
    // Get all data
    const allData = sheet.getDataRange().getValues();
    console.log('Retrieved data rows:', allData.length);
    
    const result = {
      success: true,
      sheetExists: true,
      sheetName: sheet.getName(),
      lastRow: lastRow,
      lastCol: lastCol,
      headers: allData.length > 0 ? allData[0] : [],
      totalRows: allData.length,
      dataRows: allData.length - 1, // Excluding header
      sampleData: allData.slice(0, Math.min(5, allData.length)) // First 5 rows
    };
    
    console.log('Debug result:', result);
    return result;
    
  } catch (error) {
    console.error('ERROR in debugCommentsSheet:', error.toString());
    console.error('Error stack:', error.stack);
    
    return {
      error: error.toString(),
      stack: error.stack,
      spreadsheetId: SPREADSHEET_ID
    };
  }
}

/**
 * Gets comments for an action item
 * @param {string} actionItemId - The ID of the action item
 * @returns {Array} Array of comments
 */
function getCommentsForActionItem(actionItemId) {
  try {
    // Step 1: Basic validation
    if (!actionItemId) {
      return [];
    }
    
    // Step 2: Try to access spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Step 3: Try to get the actionItemComments sheet
    const sheet = ss.getSheetByName('actionItemComments');
    
    // Step 4: If sheet doesn't exist, return empty array
    if (!sheet) {
      return [];
    }
    
    // Step 5: Check if sheet has data
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return []; // No data rows (only headers or empty)
    }
    
    // Step 6: Get all data from sheet
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    
    // Step 7: Find the actionItemId column
    const actionItemIdCol = headers.indexOf('actionItemId');
    if (actionItemIdCol === -1) {
      return []; // Column not found
    }
    
    // Step 8: Find matching comments
    const comments = [];
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      if (row[actionItemIdCol] === actionItemId) {
        // Build comment object from row data
        const comment = {};
        for (let j = 0; j < headers.length; j++) {
          comment[headers[j]] = row[j];
        }
        comments.push(comment);
      }
    }
    
    // Step 9: Sort by timestamp (newest first)
    comments.sort((a, b) => {
      const dateA = new Date(a.timestamp);
      const dateB = new Date(b.timestamp);
      return dateB - dateA; // Newest first
    });
    
    return comments;
    
  } catch (error) {
    Logger.log('ERROR in getCommentsForActionItem: ' + error.toString());
    // Return empty array on any error
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
    console.log('Adding comment for action item:', actionItemId, 'with data:', commentData);
    
    // Validate
    if (!actionItemId) {
      throw new Error('Action item ID is required');
    }
    
    if (!commentData.content) {
      throw new Error('Comment content is required');
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEM_COMMENTS);
    
    // Create the sheet if it doesn't exist
    if (!sheet) {
      console.log('ActionItemComments sheet not found, creating it...');
      sheet = createActionItemCommentsSheet();
    }
    
    // Get headers - handle case where sheet is newly created
    let headers;
    if (sheet.getLastColumn() === 0) {
      // Sheet is empty, use default headers
      headers = ['commentId', 'actionItemId', 'author', 'content', 'timestamp', 'mentionedUsers'];
    } else {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    console.log('Sheet headers:', headers);
    
    // Generate comment ID
    const commentId = 'COM-' + new Date().getTime();
    
    // Prepare comment data
    const comment = {
      commentId: commentId,
      actionItemId: actionItemId,
      content: commentData.content,
      author: commentData.author || getUserEmail(),
      timestamp: new Date().toISOString(),
      mentionedUsers: extractMentions(commentData.content).join(', ')
    };
    
    console.log('Prepared comment:', comment);
    
    // Prepare row data
    const rowData = headers.map(header => comment[header] !== undefined ? comment[header] : '');
    
    console.log('Row data to append:', rowData);
    
    // Append to sheet
    sheet.appendRow(rowData);
    
    console.log('Comment successfully added to sheet');
    
    // Process mentions in the comment
    if (commentData.content) {
      const mentions = extractMentions(commentData.content);
      console.log(`Found ${mentions.length} mentions in comment:`, mentions);
      
      if (mentions.length > 0) {
        try {
          console.log('Sending mention notifications for:', mentions);
          const actionItem = getActionItemById(actionItemId);
          const notificationResult = sendMentionNotifications(mentions, actionItem, comment);
          console.log('Mention notifications sent with result:', notificationResult);
        } catch (mentionError) {
          const errorMsg = 'Failed to send mention notifications: ' + mentionError.toString();
          console.error(errorMsg);
          Logger.log(errorMsg);
          // Don't fail the entire comment save if notifications fail
        }
      }
    }
    
    console.log('Returning comment object:', comment);
    return comment;
  } catch (error) {
    console.error('ERROR in addCommentToActionItem:', error.toString());
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
    
    console.log('Extracting mentions from text:', text);
    
    // Match @email@domain.com or @username patterns
    const mentionRegex = /@([a-zA-Z0-9._-]+(?:@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})?)/g;
    const matches = text.match(mentionRegex) || [];
    
    console.log('Found mention matches:', matches);
    
    // Process matches
    const mentions = matches.map(match => {
      // Remove @ symbol
      let username = match.substring(1);
      
      // If it already contains @, it's a full email
      if (username.includes('@')) {
        return username;
      }
      
      // If it doesn't contain a domain, add the default domain
      return username + '@rezilienthealth.com';
    });
    
    console.log('Processed mentions:', mentions);
    
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
 * Sends notifications for mentions via email and webhooks
 * @param {Array} mentions - Array of email addresses (case-insensitive)
 * @param {Object} actionItem - The action item data
 * @param {Object} comment - Optional comment data if the mention is in a comment
 */
async function sendMentionNotifications(mentions, actionItem, comment = null) {
  const NOTIFICATION_TIMEOUT_MS = 15000; // 15 seconds timeout for notifications
  const sessionId = Utilities.getUuid().substring(0, 8);
  
  console.log(`=== SEND MENTION NOTIFICATIONS [${sessionId}] START ===`);
  console.log(`Action Item: ${actionItem?.actionItemId} - ${actionItem?.title?.substring(0, 50)}...`);
  console.log(`Mentions: ${mentions?.length || 0} users`);
  console.log('Action Item Data:', JSON.stringify(actionItem, null, 2));
  console.log('Comment Data:', comment ? JSON.stringify(comment, null, 2) : 'No comment');
  
  if (!mentions || !mentions.length) {
    console.log('No mentions to process');
    return [];
  }
  
  try {
    // Normalize and deduplicate emails
    const normalizedMentions = [...new Set(
      mentions
        .filter(email => email && typeof email === 'string')
        .map(email => email.toLowerCase().trim())
    )];
    
    console.log(`Normalized ${mentions.length} mentions to ${normalizedMentions.length} unique emails`);
    
    const userEmail = getUserEmail()?.toLowerCase()?.trim() || 'unknown';
    console.log(`Current user: ${userEmail}`);
    
    // Load user data with timeout
    let userMap = {};
    try {
      console.log('Loading user map...');
      userMap = getUserMap();
      console.log(`User map loaded with ${Object.keys(userMap).length} users`);
    } catch (userError) {
      console.error('Failed to load user map:', userError);
      // Continue with empty map to at least try email notifications
    }
    
    // Process each mention with timeout protection
    const notificationPromises = normalizedMentions.map(email => {
      return new Promise(resolve => {
        const mentionId = `${email.substring(0, 3)}...${email.substring(email.indexOf('@'))}`;
        const startTime = new Date().getTime();
        
        // Store the resolve function to use in our timeout check
        let resolved = false;
        const timeoutCheck = () => {
          if (!resolved && (new Date().getTime() - startTime) > NOTIFICATION_TIMEOUT_MS) {
            console.error(`⏱️ Timeout processing mention for ${mentionId}`);
            resolve({ email, status: 'timeout', error: 'Processing timeout' });
            resolved = true;
            return true;
          }
          return false;
        };
        
        try {
          console.log(`\n--- Processing mention for: ${mentionId} ---`);
          
          // Allow self-mentions for testing, but log them
          if (email === userEmail) {
            console.log(`Processing self-mention for: ${mentionId}`);
            // Continue processing to send webhook
          }
          
          const user = userMap[email];
          if (!user) {
            console.warn(`User not found in userMap: ${mentionId}`);
            resolve({ email, status: 'skipped', error: 'User not found in userMap' });
            resolved = true;
            return;
          }
          
          console.log(`User data for ${email}:`, JSON.stringify(user, null, 2));
          
          if (!user.webhookUrl) {
            console.warn(`No webhook URL configured for user: ${email}`);
            resolve({ email, status: 'skipped', error: 'No webhook URL configured' });
            resolved = true;
            return;
          }
          
          // Prepare notification content with action item details
          const actionItemTitle = actionItem.title || 'Untitled Action Item';
          const commentContent = comment?.content || '';
          // Use getActionItemUrl to get the direct URL to the action item
          const viewUrl = getActionItemUrl(actionItem.actionItemId);
          
          // Create message for webhook
          const message = {
            cards: [{
              header: {
                title: `🔔 Mention in Action Item`,
                subtitle: `From: ${userEmail}`
              },
              sections: [{
                widgets: [{
                  textParagraph: {
                    text: `You were mentioned in: <b>${escapeHtml(actionItemTitle)}</b>`
                  }
                }, {
                  keyValue: {
                    topLabel: 'Action Item',
                    content: actionItemTitle,
                    contentMultiline: true,
                    button: {
                      textButton: {
                        text: 'VIEW IN ACTION ITEMS',
                        onClick: { 
                          openLink: { 
                            url: viewUrl 
                          } 
                        }
                      }
                    }
                  }
                }, ...(commentContent ? [{
                  textParagraph: {
                    text: `<i>${escapeHtml(commentContent.substring(0, 200))}${commentContent.length > 200 ? '...' : ''}</i>`
                  }
                }] : [])]
              }]
            }]
          };
          
          // Process notifications with timeout checks
          const processNotifications = () => {
            if (timeoutCheck()) return;
            
            console.log(`Sending webhook notification to ${email} with URL: ${user.webhookUrl ? '***URL_REDACTED***' : 'undefined'}`);
            
            // Only send webhook notification, skip email
            sendWebhookNotification(user, message, email).then(webhookResult => {
              if (!resolved) {
                console.log(`Webhook notification result for ${email}:`, webhookResult);
                const result = {
                  email,
                  webhookStatus: webhookResult.status,
                  webhookUrl: webhookResult.url ? `${webhookResult.url.substring(0, 30)}...` : undefined,
                  response: webhookResult.response,
                  timestamp: new Date().toISOString()
                };
                console.log(`Webhook notification completed for ${email}:`, result);
                resolve(result);
                resolved = true;
              }
            }).catch(error => {
              if (!resolved) {
                const errorMsg = `Error processing notifications for ${email}: ${error.toString()}`;
                console.error(errorMsg);
                const result = { 
                  email, 
                  status: 'error', 
                  error: error.toString(),
                  timestamp: new Date().toISOString()
                };
                console.log('Error result:', result);
                resolve(result);
                resolved = true;
              }
            });
          };
          
          // Start processing
          processNotifications();
          
        } catch (error) {
          if (!resolved) {
            console.error(`Error processing mention for ${email}:`, error);
            resolve({ email, status: 'error', error: error.toString() });
            resolved = true;
          }
        }
      });
    });
    
    // Wait for all notifications to complete or timeout
    console.log(`Waiting for ${notificationPromises.length} notification(s) to complete...`);
    const results = await Promise.all(notificationPromises);
    console.log(`All notifications completed. Results:`, results);
    
    // Log summary
    const summary = {
      total: results.length,
      success: results.filter(r => r.webhookStatus === 'success').length,
      skipped: results.filter(r => r.status === 'skipped').length,
      errors: results.filter(r => r.status === 'error').length,
      timeouts: results.filter(r => r.status === 'timeout').length
    };
    
    console.log(`\n=== NOTIFICATION SUMMARY [${sessionId}] ===`);
    console.log(`Total: ${summary.total}, Success: ${summary.success}, ` +
                `Skipped: ${summary.skipped}, Errors: ${summary.errors}, Timeouts: ${summary.timeouts}`);
    
    if (summary.errors > 0 || summary.timeouts > 0) {
      const failed = results.filter(r => r.status === 'error' || r.status === 'timeout');
      console.warn('Failed notifications:', failed);
    }
    
  } catch (error) {
    console.error(`❌ CRITICAL ERROR in sendMentionNotifications [${sessionId}]:`, error);
    Logger.log(`ERROR in sendMentionNotifications [${sessionId}]: ${error.toString()}\n${error.stack}`);
  } finally {
    console.log(`=== SEND MENTION NOTIFICATIONS [${sessionId}] COMPLETE ===`);
  }
}

/**
 * Sends an email notification for a mention
 * @private
 */
function sendEmailNotification(email, fromEmail, actionItem, comment, viewUrl) {
  return new Promise(resolve => {
    try {
      const subject = comment ? 
        `You were mentioned in a comment on: ${actionItem.title}` :
        `You were mentioned in: ${actionItem.title}`;
      
      const htmlBody = `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
          <h2>You were mentioned in an action item</h2>
          <p><strong>From:</strong> ${escapeHtml(fromEmail)}</p>
          <div style="background: #f5f5f5; padding: 15px; border-left: 4px solid #4285f4; margin: 15px 0;">
            <h3 style="margin-top: 0;">${escapeHtml(actionItem.title)}</h3>
            ${comment ? `<blockquote style="border-left: 3px solid #ddd; margin: 10px 0; padding-left: 15px; color: #555;">
              ${escapeHtml(comment.content)}
            </blockquote>` : ''}
          </div>
          <p>
            <a href="${viewUrl}" style="display: inline-block; background: #4285f4; color: white; 
              padding: 10px 20px; text-decoration: none; border-radius: 4px; margin-top: 10px;">
              View in Action Items
            </a>
          </p>
        </div>
      `;
      
      GmailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody,
        noReply: true,
        name: 'Action Items Notifications'
      });
      
      console.log(`✅ Email sent to: ${email}`);
      resolve({ status: 'success' });
      
    } catch (error) {
      console.error(`❌ Failed to send email to ${email}:`, error);
      resolve({ status: 'error', error: error.toString() });
    }
  });
}

/**
 * Sends a webhook notification
 * @private
 */
function sendWebhookNotification(user, message, email) {
  const requestId = Utilities.getUuid().substring(0, 8);
  const startTime = new Date();
  
  console.log(`[${requestId}] Starting webhook notification for ${email}`);
  
  return new Promise(resolve => {
    if (!user?.webhookUrl) {
      const errorMsg = `ℹ️ [${requestId}] No webhook configured for ${email}`;
      console.log(errorMsg);
      return resolve({ 
        status: 'skipped', 
        reason: 'no-webhook',
        requestId: requestId,
        timestamp: new Date().toISOString()
      });
    }
    
    const webhookUrl = user.webhookUrl.trim();
    if (!webhookUrl.startsWith('http')) {
      const errorMsg = `⚠️ [${requestId}] Invalid webhook URL for ${email}: ${webhookUrl}`;
      console.warn(errorMsg);
      return resolve({ 
        status: 'error', 
        error: 'invalid-url', 
        url: webhookUrl,
        requestId: requestId,
        timestamp: new Date().toISOString()
      });
    }
    
    const shortUrl = webhookUrl.length > 30 ? 
      `${webhookUrl.substring(0, 15)}...${webhookUrl.substring(webhookUrl.length - 15)}` : 
      webhookUrl;
    
    console.log(`[${requestId}] Sending webhook to ${email} (${shortUrl})`);
    console.log(`[${requestId}] Webhook payload:`, JSON.stringify(message, null, 2));
    
    try {
      const fetchOptions = {
        method: 'post',
        contentType: 'application/json; charset=utf-8',
        payload: JSON.stringify(message),
        muteHttpExceptions: true,
        validateHttpsCertificates: true,
        followRedirects: true,
        escaping: true,
        headers: {
          'X-Request-ID': requestId,
          'X-App-Name': 'ActionItems',
          'X-User-Email': email
        }
      };
      
      console.log(`[${requestId}] Sending webhook request to: ${webhookUrl}`);
      console.log(`[${requestId}] Request options:`, JSON.stringify({
        method: fetchOptions.method,
        contentType: fetchOptions.contentType,
        muteHttpExceptions: fetchOptions.muteHttpExceptions,
        validateHttpsCertificates: fetchOptions.validateHttpsCertificates,
        followRedirects: fetchOptions.followRedirects,
        escaping: fetchOptions.escaping,
        headers: Object.keys(fetchOptions.headers)
      }, null, 2));
      
      const startTimeMs = new Date().getTime();
      const response = UrlFetchApp.fetch(webhookUrl, fetchOptions);
      const endTimeMs = new Date().getTime();
      
      const status = response.getResponseCode();
      const responseHeaders = response.getAllHeaders();
      const responseText = response.getContentText().trim();
      const responseTimeMs = endTimeMs - startTimeMs;
      
      console.log(`[${requestId}] Webhook response received in ${responseTimeMs}ms`);
      console.log(`[${requestId}] Response status: ${status}`);
      console.log(`[${requestId}] Response headers:`, JSON.stringify(responseHeaders, null, 2));
      
      const result = {
        status: status >= 200 && status < 300 ? 'success' : 'error',
        url: webhookUrl,
        responseCode: status,
        responseTimeMs: responseTimeMs,
        responseHeaders: responseHeaders,
        responseBody: responseText,
        requestId: requestId,
        timestamp: new Date().toISOString()
      };
      
      if (result.status === 'success') {
        console.log(`✅ [${requestId}] Webhook to ${email} succeeded (${status}) in ${responseTimeMs}ms`);
      } else {
        console.error(`❌ [${requestId}] Webhook to ${email} failed (${status}): ${responseText}`);
        result.error = `HTTP ${status}: ${responseText.substring(0, 200)}`;
      }
      
      console.log(`[${requestId}] Webhook result:`, JSON.stringify(result, null, 2));
      resolve(result);
      
    } catch (error) {
      const errorMsg = `[${requestId}] Webhook to ${email} failed: ${error.toString()}`;
      console.error(errorMsg);
      console.error(`[${requestId}] Error details:`, error);
      
      const errorResult = { 
        status: 'error', 
        error: error.toString(),
        errorDetails: error.message || error,
        url: webhookUrl,
        requestId: requestId,
        timestamp: new Date().toISOString()
      };
      
      console.error(`[${requestId}] Webhook error result:`, JSON.stringify(errorResult, null, 2));
      resolve(errorResult);
    }
  });
}

/**
 * Escapes HTML special characters
 * @private
 */
function escapeHtml(unsafe) {
  if (!unsafe) return '';
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
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
 * Gets recent activity feed (comments, mentions, status changes)
 * @param {Object} filters - Filter options {type, timeframe, search}
 * @returns {Array} Array of activity items
 */
function getActivityFeed(filters = {}) {
  try {
    const activities = [];
    const now = new Date();
    let cutoffDate = new Date(0); // Default to beginning of time
    
    // Set cutoff date based on timeframe
    switch (filters.timeframe) {
      case 'today':
        cutoffDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        break;
      case 'week':
        cutoffDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
        break;
      case 'month':
        cutoffDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
        break;
      default:
        cutoffDate = new Date(0);
    }
    
    // Get comments if requested
    if (!filters.type || filters.type === 'all' || filters.type === 'comments') {
      const comments = getRecentComments(cutoffDate);
      activities.push(...comments.map(comment => ({
        type: 'comment',
        timestamp: comment.timestamp,
        actionItemId: comment.actionItemId,
        actionItemTitle: getActionItemTitle(comment.actionItemId),
        author: comment.author,
        content: comment.content,
        id: comment.commentId
      })));
    }
    
    // Get mentions if requested
    if (!filters.type || filters.type === 'all' || filters.type === 'mentions') {
      const mentions = getRecentMentions(cutoffDate);
      activities.push(...mentions);
    }
    
    // Get status changes if requested
    if (!filters.type || filters.type === 'all' || filters.type === 'status-changes') {
      const statusChanges = getRecentStatusChanges(cutoffDate);
      activities.push(...statusChanges);
    }
    
    // Sort by timestamp (most recent first)
    activities.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    // Apply search filter if provided
    let filteredActivities = activities;
    if (filters.search && filters.search.trim()) {
      const searchTerm = filters.search.toLowerCase();
      filteredActivities = activities.filter(activity => 
        (activity.content && activity.content.toLowerCase().includes(searchTerm)) ||
        (activity.actionItemTitle && activity.actionItemTitle.toLowerCase().includes(searchTerm)) ||
        (activity.author && activity.author.toLowerCase().includes(searchTerm))
      );
    }
    
    // Limit to 100 most recent items
    return filteredActivities.slice(0, 100);
    
  } catch (error) {
    Logger.log('ERROR in getActivityFeed: ' + error.toString());
    return [];
  }
}

/**
 * Gets recent comments after a cutoff date
 * @param {Date} cutoffDate - Only return comments after this date
 * @returns {Array} Array of recent comments
 */
function getRecentComments(cutoffDate) {
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
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const comment = {};
      
      headers.forEach((header, index) => {
        comment[header] = row[index];
      });
      
      // Filter by date
      const commentDate = new Date(comment.timestamp);
      if (commentDate >= cutoffDate) {
        comments.push(comment);
      }
    }
    
    return comments;
  } catch (error) {
    Logger.log('ERROR in getRecentComments: ' + error.toString());
    return [];
  }
}

/**
 * Gets recent mentions after a cutoff date
 * @param {Date} cutoffDate - Only return mentions after this date
 * @returns {Array} Array of recent mentions
 */
function getRecentMentions(cutoffDate) {
  try {
    // Get recent comments and filter for those with mentions
    const comments = getRecentComments(cutoffDate);
    const mentions = [];
    
    comments.forEach(comment => {
      const extractedMentions = extractMentions(comment.content || '');
      if (extractedMentions.length > 0) {
        mentions.push({
          type: 'mention',
          timestamp: comment.timestamp,
          actionItemId: comment.actionItemId,
          actionItemTitle: getActionItemTitle(comment.actionItemId),
          author: comment.author,
          content: comment.content,
          mentionedUsers: extractedMentions,
          id: comment.commentId + '_mention'
        });
      }
    });
    
    return mentions;
  } catch (error) {
    Logger.log('ERROR in getRecentMentions: ' + error.toString());
    return [];
  }
}

/**
 * Gets recent status changes after a cutoff date
 * @param {Date} cutoffDate - Only return status changes after this date
 * @returns {Array} Array of recent status changes
 */
function getRecentStatusChanges(cutoffDate) {
  try {
    // This would require a status change history table
    // For now, return empty array - can be implemented later
    return [];
  } catch (error) {
    Logger.log('ERROR in getRecentStatusChanges: ' + error.toString());
    return [];
  }
}

/**
 * Gets the title of an action item by ID
 * @param {string} actionItemId - The action item ID
 * @returns {string} The action item title
 */
function getActionItemTitle(actionItemId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEMS);
    
    if (!sheet) {
      return 'Unknown Item';
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return 'Unknown Item';
    }
    
    const headers = values[0];
    const idColIndex = headers.indexOf('actionItemId');
    const titleColIndex = headers.indexOf('title');
    
    if (idColIndex === -1 || titleColIndex === -1) {
      return 'Unknown Item';
    }
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][idColIndex] === actionItemId) {
        return values[i][titleColIndex] || 'Untitled Item';
      }
    }
    
    return 'Unknown Item';
  } catch (error) {
    Logger.log('ERROR in getActionItemTitle: ' + error.toString());
    return 'Unknown Item';
  }
}

/**
 * Gets the script URL
 * @returns {string} The URL of the web app
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
