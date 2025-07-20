/**
 * ActionItemOptions.js - Functions for handling ActionItem options and hierarchical structure
 */

/**
 * Builds all action item options into a hierarchical structure
 * @param {boolean} useCache - Whether to use cached options
 * @returns {Object} Hierarchical structure of action item options
 */
function clearCachedOptions() {
  cachedOptions = null;
  lastOptionsCacheTime = 0;
  Logger.log('DEBUGGING: Server-side action item options cache has been cleared.');
}

function buildActionItemOptions(useCache = true) {
  Logger.log('SIMPLE LOG: buildActionItemOptions in ActionItemOptions.js called with useCache=' + useCache);
  try {
    Logger.log('DEBUGGING: buildActionItemOptions called with useCache=' + useCache);
    const currentTime = new Date().getTime();
    if (useCache && cachedOptions && (currentTime - lastOptionsCacheTime < CACHE_EXPIRATION_SECONDS * 1000)) {
      Logger.log('DEBUGGING: Returning cached options.');
      return cachedOptions;
    }

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.ACTION_ITEM_OPTIONS);
    if (!sheet) throw new Error('actionItemOptions sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const columnMap = {};
    headers.forEach((header, index) => { columnMap[header] = index; });

    const options = {
      actionItems: {
        categories: {},
        selectionTypes: {},
        groups: {},
        optionsData: {}
      }
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || row[columnMap.active] === false) continue; // Skip empty or inactive rows

      const optionData = {};
      headers.forEach((header, index) => {
        const value = row[index];
        if (['requiresPatient', 'requiresProviderApproval', 'active', 'allowsRecurrence'].includes(header)) {
          optionData[header] = value === true || String(value).toUpperCase() === 'TRUE';
        } else {
          optionData[header] = value;
        }
      });

      const levels = [
        optionData.categoryLevel1,
        optionData.categoryLevel2,
        optionData.categoryLevel3,
        optionData.categoryLevel4,
        optionData.categoryLevel5
      ].filter(Boolean);

      if (levels.length === 0) continue;

      let currentLevel = options.actionItems.categories;
      let currentPath = '';
      for (let j = 0; j < levels.length; j++) {
        const levelName = levels[j];
        currentPath = currentPath ? `${currentPath}/${levelName}` : levelName;

        if (!currentLevel[levelName]) {
          currentLevel[levelName] = {
            displayName: levelName,
            subcategories: {}
          };
        }

        // A category path can be a template if it's the last one in its row.
        // So, we store the data for the current full path.
        // This overwrites intermediate paths with the more specific data from deeper rows, which is the desired behavior.
        options.actionItems.optionsData[currentPath] = optionData;

        currentLevel = currentLevel[levelName].subcategories;
      }
    }

    cachedOptions = options;
    lastOptionsCacheTime = currentTime;

    Logger.log('DEBUGGING: buildActionItemOptions returning options structure: ' + JSON.stringify(options));
    return options;
  } catch (error) {
    Logger.log('ERROR in buildActionItemOptions: ' + error.toString() + '\n' + error.stack);
    console.error('ERROR in buildActionItemOptions: ' + error.toString() + '\n' + error.stack);
    return { actionItems: { categories: {}, selectionTypes: {}, groups: {}, optionsData: {} } };
  }
}

/**
 * Gets action item templates from the actionItems sheet
 * @returns {Array} Array of template action items
 */
function getActionItemTemplates() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTION_ITEMS);
    
    if (!sheet) {
      return [];
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length <= 1) {
      return [];
    }
    
    const headers = values[0];
    const templates = [];
    
    // Find isTemplate column index
    const templateColIndex = headers.indexOf('isTemplate');
    
    if (templateColIndex === -1) {
      return [];
    }
    
    // Find matching templates
    for (let i = 1; i < values.length; i++) {
      const isTemplate = values[i][templateColIndex] === true || 
                        String(values[i][templateColIndex]).toUpperCase() === 'TRUE';
      
      if (isTemplate) {
        const template = {};
        headers.forEach((header, index) => {
          // Handle special fields
          if (header === 'tags' || header === 'mentionedUsers' || header === 'selectedOptions' || header === 'relatedIds') {
            // Convert comma-separated strings to arrays
            template[header] = values[i][index] ? String(values[i][index]).split(',').map(s => s.trim()).filter(Boolean) : [];
          } else if (header === 'isRecurring' || header === 'isTemplate' || 
                    header === 'faxSent' || header === 'visitInfoAttached' || 
                    header === 'facesheetAttached') {
            // Convert to boolean
            template[header] = values[i][index] === true || String(values[i][index]).toUpperCase() === 'TRUE';
          } else {
            // Standard field
            template[header] = values[i][index];
          }
        });
        templates.push(template);
      }
    }
    
    return templates;
  } catch (error) {
    Logger.log('ERROR in getActionItemTemplates: ' + error.toString());
    return [];
  }
}

/**
 * Creates a new action item from a template
 * @param {string} templateId - The ID of the template to use
 * @param {Object} overrides - Values to override in the template
 * @returns {Object} The created action item
 */
function createFromTemplate(templateId, overrides = {}) {
  try {
    // Get the template
    const template = getActionItemById(templateId);
    
    if (!template || !template.isTemplate) {
      throw new Error('Template not found or not a valid template');
    }
    
    // Create a copy of the template
    const newItem = JSON.parse(JSON.stringify(template));
    
    // Remove template-specific fields
    delete newItem.actionItemId;
    newItem.isTemplate = false;
    newItem.templateId = templateId;
    
    // Apply overrides
    Object.keys(overrides).forEach(key => {
      newItem[key] = overrides[key];
    });
    
    // Save the new item
    return saveActionItem(newItem);
  } catch (error) {
    Logger.log('ERROR in createFromTemplate: ' + error.toString());
    throw error;
  }
}

/**
 * Gets the full option data for a specific category path
 * @param {string} actionItemType - The action item type
 * @param {string} categoryPath - The category path
 * @returns {Object|null} The option data or null if not found
 */
function getOptionByPath(actionItemType, categoryPath) {
  try {
    const options = buildActionItemOptions();
    
    if (!options[actionItemType]) {
      return null;
    }
    
    const pathParts = categoryPath.split('/');
    let current = options[actionItemType].categories;
    
    // Navigate through the path
    for (let i = 0; i < pathParts.length; i++) {
      const part = pathParts[i];
      
      if (i === pathParts.length - 1) {
        // Last part - check if it's in options
        const category = current[part];
        if (!category) return null;
        
        if (category.options) {
          return category.options.find(option => option.displayName === part);
        }
        return null;
      } else {
        // Navigate deeper
        if (!current[part] || !current[part].subcategories) {
          return null;
        }
        current = current[part].subcategories;
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('ERROR in getOptionByPath: ' + error.toString());
    return null;
  }
}
