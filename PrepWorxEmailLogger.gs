/**
 * PrepWorx Email Check-in Logger - Complete Solution
 * Automatically monitors Gmail for PrepWorx shipment notifications and logs data to Google Sheets
 * 
 * Setup Instructions:
 * 1. Update CHECKIN_CONFIG.SHEET_ID with your Google Sheet ID
 * 2. Run completeSetup() function to initialize everything
 * 3. Grant required permissions when prompted
 * 4. Test with testEmailProcessing() function
 */

// ==================== CHECKIN_CONFIGURATION ====================

const CHECKIN_CONFIG = {
  // Your Google Sheet ID (found in the URL between /d/ and /edit)
  SHEET_ID: '1uK9Vv6rEQ8-WUc3core1TJtozUdLZP0bnlsf9qUsnL0',
  
  // Sheet name where data will be logged
  SHEET_NAME: 'PrepWorx Checkins',
  
  // Email settings
  EMAIL_FROM: 'beta@prepworx.io',
  EMAIL_SUBJECT_CONTAINS: 'Inbound',
  EMAIL_SUBJECT_PROCESSED: 'has been processed',
  
  // Label for processed emails (will be created automatically)
  PROCESSED_LABEL: 'PrepWorx/Processed',
  
  // Processing intervals (in minutes)
  CHECK_INTERVAL_MINUTES: 5,        // How often to check for new emails
  RATE_LIMIT_SECONDS: 30            // Minimum time between processing runs
};

// ==================== MAIN FUNCTIONS ====================

/**
 * Main function to process new PrepWorx emails
 * This function is triggered by Gmail when new emails arrive
 */
function processNewEmails() {
  try {
    Logger.info('Starting email processing...');
    
    // Rate limiting check
    if (!RateLimiter.canProcess()) {
      Logger.info('Rate limited, skipping this run');
      return;
    }
    
    // Search for unprocessed PrepWorx emails
    const searchQuery = `from:${CHECKIN_CONFIG.EMAIL_FROM} subject:"${CHECKIN_CONFIG.EMAIL_SUBJECT_CONTAINS}" subject:"${CHECKIN_CONFIG.EMAIL_SUBJECT_PROCESSED}" -label:${CHECKIN_CONFIG.PROCESSED_LABEL}`;
    const threads = GmailApp.search(searchQuery, 0, 50);
    
    Logger.info(`Found ${threads.length} unprocessed email threads`);
    
    if (threads.length === 0) {
      Logger.info('No new emails to process');
      return;
    }
    
    // Get or create the processed label
    const processedLabel = getOrCreateLabel(CHECKIN_CONFIG.PROCESSED_LABEL);
    
    // Process each thread
    let processedCount = 0;
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        if (isFromPrepWorx(message)) {
          const success = processEmail(message);
          if (success) {
            processedCount++;
            updateActivityStats('processed');
          }
        }
      }
      
      // Mark thread as processed
      thread.addLabel(processedLabel);
    }
    
    Logger.info(`Successfully processed ${processedCount} emails`);
    
  } catch (error) {
    Logger.error('Error in processNewEmails', error);
    updateActivityStats('error', { message: error.message, stack: error.stack });
    sendErrorNotification(error);
  }
}

/**
 * Process a single email message
 */
function processEmail(message) {
  try {
    Logger.info(`Processing email: ${message.getSubject()}`);
    
    // Extract email data
    const emailData = extractEmailData(message);
    
    if (!emailData) {
      Logger.warn('Could not extract data from email');
      return false;
    }
    
    // Log to Google Sheet
    const success = logToSheet(emailData);
    
    if (success) {
      Logger.info('Successfully logged data to sheet');
      return true;
    } else {
      Logger.warn('Failed to log data to sheet');
      return false;
    }
    
  } catch (error) {
    Logger.error('Error processing email', error);
    return false;
  }
}

/**
 * Check if email is from PrepWorx
 */
function isFromPrepWorx(message) {
  const from = message.getFrom().toLowerCase();
  return from.includes(CHECKIN_CONFIG.EMAIL_FROM.toLowerCase()) || 
         from.includes('prepworx');
}

// ==================== EMAIL PARSING FUNCTIONS ====================

/**
 * Extract data from PrepWorx email
 */
function extractEmailData(message) {
  try {
    const subject = message.getSubject();
    const htmlBody = message.getBody();
    const plainBody = message.getPlainBody();
    const date = message.getDate();
    
    Logger.info('Extracting data from email', { subject });
    
    // Extract shipment number from subject
    const shipmentNumber = extractShipmentNumber(subject);
    if (!shipmentNumber) {
      Logger.warn('Could not extract shipment number from subject');
      return null;
    }
    
    // Try HTML parsing first, then fall back to plain text
    let extractedData = null;
    
    if (htmlBody && htmlBody.trim() !== '') {
      extractedData = parseHtmlContent(htmlBody, shipmentNumber);
    }
    
    if (!extractedData && plainBody && plainBody.trim() !== '') {
      extractedData = parsePlainTextContent(plainBody, shipmentNumber);
    }
    
    if (!extractedData) {
      Logger.warn('Could not extract data from email content');
      return null;
    }
    
    // Combine all extracted data
    const result = {
      shipmentNumber: shipmentNumber,
      emailDate: date,
      ...extractedData
    };
    
    Logger.info('Successfully extracted data', result);
    return result;
    
  } catch (error) {
    Logger.error('Error extracting email data', error);
    return null;
  }
}

/**
 * Extract shipment number from email subject
 */
function extractShipmentNumber(subject) {
  try {
    // Pattern: "Inbound P7354920059015303168 has been processed."
    const match = subject.match(/Inbound\s+([A-Z0-9]+)/i);
    return match ? match[1] : null;
  } catch (error) {
    Logger.error('Error extracting shipment number', error);
    return null;
  }
}

/**
 * Parse HTML email content to extract shipment data
 */
function parseHtmlContent(htmlBody, shipmentNumber) {
  try {
    Logger.info('Parsing HTML content...');
    
    // Use shipment number as the main order number
    const orderNumber = shipmentNumber;
    
    // Extract the secondary code (the one between shipment number and date)
    const secondaryCode = extractSecondaryCode(htmlBody);
    
    // Extract date and time
    const dateTime = extractDateTime(htmlBody);
    
    // Extract item data from HTML table
    const items = extractItemsFromHtml(htmlBody);
    
    if (!items || items.length === 0) {
      Logger.warn('No items found in HTML');
      return null;
    }
    
    // Return data for the first item (can be extended for multiple items)
    const firstItem = items[0];
    
    return {
      orderNumber: orderNumber,
      dateTime: dateTime,
      itemName: firstItem.itemName,
      asin: firstItem.asin,
      quantity: firstItem.quantity,
      correctOrderNumber: secondaryCode || orderNumber
    };
    
  } catch (error) {
    Logger.error('Error parsing HTML content', error);
    return null;
  }
}

/**
 * Parse plain text email content as fallback
 */
function parsePlainTextContent(plainBody, shipmentNumber) {
  try {
    Logger.info('Parsing plain text content...');
    
    // Use shipment number as the main order number
    const orderNumber = shipmentNumber;
    
    // Extract the secondary code (the one between shipment number and date)
    const secondaryCode = extractSecondaryCode(plainBody);
    
    // Extract date and time
    const dateTime = extractDateTime(plainBody);
    
    // Extract item data from plain text
    const items = extractItemsFromPlainText(plainBody);
    
    if (!items || items.length === 0) {
      Logger.warn('No items found in plain text');
      return null;
    }
    
    // Return data for the first item
    const firstItem = items[0];
    
    return {
      orderNumber: orderNumber,
      dateTime: dateTime,
      itemName: firstItem.itemName,
      asin: firstItem.asin,
      quantity: firstItem.quantity,
      correctOrderNumber: secondaryCode || orderNumber
    };
    
  } catch (error) {
    Logger.error('Error parsing plain text content', error);
    return null;
  }
}

/**
 * Extract secondary code from email content (the code between shipment number and date)
 */
function extractSecondaryCode(content) {
  try {
    // Look for the pattern after shipment number and before date
    // Pattern: Pnik90cM47H5vTIke3jn
    const patterns = [
      /\n([A-Za-z0-9]{20})\n/,  // Exact 20 character alphanumeric
      /\b([A-Za-z0-9]{15,25})\b(?=\s*\d+\/\d+\/\d+)/,  // 15-25 chars before date
      /([A-Za-z0-9]{15,25})\s*\n\s*\d+\/\d+\/\d+/,  // Pattern before date line
    ];
    
    for (const pattern of patterns) {
      const match = content.match(pattern);
      if (match && match[1]) {
        const candidate = match[1].trim();
        // Validate it's not a shipment number (which starts with P and is all digits after)
        if (!candidate.match(/^P\d+$/)) {
          return candidate;
        }
      }
    }
    
    return null;
  } catch (error) {
    Logger.error('Error extracting secondary code', error);
    return null;
  }
}

/**
 * Extract date and time from email content
 */
function extractDateTime(content) {
  try {
    // Pattern: 8/19/2025, 5:20:54 PM +00:00
    const patterns = [
      /(\d{1,2}\/\d{1,2}\/\d{4},\s+\d{1,2}:\d{2}:\d{2}\s+(?:AM|PM)\s+[+-]\d{2}:\d{2})/i,
      /(\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+(?:AM|PM))/i
    ];
    
    for (const pattern of patterns) {
      const match = content.match(pattern);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
    
    return null;
  } catch (error) {
    Logger.error('Error extracting date time', error);
    return null;
  }
}

/**
 * Extract items from HTML table
 */
function extractItemsFromHtml(htmlBody) {
  try {
    const items = [];
    
    // First try to find HTML table structure
    const tableMatches = htmlBody.match(/<table[\s\S]*?<\/table>/gi);
    
    if (tableMatches) {
      for (const table of tableMatches) {
        const rows = table.match(/<tr[\s\S]*?<\/tr>/gi);
        if (rows) {
          for (const row of rows) {
            const item = parseItemFromTableRow(row);
            if (item) {
              items.push(item);
            }
          }
        }
      }
    }
    
    // If no table found, look for the pattern in the HTML
    if (items.length === 0) {
      items.push(...extractItemsFromPlainText(htmlBody));
    }
    
    return items;
  } catch (error) {
    Logger.error('Error extracting items from HTML', error);
    return [];
  }
}

/**
 * Parse item from HTML table row
 */
function parseItemFromTableRow(rowHtml) {
  try {
    // Remove HTML tags and get text content
    const text = rowHtml.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
    
    // Skip header rows
    if (text.toLowerCase().includes('item') || text.toLowerCase().includes('amount')) {
      return null;
    }
    
    return parseItemFromText(text);
  } catch (error) {
    Logger.error('Error parsing table row', error);
    return null;
  }
}

/**
 * Extract items from plain text content
 */
function extractItemsFromPlainText(content) {
  try {
    const items = [];
    
    // Look for item pattern: On Running Cloud X 3 Hunter Black 11.5 60.98101 - B0BNC66RPR	1
    const lines = content.split('\n');
    
    for (const line of lines) {
      const trimmedLine = line.trim();
      
      // Skip empty lines and headers
      if (!trimmedLine || 
          trimmedLine.toLowerCase().includes('item') || 
          trimmedLine.toLowerCase().includes('amount') ||
          trimmedLine.length < 10) {
        continue;
      }
      
      const item = parseItemFromText(trimmedLine);
      if (item) {
        items.push(item);
      }
    }
    
    return items;
  } catch (error) {
    Logger.error('Error extracting items from plain text', error);
    return [];
  }
}

/**
 * Parse individual item from text line
 */
function parseItemFromText(text) {
  try {
    // Pattern: On Running Cloud X 3 Hunter Black 11.5 60.98101 - B0BNC66RPR	1
    // Or: On Running Cloud X 3 Hunter Black 11.5 60.98101 - B0BNC66RPR 1
    
    // Extract ASIN (pattern: - B0BNC66RPR)
    const asinMatch = text.match(/-\s*([A-Z0-9]{10})/);
    if (!asinMatch) {
      return null;
    }
    
    const asin = asinMatch[1];
    
    // Extract quantity (last number, usually 1)
    const parts = text.split(/\s+/);
    let quantity = 1;
    
    // Look for quantity at the end
    for (let i = parts.length - 1; i >= 0; i--) {
      const part = parts[i];
      if (/^\d+$/.test(part) && part !== asin) {
        quantity = parseInt(part);
        break;
      }
    }
    
    // Extract item name (everything before the ASIN)
    const beforeAsin = text.substring(0, text.indexOf('-')).trim();
    
    // Remove any trailing numbers that might be part of the price/code
    const itemName = beforeAsin.replace(/\s+[\d.]+$/, '').trim();
    
    if (!itemName) {
      return null;
    }
    
    return {
      itemName: itemName,
      asin: asin,
      quantity: quantity
    };
    
  } catch (error) {
    Logger.error('Error parsing item from text', error);
    return null;
  }
}

// ==================== GOOGLE SHEETS INTEGRATION ====================

/**
 * Log extracted data to Google Sheet
 */
function logToSheet(emailData) {
  try {
    const sheet = getOrCreateSheet();
    
    if (!sheet) {
      Logger.error('Could not access Google Sheet');
      return false;
    }
    
    // Prepare row data
    const rowData = [
      emailData.dateTime || formatDate(emailData.emailDate),
      emailData.orderNumber || '',
      emailData.itemName || '',
      emailData.asin || '',
      emailData.quantity || 1,
      emailData.correctOrderNumber || emailData.orderNumber || ''
    ];
    
    // Check for duplicates
    if (isDuplicateEntry(sheet, emailData)) {
      Logger.info('Duplicate entry found, skipping...');
      return true; // Return true since it's not an error
    }
    
    // Find the next empty row
    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;
    
    // Insert the data
    const range = sheet.getRange(nextRow, 1, 1, rowData.length);
    range.setValues([rowData]);
    
    // Format the new row
    formatNewRow(sheet, nextRow, rowData.length);
    
    Logger.info(`Data logged to sheet at row ${nextRow}`, rowData);
    return true;
    
  } catch (error) {
    Logger.error('Error logging to sheet', error);
    return false;
  }
}

/**
 * Get or create the Google Sheet
 */
function getOrCreateSheet() {
  try {
    let spreadsheet;
    
    // Try to open existing spreadsheet
    try {
      spreadsheet = SpreadsheetApp.openById(CHECKIN_CONFIG.SHEET_ID);
    } catch (error) {
      Logger.warn('Could not open spreadsheet with provided ID, creating new one...');
      
      // Create new spreadsheet
      spreadsheet = SpreadsheetApp.create('PrepWorx Check-in Logger');
      Logger.info(`Created new spreadsheet: ${spreadsheet.getId()}`);
      console.log(`Please update CHECKIN_CONFIG.SHEET_ID to: ${spreadsheet.getId()}`);
    }
    
    // Get or create the specified sheet
    let sheet = spreadsheet.getSheetByName(CHECKIN_CONFIG.SHEET_NAME);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CHECKIN_CONFIG.SHEET_NAME);
      Logger.info(`Created new sheet: ${CHECKIN_CONFIG.SHEET_NAME}`);
    }
    
    return sheet;
    
  } catch (error) {
    Logger.error('Error accessing Google Sheet', error);
    return null;
  }
}

/**
 * Check if this entry already exists in the sheet
 */
function isDuplicateEntry(sheet, emailData) {
  try {
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return false; // No data rows exist
    }
    
    // Get all data
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 6);
    const values = dataRange.getValues();
    
    // Check for duplicates based on order number and item name
    for (const row of values) {
      const existingOrderNumber = row[1]; // Column B
      const existingItemName = row[2];    // Column C
      const existingASIN = row[3];        // Column D
      
      if (existingOrderNumber === emailData.orderNumber &&
          existingItemName === emailData.itemName &&
          existingASIN === emailData.asin) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    Logger.error('Error checking for duplicates', error);
    return false; // Assume not duplicate if we can't check
  }
}

/**
 * Format a newly added row
 */
function formatNewRow(sheet, rowNumber, columnCount) {
  try {
    const range = sheet.getRange(rowNumber, 1, 1, columnCount);
    
    // Alternate row colors for better readability
    if (rowNumber % 2 === 0) {
      range.setBackground('#f8f9fa');
    }
    
    // Format date column (A)
    const dateRange = sheet.getRange(rowNumber, 1);
    dateRange.setNumberFormat('mm/dd/yyyy, h:mm:ss AM/PM');
    
    // Format quantity column (E) as number
    const quantityRange = sheet.getRange(rowNumber, 5);
    quantityRange.setNumberFormat('0');
    quantityRange.setHorizontalAlignment('center');
    
    // Set borders
    range.setBorder(true, true, true, true, false, false);
    
  } catch (error) {
    Logger.error('Error formatting row', error);
  }
}

/**
 * Get or create a Gmail label
 */
function getOrCreateLabel(labelName) {
  try {
    let label = GmailApp.getUserLabelByName(labelName);
    
    if (!label) {
      // Create parent labels if needed
      const parts = labelName.split('/');
      let currentPath = '';
      
      for (let i = 0; i < parts.length; i++) {
        if (i > 0) {
          currentPath += '/';
        }
        currentPath += parts[i];
        
        const existingLabel = GmailApp.getUserLabelByName(currentPath);
        if (!existingLabel) {
          label = GmailApp.createLabel(currentPath);
          Logger.info(`Created label: ${currentPath}`);
        } else {
          label = existingLabel;
        }
      }
    }
    
    return label;
    
  } catch (error) {
    Logger.error('Error creating label', error);
    return null;
  }
}

// ==================== UTILITY FUNCTIONS ====================

/**
 * Advanced logging system
 */
class Logger {
  static log(level, message, data = null) {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] [${level}] ${message}`;
    
    console.log(logMessage);
    
    if (data) {
      console.log('Data:', JSON.stringify(data, null, 2));
    }
    
    // Save to properties for persistent logging
    this.saveToProperties(level, message, data);
  }
  
  static info(message, data = null) {
    this.log('INFO', message, data);
  }
  
  static warn(message, data = null) {
    this.log('WARN', message, data);
  }
  
  static error(message, data = null) {
    this.log('ERROR', message, data);
  }
  
  static debug(message, data = null) {
    this.log('DEBUG', message, data);
  }
  
  static saveToProperties(level, message, data) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const logs = JSON.parse(properties.getProperty('APP_LOGS') || '[]');
      
      logs.push({
        timestamp: new Date().toISOString(),
        level: level,
        message: message,
        data: data
      });
      
      // Keep only last 100 log entries
      if (logs.length > 100) {
        logs.splice(0, logs.length - 100);
      }
      
      properties.setProperty('APP_LOGS', JSON.stringify(logs));
    } catch (error) {
      console.error('Error saving to properties:', error);
    }
  }
  
  static getLogs() {
    try {
      const properties = PropertiesService.getScriptProperties();
      return JSON.parse(properties.getProperty('APP_LOGS') || '[]');
    } catch (error) {
      console.error('Error getting logs:', error);
      return [];
    }
  }
  
  static clearLogs() {
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.deleteProperty('APP_LOGS');
      Logger.info('Logs cleared');
    } catch (error) {
      console.error('Error clearing logs:', error);
    }
  }
}

/**
 * Rate limiting helper
 */
class RateLimiter {
  static canProcess() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const lastRun = properties.getProperty('LAST_PROCESS_TIME');
      const minInterval = CHECKIN_CONFIG.RATE_LIMIT_SECONDS * 1000; // Convert to milliseconds
      
      if (!lastRun) {
        properties.setProperty('LAST_PROCESS_TIME', Date.now().toString());
        return true;
      }
      
      const timeSinceLastRun = Date.now() - parseInt(lastRun);
      
      if (timeSinceLastRun >= minInterval) {
        properties.setProperty('LAST_PROCESS_TIME', Date.now().toString());
        return true;
      }
      
      Logger.info(`Rate limited. Time since last run: ${timeSinceLastRun}ms`);
      return false;
      
    } catch (error) {
      Logger.error('Error in rate limiting', error);
      return true; // Allow processing if rate limiting fails
    }
  }
}

/**
 * Format date consistently
 */
function formatDate(date) {
  try {
    if (!date) return '';
    
    if (typeof date === 'string') {
      return date;
    }
    
    const options = {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true,
      timeZoneName: 'short'
    };
    
    return date.toLocaleString('en-US', options);
    
  } catch (error) {
    Logger.error('Error formatting date', error);
    return date ? date.toString() : '';
  }
}

/**
 * Send error notification email
 */
function sendErrorNotification(error) {
  try {
    const recipient = Session.getActiveUser().getEmail();
    const subject = 'PrepWorx Email Logger Error';
    const body = `
An error occurred in the PrepWorx Email Logger:

Error: ${error.toString()}
Stack: ${error.stack || 'No stack trace available'}

Time: ${new Date()}

Please check the script logs for more details.
    `;
    
    GmailApp.sendEmail(recipient, subject, body);
    Logger.info('Error notification sent');
    
  } catch (emailError) {
    Logger.error('Could not send error notification', emailError);
  }
}

/**
 * Update activity statistics
 */
function updateActivityStats(type, data = {}) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const stats = JSON.parse(properties.getProperty('ACTIVITY_STATS') || '{}');
    
    stats.lastRun = new Date().toISOString();
    
    switch (type) {
      case 'processed':
        stats.totalProcessed = (stats.totalProcessed || 0) + 1;
        break;
      case 'error':
        stats.totalErrors = (stats.totalErrors || 0) + 1;
        stats.lastError = {
          timestamp: new Date().toISOString(),
          message: data.message || 'Unknown error',
          stack: data.stack || null
        };
        break;
    }
    
    properties.setProperty('ACTIVITY_STATS', JSON.stringify(stats));
  } catch (error) {
    Logger.error('Error updating activity stats', error);
  }
}

// ==================== SETUP AND MANAGEMENT FUNCTIONS ====================

/**
 * Complete setup wizard - runs all setup steps
 * RUN THIS FIRST after CHECKIN_CONFIGuring your SHEET_ID
 */
function completeSetup() {
  try {
    Logger.info('Starting complete setup...');
    
    // Step 1: Validate CHECKIN_CONFIGuration
    Logger.info('Step 1: Validating CHECKIN_CONFIGuration...');
    const CHECKIN_CONFIGValidation = validateCHECKIN_CONFIGuration();
    
    if (!CHECKIN_CONFIGValidation.valid) {
      Logger.error('CHECKIN_CONFIGuration validation failed', CHECKIN_CONFIGValidation.errors);
      console.error('Setup failed. Please fix CHECKIN_CONFIGuration errors:');
      CHECKIN_CONFIGValidation.errors.forEach(error => console.error('- ' + error));
      return false;
    }
    
    // Step 2: Initialize sheet
    Logger.info('Step 2: Initializing Google Sheet...');
    initializeSheet();
    
    // Step 3: Set up triggers
    Logger.info('Step 3: Setting up email triggers...');
    setupEmailTrigger();
    
    // Step 4: Test the system
    Logger.info('Step 4: Running system test...');
    const healthResult = healthCheck();
    
    if (healthResult.error) {
      Logger.error('Health check failed', healthResult.error);
      return false;
    }
    
    // Step 5: Display setup summary
    Logger.info('Step 5: Setup complete!');
    displaySetupSummary();
    
    return true;
    
  } catch (error) {
    Logger.error('Setup failed with error', error);
    console.error('Setup failed:', error.message);
    return false;
  }
}

/**
 * Set up Gmail trigger for automatic email processing
 * Run this function once to enable automatic processing
 */
function setupEmailTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processNewEmails') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new Gmail trigger
    ScriptApp.newGmailTrigger()
      .setHandlerFunction('processNewEmails')
      .create();
    
    Logger.info('Email trigger set up successfully');
    
    // Also set up a time-based trigger as backup
    ScriptApp.newTrigger('processNewEmails')
      .timeBased()
      .everyMinutes(CHECKIN_CONFIG.CHECK_INTERVAL_MINUTES)
      .create();
    
    Logger.info(`Time-based trigger set up successfully (runs every ${CHECKIN_CONFIG.CHECK_INTERVAL_MINUTES} minutes)`);
    
  } catch (error) {
    Logger.error('Error setting up triggers', error);
  }
}

/**
 * Initialize the Google Sheet with proper headers
 * Run this function once to set up your sheet
 */
function initializeSheet() {
  try {
    const sheet = getOrCreateSheet();
    
    // Check if headers already exist
    const range = sheet.getRange(1, 1, 1, 6);
    const headers = range.getValues()[0];
    
    if (headers[0] !== 'Date & Time') {
      // Set up headers
      const headerRow = [
        'Date & Time',
        'Order Number', 
        'Item Name',
        'ASIN',
        'Quantity',
        'Correct Order Number'
      ];
      
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
      
      // Format headers
      const headerRange = sheet.getRange(1, 1, 1, headerRow.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('white');
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, headerRow.length);
      
      Logger.info('Sheet initialized with headers');
    } else {
      Logger.info('Sheet already has headers');
    }
    
  } catch (error) {
    Logger.error('Error initializing sheet', error);
  }
}

/**
 * CHECKIN_CONFIGuration validator
 */
function validateCHECKIN_CONFIGuration() {
  const errors = [];
  
  // Check required CHECKIN_CONFIGuration
  if (!CHECKIN_CONFIG.SHEET_ID || CHECKIN_CONFIG.SHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
    errors.push('SHEET_ID must be CHECKIN_CONFIGured with your actual Google Sheet ID');
  }
  
  if (!CHECKIN_CONFIG.EMAIL_FROM) {
    errors.push('EMAIL_FROM must be CHECKIN_CONFIGured');
  }
  
  if (!CHECKIN_CONFIG.SHEET_NAME) {
    errors.push('SHEET_NAME must be CHECKIN_CONFIGured');
  }
  
  // Test sheet access
  try {
    const sheet = getOrCreateSheet();
    if (!sheet) {
      errors.push('Cannot access or create Google Sheet');
    }
  } catch (error) {
    errors.push(`Sheet access error: ${error.message}`);
  }
  
  // Test Gmail access
  try {
    GmailApp.getInboxThreads(0, 1);
  } catch (error) {
    errors.push(`Gmail access error: ${error.message}`);
  }
  
  if (errors.length > 0) {
    Logger.error('CHECKIN_CONFIGuration validation failed', errors);
    return { valid: false, errors: errors };
  }
  
  Logger.info('CHECKIN_CONFIGuration validation passed');
  return { valid: true, errors: [] };
}

/**
 * Health check function
 */
function healthCheck() {
  try {
    Logger.info('Starting health check...');
    
    const results = {
      timestamp: new Date().toISOString(),
      CHECKIN_CONFIGuration: validateCHECKIN_CONFIGuration(),
      triggers: checkTriggers(),
      permissions: checkPermissions(),
      recentActivity: getRecentActivity()
    };
    
    Logger.info('Health check completed', results);
    return results;
    
  } catch (error) {
    Logger.error('Health check failed', error);
    return { error: error.message };
  }
}

/**
 * Check triggers status
 */
function checkTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const emailTriggers = triggers.filter(t => t.getHandlerFunction() === 'processNewEmails');
    
    return {
      totalTriggers: triggers.length,
      emailTriggers: emailTriggers.length,
      triggers: emailTriggers.map(t => ({
        type: t.getTriggerSource().toString(),
        function: t.getHandlerFunction()
      }))
    };
  } catch (error) {
    Logger.error('Error checking triggers', error);
    return { error: error.message };
  }
}

/**
 * Check permissions
 */
function checkPermissions() {
  const permissions = {
    gmail: false,
    sheets: false,
    properties: false
  };
  
  try {
    // Test Gmail permission
    GmailApp.getInboxThreads(0, 1);
    permissions.gmail = true;
  } catch (error) {
    Logger.warn('Gmail permission check failed', error.message);
  }
  
  try {
    // Test Sheets permission
    if (CHECKIN_CONFIG.SHEET_ID !== 'YOUR_GOOGLE_SHEET_ID_HERE') {
      SpreadsheetApp.openById(CHECKIN_CONFIG.SHEET_ID);
    }
    permissions.sheets = true;
  } catch (error) {
    Logger.warn('Sheets permission check failed', error.message);
  }
  
  try {
    // Test Properties permission
    PropertiesService.getScriptProperties().getProperty('TEST');
    permissions.properties = true;
  } catch (error) {
    Logger.warn('Properties permission check failed', error.message);
  }
  
  return permissions;
}

/**
 * Get recent activity statistics
 */
function getRecentActivity() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const stats = JSON.parse(properties.getProperty('ACTIVITY_STATS') || '{}');
    
    return {
      lastRun: stats.lastRun || null,
      totalProcessed: stats.totalProcessed || 0,
      totalErrors: stats.totalErrors || 0,
      lastError: stats.lastError || null
    };
  } catch (error) {
    Logger.error('Error getting recent activity', error);
    return { error: error.message };
  }
}

/**
 * Display setup summary
 */
function displaySetupSummary() {
  try {
    const sheet = getOrCreateSheet();
    const sheetUrl = sheet ? sheet.getParent().getUrl() : 'Could not access sheet';
    const triggers = ScriptApp.getProjectTriggers();
    const emailTriggers = triggers.filter(t => t.getHandlerFunction() === 'processNewEmails');
    
    const summary = `
=== PREPWORX EMAIL LOGGER SETUP COMPLETE ===

✅ CHECKIN_CONFIGuration validated
✅ Google Sheet initialized
✅ Email triggers created (${emailTriggers.length} active)
✅ System health check passed

📊 Google Sheet: ${sheetUrl}
📧 Monitoring emails from: ${CHECKIN_CONFIG.EMAIL_FROM}
🏷️  Processed emails will be labeled: ${CHECKIN_CONFIG.PROCESSED_LABEL}

🔄 The system is now running automatically!

Next steps:
1. Send a test email or wait for real PrepWorx emails
2. Check your Google Sheet for new entries
3. Run healthCheck() anytime to verify system status

⚠️  If you encounter issues:
- Run healthCheck() for diagnostics
- Check the execution logs
- Use testEmailProcessing() to debug
    `;
    
    console.log(summary);
    Logger.info('Setup summary displayed');
    
  } catch (error) {
    Logger.error('Error displaying setup summary', error);
  }
}

/**
 * Manual test function - processes emails from the last 24 hours
 * Use this to test the system
 */
function testEmailProcessing() {
  try {
    Logger.info('Running test email processing...');
    
    // Search for PrepWorx emails from last 24 hours
    const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000);
    const searchQuery = `from:${CHECKIN_CONFIG.EMAIL_FROM} after:${Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd')}`;
    
    const threads = GmailApp.search(searchQuery, 0, 10);
    Logger.info(`Found ${threads.length} recent email threads for testing`);
    
    let processedCount = 0;
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        if (isFromPrepWorx(message)) {
          Logger.info(`Testing email: ${message.getSubject()}`);
          const emailData = extractEmailData(message);
          
          if (emailData) {
            Logger.info('Extracted data', emailData);
            processedCount++;
          }
        }
      }
    }
    
    Logger.info(`Test completed. Processed ${processedCount} emails.`);
    
  } catch (error) {
    Logger.error('Error in test', error);
  }
}

/**
 * Clear all processed labels (for testing)
 */
function clearProcessedLabels() {
  try {
    const label = GmailApp.getUserLabelByName(CHECKIN_CONFIG.PROCESSED_LABEL);
    if (label) {
      const threads = label.getThreads();
      for (const thread of threads) {
        thread.removeLabel(label);
      }
      Logger.info(`Removed processed label from ${threads.length} threads`);
    }
  } catch (error) {
    Logger.error('Error clearing processed labels', error);
  }
}

/**
 * Get sheet statistics
 */
function getSheetStats() {
  try {
    const sheet = getOrCreateSheet();
    if (!sheet) return null;
    
    const lastRow = sheet.getLastRow();
    const totalEntries = Math.max(0, lastRow - 1); // Subtract header row
    
    const stats = {
      totalEntries: totalEntries,
      lastUpdate: totalEntries > 0 ? sheet.getRange(lastRow, 1).getValue() : null,
      sheetUrl: sheet.getParent().getUrl()
    };
    
    Logger.info('Sheet statistics', stats);
    return stats;
    
  } catch (error) {
    Logger.error('Error getting sheet stats', error);
    return null;
  }
}
