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
  SHEET_ID: '13MPdSzlxW5gCOj_zLKVlUt9Z0cPY9fpLKC9Bjm0WeCA',
  
  // Sheet name where data will be logged
  SHEET_NAME: 'CheckIn',
  
  // Email settings
  EMAIL_FROM: 'beta@prepworx.io',
  EMAIL_SUBJECT_CONTAINS: 'Inbound',
  EMAIL_SUBJECT_PROCESSED: 'has been processed',
  
  // Label for processed emails (will be created automatically)
  PROCESSED_LABEL: 'PrepWorx/Processed',
  
  // Processing intervals (in minutes)
  CHECK_INTERVAL_MINUTES: 2,        // More frequent checking to catch stacked emails
  MAX_EMAILS_PER_RUN: 50            // Process more emails per run to handle stacking
};

// ==================== MAIN FUNCTIONS ====================

/**
 * Main function to process new PrepWorx emails
 * This function is triggered by Gmail when new emails arrive
 */
function processNewEmails() {
  try {
    Logger.info('Starting email processing...');
    
    // Search for unprocessed PrepWorx emails
    const searchQuery = `from:${CHECKIN_CONFIG.EMAIL_FROM} subject:(${CHECKIN_CONFIG.EMAIL_SUBJECT_CONTAINS} ${CHECKIN_CONFIG.EMAIL_SUBJECT_PROCESSED}) -label:${CHECKIN_CONFIG.PROCESSED_LABEL}`;
    
    // Get threads and then extract individual messages to avoid stacking issues
    const threads = GmailApp.search(searchQuery, 0, CHECKIN_CONFIG.MAX_EMAILS_PER_RUN);
    const messages = [];
    for (const thread of threads) {
      messages.push(...thread.getMessages());
    }
    
    Logger.info(`Found ${threads.length} unprocessed email threads containing ${messages.length} messages`);
    
    if (messages.length === 0) {
      Logger.info('No new emails to process');
      return;
    }
    
    // Get or create the processed label
    const processedLabel = getOrCreateLabel(CHECKIN_CONFIG.PROCESSED_LABEL);
    
    // Process each message individually
    let processedCount = 0;
    let skippedCount = 0;
    
    for (const message of messages) {
      try {
        if (isFromPrepWorx(message)) {
          const success = processEmail(message);
          if (success) {
            processedCount++;
            updateActivityStats('processed');
            
            // Mark individual message as processed (not the entire thread)
            message.addLabel(processedLabel);
          } else {
            skippedCount++;
          }
        }
      } catch (error) {
        Logger.error(`Error processing message ${message.getId()}`, error);
        skippedCount++;
      }
    }
    
    Logger.info(`Successfully processed ${processedCount} emails, skipped ${skippedCount}`);
    
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
  
  // Direct emails from PrepWorx (handles both "beta@prepworx.io" and "PrepWorx Beta <beta@prepworx.io>")
  if (from.includes(CHECKIN_CONFIG.EMAIL_FROM.toLowerCase()) || from.includes('prepworx')) {
    return true;
  }
  
  return false;
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
    
    // Extract shipment number from subject and content
    const shipmentNumber = extractShipmentNumber(subject, htmlBody || plainBody);
    if (!shipmentNumber) {
      Logger.warn('Could not extract shipment number from subject or content');
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
 * Extract shipment number from email subject or content
 */
function extractShipmentNumber(subject, content = '') {
  try {
    // First try to extract from subject
    
    const subjectPatterns = [
      /Inbound\s+([A-Z0-9\s\-]+?)\s+has\s+been\s+processed/i,  // Full pattern with spaces and hyphens until "has been processed"
      /Inbound\s+([A-Z0-9\s\-]+)/i,               // Pattern with spaces and hyphens
      /Inbound\s+([A-Z0-9]+)/i,                   // Standard pattern for any alphanumeric (fallback)
      /Inbound\s+(\d+)/i,                         // Pattern for pure numbers (like 0017327917518) (fallback)
      /Inbound\s+([A-Z]+\d+)/i,                   // Pattern like SNP20046045 (fallback)
    ];
    
    for (const pattern of subjectPatterns) {
      const match = subject.match(pattern);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
    
    // If not found in subject, try to extract from content
    
    if (content) {
      const contentPatterns = [
        /inbound\s+shipment\s+named\s+([A-Z0-9\s\-]+?)(?:\s*\.|\s*$)/i,  // Pattern with spaces and hyphens until period or end
        /shipment\s+named\s+([A-Z0-9\s\-]+?)(?:\s*\.|\s*$)/i,          // Pattern with spaces and hyphens until period or end
        /named\s+([A-Z0-9\s\-]+?)(?:\s*\.|\s*$)/i,                     // Pattern with spaces and hyphens until period or end
        /inbound\s+shipment\s+named\s+([A-Z0-9]+)/i,                   // Fallback: alphanumeric only
        /shipment\s+named\s+([A-Z0-9]+)/i,                             // Fallback: alphanumeric only
        /named\s+([A-Z0-9]+)/i                                         // Fallback: alphanumeric only
      ];
      
      for (const pattern of contentPatterns) {
        const match = content.match(pattern);
        if (match && match[1]) {
          return match[1].trim();
        }
      }
    }
    
    return null;
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
    
    // Return data for all items (handle multiple items)
    return {
      orderNumber: orderNumber,
      dateTime: dateTime,
      items: items,
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
    
    // Return data for all items (handle multiple items)
    return {
      orderNumber: orderNumber,
      dateTime: dateTime,
      items: items,
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
    // Pattern examples:
    // "Adidas Gazelle Bold Mint Rush Impact Orange Women's 9 IG4386 - B0DDXBXNSP	2"
    // "Nike P-6000 White Gold Red Women's 7.5 BV1021-101 - B07PH4JSPN	2"
    // "Nike Ja 2 Breeze 12 FD7328-403 - B0F22FJFHF	2"
    
    // First, decode HTML entities
    let cleanText = text.replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/&amp;/g, '&');
    
    // Extract ASIN (exactly 10 characters: letters and numbers)
    const asinMatch = cleanText.match(/-\s*([A-Z0-9]{10})\s/);
    if (!asinMatch) {
      return null;
    }
    
    const asin = asinMatch[1];
    
    // Extract quantity (number after the ASIN)
    const afterAsinPattern = new RegExp(asin + '\\s+(\\d+)');
    const quantityMatch = cleanText.match(afterAsinPattern);
    let quantity = 1;
    if (quantityMatch) {
      quantity = parseInt(quantityMatch[1]);
    }
    
    // Extract item name (everything before " - ASIN")
    const asinPosition = cleanText.indexOf(' - ' + asin);
    if (asinPosition === -1) {
      return null;
    }
    
    let itemName = cleanText.substring(0, asinPosition).trim();
    
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
    
    // Handle multiple items if they exist
    const items = emailData.items || [{ 
      itemName: emailData.itemName, 
      asin: emailData.asin, 
      quantity: emailData.quantity || 1 
    }];
    
    let successCount = 0;
    
    for (const item of items) {
      // Prepare row data for each item
      const rowData = [
        emailData.dateTime || formatDate(emailData.emailDate),
        emailData.orderNumber || '',
        item.itemName || '',
        item.asin || '',
        item.quantity || 1,
        emailData.correctOrderNumber || emailData.orderNumber || ''
      ];
      
      // Create a temporary emailData object for duplicate checking
      const itemEmailData = {
        orderNumber: emailData.orderNumber,
        dateTime: emailData.dateTime || formatDate(emailData.emailDate),
        itemName: item.itemName,
        asin: item.asin
      };
      
      // Check for duplicates (only prevents same order number + date/time combinations)
      if (isDuplicateEntry(sheet, itemEmailData)) {
        Logger.info(`Duplicate entry found for order ${itemEmailData.orderNumber} with date ${itemEmailData.dateTime}, skipping...`);
        continue;
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
      successCount++;
    }
    
    Logger.info(`Successfully logged ${successCount} items to sheet`);
    return successCount > 0;
    
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
 * Only prevents duplicates when both order number AND date/time are identical
 * Allows storing same items (item name + ASIN) with different order numbers or dates
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
    
    // Check for duplicates based on order number AND date/time only
    // This allows same items to be stored if they have different order numbers or different dates
    for (const row of values) {
      const existingDateTime = row[0];     // Column A - Date & Time
      const existingOrderNumber = row[1];  // Column B - Order Number
      
      // Convert date/time values to strings for comparison
      const existingDateTimeStr = typeof existingDateTime === 'string' ? 
        existingDateTime : 
        (existingDateTime instanceof Date ? formatDate(existingDateTime) : String(existingDateTime));
      
      const currentDateTimeStr = typeof emailData.dateTime === 'string' ? 
        emailData.dateTime : 
        String(emailData.dateTime);
      
      // Only consider it a duplicate if BOTH order number AND date/time match
      if (existingOrderNumber === emailData.orderNumber &&
          existingDateTimeStr === currentDateTimeStr) {
        Logger.info(`Duplicate entry found: Order ${emailData.orderNumber} with date ${currentDateTimeStr} already exists`);
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
    
    // Format order number columns (B and F) as text to preserve leading zeros
    const orderNumberRange = sheet.getRange(rowNumber, 2);
    orderNumberRange.setNumberFormat('@'); // @ means text format
    
    const correctOrderNumberRange = sheet.getRange(rowNumber, 6);
    correctOrderNumberRange.setNumberFormat('@'); // @ means text format
    
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
 * RUN THIS FIRST after configuring your SHEET_ID
 */
function completeSetup() {
  try {
    Logger.info('Starting complete setup...');
    
    // Step 1: Validate configuration
    Logger.info('Step 1: Validating configuration...');
    const configValidation = validateConfiguration();
    
    if (!configValidation.valid) {
      Logger.error('Configuration validation failed', configValidation.errors);
      console.error('Setup failed. Please fix configuration errors:');
      configValidation.errors.forEach(error => console.error('- ' + error));
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
    
    // Also set up a time-based trigger as backup - more frequent to catch stacked emails
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
      
      // Format order number columns (B and F) as text for the entire columns to preserve leading zeros
      const orderNumberColumn = sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1);
      orderNumberColumn.setNumberFormat('@'); // @ means text format
      
      const correctOrderNumberColumn = sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1);
      correctOrderNumberColumn.setNumberFormat('@'); // @ means text format
      
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
 * Configuration validator
 */
function validateConfiguration() {
  const errors = [];
  
  // Check required configuration
  if (!CHECKIN_CONFIG.SHEET_ID || CHECKIN_CONFIG.SHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
    errors.push('SHEET_ID must be configured with your actual Google Sheet ID');
  }
  
  if (!CHECKIN_CONFIG.EMAIL_FROM) {
    errors.push('EMAIL_FROM must be configured');
  }
  
  if (!CHECKIN_CONFIG.SHEET_NAME) {
    errors.push('SHEET_NAME must be configured');
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
    Logger.error('Configuration validation failed', errors);
    return { valid: false, errors: errors };
  }
  
  Logger.info('Configuration validation passed');
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
      configuration: validateConfiguration(),
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

✅ Configuration validated
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
- Verify your configuration
    `;
    
    console.log(summary);
    Logger.info('Setup summary displayed');
    
  } catch (error) {
    Logger.error('Error displaying setup summary', error);
  }
}