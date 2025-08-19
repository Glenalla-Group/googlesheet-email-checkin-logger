# PrepWorx Email Check-in Logger

Automatically monitors Gmail for PrepWorx shipment notifications and logs data to Google Sheets.

## Features

- **Automatic Email Monitoring**: Uses Gmail triggers to process emails in real-time
- **Smart Data Extraction**: Parses both HTML and plain text emails
- **Duplicate Detection**: Prevents duplicate entries in your spreadsheet
- **Error Handling**: Comprehensive logging and error notifications
- **Configurable Timing**: Easy-to-adjust check intervals and rate limiting
- **Health Monitoring**: Built-in diagnostics and status checking

## Setup Instructions

### 1. Create Google Apps Script Project

1. Go to [Google Apps Script](https://script.google.com)
2. Click "New Project"
3. Replace the default `Code.gs` content with the single `PrepWorxEmailLogger.gs` file

### 2. Create Google Sheet

1. Create a new Google Sheet or use an existing one
2. Copy the Sheet ID from the URL (the long string between `/d/` and `/edit`)
   - Example: `https://docs.google.com/spreadsheets/d/1ABC123...XYZ/edit#gid=0`
   - Sheet ID: `1ABC123...XYZ`

### 3. Configure the Script

1. In `PrepWorxEmailLogger.gs`, update the `CONFIG` object:
   ```javascript
   const CONFIG = {
     SHEET_ID: 'YOUR_ACTUAL_SHEET_ID_HERE',   // Replace with your Sheet ID
     SHEET_NAME: 'PrepWorx Checkins',         // Sheet tab name
     EMAIL_FROM: 'beta@prepworx.io',          // Email address to monitor
     EMAIL_SUBJECT_CONTAINS: 'Inbound',       // Subject must contain this
     EMAIL_SUBJECT_PROCESSED: 'has been processed',
     PROCESSED_LABEL: 'PrepWorx/Processed',   // Gmail label for processed emails
     
     // Processing intervals (in minutes)
     CHECK_INTERVAL_MINUTES: 5,               // How often to check for new emails
     RATE_LIMIT_SECONDS: 30                   // Minimum time between processing runs
   };
   ```

### 4. Set Up Permissions

1. Click the "Run" button on the `initializeSheet` function
2. Authorize the required permissions:
   - **Gmail**: Read and modify emails, create labels
   - **Google Sheets**: Read and write to your spreadsheet
   - **Google Apps Script**: Manage triggers and properties

### 5. Initialize the System

Run these functions in order (click the function name and press "Run"):

1. **`initializeSheet()`**: Sets up your Google Sheet with proper headers
2. **`validateConfiguration()`**: Checks if everything is configured correctly
3. **`setupEmailTrigger()`**: Creates the Gmail trigger for automatic processing
4. **`testEmailProcessing()`**: Tests the system with recent emails

### 6. Verify Setup

1. Run `healthCheck()` to verify everything is working
2. Check the execution log for any errors
3. Send a test email or wait for a real PrepWorx email

## Usage

### Automatic Operation

Once set up, the system runs automatically:
- Gmail trigger activates when new emails arrive (instant processing)
- Backup time-based trigger runs every `CHECK_INTERVAL_MINUTES` (default: 5 minutes)
- Rate limiting prevents processing more than once every `RATE_LIMIT_SECONDS` (default: 30 seconds)
- Emails are processed and data is logged to your sheet
- Processed emails are labeled to prevent reprocessing

### Manual Functions

You can run these functions manually when needed:

- **`processNewEmails()`**: Manually process pending emails
- **`testEmailProcessing()`**: Test with recent emails
- **`healthCheck()`**: Check system status
- **`getSheetStats()`**: View statistics
- **`clearProcessedLabels()`**: Remove processed labels (for testing)

### Google Sheet Columns

Your sheet will have these columns:
1. **Date & Time**: When the shipment was processed (e.g., `8/19/2025, 5:20:54 PM +00:00`)
2. **Order Number**: The shipment number from email subject (e.g., `P7354920059015303168`)
3. **Item Name**: Product description (e.g., `On Running Cloud X 3 Hunter Black 11.5`)
4. **ASIN**: Amazon product identifier (e.g., `B0BNC66RPR`)
5. **Quantity**: Number of items (e.g., `1`)
6. **Correct Order Number**: Secondary code from email body (e.g., `Pnik90cM47H5vTIke3jn`)

## Troubleshooting

### Common Issues

1. **"Cannot access spreadsheet"**
   - Verify the Sheet ID is correct
   - Check that the script has permission to access your sheet

2. **"No emails being processed"**
   - Run `healthCheck()` to verify triggers are set up
   - Check that emails match the configured criteria
   - Verify Gmail permissions are granted

3. **"Duplicate entries"**
   - The system automatically prevents duplicates
   - If you see duplicates, check the duplicate detection logic

4. **"Parsing errors"**
   - Run `testEmailProcessing()` to see detailed logs
   - Check if email format has changed

### Debugging

1. **View Logs**: Check the Apps Script execution log
2. **Health Check**: Run `healthCheck()` for system status
3. **Test Mode**: Use `testEmailProcessing()` to debug issues
4. **Manual Processing**: Run `processNewEmails()` to force processing

### Getting Help

1. Check the execution logs in Google Apps Script
2. Use the built-in health check and diagnostic functions
3. Review the email parsing logic if data extraction fails

## Email Format Requirements

The system expects emails with this format:

**From**: PrepWorx Beta <beta@prepworx.io>
**Subject**: Inbound [SHIPMENT_NUMBER] has been processed.

**Body**: Contains:
- Order number (alphanumeric code)
- Date and time
- Item table with: Item Name - ASIN [tab] Quantity

## Security Notes

- The script only processes emails from the configured sender
- Processed emails are labeled to prevent reprocessing
- All data stays within your Google account
- No external services are used

## Customization

### Timing Configuration

Easily adjust processing intervals in the CONFIG section:

```javascript
// Processing intervals (in minutes)
CHECK_INTERVAL_MINUTES: 5,        // How often to check for new emails
RATE_LIMIT_SECONDS: 30            // Minimum time between processing runs
```

**Examples:**
- Check every minute: `CHECK_INTERVAL_MINUTES: 1`
- Check every 10 minutes: `CHECK_INTERVAL_MINUTES: 10`
- Faster rate limiting: `RATE_LIMIT_SECONDS: 15`

### Other Customizations

You can also modify:
- Sheet column layout and formatting
- Email search criteria and parsing logic
- Data extraction and validation rules
- Error handling and notification settings

## File Structure

**Single File Solution**: All functionality is contained in `PrepWorxEmailLogger.gs`
- Main processing functions
- Email parsing (HTML and plain text)
- Google Sheets integration
- Error handling and logging
- Setup and configuration functions

## Version History

- **v1.0**: Initial release with automatic email processing and Google Sheets integration
- **v1.1**: Single file solution with configurable timing intervals
