# PrepWorx Email Check-in Logger

Automatically monitors Gmail for PrepWorx shipment notifications and logs data to Google Sheets.

## Recent Improvements (v2.0)

### Fixed Email Stacking Issue
- **Problem**: When multiple emails arrived in quick succession, the script would process them as a batch and mark entire threads as processed, missing subsequent individual emails.
- **Solution**: Changed from processing email threads to processing individual messages, ensuring each email is handled separately.

### Enhanced Email Processing
- **More Frequent Checking**: Reduced check interval from 5 minutes to 2 minutes to catch stacked emails faster.
- **Individual Message Processing**: Each email is now processed and labeled individually, preventing batch processing issues.
- **Better Error Handling**: Improved error handling for individual messages with detailed logging.

### New Functions for Missing Emails

#### `checkForMissedEmails()`
- Checks for emails from the last 6 hours that might have been missed
- Automatically processes any unprocessed emails found
- Use this to catch emails that might have been missed due to stacking

#### `processSpecificMissingEmails()`
- Specifically processes the missing emails you mentioned:
  - 40010104347 - ASIN B0C5QJPS52
  - 50004567675 - ASIN B0D2RLD37V
- Run this function to process these specific missing emails

#### `searchForSpecificEmails(searchTerm)`
- Search for emails containing specific terms (shipment numbers, ASINs, etc.)
- Returns detailed information about found emails
- Use this to locate specific missing emails

#### `forceReprocessAll()`
- Removes all processed labels and reprocesses all emails
- Use this if you suspect many emails were missed
- **Warning**: This will reprocess ALL emails, so use only when necessary

## Setup Instructions

1. Update `CHECKIN_CONFIG.SHEET_ID` with your Google Sheet ID
2. Run `completeSetup()` function to initialize everything
3. Grant required permissions when prompted
4. Test with `testEmailProcessing()` function

## Usage

### Automatic Processing
The script runs automatically every 2 minutes and processes new emails as they arrive.

### Manual Processing of Missing Emails
If you notice emails were missed:

1. **Check for recent missed emails**:
   ```javascript
   checkForMissedEmails()
   ```

2. **Process specific missing emails**:
   ```javascript
   processSpecificMissingEmails()
   ```

3. **Search for specific emails**:
   ```javascript
   searchForSpecificEmails("40010104347")
   searchForSpecificEmails("B0C5QJPS52")
   ```

4. **Force reprocess all emails** (use with caution):
   ```javascript
   forceReprocessAll()
   ```

### Monitoring and Debugging
- Run `healthCheck()` to verify system status
- Check execution logs for detailed information
- Use `getSheetStats()` to see processing statistics

## Configuration

The script is configured to:
- Check for new emails every 2 minutes
- Process up to 50 emails per run
- Filter emails with subject "Inbound *** has been processed"
- Label processed emails as "PrepWorx/Processed"

## Troubleshooting

### Emails Still Being Missed
1. Run `checkForMissedEmails()` to catch recent missed emails
2. Check if the emails are actually in your Gmail
3. Verify the email subjects match the expected format
4. Run `forceReprocessAll()` if many emails were missed

### Performance Issues
- The script now processes emails more frequently but individually
- Each email is processed separately to avoid stacking issues
- Check execution logs for any errors or warnings

## Support

If you continue to experience issues:
1. Run `healthCheck()` and check the results
2. Review the execution logs for errors
3. Use the manual processing functions to catch missed emails
4. Consider running `forceReprocessAll()` to reprocess all emails
