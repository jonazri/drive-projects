# Technology Stack

## Platform
- Google Apps Script (JavaScript runtime)
- Sheet-bound script attached to Google Spreadsheet

## Core APIs
- `SpreadsheetApp` - spreadsheet manipulation
- `UrlFetchApp` - HTTP requests, PDF downloads, and HEAD requests for URL detection
- `DriveApp` - Google Drive file operations
- `PropertiesService` - persisting sheet name between trigger-based executions
- `ScriptApp` - creating and managing time-based continuation triggers
- `Logger` - execution logging

## Key Libraries
None - uses native Google Apps Script services only

## Development

### Running the Script
- Open the spreadsheet
- Use custom menu: **PDF Downloader > Download PDFs**
- Or run `downloadPDFs()` from Apps Script editor

### Configuration
Edit constants at top of `Code.gs`:
- `SHARED_DRIVE_FOLDER_ID` - target folder ID from Drive URL
- `URL_COLUMN` - column letter containing source URLs (e.g., 'C')
- `PDF_LINK_COLUMN` - column letter for extracted PDF URLs (e.g., 'D')
- `DRIVE_LINK_COLUMN` - column letter for Drive shareable links (e.g., 'E')
- `START_ROW` - first data row, typically 2 to skip headers
- `MAX_EXECUTION_TIME_MS` - time limit before creating continuation trigger (default: 5 minutes)
- `CONTINUATION_DELAY_MS` - delay before continuation trigger fires (default: 10 seconds)
- `DIRECT_MODE` - optional override to skip URL detection (default: false)

### Logging
View execution logs in Apps Script editor: **View > Logs** or **Executions**

### Permissions
First run requires authorization for:
- Spreadsheet access
- Drive file creation
- External URL fetching
