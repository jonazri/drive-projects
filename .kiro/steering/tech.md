# Technology Stack

## Platform
- Google Apps Script (JavaScript runtime)
- Sheet-bound script attached to Google Spreadsheet

## Core APIs
- `SpreadsheetApp` - spreadsheet manipulation
- `UrlFetchApp` - HTTP requests and PDF downloads
- `DriveApp` - Google Drive file operations
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
- `URL_COLUMN` - column letter containing PDF URLs (e.g., 'C')
- `LINK_COLUMN` - column letter for output links (e.g., 'D')
- `START_ROW` - first data row, typically 2 to skip headers

### Logging
View execution logs in Apps Script editor: **View > Logs** or **Executions**

### Permissions
First run requires authorization for:
- Spreadsheet access
- Drive file creation
- External URL fetching
