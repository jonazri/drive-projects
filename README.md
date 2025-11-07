# PDF Downloader for Google Sheets

A Google Apps Script that automates downloading PDF files from URLs listed in a Google Spreadsheet, saving them to a Google Shared Drive folder, and updating the spreadsheet with shareable links to the downloaded files. Features intelligent URL detection, automatic execution time management for large datasets, and resume capability for interrupted processing.

## Features

- **Batch PDF Downloads**: Process multiple PDF URLs from a spreadsheet column in a single execution
- **Automatic URL Detection**: Intelligently detects whether URLs are direct PDF links or wrapper pages with embedded PDFs
- **Smart URL Extraction**: Automatically extracts actual PDF URLs from embed wrapper pages when needed
- **Shared Drive Integration**: Saves downloaded PDFs directly to a specified Google Shared Drive folder
- **Separate Column Tracking**: Maintains both extracted PDF URLs and Drive shareable links in separate columns
- **Incremental Processing**: Automatically skips already-processed rows, allowing you to resume interrupted jobs
- **Automatic Continuation**: Handles large datasets that exceed Google Apps Script's 6-minute execution limit by automatically creating continuation triggers
- **Error Handling**: Comprehensive error handling with informative error messages in the spreadsheet
- **Custom Menu**: Easy-to-use custom menu in Google Sheets for running the script
- **Highly Configurable**: Simple configuration constants for customizing columns, folder location, timing, and behavior

## Prerequisites

- A Google account with access to Google Sheets and Google Drive
- Access to a Google Shared Drive folder where PDFs will be saved
- Basic understanding of Google Apps Script (for setup and configuration)

## Installation

1. **Open your Google Spreadsheet** where you want to use this script

2. **Open the Apps Script editor**:
   - Click on `Extensions` > `Apps Script`

3. **Copy the script**:
   - Delete any existing code in the editor
   - Copy the entire contents of `Code.gs` from this repository
   - Paste it into the Apps Script editor

4. **Save the project**:
   - Click the save icon or press `Ctrl+S` (Windows) or `Cmd+S` (Mac)
   - Give your project a name (e.g., "PDF Downloader")

## Configuration

Before running the script, you need to configure the following constants at the top of the `Code.gs` file:

### Required Configuration

```javascript
const SHARED_DRIVE_FOLDER_ID = 'YOUR_SHARED_DRIVE_FOLDER_ID_HERE';
```
- **Description**: The ID of the Shared Drive folder where PDFs will be saved
- **How to find the Shared Drive folder ID**:
  1. Open Google Drive in your web browser
  2. Navigate to the Shared Drive (not "My Drive")
  3. Open the specific folder where you want PDFs saved
  4. Look at the URL in your browser's address bar
  5. The folder ID is the long string after `/folders/`
     - Example URL: `https://drive.google.com/drive/folders/1In9KsvISrVgIOMGkhbwQOXnyky4qsLmP`
     - Folder ID: `1In9KsvISrVgIOMGkhbwQOXnyky4qsLmP`
  6. Copy the folder ID and paste it into the script, replacing `'YOUR_SHARED_DRIVE_FOLDER_ID_HERE'`
- **Important**: You must have edit access to this Shared Drive folder

### Column Configuration

```javascript
const URL_COLUMN = 'C';          // Source URLs (wrapper pages or direct PDF links)
const PDF_LINK_COLUMN = 'D';     // Extracted/direct PDF URLs
const DRIVE_LINK_COLUMN = 'E';   // Google Drive shareable links
const START_ROW = 2;             // First data row (2 = skip header)
```

- **URL_COLUMN**: The column letter containing the source URLs (can be wrapper pages or direct PDF links)
- **PDF_LINK_COLUMN**: The column where extracted or direct PDF URLs will be written
- **DRIVE_LINK_COLUMN**: The column where Google Drive shareable links will be written
- **START_ROW**: The row number to start processing from (set to 2 to skip header row)

### Execution Time Configuration

```javascript
const MAX_EXECUTION_TIME_MS = 5 * 60 * 1000;    // 5 minutes
const CONTINUATION_DELAY_MS = 10 * 1000;        // 10 seconds
```

- **MAX_EXECUTION_TIME_MS**: Maximum execution time before creating a continuation trigger (default: 5 minutes, leaving 1-minute buffer before Google's 6-minute limit)
- **CONTINUATION_DELAY_MS**: Delay before the continuation trigger fires (default: 10 seconds)

### Advanced Configuration

```javascript
const DIRECT_MODE = false;  // Override automatic detection
```

- **DIRECT_MODE**: Optional override for automatic URL detection
  - `false` (default): Automatically detects whether URLs are direct PDFs or wrapper pages
  - `true`: Treats all URLs as direct PDF links, skipping detection and extraction

## Usage

### Setting up your spreadsheet

1. **Organize your data**:
   - Place source URLs in the configured URL_COLUMN (default: Column C)
   - Leave PDF_LINK_COLUMN and DRIVE_LINK_COLUMN empty (they will be filled automatically)
   - Ensure row 1 contains headers (or adjust START_ROW if your data starts elsewhere)

2. **Example spreadsheet structure**:
   ```
   | A         | B          | C (URL Column)           | D (PDF Link)      | E (Drive Link)    |
   |-----------|------------|--------------------------|-------------------|-------------------|
   | Name      | Date       | Source URL               | PDF URL           | Shareable Link    |
   | Document1 | 2024-01-01 | https://example.com/pg1  | (auto-filled)     | (auto-filled)     |
   | Document2 | 2024-01-02 | https://example.com/f.pdf| (auto-filled)     | (auto-filled)     |
   ```

3. **Column purposes**:
   - **Column C (URL_COLUMN)**: Your source URLs - can be wrapper pages with embedded PDFs or direct PDF links
   - **Column D (PDF_LINK_COLUMN)**: The script writes the actual PDF URL here (extracted from wrapper or copied from direct link)
   - **Column E (DRIVE_LINK_COLUMN)**: The script writes the Google Drive shareable link here after saving the PDF

### Running the script

1. **First-time authorization**:
   - The first time you run the script, Google will ask you to authorize it
   - Click "Review Permissions" and grant the necessary permissions
   - The script needs access to your spreadsheet and Google Drive

2. **Using the custom menu**:
   - Open your spreadsheet
   - Look for the "PDF Downloader" menu in the menu bar (appears after the sheet loads)
   - Click `PDF Downloader` > `Download PDFs`
   - The script will process all URLs in the URL_COLUMN

3. **Alternative method** (from Apps Script editor):
   - Open the Apps Script editor (`Extensions` > `Apps Script`)
   - Select the `downloadPDFs` function from the dropdown
   - Click the "Run" button

### What happens during execution

1. The script reads all non-empty URLs from the URL_COLUMN starting at START_ROW
2. For each URL:
   - **Checks if already processed**: Skips rows where DRIVE_LINK_COLUMN already has content (resume capability)
   - **Detects URL type**: Automatically determines if the URL is a direct PDF link or a wrapper page
   - **Extracts PDF URL if needed**: If it's a wrapper page, extracts the actual PDF URL from the embed tag
   - **Writes PDF URL**: Saves the extracted or direct PDF URL to PDF_LINK_COLUMN
   - **Downloads the PDF**: Fetches the PDF file from the PDF URL
   - **Saves to Shared Drive**: Creates the file in your configured Shared Drive folder
   - **Writes Drive link**: Saves the shareable link to DRIVE_LINK_COLUMN
   - **Monitors execution time**: Checks if approaching the time limit after each row
3. If any errors occur, error messages are written to the DRIVE_LINK_COLUMN for that row
4. Processing continues for all remaining URLs even if some fail
5. If execution time approaches 5 minutes, the script automatically creates a continuation trigger and exits gracefully
6. The continuation trigger fires after 10 seconds and resumes processing from where it left off

## How It Works

### Automatic Direct PDF Detection

The script intelligently detects whether each URL is a direct PDF link or a wrapper page:

1. **HEAD Request Check**: Performs a lightweight HEAD request to check the Content-Type header
2. **Content-Type Analysis**: If the header indicates `application/pdf`, treats it as a direct PDF link
3. **Extension Fallback**: If HEAD request fails, checks if the URL ends with `.pdf`
4. **Graceful Handling**: On detection failure, assumes wrapper page and attempts extraction

**Benefits**:
- No manual configuration needed for mixed URL types
- Handles both direct PDFs and wrapper pages in the same spreadsheet
- Minimal performance impact (HEAD requests are fast)
- Automatic fallback ensures processing continues

### URL Extraction from Wrapper Pages

When a wrapper page is detected, the script:

1. Fetches the HTML content of the page
2. Parses the HTML to find `<embed>` tags with `type="application/pdf"`
3. Extracts the actual PDF URL from the `src` attribute
4. Handles protocol-relative URLs (starting with `//`) by adding `https:`
5. Writes the extracted URL to PDF_LINK_COLUMN for verification

### PDF Download
- Uses `UrlFetchApp.fetch()` to download PDF content
- Validates that the content is actually a PDF file
- Handles timeouts (30-second maximum per download)
- Manages network errors gracefully
- Follows redirects automatically

### File Saving
- Generates unique filenames based on the URL or timestamp
- Saves PDFs to the configured Shared Drive folder
- Preserves folder permissions (inherits from the Shared Drive folder)
- Google Drive automatically handles duplicate filenames
- Returns the File object for link generation

### Link Generation
- Retrieves shareable links using `file.getUrl()`
- Writes links to the DRIVE_LINK_COLUMN in the same row as the source URL
- Links respect the Shared Drive folder's sharing settings
- Preserves the PDF URL in PDF_LINK_COLUMN for reference

### Incremental Processing and Resume Capability

The script supports resuming interrupted jobs:

1. **Status Check**: Before processing each row, checks if DRIVE_LINK_COLUMN already has content
2. **Skip Processed Rows**: Automatically skips rows that have already been processed
3. **Resume from Interruption**: If the script is stopped (manually or due to errors), simply run it again
4. **No Duplicate Work**: Already-downloaded PDFs won't be re-downloaded
5. **Immediate Writes**: Uses `SpreadsheetApp.flush()` after each row to ensure data is saved immediately

**Use cases**:
- Script interrupted by user
- Script stopped due to errors
- Adding new URLs to an existing list
- Re-running after fixing configuration issues

### Automatic Continuation for Large Datasets

For datasets that exceed Google Apps Script's 6-minute execution limit:

1. **Time Monitoring**: Tracks elapsed time from the start of execution
2. **Proactive Exit**: When approaching 5 minutes (leaving 1-minute buffer), stops processing new rows
3. **Trigger Creation**: Automatically creates a time-based trigger to resume in 10 seconds
4. **Sheet Persistence**: Stores the active sheet name using PropertiesService for trigger-based execution
5. **Seamless Resume**: The trigger fires and processing continues from the first unprocessed row
6. **Automatic Cleanup**: Deletes continuation triggers when all rows are processed

**Benefits**:
- No user intervention required for large datasets
- Prevents timeout errors
- Leverages incremental processing for exact resume points
- Self-cleaning (removes triggers when complete)
- Works across multiple continuation cycles if needed

## Error Handling

The script includes comprehensive error handling:

- **Failed URL extraction**: Writes "Error: Failed to extract PDF URL from page" to DRIVE_LINK_COLUMN
- **Failed downloads**: Writes "Error: Failed to download PDF" to DRIVE_LINK_COLUMN
- **Failed saves**: Writes "Error: Failed to save PDF to Shared Drive" to DRIVE_LINK_COLUMN
- **Unexpected errors**: Writes "Error: [error message]" to DRIVE_LINK_COLUMN
- **Continued processing**: Errors on one row don't stop processing of remaining rows
- **Detailed logging**: All errors are logged to the Apps Script execution log with row numbers and URLs

### Viewing execution logs

To view detailed execution logs:
1. Open the Apps Script editor (`Extensions` > `Apps Script`)
2. Click `Executions` in the left sidebar
3. Click on a specific execution to view its logs
4. Look for entries with row numbers, URLs, and error details

## Troubleshooting

### The custom menu doesn't appear
- Refresh the spreadsheet (close and reopen it)
- Make sure the script is saved and bound to the spreadsheet
- Check that the `onOpen()` function exists in the script
- Wait a few seconds after opening the spreadsheet for the menu to load

### "Permission denied" errors
- Verify you have edit access to the Shared Drive folder
- Check that the SHARED_DRIVE_FOLDER_ID is correct
- Re-authorize the script by running it manually from the Apps Script editor
- Ensure the Shared Drive folder hasn't been moved or deleted

### PDFs aren't downloading
- Verify the URLs in URL_COLUMN are accessible (try opening them in a browser)
- Check the execution logs for specific error messages
- Ensure the source URLs aren't blocked by firewalls or authentication
- For wrapper pages, verify they contain `<embed>` tags with `type="application/pdf"`
- Check if the PDF_LINK_COLUMN shows extracted URLs (helps diagnose extraction issues)

### Links aren't being written to the spreadsheet
- Verify the PDF_LINK_COLUMN and DRIVE_LINK_COLUMN configurations are correct
- Check that the column letters are valid (A-Z, AA-ZZ, etc.)
- Ensure you have edit permissions on the spreadsheet
- Look for error messages in the DRIVE_LINK_COLUMN

### Script times out or stops processing
- **Good news**: The script has automatic continuation for large datasets!
- The script will automatically create a continuation trigger when approaching the 6-minute limit
- Wait 10 seconds and the script will resume automatically
- Check the execution logs for "CONTINUATION" messages
- If continuation triggers aren't working, you can manually re-run the script (it will skip already-processed rows)

### Continuation triggers aren't cleaning up
- The script automatically deletes triggers when all rows are processed
- If triggers remain after completion, you can manually delete them:
  1. Open the Apps Script editor
  2. Click the clock icon (Triggers) in the left sidebar
  3. Delete any triggers for the `downloadPDFs` function

### Some rows are being skipped
- The script skips rows where DRIVE_LINK_COLUMN already has content (this is the resume feature)
- To reprocess a row, clear the content in DRIVE_LINK_COLUMN for that row
- Empty rows in URL_COLUMN are also skipped (this is intentional)

### Automatic URL detection isn't working
- Check the execution logs to see what type was detected for each URL
- If detection fails, the script falls back to wrapper page extraction
- You can override detection by setting `DIRECT_MODE = true` for all direct PDFs
- Verify that direct PDF URLs return `Content-Type: application/pdf` (test with a HEAD request)

### Sheet not found error after continuation trigger
- This can happen if the sheet name was changed during processing
- The script stores the sheet name in PropertiesService at the start
- To fix: manually run the script again from the spreadsheet (not the editor)
- The script will reset and use the current active sheet

## Limitations

- **Execution time**: Google Apps Script has a 6-minute maximum execution time per execution (automatically handled by continuation triggers)
- **File size**: Very large PDF files (>50MB) may cause timeouts during download
- **Rate limiting**: Google's URL Fetch service has daily quotas (20,000 URL fetch calls per day for consumer accounts)
- **Single-threaded**: URLs are processed sequentially, not in parallel
- **Wrapper page formats**: Only detects PDFs in `<embed>` tags with `type="application/pdf"` (most common format)
- **Authentication**: Cannot download PDFs that require login or authentication
- **Shared Drive access**: Requires edit access to the target Shared Drive folder

## Technical Details

- **Language**: Google Apps Script (JavaScript-based)
- **Google Services Used**:
  - `SpreadsheetApp` - Reading/writing spreadsheet data and flushing changes
  - `UrlFetchApp` - Downloading PDFs and performing HEAD requests for URL detection
  - `DriveApp` - Saving files to Shared Drive and generating shareable links
  - `PropertiesService` - Persisting sheet name and trigger IDs between executions
  - `ScriptApp` - Creating and managing time-based continuation triggers
  - `Logger` - Execution logging for debugging
- **File Format**: Google Apps Script (.gs)
- **Architecture**: Single-file script with modular helper functions

## Contributing

This is a single-file Google Apps Script project. To contribute:
1. Test your changes thoroughly in a copy of the script
2. Ensure all configuration constants are properly documented
3. Update this README if you add new features or change functionality

## License

This project is provided as-is for use with Google Apps Script.

## Support

For issues or questions:
- Check the execution logs in the Apps Script editor
- Review the error messages written to the spreadsheet
- Ensure all configuration values are correct
- Verify you have necessary permissions for the Shared Drive folder

## Required Permissions

On first run, the script will request authorization for the following permissions:

- **View and manage spreadsheets**: Required to read URLs and write links
- **View and manage files in Google Drive**: Required to save PDFs to Shared Drive
- **Connect to external services**: Required to download PDFs from URLs
- **Run when you're not present**: Required for automatic continuation triggers

**Authorization steps**:
1. Click "Review Permissions" when prompted
2. Select your Google account
3. Click "Advanced" if you see a warning
4. Click "Go to [Project Name] (unsafe)" - this is safe, it's your own script
5. Click "Allow" to grant permissions

## Version History

- **v2.0**: Enhanced automation and intelligence
  - Automatic direct PDF detection via HEAD requests
  - Separate columns for PDF URLs and Drive links
  - Incremental processing with resume capability
  - Automatic continuation triggers for large datasets
  - Sheet name persistence for trigger-based execution
  - Configurable execution time limits

- **v1.0**: Initial release with core functionality
  - PDF URL extraction from embed pages
  - Batch PDF downloads
  - Shared Drive integration
  - Custom menu interface
  - Comprehensive error handling
