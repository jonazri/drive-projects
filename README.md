# PDF Downloader for Google Sheets

A Google Apps Script that automates downloading PDF files from URLs listed in a Google Spreadsheet, saving them to a Google Shared Drive folder, and updating the spreadsheet with shareable links to the downloaded files.

## Features

- **Batch PDF Downloads**: Process multiple PDF URLs from a spreadsheet column in a single execution
- **Automatic URL Extraction**: Extracts actual PDF URLs from embed wrapper pages
- **Shared Drive Integration**: Saves downloaded PDFs directly to a specified Google Shared Drive folder
- **Link Updates**: Automatically writes shareable links back to the spreadsheet
- **Error Handling**: Comprehensive error handling with informative error messages in the spreadsheet
- **Custom Menu**: Easy-to-use custom menu in Google Sheets for running the script
- **Configurable**: Simple configuration constants for customizing column letters, folder location, and starting row

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
- **How to find**: 
  1. Open the Shared Drive folder in Google Drive
  2. Copy the folder ID from the URL: `https://drive.google.com/drive/folders/[FOLDER_ID_HERE]`

### Column Configuration

```javascript
const URL_COLUMN = 'C';
const LINK_COLUMN = 'D';
const START_ROW = 2;
```

- **URL_COLUMN**: The column letter containing the PDF URLs to download (e.g., 'A', 'B', 'C')
- **LINK_COLUMN**: The column letter where shareable links will be written (typically adjacent to URL_COLUMN)
- **START_ROW**: The row number to start processing from (set to 2 to skip header row)

## Usage

### Setting up your spreadsheet

1. **Organize your data**:
   - Place PDF URLs in the configured URL_COLUMN (default: Column C)
   - Ensure row 1 contains headers (or adjust START_ROW if your data starts elsewhere)

2. **Example spreadsheet structure**:
   ```
   | A         | B          | C (URL Column)           | D (Link Column)    |
   |-----------|------------|--------------------------|-------------------|
   | Name      | Date       | PDF URL                  | Shareable Link    |
   | Document1 | 2024-01-01 | https://example.com/pdf1 | (will be filled)  |
   | Document2 | 2024-01-02 | https://example.com/pdf2 | (will be filled)  |
   ```

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
   - Extracts the actual PDF URL (if the URL points to an embed wrapper page)
   - Downloads the PDF file
   - Saves it to the specified Shared Drive folder
   - Writes the shareable link to the LINK_COLUMN
3. If any errors occur, error messages are written to the LINK_COLUMN for that row
4. Processing continues for all remaining URLs even if some fail

## How It Works

### URL Extraction
The script can handle URLs that point to embed wrapper pages. It parses the HTML to find `<embed>` tags with `type="application/pdf"` and extracts the actual PDF URL from the `src` attribute.

### PDF Download
- Uses `UrlFetchApp.fetch()` to download PDF content
- Validates that the content is actually a PDF file
- Handles timeouts (30-second maximum per download)
- Manages network errors gracefully

### File Saving
- Generates unique filenames based on the URL or timestamp
- Saves PDFs to the configured Shared Drive folder
- Preserves folder permissions (inherits from the Shared Drive folder)
- Google Drive automatically handles duplicate filenames

### Link Generation
- Retrieves shareable links using `file.getUrl()`
- Writes links to the LINK_COLUMN in the same row as the source URL
- Links respect the Shared Drive folder's sharing settings

## Error Handling

The script includes comprehensive error handling:

- **Failed URL extraction**: Writes "Error: Failed to extract PDF URL from page" to LINK_COLUMN
- **Failed downloads**: Writes "Error: Failed to download PDF" to LINK_COLUMN
- **Failed saves**: Writes "Error: Failed to save PDF to Shared Drive" to LINK_COLUMN
- **Unexpected errors**: Writes "Error: [error message]" to LINK_COLUMN
- All errors are logged to the Apps Script execution log for debugging

### Viewing execution logs

To view detailed execution logs:
1. Open the Apps Script editor
2. Click `Executions` in the left sidebar
3. Click on a specific execution to view its logs

## Troubleshooting

### The custom menu doesn't appear
- Refresh the spreadsheet (close and reopen it)
- Make sure the script is saved and bound to the spreadsheet
- Check that the `onOpen()` function exists in the script

### "Permission denied" errors
- Verify you have edit access to the Shared Drive folder
- Check that the SHARED_DRIVE_FOLDER_ID is correct
- Re-authorize the script by running it manually from the Apps Script editor

### PDFs aren't downloading
- Verify the URLs in URL_COLUMN are accessible
- Check that the URLs point to actual PDF files
- Review the execution logs for specific error messages
- Ensure the source URLs aren't blocked by firewalls or authentication

### Links aren't being written to the spreadsheet
- Verify the LINK_COLUMN configuration is correct
- Check that the column letter is valid (A-Z, AA-ZZ, etc.)
- Ensure you have edit permissions on the spreadsheet

### Script times out
- Google Apps Script has a 6-minute execution limit for custom functions
- If processing many files, consider running the script multiple times on different ranges
- Check for network issues that might cause slow downloads

## Limitations

- **Execution time**: Google Apps Script custom menu functions have a 6-minute maximum execution time
- **File size**: Very large PDF files may cause timeouts during download
- **Rate limiting**: Google's URL Fetch service has daily quotas (20,000 URL fetch calls per day for consumer accounts)
- **Single-threaded**: URLs are processed sequentially, not in parallel

## Technical Details

- **Language**: Google Apps Script (JavaScript-based)
- **Google Services Used**:
  - SpreadsheetApp (for reading/writing spreadsheet data)
  - UrlFetchApp (for downloading PDFs)
  - DriveApp (for saving files to Shared Drive)
- **File Format**: Google Apps Script (.gs)

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

## Version History

- **v1.0**: Initial release with core functionality
  - PDF URL extraction from embed pages
  - Batch PDF downloads
  - Shared Drive integration
  - Custom menu interface
  - Comprehensive error handling
