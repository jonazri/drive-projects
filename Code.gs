/**
 * PDF Downloader for Google Sheets
 *
 * This script reads PDF URLs from a specified column in the active sheet,
 * downloads each PDF, saves it to a Google Shared Drive folder, and writes
 * the shareable link back to an adjacent column.
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

/**
 * The ID of the Shared Drive folder where PDFs will be saved.
 * To find this ID, open the folder in Google Drive and copy the ID from the URL:
 * https://drive.google.com/drive/folders/[FOLDER_ID_HERE]
 */
const SHARED_DRIVE_FOLDER_ID = '1In9KsvISrVgIOMGkhbwQOXnyky4qsLmP';

/**
 * The column letter containing the PDF URLs to download.
 * Example: 'A' for column A, 'C' for column C, etc.
 */
const URL_COLUMN = 'C';

/**
 * The column letter where shareable links to saved PDFs will be written.
 * This should typically be the column adjacent to URL_COLUMN.
 * Example: 'B' for column B, 'D' for column D, etc.
 */
const LINK_COLUMN = 'D';

/**
 * The row number to start processing URLs from.
 * Set to 2 to skip the header row, or adjust based on your sheet structure.
 */
const START_ROW = 2;

/**
 * Maximum execution time in milliseconds before creating a continuation trigger.
 * Set to 5 minutes (300000 ms) to leave a 1-minute buffer before the 6-minute
 * Google Apps Script execution time limit. This ensures the script can gracefully
 * exit and schedule a continuation before being forcefully terminated.
 */
const MAX_EXECUTION_TIME_MS = 5 * 60 * 1000; // 5 minutes = 300000 ms

/**
 * Delay in milliseconds before the continuation trigger fires.
 * Set to 10 seconds (10000 ms) to allow a brief pause between script executions,
 * preventing rapid consecutive runs and allowing time for any pending operations
 * to complete.
 */
const CONTINUATION_DELAY_MS = 10 * 1000; // 10 seconds = 10000 ms

/**
 * The name of the function to call when the continuation trigger fires.
 * This should match the main function name that processes PDFs.
 */
const TRIGGER_FUNCTION_NAME = 'downloadPDFs';

/**
 * Property key used to store continuation trigger IDs in script properties.
 * This allows us to track and manage only the triggers created by this script
 * for continuation purposes, without affecting user-created triggers.
 */
const CONTINUATION_TRIGGER_PROPERTY_KEY = 'PDF_DOWNLOADER_CONTINUATION_TRIGGERS';

// ============================================================================
// MENU SETUP
// ============================================================================

/**
 * Creates a custom menu in the spreadsheet when it opens.
 * This function runs automatically when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF Downloader')
      .addItem('Download PDFs', 'downloadPDFs')
      .addToUi();
}

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Main function to download PDFs from URLs in the spreadsheet.
 * Reads URLs from URL_COLUMN starting at START_ROW, downloads each PDF,
 * saves it to the Shared Drive folder, and writes the shareable link to LINK_COLUMN.
 * Supports incremental processing - skips rows that already have content in LINK_COLUMN,
 * allowing the script to resume from where it left off if execution halts.
 * Includes automatic time monitoring to handle Google Apps Script execution time limits.
 */
function downloadPDFs() {
  // Record start time for execution time monitoring (Requirement 8.1)
  const startTime = Date.now();
  Logger.log('Starting PDF download process at: ' + new Date(startTime).toISOString());

  // Get the active spreadsheet and sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // Get the last row with content in the URL column
  const lastRow = sheet.getLastRow();

  // If there are no rows to process, exit early
  if (lastRow < START_ROW) {
    Logger.log('No URLs to process. Sheet has no data starting from row ' + START_ROW);
    // Clean up any continuation triggers since there's nothing to process (Requirement 8.6)
    deleteContinuationTriggers();
    return;
  }

  // Read all values from both URL_COLUMN and LINK_COLUMN starting from START_ROW
  const urlColumnIndex = columnLetterToIndex(URL_COLUMN);
  const linkColumnIndex = columnLetterToIndex(LINK_COLUMN);
  const numRows = lastRow - START_ROW + 1;

  const urlRange = sheet.getRange(START_ROW, urlColumnIndex, numRows, 1);
  const urlValues = urlRange.getValues();

  const linkRange = sheet.getRange(START_ROW, linkColumnIndex, numRows, 1);
  const linkValues = linkRange.getValues();

  Logger.log('Processing ' + numRows + ' rows starting from row ' + START_ROW);

  // Loop through each row and process non-empty URLs
  for (let i = 0; i < urlValues.length; i++) {
    // Check if we should continue processing based on elapsed time (Requirement 8.2)
    if (!shouldContinueProcessing(startTime)) {
      const elapsedTime = Date.now() - startTime;
      Logger.log('CONTINUATION: Execution time limit approaching (' + (elapsedTime / 1000) + ' seconds elapsed)');
      Logger.log('CONTINUATION: Stopping at row ' + (START_ROW + i) + ' to avoid timeout');
      Logger.log('CONTINUATION: Creating trigger to resume processing in ' + (CONTINUATION_DELAY_MS / 1000) + ' seconds');

      // Create a continuation trigger to resume processing (Requirement 8.3)
      createContinuationTrigger();

      // Exit the loop gracefully (Requirement 8.5)
      return;
    }
    const currentRow = START_ROW + i;
    const url = urlValues[i][0];
    const existingLink = linkValues[i][0];

    // Skip empty cells
    if (!url || url.toString().trim() === '') {
      Logger.log('Row ' + currentRow + ': Skipping empty URL');
      continue;
    }

    // Check if LINK_COLUMN already has content - if so, skip this row (resume capability)
    if (existingLink && existingLink.toString().trim() !== '') {
      Logger.log('Row ' + currentRow + ': Skipping already processed row (has link or error message)');
      continue;
    }

    Logger.log('Row ' + currentRow + ': Processing URL: ' + url);

    try {
      // Extract the actual PDF URL from the embed wrapper page
      const extractedPdfUrl = extractPDFUrl(url);

      if (!extractedPdfUrl) {
        // Extraction failed - write error message to LINK_COLUMN
        const linkCell = sheet.getRange(currentRow, linkColumnIndex);
        linkCell.setValue('Error: Failed to extract PDF URL from page');
        SpreadsheetApp.flush(); // Ensure data is written immediately
        Logger.log('ERROR: Row ' + currentRow + ' - Failed to extract PDF URL from: ' + url);
        // Continue processing remaining URLs (Requirement 2.4)
        continue;
      }

      // Save the extracted PDF URL to column D (same as LINK_COLUMN for now)
      // Note: This will be overwritten with the shareable link if download succeeds
      const linkCell = sheet.getRange(currentRow, linkColumnIndex);
      linkCell.setValue(extractedPdfUrl);
      SpreadsheetApp.flush(); // Ensure data is written immediately
      Logger.log('Row ' + currentRow + ': Extracted PDF URL: ' + extractedPdfUrl);

      // Download the PDF from the extracted URL
      const pdfBlob = downloadPDF(extractedPdfUrl);

      if (!pdfBlob) {
        // Download failed - write error message to LINK_COLUMN (Requirement 2.3)
        linkCell.setValue('Error: Failed to download PDF');
        SpreadsheetApp.flush(); // Ensure data is written immediately
        Logger.log('ERROR: Row ' + currentRow + ' - Failed to download PDF from: ' + extractedPdfUrl);
        // Continue processing remaining URLs (Requirement 2.4)
        continue;
      }

      // Save the PDF to Shared Drive
      const savedFile = savePDFToSharedDrive(pdfBlob, extractedPdfUrl, SHARED_DRIVE_FOLDER_ID);

      if (!savedFile) {
        // Save failed - write error message to LINK_COLUMN
        linkCell.setValue('Error: Failed to save PDF to Shared Drive');
        SpreadsheetApp.flush(); // Ensure data is written immediately
        Logger.log('ERROR: Row ' + currentRow + ' - Failed to save PDF to Shared Drive for URL: ' + extractedPdfUrl);
        // Continue processing remaining URLs (Requirement 2.4)
        continue;
      }

      // Get the shareable link from the saved file
      const shareableLink = savedFile.getUrl();

      // Write the shareable link to LINK_COLUMN in the same row
      linkCell.setValue(shareableLink);
      SpreadsheetApp.flush(); // Ensure data is written immediately

      Logger.log('Row ' + currentRow + ': Successfully saved PDF and updated link: ' + shareableLink);

    } catch (error) {
      // Handle any unexpected errors (Requirement 7.4)
      const linkCell = sheet.getRange(currentRow, linkColumnIndex);
      linkCell.setValue('Error: ' + error.message);
      SpreadsheetApp.flush(); // Ensure data is written immediately
      Logger.log('ERROR: Row ' + currentRow + ' - Unexpected exception for URL: ' + url);
      Logger.log('ERROR: ' + error.message);
      Logger.log('Stack trace: ' + error.stack);
      // Continue processing remaining URLs (Requirement 2.4)
    }

    // Check time after processing each row to ensure we don't timeout mid-operation
    if (!shouldContinueProcessing(startTime)) {
      const elapsedTime = Date.now() - startTime;
      Logger.log('CONTINUATION: Execution time limit approaching after processing row ' + currentRow);
      Logger.log('CONTINUATION: Elapsed time: ' + (elapsedTime / 1000) + ' seconds');
      Logger.log('CONTINUATION: Creating trigger to resume processing in ' + (CONTINUATION_DELAY_MS / 1000) + ' seconds');

      // Create a continuation trigger to resume processing (Requirement 8.3)
      createContinuationTrigger();

      // Exit the loop gracefully (Requirement 8.5)
      return;
    }
  }

  // All rows processed successfully - clean up any continuation triggers (Requirement 8.6)
  Logger.log('All rows processed successfully');
  deleteContinuationTriggers();
  Logger.log('Finished processing URLs');
}

/**
 * Helper function to convert column letter to column index (1-based).
 * Example: 'A' -> 1, 'B' -> 2, 'Z' -> 26, 'AA' -> 27
 */
function columnLetterToIndex(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Extracts the actual PDF URL from an embed wrapper page.
 * Fetches the page HTML, parses it to find the embed tag with type="application/pdf",
 * and extracts the src attribute value.
 *
 * @param {string} pageUrl - The URL of the wrapper page containing the embedded PDF
 * @returns {string|null} The extracted PDF URL if found, null otherwise
 */
function extractPDFUrl(pageUrl) {
  try {
    Logger.log('Extracting PDF URL from: ' + pageUrl);

    // Fetch the page HTML
    const options = {
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      followRedirects: true,
      timeout: 30000 // 30 seconds timeout
    };

    const response = UrlFetchApp.fetch(pageUrl, options);
    const responseCode = response.getResponseCode();

    // Check if the request was successful
    if (responseCode !== 200) {
      Logger.log('ERROR: Failed to fetch page. HTTP status code: ' + responseCode + ' for URL: ' + pageUrl);
      return null;
    }

    // Get the HTML content
    const html = response.getContentText();

    // Parse HTML to find the embed tag with type="application/pdf"
    // Using regex to find: <embed ... type="application/pdf" ... src="..." ...>
    // This regex looks for embed tags with type="application/pdf" and extracts the src attribute
    const embedRegex = /<embed[^>]*type=["']application\/pdf["'][^>]*src=["']([^"']+)["'][^>]*>/i;
    const embedRegexAlt = /<embed[^>]*src=["']([^"']+)["'][^>]*type=["']application\/pdf["'][^>]*>/i;

    let match = html.match(embedRegex);
    if (!match) {
      match = html.match(embedRegexAlt);
    }

    if (match && match[1]) {
      let pdfUrl = match[1];

      // Handle protocol-relative URLs (starting with //)
      if (pdfUrl.startsWith('//')) {
        pdfUrl = 'https:' + pdfUrl;
        Logger.log('Converted protocol-relative URL to: ' + pdfUrl);
      }

      Logger.log('Successfully extracted PDF URL: ' + pdfUrl);
      return pdfUrl;
    }

    Logger.log('ERROR: No embed tag with type="application/pdf" found in the page: ' + pageUrl);
    return null;

  } catch (error) {
    Logger.log('ERROR: Exception while extracting PDF URL from ' + pageUrl + ' - ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
    return null;
  }
}

/**
 * Downloads a PDF from the specified URL.
 *
 * @param {string} url - The URL of the PDF to download
 * @returns {Blob|null} The PDF blob if successful, null if download fails
 */
function downloadPDF(url) {
  try {
    Logger.log('Attempting to download PDF from: ' + url);

    // Fetch the PDF from the URL with a 30-second timeout
    const options = {
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      followRedirects: true,
      timeout: 30000 // 30 seconds timeout as per requirement 2.5
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    // Check if the request was successful (HTTP 200)
    if (responseCode !== 200) {
      Logger.log('ERROR: Failed to download PDF. HTTP status code: ' + responseCode + ' for URL: ' + url);
      return null;
    }

    // Get the blob from the response
    const blob = response.getBlob();

    // Verify that the content is a PDF
    const contentType = blob.getContentType();
    if (contentType !== 'application/pdf' && !contentType.includes('pdf')) {
      Logger.log('ERROR: Downloaded content is not a PDF. Content type: ' + contentType + ' for URL: ' + url);
      return null;
    }

    Logger.log('Successfully downloaded PDF from: ' + url);
    return blob;

  } catch (error) {
    // Handle network errors and timeouts (Requirement 2.3, 2.4)
    Logger.log('ERROR: Exception while downloading PDF from ' + url + ' - ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
    return null;
  }
}

/**
 * Saves a PDF blob to the specified Shared Drive folder.
 *
 * @param {Blob} blob - The PDF blob to save
 * @param {string} url - The original URL (used for filename generation)
 * @param {string} folderId - The ID of the Shared Drive folder
 * @returns {GoogleAppsScript.Drive.File|null} The saved File object if successful, null if save fails
 */
function savePDFToSharedDrive(blob, url, folderId) {
  try {
    Logger.log('Attempting to save PDF to Shared Drive folder: ' + folderId);

    // Get the Shared Drive folder by ID
    const folder = DriveApp.getFolderById(folderId);

    // Generate a filename from the URL
    const filename = getFileName(url);

    // Set the blob name to the generated filename
    blob.setName(filename);

    // Create the file in the Shared Drive folder
    const file = folder.createFile(blob);

    Logger.log('Successfully saved PDF: ' + filename + ' (File ID: ' + file.getId() + ')');
    return file;

  } catch (error) {
    // Handle errors (e.g., invalid folder ID, permission issues)
    Logger.log('ERROR: Failed to save PDF to Shared Drive for URL: ' + url);
    Logger.log('ERROR: ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
    return null;
  }
}

/**
 * Generates a filename for the PDF based on the URL or timestamp.
 * Ensures the filename has a .pdf extension and handles special characters.
 *
 * @param {string} url - The URL to extract filename from
 * @returns {string} A valid filename with .pdf extension
 */
function getFileName(url) {
  try {
    // Try to extract filename from URL
    // Remove query parameters and fragments
    let cleanUrl = url.split('?')[0].split('#')[0];

    // Get the last part of the URL path
    const urlParts = cleanUrl.split('/');
    let filename = urlParts[urlParts.length - 1];

    // If filename is empty or doesn't look like a file, use timestamp
    if (!filename || filename.length === 0 || filename.indexOf('.') === -1) {
      filename = 'pdf_' + new Date().getTime() + '.pdf';
    } else {
      // Clean up special characters (keep only alphanumeric, dots, hyphens, underscores)
      filename = filename.replace(/[^a-zA-Z0-9._-]/g, '_');

      // Ensure .pdf extension
      if (!filename.toLowerCase().endsWith('.pdf')) {
        // Remove any existing extension and add .pdf
        const lastDotIndex = filename.lastIndexOf('.');
        if (lastDotIndex > 0) {
          filename = filename.substring(0, lastDotIndex) + '.pdf';
        } else {
          filename = filename + '.pdf';
        }
      }
    }

    // Handle potential duplicate filenames by checking if file exists
    // Note: Apps Script will automatically handle duplicates by creating a new file
    // with the same name (Google Drive allows duplicate names)

    return filename;

  } catch (error) {
    // If anything goes wrong, fall back to timestamp-based name
    Logger.log('Error generating filename, using timestamp: ' + error.message);
    return 'pdf_' + new Date().getTime() + '.pdf';
  }
}

// ============================================================================
// TIME MONITORING FUNCTIONS
// ============================================================================

/**
 * Checks if the script should continue processing based on elapsed execution time.
 * Returns false if the elapsed time exceeds MAX_EXECUTION_TIME_MS, indicating
 * that the script should stop processing and create a continuation trigger.
 *
 * @param {number} startTime - The timestamp (from Date.now()) when processing started
 * @returns {boolean} True if processing should continue, false if time limit exceeded
 */
function shouldContinueProcessing(startTime) {
  const elapsedTime = Date.now() - startTime;
  return elapsedTime < MAX_EXECUTION_TIME_MS;
}

/**
 * Creates a time-based trigger to continue processing after a delay.
 * Deletes any existing continuation triggers before creating a new one to avoid
 * duplicate triggers. The trigger will fire after CONTINUATION_DELAY_MS (10 seconds)
 * and call the function specified in TRIGGER_FUNCTION_NAME.
 *
 * This function is called when the script approaches the execution time limit
 * and needs to schedule a continuation to process remaining rows.
 */
function createContinuationTrigger() {
  try {
    Logger.log('Creating continuation trigger to resume processing in ' + (CONTINUATION_DELAY_MS / 1000) + ' seconds');

    // Delete any existing continuation triggers first to avoid duplicates
    deleteContinuationTriggers();

    // Create a new time-based trigger to fire after CONTINUATION_DELAY_MS
    const trigger = ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
        .timeBased()
        .after(CONTINUATION_DELAY_MS)
        .create();

    // Store the trigger ID in script properties so we can identify and clean up
    // only our continuation triggers without affecting user-created triggers
    const properties = PropertiesService.getScriptProperties();
    const existingTriggers = properties.getProperty(CONTINUATION_TRIGGER_PROPERTY_KEY);
    const triggerIds = existingTriggers ? JSON.parse(existingTriggers) : [];
    triggerIds.push(trigger.getUniqueId());
    properties.setProperty(CONTINUATION_TRIGGER_PROPERTY_KEY, JSON.stringify(triggerIds));

    Logger.log('Continuation trigger created successfully. Trigger ID: ' + trigger.getUniqueId());
    Logger.log('Script will resume processing in ' + (CONTINUATION_DELAY_MS / 1000) + ' seconds');

  } catch (error) {
    Logger.log('ERROR: Failed to create continuation trigger - ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
  }
}

/**
 * Deletes all continuation triggers created by this script.
 * Uses script properties to identify only the triggers created for continuation purposes,
 * ensuring that user-created triggers calling the same function are not affected.
 *
 * This function is called:
 * 1. Before creating a new continuation trigger (to avoid duplicates)
 * 2. When all processing is complete (to clean up)
 */
function deleteContinuationTriggers() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const storedTriggers = properties.getProperty(CONTINUATION_TRIGGER_PROPERTY_KEY);

    if (!storedTriggers) {
      Logger.log('No continuation triggers to delete');
      return;
    }

    const triggerIds = JSON.parse(storedTriggers);
    const allTriggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;

    // Delete only triggers whose IDs are stored in our property
    for (let i = 0; i < allTriggers.length; i++) {
      const trigger = allTriggers[i];
      const triggerId = trigger.getUniqueId();

      if (triggerIds.includes(triggerId)) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
        Logger.log('Deleted continuation trigger with ID: ' + triggerId);
      }
    }

    // Clear the stored trigger IDs
    properties.deleteProperty(CONTINUATION_TRIGGER_PROPERTY_KEY);

    if (deletedCount > 0) {
      Logger.log('Deleted ' + deletedCount + ' continuation trigger(s)');
    }

  } catch (error) {
    Logger.log('ERROR: Failed to delete continuation triggers - ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
  }
}
