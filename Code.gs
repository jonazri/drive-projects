/**
 * PDF Downloader for Google Sheets
 * Automates batch downloading of PDFs from URLs, with automatic URL detection
 * and execution time management for large datasets.
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

/** Shared Drive folder ID where PDFs will be saved (from folder URL) */
const SHARED_DRIVE_FOLDER_ID = '1In9KsvISrVgIOMGkhbwQOXnyky4qsLmP';

/** Column containing source URLs (wrapper pages or direct PDF links) */
const URL_COLUMN = 'C';

/** Column for extracted/direct PDF URLs */
const PDF_LINK_COLUMN = 'D';

/** Column for Google Drive shareable links */
const DRIVE_LINK_COLUMN = 'E';

/** First row to process (2 = skip header) */
const START_ROW = 2;

/** Max execution time (5 min) before creating continuation trigger */
const MAX_EXECUTION_TIME_MS = 5 * 60 * 1000;

/** Delay (10 sec) before continuation trigger fires */
const CONTINUATION_DELAY_MS = 10 * 1000;

/** Function name for continuation triggers */
const TRIGGER_FUNCTION_NAME = 'downloadPDFs';

/** Property key for storing continuation trigger IDs */
const CONTINUATION_TRIGGER_PROPERTY_KEY = 'PDF_DOWNLOADER_CONTINUATION_TRIGGERS';

/** Property key for storing active sheet name between executions */
const SHEET_NAME_PROPERTY_KEY = 'PDF_DOWNLOADER_ACTIVE_SHEET_NAME';

/**
 * Direct mode override (default: false for automatic detection)
 * true = treat all URLs as direct PDFs, false = auto-detect URL type
 */
const DIRECT_MODE = false;

// ============================================================================
// MENU SETUP
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PDF Downloader')
    .addItem('Download PDFs', 'downloadPDFs')
    .addToUi();
}

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Main function to download PDFs from URLs in the spreadsheet.
 * Supports automatic URL detection, incremental processing, and time-based continuation.
 */
function downloadPDFs() {
  const startTime = Date.now();
  Logger.log('Starting PDF download process at: ' + new Date(startTime).toISOString());

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrRestoreSheet(spreadsheet);

  if (!sheet) return;

  const lastRow = sheet.getLastRow();

  if (lastRow < START_ROW) {
    Logger.log('No URLs to process. Sheet has no data starting from row ' + START_ROW);
    cleanupAndExit();
    return;
  }

  const urlColumnIndex = columnLetterToIndex(URL_COLUMN);
  const pdfLinkColumnIndex = columnLetterToIndex(PDF_LINK_COLUMN);
  const driveLinkColumnIndex = columnLetterToIndex(DRIVE_LINK_COLUMN);
  const numRows = lastRow - START_ROW + 1;

  const urlValues = sheet.getRange(START_ROW, urlColumnIndex, numRows, 1).getValues();
  const driveLinkValues = sheet.getRange(START_ROW, driveLinkColumnIndex, numRows, 1).getValues();

  Logger.log('Processing ' + numRows + ' rows starting from row ' + START_ROW);

  for (let i = 0; i < urlValues.length; i++) {
    if (!shouldContinueProcessing(startTime)) {
      logContinuation(START_ROW + i);
      createContinuationTrigger();
      return;
    }

    const currentRow = START_ROW + i;
    const url = urlValues[i][0];
    const existingDriveLink = driveLinkValues[i][0];

    if (!url || url.toString().trim() === '') {
      Logger.log('Row ' + currentRow + ': Skipping empty URL');
      continue;
    }

    if (existingDriveLink && existingDriveLink.toString().trim() !== '') {
      Logger.log('Row ' + currentRow + ': Skipping already processed row');
      continue;
    }

    Logger.log('Row ' + currentRow + ': Processing URL: ' + url);

    try {
      const pdfUrl = getPDFUrl(url, currentRow, sheet, pdfLinkColumnIndex);

      if (!pdfUrl) {
        writeError(sheet, currentRow, driveLinkColumnIndex, 'Failed to extract PDF URL from page');
        continue;
      }

      const pdfBlob = downloadPDF(pdfUrl);
      if (!pdfBlob) {
        writeError(sheet, currentRow, driveLinkColumnIndex, 'Failed to download PDF');
        continue;
      }

      const savedFile = savePDFToSharedDrive(pdfBlob, pdfUrl, SHARED_DRIVE_FOLDER_ID);
      if (!savedFile) {
        writeError(sheet, currentRow, driveLinkColumnIndex, 'Failed to save PDF to Shared Drive');
        continue;
      }

      const shareableLink = savedFile.getUrl();
      sheet.getRange(currentRow, driveLinkColumnIndex).setValue(shareableLink);
      SpreadsheetApp.flush();

      Logger.log('Row ' + currentRow + ': Successfully saved PDF and updated Drive link');

    } catch (error) {
      writeError(sheet, currentRow, driveLinkColumnIndex, error.message);
      Logger.log('ERROR: Row ' + currentRow + ' - ' + error.message);
    }

    if (!shouldContinueProcessing(startTime)) {
      logContinuation(currentRow);
      createContinuationTrigger();
      return;
    }
  }

  Logger.log('All rows processed successfully');
  cleanupAndExit();
}

/**
 * Converts column letter to column index (1-based).
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

/**
 * Gets or restores the active sheet using PropertiesService.
 * Returns null if stored sheet name is invalid.
 */
function getOrRestoreSheet(spreadsheet) {
  const properties = PropertiesService.getScriptProperties();
  const storedSheetName = properties.getProperty(SHEET_NAME_PROPERTY_KEY);

  if (storedSheetName) {
    Logger.log('Resuming from trigger - retrieving sheet: ' + storedSheetName);
    const sheet = spreadsheet.getSheetByName(storedSheetName);

    if (!sheet) {
      Logger.log('ERROR: Stored sheet "' + storedSheetName + '" not found');
      properties.deleteProperty(SHEET_NAME_PROPERTY_KEY);
      return null;
    }

    return sheet;
  }

  const sheet = spreadsheet.getActiveSheet();
  properties.setProperty(SHEET_NAME_PROPERTY_KEY, sheet.getName());
  Logger.log('Stored sheet name for continuation: ' + sheet.getName());
  return sheet;
}

/**
 * Determines the PDF URL from a source URL using automatic detection or DIRECT_MODE.
 * Writes the PDF URL to PDF_LINK_COLUMN and returns it.
 */
function getPDFUrl(url, row, sheet, pdfLinkColumnIndex) {
  let pdfUrl;

  if (DIRECT_MODE) {
    pdfUrl = url;
    Logger.log('Row ' + row + ': Using direct PDF URL (DIRECT_MODE)');
  } else {
    const isDirect = isDirectPDFUrl(url);

    if (isDirect) {
      pdfUrl = url;
      Logger.log('Row ' + row + ': Detected direct PDF URL');
    } else {
      Logger.log('Row ' + row + ': Detected wrapper page, extracting PDF URL');
      pdfUrl = extractPDFUrl(url);

      if (pdfUrl) {
        Logger.log('Row ' + row + ': Extracted PDF URL: ' + pdfUrl);
      }
    }
  }

  if (pdfUrl) {
    sheet.getRange(row, pdfLinkColumnIndex).setValue(pdfUrl);
    SpreadsheetApp.flush();
  }

  return pdfUrl;
}

/**
 * Writes an error message to the specified cell and logs it.
 */
function writeError(sheet, row, columnIndex, message) {
  sheet.getRange(row, columnIndex).setValue('Error: ' + message);
  SpreadsheetApp.flush();
  Logger.log('ERROR: Row ' + row + ' - ' + message);
}

/**
 * Logs continuation information.
 */
function logContinuation(row) {
  Logger.log('CONTINUATION: Execution time limit approaching');
  Logger.log('CONTINUATION: Stopping at row ' + row + ' to avoid timeout');
  Logger.log('CONTINUATION: Creating trigger to resume in ' + (CONTINUATION_DELAY_MS / 1000) + ' seconds');
}

/**
 * Cleans up triggers and properties, then exits.
 */
function cleanupAndExit() {
  deleteContinuationTriggers();
  PropertiesService.getScriptProperties().deleteProperty(SHEET_NAME_PROPERTY_KEY);
  Logger.log('Cleanup complete');
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Checks if a URL points directly to a PDF file.
 * Uses HEAD request to check Content-Type, falls back to .pdf extension check.
 */
function isDirectPDFUrl(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'head',
      muteHttpExceptions: true,
      followRedirects: true,
      timeout: 10000
    });

    if (response.getResponseCode() === 200) {
      const headers = response.getHeaders();
      const contentType = headers['Content-Type'] || headers['content-type'] || '';

      if (contentType.includes('application/pdf')) {
        Logger.log('Direct PDF detected via Content-Type');
        return true;
      }
    }

    const cleanUrl = url.toLowerCase().split('?')[0].split('#')[0];
    if (cleanUrl.endsWith('.pdf')) {
      Logger.log('Direct PDF detected via .pdf extension');
      return true;
    }

    return false;

  } catch (error) {
    Logger.log('Error checking URL type, assuming wrapper page: ' + error.message);
    return false;
  }
}

/**
 * Extracts the PDF URL from an embed wrapper page.
 * Parses HTML to find embed tag with type="application/pdf" and extracts src attribute.
 */
function extractPDFUrl(pageUrl) {
  try {
    const response = UrlFetchApp.fetch(pageUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      timeout: 30000
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('Failed to fetch page. HTTP status: ' + response.getResponseCode());
      return null;
    }

    const html = response.getContentText();
    const embedRegex = /<embed[^>]*type=["']application\/pdf["'][^>]*src=["']([^"']+)["'][^>]*>/i;
    const embedRegexAlt = /<embed[^>]*src=["']([^"']+)["'][^>]*type=["']application\/pdf["'][^>]*>/i;

    let match = html.match(embedRegex) || html.match(embedRegexAlt);

    if (match && match[1]) {
      let pdfUrl = match[1];

      if (pdfUrl.startsWith('//')) {
        pdfUrl = 'https:' + pdfUrl;
      }

      return pdfUrl;
    }

    Logger.log('No embed tag with type="application/pdf" found');
    return null;

  } catch (error) {
    Logger.log('Exception while extracting PDF URL: ' + error.message);
    return null;
  }
}

/**
 * Downloads a PDF from the specified URL.
 */
function downloadPDF(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      timeout: 30000
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('Failed to download PDF. HTTP status: ' + response.getResponseCode());
      return null;
    }

    const blob = response.getBlob();
    const contentType = blob.getContentType();

    if (!contentType.includes('pdf')) {
      Logger.log('Downloaded content is not a PDF. Content type: ' + contentType);
      return null;
    }

    Logger.log('Successfully downloaded PDF');
    return blob;

  } catch (error) {
    Logger.log('Exception while downloading PDF: ' + error.message);
    return null;
  }
}

/**
 * Saves a PDF blob to the specified Shared Drive folder.
 */
function savePDFToSharedDrive(blob, url, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const filename = getFileName(url);
    blob.setName(filename);
    const file = folder.createFile(blob);

    Logger.log('Successfully saved PDF: ' + filename);
    return file;

  } catch (error) {
    Logger.log('Failed to save PDF to Shared Drive: ' + error.message);
    return null;
  }
}

/**
 * Generates a filename for the PDF based on the URL or timestamp.
 * Ensures .pdf extension and handles special characters.
 */
function getFileName(url) {
  try {
    const cleanUrl = url.split('?')[0].split('#')[0];
    const urlParts = cleanUrl.split('/');
    let filename = urlParts[urlParts.length - 1];

    if (!filename || filename.indexOf('.') === -1) {
      return 'pdf_' + new Date().getTime() + '.pdf';
    }

    filename = filename.replace(/[^a-zA-Z0-9._-]/g, '_');

    if (!filename.toLowerCase().endsWith('.pdf')) {
      const lastDotIndex = filename.lastIndexOf('.');
      filename = lastDotIndex > 0
        ? filename.substring(0, lastDotIndex) + '.pdf'
        : filename + '.pdf';
    }

    return filename;

  } catch (error) {
    return 'pdf_' + new Date().getTime() + '.pdf';
  }
}

// ============================================================================
// TIME MONITORING FUNCTIONS
// ============================================================================

/**
 * Checks if the script should continue processing based on elapsed time.
 */
function shouldContinueProcessing(startTime) {
  return (Date.now() - startTime) < MAX_EXECUTION_TIME_MS;
}

/**
 * Creates a time-based trigger to continue processing after a delay.
 */
function createContinuationTrigger() {
  try {
    deleteContinuationTriggers();

    const trigger = ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .timeBased()
      .after(CONTINUATION_DELAY_MS)
      .create();

    const properties = PropertiesService.getScriptProperties();
    const existingTriggers = properties.getProperty(CONTINUATION_TRIGGER_PROPERTY_KEY);
    const triggerIds = existingTriggers ? JSON.parse(existingTriggers) : [];
    triggerIds.push(trigger.getUniqueId());
    properties.setProperty(CONTINUATION_TRIGGER_PROPERTY_KEY, JSON.stringify(triggerIds));

    Logger.log('Continuation trigger created. ID: ' + trigger.getUniqueId());

  } catch (error) {
    Logger.log('Failed to create continuation trigger: ' + error.message);
  }
}

/**
 * Deletes all continuation triggers created by this script.
 */
function deleteContinuationTriggers() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const storedTriggers = properties.getProperty(CONTINUATION_TRIGGER_PROPERTY_KEY);

    if (!storedTriggers) return;

    const triggerIds = JSON.parse(storedTriggers);
    const allTriggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;

    for (let i = 0; i < allTriggers.length; i++) {
      const trigger = allTriggers[i];
      if (triggerIds.includes(trigger.getUniqueId())) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }

    properties.deleteProperty(CONTINUATION_TRIGGER_PROPERTY_KEY);

    if (deletedCount > 0) {
      Logger.log('Deleted ' + deletedCount + ' continuation trigger(s)');
    }

  } catch (error) {
    Logger.log('Failed to delete continuation triggers: ' + error.message);
  }
}
