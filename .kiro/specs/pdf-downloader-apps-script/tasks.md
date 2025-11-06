# Implementation Plan

- [x] 1. Create the Apps Script file with configuration constants
  - Create Code.gs file with SHARED_DRIVE_FOLDER_ID, URL_COLUMN, LINK_COLUMN, and START_ROW constants
  - Add comments explaining each configuration option
  - _Requirements: 5.1, 5.2, 5.3, 5.4_

- [x] 2. Implement main downloadPDFs function
  - Get active spreadsheet and sheet
  - Read URL column range starting from START_ROW
  - Loop through each row and process non-empty URLs
  - _Requirements: 1.1, 1.2, 1.4, 1.5_

- [x] 3. Implement PDF download functionality
  - Create downloadPDF helper function that uses UrlFetchApp.fetch()
  - Return blob from successful download
  - Handle network errors with try-catch
  - _Requirements: 2.1, 2.3, 2.4, 2.5_

- [x] 4. Implement file saving to Shared Drive
  - Create savePDFToSharedDrive helper function
  - Get folder using DriveApp.getFolderById()
  - Generate filename using getFileName helper function
  - Create file in folder using folder.createFile()
  - Return File object
  - _Requirements: 3.1, 3.2, 3.3, 3.5_

- [x] 5. Implement filename generation
  - Create getFileName helper function
  - Extract filename from URL or use timestamp
  - Ensure .pdf extension
  - Handle special characters
  - _Requirements: 3.2_

- [x] 6. Implement spreadsheet link updates
  - Get shareable link from saved file using file.getUrl()
  - Write link to LINK_COLUMN in the same row
  - Write error message to LINK_COLUMN if download or save fails
  - _Requirements: 4.1, 4.2, 4.3, 4.5_

- [x] 7. Extract actual PDF URL from embed wrapper pages
  - Create extractPDFUrl helper function that fetches the page HTML
  - Parse HTML to find the embed tag with type="application/pdf"
  - Extract the src attribute value from the embed tag
  - Handle protocol-relative URLs (starting with //) by adding https:
  - Write the extracted PDF URL to column D
  - Return the extracted PDF URL for downloading
  - _Requirements: 1.1, 2.1_

- [x] 8. Update main function to use extracted PDF URLs
  - Modify downloadPDFs to call extractPDFUrl for each page URL
  - Save extracted PDF URL to column D
  - Use the extracted PDF URL for the download operation
  - Keep existing download and save logic
  - _Requirements: 1.5, 2.1, 4.1_

- [x] 9. Add error handling and logging
  - Wrap download and save operations in try-catch blocks
  - Log errors to console with row number and URL
  - Continue processing remaining URLs after errors
  - _Requirements: 2.3, 2.4, 7.4_

- [x] 10. Create custom menu for running the script
  - Implement onOpen() function that runs when spreadsheet opens
  - Add custom menu to spreadsheet UI using SpreadsheetApp.getUi().createMenu()
  - Add menu item that triggers downloadPDFs() function
  - _Requirements: 6.1, 6.2, 6.3_

- [x] 11. Implement incremental processing with resume capability
  - Modify main function to write results to spreadsheet immediately after each row completes
  - Check if LINK_COLUMN already has content before processing each row
  - Skip rows that already have a link or error message in LINK_COLUMN
  - Allow script to resume from the first unprocessed row if execution halts
  - Use SpreadsheetApp.flush() after each row update to ensure data is written
  - _Requirements: 1.4, 1.5, 4.1_

- [x] 12. Add execution time monitoring configuration
  - Add MAX_EXECUTION_TIME_MS constant (default 5 minutes = 300000 ms)
  - Add CONTINUATION_DELAY_MS constant (default 10 seconds = 10000 ms)
  - Add TRIGGER_FUNCTION_NAME constant set to 'downloadPDFs'
  - Add TRIGGER_UNIQUE_NAME constant for identifying continuation triggers
  - Add comments explaining the time buffer strategy and continuation delay
  - _Requirements: 8.7_

- [x] 13. Implement time monitoring helper function
  - Create shouldContinueProcessing(startTime) function
  - Calculate elapsed time using Date.now() - startTime
  - Return false if elapsed time exceeds MAX_EXECUTION_TIME_MS
  - Return true if still within time limit
  - _Requirements: 8.1, 8.2_

- [x] 14. Implement continuation trigger management
  - Create createContinuationTrigger() function
  - Delete existing triggers with TRIGGER_UNIQUE_NAME using deleteContinuationTriggers()
  - Create new time-based trigger using ScriptApp.newTrigger()
  - Set trigger to fire after CONTINUATION_DELAY_MS (10 seconds) using .timeBased().after(CONTINUATION_DELAY_MS)
  - Set unique name on trigger for identification
  - Log trigger creation event
  - _Requirements: 8.3, 8.4, 8.8_

- [x] 15. Implement trigger cleanup function
  - Create deleteContinuationTriggers() function
  - Get all project triggers using ScriptApp.getProjectTriggers()
  - Loop through triggers and identify those matching TRIGGER_UNIQUE_NAME
  - Delete matching triggers using ScriptApp.deleteTrigger()
  - Log trigger deletion events
  - _Requirements: 8.6, 8.8_

- [x] 16. Integrate time monitoring into main function
  - Record start time at beginning of downloadPDFs() using Date.now()
  - Before processing each row, call shouldContinueProcessing(startTime)
  - If shouldContinueProcessing returns false, call createContinuationTrigger() and exit loop
  - After processing all rows successfully, call deleteContinuationTriggers()
  - Log continuation events when time limit is reached
  - _Requirements: 8.1, 8.2, 8.3, 8.5, 8.6, 8.8_
