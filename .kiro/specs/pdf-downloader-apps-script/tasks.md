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

- [x] 17. Fix sheet identification for time-based trigger continuation (BUG)
  - Use PropertiesService.getScriptProperties() to persist the active sheet name between executions
  - At the start of downloadPDFs(), get the sheet using getActiveSheet() and store its name in script properties
  - When triggered by a time-based trigger, check if a sheet name exists in script properties
  - If sheet name exists in properties, use spreadsheet.getSheetByName() to get that sheet
  - If no sheet name in properties (first run), use getActiveSheet() and store the name
  - Clear the sheet name from script properties after all rows are processed successfully
  - Add error handling if stored sheet name doesn't exist
  - _Requirements: 1.1, 8.3_

- [x] 18. Separate PDF link and Drive link columns
  - Add new configuration constant PDF_LINK_COLUMN set to 'D' for extracted PDF URLs
  - Add new configuration constant DRIVE_LINK_COLUMN set to 'E' for Drive shareable links
  - Update extractPDFUrl to write extracted PDF URL to PDF_LINK_COLUMN (column D)
  - Update main function to write Drive shareable link to DRIVE_LINK_COLUMN (column E) instead of overwriting column D
  - Update resume logic to check DRIVE_LINK_COLUMN (column E) for existing content to determine if row is processed
  - Preserve PDF link in column D throughout the process
  - _Requirements: 4.1, 4.2, 4.3_

- [x] 19. Add DIRECT_MODE configuration for direct PDF URLs
  - Add new boolean configuration constant DIRECT_MODE (default false)
  - Add comment explaining DIRECT_MODE: when true, URLs in URL_COLUMN are direct PDF links; when false, URLs are wrapper pages with embed tags
  - Update main function to check DIRECT_MODE before processing each row
  - When DIRECT_MODE is true, skip extractPDFUrl() and use URL_COLUMN value directly as the PDF URL
  - When DIRECT_MODE is false, call extractPDFUrl() to extract PDF URL from embed tag
  - When DIRECT_MODE is true, optionally skip writing to PDF_LINK_COLUMN or copy URL_COLUMN value to PDF_LINK_COLUMN
  - Ensure download logic works the same regardless of mode
  - _Requirements: 1.1, 2.1_

- [x] 20. Implement automatic direct PDF detection
  - Create isDirectPDFUrl helper function that checks if a URL points directly to a PDF
  - Perform a HEAD request using UrlFetchApp.fetch with method: 'head' and muteHttpExceptions: true
  - Check the Content-Type header for 'application/pdf'
  - Check if URL ends with .pdf extension as a fallback
  - Update main function to call isDirectPDFUrl for each URL before processing
  - If isDirectPDFUrl returns true, use the URL directly as the PDF URL (direct mode behavior)
  - If isDirectPDFUrl returns false, call extractPDFUrl to extract from embed tag (standard mode behavior)
  - Remove dependency on DIRECT_MODE constant or make it an override option
  - _Requirements: 1.1, 2.1_

- [x] 21. Code cleanup and refactoring
  - Review code for DRY (Don't Repeat Yourself) violations and extract repeated logic into helper functions
  - Identify and remove dead code or unused variables
  - Look for overly complex conditional logic and simplify where possible
  - Consolidate duplicate error handling patterns
  - Simplify nested conditionals using early returns or guard clauses
  - Remove unnecessary comments that don't add value
  - Verify all comments are accurate and reflect the actual behavior of the code they describe
  - Update or remove outdated comments that no longer match the implementation
  - Ensure consistent naming conventions throughout
  - Apply KISS (Keep It Simple, Stupid) principle to reduce complexity
  - Verify all functions have a single, clear responsibility
  - _Requirements: 7.4_

- [x] 22. Update README documentation
  - Create or update README.md file with current project overview
  - Document all configuration constants and their purpose
  - Explain the automatic direct PDF detection feature
  - Document the incremental processing and resume capability
  - Explain the automatic continuation trigger system for long-running jobs
  - Provide setup instructions including how to get Shared Drive folder ID
  - Document column structure (URL_COLUMN, PDF_LINK_COLUMN, DRIVE_LINK_COLUMN)
  - Include usage instructions for running from custom menu
  - Add troubleshooting section for common issues
  - Document required permissions and first-run authorization
  - _Requirements: 5.1, 5.2, 5.3, 5.4, 6.1, 6.2_
