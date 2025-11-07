# Requirements Document

## Introduction

This feature provides a Google Apps Script solution that automates the process of downloading PDF files from URLs listed in a Google Spreadsheet. The script reads URLs from a specified column, downloads each PDF, saves it to a designated folder in Google Shared Drive, and updates the adjacent column with a link to the saved PDF file. This sheet-bound script can be executed directly from the spreadsheet interface.

## Glossary

- **PDF Downloader Script**: The Google Apps Script that performs URL reading, PDF downloading, file saving, and link updating operations
- **Source Column**: The spreadsheet column containing URLs pointing to PDF files
- **Target Column**: The spreadsheet column where links to saved PDFs will be written
- **Shared Drive Folder**: The designated folder in Google Shared Drive where downloaded PDFs will be stored
- **Active Sheet**: The currently selected sheet in the Google Spreadsheet where the script operates

## Requirements

### Requirement 1

**User Story:** As a spreadsheet user, I want the script to read URLs from a column in my sheet, so that I can process multiple PDF downloads in batch

#### Acceptance Criteria

1. WHEN the PDF Downloader Script executes, THE PDF Downloader Script SHALL identify the Source Column containing PDF URLs
2. THE PDF Downloader Script SHALL read all non-empty cells in the Source Column starting from row 2
3. THE PDF Downloader Script SHALL validate that each cell contains a properly formatted URL
4. THE PDF Downloader Script SHALL skip rows where the Source Column is empty
5. THE PDF Downloader Script SHALL process URLs sequentially from top to bottom

### Requirement 2

**User Story:** As a spreadsheet user, I want the script to download PDFs from the URLs, so that I can store them in my organization's Shared Drive

#### Acceptance Criteria

1. WHEN the PDF Downloader Script processes a URL, THE PDF Downloader Script SHALL fetch the content from the URL
2. THE PDF Downloader Script SHALL verify that the fetched content is a PDF file
3. IF a URL fetch fails, THEN THE PDF Downloader Script SHALL log an error message with the row number and URL
4. IF a URL fetch fails, THEN THE PDF Downloader Script SHALL continue processing remaining URLs
5. THE PDF Downloader Script SHALL handle network timeouts with a maximum wait time of 30 seconds per download

### Requirement 3

**User Story:** As a spreadsheet user, I want downloaded PDFs saved to a specific Shared Drive folder, so that my team can access them in a centralized location

#### Acceptance Criteria

1. THE PDF Downloader Script SHALL save each downloaded PDF to the designated Shared Drive Folder
2. THE PDF Downloader Script SHALL generate a unique filename for each PDF based on the original URL or timestamp
3. IF a file with the same name exists in the Shared Drive Folder, THEN THE PDF Downloader Script SHALL append a numeric suffix to create a unique filename
4. THE PDF Downloader Script SHALL set appropriate sharing permissions on saved files to match the Shared Drive Folder permissions
5. WHEN a PDF is successfully saved, THE PDF Downloader Script SHALL retrieve the file's shareable link

### Requirement 4

**User Story:** As a spreadsheet user, I want the script to add links to saved PDFs in the adjacent column, so that I can easily access the downloaded files

#### Acceptance Criteria

1. WHEN a PDF is successfully saved to the Shared Drive Folder, THE PDF Downloader Script SHALL write the shareable link to the Target Column
2. THE Target Column SHALL be the column immediately adjacent to the Source Column
3. THE PDF Downloader Script SHALL write the link in the same row as the source URL
4. THE PDF Downloader Script SHALL format the cell as a hyperlink with descriptive text
5. IF a download fails for a specific URL, THEN THE PDF Downloader Script SHALL write an error message in the Target Column for that row

### Requirement 5

**User Story:** As a spreadsheet user, I want to configure the script settings easily, so that I can specify which column contains URLs and where to save files

#### Acceptance Criteria

1. THE PDF Downloader Script SHALL provide a configuration section at the top of the script file with clearly labeled constants
2. THE PDF Downloader Script SHALL allow configuration of the Source Column letter or index
3. THE PDF Downloader Script SHALL allow configuration of the Shared Drive Folder ID
4. THE PDF Downloader Script SHALL allow configuration of the sheet name to process
5. THE PDF Downloader Script SHALL validate configuration values before processing begins

### Requirement 6

**User Story:** As a spreadsheet user, I want to run the script from the spreadsheet menu, so that I can execute it without opening the script editor

#### Acceptance Criteria

1. WHEN the spreadsheet opens, THE PDF Downloader Script SHALL add a custom menu to the spreadsheet menu bar
2. THE custom menu SHALL contain an option labeled "Download PDFs"
3. WHEN the user selects the "Download PDFs" menu option, THE PDF Downloader Script SHALL execute the download process
4. WHILE the PDF Downloader Script is executing, THE PDF Downloader Script SHALL display a progress indicator or toast notification
5. WHEN the PDF Downloader Script completes, THE PDF Downloader Script SHALL display a summary message with the count of successful and failed downloads

### Requirement 7

**User Story:** As a spreadsheet user, I want clear error messages when something goes wrong, so that I can troubleshoot issues effectively

#### Acceptance Criteria

1. IF the Shared Drive Folder ID is invalid or inaccessible, THEN THE PDF Downloader Script SHALL display an error message and halt execution
2. IF the specified sheet name does not exist, THEN THE PDF Downloader Script SHALL display an error message and halt execution
3. IF the Source Column is empty or invalid, THEN THE PDF Downloader Script SHALL display an error message and halt execution
4. WHEN an error occurs during processing, THE PDF Downloader Script SHALL log the error details to the Apps Script execution log
5. THE PDF Downloader Script SHALL provide user-friendly error messages in the spreadsheet UI using toast notifications or alerts

### Requirement 8

**User Story:** As a spreadsheet user processing large datasets, I want the script to automatically handle Google Apps Script execution time limits, so that all my PDFs are downloaded without manual intervention

#### Acceptance Criteria

1. THE PDF Downloader Script SHALL monitor elapsed execution time during processing
2. WHEN elapsed execution time reaches 5 minutes, THE PDF Downloader Script SHALL stop processing new rows
3. WHEN the PDF Downloader Script stops due to time limits, THE PDF Downloader Script SHALL create a time-based trigger to continue processing
4. THE time-based trigger SHALL execute the PDF Downloader Script within 10 seconds of the previous execution stopping
5. WHEN the PDF Downloader Script resumes via trigger, THE PDF Downloader Script SHALL continue from the first unprocessed row
6. WHEN all rows are processed, THE PDF Downloader Script SHALL delete any remaining continuation triggers
7. THE PDF Downloader Script SHALL allow configuration of the maximum execution time threshold and continuation delay
8. THE PDF Downloader Script SHALL log time-based continuation events to the execution log
9. THE PDF Downloader Script SHALL persist the active sheet name using PropertiesService between trigger-based executions
10. WHEN triggered by a time-based trigger, THE PDF Downloader Script SHALL retrieve the sheet name from script properties and process that specific sheet

### Requirement 9

**User Story:** As a spreadsheet user, I want to see both the extracted PDF URL and the final Drive link in separate columns, so that I can verify the extraction worked correctly and access the saved file

#### Acceptance Criteria

1. THE PDF Downloader Script SHALL write extracted PDF URLs to a dedicated PDF Link Column
2. THE PDF Downloader Script SHALL write Drive shareable links to a separate Drive Link Column
3. THE PDF Downloader Script SHALL preserve the extracted PDF URL when writing the Drive link
4. THE PDF Downloader Script SHALL allow configuration of both PDF Link Column and Drive Link Column separately
5. WHEN resuming processing, THE PDF Downloader Script SHALL check the Drive Link Column to determine if a row has been processed

### Requirement 10

**User Story:** As a spreadsheet user with direct PDF URLs, I want the script to handle both wrapper pages and direct PDF links, so that I don't need to manually extract URLs first

#### Acceptance Criteria

1. THE PDF Downloader Script SHALL provide a DIRECT_MODE configuration option
2. WHEN DIRECT_MODE is true, THE PDF Downloader Script SHALL use URLs from the Source Column directly as PDF URLs
3. WHEN DIRECT_MODE is false, THE PDF Downloader Script SHALL extract PDF URLs from embed tags in wrapper pages
4. THE PDF Downloader Script SHALL allow users to toggle between direct and wrapper page modes via configuration
5. THE PDF Downloader Script SHALL document the DIRECT_MODE setting with clear comments explaining both modes

### Requirement 11

**User Story:** As a spreadsheet user with mixed URL types, I want the script to automatically detect whether a URL is a direct PDF or wrapper page, so that I don't need to configure modes or separate my data

#### Acceptance Criteria

1. THE PDF Downloader Script SHALL automatically detect if a URL points directly to a PDF file
2. THE PDF Downloader Script SHALL perform a HEAD request to check the Content-Type header for PDF detection
3. IF the Content-Type header indicates application/pdf, THEN THE PDF Downloader Script SHALL treat the URL as a direct PDF link
4. IF the URL ends with .pdf extension, THEN THE PDF Downloader Script SHALL treat the URL as a direct PDF link
5. IF automatic detection indicates a direct PDF, THEN THE PDF Downloader Script SHALL skip the embed tag extraction step
6. IF automatic detection indicates a wrapper page, THEN THE PDF Downloader Script SHALL extract the PDF URL from embed tags
7. THE PDF Downloader Script SHALL handle HEAD request failures gracefully and fall back to embed tag extraction
