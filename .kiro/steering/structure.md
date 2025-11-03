# Project Structure

## File Organization

```
/
├── Code.gs                    # Main script file (single-file architecture)
└── .kiro/
    ├── specs/                 # Feature specifications
    │   └── pdf-downloader-apps-script/
    │       ├── requirements.md
    │       ├── design.md
    │       └── tasks.md
    └── steering/              # AI assistant guidance documents
```

## Code Structure (Code.gs)

The script follows a top-down organization:

1. **File header** - JSDoc comment describing the script
2. **Configuration constants** - User-configurable settings (SHARED_DRIVE_FOLDER_ID, URL_COLUMN, LINK_COLUMN, START_ROW)
3. **Menu setup** - `onOpen()` function for custom spreadsheet menu
4. **Main function** - `downloadPDFs()` orchestrates the entire workflow
5. **Helper functions** - Modular functions for specific tasks:
   - `columnLetterToIndex()` - column letter to index conversion
   - `extractPDFUrl()` - HTML parsing to extract embedded PDF URLs
   - `downloadPDF()` - URL fetching and blob retrieval
   - `savePDFToSharedDrive()` - Drive file creation
   - `getFileName()` - filename generation from URLs

## Conventions

- Single-file architecture (typical for Apps Script projects)
- Constants in UPPER_SNAKE_CASE
- Functions in camelCase
- Comprehensive JSDoc comments for all functions
- Section dividers using comment blocks with `=` characters
- Error handling with try-catch and null returns
- Detailed logging for debugging and monitoring
