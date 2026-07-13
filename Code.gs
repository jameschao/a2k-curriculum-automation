// ============================================================================
// CONFIGURATION - Config Sheet
// ============================================================================
// The "config" sheet contains parent folder URLs for each lesson sheet.
// Structure (tab-separated):
//   sheet | workspace_root | versions_root | published_root
//   1st   | https://...    | https://...   | https://...
//
// Each row after the header corresponds to a sheet name and its folder URLs.

// ============================================================================
// CONFIGURATION - Column Headers
// ============================================================================

const COLUMNS = {
  ID: 'id',
  STATUS: 'status',
  WORKSPACE_FOLDER_URL: 'workspace_folder_url',
  LATEST_VERSION_FOLDER_URL: 'latest_version_folder_url',
  PUBLISH_FOLDER_URL: 'publish_folder_url',
  LAST_PUBLISH_TIME: 'last_publish_time'
};

// ============================================================================
// CONFIGURATION - Status Values
// ============================================================================

const STATUS_VALUES = {
  NOT_STARTED: 'Not Started',
  IN_DEVELOPMENT: 'In Development',
  IN_REVIEW: 'In Review',
  READY_TO_PUBLISH: 'Ready to Publish',
  PUBLISHING_IN_PROGRESS: 'Publishing in Progress',
  PUBLISHED: 'Published',
  BLOCKED: 'Blocked'
};

// ============================================================================
// MENU INITIALIZATION
// ============================================================================

/**
 * The onOpen function runs automatically every time the spreadsheet loads.
 * It builds the custom menu structure inside the Google Sheets UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Instantiates the top-level main menu container
  ui.createMenu('🚀 A2K Publish')
    .addItem('Publish lesson', 'publishLesson') // (Display Label, Target Function Name)

    // Adds a visual dividing line to group operational commands
    .addSeparator()

    // Nests a sub-menu for validation
    .addSubMenu(ui.createMenu('Validation')
      .addItem('Validate Active Sheet', 'validateActiveSheetStructure')
      .addItem('Validate Config Sheet', 'validateConfigSheetStructure')
    )

    // Renders the built structure into the spreadsheet main header bar
    .addToUi();
}

// ============================================================================
// ENTRY POINT FUNCTIONS TIED TO MENU ACTIONS
// ============================================================================

/**
 * Main automation workflow for publishing a lesson.
 *
 * Steps:
 * 1. Validates sheet structure (column headers and status dropdown values)
 * 2. Validates lesson (ID from selected row and status must be "Ready to Publish")
 * 3. Sets status to "Publishing in Progress" and creates timestamped version folder with PDFs
 * 4. Publishes files to publish folder
 * 5. Updates spreadsheet row with folder URLs, timestamp, and sets status to "Published"
 *
 * User must have a cell selected in the lesson's row before running this function.
 */
function publishLesson() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('=== Starting publishLesson ===');

  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    Logger.log(`Active sheet: "${sheetName}"`);

    // Step 1: Validate config sheet and active sheet structure
    ss.toast('Validating configuration and sheet structure...', '🔍 Step 1/4', 3);
    Logger.log('Step 1: Validating configuration and sheet structure...');

    // Step 1a: Validate config sheet
    Logger.log('  1a: Validating config sheet...');
    const configValidation = performConfigSheetValidation();

    if (!configValidation.isValid) {
      Logger.log('  FAILED: Config sheet validation failed');
      const message = formatConfigValidationMessage(configValidation);
      ui.alert(
        'Config Sheet Validation Failed',
        'The config sheet is invalid. Please fix the following issues before publishing:\n\n' + message,
        ui.ButtonSet.OK
      );
      return;
    }

    Logger.log('  SUCCESS: Config sheet is valid');

    // Get configuration from config sheet
    Logger.log('  Loading configuration from config sheet...');
    const config = getSheetConfig(sheetName);
    Logger.log('  SUCCESS: Configuration loaded');

    // Parse folder URLs into context object
    Logger.log('  Parsing parent folder URLs...');
    const context = {
      workspacesRootId: parseFolderId(config.workspace_root),
      versionsRootId: parseFolderId(config.versions_root),
      publishedRootId: parseFolderId(config.published_root)
    };
    Logger.log('  SUCCESS: All parent folder URLs parsed');

    // Step 1b: Validate active sheet structure
    Logger.log('  1b: Validating active sheet structure...');
    const structureValidation = performSheetValidation(sheet);

    if (!structureValidation.isValid) {
      Logger.log('  FAILED: Sheet structure validation failed');
      const message = formatValidationMessage(structureValidation);
      ui.alert(
        'Sheet Structure Validation Failed',
        'The sheet structure is invalid. Please fix the following issues before publishing:\n\n' + message,
        ui.ButtonSet.OK
      );
      return;
    }

    Logger.log('  SUCCESS: Sheet structure is valid');
    Logger.log('SUCCESS: All validation checks passed');

    // Step 2: Validate lesson (ID and status)
    ss.toast('Validating lesson...', '📋 Step 2/4', 3);
    Logger.log('Step 2: Validating lesson ID and status...');

    // Get and validate the selected lesson name
    Logger.log('  - Validating lesson ID selection...');
    const lessonId = getValidatedLessonName();

    if (!lessonId) {
      Logger.log('FAILED: No valid lesson name selected');
      ui.alert(
        'Invalid Selection',
        'Please select a valid lesson row before running this automation.',
        ui.ButtonSet.OK
      );
      return;
    }

    Logger.log(`  - SUCCESS: Validated lesson ID: "${lessonId}"`);

    // Validate lesson status
    Logger.log('  - Validating lesson status...');
    const columnIndices = getColumnIndices(sheet);
    const lessonRow = findLessonRow(sheet, columnIndices, lessonId);

    if (lessonRow === -1) {
      Logger.log('  - FAILED: Could not find lesson row');
      ui.alert(
        'Error',
        `Could not find lesson "${lessonId}" in the sheet.`,
        ui.ButtonSet.OK
      );
      return;
    }

    const statusColumnIndex = columnIndices[COLUMNS.STATUS];
    if (!statusColumnIndex) {
      Logger.log('  - FAILED: Status column not found');
      ui.alert(
        'Error',
        'Status column not found in the sheet.',
        ui.ButtonSet.OK
      );
      return;
    }

    const currentStatus = sheet.getRange(lessonRow, statusColumnIndex).getValue();
    Logger.log(`  - Current status: "${currentStatus}"`);

    if (currentStatus !== STATUS_VALUES.READY_TO_PUBLISH) {
      Logger.log(`  - FAILED: Lesson status is "${currentStatus}", not "Ready to Publish"`);
      ui.alert(
        'Cannot Publish',
        `Lesson "${lessonId}" has status "${currentStatus}".\n\n` +
        `Only lessons with status "${STATUS_VALUES.READY_TO_PUBLISH}" can be published.`,
        ui.ButtonSet.OK
      );
      return;
    }

    Logger.log('  - SUCCESS: Lesson status is "Ready to Publish"');
    Logger.log('SUCCESS: Lesson validation complete');

    // Update status to "Publishing in Progress"
    Logger.log('Updating status to "Publishing in Progress"...');
    sheet.getRange(lessonRow, statusColumnIndex).setValue(STATUS_VALUES.PUBLISHING_IN_PROGRESS);
    SpreadsheetApp.flush(); // Force the update to be visible immediately
    Logger.log('SUCCESS: Status updated to "Publishing in Progress"');

    // Step 3: Create timestamped version folder with PDFs
    ss.toast('Creating version folder and generating PDFs...', '📁 Step 3/4', 5);
    Logger.log('Step 3: Creating version folder and generating PDFs...');
    const {versionFolderId, workspaceFolderId} = createVersionFolder(lessonId, context);
    Logger.log(`SUCCESS: Version folder created with ID: ${versionFolderId}`);

    // Step 4: Publish files to publish folder
    ss.toast('Publishing files to publish folder...', '🚀 Step 4/4', 5);
    Logger.log('Step 4: Publishing files to publish folder...');
    const publishFolderId = publishVersionFiles(lessonId, versionFolderId, context);
    Logger.log('SUCCESS: Files published successfully');

    // Update spreadsheet row with folder URLs, timestamp, and status
    Logger.log('Updating spreadsheet row...');
    updateSpreadsheetRow({row: lessonRow, lessonId: lessonId}, columnIndices, workspaceFolderId, versionFolderId, publishFolderId);
    Logger.log('SUCCESS: Spreadsheet row updated');

    Logger.log(`=== Automation completed successfully for lesson: ${lessonId} ===`);

    ss.toast('Publishing completed successfully! ✅', 'Success', 5);
    ui.alert(
      'Success',
      `Automation completed for lesson: ${lessonId}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`FATAL ERROR in publishLesson: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    ui.alert(
      'Error',
      `Automation failed: ${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Validates the sheet structure and returns validation results.
 * Checks for required column headers and status dropdown values.
 *
 * @param {Sheet} sheet - The sheet to validate
 * @returns {Object} Validation results object containing:
 *   - isValid: boolean indicating if validation passed
 *   - missingColumns: array of missing column names
 *   - additionalColumns: array of extra column names (informational only)
 *   - missingStatusValues: array of missing status dropdown values
 *   - unexpectedStatusValues: array of unexpected status dropdown values
 */
function performSheetValidation(sheet) {
  Logger.log('=== Starting performSheetValidation ===');

  const validationResults = {
    isValid: true,
    missingColumns: [],
    additionalColumns: [],
    missingStatusValues: [],
    unexpectedStatusValues: []
  };

  // Step 1: Validate column headers
  Logger.log('Step 1: Validating column headers...');
  const columnIndices = getColumnIndices(sheet);
  const expectedColumns = Object.values(COLUMNS);
  const actualColumns = Object.keys(columnIndices);

  // Check for missing columns (FAIL validation)
  expectedColumns.forEach(expectedCol => {
    if (!columnIndices[expectedCol]) {
      validationResults.missingColumns.push(expectedCol);
      validationResults.isValid = false;
    }
  });

  // Check for additional columns (acceptable, just informational)
  actualColumns.forEach(actualCol => {
    if (!expectedColumns.includes(actualCol)) {
      validationResults.additionalColumns.push(actualCol);
    }
  });

  Logger.log(`  - Missing columns: ${validationResults.missingColumns.length}`);
  Logger.log(`  - Additional columns: ${validationResults.additionalColumns.length}`);

  // Step 2: Validate status column dropdown values
  Logger.log('Step 2: Validating status column dropdown values...');
  const statusColumnIndex = columnIndices[COLUMNS.STATUS];

  if (statusColumnIndex) {
    const expectedStatusValues = Object.values(STATUS_VALUES);
    const actualStatusValues = getDropdownValues(sheet, statusColumnIndex);

    if (actualStatusValues) {
      // Check for missing status values
      expectedStatusValues.forEach(expectedStatus => {
        if (!actualStatusValues.includes(expectedStatus)) {
          validationResults.missingStatusValues.push(expectedStatus);
          validationResults.isValid = false;
        }
      });

      // Check for unexpected status values
      actualStatusValues.forEach(actualStatus => {
        if (!expectedStatusValues.includes(actualStatus)) {
          validationResults.unexpectedStatusValues.push(actualStatus);
          validationResults.isValid = false;
        }
      });

      Logger.log(`  - Missing status values: ${validationResults.missingStatusValues.length}`);
      Logger.log(`  - Unexpected status values: ${validationResults.unexpectedStatusValues.length}`);
    } else {
      Logger.log('  - WARNING: No data validation found on status column');
      validationResults.isValid = false;
      validationResults.missingStatusValues = expectedStatusValues;
    }
  } else {
    Logger.log('  - WARNING: Status column not found, skipping status validation');
  }

  Logger.log(`=== Validation completed: ${validationResults.isValid ? 'PASSED' : 'FAILED'} ===`);
  return validationResults;
}

/**
 * Formats validation results into a user-friendly message for display to the user.
 *
 * @param {Object} validationResults - The validation results object from performSheetValidation()
 * @returns {string} Formatted message with validation status and any issues found
 */
function formatValidationMessage(validationResults) {
  let message = '';

  if (validationResults.isValid) {
    message = '✅ Validation Passed\n\nAll required column headers and status values are correct.';

    // Mention additional columns if present
    if (validationResults.additionalColumns.length > 0) {
      message += '\n\nAdditional columns found (acceptable):\n  • ' +
                 validationResults.additionalColumns.join('\n  • ');
    }
  } else {
    message = '❌ Validation Failed\n\n';
    const issues = [];

    if (validationResults.missingColumns.length > 0) {
      issues.push('Missing Columns:\n  • ' + validationResults.missingColumns.join('\n  • '));
    }

    if (validationResults.missingStatusValues.length > 0) {
      issues.push('Missing Status Values:\n  • ' + validationResults.missingStatusValues.join('\n  • '));
    }

    if (validationResults.unexpectedStatusValues.length > 0) {
      issues.push('Unexpected Status Values:\n  • ' + validationResults.unexpectedStatusValues.join('\n  • '));
    }

    message += issues.join('\n\n');

    // Mention additional columns if present (at the end, as informational)
    if (validationResults.additionalColumns.length > 0) {
      message += '\n\nAdditional columns found (acceptable):\n  • ' +
                 validationResults.additionalColumns.join('\n  • ');
    }
  }

  return message;
}

/**
 * Validates the sheet structure and shows a UI alert with the results.
 * This is a wrapper around performSheetValidation() that displays results to the user.
 *
 * @param {Sheet} sheet - The sheet to validate
 */
function validateSheetStructure(sheet) {
  const ui = SpreadsheetApp.getUi();

  try {
    const validationResults = performSheetValidation(sheet);
    const message = formatValidationMessage(validationResults);

    ui.alert(
      validationResults.isValid ? 'Validation Passed' : 'Validation Failed',
      message,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`FATAL ERROR in validateSheetStructure: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    ui.alert(
      'Error',
      `Validation failed with error: ${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Entry point for the "Validate Active Sheet" menu item.
 * Gets the currently active sheet and validates its structure.
 */
function validateActiveSheetStructure() {
  const sheet = SpreadsheetApp.getActiveSheet();
  validateSheetStructure(sheet);
}

/**
 * Validates the config sheet structure and returns validation results.
 * Checks for:
 * - Config sheet exists
 * - Required column headers are present
 * - Active sheet has a configuration row
 * - All folder URL values are present and valid Google Drive URLs
 *
 * @returns {Object} Validation results object containing:
 *   - isValid: boolean indicating if validation passed
 *   - configSheetExists: boolean indicating if config sheet exists
 *   - missingColumns: array of missing column names
 *   - additionalColumns: array of extra column names (informational only)
 *   - activeSheetHasConfig: boolean indicating if active sheet has a config row
 *   - activeSheetName: name of the active sheet
 *   - invalidUrls: array of objects describing invalid URLs (column name and value)
 *   - missingUrls: array of column names with missing/empty URLs
 */
function performConfigSheetValidation() {
  Logger.log('=== Starting performConfigSheetValidation ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeSheetName = activeSheet.getName();

  const validationResults = {
    isValid: true,
    configSheetExists: false,
    missingColumns: [],
    additionalColumns: [],
    activeSheetHasConfig: false,
    activeSheetName: activeSheetName,
    invalidUrls: [],
    missingUrls: []
  };

  // Don't validate the config sheet itself
  if (activeSheetName === 'config') {
    Logger.log('Active sheet is the config sheet itself - skipping validation');
    validationResults.isValid = false;
    validationResults.isConfigSheet = true;
    return validationResults;
  }

  // Step 1: Check if config sheet exists
  Logger.log('Step 1: Checking if config sheet exists...');
  const configSheet = ss.getSheetByName('config');

  if (!configSheet) {
    Logger.log('FAILED: Config sheet does not exist');
    validationResults.isValid = false;
    validationResults.configSheetExists = false;
    return validationResults;
  }

  validationResults.configSheetExists = true;
  Logger.log('SUCCESS: Config sheet exists');

  // Step 2: Validate column headers
  Logger.log('Step 2: Validating column headers...');
  const expectedColumns = ['sheet', 'workspace_root', 'versions_root', 'published_root'];
  const lastRow = configSheet.getLastRow();
  const lastCol = configSheet.getLastColumn();

  if (lastRow < 1) {
    Logger.log('FAILED: Config sheet is empty');
    validationResults.isValid = false;
    validationResults.missingColumns = expectedColumns;
    return validationResults;
  }

  const headers = configSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalizedHeaders = headers.map(h => String(h).toLowerCase().trim());

  // Check for missing columns
  expectedColumns.forEach(expectedCol => {
    if (!normalizedHeaders.includes(expectedCol)) {
      validationResults.missingColumns.push(expectedCol);
      validationResults.isValid = false;
    }
  });

  // Check for additional columns (informational only)
  normalizedHeaders.forEach(actualCol => {
    if (actualCol && !expectedColumns.includes(actualCol)) {
      validationResults.additionalColumns.push(actualCol);
    }
  });

  Logger.log(`  - Missing columns: ${validationResults.missingColumns.length}`);
  Logger.log(`  - Additional columns: ${validationResults.additionalColumns.length}`);

  // If missing columns, stop here
  if (validationResults.missingColumns.length > 0) {
    return validationResults;
  }

  // Get column indices
  const sheetColIndex = normalizedHeaders.indexOf('sheet');
  const workspaceRootIndex = normalizedHeaders.indexOf('workspace_root');
  const versionsRootIndex = normalizedHeaders.indexOf('versions_root');
  const publishedRootIndex = normalizedHeaders.indexOf('published_root');

  // Step 3: Check if active sheet has a configuration row
  Logger.log(`Step 3: Checking if active sheet "${activeSheetName}" has a config row...`);

  if (lastRow < 2) {
    Logger.log('FAILED: Config sheet has no data rows');
    validationResults.isValid = false;
    validationResults.activeSheetHasConfig = false;
    return validationResults;
  }

  const data = configSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let configRow = null;

  for (let i = 0; i < data.length; i++) {
    const rowSheetName = String(data[i][sheetColIndex]).trim();
    if (rowSheetName === activeSheetName) {
      configRow = data[i];
      Logger.log(`SUCCESS: Found config row for sheet "${activeSheetName}"`);
      validationResults.activeSheetHasConfig = true;
      break;
    }
  }

  if (!configRow) {
    Logger.log(`FAILED: No config row found for sheet "${activeSheetName}"`);
    validationResults.isValid = false;
    validationResults.activeSheetHasConfig = false;
    return validationResults;
  }

  // Step 4: Validate that all URL values are present and valid
  Logger.log('Step 4: Validating folder URL values...');

  const urlColumns = [
    { name: 'workspace_root', index: workspaceRootIndex, value: configRow[workspaceRootIndex] },
    { name: 'versions_root', index: versionsRootIndex, value: configRow[versionsRootIndex] },
    { name: 'published_root', index: publishedRootIndex, value: configRow[publishedRootIndex] }
  ];

  urlColumns.forEach(col => {
    const urlValue = String(col.value).trim();

    // Check if URL is missing or empty
    if (!urlValue) {
      Logger.log(`  - FAILED: ${col.name} is empty`);
      validationResults.missingUrls.push(col.name);
      validationResults.isValid = false;
      return;
    }

    // Try to parse the folder ID to validate URL format
    try {
      parseFolderId(urlValue);
      Logger.log(`  - SUCCESS: ${col.name} is a valid URL`);
    } catch (error) {
      Logger.log(`  - FAILED: ${col.name} has invalid URL: ${urlValue}`);
      validationResults.invalidUrls.push({
        column: col.name,
        value: urlValue,
        error: error.message
      });
      validationResults.isValid = false;
    }
  });

  Logger.log(`=== Config validation completed: ${validationResults.isValid ? 'PASSED' : 'FAILED'} ===`);
  return validationResults;
}

/**
 * Formats config validation results into a user-friendly message.
 *
 * @param {Object} validationResults - The validation results from performConfigSheetValidation()
 * @returns {string} Formatted message with validation status and any issues found
 */
function formatConfigValidationMessage(validationResults) {
  let message = '';

  // Special case: validating the config sheet itself
  if (validationResults.isConfigSheet) {
    return 'Cannot validate the config sheet itself.\n\nPlease switch to a lesson sheet to validate its configuration.';
  }

  if (validationResults.isValid) {
    message = `✅ Config Validation Passed\n\nSheet "${validationResults.activeSheetName}" has valid configuration.`;

    if (validationResults.additionalColumns.length > 0) {
      message += '\n\nAdditional columns in config sheet (acceptable):\n  • ' +
                 validationResults.additionalColumns.join('\n  • ');
    }
  } else {
    message = '❌ Config Validation Failed\n\n';
    const issues = [];

    if (!validationResults.configSheetExists) {
      issues.push('Config sheet not found.\nPlease create a sheet named "config" with columns:\n  • sheet\n  • workspace_root\n  • versions_root\n  • published_root');
    } else if (validationResults.missingColumns.length > 0) {
      issues.push('Missing Columns in config sheet:\n  • ' + validationResults.missingColumns.join('\n  • '));
    } else if (!validationResults.activeSheetHasConfig) {
      issues.push(`No configuration row found for sheet "${validationResults.activeSheetName}".\n\nPlease add a row in the config sheet with:\n  • sheet = "${validationResults.activeSheetName}"\n  • workspace_root = <folder URL>\n  • versions_root = <folder URL>\n  • published_root = <folder URL>`);
    } else {
      if (validationResults.missingUrls.length > 0) {
        issues.push('Missing or Empty URLs:\n  • ' + validationResults.missingUrls.join('\n  • '));
      }

      if (validationResults.invalidUrls.length > 0) {
        const invalidUrlDetails = validationResults.invalidUrls.map(item =>
          `${item.column}: ${item.value}\n    Error: ${item.error}`
        );
        issues.push('Invalid Google Drive URLs:\n  • ' + invalidUrlDetails.join('\n  • '));
      }
    }

    message += issues.join('\n\n');

    if (validationResults.additionalColumns.length > 0) {
      message += '\n\nAdditional columns in config sheet (acceptable):\n  • ' +
                 validationResults.additionalColumns.join('\n  • ');
    }
  }

  return message;
}

/**
 * Validates the config sheet and shows a UI alert with the results.
 * Entry point for the "Validate Config Sheet" menu item.
 */
function validateConfigSheetStructure() {
  const ui = SpreadsheetApp.getUi();

  try {
    const validationResults = performConfigSheetValidation();
    const message = formatConfigValidationMessage(validationResults);

    ui.alert(
      validationResults.isValid ? 'Config Validation Passed' : 'Config Validation Failed',
      message,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`FATAL ERROR in validateConfigSheetStructure: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    ui.alert(
      'Error',
      `Config validation failed with error: ${error.message}`,
      ui.ButtonSet.OK
    );
  }
}


// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Reads configuration from the "config" sheet for the specified sheet name.
 * Returns an object with workspace_root, versions_root, and published_root URLs.
 *
 * @param {string} sheetName - The name of the sheet to look up
 * @returns {Object} Configuration object containing:
 *   - workspace_root: URL to the workspace parent folder
 *   - versions_root: URL to the versions parent folder
 *   - published_root: URL to the published parent folder
 * @throws {Error} If the config sheet doesn't exist or the sheet name isn't found
 */
function getSheetConfig(sheetName) {
  Logger.log(`  - Looking up config for sheet: "${sheetName}"`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');

  if (!configSheet) {
    throw new Error('Config sheet not found. Please create a sheet named "config" with columns: sheet, workspace_root, versions_root, published_root');
  }

  // Read all data from config sheet
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('Config sheet is empty. Please add configuration rows.');
  }

  const data = configSheet.getRange(1, 1, lastRow, configSheet.getLastColumn()).getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());

  // Find column indices
  const sheetColIndex = headers.indexOf('sheet');
  const workspaceRootIndex = headers.indexOf('workspace_root');
  const versionsRootIndex = headers.indexOf('versions_root');
  const publishedRootIndex = headers.indexOf('published_root');

  if (sheetColIndex === -1 || workspaceRootIndex === -1 || versionsRootIndex === -1 || publishedRootIndex === -1) {
    throw new Error('Config sheet is missing required columns: sheet, workspace_root, versions_root, published_root');
  }

  // Find the row matching the sheet name
  for (let i = 1; i < data.length; i++) {
    const rowSheetName = String(data[i][sheetColIndex]).trim();
    if (rowSheetName === sheetName) {
      Logger.log(`  - SUCCESS: Found config for sheet "${sheetName}"`);
      return {
        workspace_root: String(data[i][workspaceRootIndex]).trim(),
        versions_root: String(data[i][versionsRootIndex]).trim(),
        published_root: String(data[i][publishedRootIndex]).trim()
      };
    }
  }

  throw new Error(`No configuration found for sheet "${sheetName}" in the config sheet`);
}

/**
 * Parses a Google Drive folder URL and extracts the folder ID.
 *
 * @param {string} url - A Google Drive folder URL
 * @returns {string} The folder ID
 * @throws {Error} If the URL format is invalid
 *
 * Supported formats:
 * - https://drive.google.com/drive/folders/FOLDER_ID
 * - https://drive.google.com/drive/folders/FOLDER_ID?usp=drive_link
 */
function parseFolderId(url) {
  if (!url || typeof url !== 'string') {
    throw new Error('Folder URL is required');
  }

  const trimmed = url.trim();

  // Parse URL and extract folder ID
  // Format: https://drive.google.com/drive/folders/FOLDER_ID
  // or:     https://drive.google.com/drive/folders/FOLDER_ID?usp=drive_link
  const match = trimmed.match(/\/folders\/([a-zA-Z0-9_-]+)/);

  if (!match) {
    throw new Error(`Invalid Google Drive folder URL format: ${trimmed}`);
  }

  return match[1];
}

/**
 * Gets the lesson ID from the "id" column of the currently selected row.
 * Validates that a cell is selected and it's in a data row (not the header).
 *
 * @returns {string|null} The lesson ID if valid, or null if invalid/not selected
 */
function getValidatedLessonName() {
  Logger.log('  - Getting current cell selection...');
  const sheet = SpreadsheetApp.getActiveSheet();
  const currentCell = SpreadsheetApp.getCurrentCell();

  if (!currentCell) {
    Logger.log('  - FAILED: No cell selected');
    return null;
  }

  const selectedRow = currentCell.getRow();
  Logger.log(`  - Selected row: ${selectedRow}`);

  // Row 1 is the header, data starts at row 2
  if (selectedRow < 2) {
    Logger.log('  - FAILED: Selected row is the header row');
    return null;
  }

  // Get column indices to find the "id" column
  const columnIndices = getColumnIndices(sheet);
  const idColumnIndex = columnIndices[COLUMNS.ID];

  if (!idColumnIndex) {
    Logger.log('  - FAILED: "id" column not found in sheet');
    return null;
  }

  // Get the value from the "id" column of the selected row
  const lessonId = sheet.getRange(selectedRow, idColumnIndex).getValue();
  Logger.log(`  - Lesson ID from "id" column: "${lessonId}"`);

  if (!lessonId || typeof lessonId !== 'string' || lessonId.trim() === '') {
    Logger.log('  - FAILED: Lesson ID is empty or invalid');
    return null;
  }

  return lessonId.trim();
}

/**
 * Creates a timestamped version folder and generates PDFs for all Google Docs.
 * Copies all files from the workspace folder to the version folder and creates
 * a VERSION document with the timestamp in its name.
 *
 * @param {string} lessonId - The lesson identifier (used to find the workspace folder)
 * @param {Object} context - Context object containing parsed folder IDs:
 *   - workspacesRootId: ID of the workspaces parent folder
 *   - versionsRootId: ID of the versions parent folder
 *   - publishedRootId: ID of the published parent folder
 * @returns {Object} Object containing:
 *   - versionFolderId: ID of the created version folder
 *   - workspaceFolderId: ID of the source workspace folder
 */
function createVersionFolder(lessonId, context) {
  Logger.log(`  - Getting versions parent folder (ID: ${context.versionsRootId})...`);
  const versionsParentFolder = DriveApp.getFolderById(context.versionsRootId);
  Logger.log(`  - SUCCESS: Found versions parent folder: "${versionsParentFolder.getName()}"`);

  // Get or create the lesson's versions folder
  Logger.log(`  - Getting or creating lesson versions folder: "${lessonId}"...`);
  const lessonVersionsFolder = getOrCreateFolder(versionsParentFolder, lessonId);
  Logger.log(`  - SUCCESS: Lesson versions folder ready`);

  // Generate version ID
  const versionId = generateVersionId();
  Logger.log(`  - Creating version folder: "${versionId}"...`);

  // Create the version folder
  const versionFolder = lessonVersionsFolder.createFolder(versionId);
  Logger.log(`  - SUCCESS: Version folder created`);

  // Get the lesson's workspace folder
  Logger.log(`  - Getting workspace parent folder (ID: ${context.workspacesRootId})...`);
  const workspaceParentFolder = DriveApp.getFolderById(context.workspacesRootId);
  Logger.log(`  - Finding workspace folder for lesson: "${lessonId}"...`);

  const workspaceFolders = workspaceParentFolder.getFoldersByName(lessonId);

  if (!workspaceFolders.hasNext()) {
    Logger.log(`  - FAILED: Workspace folder not found for lesson: "${lessonId}"`);
    throw new Error(`Workspace folder not found for lesson: ${lessonId}`);
  }

  const workspaceFolder = workspaceFolders.next();
  Logger.log(`  - SUCCESS: Found workspace folder`);

  // Copy all files from workspace to version folder, generating PDFs for Google Docs
  Logger.log(`  - Copying files and generating PDFs...`);
  const files = workspaceFolder.getFiles();
  let fileCount = 0;
  let pdfCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    // Copy the original file
    Logger.log(`    - Copying: "${fileName}"`);
    file.makeCopy(fileName, versionFolder);
    fileCount++;

    // If it's a Google Doc, also generate a PDF
    // Note: PDF generation only works for Google Docs, not DOCX files
    if (mimeType === MimeType.GOOGLE_DOCS) {
      const pdfBlob = file.getAs(MimeType.PDF);
      const pdfFileName = fileName.replace(/\.(docx?|gdoc)$/i, '') + '.pdf';
      Logger.log(`    - Generating PDF: "${pdfFileName}"`);
      versionFolder.createFile(pdfBlob).setName(pdfFileName);
      pdfCount++;
    }
  }

  Logger.log(`  - SUCCESS: Copied ${fileCount} files, generated ${pdfCount} PDFs`);

  // Create the VERSION Google Doc with version ID in the name
  const versionDocName = `VERSION - ${versionId}`;
  Logger.log(`  - Creating VERSION document: "${versionDocName}"...`);
  const versionDoc = DocumentApp.create(versionDocName);
  const versionDocFile = DriveApp.getFileById(versionDoc.getId());

  // Leave the document content empty (just close it)
  versionDoc.saveAndClose();

  // Move the VERSION doc into the version folder
  versionDocFile.moveTo(versionFolder);
  Logger.log(`  - SUCCESS: VERSION document created`);

  return {
    versionFolderId: versionFolder.getId(),
    workspaceFolderId: workspaceFolder.getId()
  };
}

/**
 * Publishes files by copying them from the version folder to the publish folder.
 * Deletes all existing files in the publish folder before copying new ones.
 *
 * @param {string} lessonId - The lesson identifier (used to find/create the publish folder)
 * @param {string} versionFolderId - The ID of the version folder to copy from
 * @param {Object} context - Context object containing parsed folder IDs:
 *   - workspacesRootId: ID of the workspaces parent folder
 *   - versionsRootId: ID of the versions parent folder
 *   - publishedRootId: ID of the published parent folder
 * @returns {string} The ID of the publish folder
 */
function publishVersionFiles(lessonId, versionFolderId, context) {
  Logger.log(`  - Getting version folder (ID: ${versionFolderId})...`);
  const versionFolder = DriveApp.getFolderById(versionFolderId);
  Logger.log(`  - SUCCESS: Found version folder`);

  Logger.log(`  - Getting published parent folder (ID: ${context.publishedRootId})...`);
  const publishedParentFolder = DriveApp.getFolderById(context.publishedRootId);
  Logger.log(`  - SUCCESS: Found published parent folder: "${publishedParentFolder.getName()}"`);

  // Get or create the lesson's published folder
  Logger.log(`  - Getting or creating publish folder for lesson: "${lessonId}"...`);
  const publishFolder = getOrCreateFolder(publishedParentFolder, lessonId);
  Logger.log(`  - SUCCESS: Publish folder ready`);

  // Delete all existing files in the publish folder
  Logger.log(`  - Deleting existing files in publish folder...`);
  const existingFiles = publishFolder.getFiles();
  let deletedCount = 0;

  while (existingFiles.hasNext()) {
    const file = existingFiles.next();
    const fileName = file.getName();
    Logger.log(`    - Deleting: "${fileName}"`);
    file.setTrashed(true);
    deletedCount++;
  }

  Logger.log(`  - SUCCESS: Deleted ${deletedCount} existing files`);

  // Copy all files from version folder to publish folder
  Logger.log(`  - Copying files to publish folder...`);
  const files = versionFolder.getFiles();
  let publishedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    Logger.log(`    - Publishing: "${fileName}"`);
    file.makeCopy(fileName, publishFolder);
    publishedCount++;
  }

  Logger.log(`  - SUCCESS: Published ${publishedCount} files to publish folder`);

  return publishFolder.getId();
}

/**
 * Gets an existing folder by name, or creates it if it doesn't exist
 * @param {Folder} parentFolder - The parent folder to search in
 * @param {string} folderName - The name of the folder to find or create
 * @returns {Folder} The found or created folder
 */
function getOrCreateFolder(parentFolder, folderName) {
  const existingFolders = parentFolder.getFoldersByName(folderName);

  if (existingFolders.hasNext()) {
    Logger.log(`    - Found existing folder: "${folderName}"`);
    return existingFolders.next();
  }

  Logger.log(`    - Creating new folder: "${folderName}"`);
  return parentFolder.createFolder(folderName);
}

/**
 * Generates a version ID based on the current timestamp
 * Format: yyyy-MM-dd_HH:mm:ss (24-hour format, zero-padded)
 * @returns {string} The version ID
 */
function generateVersionId() {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH:mm:ss');
}

/**
 * Gets a map of column names to their 1-indexed column numbers
 * @param {Sheet} sheet - The sheet to read headers from
 * @returns {Object} Map of lowercase column names to 1-indexed column numbers
 */
function getColumnIndices(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnIndices = {};

  headerRow.forEach((header, index) => {
    const normalizedHeader = String(header).toLowerCase().trim();
    if (normalizedHeader) {
      columnIndices[normalizedHeader] = index + 1; // 1-indexed for Sheets
    }
  });

  return columnIndices;
}

/**
 * Gets the dropdown values from a data validation rule in a column
 * @param {Sheet} sheet - The sheet to read from
 * @param {number} columnIndex - The 1-indexed column number
 * @returns {Array<string>|null} Array of dropdown values, or null if no data validation
 */
function getDropdownValues(sheet, columnIndex) {
  // Get data validation from the first data row (row 2)
  const range = sheet.getRange(2, columnIndex);
  const validation = range.getDataValidation();

  if (!validation) {
    Logger.log(`  - No data validation found in column ${columnIndex}`);
    return null;
  }

  const criteria = validation.getCriteriaType();

  // Check if it's a list validation (dropdown)
  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    const values = validation.getCriteriaValues()[0];
    Logger.log(`  - Found dropdown values: ${values}`);
    return values;
  } else if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    // If dropdown is based on a range, get values from that range
    const sourceRange = validation.getCriteriaValues()[0];
    const values = sourceRange.getValues().flat().filter(v => v !== '');
    Logger.log(`  - Found dropdown values from range: ${values}`);
    return values;
  }

  return null;
}

/**
 * Finds the row index for a lesson by its id
 * @param {Sheet} sheet - The sheet to search in
 * @param {Object} columnIndices - Map of column names to indices
 * @param {string} lessonId - The lesson id to find
 * @returns {number} The 1-indexed row number, or -1 if not found
 */
function findLessonRow(sheet, columnIndices, lessonId) {
  const idColumnIndex = columnIndices[COLUMNS.ID];
  if (!idColumnIndex) {
    Logger.log('  - WARNING: "id" column not found in sheet');
    return -1;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }

  const idColumn = sheet.getRange(2, idColumnIndex, lastRow - 1, 1).getValues();

  for (let i = 0; i < idColumn.length; i++) {
    if (String(idColumn[i][0]).trim() === lessonId) {
      return i + 2; // +2 because we started from row 2
    }
  }

  return -1;
}

/**
 * Updates the spreadsheet row with folder URLs, publish timestamp, and status.
 * Creates HYPERLINK formulas for folder URLs and sets the status to "Published".
 *
 * @param {Object} rowInfo - Object containing:
 *   - row: The 1-indexed row number to update
 *   - lessonId: The lesson identifier
 * @param {Object} columnIndices - Map of column names to 1-indexed column numbers
 * @param {string} workspaceFolderId - The ID of the workspace folder
 * @param {string} versionFolderId - The ID of the version folder
 * @param {string} publishFolderId - The ID of the publish folder
 */
function updateSpreadsheetRow(rowInfo, columnIndices, workspaceFolderId, versionFolderId, publishFolderId) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const targetRow = rowInfo.row;
  const lessonId = rowInfo.lessonId;

  Logger.log(`  - Updating row ${targetRow} for lesson: "${lessonId}"`);

  // Get folder objects and URLs
  const workspaceFolder = DriveApp.getFolderById(workspaceFolderId);
  const versionFolder = DriveApp.getFolderById(versionFolderId);
  const publishFolder = DriveApp.getFolderById(publishFolderId);

  const workspaceUrl = workspaceFolder.getUrl();
  const versionUrl = versionFolder.getUrl();
  const publishUrl = publishFolder.getUrl();
  const versionTimestamp = versionFolder.getName(); // The folder name IS the version ID

  // Convert version ID to ISO format for last_publish_time
  // Version ID format: yyyy-MM-dd_HH:mm:ss
  // ISO format:        yyyy-MM-ddTHH:mm:ss
  const publishTime = versionTimestamp.replace('_', 'T');

  Logger.log('  - Creating hyperlinked formulas...');

  // Update the cells with hyperlinked formulas
  if (columnIndices[COLUMNS.WORKSPACE_FOLDER_URL]) {
    const workspaceFormula = `=HYPERLINK("${workspaceUrl}", "folder")`;
    sheet.getRange(targetRow, columnIndices[COLUMNS.WORKSPACE_FOLDER_URL]).setFormula(workspaceFormula);
    Logger.log(`    - Set workspace_folder_url`);
  }

  if (columnIndices[COLUMNS.LATEST_VERSION_FOLDER_URL]) {
    const versionFormula = `=HYPERLINK("${versionUrl}", "${versionTimestamp}")`;
    sheet.getRange(targetRow, columnIndices[COLUMNS.LATEST_VERSION_FOLDER_URL]).setFormula(versionFormula);
    Logger.log(`    - Set latest_version_folder_url with timestamp: ${versionTimestamp}`);
  }

  if (columnIndices[COLUMNS.PUBLISH_FOLDER_URL]) {
    const publishFormula = `=HYPERLINK("${publishUrl}", "folder")`;
    sheet.getRange(targetRow, columnIndices[COLUMNS.PUBLISH_FOLDER_URL]).setFormula(publishFormula);
    Logger.log(`    - Set publish_folder_url`);
  }

  if (columnIndices[COLUMNS.LAST_PUBLISH_TIME]) {
    sheet.getRange(targetRow, columnIndices[COLUMNS.LAST_PUBLISH_TIME]).setValue(publishTime);
    Logger.log(`    - Set last_publish_time: ${publishTime}`);
  }

  // Update status to "Published"
  if (columnIndices[COLUMNS.STATUS]) {
    sheet.getRange(targetRow, columnIndices[COLUMNS.STATUS]).setValue(STATUS_VALUES.PUBLISHED);
    Logger.log(`    - Set status: ${STATUS_VALUES.PUBLISHED}`);
  }

  SpreadsheetApp.flush(); // Force the update to be visible immediately

  Logger.log('  - SUCCESS: All columns updated');
}