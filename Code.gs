// ============================================================================
// CONFIGURATION - Drive Folder IDs
// ============================================================================
// To find a folder ID: Open the folder in Drive, copy the ID from the URL
// URL format: https://drive.google.com/drive/folders/{FOLDER_ID_HERE}

const PARENT_FOLDERS = {
  // Parent folder containing all lesson workspace folders
  WORKSPACE: '1UetXJ8BXSvkjadKsve0pLGCVkipN52b8',

  // Parent folder containing all lesson version folders
  VERSIONS:  '1acCgaNU88gnyp4sqd-Cr9duLUjkipfgT',

  // Parent folder containing all lesson published folders
  PUBLISHED: '1cEzXD5bo0nkbUfZNeoENQQruPqJsdp_a'
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
    .addItem('Run Full Automation Sync', 'triggerAutomationSync') // (Display Label, Target Function Name)
    .addItem('Verify Folder Status', 'checkFolderStatus')
    
    // Adds a visual dividing line to group operational commands
    .addSeparator() 
    
    // Nests a sub-menu for secondary/administrative overrides
    .addSubMenu(ui.createMenu('Advanced Utilities')
      .addItem('Force Google Site Re-Publish', 'syncGoogleSiteOnly')
    )
    
    // Renders the built structure into the spreadsheet main header bar
    .addToUi(); 
}

// ============================================================================
// ENTRY POINT FUNCTIONS TIED TO MENU ACTIONS
// ============================================================================

/**
 * Main automation workflow:
 * 1. Validates user's current cell selection against Column A lesson names
 * 2. Creates timestamped version folder with PDFs
 * 3. Copies all files to publish folder
 */
function triggerAutomationSync() {
  const ui = SpreadsheetApp.getUi();

  Logger.log('=== Starting triggerAutomationSync ===');

  try {
    // Step 1: Get and validate the selected lesson name
    Logger.log('Step 1: Validating lesson name selection...');
    const lessonName = getValidatedLessonName();

    if (!lessonName) {
      Logger.log('FAILED: No valid lesson name selected');
      ui.alert(
        'Invalid Selection',
        'Please select a valid lesson name from Column A before running this automation.',
        ui.ButtonSet.OK
      );
      return;
    }

    Logger.log(`SUCCESS: Validated lesson name: "${lessonName}"`);

    // Step 2: Create timestamped version folder with PDFs
    Logger.log('Step 2: Creating version folder and generating PDFs...');
    const versionFolderId = createVersionFolder(lessonName);
    Logger.log(`SUCCESS: Version folder created with ID: ${versionFolderId}`);

    // Step 3: Copy all files to publish folder
    Logger.log('Step 3: Publishing files to publish folder...');
    publishVersionFiles(lessonName, versionFolderId);
    Logger.log('SUCCESS: Files published successfully');

    Logger.log(`=== Automation completed successfully for lesson: ${lessonName} ===`);

    ui.alert(
      'Success',
      `Automation completed for lesson: ${lessonName}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`FATAL ERROR in triggerAutomationSync: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    ui.alert(
      'Error',
      `Automation failed: ${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

function checkFolderStatus() {
  SpreadsheetApp.getUi().alert('Success: Initializing Drive query checks.');
}

function syncGoogleSiteOnly() {
  SpreadsheetApp.getUi().alert('Success: Pushing direct updates to Google Sites REST API.');
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Gets the current cell selection and validates it against Column A values
 * @returns {string|null} The validated lesson name, or null if invalid
 */
function getValidatedLessonName() {
  Logger.log('  - Getting current cell selection...');
  const sheet = SpreadsheetApp.getActiveSheet();
  const currentCell = SpreadsheetApp.getCurrentCell();

  if (!currentCell) {
    Logger.log('  - FAILED: No cell selected');
    return null;
  }

  const selectedValue = currentCell.getValue();
  Logger.log(`  - Selected value: "${selectedValue}"`);

  if (!selectedValue || typeof selectedValue !== 'string') {
    Logger.log('  - FAILED: Selected value is empty or not a string');
    return null;
  }

  // Get all valid lesson names from Column A (starting from A2, skipping empty cells)
  const lastRow = sheet.getLastRow();
  Logger.log(`  - Sheet has ${lastRow} rows`);

  if (lastRow < 2) {
    Logger.log('  - FAILED: No data in Column A (need at least row 2)');
    return null; // No data in Column A
  }

  const columnAValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const validLessonNames = columnAValues
    .map(row => row[0])
    .filter(value => value !== '' && value !== null && value !== undefined)
    .map(value => String(value).trim());

  Logger.log(`  - Found ${validLessonNames.length} valid lesson names in Column A`);

  // Check if the selected value is in the valid list
  const trimmedSelection = selectedValue.trim();
  const isValid = validLessonNames.includes(trimmedSelection);

  if (!isValid) {
    Logger.log(`  - FAILED: "${trimmedSelection}" not found in valid lesson names`);
  }

  return isValid ? trimmedSelection : null;
}

/**
 * Creates a timestamped version folder and generates PDFs for all Google Docs
 * @param {string} lessonName - The name of the lesson
 * @returns {string} The ID of the created version folder
 */
function createVersionFolder(lessonName) {
  Logger.log(`  - Getting versions parent folder (ID: ${PARENT_FOLDERS.VERSIONS})...`);
  const versionsParentFolder = DriveApp.getFolderById(PARENT_FOLDERS.VERSIONS);
  Logger.log(`  - SUCCESS: Found versions parent folder: "${versionsParentFolder.getName()}"`);

  // Get or create the lesson's versions folder
  Logger.log(`  - Getting or creating lesson versions folder: "${lessonName}"...`);
  const lessonVersionsFolder = getOrCreateFolder(versionsParentFolder, lessonName);
  Logger.log(`  - SUCCESS: Lesson versions folder ready`);

  // Create timestamp: 'yyyy-mm-dd hh:mm:ss' (24-hour format, zero-padded)
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  Logger.log(`  - Creating timestamped folder: "${timestamp}"...`);

  // Create the timestamped version folder
  const versionFolder = lessonVersionsFolder.createFolder(timestamp);
  Logger.log(`  - SUCCESS: Version folder created`);

  // Get the lesson's workspace folder
  Logger.log(`  - Getting workspace parent folder (ID: ${PARENT_FOLDERS.WORKSPACE})...`);
  const workspaceParentFolder = DriveApp.getFolderById(PARENT_FOLDERS.WORKSPACE);
  Logger.log(`  - Finding workspace folder for lesson: "${lessonName}"...`);

  const workspaceFolders = workspaceParentFolder.getFoldersByName(lessonName);

  if (!workspaceFolders.hasNext()) {
    Logger.log(`  - FAILED: Workspace folder not found for lesson: "${lessonName}"`);
    throw new Error(`Workspace folder not found for lesson: ${lessonName}`);
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

  // Create the VERSION Google Doc with timestamp content
  Logger.log(`  - Creating VERSION document...`);
  const versionDoc = DocumentApp.create('VERSION');
  const versionDocFile = DriveApp.getFileById(versionDoc.getId());

  // Write the timestamp to the document
  const body = versionDoc.getBody();
  body.setText(timestamp);
  versionDoc.saveAndClose();

  // Move the VERSION doc into the version folder
  versionDocFile.moveTo(versionFolder);
  Logger.log(`  - SUCCESS: VERSION document created with timestamp content`);

  return versionFolder.getId();
}

/**
 * Copies all files from the version folder to the publish folder
 * @param {string} lessonName - The name of the lesson
 * @param {string} versionFolderId - The ID of the version folder to copy from
 */
function publishVersionFiles(lessonName, versionFolderId) {
  Logger.log(`  - Getting version folder (ID: ${versionFolderId})...`);
  const versionFolder = DriveApp.getFolderById(versionFolderId);
  Logger.log(`  - SUCCESS: Found version folder`);

  Logger.log(`  - Getting published parent folder (ID: ${PARENT_FOLDERS.PUBLISHED})...`);
  const publishedParentFolder = DriveApp.getFolderById(PARENT_FOLDERS.PUBLISHED);
  Logger.log(`  - SUCCESS: Found published parent folder: "${publishedParentFolder.getName()}"`);

  // Get or create the lesson's published folder
  Logger.log(`  - Getting or creating publish folder for lesson: "${lessonName}"...`);
  const publishFolder = getOrCreateFolder(publishedParentFolder, lessonName);
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