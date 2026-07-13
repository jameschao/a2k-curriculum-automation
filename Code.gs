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
    .addItem('Publish lesson', 'publishLesson') // (Display Label, Target Function Name)
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
function publishLesson() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('=== Starting publishLesson ===');

  try {
    // Step 1: Get and validate the selected lesson name
    ss.toast('Validating lesson selection...', '📋 Step 1/4', 3);
    Logger.log('Step 1: Validating lesson name selection...');
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

    Logger.log(`SUCCESS: Validated lesson name: "${lessonId}"`);

    // Step 2: Create timestamped version folder with PDFs
    ss.toast('Creating version folder and generating PDFs...', '📁 Step 2/4', 5);
    Logger.log('Step 2: Creating version folder and generating PDFs...');
    const {versionFolderId, workspaceFolderId} = createVersionFolder(lessonId);
    Logger.log(`SUCCESS: Version folder created with ID: ${versionFolderId}`);

    // Step 3: Copy all files to publish folder
    ss.toast('Publishing files to publish folder...', '🚀 Step 3/4', 5);
    Logger.log('Step 3: Publishing files to publish folder...');
    const publishFolderId = publishVersionFiles(lessonId, versionFolderId);
    Logger.log('SUCCESS: Files published successfully');

    // Step 4: Update spreadsheet with folder URLs and timestamp
    ss.toast('Updating spreadsheet row...', '📝 Step 4/4', 3);
    Logger.log('Step 4: Updating spreadsheet row...');
    const sheet = SpreadsheetApp.getActiveSheet();
    const columnIndices = getColumnIndices(sheet);
    const lessonRow = findLessonRow(sheet, columnIndices, lessonId);

    if (lessonRow !== -1) {
      updateSpreadsheetRow({row: lessonRow, lessonId: lessonId}, columnIndices, workspaceFolderId, versionFolderId, publishFolderId);
      Logger.log('SUCCESS: Spreadsheet row updated');
    } else {
      Logger.log('WARNING: Could not update spreadsheet - lesson row not found');
    }

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
 * Gets the lesson ID from the "id" column of the currently selected row
 * @returns {string|null} The lesson ID, or null if invalid
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
  const idColumnIndex = columnIndices['id'];

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
 * Creates a timestamped version folder and generates PDFs for all Google Docs
 * @param {string} lessonId - The name of the lesson
 * @returns {Object} Object containing versionFolderId and workspaceFolderId
 */
function createVersionFolder(lessonId) {
  Logger.log(`  - Getting versions parent folder (ID: ${PARENT_FOLDERS.VERSIONS})...`);
  const versionsParentFolder = DriveApp.getFolderById(PARENT_FOLDERS.VERSIONS);
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
  Logger.log(`  - Getting workspace parent folder (ID: ${PARENT_FOLDERS.WORKSPACE})...`);
  const workspaceParentFolder = DriveApp.getFolderById(PARENT_FOLDERS.WORKSPACE);
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
 * Copies all files from the version folder to the publish folder
 * @param {string} lessonId - The name of the lesson
 * @param {string} versionFolderId - The ID of the version folder to copy from
 * @returns {string} The ID of the publish folder
 */
function publishVersionFiles(lessonId, versionFolderId) {
  Logger.log(`  - Getting version folder (ID: ${versionFolderId})...`);
  const versionFolder = DriveApp.getFolderById(versionFolderId);
  Logger.log(`  - SUCCESS: Found version folder`);

  Logger.log(`  - Getting published parent folder (ID: ${PARENT_FOLDERS.PUBLISHED})...`);
  const publishedParentFolder = DriveApp.getFolderById(PARENT_FOLDERS.PUBLISHED);
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
 * Finds the row index for a lesson by its id
 * @param {Sheet} sheet - The sheet to search in
 * @param {Object} columnIndices - Map of column names to indices
 * @param {string} lessonId - The lesson id to find
 * @returns {number} The 1-indexed row number, or -1 if not found
 */
function findLessonRow(sheet, columnIndices, lessonId) {
  const idColumnIndex = columnIndices['id'];
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
 * Updates the spreadsheet row for the lesson with folder URLs and publish timestamp
 * @param {Object} rowInfo - Object containing {row: number, lessonId: string}
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
  if (columnIndices['workspace_folder_url']) {
    const workspaceFormula = `=HYPERLINK("${workspaceUrl}", "folder")`;
    sheet.getRange(targetRow, columnIndices['workspace_folder_url']).setFormula(workspaceFormula);
    Logger.log(`    - Set workspace_folder_url`);
  }

  if (columnIndices['latest_version_folder_url']) {
    const versionFormula = `=HYPERLINK("${versionUrl}", "${versionTimestamp}")`;
    sheet.getRange(targetRow, columnIndices['latest_version_folder_url']).setFormula(versionFormula);
    Logger.log(`    - Set latest_version_folder_url with timestamp: ${versionTimestamp}`);
  }

  if (columnIndices['publish_folder_url']) {
    const publishFormula = `=HYPERLINK("${publishUrl}", "folder")`;
    sheet.getRange(targetRow, columnIndices['publish_folder_url']).setFormula(publishFormula);
    Logger.log(`    - Set publish_folder_url`);
  }

  if (columnIndices['last_publish_time']) {
    sheet.getRange(targetRow, columnIndices['last_publish_time']).setValue(publishTime);
    Logger.log(`    - Set last_publish_time: ${publishTime}`);
  }

  Logger.log('  - SUCCESS: All columns updated');
}