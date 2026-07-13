/**
 * The onOpen function runs automatically every time the spreadsheet loads.
 * It builds the custom menu structure inside the Google Sheets UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Instantiates the top-level main menu container
  ui.createMenu('⚙️ Master Control')
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

// --- Entry Point Functions Tied to Menu Actions ---

function triggerAutomationSync() {
  // Simple browser notification to confirm the UI connection works
  SpreadsheetApp.getUi().alert('Success: Initializing Master Control processing pipeline.');
  // POC cell modification logic will be invoked here
}

function checkFolderStatus() {
  SpreadsheetApp.getUi().alert('Success: Initializing Drive query checks.');
}

function syncGoogleSiteOnly() {
  SpreadsheetApp.getUi().alert('Success: Pushing direct updates to Google Sites REST API.');
}