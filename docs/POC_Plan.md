# Architecture & Implementation Documentation

**Project:** Master Control Sheet Automation (File Syncing & Google Sites Publishing)

**Target Audience:** Engineering Team

This document outlines the architectural strategy, immediate Proof of Concept (POC) workflow, and decoupled execution model for the organization's Master Control automation tool.

---

## 1. Automation Goals & Core Operations

The primary objective of this system is to automate file management and documentation publishing across separate corporate storage boundaries and internal web assets. The system must reliably execute the following operations:

* **Targeted Directory Generation:** Create new project directories within a designated target Shared Drive (inheriting or explicitly establishing structural parent-child relationships) and programmatically update user/group permissions.
* **Existential Checks:** Query specific parent directories to verify whether a folder with a given project name already exists before executing downstream operations.
* **Cross-Boundary Resource Copying:** Duplicate files (including Google Docs, Sheets, Office files, and media assets/videos) from a source Shared Drive to a destination Shared Drive.
* **Dynamic PDF Compilation:** Intercept all Microsoft Word (`.docx`) and native Google Docs during the transfer process, convert them to PDF format on the fly, and place them into the newly created target directory.
* **Decoupled System Status Updates:** Intercept user updates in the spreadsheet grid and programmatically manipulate specific cell states (values and formulas) to track system execution.
* **Dynamic Workspace Communications:** Generate and dispatch automated emails via Gmail containing dynamic text fields populated straight from the spreadsheet data.
* **Dynamic Web Publishing:** Programmatically construct new pages within a modern Google Site (leveraging layout templates via string interpolation), embed public file viewing links, and expose them to the team.

---

## 2. Immediate POC Strategy (Rapid Prototyping)

To quickly validate the primary integration points before configuring the local toolchain, the initial prototype will be built directly inside the cloud environment using zero-setup tools.

* **Development Environment:** The **In-Browser Apps Script Editor** accessible via *Extensions > Apps Script* inside the Master Control sheet. This ensures immediate OAuth handshakes and instant debugger visibility.
* **Source Control Bridge:** The **Google Apps Script GitHub Assistant** Chrome extension will be utilized. This tool injects a Git panel directly into the browser editor UI.
* **Workflow:** Developers will use the extension to visually diff files and push code snapshots with commit messages directly to a remote GitHub repository.
* **Warning:** Pulling from GitHub via the extension executes a total file overwrite in the browser editor. The script state must be pushed before attempting a pull.


* **User Interface Layer:** The POC will utilize a **Custom Menu** injected into the Google Sheets main menu bar using the native `onOpen()` function. This prevents UI clutter and provides an intuitive, native feel for testing.

### POC Custom Menu Implementation Code

Drop this baseline implementation into your browser-based `Main.gs` file to establish the UI entry points:

```javascript
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

```

---

## 3. Script Interaction & Asynchronous Execution Model

To keep the Master Control sheet responsive for end-users, the architecture utilizes a decoupled, asynchronous model where the Container-Bound script acts as the UI frontend, the Spreadsheet acts as a Shared Message Queue, and a Standalone script acts as the heavy-lifting backend worker.

### The Step-by-Step Lifecycle

1. **User Trigger:** The user navigates to the custom menu bar (**⚙️ Master Control > Run Full Automation Sync**) and clicks the menu item.
2. **State Initiation (Bound Script):** The container-bound script executes instantly. It identifies the target row and updates the target cell in the "Status" column to `"pending publish"`. Its execution then immediately terminates, keeping the user interface snappy.
3. **The Handshake (`onChange`):** The structural modification to the cell triggers an installable `onChange` listener on the **Standalone Script**.
4. **Backend Processing (Standalone Script):**
* The standalone script wakes up and verifies that `e.changeType === 'EDIT'`.
* It opens the spreadsheet programmatically using `SpreadsheetApp.openById('MC_SHEET_ID')`.
* It scans the status column for *all* rows marked `"pending publish"`.
* It executes the heavy API tasks: copying files across Shared Drives utilizing the **Advanced Drive API** (`supportsAllDrives: true`), generating PDFs on the fly, and creating pages via HTTP requests (`UrlFetchApp`) to the **New Google Sites REST API**.


5. **State Completion:** Once the API processes complete successfully, the standalone script rewrites the row's status cell to `"published"`.

### Crucial Engineering Guardrails

> ⚠️ **Loop Prevention:** Writing `"published"` back to the spreadsheet will trigger the standalone script's `onChange` listener a second time. The very first line of the standalone processing function must evaluate the cell value. If the value is `"published"`, it must execute an immediate `return;` statement to prevent an infinite loop.

> ⚠️ **Idempotency & Bulk Processing:** Installable triggers are subject to queueing delays under Google server load. The standalone script must be written to pull and process *all* rows currently marked `"pending publish"` in a single run, rather than assuming it was triggered for only a single event.

---

## 4. Long-Term Production Strategy (CLASP Setup)

For the final, enterprise-grade deployment, the automation must follow standard software engineering practices. We will bypass the browser UI to ensure type safety, version control, and multi-user maintainability.

* **Local Development Environment:** Code will be developed locally using **VS Code** and written in **TypeScript**. Autocomplete and compile-time type checking will be handled via the `@types/google-apps-script` package.
* **Version Management:** The codebase will live in a **GitHub** repository, enabling standard branching, pull requests, and peer reviews.
* **Deployment Mechanism:** Google's official command-line tool, **CLASP** (`npm install -g @google/clasp`), will manage the connection to the cloud. Running `clasp push` will automatically transpile the TypeScript files into vanilla JavaScript and upload them directly to Google's servers.
* **Insulation & Governance:**
* The Master Control Google Sheet and associated scripts **must** reside in a Google **Shared Drive** to ensure corporate ownership survives employee turnover.
* The project will be bound to a **Standard Google Cloud Project (GCP)** inside the company console to enable advanced centralized logging (Cloud Logging) and enterprise API quota management.