/**
 * @OnlyCurrentDoc
 */

// --- Configuration for JSON Exporter ---
const ROOT_EXPORT_FOLDER_NAME = "Sheet AI Exports";
const ARCHIVE_SUBFOLDER_NAME = "Archive";
const CURRENT_EXPORT_FILE_NAME = "current_sheet_data_sample.json";
const MAIN_SHEET_DATA_ROWS = 50; 
const REFERENCE_SHEET_DATA_ROWS = 10; 
const ERROR_LOG_SHEET_NAME = "ScriptErrorLog"; 
const MAX_ERROR_LOG_ENTRIES = 200; 

/**
 * NEW - Central location for project context notes provided to the AI.
 * Manually update these notes as the project evolves.
 * @returns {string[]} An array of context notes.
 */
function _getProjectContextNotes_() {
  return [
    "Context: This is an F&M Survey Analysis & Reporting Tool.",
    "Code.gs: Handles global utilities (error logging) and the JSON export process.",
    "OnOpen.gs: Contains the master onOpen(e) function that creates ALL custom menus.",
    "TrendAnalysis.gs: Contains data preparation and aggregation tools. It creates the vital 'SiteScoreSummary' sheet.",
    "Reporting.gs: Contains all logic for generating reports. It has been rebuilt to use 'TransformedData' as the single source of truth for all calculations."
  ];
}


/**
 * Wrapper function for the menu item to call initiateExportProcess without parameters.
 */
function initiateExportProcessMenu() {
    initiateExportProcess(); 
}


/**
 * Logs an error to a dedicated, hidden sheet and manages its size.
 * This function can be called by any script in the project.
 * @param {string} functionName The name of the function where the error occurred.
 * @param {Error} errorObject The error object caught.
 */
function logErrorToSheet(functionName, errorObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorLogSheet = ss.getSheetByName(ERROR_LOG_SHEET_NAME);

    if (!errorLogSheet) {
      errorLogSheet = ss.insertSheet(ERROR_LOG_SHEET_NAME);
      errorLogSheet.appendRow(["Timestamp", "Function Name", "Error Message", "Stack Trace"]);
      errorLogSheet.getRange("A1:D1").setFontWeight("bold");
      errorLogSheet.setFrozenRows(1);
      errorLogSheet.hideSheet(); 
      Logger.log(`Created and hid error log sheet: ${ERROR_LOG_SHEET_NAME}`);
    } else {
      if (errorLogSheet.isSheetHidden() === false) {
        errorLogSheet.hideSheet(); 
        Logger.log(`Error log sheet "${ERROR_LOG_SHEET_NAME}" was visible and has been re-hidden.`);
      }
    }

    const timestamp = new Date();
    errorLogSheet.appendRow([
      timestamp,
      functionName,
      errorObject.message || String(errorObject),
      errorObject.stack || ""
    ]);

    const lastRow = errorLogSheet.getLastRow();
    const currentErrorEntries = lastRow - 1; 

    if (currentErrorEntries > MAX_ERROR_LOG_ENTRIES) {
      const numRowsToDelete = currentErrorEntries - MAX_ERROR_LOG_ENTRIES;
      errorLogSheet.deleteRows(2, numRowsToDelete); 
      Logger.log(`Trimmed ${numRowsToDelete} old error(s) from "${ERROR_LOG_SHEET_NAME}". Kept ${MAX_ERROR_LOG_ENTRIES} entries.`);
    }

  } catch (e) {
    Logger.log(`CRITICAL: Failed to log error to sheet "${ERROR_LOG_SHEET_NAME}". Error: ${e.toString()}`);
    Logger.log(`Original Error in ${functionName}: Message - ${errorObject.message || String(errorObject)}, Stack - ${errorObject.stack || ""}`);
  }
}

/**
 * Initiates the data export process. 
 */
function initiateExportProcess(suppressFinalAlert = false) { 
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let directiveText = "";

    // NEW: Prompt for the current directive/goal.
    if (!suppressFinalAlert) {
        const directiveResponse = ui.prompt(
            'Set Current Directive',
            'What is the primary goal for this export? (e.g., "Fix the benchmark feature.")',
            ui.ButtonSet.OK_CANCEL
        );
        if (directiveResponse.getSelectedButton() !== ui.Button.OK) {
            ui.alert("Export Cancelled", "Export process cancelled by user.", ui.ButtonSet.OK);
            return null;
        }
        directiveText = directiveResponse.getResponseText().trim();
    }

    const allSheetNames = ss.getSheets().map(sheet => sheet.getName());
    if (allSheetNames.length === 0) {
      if (!suppressFinalAlert) {
          ui.alert("No Sheets Found", "This spreadsheet does not contain any sheets to export.", ui.ButtonSet.OK);
      }
      return null;
    }
    
    let mainSheetName = "TransformedData"; 

    if (!suppressFinalAlert) { 
        const mainSheetNameResponse = ui.prompt(
            'Select Main Sheet', 
            `Enter the name of the sheet for main export (e.g., TransformedData).\nDefault is "${mainSheetName}". Leave blank to use default.`, 
            ui.ButtonSet.OK_CANCEL
        );

        if (mainSheetNameResponse.getSelectedButton() !== ui.Button.OK) {
            ui.alert("Export Cancelled", "Export process cancelled by user.", ui.ButtonSet.OK);
            return null;
        }
        const responseText = mainSheetNameResponse.getResponseText().trim();
        if (responseText) { 
            mainSheetName = responseText;
        }
    } else {
        Logger.log(`Programmatic export triggered. Using default main sheet: "${mainSheetName}"`);
    }
    
    const mainSheet = ss.getSheetByName(mainSheetName);
    if (!mainSheet) {
      if (!suppressFinalAlert) {
          ui.alert("Sheet Not Found", `Sheet "${mainSheetName}" not found.`, ui.ButtonSet.OK);
      } else {
          Logger.log(`Sheet Not Found: "${mainSheetName}" during programmatic export.`);
      }
      logErrorToSheet('initiateExportProcess', new Error(`Sheet Not Found: "${mainSheetName}"`));
      return null; 
    }
    
    return exportSpreadsheetDataForAI(mainSheet, suppressFinalAlert, directiveText); 

  } catch (e) {
    logErrorToSheet('initiateExportProcess', e);
    if (!suppressFinalAlert) {
        ui.alert('Error', `An error occurred in initiateExportProcess: ${e.message}. Check ${ERROR_LOG_SHEET_NAME} (if accessible) and Logger.`, ui.ButtonSet.OK);
    }
    return null; 
  }
}

/**
 * Main function to export spreadsheet data to a JSON file.
 */
function exportSpreadsheetDataForAI(mainSheet, suppressFinalAlert = false, directive = "") { 
  const ui = SpreadsheetApp.getUi();
  
  if (!suppressFinalAlert) {
    const htmlOutput = HtmlService.createHtmlOutput("<p>Processing... Please wait.</p><p>This dialog will close automatically.</p>").setWidth(350).setHeight(120);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Exporting Data"); 
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    
    const exportData = {
      current_directive: directive || "No directive provided for this export.",
      project_context_notes: _getProjectContextNotes_(),
      exportTimestamp: new Date().toISOString(),
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      mainSheetData: null,
      referenceSheetsData: [],
      errorLogEntries: [] 
    };

    Logger.log(`Processing main sheet: ${mainSheet.getName()}`);
    exportData.mainSheetData = getSheetExportDataInternal(mainSheet, MAIN_SHEET_DATA_ROWS);

    allSheets.forEach(sheet => {
        if (sheet.getSheetId() !== mainSheet.getSheetId() && sheet.getName() !== ERROR_LOG_SHEET_NAME) {
            Logger.log(`Processing reference sheet: ${sheet.getName()}`);
            exportData.referenceSheetsData.push(getSheetExportDataInternal(sheet, REFERENCE_SHEET_DATA_ROWS));
        }
    });

    const errorLogSheet = ss.getSheetByName(ERROR_LOG_SHEET_NAME);
    if (errorLogSheet) {
      const lastErrorRow = errorLogSheet.getLastRow();
      if (lastErrorRow > 1) { 
        const numToRead = Math.min(lastErrorRow - 1, 5); 
        const errorDataRange = errorLogSheet.getRange(Math.max(2, lastErrorRow - numToRead + 1), 1, numToRead, 4);
        const errorValues = errorDataRange.getValues();
        
        const errorEntries = errorValues.map(row => ({
          timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
          functionName: String(row[1]),
          errorMessage: String(row[2]),
          stackTrace: String(row[3])
        })).filter(entry => entry.timestamp && entry.functionName); 

        errorEntries.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        exportData.errorLogEntries = errorEntries.slice(0, 2); 
      }
    }

    const jsonString = JSON.stringify(exportData, null, 2);
    const rootFolder = getOrCreateFolderInternal(ROOT_EXPORT_FOLDER_NAME, null);
    const archiveFolder = getOrCreateFolderInternal(ARCHIVE_SUBFOLDER_NAME, rootFolder);
    archivePreviousExportInternal(rootFolder, archiveFolder, CURRENT_EXPORT_FILE_NAME);
    const newFile = rootFolder.createFile(CURRENT_EXPORT_FILE_NAME, jsonString, MimeType.PLAIN_TEXT);
    
    if (!suppressFinalAlert) { 
      ui.alert('Export Complete!', `Data exported to: ${newFile.getName()}\nLink: ${newFile.getUrl()}`, ui.ButtonSet.OK);
    }
    
    return { 
      jsonString: jsonString, 
      fileUrl: newFile.getUrl(), 
      fileName: newFile.getName() 
    };

  } catch (e) {
    logErrorToSheet('exportSpreadsheetDataForAI', e);
    if (!suppressFinalAlert) {
        ui.alert('Export Error', `An error occurred during JSON export: ${e.message}. Check ${ERROR_LOG_SHEET_NAME} (if accessible) and Logger.`, ui.ButtonSet.OK);
    }
    return null; 
  }
}

/**
 * CORRECTED - Extracts headers and data rows. Exports ALL rows for report sheets.
 */
function getSheetExportDataInternal(sheet, numDataRows) {
  const sheetName = sheet.getName();
  const lastCol = sheet.getLastColumn();
  const lastRowInSheet = sheet.getLastRow();
  let headers = [];
  const dataRows = [];

  if (lastRowInSheet === 0 || lastCol === 0) {
    return { sheetName: sheetName, headers: ["Note: Sheet is empty."], dataRows: [] };
  }
  if (lastRowInSheet >= 1) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(header => String(header));
  } else {
    headers = ["Note: No header row found."];
  }

  const actualDataRowsAvailable = Math.max(0, lastRowInSheet - 1);
  
  // This is the corrected logic
  let rowsToFetch;
  if (sheetName.startsWith("Report -") || sheetName.startsWith("Theme Analysis -")) {
      rowsToFetch = actualDataRowsAvailable;
  } else {
      rowsToFetch = Math.min(numDataRows, actualDataRowsAvailable);
  }

  if (rowsToFetch > 0) {
    const dataRange = sheet.getRange(2, 1, rowsToFetch, lastCol); 
    const values = dataRange.getValues();
    const formulas = dataRange.getFormulas(); 
    for (let i = 0; i < rowsToFetch; i++) {
      dataRows.push({
        rowNumberInSheet: i + 2, 
        values: values[i].map(val => String(val)), 
        formulas: formulas[i].map(form => String(form)) 
      });
    }
  }
  return { sheetName: sheetName, headers: headers, dataRows: dataRows };
}

/**
 * Gets or creates a folder.
 */
function getOrCreateFolderInternal(folderName, parentFolder) {
  const root = parentFolder ? parentFolder : DriveApp.getRootFolder();
  const folders = root.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return root.createFolder(folderName);
  }
}

/**
 * Archives the previous export file.
 */
function archivePreviousExportInternal(rootExportFolder, archiveFolder, fileNameToArchive) {
  const files = rootExportFolder.getFilesByName(fileNameToArchive);
  if (files.hasNext()) {
    const fileToArchive = files.next();
    const timestamp = new Date().toISOString().replace(/:/g, "-").replace(/\..+/, ""); 
    const archiveFileName = `archived_${timestamp}_${fileToArchive.getName()}`;
    
    fileToArchive.moveTo(archiveFolder);
    const movedFile = DriveApp.getFileById(fileToArchive.getId()); 
    movedFile.setName(archiveFileName);
    Logger.log(`Successfully archived "${fileToArchive.getName()}" as "${archiveFileName}".`);
  } else {
    Logger.log(`No previous export file named "${fileNameToArchive}" found to archive.`);
  }
}

/**
 * Deletes all superseded, non-consolidated report sheets.
 */
function deleteOldReports() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Confirm Deletion', 'This will permanently delete all old report sheets that have been replaced by the new consolidated reports. Are you sure you want to continue?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) {
        ui.alert('Operation Cancelled.');
        return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    const sheetsToDelete = [];

    const oldReportPatterns = [
        "Report - Latest Scores",
        "Report - Score Changes",
        "Report - Consistent Performers",
        "Report - Service Trends",
        "Theme - "
    ];

    allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (oldReportPatterns.some(pattern => sheetName.startsWith(pattern))) {
            sheetsToDelete.push(sheetName);
            ss.deleteSheet(sheet);
        }
    });

    if (sheetsToDelete.length > 0) {
        ui.alert('Cleanup Complete', `The following ${sheetsToDelete.length} sheets were deleted:\n\n${sheetsToDelete.join('\n')}`);
    } else {
        ui.alert('No old reports found to delete.');
    }
}
