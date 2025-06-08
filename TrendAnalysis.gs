/**
 * @OnlyCurrentDoc
 */

// --- Configuration ---
const HALF_SEQUENCE_SHEET_NAME = "HalfSequenceLookup"; 
const SITE_SCORE_SUMMARY_SHEET_NAME = "SiteScoreSummary"; 
const TREND_SETTINGS_PROPERTY_KEY = "trendAnalysisSettings"; 

// --- AI Interaction Note (For Future AI Model Assistance) ---
// The 'Trend Analysis' menu is created by the onOpen function in Code.gs.
// This script file contains the functions for trend analysis tools.
// It calls logErrorToSheet which is expected to be globally available from Code.gs.
// --- End AI Interaction Note ---

// No onOpen(e) function here; menu creation is handled by Code.gs.

/**
 * Main function to set up or refresh trend analysis columns.
 */
function setupTrendColumnsWithImputation() {
  const ui = SpreadsheetApp.getUi();
  const functionName = 'setupTrendColumnsWithImputation';

  try {
    const settings = {}; 

    const sheetNameResponse = ui.prompt(functionName, 'Enter name of sheet for trend analysis (e.g., JohnLewis):', ui.ButtonSet.OK_CANCEL);
    if (sheetNameResponse.getSelectedButton() !== ui.Button.OK || !sheetNameResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
    settings.targetSheetName = sheetNameResponse.getResponseText().trim();

    const halfColResponse = ui.prompt(functionName, 'Enter column letter for "Half" identifiers (e.g., A):', ui.ButtonSet.OK_CANCEL);
    if (halfColResponse.getSelectedButton() !== ui.Button.OK || !halfColResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
    settings.halfColLetter = halfColResponse.getResponseText().trim().toUpperCase();

    const siteIdColResponse = ui.prompt(functionName, 'Enter column letter for "Site Identifier" (e.g., Q for Site No.):', ui.ButtonSet.OK_CANCEL);
    if (siteIdColResponse.getSelectedButton() !== ui.Button.OK || !siteIdColResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
    settings.siteIdColLetter = siteIdColResponse.getResponseText().trim().toUpperCase();

    const actualScoreColResponse = ui.prompt(functionName, 'Enter column letter for the "Actual Score" to track (e.g., S for Total Score):', ui.ButtonSet.OK_CANCEL);
    if (actualScoreColResponse.getSelectedButton() !== ui.Button.OK || !actualScoreColResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
    settings.actualScoreColLetter = actualScoreColResponse.getResponseText().trim().toUpperCase();
    
    const scoreNameResponse = ui.prompt(functionName, 'Enter descriptive name of this score for new headers (e.g., "Total Score"):', ui.ButtonSet.OK_CANCEL);
    if (scoreNameResponse.getSelectedButton() !== ui.Button.OK || !scoreNameResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
    settings.scoreName = scoreNameResponse.getResponseText().trim();

    const scoreTypeResponseText = ui.prompt(functionName, `Is "${settings.scoreName}" an 'Overall Survey Score' or an 'Individual Question Score'?\nEnter 'Overall' or 'Individual':`, ui.ButtonSet.OK_CANCEL).getResponseText();
    if (!scoreTypeResponseText) { ui.alert('Operation Cancelled.'); return; }
    settings.scoreType = scoreTypeResponseText.trim().toLowerCase();

    settings.isIndividualQuestionScore = (settings.scoreType === 'individual');
    settings.questionIdColLetter = "";
    if (settings.isIndividualQuestionScore) {
      const questionIdColResponse = ui.prompt(functionName, 'If tracking individual question score, enter column letter for "Question ID" (e.g., X):', ui.ButtonSet.OK_CANCEL);
      if (questionIdColResponse.getSelectedButton() !== ui.Button.OK || !questionIdColResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
      settings.questionIdColLetter = questionIdColResponse.getResponseText().trim().toUpperCase();
    } else if (settings.scoreType !== 'overall') {
      ui.alert('Invalid Score Type', 'Score type must be "Overall" or "Individual". Operation cancelled.', ui.ButtonSet.OK); return;
    }

    const imputationResponseText = ui.prompt(functionName, `If an actual score is missing, how should it be filled?\n1. None \n2. LOCF (Last Observation Carried Forward)\nEnter 'None' or 'LOCF':`, ui.ButtonSet.OK_CANCEL).getResponseText();
    if (!imputationResponseText) { ui.alert('Operation Cancelled.'); return; }
    settings.imputationMethod = imputationResponseText.trim().toUpperCase();
    if (settings.imputationMethod !== "NONE" && settings.imputationMethod !== "LOCF") {
      ui.alert('Invalid Imputation Method', 'Enter "None" or "LOCF". Operation cancelled.', ui.ButtonSet.OK); return;
    }

    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.targetSheetName);
    if (!targetSheet) {
        ui.alert('Error', `Target sheet "${settings.targetSheetName}" not found. Cannot determine column placement.`, ui.ButtonSet.OK);
        return;
    }
    settings.firstNewColIndex = targetSheet.getLastColumn() + 1; 
    Logger.log(`New trend columns for "${settings.scoreName}" will start at column index: ${settings.firstNewColIndex} in sheet "${settings.targetSheetName}"`);

    _executeTrendAnalysisLogic(settings);

    const specificSettingsKey = `${TREND_SETTINGS_PROPERTY_KEY}_${settings.targetSheetName}_${settings.scoreName.replace(/\s+/g, '_')}`;
    PropertiesService.getDocumentProperties().setProperty(specificSettingsKey, JSON.stringify(settings));
    Logger.log(`Trend settings for "${settings.scoreName}" on sheet "${settings.targetSheetName}" saved to Document Properties with key: ${specificSettingsKey}`);

  } catch (e) {
    try { logErrorToSheet(functionName, e); } catch (le) { Logger.log("Failed to call logErrorToSheet from TrendAnalysis.gs: " + le.message); }
    ui.alert('Error', `An error occurred in ${functionName}: ${e.message}. Check ScriptErrorLog and Logger.`, ui.ButtonSet.OK);
  }
}

/**
 * Quick refresh function that uses stored settings.
 */
function quickRefreshTrendData() {
  const ui = SpreadsheetApp.getUi();
  const functionName = 'quickRefreshTrendData';
  Logger.log(`Starting ${functionName}`);

  const allProperties = PropertiesService.getDocumentProperties().getKeys();
  const trendSettingKeys = allProperties.filter(key => key.startsWith(TREND_SETTINGS_PROPERTY_KEY + "_"));
  
  if (trendSettingKeys.length === 0) {
    ui.alert('No Saved Settings', 'No trend analyses have been configured yet. Please run the full "3. Setup/Add Trend Columns with Imputation" first.', ui.ButtonSet.OK);
    return;
  }

  let settingsKeyToUse;
  if (trendSettingKeys.length === 1) {
    settingsKeyToUse = trendSettingKeys[0];
  } else {
    const choices = trendSettingKeys.map(key => {
      const parts = key.substring(TREND_SETTINGS_PROPERTY_KEY.length + 1).split('_');
      const sheetName = parts.shift(); 
      const scoreName = parts.join(' '); 
      return `Sheet: ${sheetName}, Score: ${scoreName} (Key: ${key})`;
    });

    const choiceResponse = ui.prompt('Select Trend to Refresh', 'Multiple trend configurations found. Which one do you want to refresh?\n\nEnter the full key (e.g., ' + TREND_SETTINGS_PROPERTY_KEY + '_SheetName_Score_Name):\n\n' + choices.join('\n'), ui.ButtonSet.OK_CANCEL);
    
    if (choiceResponse.getSelectedButton() !== ui.Button.OK || !choiceResponse.getResponseText()) {
      ui.alert('Operation Cancelled', 'Quick refresh cancelled.', ui.ButtonSet.OK);
      return;
    }
    const chosenKey = choiceResponse.getResponseText().trim();
    if (!trendSettingKeys.includes(chosenKey)) {
      ui.alert('Invalid Key', 'The entered key does not match any saved trend configurations.', ui.ButtonSet.OK);
      return;
    }
    settingsKeyToUse = chosenKey;
  }
  
  const storedSettingsString = PropertiesService.getDocumentProperties().getProperty(settingsKeyToUse);
  if (!storedSettingsString) {
    ui.alert('Error', 'Could not retrieve the selected settings. The key might have been deleted or an error occurred.', ui.ButtonSet.OK);
    return;
  }

  try {
    const settings = JSON.parse(storedSettingsString);
    Logger.log(`Retrieved settings for key "${settingsKeyToUse}": ${JSON.stringify(settings)}`);
    
    const htmlOutput = HtmlService.createHtmlOutput("<p>Refreshing trend data for: <b>" + settings.scoreName + "</b> on sheet <b>" + settings.targetSheetName + "</b>... Please wait.</p><p>This dialog will close automatically.</p>")
                                  .setWidth(400).setHeight(150);
    ui.showModalDialog(htmlOutput, "Quick Refreshing Trends");

    _executeTrendAnalysisLogic(settings); 
  } catch (e) {
    try { logErrorToSheet(functionName, e); } catch (le) { Logger.log("Failed to call logErrorToSheet from TrendAnalysis.gs: " + le.message); }
    ui.alert('Error', `An error occurred during quick refresh for "${settingsKeyToUse}": ${e.message}. Check ScriptErrorLog and Logger.`, ui.ButtonSet.OK);
  }
}


/**
 * Internal worker function to perform trend calculations and write data.
 * @param {object} settings The settings object containing all necessary parameters.
 */
function _executeTrendAnalysisLogic(settings) {
  const ui = SpreadsheetApp.getUi(); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const functionName = '_executeTrendAnalysisLogic';

  const sheet = ss.getSheetByName(settings.targetSheetName);
  if (!sheet) { 
    Logger.log(`Target sheet "${settings.targetSheetName}" not found during execution.`);
    ui.alert('Error', `Target sheet "${settings.targetSheetName}" not found. Please run full setup again or check sheet name.`, ui.ButtonSet.OK); 
    return; 
  }

  const halfColIndex = letterToColumn(settings.halfColLetter);
  const siteIdColIndex = letterToColumn(settings.siteIdColLetter);
  const actualScoreColIndex = letterToColumn(settings.actualScoreColLetter);
  let questionIdColIndex = -1;
  if (settings.isIndividualQuestionScore && settings.questionIdColLetter) {
    questionIdColIndex = letterToColumn(settings.questionIdColLetter);
  }

  const coreColsValid = [halfColIndex, siteIdColIndex, actualScoreColIndex].every(idx => !isNaN(idx) && idx >= 1);
  const individualColValid = settings.isIndividualQuestionScore ? (!isNaN(questionIdColIndex) && questionIdColIndex >= 1) : true;
  if (!coreColsValid || !individualColValid) {
    ui.alert('Error', 'Invalid column configuration. Please run full setup again or check saved settings.', ui.ButtonSet.OK); 
    return;
  }

  const lastDataRow = sheet.getLastRow();
  if (lastDataRow < 2) { 
    ui.alert('Not enough data', `Sheet "${settings.targetSheetName}" needs a header and at least one data row.`, ui.ButtonSet.OK); 
    return; 
  }
  
  const firstNewColIndex = settings.firstNewColIndex;
  if (!firstNewColIndex || isNaN(firstNewColIndex) || firstNewColIndex < 1) {
      Logger.log(`Error: firstNewColIndex is invalid or missing in settings for ${settings.targetSheetName} - ${settings.scoreName}. Value: ${firstNewColIndex}`);
      ui.alert('Configuration Error', `Could not determine where to write trend columns for "${settings.scoreName}". Please run the full setup again for this score on this sheet.`, ui.ButtonSet.OK);
      return;
  }

  const prevHalfIdColNum = firstNewColIndex;
  const prevActualScoreColNum = firstNewColIndex + 1;
  const imputedFinalScoreColNum = firstNewColIndex + 2;
  const scoreBasisColNum = firstNewColIndex + 3;
  const changeInScoreColNum = firstNewColIndex + 4;

  if (prevHalfIdColNum <= sheet.getMaxColumns() && lastDataRow > 1) {
    const rowsToClear = sheet.getMaxRows() -1 ; 
    if (rowsToClear > 0) {
      Logger.log(`Clearing content from ${settings.targetSheetName}, Row 2, Col ${prevHalfIdColNum} for ${rowsToClear} rows, 5 columns wide.`);
      sheet.getRange(2, prevHalfIdColNum, rowsToClear, 5).clearContent().clearNote();
    }
  } else {
      Logger.log(`Skipping clear for ${settings.targetSheetName} as prevHalfIdColNum (${prevHalfIdColNum}) might be out of bounds or no data rows.`);
  }

  const refreshTimestamp = `Last Refreshed: ${new Date().toLocaleString()}`;
  sheet.getRange(1, prevHalfIdColNum).setValue("Previous Half ID").setFontWeight("bold");
  sheet.getRange(1, prevActualScoreColNum).setValue(`Previous Actual ${settings.scoreName}`).setFontWeight("bold");
  sheet.getRange(1, imputedFinalScoreColNum).setValue(`Imputed/Final ${settings.scoreName}`).setFontWeight("bold");
  sheet.getRange(1, scoreBasisColNum).setValue(`${settings.scoreName} Basis`).setFontWeight("bold");
  sheet.getRange(1, changeInScoreColNum).setValue(`Change in ${settings.scoreName} (vs Prev Actual)`).setNote(refreshTimestamp).setFontWeight("bold");

  Logger.log(`Reading data from "${settings.targetSheetName}" for calculations...`);
  const colsToReadFromSource = [halfColIndex, siteIdColIndex, actualScoreColIndex];
  if (settings.isIndividualQuestionScore && questionIdColIndex > 0) {
    colsToReadFromSource.push(questionIdColIndex);
  }
  const maxColToReadFromSource = Math.max(...colsToReadFromSource, sheet.getLastColumn()); 
  
  if (lastDataRow -1 <= 0) {
      Logger.log("No data rows to process in " + settings.targetSheetName);
      ui.alert('Success!', `Trend columns for "${settings.scoreName}" updated. No data rows to process.`, ui.ButtonSet.OK);
      return;
  }
  const sourceDataValues = sheet.getRange(2, 1, lastDataRow - 1, maxColToReadFromSource).getValues();

  const halfSequenceSheet = ss.getSheetByName(HALF_SEQUENCE_SHEET_NAME);
  let halfSequenceMap = new Map();
  if (halfSequenceSheet) {
    const seqLastRow = halfSequenceSheet.getLastRow();
    if (seqLastRow > 1) {
      const seqValues = halfSequenceSheet.getRange(2, 1, seqLastRow - 1, 2).getValues();
      seqValues.forEach(row => { if (row[0]) halfSequenceMap.set(String(row[0]).trim(), String(row[1] || "").trim()); });
    }
    Logger.log(`Loaded ${halfSequenceMap.size} entries from ${HALF_SEQUENCE_SHEET_NAME}.`);
  } else {
    Logger.log(`Warning: "${HALF_SEQUENCE_SHEET_NAME}" not found. Previous Half IDs will be blank.`);
  }

  const actualScoresMap = new Map();
  sourceDataValues.forEach((row) => {
    const currentHalf = String(row[halfColIndex - 1]).trim();
    const siteId = String(row[siteIdColIndex - 1]).trim();
    const actualScoreRaw = row[actualScoreColIndex - 1]; 
    const actualScore = (actualScoreRaw === "" || actualScoreRaw === null || isNaN(parseFloat(actualScoreRaw))) ? null : parseFloat(actualScoreRaw);

    if (siteId && currentHalf) {
      let key = `${siteId}_${currentHalf}`;
      if (settings.isIndividualQuestionScore) {
        const questionId = String(row[questionIdColIndex - 1]).trim();
        if (questionId) key += `_${questionId}`;
        else return; 
      }
      if (actualScore !== null) { 
          actualScoresMap.set(key, actualScore);
      }
    }
  });
  Logger.log(`Populated actualScoresMap with ${actualScoresMap.size} entries.`);

  const outputDataBlock = [];

  for (let i = 0; i < sourceDataValues.length; i++) {
    const currentRowData = sourceDataValues[i];
    const currentHalf = String(currentRowData[halfColIndex - 1]).trim();
    const siteId = String(currentRowData[siteIdColIndex - 1]).trim();
    const actualCurrentScoreRaw = currentRowData[actualScoreColIndex - 1];
    const actualCurrentScore = (actualCurrentScoreRaw === "" || actualCurrentScoreRaw === null || isNaN(parseFloat(actualCurrentScoreRaw))) ? null : parseFloat(actualCurrentScoreRaw);
    
    let questionId = "";
    if (settings.isIndividualQuestionScore && questionIdColIndex > 0) {
        questionId = String(currentRowData[questionIdColIndex - 1]).trim();
    }
    
    if (!siteId || !currentHalf) { 
        outputDataBlock.push(["", "", "", "Missing ID/Half", ""]);
        continue;
    }
     if (settings.isIndividualQuestionScore && !questionId) { 
        outputDataBlock.push(["", "", "", "Missing Q ID", ""]);
        continue;
    }

    const prevHalfId = halfSequenceMap.get(currentHalf) || "";
    let prevActualScore = null; 

    if (prevHalfId) { 
      let prevScoreKey = `${siteId}_${prevHalfId}`;
      if (settings.isIndividualQuestionScore) prevScoreKey += `_${questionId}`;
      
      if (actualScoresMap.has(prevScoreKey)) {
        prevActualScore = actualScoresMap.get(prevScoreKey); 
      }
    }

    let imputedFinalScore = actualCurrentScore;
    let scoreBasis = actualCurrentScore !== null ? "Actual" : "Missing";

    if (actualCurrentScore === null && settings.imputationMethod === "LOCF") {
      if (prevActualScore !== null) {
        imputedFinalScore = prevActualScore; 
        scoreBasis = "Imputed (LOCF)";
      } else {
        scoreBasis = "Missing (No LOCF)";
      }
    }
    
    let changeInScore = "";
    if (imputedFinalScore !== null && prevActualScore !== null) {
      changeInScore = imputedFinalScore - prevActualScore;
    }
    
    outputDataBlock.push([
      prevHalfId,
      prevActualScore === null ? "" : prevActualScore,
      imputedFinalScore === null ? "" : imputedFinalScore,
      scoreBasis,
      changeInScore === null ? "" : changeInScore 
    ]);
  }

  Logger.log(`Writing ${outputDataBlock.length} rows of calculated values to sheet "${settings.targetSheetName}" starting at column ${prevHalfIdColNum}.`);
  if (outputDataBlock.length > 0) {
    sheet.getRange(2, prevHalfIdColNum, outputDataBlock.length, 5).setValues(outputDataBlock);
  }
  SpreadsheetApp.flush(); 
  ui.alert('Success!', `Trend columns for "${settings.scoreName}" on sheet "${settings.targetSheetName}" have been calculated and updated.`, ui.ButtonSet.OK);
}


/**
 * Generates a Site-by-Half score summary table with additional site details.
 */
function generateSiteScoreSummaryTable() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const functionName = 'generateSiteScoreSummaryTable';

  try {
    const sourceSheetNameResponse = ui.prompt(functionName, 'Enter source sheet name (e.g., TransformedData):', ui.ButtonSet.OK_CANCEL);
    if (sourceSheetNameResponse.getSelectedButton() !== ui.Button.OK || !sourceSheetNameResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
    const sourceSheetName = sourceSheetNameResponse.getResponseText().trim();
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) { ui.alert('Error', `Source sheet "${sourceSheetName}" not found.`, ui.ButtonSet.OK); return; }

    const halfColLetter = ui.prompt(functionName, `Enter "Half" column letter in "${sourceSheetName}" (e.g., A):`, ui.ButtonSet.OK_CANCEL).getResponseText()?.trim().toUpperCase();
    if (!halfColLetter) { ui.alert('Operation Cancelled.'); return; }
    const halfColIndex = letterToColumn(halfColLetter);

    const siteIdColLetter = ui.prompt(functionName, `Enter "Site No." column letter in "${sourceSheetName}" (e.g., Q):`, ui.ButtonSet.OK_CANCEL).getResponseText()?.trim().toUpperCase();
    if (!siteIdColLetter) { ui.alert('Operation Cancelled.'); return; }
    const siteIdColIndex = letterToColumn(siteIdColLetter);

    const siteNameColLetter = ui.prompt(functionName, `Enter "Site Name" column letter in "${sourceSheetName}" (e.g., R for Site):`, ui.ButtonSet.OK_CANCEL).getResponseText()?.trim().toUpperCase();
    if (!siteNameColLetter) { ui.alert('Operation Cancelled.'); return; }
    const siteNameColIndex = letterToColumn(siteNameColLetter);

    const brandColLetter = ui.prompt(functionName, `Enter "Brand" column letter in "${sourceSheetName}" (e.g., I):`, ui.ButtonSet.OK_CANCEL).getResponseText()?.trim().toUpperCase();
    if (!brandColLetter) { ui.alert('Operation Cancelled.'); return; }
    const brandColIndex = letterToColumn(brandColLetter);
    
    const cbreSectorColLetter = ui.prompt(functionName, `Enter "CBRE Sector" column letter in "${sourceSheetName}" (e.g., N):`, ui.ButtonSet.OK_CANCEL).getResponseText()?.trim().toUpperCase();
    if (!cbreSectorColLetter) { ui.alert('Operation Cancelled.'); return; }
    const cbreSectorColIndex = letterToColumn(cbreSectorColLetter);

    const scoreColLetter = ui.prompt(functionName, `Enter "Score" column (e.g. Imputed/Final Total Score) letter to summarize in "${sourceSheetName}" (e.g., AC):`, ui.ButtonSet.OK_CANCEL).getResponseText()?.trim().toUpperCase();
    if (!scoreColLetter) { ui.alert('Operation Cancelled.'); return; }
    const scoreColIndex = letterToColumn(scoreColLetter);

    const summarySheetNameResponse = ui.prompt(functionName, 'Enter name for the new summary sheet (default: SiteScoreSummary):', ui.ButtonSet.OK_CANCEL);
    let summarySheetName = SITE_SCORE_SUMMARY_SHEET_NAME; 
    if (summarySheetNameResponse.getSelectedButton() === ui.Button.OK && summarySheetNameResponse.getResponseText()) {
      summarySheetName = summarySheetNameResponse.getResponseText().trim();
    }

    const requiredColIndices = [halfColIndex, siteIdColIndex, siteNameColIndex, brandColIndex, cbreSectorColIndex, scoreColIndex];
    if (requiredColIndices.some(isNaN) || requiredColIndices.some(idx => idx < 1)) {
      ui.alert('Error', 'One or more invalid column letters provided.', ui.ButtonSet.OK); return;
    }
    
    const lastRowSource = sourceSheet.getLastRow();
    if (lastRowSource < 2) { ui.alert('No Data', `Source sheet "${sourceSheetName}" has no data below header.`, ui.ButtonSet.OK); return; }
    
    const maxColNeededInSource = Math.max(...requiredColIndices);
    if (maxColNeededInSource > sourceSheet.getLastColumn()) {
        ui.alert('Error', 'One of specified column letters is beyond actual width of source sheet.', ui.ButtonSet.OK); return;
    }
    const dataRangeValues = sourceSheet.getRange(2, 1, lastRowSource - 1, sourceSheet.getLastColumn()).getValues();

    const allHalfValues = dataRangeValues.map(row => String(row[halfColIndex - 1]).trim()).filter(val => val !== "");
    const uniqueParsedHalves = Array.from(new Set(allHalfValues))
                                  .map(halfStr => parseHalfString(halfStr)) 
                                  .filter(parsed => parsed && (parsed.year !== 0 || parsed.sNum !== 0 || parsed.type !== '')); 
    uniqueParsedHalves.sort((a, b) => (a.year !== b.year) ? a.year - b.year : a.sNum - b.sNum);
    const sortedHalfStrings = uniqueParsedHalves.map(ph => ph.original);

    if (sortedHalfStrings.length === 0) { ui.alert('No Parsable Halves', 'Could not find/parse "Half" values for column headers from the source sheet.', ui.ButtonSet.OK); return;}

    const siteDataMap = new Map(); 
    dataRangeValues.forEach(row => {
      const siteId = String(row[siteIdColIndex - 1]).trim();
      if (!siteId) return; 
      
      const half = String(row[halfColIndex - 1]).trim();
      const scoreValueRaw = row[scoreColIndex - 1];
      const score = (scoreValueRaw === "" || scoreValueRaw === null || isNaN(parseFloat(scoreValueRaw))) ? null : parseFloat(scoreValueRaw);

      const siteName = String(row[siteNameColIndex - 1]).trim();
      const brand = String(row[brandColIndex - 1]).trim();
      const cbreSector = String(row[cbreSectorColIndex - 1]).trim();

      if (!siteDataMap.has(siteId)) {
        siteDataMap.set(siteId, { siteName: siteName, brand: brand, cbreSector: cbreSector, scores: {} });
      }
      if (half && score !== null) { 
          siteDataMap.get(siteId).scores[half] = score;
      }
    });
    
    const uniqueSiteIds = Array.from(siteDataMap.keys()).sort((a,b) => String(a).localeCompare(String(b), undefined, {numeric: true})); 
    if (uniqueSiteIds.length === 0) { ui.alert('No Valid Site IDs', 'Could not find valid Site IDs to create rows.', ui.ButtonSet.OK); return;}

    let summarySheet = ss.getSheetByName(summarySheetName);
    if (summarySheet) {
      const clearResponse = ui.alert('Sheet Exists', `Sheet "${summarySheetName}" already exists. Clear and regenerate?`, ui.ButtonSet.YES_NO);
      if (clearResponse === ui.Button.YES) summarySheet.clearContents();
      else { ui.alert('Operation Cancelled', `Sheet "${summarySheetName}" not modified.`, ui.ButtonSet.OK); return; }
    } else {
      summarySheet = ss.insertSheet(summarySheetName);
    }

    const headers = ["Site No.", "Site Name", "Brand", "CBRE Sector", ...sortedHalfStrings];
    summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    summarySheet.setFrozenRows(1);
    
    const outputMatrix = [];
    uniqueSiteIds.forEach(siteId => {
      const siteDetails = siteDataMap.get(siteId);
      const rowData = [siteId, siteDetails.siteName, siteDetails.brand, siteDetails.cbreSector];
      sortedHalfStrings.forEach(half => {
        const score = siteDetails.scores[half];
        rowData.push(score !== undefined && score !== null ? score : ""); 
      });
      outputMatrix.push(rowData);
    });

    if (outputMatrix.length > 0) {
      summarySheet.getRange(2, 1, outputMatrix.length, headers.length).setValues(outputMatrix);
      if (summarySheet.getLastColumn() > 0) {
        summarySheet.autoResizeColumns(1, summarySheet.getLastColumn());
      }
    }
    ui.alert('Success!', `"${summarySheetName}" generated/updated with ${outputMatrix.length} sites.`, ui.ButtonSet.OK);
  } catch (e) {
    try { logErrorToSheet(functionName, e); } catch (le) { Logger.log("Failed to call logErrorToSheet from TrendAnalysis.gs: " + le.message); }
    ui.alert('Error', `An error occurred in ${functionName}: ${e.message}. Check ScriptErrorLog and Logger.`, ui.ButtonSet.OK);
  }
}

/**
 * Converts a 1-based column index to a letter (e.g., 1 -> A, 27 -> AA).
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Parses a "Half" string (e.g., "S1-AW23") into components for sorting.
 * Returns an object { original: string, year: number, sNum: number, type: string }
 * Returns null if input is not a string or is empty.
 * For unmatchable formats, year/sNum will be 0 and type '', but original is preserved.
 */
function parseHalfString(halfString) {
  if (typeof halfString !== 'string' || halfString.trim() === "") return null;
  const trimmedHalfString = halfString.trim();
  const match = trimmedHalfString.match(/^S(\d+)-([A-Z]{2})(\d{2})$/i); 
  if (match) {
    const sNum = parseInt(match[1], 10);
    const type = match[2].toUpperCase();
    const yearSuffix = parseInt(match[3], 10);
    const year = 2000 + yearSuffix; 
    return { original: trimmedHalfString, year: year, sNum: sNum, type: type };
  }
  Logger.log(`Could not parse half string: "${trimmedHalfString}" using regex. Returning with default year/sNum.`);
  return { original: trimmedHalfString, year: 0, sNum: 0, type: '' }; 
}


/**
 * Generates a lookup table for "Current Half" and "Previous Half".
 * This table is crucial for time-based comparisons like LOCF or score changes.
 */
function generateHalfSequenceTable() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const functionName = 'generateHalfSequenceTable';

  try {
    const sourceSheetNameResponse = ui.prompt(functionName, 'Enter name of sheet containing "Half" identifiers (e.g., TransformedData):', ui.ButtonSet.OK_CANCEL);
    if (sourceSheetNameResponse.getSelectedButton() !== ui.Button.OK || !sourceSheetNameResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
    const sourceSheetName = sourceSheetNameResponse.getResponseText().trim();
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) { ui.alert('Error', `Source sheet "${sourceSheetName}" not found.`, ui.ButtonSet.OK); return; }

    const halfColLetterResponse = ui.prompt(functionName, `Enter column letter for "Half" identifiers in sheet "${sourceSheetName}" (e.g., A):`, ui.ButtonSet.OK_CANCEL);
    if (halfColLetterResponse.getSelectedButton() !== ui.Button.OK || !halfColLetterResponse.getResponseText()) { ui.alert('Operation Cancelled.'); return; }
    const halfColLetter = halfColLetterResponse.getResponseText().trim().toUpperCase();
    const halfColIndex = letterToColumn(halfColLetter); 

    if (isNaN(halfColIndex) || halfColIndex < 1) { ui.alert('Error', `Invalid column letter "${halfColLetter}".`, ui.ButtonSet.OK); return; }

    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) { ui.alert('No Data', `Sheet "${sourceSheetName}" has no data below header in column ${halfColLetter}.`, ui.ButtonSet.OK); return; }

    const halfValuesRange = sourceSheet.getRange(2, halfColIndex, lastRow - 1, 1);
    const halfValues = halfValuesRange.getValues().flat().filter(val => val !== null && val !== undefined && String(val).trim() !== "");

    if (halfValues.length === 0) { ui.alert('No "Half" Values', `No "Half" identifiers found or all blank in column ${halfColLetter} of sheet "${sourceSheetName}".`, ui.ButtonSet.OK); return; }
    
    const uniqueParsedHalves = Array.from(new Set(halfValues.map(h => String(h).trim()))) 
                                  .map(halfStr => parseHalfString(halfStr)) 
                                  .filter(parsed => parsed !== null && (parsed.year !== 0 || parsed.sNum !== 0 || parsed.type !== '')); 

    if (uniqueParsedHalves.length === 0) { 
        ui.alert('No Parsable "Half" Values', 'No "Half" identifiers could be properly parsed (e.g., S1-AW23). Check data format in the source column.', ui.ButtonSet.OK); 
        return; 
    }
    
    uniqueParsedHalves.sort((a, b) => {
      if (a.year !== b.year) return a.year - b.year;
      return a.sNum - b.sNum; 
    });

    let lookupSheet = ss.getSheetByName(HALF_SEQUENCE_SHEET_NAME);
    if (lookupSheet) {
      const clearResponse = ui.alert('Sheet Exists', `Sheet "${HALF_SEQUENCE_SHEET_NAME}" already exists. Clear and regenerate?`, ui.ButtonSet.YES_NO);
      if (clearResponse === ui.Button.YES) lookupSheet.clearContents();
      else { ui.alert('Operation Cancelled', `Sheet "${HALF_SEQUENCE_SHEET_NAME}" not modified.`, ui.ButtonSet.OK); return; }
    } else {
      lookupSheet = ss.insertSheet(HALF_SEQUENCE_SHEET_NAME);
    }

    lookupSheet.getRange("A1").setValue("Current Half").setFontWeight("bold");
    lookupSheet.getRange("B1").setValue("Previous Half").setFontWeight("bold");
    lookupSheet.setFrozenRows(1);

    const outputData = [];
    uniqueParsedHalves.forEach((parsedHalf, i) => {
        const currentHalfOriginal = parsedHalf.original; 
        const previousHalfOriginal = (i > 0) ? uniqueParsedHalves[i-1].original : "";
        outputData.push([currentHalfOriginal, previousHalfOriginal]);
    });
    
    if (outputData.length > 0) {
      lookupSheet.getRange(2, 1, outputData.length, 2).setValues(outputData);
      if (lookupSheet.getLastColumn() > 0) {
        lookupSheet.autoResizeColumns(1, lookupSheet.getLastColumn());
      }
    }
    ui.alert('Success!', `"${HALF_SEQUENCE_SHEET_NAME}" generated/updated with ${outputData.length} unique "Half" periods.`, ui.ButtonSet.OK);
  } catch (e) {
    try { logErrorToSheet(functionName, e); } catch (le) { Logger.log("Failed to call logErrorToSheet from TrendAnalysis.gs: " + le.message); }
    ui.alert('Error', `An error occurred in ${functionName}: ${e.message}. Check ScriptErrorLog and Logger.`, ui.ButtonSet.OK);
  }
}

/**
 * Converts a column letter (e.g., A, AA) to its 1-based column index.
 * @param {string} letter The column letter(s).
 * @return {number} The 1-based column index, or NaN if invalid.
 */
function letterToColumn(letter) {
  if (typeof letter !== 'string' || !/^[A-Z]+$/i.test(letter)) return NaN; 
  let column = 0, length = letter.length;
  letter = letter.toUpperCase(); 
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

