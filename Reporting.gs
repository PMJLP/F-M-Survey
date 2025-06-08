/**
 * @OnlyCurrentDoc
 */

// --- Configuration ---
const REPORT_SHEET_PREFIX = "Report -";
const THEME_ANALYSIS_PREFIX = "Theme Analysis -";
const DATA_SHEETS = {
  SCORES_SUMMARY: "JLSiteScoreSummary",
  TRANSFORMED: "TransformedData",
  REAL_ESTATE: "JLPRealEstate",
  CBRE_INFO: "CBRESiteInfo",
  QUESTION_THEMES: "QuestionThemes"
};

const REPORT_SHEET_NAMES = {
  LATEST_SCORES: "Latest Scores",
  SCORE_CHANGES: "Score Changes",
  CONSISTENT_PERFORMERS: "Consistent Performers",
  SERVICE_TRENDS: "Service Trends",
  CONSOLIDATED_SUPPORT: "Consolidated Support Needs",
  BY_CBRE_REGION: "By CBRE Region",
  FOCUS_AREAS: "Focus Areas"
};


// --- Helper Functions ---

function getOrCreateReportSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function applyConditionalFormatting_(range, type) {
  if (!range) return;
  const sheet = range.getSheet();
  const rules = sheet.getConditionalFormatRules();
  let cfRuleBuilder;

  switch (type) {
    case "score":
      cfRuleBuilder = SpreadsheetApp.newConditionalFormatRule().setGradientMaxpoint('#63BE7B').setGradientMidpointWithValue('#FFEB84', SpreadsheetApp.InterpolationType.NUMBER, '1').setGradientMinpoint('#F8696B');
      break;
    case "trend":
    case "change":
      cfRuleBuilder = SpreadsheetApp.newConditionalFormatRule().setGradientMaxpoint('#63BE7B').setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0').setGradientMinpoint('#F8696B');
      break;
    case "level":
      const levelColors = { "High": "#D82121", "Moderate-High": "#F8696B", "Moderate": "#FFEB84", "Moderate-Low": "#C6E0B4", "Low": "#63BE7B", "Monitor": "#BFBFBF" };
      for (const level in levelColors) {
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(level).setBackground(levelColors[level]).setRanges([range]).build());
      }
      sheet.setConditionalFormatRules(rules);
      return;
  }
  if (cfRuleBuilder) {
    rules.push(cfRuleBuilder.setRanges([range]).build());
  }
  sheet.setConditionalFormatRules(rules);
}

function _getSiteTypeInfoMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const realEstateSheet = ss.getSheetByName(DATA_SHEETS.REAL_ESTATE);
  const cbreInfoSheet = ss.getSheetByName(DATA_SHEETS.CBRE_INFO);
  if (!realEstateSheet || !cbreInfoSheet) throw new Error(`Dependent sheets for site type classification are missing.`);
  
  const typeInfoMap = new Map();
  const realEstateData = realEstateSheet.getDataRange().getValues().slice(1);
  const cbreInfoData = cbreInfoSheet.getDataRange().getValues().slice(1);

  realEstateData.forEach(row => {
    const siteNo = String(row[0]).trim();
    if (siteNo) {
      if (!typeInfoMap.has(siteNo)) typeInfoMap.set(siteNo, { jlpFunc: '', cbreComm: '' });
      typeInfoMap.get(siteNo).jlpFunc = String(row[5]).trim().toLowerCase();
    }
  });
  cbreInfoData.forEach(row => {
    const siteNo = String(row[1]).trim();
    if (siteNo) {
      if (!typeInfoMap.has(siteNo)) typeInfoMap.set(siteNo, { jlpFunc: '', cbreComm: '' });
      typeInfoMap.get(siteNo).cbreComm = String(row[5]).trim().toLowerCase();
    }
  });
  return typeInfoMap;
}

function _classifySiteType_(jlpFunc, cbreComm, brand) {
  if (jlpFunc.includes("cdh") || jlpFunc.includes("customer delivery hub") || jlpFunc.includes("central distribution hub") || cbreComm.includes("cdh")) {
    return "CDH";
  }
  if ((jlpFunc && jlpFunc !== "waitrose") || (!jlpFunc && cbreComm && cbreComm !== "waitrose" && !cbreComm.includes("cdh")) || (!jlpFunc && !cbreComm && brand === "John Lewis")) {
    return "Shop";
  }
  return null;
}

function getSiteData_(siteTypeToFilter) {
  const functionName = `getSiteData_ (${siteTypeToFilter})`;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const jlScoresSheet = ss.getSheetByName(DATA_SHEETS.SCORES_SUMMARY);
    if (!jlScoresSheet) throw new Error(`Sheet "${DATA_SHEETS.SCORES_SUMMARY}" not found.`);
    
    const typeInfoMap = _getSiteTypeInfoMap_();
    const jlScoresData = jlScoresSheet.getDataRange().getValues();
    const headers = jlScoresData.shift();
    const sites = [];

    const siteNoCol = headers.indexOf("Site No.");
    const siteNameCol = headers.indexOf("Site Name");
    const brandCol = headers.indexOf("Brand");
    if (siteNoCol === -1 || brandCol === -1 || siteNameCol === -1) throw new Error(`Required columns not found in score summary.`);

    jlScoresData.forEach(row => {
      const siteNo = String(row[siteNoCol]).trim();
      const brand = String(row[brandCol]).trim();
      const typeInfo = typeInfoMap.get(siteNo) || { jlpFunc: '', cbreComm: '' };
      const currentSiteType = _classifySiteType_(typeInfo.jlpFunc, typeInfo.cbreComm, brand);
      if (brand === "John Lewis" && siteNo && currentSiteType === siteTypeToFilter) {
        const siteObject = {};
        headers.forEach((header, i) => {
          const key = header.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
          if (key) {
            const val = parseFloat(row[i]);
            siteObject[key] = isNaN(val) ? row[i] : val;
          }
        });
        siteObject.type = currentSiteType;
        siteObject.siteName = row[siteNameCol];
        siteObject.siteNo = siteNo;
        sites.push(siteObject);
      }
    });
    return sites;
  } catch (e) {
    logErrorToSheet(functionName, e);
    SpreadsheetApp.getUi().alert('Error', `An error occurred in ${functionName}: ${e.message}.`);
    return [];
  }
}

function _mapThemeToCategory(themeName) {
    const theme = themeName.toLowerCase();
    if (theme.includes("portal")) return "Systems";
    if (theme.includes("cleaning")) return "Cleaning";
    if (theme.includes("cbre")) return "Service";
    return null;
}

function _calculateAllSiteThemeMetrics() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const themesSheet = ss.getSheetByName(DATA_SHEETS.QUESTION_THEMES);
    const transformedSheet = ss.getSheetByName(DATA_SHEETS.TRANSFORMED);
    if (!themesSheet || !transformedSheet) throw new Error("Required data sheets for themes not found.");

    const questionToThemeMap = new Map(themesSheet.getDataRange().getValues().slice(1).map(r => [r[0], r[2]]));
    const siteTypeInfoMap = _getSiteTypeInfoMap_();
    
    const transformedData = transformedSheet.getDataRange().getValues();
    const headers = transformedData.shift();
    const siteNoCol = headers.indexOf('Site No.');
    const siteNameCol = headers.indexOf('Site');
    const brandCol = headers.indexOf('Brand');
    const questionIdCol = headers.indexOf('Question ID');
    const scoreCol = headers.indexOf('Score');
    const halfCol = headers.indexOf('Half');

    const siteMetrics = {}; 

    transformedData.forEach(row => {
        const brand = String(row[brandCol]).trim();
        if (brand !== 'John Lewis') return;

        const siteNo = String(row[siteNoCol]).trim();
        const siteName = String(row[siteNameCol]).trim();
        const questionId = String(row[questionIdCol]).trim();
        const score = parseFloat(row[scoreCol]);
        const half = String(row[halfCol]).trim();
        
        if (!siteNo || !questionId || isNaN(score) || !half) return;

        const theme = questionToThemeMap.get(questionId);
        if (!theme) return;
        
        const typeInfo = siteTypeInfoMap.get(siteNo) || { jlpFunc: '', cbreComm: '' };
        const siteType = _classifySiteType_(typeInfo.jlpFunc, typeInfo.cbreComm, brand);
        if (!siteType) return;

        if (!siteMetrics[siteNo]) siteMetrics[siteNo] = { siteName: siteName, siteType: siteType, themes: {} };
        if (!siteMetrics[siteNo].themes[theme]) siteMetrics[siteNo].themes[theme] = { scoresByHalf: {} };
        if (!siteMetrics[siteNo].themes[theme].scoresByHalf[half]) siteMetrics[siteNo].themes[theme].scoresByHalf[half] = { sum: 0, count: 0 };
        
        siteMetrics[siteNo].themes[theme].scoresByHalf[half].sum += score;
        siteMetrics[siteNo].themes[theme].scoresByHalf[half].count++;
    });

    for (const siteNo in siteMetrics) {
        for (const theme in siteMetrics[siteNo].themes) {
            const themeData = siteMetrics[siteNo].themes[theme];
            const allScores = [];
            let latestScore = null, prevScore = null, trendSum = 0;
            const sortedHalves = Object.keys(themeData.scoresByHalf).sort();
            
            sortedHalves.forEach((half, i) => {
                const avg = themeData.scoresByHalf[half].sum / themeData.scoresByHalf[half].count;
                themeData.scoresByHalf[half].average = avg;
                allScores.push(avg);
                if (i === sortedHalves.length - 1) latestScore = avg;
                if (i === sortedHalves.length - 2) prevScore = avg;
                if (i > 0) trendSum += (avg - themeData.scoresByHalf[sortedHalves[i-1]].average);
            });

            themeData.latestScore = latestScore;
            themeData.scoreChange = (latestScore !== null && prevScore !== null) ? latestScore - prevScore : null;
            themeData.averageScore = allScores.reduce((a, b) => a + b, 0) / allScores.length;
            themeData.surveyCount = allScores.length;
            themeData.trend = (sortedHalves.length > 1) ? trendSum : null;
        }
    }
    return siteMetrics;
}


// --- Menu-Facing Wrapper Functions ---
function generateLatestSurveyReport_Shops() { _generateLatestSurveyReport("Shop"); }
function generateScoreChangesReport_Shops() { _generateScoreChangesReport("Shop"); }
function generateConsistentPerformersReport_Shops() { _generateConsistentPerformersReport("Shop"); }
function generateServiceTrendsReport_Shops() { _generateServiceTrendsReport("Shop"); }
function generateReportByCBRERegion_Shops() { _generateReportByCBRERegion("Shop"); }
function generateLatestSurveyReport_CDHs() { _generateLatestSurveyReport("CDH"); }
function generateScoreChangesReport_CDHs() { _generateScoreChangesReport("CDH"); }
function generateConsistentPerformersReport_CDHs() { _generateConsistentPerformersReport("CDH"); }
function generateServiceTrendsReport_CDHs() { _generateServiceTrendsReport("CDH"); }
function generateReportByCBRERegion_CDHs() { _generateReportByCBRERegion("CDH"); }
function generateConsolidatedSupportNeedsReport() { _generateConsolidatedSupportNeedsReport(); }
function generateFocusReport() { _generateFocusReport(); }
function generateAllThemeBreakdownReports() { _generateAllThemeBreakdownReports(); }

function generateAllStandardReports() {
    const ui = SpreadsheetApp.getUi();
    const functionName = 'generateAllStandardReports';
    ui.showModalDialog(HtmlService.createHtmlOutput("<p>Generating all standard reports...<br>This may take a moment.</p>"), "Processing...");
    try {
        _generateLatestSurveyReport("Shop");
        _generateLatestSurveyReport("CDH");
        _generateScoreChangesReport("Shop");
        _generateScoreChangesReport("CDH");
        _generateConsistentPerformersReport("Shop");
        _generateConsistentPerformersReport("CDH");
        _generateServiceTrendsReport("Shop");
        _generateServiceTrendsReport("CDH");
        _generateReportByCBRERegion("Shop");
        _generateReportByCBRERegion("CDH");
        ui.alert('Success!', 'All standard reports have been generated/refreshed.', ui.ButtonSet.OK);
    } catch(e) {
        logErrorToSheet(functionName, e);
        ui.alert('An error occurred while generating reports. Please check the ScriptErrorLog sheet.');
    }
}


// --- Worker Functions ---

function _generateLatestSurveyReport(siteType) {
  const reportSheetName = `${REPORT_SHEET_PREFIX} ${REPORT_SHEET_NAMES.LATEST_SCORES} (${siteType}s)`;
  const sheet = getOrCreateReportSheet_(reportSheetName);
  const sites = getSiteData_(siteType).filter(s => s.s4ss25 !== null);
  sheet.appendRow([`Top/Bottom 6 ${siteType}s - Latest Survey (S4)`]).getRange("A1").setFontWeight("bold");
  sheet.appendRow(['']);
  if (sites.length < 1) { sheet.appendRow(["No sites with latest scores found."]); return; }
  
  const headers = ["Rank", "Site Name", "Score"];
  sheet.appendRow(["Top 6 Performing"]).getRange(sheet.getLastRow(), 1).setFontWeight("bold");
  sheet.appendRow(headers);
  const topStartRow = sheet.getLastRow() + 1;
  const topData = sites.sort((a,b) => b.s4ss25 - a.s4ss25).slice(0,6).map((s,i) => [i+1, s.sitename, s.s4ss25]);
  if (topData.length > 0) {
    const range = sheet.getRange(topStartRow, 1, topData.length, 3);
    range.setValues(topData).setNumberFormat('0.00');
    applyConditionalFormatting_(range.offset(0, 2, topData.length, 1), 'score');
  }
  
  sheet.appendRow(['']);
  sheet.appendRow(["Bottom 6 Performing"]).getRange(sheet.getLastRow(), 1).setFontWeight("bold");
  sheet.appendRow(headers);
  const bottomStartRow = sheet.getLastRow() + 1;
  const bottomData = sites.sort((a,b) => a.s4ss25 - b.s4ss25).slice(0,6).map((s,i) => [i+1, s.sitename, s.s4ss25]);
  if (bottomData.length > 0) {
    const range = sheet.getRange(bottomStartRow, 1, bottomData.length, 3);
    range.setValues(bottomData).setNumberFormat('0.00');
    applyConditionalFormatting_(range.offset(0, 2, bottomData.length, 1), 'score');
  }
  sheet.autoResizeColumns(1, 3);
}

function _generateScoreChangesReport(siteType) {
  const reportSheetName = `${REPORT_SHEET_PREFIX} ${REPORT_SHEET_NAMES.SCORE_CHANGES} (${siteType}s)`;
  const sheet = getOrCreateReportSheet_(reportSheetName);
  const sites = getSiteData_(siteType).map(s => ({...s, change: (s.s4ss25 !== null && s.s3aw24 !== null) ? s.s4ss25 - s.s3aw24 : null})).filter(s => s.change !== null);
  sheet.appendRow([`Top/Bottom 6 ${siteType}s - Score Changes (S3 to S4)`]).getRange("A1").setFontWeight("bold");
  sheet.appendRow(['']);
  if (sites.length < 1) { sheet.appendRow(["No sites with comparable scores found."]); return; }

  const headers = ["Rank", "Site Name", "Change"];
  sheet.appendRow(["Top 6 Improvers"]).getRange(sheet.getLastRow(), 1).setFontWeight("bold");
  sheet.appendRow(headers);
  const topStartRow = sheet.getLastRow() + 1;
  const topData = sites.sort((a,b) => b.change - a.change).slice(0,6).map((s,i) => [i+1, s.sitename, s.change]);
  if (topData.length > 0) {
    const range = sheet.getRange(topStartRow, 1, topData.length, 3);
    range.setValues(topData).setNumberFormat('0.00');
    applyConditionalFormatting_(range.offset(0, 2, topData.length, 1), 'change');
  }

  sheet.appendRow(['']);
  sheet.appendRow(["Top 6 Decliners"]).getRange(sheet.getLastRow(), 1).setFontWeight("bold");
  sheet.appendRow(headers);
  const bottomStartRow = sheet.getLastRow() + 1;
  const bottomData = sites.sort((a,b) => a.change - b.change).slice(0,6).map((s,i) => [i+1, s.sitename, s.change]);
  if (bottomData.length > 0) {
    const range = sheet.getRange(bottomStartRow, 1, bottomData.length, 3);
    range.setValues(bottomData).setNumberFormat('0.00');
    applyConditionalFormatting_(range.offset(0, 2, bottomData.length, 1), 'change');
  }
  sheet.autoResizeColumns(1, 3);
}

function _generateConsistentPerformersReport(siteType) {
  const reportSheetName = `${REPORT_SHEET_PREFIX} ${REPORT_SHEET_NAMES.CONSISTENT_PERFORMERS} (${siteType}s)`;
  const sheet = getOrCreateReportSheet_(reportSheetName);
  const sites = getSiteData_(siteType).filter(s => s.average !== null);
  sheet.appendRow([`Top/Bottom 6 Consistently Performing ${siteType}s`]).getRange("A1").setFontWeight("bold");
  sheet.appendRow(['']);
  if (sites.length < 1) { sheet.appendRow(["No sites with average scores found."]); return; }
  
  const headers = ["Rank", "Site Name", "Average Score"];
  sheet.appendRow(["Top 6 Consistent Performers"]).getRange(sheet.getLastRow(), 1).setFontWeight("bold");
  sheet.appendRow(headers);
  const topStartRow = sheet.getLastRow() + 1;
  const topData = sites.sort((a,b) => b.average - a.average).slice(0,6).map((s,i) => [i+1, s.sitename, s.average]);
  if (topData.length > 0) {
    const range = sheet.getRange(topStartRow, 1, topData.length, 3);
    range.setValues(topData).setNumberFormat('0.00');
    applyConditionalFormatting_(range.offset(0, 2, topData.length, 1), 'score');
  }

  sheet.appendRow(['']);
  sheet.appendRow(["Bottom 6 Consistent Performers"]).getRange(sheet.getLastRow(), 1).setFontWeight("bold");
  sheet.appendRow(headers);
  const bottomStartRow = sheet.getLastRow() + 1;
  const bottomData = sites.sort((a,b) => a.average - b.average).slice(0,6).map((s,i) => [i+1, s.sitename, s.average]);
  if (bottomData.length > 0) {
    const range = sheet.getRange(bottomStartRow, 1, bottomData.length, 3);
    range.setValues(bottomData).setNumberFormat('0.00');
    applyConditionalFormatting_(range.offset(0, 2, bottomData.length, 1), 'score');
  }
  sheet.autoResizeColumns(1, 3);
}

function _generateServiceTrendsReport(siteType) {
  const reportSheetName = `${REPORT_SHEET_PREFIX} ${REPORT_SHEET_NAMES.SERVICE_TRENDS} (${siteType}s)`;
  const sheet = getOrCreateReportSheet_(reportSheetName);
  const sites = getSiteData_(siteType).filter(s => s.totaltrend !== null);
  sheet.appendRow([`Top/Bottom 6 ${siteType}s - F&M Service Trends`]).getRange("A1").setFontWeight("bold");
  sheet.appendRow(['']);
  if (sites.length < 1) { sheet.appendRow(["No sites with trend data found."]); return; }

  const headers = ["Rank", "Site Name", "Total Trend Score"];
  sheet.appendRow(["Top 6 Positive Trends"]).getRange(sheet.getLastRow(), 1).setFontWeight("bold");
  sheet.appendRow(headers);
  const topStartRow = sheet.getLastRow() + 1;
  const topData = sites.sort((a,b) => b.totaltrend - a.totaltrend).slice(0,6).map((s,i) => [i+1, s.sitename, s.totaltrend]);
  if (topData.length > 0) {
    const range = sheet.getRange(topStartRow, 1, topData.length, 3);
    range.setValues(topData).setNumberFormat('0.00');
    applyConditionalFormatting_(range.offset(0, 2, topData.length, 1), 'trend');
  }

  sheet.appendRow(['']);
  sheet.appendRow(["Top 6 Negative Trends"]).getRange(sheet.getLastRow(), 1).setFontWeight("bold");
  sheet.appendRow(headers);
  const bottomStartRow = sheet.getLastRow() + 1;
  const bottomData = sites.sort((a,b) => a.totaltrend - b.totaltrend).slice(0,6).map((s,i) => [i+1, s.sitename, s.totaltrend]);
  if (bottomData.length > 0) {
    const range = sheet.getRange(bottomStartRow, 1, bottomData.length, 3);
    range.setValues(bottomData).setNumberFormat('0.00');
    applyConditionalFormatting_(range.offset(0, 2, bottomData.length, 1), 'trend');
  }
  sheet.autoResizeColumns(1, 3);
}

function _generateReportByCBRERegion(siteType) {
  const functionName = `_generateReportByCBRERegion (${siteType})`;
  try {
      const sheetName = `${REPORT_SHEET_PREFIX} ${REPORT_SHEET_NAMES.BY_CBRE_REGION} (${siteType}s)`;
      const reportSheet = getOrCreateReportSheet_(sheetName);
      const sites = getSiteData_(siteType);

      reportSheet.appendRow([`F&M Performance by CBRE Region - ${siteType}s`]).getRange("A1").setFontWeight("bold");
      reportSheet.appendRow(['']);

      if (sites.length === 0) {
        reportSheet.appendRow([`No ${siteType} data found.`]);
        return;
      }

      const regionalData = {}; 
      sites.forEach(site => {
        const region = site.cbresector || "Unassigned";
        if (!regionalData[region]) {
          regionalData[region] = { totalS4: 0, countS4: 0, totalAvgScore: 0, sumSurveyCountsForAvg: 0, totalTrend: 0, countTrend: 0, siteCount: 0 };
        }
        regionalData[region].siteCount++;
        if (site.s4ss25 !== null) { regionalData[region].totalS4 += site.s4ss25; regionalData[region].countS4++; }
        const surveyCount = [site.s1aw23, site.s2ss24, site.s3aw24, site.s4ss25].filter(s => s !== null).length;
        if (site.average !== null && surveyCount > 0) { 
          regionalData[region].totalAvgScore += (site.average * surveyCount); 
          regionalData[region].sumSurveyCountsForAvg += surveyCount;
        }
        if (site.totaltrend !== null) { regionalData[region].totalTrend += site.totaltrend; regionalData[region].countTrend++; }
      });

      const reportHeaders = ["CBRE Region/Sector", `Avg S4 Score`, `Overall Weighted Avg Score`, `Avg Total Trend`, `Number of Sites`];
      reportSheet.appendRow(reportHeaders).getRange(reportSheet.getLastRow(), 1, 1, reportHeaders.length).setFontWeight("bold");
      
      const dataStartRow = reportSheet.getLastRow() + 1;
      const sortedRegions = Object.keys(regionalData).sort();
      const outputData = [];
      for (const region of sortedRegions) { 
        const data = regionalData[region];
        const avgS4 = data.countS4 > 0 ? (data.totalS4 / data.countS4) : "N/A";
        const overallAvg = data.sumSurveyCountsForAvg > 0 ? (data.totalAvgScore / data.sumSurveyCountsForAvg) : "N/A";
        const avgTrend = data.countTrend > 0 ? (data.totalTrend / data.countTrend) : "N/A";
        outputData.push([region, avgS4, overallAvg, avgTrend, data.siteCount]);
      }

      if (outputData.length > 0) {
          const dataRange = reportSheet.getRange(dataStartRow, 1, outputData.length, outputData[0].length);
          dataRange.setValues(outputData);
          reportSheet.getRange(dataStartRow, 2, outputData.length, 3).setNumberFormat('0.00');
      }

      const dataRangeForChart = reportSheet.getRange(dataStartRow-1, 1, outputData.length+1, 2);
      const chartBuilder = reportSheet.newChart().setChartType(Charts.ChartType.BAR).addRange(dataRangeForChart).setOption('title', `Average Scores by CBRE Region for ${siteType}s`).setPosition(2, 6, 0, 0);
      reportSheet.insertChart(chartBuilder.build());

      reportSheet.autoResizeColumns(1, reportHeaders.length);
  } catch (e) {
      logErrorToSheet(functionName, e);
  }
}

function _generateConsolidatedSupportNeedsReport() {
    const functionName = '_generateConsolidatedSupportNeedsReport';
    const ui = SpreadsheetApp.getUi();
    try {
        const reportSheetName = `${REPORT_SHEET_PREFIX} ${REPORT_SHEET_NAMES.CONSOLIDATED_SUPPORT}`;
        const reportSheet = getOrCreateReportSheet_(reportSheetName);
        const allSites = [...getSiteData_("Shop"), ...getSiteData_("CDH")];
        const headers = ["Site Type", "Site No.", "Site Name", "Support Need Level", "Reasoning Notes"];
        reportSheet.appendRow([REPORT_SHEET_NAMES.CONSOLIDATED_SUPPORT]).getRange(1, 1, 1, headers.length).merge().setFontWeight('bold').setHorizontalAlignment('center');
        reportSheet.appendRow(['']);
        reportSheet.appendRow(headers).getRange(reportSheet.getLastRow(), 1, 1, headers.length).setFontWeight("bold");

        const supportLevelOrder = {"High": 1, "Moderate-High": 2, "Moderate": 3, "Moderate-Low": 4, "Monitor": 5, "Low": 6 };
        
        const allSitesData = allSites.map(site => {
            let reasoning = [];
            const s4score = site.s4ss25, s3score = site.s3aw24, avgScore = site.average, trend = site.totaltrend;
            const surveyCount = [site.s1aw23, site.s2ss24, site.s3aw24, site.s4ss25].filter(s => s !== null).length;
            
            if (s4score !== null && s4score <= 15) reasoning.push("Low S4 Score");
            if (s3score !== null && s4score !== null && (s4score - s3score) < -3) reasoning.push("Significant Drop S3-S4");
            if (avgScore !== null && avgScore <= 16) reasoning.push(surveyCount > 1 ? "Low Avg Score (multiple surveys)" : "Low Avg Score");
            if (trend !== null && trend < 0) reasoning.push("Negative Trend");
            
            let supportNeed = "Monitor";
            if (reasoning.length >= 3) supportNeed = "High";
            else if (reasoning.length >= 2) supportNeed = "Moderate-High";
            else if (reasoning.length === 1) supportNeed = "Moderate-Low";
            else if (s4score !== null && s4score >= 20 && (trend === null || trend >= 0)) supportNeed = "Low";
            
            return {
                siteType: site.type, siteNo: site.siteno, siteName: site.sitename,
                supportNeedLevelText: supportNeed, reasoningNotesText: reasoning.join("; ") || "General performance."
            };
        }).sort((a,b) => (supportLevelOrder[a.supportNeedLevelText] || 99) - (supportLevelOrder[b.supportNeedLevelText] || 99));

        const outputRows = allSitesData.map(s => [s.siteType, s.siteNo, s.siteName, s.supportNeedLevelText, s.reasoningNotesText]);
        
        if (outputRows.length > 0) {
            const dataStartRow = reportSheet.getLastRow() + 1;
            const dataRange = reportSheet.getRange(dataStartRow, 1, outputRows.length, headers.length);
            dataRange.setValues(outputRows);
            applyConditionalFormatting_(reportSheet.getRange(dataStartRow, 4, outputRows.length, 1), "level");
        }
        reportSheet.autoResizeColumns(1, headers.length);
    } catch (e) {
        logErrorToSheet(functionName, e);
    }
}

function _generateFocusReport() {
    const functionName = '_generateFocusReport';
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(HtmlService.createHtmlOutput("<p>Generating Focus Areas Report...</p>").setWidth(350).setHeight(100), "Processing");

    try {
        const reportSheetName = `${REPORT_SHEET_PREFIX} ${REPORT_SHEET_NAMES.FOCUS_AREAS}`;
        const reportSheet = getOrCreateReportSheet_(reportSheetName);
        
        const allSites = [...getSiteData_("Shop"), ...getSiteData_("CDH")];
        const supportLevelOrder = {"High": 1, "Moderate-High": 2, "Moderate": 3, "Moderate-Low": 4, "Monitor": 5, "Low": 6 };
        
        const allSitesData = allSites.map(site => {
            let reasoning = [];
            const s4score = site.s4ss25, s3score = site.s3aw24, avgScore = site.average, trend = site.totaltrend;
            const surveyCount = [site.s1aw23, site.s2ss24, site.s3aw24, site.s4ss25].filter(s => s !== null).length;
            if (s4score !== null && s4score <= 15) reasoning.push("Low S4 Score");
            if (s3score !== null && s4score !== null && (s4score - s3score) < -3) reasoning.push("Significant Drop S3-S4");
            if (avgScore !== null && avgScore <= 16) reasoning.push(surveyCount > 1 ? "Low Avg Score (multiple surveys)" : "Low Avg Score");
            if (trend !== null && trend < 0) reasoning.push("Negative Trend");
            
            let supportNeed = "Monitor";
            if (reasoning.length >= 3) supportNeed = "High";
            else if (reasoning.length === 2) supportNeed = "Moderate-High";
            else if (reasoning.length === 1) supportNeed = "Moderate-Low";
            else if (s4score !== null && s4score >= 20 && (trend === null || trend >= 0)) supportNeed = "Low";
            
            return {
                siteType: site.type, siteNo: site.siteno, siteName: site.sitename,
                supportNeedLevelText: supportNeed, reasoningNotesText: reasoning.join("; ") || "General performance."
            };
        }).sort((a,b) => (supportLevelOrder[a.supportNeedLevelText] || 99) - (supportLevelOrder[b.supportNeedLevelText] || 99));

        const siteCategoryScores = _calculateAllSiteThemeMetrics();

        const title = "F&M Focus Areas (Support Needs vs. Theme Performance)";
        const headers = ["Site Type", "Site No.", "Site Name", "Support Need", "Systems Score", "Cleaning Score", "Service Score", "Reasoning"];
        reportSheet.appendRow([title]).getRange(1, 1, 1, headers.length).merge().setFontWeight('bold').setHorizontalAlignment('center');
        reportSheet.appendRow(['']);
        reportSheet.appendRow(headers).getRange(reportSheet.getLastRow(), 1, 1, headers.length).setFontWeight("bold");

        const outputRows = allSitesData.map(site => {
            const scores = siteCategoryScores[site.siteNo] || {};
            const systemsScore = scores.Systems ? scores.Systems.sum / scores.Systems.count : null;
            const cleaningScore = scores.Cleaning ? scores.Cleaning.sum / scores.Cleaning.count : null;
            const serviceScore = scores.Service ? scores.Service.sum / scores.Service.count : null;
            return [site.siteType, site.siteNo, site.siteName, site.supportNeedLevelText, systemsScore, cleaningScore, serviceScore, site.reasoningNotesText];
        });

        if (outputRows.length > 0) {
            const dataStartRow = reportSheet.getLastRow() + 1;
            const dataRange = reportSheet.getRange(dataStartRow, 1, outputRows.length, headers.length);
            dataRange.setValues(outputRows);
            dataRange.offset(0, 4, outputRows.length, 3).setNumberFormat('0.00');

            applyConditionalFormatting_(reportSheet.getRange(dataStartRow, 4, outputRows.length, 1), "level");
            applyConditionalFormatting_(reportSheet.getRange(dataStartRow, 5, outputRows.length, 1), "score");
            applyConditionalFormatting_(reportSheet.getRange(dataStartRow, 6, outputRows.length, 1), "score");
            applyConditionalFormatting_(reportSheet.getRange(dataStartRow, 7, outputRows.length, 1), "score");
        }
        reportSheet.autoResizeColumns(1, headers.length);
        ui.alert('Success!', `Report '${reportSheet.getName()}' has been generated.`, ui.ButtonSet.OK);
    } catch (e) {
        logErrorToSheet(functionName, e);
    }
}

function _generateAllThemeBreakdownReports() { /* To be implemented */ }
