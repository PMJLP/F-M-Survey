/**
 * @OnlyCurrentDoc
 */
function onOpen(e) { 
  const ui = SpreadsheetApp.getUi();
  try {
    const mainMenu = ui.createMenu('F&M Survey Tools');

    // Sub-menu for Data Preparation
    mainMenu.addSubMenu(ui.createMenu('1. Data Preparation')
      .addItem('Generate Half Sequence Table', 'generateHalfSequenceTable')
      .addItem('Generate Site Score Summary Table', 'generateSiteScoreSummaryTable')
      .addSeparator()
      .addItem('Setup/Add Trend Columns', 'setupTrendColumnsWithImputation')
      .addItem('Quick Refresh Trend Data', 'quickRefreshTrendData'));

    // Sub-menu for Generating Reports
    mainMenu.addSubMenu(ui.createMenu('2. Generate Reports')
      .addItem('Generate ALL Standard Reports', 'generateAllStandardReports')
      .addSeparator()
      .addSubMenu(ui.createMenu('Standard Reports (by Site Type)')
        .addSubMenu(ui.createMenu('Shops')
          .addItem('Latest Scores', 'generateLatestSurveyReport_Shops')
          .addItem('Score Changes', 'generateScoreChangesReport_Shops')
          .addItem('Consistent Performers', 'generateConsistentPerformersReport_Shops')
          .addItem('Service Trends', 'generateServiceTrendsReport_Shops')
          .addItem('Performance by CBRE Region', 'generateReportByCBRERegion_Shops'))
        .addSubMenu(ui.createMenu('CDHs')
          .addItem('Latest Scores', 'generateLatestSurveyReport_CDHs')
          .addItem('Score Changes', 'generateScoreChangesReport_CDHs')
          .addItem('Consistent Performers', 'generateConsistentPerformersReport_CDHs')
          .addItem('Service Trends', 'generateServiceTrendsReport_CDHs')
          .addItem('Performance by CBRE Region', 'generateReportByCBRERegion_CDHs')))
      .addSeparator()
      .addSubMenu(ui.createMenu('Consolidated & Themed Reports')
        .addItem('Generate Consolidated Support Needs Report', 'generateConsolidatedSupportNeedsReport')
        .addItem('Generate Focus Areas Report', 'generateFocusReport')
        .addSeparator()
        .addItem('Generate ALL Consolidated Theme Reports', 'generateAllThemeBreakdownReports')));

    // Sub-menu for Utilities
    mainMenu.addSubMenu(ui.createMenu('Utilities')
      .addItem('Export Sheet Data for AI', 'initiateExportProcessMenu')
      .addSeparator()
      .addItem('Delete Old Reports', 'deleteOldReports')); 
      
    mainMenu.addToUi();
  } catch (err) {
    try { logErrorToSheet("onOpen", err); } catch (e) { Logger.log("Critical Error in onOpen: " + err.message); }
  }
}
