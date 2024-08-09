/**
 * The event handler triggered when the options in Individual Company Dashboard changes
 * @depricated
 * @param {SpreadsheetApp.Spreadsheet} source The spreadsheet being changed.
 * @param {string} sheetName The name of the sheet that was changed. Ignores if the sheet name is not `Individual Company Dashboard`
 * @param {number} columnOfChange The index of the column that was changed. Ignores if the index is not 1.
 * @param {number} rowOfChange The index of the row that changed. Index starts at 1. Will ignore if the row is not within `dashboardOptions` (see function body).
 * @param {string} valueOfChange The value that was changed within `Individual Company Dashboard`. Assumes that the value is a `string` due to there being only one option currently (as of July 18, 2024).
 */
function onIndividualCompanyDashboardChange(source, sheetName, columnOfChange, rowOfChange, valueOfChange) {
    // Ignore all edits that are not within the dropdown
    if (sheetName !== "Individual Company Dashboard") return;
    if (columnOfChange !== 1) return;
  
    const dashboardOptions = Object.freeze({
      Company: 2
    });
  
    if (!Object.values(dashboardOptions).includes(rowOfChange)) return;
  
    Logger.log("WARNING: This code is BROKEN. I do not know why, but I get the error Exception: Service Spreadsheets failed while accessing document with id <Spreadsheet ID>");
    Logger.log("Instead, it is being done soley on the sheet itself using built in features.");
    return;
    const sheet = source.getActiveSheet();
  
    const enableLoadingIndicator = true;
  
    /**
     * @type {SpreadsheetApp.Range}
     */
    let loadingRange;
    /**
     * The status to display to the user if `enableLoadingIndicator` is true
     * @type {string}
     */
    let currentStatus = "Loading...";
    /**
     * Used for the case that there is an error. Instead of just throwing an error, it will hold the error until everything has finished, and then it will throw it at the end.
     * @type {Error}
     */
    let throwError;
    if (enableLoadingIndicator) {
      loadingRange = sheet.getRange("B1");
      loadingRange.setValue(currentStatus);
      SpreadsheetApp.flush();
    }
  
    try {
      updateIndividualCompanyChart(sheet, valueOfChange);
    } catch (error) {
      Logger.log("Caught error: %s", error);
      throwError = error;
      currentStatus = "Failed!";
    }
  
    if (enableLoadingIndicator) {
      // Note: This is SLOW, but it is here for UX. Not a necessary feature.
      // Keep at the bottom of this function. Moving it elsewhere will cause bottlenecks
      currentStatus = currentStatus === "Loading..." ? "Done!" : currentStatus;
      loadingRange.setValue(currentStatus);
      SpreadsheetApp.flush(); // Update the sheet with what it has currently
  
      // Sleep for half a second
      const NUMBER_OF_SLEEP_SECONDS = 0.5;
      Utilities.sleep(NUMBER_OF_SLEEP_SECONDS * 1000);
      loadingRange.clearContent();
      if (throwError !== undefined) throw throwError; // Throw it at the end so it shows up on 
      return; // Return here to force this being the last call in the function
    }
  }
  
  /**
   * @depricated
   * @param {SpreadsheetApp.Sheet} sourceSheet The sheet that contains the Individual Company graphs.
   * @param {string} newCompany The company that the dropdown in Individual Company Dashboard was changed to.
   * @returns {SpreadsheetApp.EmbeddedChart[]} A list of all charts that were changed.
   */
  function updateIndividualCompanyChart(sourceSheet, newCompany) {
    checkVariablesDefined({ sourceSheet, newCompany });
  
    const charts = sourceSheet.getCharts();
    const changedCharts = [];
    charts.forEach(chart => {
      const newTitle = "Revenue vs. Expenses";
      const newSubtitle = newCompany;
      /**
       * @type {UpdateChartOptions}
       */
      const newChartOptions = {
        title: newTitle,
        subtitle: newSubtitle,
      };
      const changedChart = updateChart(sourceSheet, chart, newChartOptions);
      changedCharts.push(changedChart);
    });
  
    return changedCharts;
  }