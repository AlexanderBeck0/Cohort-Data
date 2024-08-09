// #####
// A switch statement to route where edits go to.
// #####

/**
 * @typedef {Object} onSheetEditEvent The event when a sheet it edited.
 * @property {ScriptApp.AuthMode} authMode A value from the [ScriptApp.AuthMode](https://developers.google.com/apps-script/reference/script/auth-mode) enum.
 * @property {(string | number | boolean | Date)| undefined} oldValue Cell value prior to the edit, if any. Only available if the edited range is a single cell. Will be undefined if the cell had no previous content.
 * @property {SpreadsheetApp.Range} range A [Range](https://developers.google.com/apps-script/reference/spreadsheet/range) object, representing the cell or range of cells that were edited.
 * @property {SpreadsheetApp.Spreadsheet} source A [Spreadsheet](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet) object, representing the Google Sheets file to which the script is bound.
 * @property {string} triggerUid ID of trigger that produced this event (installable triggers only). Example trigger ID: `4034124084959907503`
 * @property {User} user A [User](https://developers.google.com/apps-script/reference/base/user) object, representing the active user, if available ([depending on a complex set of security restrictions](https://developers.google.com/apps-script/reference/base/session#getActiveUser())).
 * @property {string | number | boolean | Date} value New cell value after the edit. Only available if the edited range is a single cell.
 * @see [Google documentation](https://developers.google.com/apps-script/guides/triggers/events#Google%20Sheets-events)
 */

/**
 * The event handler triggered when anything changes
 * @param {onSheetEditEvent} e The onEdit() event
 */
function onSheetEdit(e) {
  const sheetName = e.range.getSheet().getSheetName();
  switch (sheetName) {
    case "Relative Year Summary":
      Logger.log("Relative Year Summary Change");
      const success = onRYSEdit(e.source, sheetName, e.range.getColumn(), e.range.getRow());
      // Use success to prevent any change on RYS updating growth dashboard
      if (success) {
        const sharedRange = e.source.getSheetByName("Company Growth Dashboard").getRange("A17");
        sharedRange.setValue(e.value); // Set shared range to whatever RYS option was changed to
        SpreadsheetApp.flush();
        refreshGrowthDashboard(e.source);
      }
      break;
    case "Budget Filter":
      Logger.log("Budget Filter Change");
      onBudgetFilterChange(e.source, sheetName, e.range.getColumn(), e.range.getRow());
      break;
    case "Company Growth Dashboard":
      Logger.log("Company Growth Dashboard Change");
      onGrowthDashboardChange(e.source, sheetName, e.range.getColumn(), e.range.getRow());
      break;
    case "Companies With Missing Data":
      Logger.log("Companies With Missing Data Change");
      onMissingDataSheetChange(e.source, sheetName, e.range.getColumn(), e.range.getRow());
      break;
    case "Individual Company Dashboard":
      // Commented out due to it being depricated
      // Logger.log("Individual Company Dashboard Change");
      // onIndividualCompanyDashboardChange(e.source, sheetName, e.range.getColumn(), e.range.getRow(), e.value);
      break;
    default:
      return;
  }
}
