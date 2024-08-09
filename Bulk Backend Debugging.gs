// ######
// For doing everything in bulk. Not used in the front end at all. Done in the event that something is changed and someone needs to go an recalculate everything.
// ######

function createTriggers() {
  const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(activeSpreadSheet).onFormSubmit().create();
  // ScriptApp.newTrigger('updateNamedRanges').forSpreadsheet(activeSpreadSheet).onChange().create();
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(activeSpreadSheet).onEdit().create();
  ScriptApp.newTrigger('updateCompanyListOnForm').timeBased().atHour(0).onMonthDay(1).create();
}

/**
 * Deletes all triggers and then calls `createTrigger()`
 * 
 * **WARNING** Will delete **all** triggers
 */
function refreshTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  createTriggers();
}

/**
 * Recalculates all the relative year summary values
 */
function recalculateAllRelativeYearSummaryValues() {
  updateYearValues();
}