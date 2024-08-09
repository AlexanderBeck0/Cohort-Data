function onOpen() {
    const ui = SpreadsheetApp.getUi();
  
    ui.createMenu("COMPANY NAME")
      .addItem("Refresh relative year summary", 'updateYearValues')
      .addItem("Refresh company list", "recalculateCompanyList")
      .addItem("Refresh raw data", "refreshRawData")
      .addToUi();
  }
  
  /**
   * Recalculates all the differences in Raw Data
   */
  function recalculateAllDifferences() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName("Raw Data");
    const relavantDataRange = rawDataSheet.getRange(String.fromCharCode(rawDataColumns.Revenue + 65) + "2:" + String.fromCharCode(rawDataColumns.Difference + 65));
    const values = relavantDataRange.getValues();
    const newDifferences = values.map(row => [row[0], row[1], calculateDifference(row[0], row[1])]);
    relavantDataRange.setValues(newDifferences);
  }
  
  function recalculateFundedInRawData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName("Raw Data");
    const companyListSheet = spreadsheet.getSheetByName("Company List");
    const relavantDataRange = rawDataSheet.getRange(String.fromCharCode(rawDataColumns.CompanyID + 65) + "2:" + String.fromCharCode(rawDataColumns.Funded + 65) + rawDataSheet.getLastRow());
    const values = relavantDataRange.getValues();
    const companyListValues = companyListSheet.getRange("A2:D" + companyListSheet.getLastRow()).getValues();
    const companyFundedMap = new Map();
    companyListValues.forEach(company => companyFundedMap.set(company[0], company[3]));
    values.forEach((row, rowIndex) => {
      const company = getDerivedFromID(row[0]).company;
      if (companyFundedMap.has(company)) {
        values[rowIndex][rawDataColumns.Funded - rawDataColumns.CompanyID] = companyFundedMap.get(company);
      }
    });
    relavantDataRange.setValues(values);
  }
  
  function refreshRawData() {
    recalculateAllDifferences();
    recalculateFundedInRawData();
  }
  
  function recalculateCompanyList() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName("Raw Data");
    const companyListSheet = spreadsheet.getSheetByName("Company List");
    const rawDataRange = rawDataSheet.getRange(String.fromCharCode(rawDataColumns.Company + 65) + "2:" + String.fromCharCode(rawDataColumns.CohortYear + 65));
    const rawData = rawDataRange.getValues();
    /**
     * @type {Object.<string, Set|string>}
     */
    const lookup = {};
    rawData.forEach(row => {
      // row[0] is CompanyID
      // row[1] is CohortYear
      if (lookup[row[0]] === undefined) {
        lookup[row[0]] = new Set();
      }
      lookup[row[0]].add(+row[1]);
    });
  
    // 2 to skip header, 3 is the column of Cohort Years
    const companyListRange = companyListSheet.getRange(2, CompanyListColumns.Company + 1, companyListSheet.getLastRow() - 1, CompanyListColumns.CohortYears + 1);
    const companyListValues = companyListRange.getValues();
    const currentYear = new Date().getFullYear();
    companyListValues.forEach(row => {
      if (lookup[row[0]] === undefined) return;
      const yearArray = Array.from(lookup[row[0]]).sort();
      row[2] = yearArray.join(", ");
      row[1] = String(bulkCheckIsActive(yearArray, currentYear, 5));
    });
    companyListRange.setValues(companyListValues);
  }