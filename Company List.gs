// ######
// A list of all the unique companies. There is a trigger that runs monthly that updates Active based on yearThreshold.
// Also stores if a company is funded or not, as well as their funding amount (which currently does nothing).
// ######

const CompanyListColumns = Object.freeze({
  Company: 0,
  Active: 1,
  CohortYears: 2,
  Funded: 3,
  FundingAmount: 4,
});

/**
 * Updates the Cohort Years in `Company List`. Also checks if the company is active or not.
 * @param {SpreadsheetApp.Sheet} companyListSheet The sheet to retrieve the data from.
 * @param {string|number} cohortYear The cohort year to update or add.
 * @param {number} rowIndexOfCompany The non-negative row index of the company. Index starts at 1.
 * @param {number} cohortYearColumnIndex The non-negative column index to place the cohortYear in.
 * @param {string} [delim = ", "] The delimiter to split the years with. Defaults to `", "`.
 * @see `recalculateCompanyList()`
 */
function updateOrAddCohortYearList(companyListSheet, cohortYear, rowIndexOfCompany, cohortYearColumnIndex, delim = ", ") {
  checkVariablesDefined({ companyListSheet, rowIndexOfCompany, cohortYearColumnIndex });
  checkVariablesGreaterThanZero({ rowIndexOfCompany, cohortYearColumnIndex });
  if (delim === undefined) {
    delim = ", ";
  }

  const cohortYearsRange = companyListSheet.getRange(rowIndexOfCompany, cohortYearColumnIndex, 1);
  let currentCohortYears = String(cohortYearsRange.getValue()); // Note: Must force this as a String as for some reason the 4 times I do it below don't seem to do the trick...
  if (currentCohortYears) {
    // Check if the current cohort year is already within cohortYears
    const cohortYears = String(currentCohortYears).split(delim);
    if (!cohortYears.includes(String(cohortYear))) {
      // Add the new Cohort Year
      cohortYears.push(String(cohortYear));
      const sortedYears = cohortYears.sort();
      currentCohortYears = sortedYears.join(delim);
      cohortYearsRange.setValue(currentCohortYears);
    }
  } else {
    // It is empty. Add data to the cohort
    currentCohortYears = String(cohortYear);
    cohortYearsRange.setValue(currentCohortYears);
  }

  // Check if the company is still active
  const activeRange = companyListSheet.getRange(rowIndexOfCompany, CompanyListColumns.Active + 1, 1);
  const isActive = checkIsActive(currentCohortYears, 5, delim);
  activeRange.setValue(String(isActive));
}

/**
 * @param {SpreadsheetApp.Sheet} companyListSheet The sheet to retrieve the data from.
 * @param {string} company The company to add to the Company List.
 * @param {string|number} cohortYear The cohort year to update or add.
 * @returns {number} The index that the company was added to.
 */
function addToCompanyList(companyListSheet, company, cohortYear) {
  const firstEmptyRow = companyListSheet.getLastRow() + 1;
  // Use CohortYears to set all the data up to it
  companyListSheet.getRange(firstEmptyRow, 1, 1, CompanyListColumns.CohortYears + 1).setValues([[company, checkIsActive(String(cohortYear)), String(cohortYear)]]);
  return firstEmptyRow;
}

/**
 * @param {SpreadsheetApp.Sheet} companyListSheet The sheet to place the data in.
 * @param {string} company The company name to add `funds` to in `companyListSheet`.
 * @param {number|undefined} funds The amount of funds granted. Undefined if the company is just being ticked as funded without a funding amount.
 * @param {number|undefined} index The non-negative index of the company in Company List. If it is not provided, it will be found using `rowIndexOfValue()`.
 */
function setCompanyAsFunded(companyListSheet, company, funds, index) {
  if (index === undefined) {
    // 2 to skip header row
    index = rowIndexOfValue(company, 1, companyListSheet, 2);
    if (index === -1) {
      throw "Company not found! Cannot set company as funded.";
    }
  }
  const range = companyListSheet.getRange(index, 1, 1, companyListSheet.getLastColumn());
  const values = range.getValues();
  // Mark as funded
  // 3 is column of Funded
  values[0][CompanyListColumns.Funded] = String(true);

  // Add funds if there are any
  if (funds !== undefined && !isNaN(+funds)) {
    // 4 is column of Funding Amount
    values[0][CompanyListColumns.FundingAmount] = funds;
  }

  range.setValues(values);
}

/**
 * @param {string} cohortYears The list of cohortYears to check if they are active or not.
 * @param {number} [yearThreshold = 5] The non-negative number of years before a company is considered retired. Defaults to 5. Minimum of 1.
 * @param {string} [delim = ", "] The delimiter used to split up `cohortYears`. Defaults to `", "`.
 * @returns {boolean} `true` if at least one of the cohort years within `cohortYears` is within the `yearThreshold` threshold, and `false` if none of them are within the threshold. Also returns `false` if `cohortYears` is empty.
 * @see `bulkCheckIsActive()`
 */
function checkIsActive(cohortYears, yearThreshold = 5, delim = ", ") {
  checkVariablesDefined({ cohortYears });

  if (cohortYears.length === 0) {
    return false;
  }
  if (yearThreshold === undefined) {
    yearThreshold = 5;
  }
  checkVariablesGreaterThanZero({ yearThreshold });
  if (delim === undefined) {
    delim = ", ";
  }

  const currentYear = new Date().getFullYear();
  const years = cohortYears.split(delim).map(year => +year);
  const mostRecentYear = Math.max(...years);
  return mostRecentYear >= currentYear - yearThreshold;
}

/**
 * Used as opposed to checkIsActive when there are many checks to do. Providing the currentYear beforehand reduces the number of computations.
 * 
 * Calls `checkIsActive()` if `currentYear` is `undefined`. 
 * @param {number[]} cohortYears The list of cohortYears to check if they are active or not.
 * @param {number} currentYear The current year.
 * @param {number} [yearThreshold = 5] The non-negative number of years before a company is considered retired. Defaults to 5. Minimum of 1.
 * @param {string} [delim = ", "] The delimiter used to split up `cohortYears`. Defaults to `", "`.
 * @returns {boolean} `true` if at least one of the cohort years within `cohortYears` is within the `yearThreshold` threshold, and `false` if none of them are within the threshold. Also returns `false` if `cohortYears` is empty.
 * @see `checkIsActive()`
 */
function bulkCheckIsActive(cohortYears, currentYear, yearThreshold = 5, delim = ", ") {
  checkVariablesDefined({ cohortYears });
  if (cohortYears.length === 0) {
    return false;
  }
  if (yearThreshold === undefined) {
    yearThreshold = 5;
  }
  checkVariablesGreaterThanZero({ yearThreshold });
  if (delim === undefined) {
    delim = ", ";
  }
  if (currentYear === undefined) {
    return checkIsActive(cohortYears.join(delim), delim, yearThreshold);
  }

  const mostRecentYear = Math.max(...cohortYears);
  return mostRecentYear >= currentYear - yearThreshold;
}

/**
 * @returns {string[]} All the companies in Company List that are marked as Active.
 */
function getActiveCompanies() {
  const companyListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Company List");
  // I'm too lazy to translate this into using CompanyListColumns
  const companyListValues = companyListSheet.getRange("A2:B" + companyListSheet.getLastRow()).getValues();
  const activeCompanies = companyListValues.filter(row => row[1]); // Only keep the companies that are marked as active
  const activeCompanyNames = activeCompanies.map(row => row[0]); // Get only the names
  return activeCompanyNames;
}

/**
 * @param {boolean} [includeFunded = false] `false` if the output should just be the strings of CompanyIDs, and `true` if it should be an object with CompanyID and Funded. Defaults to `false`.
 * @returns {string[]|{CompanyID: string, Funded: boolean}[]} A list of all the unique `CompanyID`'s in Company List if `includeFunded` is `false`, or a list of objects with unique `CompanyID`s, and their associated `Funded` properties if `includeFunded` is `true`. 
 */
function getAllCompanyIDs(includeFunded = false) {
  /**
   * @type {Set<string|{CompanyID: string, Funded: boolean}>}
   */
  const uniqueCompanyIDs = new Set();
  const companyListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Company List");
  // I'm too lazy to translate this into using CompanyListColumns
  const companyListValues = companyListSheet.getRange("A2:" + (includeFunded ? "D" : "C") + companyListSheet.getLastRow()).getValues();
  companyListValues.forEach(row => {
    const company = row[0];
    const cohortYears = String(row[2]).split(", ");
    cohortYears.forEach(year => {
      const companyID = `${company}-\$-${year}`;
      uniqueCompanyIDs.add(includeFunded ? { CompanyID: companyID, Funded: row[3] } : companyID);
    });
  });
  return Array.from(uniqueCompanyIDs);
}