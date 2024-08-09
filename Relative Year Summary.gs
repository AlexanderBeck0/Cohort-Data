// #####
// Code for Relative Year Summary. 
// #####

/**
 * The event handler triggered when the options in Relative Year Summary changes.
 * @param {SpreadsheetApp.Spreadsheet} source The spreadsheet the change was made to.
 * @param {string} sheetName The name of the sheet that was changed. Ignores if the sheet name is not `Relative Year Summary`
 * @param {number} columnOfChange The index of the column that was changed. Ignores if the index is not `2`.
 * @param {number} rowOfChange The index of the row that changed. Index starts at 1. Will ignore if the row is not within `dashboardOptions` (see function body).
 * @returns {boolean} A flag representing if the values were successfully changed or not.
 */
function onRYSEdit(source, sheetName, columnOfChange, rowOfChange) {
  // Ignore all edits that are not within the only button
  if (sheetName !== "Relative Year Summary") return false;
  if (columnOfChange !== 3) return false;

  const dashboardOptions = Object.freeze({
    CarryOver: 2
  });

  if (!Object.values(dashboardOptions).includes(rowOfChange)) return false; // Redundant since there is only one value, but will keep for consistency purposes 

  const doCarryOver = source.getSheetByName(sheetName).getRange("C2").getValue();
  const addedData = updateYearValues(doCarryOver); // Note: this assumes that there is only one option in RYS

  return true;
}

/**
 * @returns {{CompanyID: number, Company: number, CohortYear: number, [key: string]: number}} The Relative Year Summary Columns. Also includes YearNDifference and YearNPercentage.
 */
function generateRYSColumns() {
  const RYSColumns = {
    CompanyID: 0,
    Company: 1,
    CohortYear: 2
  };

  // YearNDifference and YearNPercentage columns
  // Note: The year columns must be the last ones. This means that all other data columns, such as CompanyID, Company, and CohortYear must come 
  // before them.
  const previousIndex = Math.max(...Object.values(RYSColumns));
  for (let i = 1, j = previousIndex + 1; i <= NUMBER_OF_YEARS; i++, j += 2) {
    RYSColumns[`Year${i}Difference`] = j;
    RYSColumns[`Year${i}Percentage`] = j + 1;
  }
  return RYSColumns;
}

/**
 * Updates Year 1-`NUMBER_OF_YEARS` Difference and Percentage values in Relative Year Summary
 * 
 * Uses the following properties from `rawDataColumns`:
 * - CompanyID
 * - CohortYear
 * - Year
 * - Difference
 * @param {boolean | undefined} [doCarryOver = undefined] Whether the Relative Year Summary should carry over its values. This means adding the inverse of the difference from the previous year to the next year.
 * @returns {(string | number | boolean)[]} The data added to the sheet
 */
function updateYearValues(doCarryOver = undefined) {
  // Originally had 3 columns that represented Year 1, 3, and 4
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = spreadsheet.getSheetByName("Raw Data");
  const RYS = spreadsheet.getSheetByName("Relative Year Summary");

  const carryOverOptionRange = RYS.getRange("C2");
  // Get the option for doCarryOver if it is not provided
  if (doCarryOver === undefined || doCarryOver === null) doCarryOver = carryOverOptionRange.getValue();

  if (doCarryOver !== carryOverOptionRange.getValue()) {
    carryOverOptionRange.setValue(doCarryOver);
    SpreadsheetApp.flush();
  }

  const lookupTable = createIDYearLookupTable(rawDataSheet, doCarryOver);

  /**
   * A list of all the unique company ids.
   * @type {string[]}
   */
  const uniqueCompanyIDs = getAllCompanyIDs(false); // false to only get company id's

  /**
   * The Relative Year Summary data. Also includes YearNDifference and YearNPercentage.
   * @type {(string | number | boolean)[]}
   */
  const data = [];
  uniqueCompanyIDs.forEach(id => {
    const [company, cohortYear] = id.split("-$-");
    const years = calculateYearsFromCohortYear(cohortYear);
    const companyData = { CompanyID: id, Company: company, CohortYear: cohortYear };
    let previousDifference = 0;
    for (let i = 0; i < years.length - 1; i++) {
      const yearValue1 = lookupTable.get(`${id}-${years[i]}`);
      const yearValue2 = lookupTable.get(`${id}-${years[i + 1]}`);

      const { difference, percentage } = calculateDifferenceAndPercentage(yearValue1, yearValue2, doCarryOver, previousDifference, "");
      companyData[`Year${i + 1}Difference`] = difference;
      companyData[`Year${i + 1}Percentage`] = percentage;
      previousDifference = difference;
    }
    data.push(companyData);
  });

  // RYS.getRange(3, 1, data.length, data[0].length).setValues(data);
  const columns = Object.keys(generateRYSColumns());
  const yearlyFormats = {
    ...Object.fromEntries(columns.filter(column => column.includes("Difference")).map(column => [column, "currency"])),
    ...Object.fromEntries(columns.filter(column => column.includes("Percentage")).map(column => [column, "percent"])),
  };

  const range = addDataToSheet(data, RYS, columns, null, [], false, false, 3, 1);
  SpreadsheetApp.flush();
  updateColumnNumberFormats(RYS, columns.filter(column => !["CompanyID, Company, CohortYear"].includes(column)), yearlyFormats, range);
  return data;
}

/**
 * Uses the following properties from `rawDataColumns`:
 * - CompanyID
 * - Year
 * - Difference
 * @private
 * @param {SpreadsheetApp.Sheet} rawDataSheet The Raw Data sheet
 * @returns {Map.<string, number>} A lookup table with the key being `CompanyID-Year` and the value being the `Difference` value of the company at that year
 */
function createIDYearLookupTable(rawDataSheet) {
  /**
   * @type {LookupTableCallback}
   */
  const processData = (lookupTable, row, key, order, columns, includeColumns, keyColumn) => {
    const year = row[order.indexOf(columns["Year"])];
    const difference = row[order.indexOf(columns["Difference"])];
    lookupTable.set(`${key}-${year}`, difference);
  }

  const includeColumns = ["CompanyID", "Year", "Difference"];
  const requiredKeys = ["CompanyID", "Year", "Difference"];
  const keyColumn = "CompanyID";
  return createLookupTable(rawDataSheet, rawDataColumns, includeColumns, requiredKeys, keyColumn, processData);
}

/**
 * Calculates the difference and percentage between firstValue and secondValue.
 * 
 * Difference: `secondValue - firstValue`
 * 
 * Percentage: `(secondValue - firstValue) / firstValue`
 * 
 * Edge cases:
 * - `1 / 0`: Infinity
 * - `0 / 0`: Infinity
 * - `undefined / 1`: `defaultValue`
 * - `1 / undefined`: `defaultValue`
 * - `undefined / undefined`: `defaultValue`
 * 
 * @param {number} firstValue The first number. Note, if `firstValue` is negative, it will use the absolute value when dividing. Additionally, if `firstValue` is 0, it is a guaranteed edge case
 * @param {number} secondValue The second number.
 * @param {boolean} [doCarryOver = false] Flag to determine if carry-over should be applied.
 * @param {number} [previousDifference=0] The difference from the previous year (only used if doCarryOver is true).
 * @param {string} [defaultValue=""] The default value if there is an issue (such as an edge case)
 * @returns {{difference: number, percentage: number}} An Object with a difference and percentage value. If it hits an edge case, the value will be `defaultValue`.
 */
function calculateDifferenceAndPercentage(firstValue, secondValue, doCarryOver = false, previousDifference = 0, defaultValue = "") {
  if (defaultValue === undefined) {
    defaultValue = "";
  }

  let difference = defaultValue;
  let percentage = defaultValue;

  if (firstValue === undefined || secondValue === undefined) {
    // There is at least 1 empty value. Return empty for both
    return { difference: difference, percentage: percentage };
  }

  if (firstValue === defaultValue || secondValue === defaultValue) {
    // Empty values
    // Return empty for both
    return { difference: difference, percentage: percentage };
  }

  difference = secondValue - firstValue;

  if (doCarryOver) {
    difference += previousDifference === "" ? 0 : previousDifference;
  }

  if (firstValue !== 0) {
    // Use Math.abs() to ensure that percentage is only negative when difference is negative
    percentage = difference / Math.abs(firstValue);
  } else {
    // 0/0
    percentage = "Infinity";
  }

  return { difference: difference, percentage: percentage };
}