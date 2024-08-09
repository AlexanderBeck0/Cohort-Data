// #####
// This are for finding the companies with data missing from years.
// #####

/**
 * @typedef {Object} MissingDataDashboardValues
 * @property {boolean} Refresh Whether the values should be refreshed or not.
 * @property {boolean} SeparateFundedNonFunded Whether the Funded data should be separated from Non-Funded data. If false, they will be combined.
 */

/**
 * The event handler triggered when the options in Companies With Missing Data changes.
 * @param {SpreadsheetApp.Spreadsheet} source The spreadsheet being changed.
 * @param {string} sheetName The name of the sheet that was changed. Ignores if the sheet name is not `Companies With Missing Data`.
 * @param {number} columnOfChange The index of the column that was changed. Ignores if the index is not 1.
 * @param {number} rowOfChange The index of the row that changed. Index starts at 1. Will ignore if the row is not within `dashboardOptions` (see function body).
 */
function onMissingDataSheetChange(source, sheetName, columnOfChange, rowOfChange) {
  // Ignore all edits that are not within the dropdown
  if (sheetName !== "Companies With Missing Data") return;
  if (columnOfChange !== 1) return;

  const dashboardOptions = Object.freeze({
    Refresh: 2,
    SeparateFundedNonFunded: 5,
  });

  if (rowOfChange !== dashboardOptions.Refresh) return; // Since it only works with Refresh anyways, why waste time doing the rest and instead only check if it is updated?

  if (!Object.values(dashboardOptions).includes(rowOfChange)) return;

  const enableLoadingIndicator = true;

  // const missingDataSheet = spreadsheet.getSheetByName("Companies With Missing Data");
  const missingDataSheet = source.getActiveSheet();
  const rawDataSheet = source.getSheetByName("Raw Data");

  /**
   * @type {MissingDataDashboardValues}
   */
  const currentOptions = getCurrentOptions(missingDataSheet, dashboardOptions);

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
    loadingRange = missingDataSheet.getRange(dashboardOptions.Refresh, 2, 1, 1);
    loadingRange.setValue(currentStatus);
    SpreadsheetApp.flush();
  }

  if (currentOptions.Refresh) {
    try {
      Logger.log("Refreshing Missing Data...");
      const startingRow = 4;
      const startingColumn = 2;
      refreshCompaniesWithMissingData(rawDataSheet, rawDataColumns, missingDataSheet, startingRow, startingColumn, currentOptions.SeparateFundedNonFunded);
    } catch (error) {
      Logger.log("Caught error: %s", error);
      throwError = error;
      currentStatus = "Failed!";
    } finally {
      // Toggle the Refresh button
      const doRefreshRange = missingDataSheet.getRange(dashboardOptions.Refresh, 1, 1, 1);
      doRefreshRange.setValue(false);
    }
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
 * @param {SpreadsheetApp.Sheet} sourceSheet The sheet to retrieve the missing data from.
 * @param {Object} sourceColumns The column indices of sourceSheet
 * @param {SpreadsheetApp.Sheet} targetSheet The sheet to place the companies with missing data in.
 * @param {number} [startingRow=1] The non-negative row to start placing the data in. Index starts at 1.
 * @param {number} [startingColumn=1] The non-negative column to start placing data in. Index starts at 1.
 * @param {boolean} [separateFunded = false] Whether the Funded and Non-Funded companies should be separated
 */
function refreshCompaniesWithMissingData(sourceSheet, sourceColumns, targetSheet, startingRow = 1, startingColumn = 1, separateFunded = false) {
  checkVariablesTruthy({ sourceSheet, sourceColumns, targetSheet, startingRow, startingColumn, separateFunded });
  checkVariablesGreaterThanZero({ startingRow, startingColumn });
  if (typeof separateFunded !== "boolean") {
    throw new TypeError("separateFunded must be of type boolean!");
  }

  const includeColumns = ["CompanyID", "Year", "Revenue", "Expenses", "Difference"];
  // if (separateFunded) includeColumns.push("Funded");
  const lookup = createCompaniesMissingDataLookup(sourceSheet, sourceColumns, includeColumns);
  const allCompanyIDs = getAllCompanyIDs(separateFunded);
  const companiesMissingData = getCompaniesMissingData(lookup, allCompanyIDs, true, separateFunded);
  const outputColumns = ["Company", "Years"];
  const columnOffset = outputColumns.length + 1;
  if (separateFunded) {
    const addFundedRange = addDataToSheet(companiesMissingData.filter(company => company.Funded), targetSheet, outputColumns, null, [], true, false, startingRow, startingColumn);
    const addNonFundedRange = addDataToSheet(companiesMissingData.filter(company => !company.Funded), targetSheet, outputColumns, null, [], true, false, startingRow, startingColumn + columnOffset);
    targetSheet.getRange(startingRow - 1, startingColumn, 1, startingColumn + columnOffset).clearContent();
    targetSheet.getRange(startingRow - 1, startingColumn, 1, 1).setValue("Funded");
    targetSheet.getRange(startingRow - 1, startingColumn + columnOffset, 1, 1).setValue("Non-Funded");
  } else {
    const addedRange = addDataToSheet(companiesMissingData, targetSheet, outputColumns, null, [], true, false, startingRow, startingColumn);
    targetSheet.getRange(startingRow - 1, startingColumn, 1, startingColumn + columnOffset).clearContent();
    targetSheet.getRange(startingRow - 1, startingColumn, 1, 1).setValue("Funded & Non-Funded Companies Missing Data")
  }
}

/**
 * A function that is given a list of the unique company ids of the form `Comapny-$-CohortYear`. It is also given a lookup table with the key of `CompanyID-Year`.
 * @param {Map.<string, LookupTableValue>} lookup A map with the key as `CompanyID-Year` and the value as an object containing the different values.
 * @param {string[] | {CompanyID: string, Funded: boolean}[]} uniqueCompanyIDs A list of all the unique company ids in Company List, or a list of objects with unique `CompanyID`s, and their associated `Funded` properties if `separateFunded` is true
 * @param {boolean} [mergeRepeatCompanies = true] Whether the years of companies that have multiple cohort years should be merged.
 * @param {boolean} [separateFunded = false] Whether the Funded and Non-Funded companies should be separated. Defaults to false.
 * @returns {{Company: string, Years: string, Funded?: boolean}[]} A list of companies and the years they are missing data.
 */
function getCompaniesMissingData(lookup, uniqueCompanyIDs, mergeRepeatCompanies = true, separateFunded = false) {
  // I'm too lazy to check if these variables are truthy/defined
  checkVariablesDefined({ lookup, uniqueCompanyIDs, mergeRepeatCompanies });
  if (typeof mergeRepeatCompanies !== "boolean") {
    throw new TypeError("Parameter 'mergeRepeatCompanies' must be a boolean! Instead got " + typeof mergeRepeatCompanies + ".");
  }
  /**
   * An array containing all the companies in uniqueCompanyID, and the years they are missing data
   * @type {{Company: string, Years: string, Funded?: boolean}[]}
   */
  const companiesMissingData = [];
  uniqueCompanyIDs.forEach(companyID => {
    /**
     * Used instead of just companyID because companyID could be either a string or an object with a CompanyID property (which is a string)
     * @type {string}
     */
    const useCompanyID = companyID.CompanyID || companyID;
    const { company, cohortYear } = getDerivedFromID(useCompanyID);
    const years = calculateYearsFromCohortYear(cohortYear);
    /**
     * An array containing the missing years for a given company
     * @type {number[]}
     */
    const missingYears = [];
    years.forEach(year => {
      if (!lookup.has(`${useCompanyID}-${year}`)) {
        // Year is missing from the lookup table
        missingYears.push(year);
        return;
      }
      // The lookup has the year. Check if there is any data
      /**
       * @type {LookupTableValue}
       */
      const yearEntry = lookup.get(`${useCompanyID}-${year}`);
      const revenue = yearEntry["Revenue"];
      const expenses = yearEntry["Expenses"];
      const difference = yearEntry["Difference"];
      /**
       * @type {(value: number|string) => boolean}
       */
      const isNaNOrEmpty = (value) => Number.isNaN(value) || value === "";
      // While not expected behaviour, it will count "Less than 50,000" as valid. However, it will still add it to missingYears because difference won't calculate properly
      if (isNaNOrEmpty(revenue) || isNaNOrEmpty(expenses) || isNaNOrEmpty(difference)) {
        missingYears.push(year);
      }
    });
    // Only add to missingCompanies if there is at least 1 year
    if (missingYears.length > 0) {
      const objToAdd = { Company: company, Years: missingYears.sort().join(", ") };
      if (separateFunded) {
        try {
          objToAdd["Funded"] = companyID.Funded;
        } catch (error) {
          throw "Error! Unique Company ID's does not have a Funded property!";
        }
      }
      companiesMissingData.push(objToAdd);
    }
  });
  return mergeRepeatCompanies ? mergeCompaniesMissingData(companiesMissingData) : companiesMissingData;
}

/**
 * Merge duplicate companies (companies that have multiple cohort years).
 * @param {{Company: string, Years: string, Funded?: boolean}[]} companies The companies that are being merged. Must have `Company` and `Year` properties.
 * @returns {{Company: string, Years: string, Funded?: boolean}[]} A shallow copy of `companies` with companies that are duplicates merged. 
 */
function mergeCompaniesMissingData(companies) {
  companies.forEach(company => checkObjectPropertiesDefined(company, ["Company", "Years"]));
  // Would definitely be more efficient to look at Company List and see which companies have multiple cohort years, but that would require rewriting how CompanyIDs are passed to this function
  /**
   * A map to keep track of merged companies.
   * @type {Map<string, {Company: string, Years: string}|{Company: string, Years: string, Funded: boolean}>}
   */
  const companyMap = new Map();
  companies.forEach(company => {
    if (!companyMap.has(company.Company)) {
      companyMap.set(company.Company, { ...company });
      return;
    }
    // Merge the years since the map does not already have it
    const existingCompany = companyMap.get(company.Company);
    const existingYears = new Set(existingCompany.Years.split(", "));
    const newYears = company.Years.split(", ");
    newYears.forEach(year => existingYears.add(year));
    existingCompany.Years = Array.from(existingYears).sort().join(", ");
  });
  return Array.from(companyMap.values());
}

/**
 * Creates a lookup table with the key as `CompanyID-Year` and the value as a `LookupTableValue`.
 * @param {SpreadsheetApp.Sheet} sourceSheet The source sheet.
 * @param {RawDataColumns} sourceColumns The source columns.
 * @param {string[]} includeColumns A string array representing keys of `sourceColumns`. MUST include `CompanyID`.
 * @returns {Map.<string, LookupTableValue>} A map with the key as `CompanyID-Year` and the value as an object containing the different values.
 */
function createCompaniesMissingDataLookup(sourceSheet, sourceColumns, includeColumns) {
  checkVariablesDefined({ sourceSheet, sourceColumns, includeColumns });
  checkArrayContainsElements(includeColumns, "CompanyID", "Year")

  // Ensure that all columns are valid
  includeColumns.forEach(column => {
    if (sourceColumns[column] === undefined) {
      throw `Unknown key provided in includeColumns: "${column}". Could not find column.`;
    }
  });

  /**
   * Used for getting the relative index of a column in sourceColumns.
   * Example of how to use order with `sourceColumns` and a valid `column` string:
   * ```javascript
   * row[order.indexOf(sourceColumns[column])]
   * ```
   * @type {number[]}
   */
  const order = includeColumns.map(key => sourceColumns[key]);
  const filteredData = getDataInOrder(sourceSheet, sourceColumns, includeColumns, order);

  /**
   * @type {Map.<string, LookupTableValue>}
   */
  const lookupTable = new Map();
  filteredData.forEach(row => {
    const companyID = row[order.indexOf(sourceColumns.CompanyID)];
    const year = +row[order.indexOf(sourceColumns.Year)];
    if (!companyID || companyID === "") return;
    if (!year) return;
    const newID = `${companyID}-${year}`;
    const existingData = lookupTable.get(newID) || {};
    const newData = { ...existingData };
    includeColumns.filter(column => column !== "CompanyID" && column !== "Year").forEach(column => newData[column] = row[order.indexOf(sourceColumns[column])]);
    lookupTable.set(newID, newData);
  });
  return lookupTable;
}
