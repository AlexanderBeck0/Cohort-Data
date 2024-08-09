// ######
// The first dashboard I made. There were lots of iterations, and lots of abstracting. You can see most of the old code in Graveyard, but be warned; there is a lot.
// Now there is a loading indicator, controls to show how you want to sort the data, a filter for what data to keep, what data to show, as well as split it into multiple graphs.
// There is also the option to change the graphs, and I implemented a needlessly complex filter system (see getValuesWithFilter() in Dashboard Utilities) to make for more customizability.
// Overall, follow the format of Budget Filter if you want to create your own dashboard. You could also do what I did with createLookupTable() in Helper Functions and abstract the creation of dashboards, but I have not gotten around to that.
// Perhaps a better way to do this would be to make it oop, and have an interface/abstracted class for making dashboards.
// ######
/**
 * @typedef {Object} BudgetFilterOptionValues
 * @property {boolean} ShowFunded Show the companies that have been funded.
 * @property {boolean} ShowNonFunded Show the companies that have not been funded.
 * @property {boolean} ShowAllYears Exclusive to `ShowMostRecentYears`.
 * @property {boolean} ShowMostRecentYears Exclusive to `ShowAllYears`.
 * @property {string} BudgetFilter The budget filter. See `changeBudgetFilter()`.
 * @property {string} SortBy The direction to sort by.
 * @property {string} SortByColumn The column value to sort by.
 * @property {string} DisplayColumn The column value to display on the charts.
 * @property {boolean} EmptyData Show companies with empty data.
 */

/**
 * The event handler triggered when the options in Budget Filter changes.
 * @param {SpreadsheetApp.Spreadsheet} source The spreadsheet being changed.
 * @param {string} sheetName The name of the sheet that was changed. Ignores if the sheet name is not `Budget Filter`
 * @param {number} columnOfChange The index of the column that was changed. Ignores if the index is not 1.
 * @param {number} rowOfChange The index of the row that changed. Index starts at 1. Will ignore if the row is not within `dashboardOptions` (see function body).
 */
function onBudgetFilterChange(source, sheetName, columnOfChange, rowOfChange) {
  // Ignore all edits that are not within Budget Filter
  if (sheetName !== "Budget Filter") return;

  // Ignore all edits that are not within the A column of Budget Filter
  if (columnOfChange !== 1) return;

  /**
   * The row indices of the controllable options in the dashboard. Any change not within these will be ignored.
   * @readonly
   * @see `BudgetFilterOptionValues`
   */
  const dashboardOptions = Object.freeze({
    ShowFunded: 2,
    ShowNonFunded: 4,
    ShowAllYears: 7,
    ShowMostRecentYears: 9,
    BudgetFilter: 12,
    SortBy: 14,
    SortByColumn: 16,
    DisplayColumn: 19,
    EmptyData: 22,
  });

  // Ignore change if it is not within the row index marked by `dashboardOptions`
  if (!Object.values(dashboardOptions).includes(rowOfChange)) return;

  // const budgetSheet = spreadsheet.getSheetByName("Budget Filter");
  const budgetSheet = source.getActiveSheet();
  const rawDataSheet = source.getSheetByName("Raw Data");

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
    loadingRange = budgetSheet.getRange("B1");
    loadingRange.setValue(currentStatus);
    SpreadsheetApp.flush();
  }

  // Make ShowMostRecentYears and ShowAllYears mutually exclusive
  if (rowOfChange === dashboardOptions.ShowAllYears) {
    // Toggle ShowMostRecentYears
    budgetSheet.getRange(dashboardOptions.ShowMostRecentYears, 1, 1, 1).setValue(false);
  } else if (rowOfChange === dashboardOptions.ShowMostRecentYears) {
    // Toggle ShowAllYears
    budgetSheet.getRange(dashboardOptions.ShowAllYears, 1, 1, 1).setValue(false);
  }

  try {
    /**
     * @type {BudgetFilterOptionValues}
     */
    const currentOptions = getCurrentOptions(budgetSheet, dashboardOptions);

    changeBudgetFilter(rawDataSheet, budgetSheet, currentOptions);
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
 * @param {SpreadsheetApp.Sheet} sourceSheet The sheet to get the data from.
 * @param {SpreadsheetApp.Sheet} targetSheet The sheet to place the output data.
 * @param {BudgetFilterOptionValues} currentOptions The current options within the sheet.
 * @see `addDataToSheet()`.
 */
function changeBudgetFilter(sourceSheet, targetSheet, currentOptions) {
  /**   
   * @typedef {Object} MinAndMax
   * @property {string} name
   * @property {number} [minimum]
   * @property {number} [maximum]
   */

  /**
   * @type {MinAndMax[]}
   */
  let filterOptions = [
    { name: "No budget" },
    { name: "<50,000", maximum: 50000 },
    { name: "50,000 - 100,000", minimum: 50000, maximum: 100000 },
    { name: "100,000 - 250,000", minimum: 100000, maximum: 250000 },
    { name: "250,000 - 1,000,000", minimum: 250000, maximum: 1000000 },
    { name: ">1,000,000", minimum: 1000000 },
  ];
  // Freeze all the options in filterOptions
  filterOptions = filterOptions.map(option => Object.freeze(option));

  const includeColumns = ["Revenue", "Expenses", "Difference", "CompanyID", "Funded"];
  if (currentOptions.ShowMostRecentYears) includeColumns.push("Year");

  // Defaults to Revenue
  const columnToSortBy = includeColumns.includes(currentOptions.SortByColumn) ? currentOptions.SortByColumn : "Revenue";
  /**
   * @type {SortByOption[]}
   */
  const sortByOptions = [
    {
      name: "Alphabetical (A-Z)",
      sortFunction:
        /**
         * @param {{CompanyID: string, [key: string]: string | number | boolean}} a The first element for comparison. Will never be `undefined`.
         * @param {{CompanyID: string, [key: string]: string | number | boolean}} b The second element for comparison. Will never be `undefined`.
         */
        (a, b) => a.CompanyID.localeCompare(b.CompanyID)
    },
    {
      name: "Alphabetical (Z-A)",
      sortFunction:
        /**
         * @param {{CompanyID: string, [key: string]: string | number | boolean}} a The first element for comparison. Will never be `undefined`.
         * @param {{CompanyID: string, [key: string]: string | number | boolean}} b The second element for comparison. Will never be `undefined`.
         */
        (a, b) => b.CompanyID.localeCompare(a.CompanyID)
    },
    {
      name: "Highest First",
      sortFunction:
        /**
         * @param {{CompanyID: string, [key: string]: string | number | boolean}} a The first element for comparison. Will never be `undefined`.
         * @param {{CompanyID: string, [key: string]: string | number | boolean}} b The second element for comparison. Will never be `undefined`.
         */
        (a, b) => {
          const aColumns = Object.keys(a).filter(key => key.includes(columnToSortBy)).map(key => a[key]);
          const bColumns = Object.keys(b).filter(key => key.includes(columnToSortBy)).map(key => b[key]);
          const aMax = Math.max(...aColumns);
          const bMax = Math.max(...bColumns);
          return bMax - aMax;
        }
    },
    {
      name: "Lowest First",
      sortFunction:
        /**
         * @param {{CompanyID: string, [key: string]: string | number | boolean}} a The first element for comparison. Will never be `undefined`.
         * @param {{CompanyID: string, [key: string]: string | number | boolean}} b The second element for comparison. Will never be `undefined`.
         */
        (a, b) => {
          const aColumns = Object.keys(a).filter(key => key.includes(columnToSortBy)).map(key => a[key]);
          const bColumns = Object.keys(b).filter(key => key.includes(columnToSortBy)).map(key => b[key]);
          const aMin = Math.min(...aColumns);
          const bMin = Math.min(...bColumns);
          return aMin - bMin;
        }
    },
  ];

  const sortBy = sortByOptions.find(sortOption => sortOption.name === currentOptions.SortBy);
  const budgetOption = filterOptions.find(option => option.name === currentOptions.BudgetFilter);
  const lookupTable = createREDSumLookup(sourceSheet, rawDataColumns, includeColumns, currentOptions.ShowMostRecentYears);

  /**
   * An object that contains all the filters that rely on a single field.
   * @type {Object.<string, function(*): boolean>|null}
   */
  const singularFilterConfig = {
    Difference: (value) => (budgetOption.minimum === undefined || value >= budgetOption.minimum) && (budgetOption.maximum === undefined || value <= budgetOption.maximum),
    Funded: (value) => (currentOptions.ShowFunded && value) || (currentOptions.ShowNonFunded && !value),
  };

  /**
   * An array of objects that contain the filters that rely on multiple fields.
   * @type {Object.<string, function(*): boolean>[]|null}
   */
  const complexFilterConfigs =
    [{
      Revenue: (value) => (!currentOptions.EmptyData && value !== "" && value !== null) || (currentOptions.EmptyData),
      Expenses: (value) => (!currentOptions.EmptyData && value !== "" && value !== null) || (currentOptions.EmptyData),
    }];

  const data = getValuesWithFilter(lookupTable, singularFilterConfig, complexFilterConfigs);

  if (sortBy) {
    data.sort((a, b) => sortBy.sortFunction(a, b));
  } else {
    Logger.log(`Unknown SortBy option provided. Provided "${currentOptions.SortBy}". Not sorting...`);
  }

  /**
   * For each key in columns, if the key in scales is within columns, it will scale the resulting number by number * (10^scale).
   * Example for scaling Revenue down by 1000: 
   * ```javascript
   * Revenue: -3
   * ```
   * @type {Object.<string, number>}
   */
  const scales = {
    Revenue: -3,
    Expenses: -3,
    Difference: -3
  };

  const excludedColumns = ["CompanyID", "Funded"];
  const splitCompanyID = true;
  const includeColumnHeaders = true;
  const startingRow = 29;
  const startingColumn = 4;
  // const newRange = addDataToSheet(data, targetSheet, includeColumns, excludedColumns, true, true, startingRow, startingColumn); // Use for if you want them combined

  // Add the data to the sheet
  const fundedData = data.filter(company => currentOptions.ShowFunded && company.Funded);
  const nonFundedData = data.filter(company => currentOptions.ShowNonFunded && !company.Funded);
  const fundedRange = addDataToSheet(fundedData, targetSheet, includeColumns, scales, excludedColumns, includeColumnHeaders, splitCompanyID, startingRow, startingColumn);

  const lastColumn = fundedRange.getLastColumn();
  const nonFundedStartingColumn = lastColumn === 9 ? lastColumn + 2 : lastColumn + 3;
  const nonFundedRange = addDataToSheet(nonFundedData, targetSheet, includeColumns, scales, excludedColumns, includeColumnHeaders, splitCompanyID, startingRow, nonFundedStartingColumn);

  // Estimate what the column indices are
  let estimatedColumns = Array.from(includeColumns);
  if (splitCompanyID && estimatedColumns.includes("CompanyID")) {
    estimatedColumns.unshift("Company", "Cohort Year");
    estimatedColumns.splice(estimatedColumns.indexOf("CompanyID"), 1);
  }
  estimatedColumns = estimatedColumns.filter(column => !excludedColumns.includes(column));

  updateBudgetFilterGraphs(targetSheet, [fundedRange, nonFundedRange], estimatedColumns, currentOptions);
  // _changeFormatting(); // DON'T USE!!!
}

/**
 * @param {SpreadsheetApp.Sheet} sourceSheet The source sheet.
 * @param {RawDataColumns} sourceColumns The column indices for `sourceSheet`.
 * @param {string[]} includeColumns An array of strings representing keys of `sourceColumns`. MUST include `CompanyID`.
 * @param {boolean} mostRecent A boolean representing if the data should only include the most recent year. MUST include `Year` in `includeColumns` if this is true.
 * @returns {Map.<string, LookupTableValue>} A map of all the values that are listed within includeColumns. Key is `CompanyID`.
 */
function createREDSumLookup(sourceSheet, sourceColumns, includeColumns, mostRecent) {
  if (mostRecent && !includeColumns.includes("Year")) {
    throw "Cannot get most recent year data without Year column. Provide Year column in includeColumns.";
  }

  /**
   * Used to prevent fields from being summed that are not meant to be
   * @type {Set<string>}
   */
  const excludeFromSum = new Set(["CompanyID", "Cohort Year", "CohortYear", "Year"]);
  /**
   * @type {LookupTableCallback}
   */
  const processData = (lookupTable, row, key, order, columns, includeColumns, keyColumn) => {
    const previousData = lookupTable.get(key) || {};
    const newData = {};
    if (mostRecent) {
      if (previousData["Year"] === undefined || (+previousData["Year"] < +row[order.indexOf(columns.Year)] && row.every(val => val !== ""))) {
        includeColumns.filter(column => column !== keyColumn).forEach(column => newData[column] = row[order.indexOf(columns[column])]);
        lookupTable.set(key, newData);
      }
      return;
    }

    // Sum up all the data
    includeColumns.filter(column => column !== keyColumn).forEach(column => {
      newData[column] = previousData[column] ?
        (typeof +previousData[column] === "number" && !excludeFromSum.has(column) ?
          +previousData[column] + +row[order.indexOf(columns[column])] :
          previousData[column]) :
        row[order.indexOf(columns[column])]
    });
    lookupTable.set(key, newData);
  };

  const requiredKeys = ["CompanyID"];
  const keyColumn = "CompanyID";
  return createLookupTable(sourceSheet, sourceColumns, includeColumns, requiredKeys, keyColumn, processData);
}

/**
 * @param {SpreadsheetApp.Sheet} sheet The sheet sheet to update the graphs in.
 * @param {SpreadsheetApp.Range[]} ranges A list of ranges where the data was entered. **Note**: Assumes the number of ranges matchs the number of charts and are in the same order.
 * @param {string[]} columns The order of column indices of the data.
 * @param {BudgetFilterOptionValues} currentOptions The current options within the sheet.
 * @returns {SpreadsheetApp.EmbeddedChart[]} A list of charts that were updated.
 */
function updateBudgetFilterGraphs(sheet, ranges, columns, currentOptions) {
  checkVariablesDefined({ sheet });
  const newHAxisTitle = `The ${currentOptions.DisplayColumn} (in thousands of \$), sorted by ${currentOptions.SortByColumn} (${currentOptions.SortBy})`;
  const charts = sheet.getCharts();
  const returnCharts = [];
  charts.forEach((chart, index) => {
    /**
     * @type {UpdateChartOptions}
     */
    const newChartOptions = {
      useFirstColumnAsDomain: true,
      hAxis: {
        title: newHAxisTitle,
        textPosition: 'out',
        minorGridlines: {
          count: 0
        },
      },
    };

    let chartBuilder = chart.modify();
    chartBuilder = chartBuilder.clearRanges();
    const companyRange = sheet.getRange(ranges[index].getRow(), ranges[index].getColumn() + columns.indexOf("Company"), ranges[index].getNumRows(), 1);
    const columnOfInterestRange = sheet.getRange(ranges[index].getRow(), ranges[index].getColumn() + columns.indexOf(currentOptions.DisplayColumn), ranges[index].getNumRows(), 1);
    chartBuilder = chartBuilder.addRange(companyRange).addRange(columnOfInterestRange);
    const newChart = updateChart(sheet, chartBuilder, newChartOptions);
    returnCharts.push(newChart);
    // Logger.log(newChart.getOptions().get("hAxis.format"))
    // Logger.log(newChart.getOptions().get("hAxis.title").toString())
  });
  return returnCharts;
}