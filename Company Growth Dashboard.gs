// ######
// Another dashboard similar to Budget Filter. Some slight differences, but can be used as an example for creating another dashboard.
// ######

/**
 * @typedef {Object} GrowthOptionValues
 * @property {boolean} ShowFunded Show the companies that have been funded
 * @property {boolean} ShowNonFunded Show the companies that have not been funded
 * @property {boolean} Year1 Show Year 1's values
 * @property {boolean} Year2 Show Year 2's values
 * @property {boolean} Year3 Show Year 3's values
 * @property {boolean} Year4 Show Year 4's values
 * @property {boolean} CarryOver Flag to determine if the data from the previous year should be carried over.
 * @property {string} SortBy The direction to sort by
 * @property {boolean} EmptyData Show companies with empty data
 */

/**
 * The event handler triggered when the options in Company Growth Dashboard changes
 * @param {SpreadsheetApp.Spreadsheet} source The spreadsheet being changed.
 * @param {string} sheetName The name of the sheet that was changed. Ignores if the sheet name is not `Company Growth Dashboard`
 * @param {number} columnOfChange The index of the column that was changed. Ignores if the index is not 1.
 * @param {number} rowOfChange The index of the row that changed. Index starts at 1. Will ignore if the row is not within `dashboardOptions` (see function body).
 */
function onGrowthDashboardChange(source, sheetName, columnOfChange, rowOfChange) {
  // Ignore all edits that are not within the dropdown
  if (sheetName !== "Company Growth Dashboard") return;
  if (columnOfChange !== 1) return;

  /**
   * @type {GrowthOptionValues}
   */
  const dashboardOptions = Object.freeze({
    ShowFunded: 3,
    ShowNonFunded: 5,
    Year1: 8,
    Year2: 10,
    Year3: 12,
    Year4: 14,
    CarryOver: 17,
    SortBy: 20,
    EmptyData: 23,
  });

  if (!Object.values(dashboardOptions).includes(rowOfChange)) return;

  const dashboardSheet = source.getSheetByName(sheetName); // getActiveSheet() is much faster, but can lead to unintended issues
  const RYSSheet = source.getSheetByName("Relative Year Summary");
  const companyListSheet = source.getSheetByName("Company List");

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
    loadingRange = dashboardSheet.getRange("B2");
    loadingRange.setValue(currentStatus);
    SpreadsheetApp.flush();
  }

  try {
    /**
     * @type {GrowthOptionValues}
     */
    const currentOptions = getCurrentOptions(dashboardSheet, dashboardOptions);
    if (rowOfChange === dashboardOptions.CarryOver) updateYearValues(currentOptions.CarryOver);

    changeGrowthDashboardFilter(RYSSheet, dashboardSheet, companyListSheet, currentOptions);
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
 * Runs `onGrowthDashboardChange()` with `source` (if it is provided) and uses `A3` as the simulated change. 
 * @param {SpreadsheetApp.Spreadsheet | undefined} [source = undefined] The active spreadsheet.
 */
function refreshGrowthDashboard(source = undefined) {
  Logger.log("Refreshing Growth Dashboard...");
  if (source === undefined) source = SpreadsheetApp.getActiveSpreadsheet();
  onGrowthDashboardChange(source, "Company Growth Dashboard", 1, 3);
}

/**
 * @param {SpreadsheetApp.Sheet} sourceSheet The sheet to get the data from
 * @param {SpreadsheetApp.Sheet} targetSheet The sheet to place the output data
 * @param {SpreadsheetApp.Sheet} companyListSheet The sheet that holds where the funded data is
 * @param {GrowthOptionValues} currentOptions The current options within the sheet
 * @see `addDataToSheet()`
 */
function changeGrowthDashboardFilter(sourceSheet, targetSheet, companyListSheet, currentOptions) {
  const columnToSortBy = "Percentage"; // Change to change what column to sort by (Percentage, Difference, Revenue, Expenses, etc.)
  /**
   * @type {SortByOption[]}
   * @see `columnToSortBy`
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

  /**
   * @type {string[]}
   */
  const includeColumns = ["CompanyID"];

  // Add all the necessary years
  for (let i = NUMBER_OF_YEARS - 1; i > 0; i--) {
    includeColumns.unshift(`Year${i}Percentage`);
    includeColumns.unshift(`Year${i}Difference`);
  }

  /**
   * A list that represents which years to include.
   * @type {boolean[]}
   */
  const includeYears = [currentOptions.Year1, currentOptions.Year2, currentOptions.Year3, currentOptions.Year4];

  const companyListColumns = { Company: 0, Active: 1, CohortYears: 2, Funded: 3, FundingAmount: 4 };
  const fundedLookup = createFundedLookup(companyListSheet, companyListColumns, ["Company", "Funded"]);
  const lookupTable = createGrowthLookup(sourceSheet, generateRYSColumns(), includeColumns, includeYears, fundedLookup);

  /**
   * An object that contains all the filters that rely on a single field.
   * @type {Object.<string, function(*): boolean>|null}
   */
  const singularFilterConfig = null;

  const emptyDataConfigs = generateYearlyFilterConfigs(includeYears, currentOptions.EmptyData);
  /**
   * An array of objects that contain the filters that rely on multiple fields.
   * @type {Object.<string, function(*): boolean>[]|null}
   */
  const complexFilterConfigs = [...emptyDataConfigs];

  const data = getValuesWithFilter(lookupTable, singularFilterConfig, complexFilterConfigs);

  /**
   * A list of all the columns to exclude from the output. 
   * @type {string[]}
   */
  const excludedColumns = ["CompanyID", "Funded"];

  includeYears.forEach((year, index) => {
    if (!year) {
      excludedColumns.push(`Year${index + 1}Difference`);
      excludedColumns.push(`Year${index + 1}Percentage`);
    }
  });

  if (sortBy) {
    data.sort((a, b) => sortBy.sortFunction(a, b));
  } else {
    Logger.log(`Unknown SortBy option provided. Provided "${currentOptions.SortBy}". Not sorting...`);
  }

  const includeColumnHeaders = true;
  const splitCompanyID = true;

  // Estimate what the column indices are
  let estimatedColumns = Array.from(includeColumns);
  if (splitCompanyID && estimatedColumns.includes("CompanyID")) {
    estimatedColumns.unshift("Company", "Cohort Year");
    estimatedColumns.splice(estimatedColumns.indexOf("CompanyID"), 1);
  }
  estimatedColumns = estimatedColumns.filter(column => !excludedColumns.includes(column));

  const scales = {
    ...Object.fromEntries(estimatedColumns.filter(column => column.includes("Percentage")).map(column => [column, 2])),
  };

  // 29 for row 29, 4 for column D
  const fundedRange = addDataToSheet(data.filter(company => currentOptions.ShowFunded && company.Funded), targetSheet, includeColumns, scales, excludedColumns, includeColumnHeaders, splitCompanyID, 28, 4);

  // 29 for row 29, 15 for column O
  const nonFundedRange = addDataToSheet(data.filter(company => currentOptions.ShowNonFunded && !company.Funded), targetSheet, includeColumns, scales, excludedColumns, includeColumnHeaders, splitCompanyID, 28, 15);

  // Logger.log(includeColumns.filter(column => !excludedColumns.includes(column) && column.includes("Percentage")).map(column => column.charAt(4)));
  // Logger.log(includeColumns.filter(column => column.includes("Percentage")).map((column, index) => {
  //   const label = `${ordinal(index + 1)} Year Growth (%)`;
  //   return {
  //     targetAxisIndex: index,
  //     labelInLegend: label,
  //     visibleInLegend: currentOptions[`Year${index + 1}`]
  //   };
  // }));
  // Logger.log(targetSheet.getCharts()[0].getOptions().getOrDefault("series.0.visibleInLegend"));

  // It is certainly not pretty, but I had to scale the percentage because otherwise the number format on the graph would be 200 instead of 20000 like it is supposed to be.
  // I spent 4.5 hours trying to get it to work, with literally no progress whatsoever.
  /**
   * @type {{[column: string]: Format}}
   */
  const formats = {
    ["Cohort Year"]: "none",
    ...Object.fromEntries(estimatedColumns.filter(column => column.includes("Difference") || column.includes("Percentage")).map(column => [column, column.includes("Difference") ? "currency" : "decimal"])),
  };
  updateColumnNumberFormats(targetSheet, estimatedColumns, formats, fundedRange);
  updateColumnNumberFormats(targetSheet, estimatedColumns, formats, nonFundedRange);
  updateCompanyGrowthCharts(targetSheet, [fundedRange, nonFundedRange], estimatedColumns, currentOptions);
}

/**
 * @param {boolean[]} includeYears An array of booleans representing which years should be included
 * @param {boolean} includeEmptyData Show companies with empty data in all of the years that are `true` in `includeYears`
 * @returns {Object.<string, function(*): boolean>[]}
 */
function generateYearlyFilterConfigs(includeYears, includeEmptyData) {
  const filterConfigs = [];
  for (let i = 1; i <= includeYears.length; i++) {
    if (includeYears[i - 1]) {
      filterConfigs.push({
        [`Year${i}Difference`]: (value) => (
          includeEmptyData || (!includeEmptyData && value !== "" && value !== null)
        ),
        [`Year${i}Percentage`]: (value) => (
          includeEmptyData || (!includeEmptyData && value !== "" && value !== null)
        ),
      });
    }
  }
  return filterConfigs;
}

/**
 * @param {SpreadsheetApp.Sheet} companyListSheet The Company List sheet
 * @param {Object.<string, number>} companyListColumns The column indices for `companyListSheet`
 * @param {string[]} includeColumns The columns to include. Must include `Company` and `Funded`. The fewer columns in `includeColumns`, the faster the lookup is generated
 * @returns {Map.<string, LookupTableValue>} An object with the keys as Company's (**NOT** CompanyID) and values of what are in `includeColumns`
 */
function createFundedLookup(companyListSheet, companyListColumns, includeColumns) {
  if (includeColumns.includes("CompanyID")) {
    throw "Cannot include CompanyID in the include columns for Funded lookup. Use Company as key instead.";
  }

  /**
   * @type {LookupTableCallback}
   */
  const processData = (lookupTable, row, key, order, columns, includeColumns, keyColumn) => {
    if (lookupTable.get(key) !== undefined) {
      throw "Duplicate company found! Cannot create lookup table. Company found: " + key + ". Please check Company List for duplicates.";
    }
    const newData = {};
    includeColumns.filter(column => column !== keyColumn).forEach(column => newData[column] = row[order.indexOf(columns[column])]);
    lookupTable.set(key, newData);
  };

  const requiredKeys = ["Company", "Funded"];
  const keyColumn = "Company";
  return createLookupTable(companyListSheet, companyListColumns, includeColumns, requiredKeys, keyColumn, processData);
}

/**
 * @param {SpreadsheetApp.Sheet} sourceSheet The source sheet.
 * @param {{CompanyID: number, [key: string]: number}} sourceColumns.
 * @param {string[]} includeColumns A string array representing keys of `sourceColumns`. MUST include `CompanyID`.
 * @param {boolean[]} includeYears An array of booleans representing which years should be included.
 * @param {Map.<string, {Funded: number, [key: string]: number}>} fundedLookup The funded lookup table with the key as the company and the value as.
 * @returns {Map.<string, LookupTableValue>} A map of all the values that are listed within includeColumns. Key is `CompanyID`. Returns an empty map if there are no values
 */
function createGrowthLookup(sourceSheet, sourceColumns, includeColumns, includeYears, fundedLookup) {
  checkVariablesTruthy({ includeYears, fundedLookup });
  checkArrayContainsElements(includeColumns, "CompanyID")
  // No years are included (all are false), so no need to do any API queries
  if (!includeYears.some(year => year)) {
    return new Map();
  }

  // Filters all the years not included in includeYears
  /**
   * @type {Set<string>}
   */
  const excludedColumns = new Set();
  includeYears.forEach((year, index) => {
    if (!year) {
      excludedColumns.add(`Year${index + 1}Difference`);
      excludedColumns.add(`Year${index + 1}Percentage`);
    }
  });

  /**
   * @type {LookupTableCallback}
   */
  const processData = (lookupTable, row, key, order, columns, includeColumns, keyColumn) => {
    if (lookupTable.get(key) !== undefined) {
      throw "Duplicate company found! Cannot create lookup table";
    }
    const newData = {};

    // Ignore funded if it is not part of source columns
    if (columns["Funded"] === undefined) excludedColumns.add("Funded");

    // Get all the values excluding the rows within excludeColumns
    // This ensures that the years that were clicked in the dashboard are not considered
    includeColumns.filter(column => column !== keyColumn && !excludedColumns.has(keyColumn)).forEach(column => newData[column] = row[order.indexOf(columns[column])]);

    // Add Funded to newData if it was not already in it
    if (excludedColumns.has("Funded")) {
      let company = key;
      if (keyColumn === "CompanyID") company = getDerivedFromID(key).company;
      const isFunded = fundedLookup.get(company)["Funded"];
      newData["Funded"] = isFunded;
    }

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
 * @param {GrowthOptionValues} currentOptions The current options within the sheet.
 * @returns {SpreadsheetApp.EmbeddedChart[]} A list of charts that were updated.
 */
function updateCompanyGrowthCharts(sheet, ranges, columns, currentOptions) {
  checkVariablesDefined({ sheet });
  const charts = sheet.getCharts();
  const returnCharts = [];
  charts.forEach((chart, index) => {
    /**
     * @type {UpdateChartOptions}
     */
    const newChartOptions = {
      useFirstColumnAsDomain: true,
      series: columns.filter(column => column.includes("Percentage")).map((column, i) => {
        const label = `${ordinal(i + 1)} Year Growth (%)`;
        return {
          targetAxisIndex: i,
          labelInLegend: label,
          visibleInLegend: currentOptions[`Year${i + 1}`],
        };
      }),
      hAxis: {
        title: "Percent Growth (%)",
        minorGridlines: {
          count: 0
        },
      },
    };

    let chartBuilder = chart.modify();
    const newChart = updateChart(sheet, chartBuilder, newChartOptions);
    returnCharts.push(newChart);
  });
  return returnCharts;
}