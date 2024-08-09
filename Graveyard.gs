// #########
// Welcome to the graveyard! This is all code that is depricated, but I could not bring myself to delete.
// Perhaps you could find a use for it... (be warned, a lot of it is super slow, which is why it was depricated in the first place)
// #########

// #region Autocorrect
function levenshtein(a, b) {
    var tmp;
    if (a.length === 0) { return b.length; }
    if (b.length === 0) { return a.length; }
    if (a.length > b.length) { tmp = a; a = b; b = tmp; }
  
    var i, j, res, alen = a.length, blen = b.length, row = Array(alen);
    for (i = 0; i <= alen; i++) { row[i] = i; }
  
    for (i = 1; i <= blen; i++) {
      res = i;
      for (j = 1; j <= alen; j++) {
        tmp = row[j - 1];
        row[j - 1] = res;
        res = b[i - 1] === a[j - 1] ? tmp : Math.min(tmp + 1, Math.min(res + 1, row[j] + 1));
      }
    }
    return res;
  }
  // #endregion
  
  // #region Budget Filter
  /**
   * @depricated
   */
  function getOldData(sourceSheet, sourceColumns, includeColumns) {
    /**
   * Used for getting the relative index of a column in sourceColumns.
   * Example of how to use order with `sourceColumns` and a valid `column` string:
   * ```javascript
   * row[order.indexOf(sourceColumns[column])]
   * ```
   * @type {number[]}
   */
    const order = includeColumns.map(key => sourceColumns[key]);
    const columnData = order.map(index => {
      // 2 to skip the header row
      const range = sourceSheet.getRange(2, index + 1, sourceSheet.getLastRow() - 1, 1);
      return range.getValues().flat(); // Flatten the 2D array to 1D
    });
  
    // Transpose the fetched data so that each row is an array of column values
    const numRows = columnData[0].length;
    const filteredData = Array.from({ length: numRows }, (_, rowIndex) =>
      columnData.map(column => column[rowIndex])
    );
    return filteredData;
  }
  
  /**
   * @typedef {Object} TableFormattingHeaderOptions
   * @property {boolean|null} addTopBorder Whether the top border should be added to the headers. Use `null` if there should be no change from the already existing value
   * @property {boolean|null} addBottomBorder Whether the bottom border should be added to the headers. Use `null` if there should be no change from the already existing value
   * @property {boolean|null} addLeftBorder Whether the left border should be added to the headers. Use `null` if there should be no change from the already existing value
   * @property {boolean|null} addRightBorder Whether the right border should be added to the headers. Use `null` if there should be no change from the already existing value
   * @property {boolean|null} addVerticalBorder true for internal vertical borders, false for none, null for no change.
   * @property {boolean|null} addHorizontalBorder true for internal horizontal borders, false for none, null for no change.
   * @property {boolean} bold Whether the headers should be bolded
   * @property {boolean} underline Whether the headers should be underlined
   * @property {boolean} italics Whether the headers shpuld be italicized
   * @property {string} backgroundColor A color code in CSS notation (such as '#ffffff' or 'white'); a null value resets the color.
   */
  
  /**
   * @typedef {Object} TableFormattingFrameOptions
   * @property {string} backgroundColor A color code in CSS notation (such as '#ffffff' or 'white'); a null value resets the color.
   */
  
  /**
   * @depricated
   * @see `updateTableFormatting()`
   */
  function _changeFormatting() {
    const tableFormatting = {
      /**
       * @type {TableFormattingHeaderOptions}
       */
      headers: {
        addTopBorder: false,
        addBottomBorder: true,
        addLeftBorder: false,
        addRightBorder: false,
        addVerticalBorder: null,
        addHorizontalBorder: null,
        bold: true,
        underline: false,
        italics: false,
        backgroundColor: "white",
      },
      /**
       * @type {TableFormattingFrameOptions}
       */
      frame: {
        backgroundColor: "dark gray 2",
      },
    };
    updateTableFormatting(targetSheet, newRange, tableFormatting, true);
  }
  
  /**
   * @depricated
   * Updates the formatting of the cells AROUND `dataRange` 
   * @param {SpreadsheetApp.Sheet} dataSheet The sheet to change the formatting in
   * @param {SpreadsheetApp.Range} dataRange The range of the data
   * @param {Object} formattingConfig The config of the formatting changes to apply around `dataRange`
   * @param {TableFormattingHeaderOptions} formattingConfig.headers The config for the header formatting
   * @param {TableFormattingFrameOptions} formattingConfig.frame The config for the formatting of the outside of the table
   * @param {boolean} areHeadersInRange Whether there are headers within `dataRange`. If there are not, the headers will not be changed
   * @see `_changeFormatting()`
   */
  function updateTableFormatting(dataSheet, dataRange, formattingConfig, areHeadersInRange) {
    checkVariablesDefined({ dataSheet, dataRange });
  
    if (typeof areHeadersInRange !== "boolean") {
      throw "areHeadersInRange must be of type boolean!";
    }
  
    const startRow = dataRange.getRow();
    const endRow = dataRange.getNumRows() - 1;
    const startColumn = dataRange.getColumn();
    const endColumn = dataRange.getNumColumns() - 1;
  
    if (areHeadersInRange && formattingConfig.headers) {
      let lastColumn = dataRange.getLastColumn();
      let headersRange = dataSheet.getRange(startRow, startColumn, 1, lastColumn);
      const headerValues = headersRange.getValues()[0];
      lastColumn = headerValues.indexOf("");
      if (lastColumn !== -1) {
        headersRange = dataSheet.getRange(startRow, startColumn, 1, lastColumn);
        dataSheet.getRange(startRow, lastColumn + 1, 1, headerValues.length - lastColumn).clearFormat();
      }
      // Change the background color
      if (formattingConfig.headers.backgroundColor !== undefined) {
        headersRange.setBackground(formattingConfig.headers.backgroundColor);
      }
  
      // Add the header borders. Null if they are not included in the config.
      headersRange.setBorder(formattingConfig.headers.addTopBorder || null, formattingConfig.headers.addLeftBorder || null, formattingConfig.headers.addBottomBorder || null, formattingConfig.headers.addRightBorder || null, formattingConfig.headers.addVerticalBorder || null, formattingConfig.headers.addHorizontalBorder || null);
  
      // Update the headers style to include the config options
      const headerStyles = SpreadsheetApp.newTextStyle()
        .setBold(formattingConfig.headers.bold || false)
        .setUnderline(formattingConfig.headers.underline || false)
        .setItalic(formattingConfig.headers.italics || false)
        .build();
      headersRange.setTextStyle(headerStyles);
    }
  
    if (formattingConfig.frame) {
      // Checks to ensure that only changeable ranges are changed
      const maxRows = dataSheet.getMaxRows();
      const maxColumns = dataSheet.getMaxColumns();
      const canChangeTop = startRow > 1;
      const canChangeBottom = endRow < maxRows;
      const canChangeLeft = startColumn > 1;
      const canChangeRight = endColumn < maxColumns;
  
      const topRange = canChangeTop ? dataSheet.getRange(startRow - 1, startColumn - 1, 1, dataRange.getNumColumns() + 2) : null;
      const bottomRange = canChangeBottom ? dataSheet.getRange(endRow + 1, startColumn - 1, 1, dataRange.getNumColumns() + 2) : null;
      const leftRange = canChangeLeft ? dataSheet.getRange(startRow, startColumn - 1, dataRange.getNumRows(), 1) : null;
      const rightRange = canChangeRight ? dataSheet.getRange(startRow, endColumn + 1, dataRange.getNumRows(), 1) : null;
  
      // Change the background colors
      if (formattingConfig.frame.backgroundColor !== undefined) {
        if (topRange) topRange.setBackground(formattingConfig.frame.backgroundColor);
        if (bottomRange) bottomRange.setBackground(formattingConfig.frame.backgroundColor);
        if (leftRange) leftRange.setBackground(formattingConfig.frame.backgroundColor);
        if (rightRange) rightRange.setBackground(formattingConfig.frame.backgroundColor);
      }
    }
  }
  
  /**
   * @depricated
   * **Super slow**. Calculates the sum of Differences in Raw Data for every company, and returns a lookup table with the differences.
   * @param {SpreadsheetApp.Sheet} rawDataSheet The rawDataSheet
   * @returns {Object.<string, number>} An object containing a company's CompanyID and the sum of all its differences
   * @see `sumDifferences()`
   */
  function sumDifferencesOld(rawDataSheet) {
    const companyIDValues = rawDataSheet.getRange("G2:G" + rawDataSheet.getLastRow()).getValues().flat().filter(id => id !== "");
    const differences = createIDYearLookupTable(rawDataSheet);
    const uniqueCompanyIDs = [...new Set(companyIDValues)];
    const differenceSums = {};
  
    uniqueCompanyIDs.forEach(uniqueCompanyID => {
      const cohortYear = getDerivedFromID(uniqueCompanyID).cohortYear;
      const years = calculateYearsFromCohortYear(cohortYear);
      const differencesOfUniqueCompany = years.map(year => +differences[`${uniqueCompanyID}-${year}`]);
  
      differenceSums[uniqueCompanyID] = differencesOfUniqueCompany.reduce((sum, value) => sum + (isNaN(value) || value === null ? 0 : value), 0);
    });
  
    return differenceSums;
  }
  
  /**
   * @depricated
   * Calculates the sum of Differences in Raw Data for every company, and returns a lookup table with the differences
   * @param {SpreadsheetApp.Sheet} rawDataSheet The rawDataSheet
   * @returns {Object.<string, number>} An object containing a company's CompanyID and the sum of all its differences
   * @see `sumDifferencesOld()`
   */
  function sumDifferences(rawDataSheet) {
    // Have to use order to get the correct column within row below
    const order = [rawDataColumns.CompanyID, rawDataColumns.Difference].sort();
    const dataRange = rawDataSheet.getRange(String.fromCharCode(65 + rawDataColumns.CompanyID) + "2:" + String.fromCharCode(65 + rawDataColumns.Difference) + rawDataSheet.getLastRow());
    const data = dataRange.getValues();
    // Create a map of all the sums with key = CompanyID and value = the sum of differences
    const lookupTable = {};
    data.forEach(row => {
      const companyID = row[order.indexOf(rawDataColumns.CompanyID)];
      const difference = parseFloat(row[order.indexOf(rawDataColumns.Difference)]) || 0;
      if (companyID !== "" && companyID !== null && companyID !== undefined) {
        if (!lookupTable[companyID]) {
          lookupTable[companyID] = 0;
        }
        lookupTable[companyID] += difference;
      }
    });
    return lookupTable;
  }
  
  /**
   * @depricated
   * Sums all the differences between 
   * @param {SpreadsheetApp.Sheet} rawDataSheet The Raw Data sheet
   * @param {number|undefined} minimum The minimum Difference to get. `undefined` if there is no minimum.
   * @param {number|undefined} maximum The maximum Difference to get. `undefined` if there is no maxmimum.
   * @returns {{companyID: string, differenceSum: number}[]} An array of objects containing a company's CompanyID and the sum of differences 
   */
  function getValuesWithMinAndMax(rawDataSheet, minimum, maximum) {
    const lookupTable = sumDifferences(rawDataSheet);
    const values = [];
    for (const [key, value] of Object.entries(lookupTable)) {
      if ((minimum === undefined || value >= minimum) && (maximum === undefined || value <= maximum)) {
        values.push({ companyID: key, differenceSum: value })
      }
    }
    return values;
  }
  
  /**
   * @depricated 
   * @param {Map.<string, LookupTableValue>} lookupTable The lookup table to get the values from. Key must be `CompanyID`.
   * @param {Object} currentOptions An object with `dashboardOptions` mapped to their respective values.
   * @param {boolean} currentOptions.ShowAllYears
   * @param {boolean} currentOptions.ShowMostRecentYears
   * @param {string} currentOptions.BudgetFilter
   * @param {boolean} currentOptions.EmptyData
   * @param {boolean} currentOptions.ShowFunded
   * @param {boolean} currentOptions.ShowNonFunded
   * @returns {Array.<{CompanyID: string, [key: string]: any}>} An array of objects containing a company's CompanyID and other key values within `lookupTable`. Only keeps those that hold true with `singularFilterConfig` and `complexFilterConfigs`
   * @see `getValuesWithFilter()` 
   */
  function _getValuesWithFilter(lookupTable, currentOptions) {
    const filtersImplementedElsewhere = [currentOptions.ShowAllYears, currentOptions.ShowMostRecentYears];
    const orFilters = [[currentOptions.ShowFunded, currentOptions.ShowNonFunded]];
    const andFilters = [[currentOptions.BudgetFilter]];
    const notYetImplemented = [currentOptions.EmptyData];
  
    const numFiltersElsewhere = filtersImplementedElsewhere.length;
    const numAndFilters = andFilters.length;
    const numOrFilters = orFilters.length;
    const numNotImplementedFilters = notYetImplemented.length;
    const values = [];
    lookupTable.forEach((value, key) => {
      // let includeEntry = false;
      let filtersMatched = 0;
      Object.keys(value).forEach(individualKey => {
        switch (individualKey) { // Change from Difference to something else to change what the Budget Filter dropdown affects
          case "Difference":
            if ((minimum === undefined || value[individualKey] >= minimum) && (maximum === undefined || value[individualKey] <= maximum)) {
              // includeEntry = true;
              filtersMatched++;
            }
            break;
          case "Funded":
            if (currentOptions.ShowFunded && value[individualKey]) {
              Logger.log(value)
              // includeEntry = true;
              filtersMatched++;
            }
            if (currentOptions.ShowNonFunded && !value[individualKey]) {
              // includeEntry = true;
              filtersMatched++;
            }
            break;
          default:
            break;
        }
      });
  
      // if (includeEntry) {
      //   values.push({ CompanyID: key, ...value });
      // }
      if (filtersMatched - numAndFilters - numOrFilters >= 0) {
        values.push({ CompanyID: key, ...value });
      }
    });
  }
  
  /**
   * @depricated
   * @param {SpreadsheetApp.Sheet} sourceSheet The source sheet.
   * @param {RawDataColumns} sourceColumns The column indices for `sourceSheet`.
   * @param {string[]} includeColumns An array of strings representing keys of `sourceColumns`. MUST include `CompanyID`.
   * @param {boolean} mostRecent A boolean representing if the data should only include the most recent year. MUST include `Year` in `includeColumns` if this is true.
   * @returns {Map.<string, LookupTableValue>} A map of all the values that are listed within includeColumns. Key is `CompanyID`.
   */
  function _createREDSumLookup(sourceSheet, sourceColumns, includeColumns, mostRecent) {
    // if (!includeColumns.includes("CompanyID")) {
    //   throw "Must include CompanyID for the RED lookup!";
    // }
    checkArrayContainsElements(includeColumns, "CompanyID")
  
    // Ensure that all columns are valid
    includeColumns.forEach(column => {
      if (sourceColumns[column] === undefined) {
        throw "Unknown key provided: " + column + ". Could not find column.";
      }
    });
  
    if (mostRecent && !includeColumns.includes("Year")) {
      throw "Cannot get most recent year data without Year column. Provide Year column in includeColumns.";
    }
  
    /**
     * Used for getting the relative index of a column in `sourceColumns`.
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
      if (!companyID || companyID === "") return;
      if (!lookupTable.has(companyID)) lookupTable.set(companyID, {});
  
      const previousData = lookupTable.get(companyID);
      if (mostRecent) {
        if (previousData["Year"] === undefined || +previousData["Year"] < +row[order.indexOf(sourceColumns.Year)]) {
          // Does not matter if there is an entry or not. Replace or add it.
          const newData = {};
          // Skips CompanyID since the key is CompanyID. No need to store it
          includeColumns.filter(column => column !== "CompanyID").forEach(column => newData[column] = row[order.indexOf(sourceColumns[column])]);
          lookupTable.set(companyID, newData);
        }
        return;
      }
  
      // mostRecent is false, so sum up all the previous data
      if (Object.keys(previousData).length === 0) {
        // Hasn't been initialized yet
        const newData = {};
        // Skips CompanyID since the key is CompanyID. No need to store it
        includeColumns.filter(column => column !== "CompanyID").forEach(column => newData[column] = row[order.indexOf(sourceColumns[column])]);
        lookupTable.set(companyID, newData);
        return;
      }
  
      // Add on to the previous data
      Object.keys(previousData).forEach(key => {
        // Skips CompanyID since the key is CompanyID. No need to store it
        if (typeof +previousData[key] === "number" && key !== "CompanyID" && key !== "Cohort Year" && key !== "Year") {
          previousData[key] = +previousData[key] + +row[order.indexOf(sourceColumns[key])];
        }
      });
      lookupTable.set(companyID, previousData);
    });
    return lookupTable;
  }
  // #endregion
  
  // #region New Data
  /**
   * @depricated
   * Not using this anymore, but will leave the code in.
   * To use, create a trigger using:
   * ScriptApp.newTrigger('updateNamedRanges').forSpreadsheet(activeSpreadSheet).onChange().create();
   * @param {Event} e The onChange event
   */
  function updateNamedRanges(e) {
    // Ensure that it only updates the ranges when rows are added
    if (e.changeType !== 'INSERT_ROW') {
      return;
    }
  
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName("Raw Data");
    const namedRanges = ["Company", "CohortYear", "Year", "Revenue", "Expenses", "Difference", "CompanyID"];
  
    // Do updates in batch to improve performance
    const updates = [];
    for (const namedRangeString of namedRanges) {
      const namedRange = spreadsheet.getRangeByName(namedRangeString);
  
      if (namedRange) {
        // The range exists
  
        const lastRow = rawDataSheet.getLastRow();
        const newRange = rawDataSheet.getRange(2, namedRange.getColumn(), lastRow, namedRange.getWidth());
  
        // Add the range to the batch
        updates.push({ name: namedRangeString, range: newRange });
      } else {
        // The named range was not found
        Logger.log("Unknown named range accessed! Range: " + namedRangeString);
      }
    }
  
  
    // Update all the ranges
    updates.forEach(update => {
      spreadsheet.setNamedRange(update.name, update.range);
    });
  }
  
  /**
   * @depricated As of July 19, 2024
   * Uses the following from `rawDataColumns`:
   */
  function updateDerivedFields() {
    /*
    ###################
    TODO
    ###################
    Make update add new data to Budget Filter as well (without having to recalculate all of the data)
    */
  
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName("Raw Data");
    const dependentSheetNames = ["Relative Year Summary"];
  
    const dependentSheets = dependentSheetNames.map(sheetName => ({
      name: sheetName,
      sheet: spreadsheet.getSheetByName(sheetName),
      idColumn: "A",
      companyColumn: "B",
      yearColumn: "C",
      startingRow: 2, // Important: Indexing starts at 1
    }));
  
    // An example of overriding the defaults would be
    const RYS = dependentSheets.find(dependentSheet => dependentSheet.name === "Relative Year Summary");
    RYS.startingRow = 3;
  
    // Here is an example of how to rearange the data
    // const mySheet = dependentSheets.find(dependentSheet => dependentSheet.name === "My Sheet");
    // mySheet.yearColumn = "A";
    // mySheet.idColumn = "B";
    // mySheet.companyColumn = "C";
    // mySheet.startingRow = 5;
  
    updateCompanyIDsAndDerivedFields(rawDataSheet, rawDataColumns, dependentSheets);
    updateYearValues();
  }
  // #endregion
  
  // #region Dashboard Utilities
  /** 
   * @descripton For a vAxis in UpdateChartOptions
   * @property {number | undefined} vAxis.direction The direction in which the values along the vertical axis grow. By default, low values are on the bottom of the chart. Specify `-1` to reverse the order of the values. **Note**: Only takes in `1` or `-1`.
   * @property {Object | null | undefined} vAxis.gridlines An object with members to configure the gridlines on the vertical axis. Note that vertical axis gridlines are drawn horizontally.
   * @property {string | undefined} vAxis.gridlines.color The color of the vertical gridlines inside the chart area. Specify a valid HTML color string.
   * @property {number | undefined} vAxis.gridlines.count The approximate number of horizontal gridlines inside the chart area. If you specify a positive number for `gridlines.count`, it will be used to compute the `minSpacing` between gridlines. You can specify a value of `1` to only draw one gridline, or `0` to draw no gridlines. Specify `-1`, which is the default, to automatically compute the number of gridlines based on other options.
   * @property {boolean | undefined} vAxis.logScale If `true`, makes the vertical axis a logarithmic scale. **Note**: All values must be positive.
   * @property {number | undefined} vAxis.maxValue Moves the max value of the vertical axis to the specified value; this will be upward in most charts. Ignored if this is set to a value smaller than the maximum y-value of the data. `vAxis.viewWindow.max` overrides this property.
   * @property {Object} vAxis.minorGridlines An object with members to configure the minor gridlines on the vertical axis, similar to the vAxis.gridlines option.
   * @property {string | null | undefined} vAxis.minorGridlines.color The color of the vertical minor gridlines inside the chart area. Specify a valid HTML color string.
   * @property {number | undefined} vAxis.minorGridlines.count The `minorGridlines.count` option is mostly deprecated, except for disabling minor gridlines by setting the count to `0`. The number of minor gridlines depends on the interval between major gridlines and the minimum required space.
   * @property {number | null | undefined} vAxis.minValue Moves the min value of the vertical axis to the specified value; this will be downward in most charts. Ignored if this is set to a value greater than the minimum y-value of the data. `vAxis.viewWindow.min` overrides this property.
   * @property {string | null | undefined} vAxis.textPosition Position of the vertical axis text, relative to the chart area. Supported values: 'out', 'in', 'none'.
   * @property {TextStyle | undefined} vAxis.textStyle An object that specifies the vertical axis text style.
   * @property {string | null | undefined} vAxis.title Specifies a title for the vertical axis. Set to `null` to leave as is, and `undefined` to remove.
   * @property {TextStyle | undefined} vAxis.titleTextStyle An object that specifies the vertical axis title text style.
   * @property {Object | null | undefined} vAxis.viewWindow Specifies the cropping range of the vertical axis.
   * @property {number | null | undefined} vAxis.viewWindow.max The maximum vertical data value to render. Ignored when `vAxis.viewWindowMode` is 'pretty' or 'maximized'.
   * @property {number | null | undefined} vAxis.viewWindow.min The minimum vertical data value to render. Ignored when `vAxis.viewWindowMode` is 'pretty' or 'maximized'.
   */
  { }; // used to hide from any intellisense
  // #endregion
  
  // #region Company Growth Dashboard
  /**
   * @depricated
   * @param {SpreadsheetApp.Sheet} companyListSheet The Company List sheet
   * @param {Object.<string, number>} companyListColumns The column indices for `companyListSheet`
   * @param {string[]} includeColumns The columns to include. Must include `Company` and `Funded`. The fewer columns in `includeColumns`, the faster the lookup is generated
   * @returns {Object.<string, LookupTableValue>} An object with the keys as Company's (**NOT** CompanyID) and values of what are in `includeColumns`
   * @see `createFundedLookup`
   */
  function _createFundedLookup(companyListSheet, companyListColumns, includeColumns) {
    checkVariablesDefined(includeColumns);
    checkArrayContainsElements(includeColumns, "Company", "Funded");
  
    if (includeColumns.includes("CompanyID")) {
      throw "Cannot include CompanyID in the include columns for Funded lookup. Use Company as key instead.";
    }
  
    includeColumns.forEach(column => {
      if (companyListColumns[column] === undefined) {
        throw "Unknown column key provided: " + column + ". Check the company list columns and include columns.";
      }
    });
  
    checkVariablesDefined({ companyListSheet, companyListColumns });
    if (companyListColumns["Company"] === undefined) {
      throw "Could not find Company column in companyListColumns!";
    }
    if (companyListColumns["Funded"] === undefined) {
      throw "Could not find Funded column in companyListColumns!";
    }
  
    /**
     * Used for getting the relative index of a column in `companyListColumns`.
     * Example of how to use order with `companyListColumns` and a valid `column` string:
     * ```javascript
     * row[order.indexOf(companyListColumns[column])]
     * ```
     * @type {number[]}
     */
    const order = includeColumns.map(key => companyListColumns[key]);
    const filteredData = getDataInOrder(companyListSheet, companyListColumns, includeColumns, order);
  
    /**
     * @type {Object.<string, LookupTableValue>}
     */
    const lookupTable = {};
    filteredData.forEach(row => {
      const company = row[order.indexOf(companyListColumns.Company)];
      if (!company || company === "") return;
      if (lookupTable[company] !== undefined) {
        throw "Duplicate company found! Cannot create lookup table";
      }
      const newData = {};
      includeColumns.filter(column => column !== "Company").forEach(column => newData[column] = row[order.indexOf(companyListColumns[column])]);
      // lookupTable.set(company, newData);
      lookupTable[company] = newData;
    });
  
    return lookupTable;
  }
  
  /**
   * @depricated
   * @param {SpreadsheetApp.Sheet} sourceSheet The source sheet
   * @param {{CompanyID: number, [key: string]: number}} sourceColumns
   * @param {string[]} includeColumns A string array representing keys of `sourceColumns`. MUST include `CompanyID`
   * @param {boolean[]} includeYears An array of booleans representing which years should be included
   * @param {Object.<string, {Funded: number, [key: string]: number}>} fundedLookup The funded lookup table with the key as the company and the value as 
   * @returns {Map.<string, LookupTableValue>} A map of all the values that are listed within includeColumns. Key is `CompanyID`. Returns an empty map if there are no values
   * @see `createGrowthLookup`
   */
  function _createGrowthLookup(sourceSheet, sourceColumns, includeColumns, includeYears, fundedLookup) {
    checkVariablesTruthy({ includeColumns, includeYears, fundedLookup });
    checkArrayContainsElements(includeColumns, "CompanyID")
    // No years are included (all are false), so no need to do any API queries
    if (!includeYears.some(year => year)) {
      return new Map();
    }
  
    // Ensure that all columns are valid
    includeColumns.forEach(column => {
      if (sourceColumns[column] === undefined && column !== "Funded") {
        throw "Unknown column key provided: " + column + ". Check the source columns and include columns.";
      }
    });
  
    // Have to use order to get the correct column within row below
    const order = includeColumns.map(key => sourceColumns[key]);
    const filteredData = getDataInOrder(sourceSheet, sourceColumns, includeColumns, order);
  
  
    // Create a map of all the sums with key = CompanyID and value = an object containing all the information
    /**
     * @type {Map.<string, Object.<string, any>>}
     */
    const lookupTable = new Map();
    filteredData.forEach(row => {
      const companyID = row[order.indexOf(sourceColumns.CompanyID)];
      if (!companyID || companyID === "") return;
      if (lookupTable.has(companyID)) {
        throw "Duplicate CompanyID found. Cannot continue creating lookup table";
      }
      const newData = {};
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
  
      // Ignore funded if it is not part of source columns
      if (sourceColumns["Funded"] === undefined) excludedColumns.add("Funded");
  
      // Get all the values excluding the rows within excludeColumns
      // This ensures that the years that were clicked in the dashboard are not considered
      includeColumns.filter(column => column !== "CompanyID" && !excludedColumns.has(column)).forEach(column => newData[column] = row[order.indexOf(sourceColumns[column])]);
  
      // Add Funded to newData if it was not already in it
      if (excludedColumns.has("Funded")) {
        const company = getDerivedFromID(companyID).company;
        const isFunded = fundedLookup[company]["Funded"];
        newData["Funded"] = isFunded;
      }
  
      lookupTable.set(companyID, newData);
    });
    return lookupTable;
  }
  // #endregion
  // #region Relative Year Summary (RYS)
  /**
   * @depricated As of July 19, 2024
   * Uses the following properties from `rawDataColumns`:
   * - CompanyID
   * - Year
   * - Difference
   * @private
   * @param {SpreadsheetApp.Sheet} rawDataSheet The Raw Data sheet
   * @returns {Map.<string, number>} A lookup table with the key being `CompanyID-Year` and the value being the `Difference` value of the company at that year
   * @see `createIDYearLookupTable()`
   */
  function _createIDYearLookupTable(rawDataSheet) {
    // Note: This does not use createLookupTable, and I am too lazy to update it to do so.
    const rawData = rawDataSheet.getDataRange().getValues();
    const lookupTable = new Map();
    // Start at 1 to skip headers
    const rawDataLength = rawData.length; // Used to improve performance
    for (let i = 1; i < rawDataLength; i++) {
      const companyID = rawData[i][rawDataColumns.CompanyID];
      const year = rawData[i][rawDataColumns.Year];
      const difference = rawData[i][rawDataColumns.Difference];
      // Less readable without but slightly (and I mean SLIGHTLY) faster
      // const key = `${companyID}-${year}`;
      lookupTable.set(`${companyID}-${year}`, difference);
    }
    return lookupTable;
  }
  
  /*
  Originally used the following spreadsheet formula:
  ARRAYFORMULA(IF(A3:A <> "", GET_DIFFERENCE_FROM_YEAR(A3:A, D3:D), ""))
  With A3:A being the CompanyID column, D3:D being Year 1 column, and
  GET_DIFFERENCE_FROM_YEAR being
  IFERROR(
          VLOOKUP(id & "-" & (year_range),
          {CompanyID & "-" & Year, Difference},
          2, FALSE),
        ""
        )
  */
  
  /**
   * @depricated As of July 19, 2024
   * Gets all the unique CompanyID's from `sourceSheet` and puts them into the `targetSheets`. Also adds Company and Cohort Year.
   * @param {SpreadsheetApp.Sheet} sourceSheet The source sheet for the data
   * @param {Object.<string, number>} sourceColumns The indices of the columns of the source sheet
   * @param {number} sourceColumns.CompanyID The index of the CompanyID column in `sourceSheet `
   * @param {Object[]} targetSheets A non-empty array of objects to add the data to
   * @param {string} targetSheets[].name The name of the target sheet
   * @param {SpreadsheetApp.Sheet} targetSheets[].sheet The target sheet
   * @param {string} targetSheets[].idColumn The A1 notation column for CompanyID. Defaults to "A"
   * @param {string} targetSheets[].companyColumn The A1 notation column for Company. Defaults to "B"
   * @param {string} targetSheets[].yearColumn The A1 notation column letter for Cohort Year. Defaults to "C"
   * @param {number} targetSheets[].startingRow The row to start inserting data at. Defaults to 2
   */
  function _updateCompanyIDsAndDerivedFields(sourceSheet, sourceColumns, targetSheets) {
    checkVariablesTruthy({ sourceSheet, sourceColumns, targetSheets });
    if (targetSheets.length === 0) {
      throw "Target sheets must contain at least one item";
    }
  
    // Use A1 notation to not call updateNamedRanges
    // TODO: Get A1 notation from the index
    // What was the point of this TODO? I don't remember
    const companyIDColumn = sourceSheet.getRange("G2:G" + sourceSheet.getLastRow());
  
    // Get all non-empty ids
    // getDataValues() might work better here
    const ids = companyIDColumn.getValues().flat().filter(id => id !== "");
  
    // Leverage a Set's property of having only unique values
    const uniqueIDs = [...new Set(ids)];
  
    /**
     * Generates the company and cohort year data
     * @type {(string | number)[][]}
     */
    const uniqueData = uniqueIDs.map(id => {
      const [company, cohortYear] = id.split('-$-');
      return [id, company, cohortYear];
    });
  
    targetSheets.forEach(targetSheet => {
      // Object properties checks
      checkObjectPropertiesDefined(targetSheet, ["name", "sheet", "idColumn", "companyColumn", "yearColumn", "startingRow"]);
  
      // Do the ranges in bulk (fast)
      if (targetSheet.idColumn === "A" && targetSheet.companyColumn === "B" && targetSheet.yearColumn === "C") {
        // Remove the old IDs
        // This is an expensive operation since it requires reentering all the unique ids
        targetSheet.sheet.getRange("A" + String(targetSheet.startingRow) + ":C" + targetSheet.sheet.getLastRow()).clear();
  
        targetSheet.sheet.getRange(targetSheet.startingRow, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
      } else {
        const columns = [targetSheet.idColumn, targetSheet.companyColumn, targetSheet.yearColumn];
        // Do each range separately (slow)
        for (let i = 0; i < 3; i++) {
          const range = columns[i] + targetSheet.startingRow + ":" + columns[i];
          targetSheet.sheet.getRange(range + targetSheet.sheet.getLastRow()).clear();
  
          // 65 is ASCII value of A
          // Index starts at 1 so add 1 to starting column
          targetSheet.sheet.getRange(targetSheet.startingRow, columns[i].charCodeAt(0) - 65 + 1, uniqueData.length, 1).setValues(uniqueData.map(data => [data[i]]));
        }
      }
    });
  }
  
  /**
   * @depricated As of July 19, 2024
   * Gets all the unique CompanyID's from `sourceSheet` and puts them into the `targetSheets`. Also adds Company and Cohort Year.
   * @param {SpreadsheetApp.Sheet} sourceSheet The source sheet for the data
   * @param {Object.<string, number>} sourceColumns The indices of the columns of the source sheet
   * @param {number} sourceColumns.CompanyID The index of the CompanyID column in `sourceSheet `
   * @param {Object[]} targetSheets A non-empty array of objects to add the data to
   * @param {string} targetSheets[].name The name of the target sheet
   * @param {SpreadsheetApp.Sheet} targetSheets[].sheet The target sheet
   * @param {string} targetSheets[].idColumn The A1 notation column for CompanyID. Defaults to "A"
   * @param {string} targetSheets[].companyColumn The A1 notation column for Company. Defaults to "B"
   * @param {string} targetSheets[].yearColumn The A1 notation column letter for Cohort Year. Defaults to "C"
   * @param {number} targetSheets[].startingRow The row to start inserting data at. Defaults to 2
   * @see `updateYearValues()`
   */
  function updateCompanyIDsAndDerivedFields(sourceSheet, sourceColumns, ...targetSheets) {
    checkVariablesTruthy({ sourceSheet, sourceColumns, targetSheets });
    if (targetSheets.length === 0) {
      throw "Target sheets must contain at least one item";
    }
  
    // Use A1 notation to not call updateNamedRanges
    // TODO: Get A1 notation from the index
    // What was the point of this TODO? I don't remember
    const companyIDColumn = sourceSheet.getRange("G2:G" + sourceSheet.getLastRow());
  
    // Get all non-empty ids
    // getDataValues() might work better here
    const ids = companyIDColumn.getValues().flat().filter(id => id !== "");
  
    // Leverage a Set's property of having only unique values
    const uniqueIDs = [...new Set(ids)];
  
    /**
     * Generates the company and cohort year data
     * @type {(string | number)[][]}
     */
    const uniqueData = uniqueIDs.map(id => {
      const [company, cohortYear] = id.split('-$-');
      return [id, company, cohortYear];
    });
  
    targetSheets.forEach(targetSheet => {
      // Object properties checks
      checkObjectPropertiesDefined(targetSheet, ["name", "sheet", "idColumn", "companyColumn", "yearColumn", "startingRow"]);
  
      if (targetSheet.idColumn === "A" && targetSheet.companyColumn === "B" && targetSheet.yearColumn === "C") {
        // Remove the old IDs
        // This is an expensive operation since it requires reentering all the unique ids
        targetSheet.sheet.getRange("A" + String(targetSheet.startingRow) + ":C" + targetSheet.sheet.getLastRow()).clear();
  
        targetSheet.sheet.getRange(targetSheet.startingRow, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
        return;
      }
  
      // For the case they are out of order, use this instead
      const order = [...[targetSheet.idColumn, targetSheet.companyColumn, targetSheet.yearColumn].map(col => col.charCodeAt(0) - 65 + 1)];
      getDataInOrder(sourceSheet, sourceColumns, ["CompanyID", "Company", "CohortYear"], order);
      targetSheet.sheet.getRange(targetSheet.startingRow, Math.min(...order), uniqueData.length, order.length);
    });
  }
  
  /**
   * @depricated as of July 19, 2024
   * Updates Year 1-`NUMBER_OF_YEARS` Difference and Percentage values in Relative Year Summary
   * 
   * Uses the following properties from `rawDataColumns`:
   * - CompanyID
   * - CohortYear
   * - Year
   * - Difference
   * @param {boolean | undefined} [doCarryOver = undefined] Whether the Relative Year Summary should carry over its values. This means adding the inverse of the difference from the previous year to the next year.
   * @see `updateYearValues()`
   */
  function _updateYearValues(doCarryOver = undefined) {
    /* 
    Originally used the following spreadsheet formula:
    ARRAYFORMULA(IF(A3:A <> "", GET_DIFFERENCE_FROM_YEAR(A3:A, D3:D), ""))
    With A3:A being the CompanyID column, D3:D being Year 1 column, and
    GET_DIFFERENCE_FROM_YEAR being 
    IFERROR(
            VLOOKUP(id & "-" & (year_range), 
            {CompanyID & "-" & Year, Difference}, 
            2, FALSE),
          ""
          )
    */
  
    // Originally had 3 columns that represented Year 1, 3, and 4
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName("Raw Data");
    const RYS = spreadsheet.getSheetByName("Relative Year Summary");
    const RYSColumns = generateRYSColumns();
  
    const lookupTable = createIDYearLookupTable(rawDataSheet, doCarryOver);
  
    const RYSData = RYS.getDataRange().getValues();
  
    // Process each row in RYSData
    for (let i = 2; i < RYSData.length; i++) {
      const companyID = RYSData[i][RYSColumns.CompanyID];
      const cohortYear = RYSData[i][RYSColumns.CohortYear];
      const years = calculateYearsFromCohortYear(cohortYear);
  
      // Calculate differences and percentages
      for (let j = 0; j < years.length - 1; j++) {
        const yearValue1 = lookupTable.get(`${companyID}-${years[j]}`);
        const yearValue2 = lookupTable.get(`${companyID}-${years[j + 1]}`);
  
        // Desctructure difference and percentage
        const { difference, percentage } = calculateDifferenceAndPercentage(yearValue1, yearValue2);
  
        RYSData[i][RYSColumns[`Year${j + 1}Difference`]] = difference;
        RYSData[i][RYSColumns[`Year${j + 1}Percentage`]] = percentage;
      }
    }
    // length - 2 to skip the headers
    RYS.getRange(3, 1, RYSData.length - 2, RYSData[0].length).setValues(RYSData.slice(2));
  }
  
  
  // #endregion