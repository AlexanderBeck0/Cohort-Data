// #####
// Misc. helper functions for throughout the code. Note that these are as optimized as possible, since they are used at any point throughout the code.
// #####

// Global type definitions
/**
 * @typedef {Object} RawDataColumns
 * @property {number} Company The non-negative index of the Company column in Raw Data
 * @property {number} CohortYear The non-negative index of the Cohort Year column in Raw Data
 * @property {number} Year The non-negative index of the Year column in Raw Data
 * @property {number} Revenue The non-negative index of the Revenue column in Raw Data
 * @property {number} Expenses The non-negative index of the Expenses column in Raw Data
 * @property {number} Difference The non-negative index of the Difference column in Raw Data
 * @property {number} CompanyID The non-negative index of the Company ID column in Raw Data
 * @property {number} CleanedCompany The non-negative index of the Cleaned Company column in Raw Data
 * @property {number} Funded The non-negative index of the Funded column in Raw Data
 */

/**   
 * @typedef {Object} SortByOption
 * @property {string} name The name of the sorting option
 * @property {(a: {CompanyID: string, [key: string]: string | number | boolean}, b: {CompanyID: string, [key: string]: string | number | boolean}) => number} sortFunction A function that determines the order of the elements. It should return a number where:
 * - A negative value indicates that `a` should come before `b`.
 * - A positive value indicates that `a` should come after `b`.
 * - Zero or `NaN` indicates that `a` and `b` are considered equal.
 */

/**
 * @typedef LookupTableValue
 * @type {Object.<string, string|boolean|number|Date>>}
 */

/**
 * @typedef Format
 * @type {'none' | 'decimal' | 'currency' | 'currency_rounded' | 'percent'}
 * Supports the following types:
 * - `'none'`: displays numbers with no formatting (e.g., 8000000)
 * - `'decimal'`: displays numbers with thousands separators (e.g., 8,000,000)
 * - `'currency'`: displays numbers in USD (e.g., $8,000,000.00)
 * - `'currency_rounded'`: displays rounded numbers in USD (e.g. 80.2 -> $80) 
 * - `'percent'`: displays numbers as percentages (e.g., 800,000,000%)
 */

const FORMAT_VALUES = Object.freeze({
  "none": '0.###############',
  "decimal": '#,##0.00',
  "currency": '"$"#,##0.00',
  "currency_rounded": '"$"#,##0',
  "percent": '0.00%'
});

// Helper functions

/**
 * An ordinal is "1st", "2nd", "3rd", etc.
 * @param {number} n The number to convert to an ordinal. Only supports 0-99.
 * @returns {string} The ordinal of the number.
 * @see [Source](https://gist.github.com/jlbruno/1535691)
 */
function ordinal(n) {
  var s = ["th", "st", "nd", "rd"];
  var v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

/**
 * Converts column index to A1 notation (e.g., 1 -> A, 2 -> B, 27 -> AA).
 * @param {number} index The index to convert into A1 notation.
 * @returns {string} The column index in A1 notation (e.g., 1 -> A, 2 -> B, 27 -> AA).
 */
function columnIndexToA1(index) {
  let columnLetter = '';
  while (index > 0) {
    let remainder = (index - 1) % 26;
    columnLetter = String.fromCharCode(remainder + 65) + columnLetter;
    index = Math.floor((index - 1) / 26);
  }
  return columnLetter;
}

/**
 * @param {string} companyID The CompanyID in the form `Company-$-CohortYear`
 * @returns {{company: string, cohortYear: string}} The company and cohort year of `companyID`
 */
function getDerivedFromID(companyID) {
  const [company, cohortYear] = companyID.split("-$-");
  return { company: company, cohortYear: cohortYear };
}

/**
 * @param {number} cohortYear The cohort year to calculate the different years from. Non-negative.
 * @returns {number[]} The years that a company is in a cohort. [`cohortYear - 1`, `cohortYear`, `cohortYear + 1`, ..., `cohortYear + NUMBER_OF_YEARS`]
 * @see `NUMBER_OF_YEARS`
 */
function calculateYearsFromCohortYear(cohortYear) {
  if (+cohortYear < 0) {
    throw "Cohort Year cannot be less than 0!";
  }

  const years = [+cohortYear - 1];
  for (let i = 0; i < NUMBER_OF_YEARS - 1; i++) {
    years.push(+cohortYear + i);
  }
  return years;
  // return [cohortYear - 1, cohortYear, cohortYear + 1, cohortYear + 2, cohortYear + 3]
};

/**
 * @param {number} firstValue The first value
 * @param {number} secondValue The second value
 * @param {string} [defaultValue=""] The value to put if at least one of the two values is undefined or this value
 * @returns {number|string} `defaultValue` if either `firstValue` or `secondValue` are undefined/equal to `defaultValue`, or `firstValue - secondValue` if they are both numbers
 */
function calculateDifference(firstValue, secondValue, defaultValue = "") {
  let difference = defaultValue;
  const isNotUndefined = firstValue !== undefined && secondValue !== undefined;
  const bothAreNotEmpty = firstValue !== defaultValue && secondValue !== defaultValue;
  if (isNotUndefined && bothAreNotEmpty) {
    difference = firstValue - secondValue;
  }
  return difference;
}

/**
 * Copies `numberOfRows` rows and `numberOfColumns` columns from `sourceSheet` starting at `sourceRowIndex` in column `sourceColumnIndex`, and copying it to `targetSheet` starting at `targetRowIndex` in column `targetColumnIndex`. Turns out that `copyTo()` is already a function, but oh well.
 * @param {SpreadsheetApp.Sheet} sourceSheet The sheet to get the data from.
 * @param {number} [sourceRowIndex = sourceSheet.getLastRow()] The non-negative row index to start copying data from. Index starts at 1. Defaults to the last row with data in the sheet.
 * @param {number} [sourceColumnIndex = 1] The non-negative column index to start copying data from. Index starts at 1. Defaults to 1.
 * @param {SpreadsheetApp.Sheet} targetSheet the sheet to add the data to.
 * @param {number} [targetRowIndex = targetSheet.getLastRow() + 1] The non-negative row index to start adding the data to. Index starts at 1. Defaults to targetSheet.getLastRow() + 1.
 * @param {number} [targetColumnIndex = 1] The non-negative column index to start adding the data to. Index starts at 1. Defaults to 1.
 * @param {number} [numberOfRows = 1] The non-negative number of rows to copy. Defaults to 1.
 * @param {number} [numberOfColumns = sourceSheet.getLastColumn()] The non-negative number of columns to copy. Defaults to sourceSheet.getLastColumn().
 * @param {boolean} [includeTimestamp = true] Whether there should be a Timestamp column appended to the end of the data to copy.
 */
function copyRowsToSheet(sourceSheet, sourceRowIndex, sourceColumnIndex, targetSheet, targetRowIndex, targetColumnIndex = 1, numberOfRows = 1, numberOfColumns = undefined, includeTimestamp = true) {
  /*if (sourceSheet === undefined) {
    throw "sourceSheet cannot be undefined!";
  }
  if (targetSheet === undefined) {
    throw "targetSheet cannot be undefined!";
  }*/
  checkVariablesDefined({ sourceSheet, targetSheet });

  // Default values
  if (numberOfRows === undefined) {
    numberOfRows = 1;
  }
  if (numberOfColumns === undefined) {
    numberOfColumns = sourceSheet.getLastColumn();
  }
  if (targetColumnIndex === undefined) {
    targetColumnIndex = 1;
  }
  if (sourceColumnIndex === undefined) {
    sourceColumnIndex = 1;
  }
  if (sourceRowIndex === undefined) {
    sourceRowIndex = sourceSheet.getLastRow();
  }
  if (targetRowIndex === undefined) {
    targetRowIndex = targetSheet.getLastRow() + 1;
  }
  if (includeTimestamp === undefined) {
    includeTimestamp = true;
  }
  checkVariablesGreaterThanZero({ sourceRowIndex, sourceColumnIndex, targetRowIndex, targetColumnIndex });

  let rowsToCopy = sourceSheet.getRange(sourceRowIndex, sourceColumnIndex, numberOfRows, numberOfColumns).getValues();

  if (includeTimestamp) {
    const timeZone = 'America/New_York'
    const currentDate = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");
    rowsToCopy.forEach(row => row.push(currentDate));
  }

  targetSheet.getRange(targetRowIndex, targetColumnIndex, rowsToCopy.length, rowsToCopy[0].length).setValues(rowsToCopy);
}

/**
 * Gets the row index of `str` in `sheet`, starting at `startIndex`. Returns -1 if there is no match found.
 * @param {string} str The string to look for
 * @param {number} columnIndex A non-negative number. The index of the column to look at. Index starts at 1.
 * @param {SpreadsheetApp.Sheet} sheet The Sheet to look for `str` in
 * @param {number} [startIndex=1] The starting index to start looping from. Defaults to 1. Non-negative. Index starts at 1.
 * @returns {number} The row index of `str`, or -1 if it was not found
 */
function rowIndexOfValue(str, columnIndex, sheet, startIndex = 1) {
  // There likely already exists a function within the Spreadsheets API, but I already implemented this so...
  // Future me here. There does exist this already. Oh well.
  // Parameter validation
  checkVariablesGreaterThanZero({ columnIndex, startIndex });
  const maxColumns = sheet.getMaxColumns(); // Used instead of inline to improve performance
  if (columnIndex > maxColumns) {
    throw "Column Index must be less than the number of columns in the sheet!";
  }
  if (startIndex > maxColumns) {
    throw "Start Index must be less than the number of columns in the sheet!";
  }

  // Start search
  const values = sheet.getRange(1, columnIndex, sheet.getLastRow(), 1).getValues();
  // const values = sheet.getDataRange().getValues();
  /*const valuesLength = values.length;
  for (let i = startIndex; i < valuesLength; i++) {
    const row = values[i];
    if (row[0] === str) {
      return i;
    }
  }
  return -1;*/
  // More efficient version of what is commented out above
  // Feel free to delete, it is there for records sake
  const index = values.slice(startIndex).findIndex(row => row[0] === str);
  return index === -1 ? -1 : index + startIndex;
}

/**
 * Gets the data from `sourceSheet`, with the data in the order of `order`. 
 * If `order` is `[1.0, 3.0, 2.0]`, then the return value would be a 2D array of the values 
 * in `sourceSheet` at the indices of `order` (based on `sourceColumns`).
 * Example of how to use order with `sourceColumns` and a valid `column` string:
 * ```javascript
 * row[order.indexOf(sourceColumns[column])]
 * ```
 * @param {SpreadsheetApp.Sheet} sourceSheet The source sheet.
 * @param {RawDataColumns} sourceColumns The column indices for `sourceSheet`.
 * @param {string[]} includeColumns An array of strings representing keys of `sourceColumns`.
 * @param {number[]} [order = includeColumns.map(key => sourceColumns[key])] The order of the column indices to place the data in.
 * @returns {Object[][]} A 2D array of the values in `sourceSheet` at the indices of `order`.
 */
function getDataInOrder(sourceSheet, sourceColumns, includeColumns, order = undefined) {
  if (order === undefined) {
    /**
     * Used for getting the relative index of a column in `sourceColumns`.
     * Example of how to use order with `sourceColumns` and a valid `column` string:
     * ```javascript
     * row[order.indexOf(sourceColumns[column])]
     * ```
     * @type {number[]}
     */
    order = includeColumns.map(key => sourceColumns[key]);
  }

  const minColumn = Math.min(...order);
  const maxColumn = Math.max(...order);
  const numColumns = maxColumn - minColumn + 1;

  const range = sourceSheet.getRange(2, minColumn + 1, sourceSheet.getLastRow() - 1, numColumns);
  const values = range.getValues();

  const columnData = order.map(index => values.map(row => row[index - minColumn]));
  // Transpose the fetched data so that each row is an array of column values
  const numRows = columnData[0].length;
  const filteredData = Array.from({ length: numRows }, (_, rowIndex) =>
    columnData.map(column => column[rowIndex])
  );
  return filteredData;
}


/**
 * The callback to use for creating a lookup table. **Note**: This is looped through, so it is highly recommended to not use any loops within. 
 * @callback LookupTableCallback
 * @param {Map.<string, LookupTableValue>} lookupTable The lookup table with the key as `key`. Is modified.
 * @param {LookupTableValue} row The particular row of data.
 * @param {string} key They key to use for the lookup table.
 * @param {number[]} order Used for getting the relative index of a column in `columns`. \
  * Example of how to use order with `sourceColumns` and a valid `column` string:
  * ```javascript
  * row[order.indexOf(columns[column])]
  * ```
 * @param {Object.<string, number>} columns The column indices.
 * @param {string[]} includeColumns The columns to include, in the provided order.
 * @param {string} keyColumn The column that holds the key value.
 */

/**
 * Utility function to streamline the creation of lookup tables
 * @param {SpreadsheetApp.Sheet} sheet The source sheet.
 * @param {Object.<string, number>} columns The column indices for `sheet`.
 * @param {string[]} includeColumns The columns to include.
 * @param {string[]} requiredKeys Keys that must be included in `includeColumns`.
 * @param {string} keyColumn The column name to use as the key for the lookup table.
 * @param {LookupTableCallback} processDataCallback A callback function to process each row and add data to the lookup table.
 * @returns {Map.<string, LookupTableValue>} The lookup table.
 */
function createLookupTable(sheet, columns, includeColumns, requiredKeys, keyColumn, processDataCallback) {
  checkVariablesDefined({ sheet, columns, includeColumns, keyColumn });
  checkArrayContainsElements(Object.keys(columns), keyColumn, ...requiredKeys);
  checkArrayContainsElements(includeColumns, ...requiredKeys);

  // Ensure that all columns are valid
  includeColumns.forEach(column => {
    if (columns[column] === undefined) {
      throw new Error(`Unknown column key provided: ${column}. Check the columns and include columns.`);
    }
  });

  /**
  * Used for getting the relative index of a column in `sourceColumns`.
  * Example of how to use order with `sourceColumns` and a valid `column` string:
  * ```javascript
  * row[order.indexOf(sourceColumns[column])]
  * ```
  * @type {number[]}
  */
  const order = includeColumns.map(key => columns[key]);
  const filteredData = getDataInOrder(sheet, columns, includeColumns, order);

  /**
   * @type {Map.<string, LookupTableValue>}
   */
  const lookupTable = new Map();
  filteredData.forEach(row => {
    const key = row[order.indexOf(columns[keyColumn])];
    if (!key || key === "") return;
    processDataCallback(lookupTable, row, key, order, columns, includeColumns, keyColumn);
  });
  return lookupTable;
}

/**
 * Checks if the given variables are defined.
 * 
 * This function takes an object where the keys are variable names and the values are the variables themselves.
 * It throws an error if any of the variables are `undefined`, with a message indicating which variable must be defined.
 * 
 * @param {Object} variables An object containing variables to check.
 * @throws Throws an error if any variable or specified property is `undefined`.
 * 
 * Example:
 * ```javascript
 * checkVariablesDefined({ sourceSheet, targetSheet, startingRow, startingColumn });
 * ```
 * @see `checkObjectPropertiesDefined()`
 */
function checkVariablesDefined(variables) {
  for (const key in variables) {
    if (variables[key] === undefined) {
      throw `${key} must be defined!`;
    } else if (typeof variables[key] === "object" && variables !== null) {
      checkVariablesDefined(variables[key]); // Recursively check nested objects
    }
  }
}

/**
 * Checks if the specified properties of an object are defined.
 * 
 * This function takes an object and a list of property names to check.
 * It throws an error if any of the specified properties are `undefined`, with a message indicating which property must be defined.
 * 
 * @param {Object} obj The object containing properties to check.
 * @param {string[]} properties The list of property names to check.
 * @throws Throws an error if any property is `undefined`.
 * 
 * Example:
 * ```javascript
 * checkObjectPropertiesDefined(targetSheet, ["name", "sheet", "idColumn", "companyColumn", "yearColumn", "startingRow"]);
 * ```
 * @see `checkVariablesDefined()`
 */
function checkObjectPropertiesDefined(obj, properties) {
  properties.forEach(property => {
    if (obj[property] === undefined) {
      throw `${property} must be defined!`;
    }
  });
}

/**
 * Checks if the given variables are truthy.
 * 
 * This function takes an object where the keys are variable names and the values are the variables themselves.
 * It throws an error if any of the variables are `undefined`, `null`, `""`, `NaN`, or `[]`, with a message indicating which variable must be defined.
 * 
 * @param {Object} variables An object containing variables to check.
 * @throws Throws an error if any variable or specified property is falsy.
 * 
 * Example:
 * ```javascript
 * checkVariablesTruthy({ sourceSheet, targetSheet, startingRow, startingColumn });
 * ```
 * @see `checkObjectPropertiesTruthy()`
 */
function checkVariablesTruthy(variables) {
  for (const key in variables) {
    if (variables[key] === undefined || variables[key] === null || variables === "" || Number.isNaN(variables[key]) || (Array.isArray(variables[key]) && variables[key].length === 0)) {
      throw `${key} must be defined! Recieved: ${variables[key]}.`;
    } else if (typeof variables[key] === "object" && variables !== null) {
      checkVariablesTruthy(variables[key]); // Recursively check nested objects
    }
  }
}

/**
 * Checks if the specified properties of an object are truthy.
 * 
 * This function takes an object and a list of property names to check.
 * It throws an error if any of the specified properties are `undefined`, `null`, `""`, `NaN`, or `[]`, with a message indicating which property must be defined.
 * 
 * @param {Object} obj The object containing properties to check.
 * @param {string[]} properties The list of property names to check.
 * @throws Throws an error if any property is falsy.
 * 
 * Example:
 * ```javascript
 * checkObjectPropertiesTruthy(targetSheet, ["name", "sheet", "idColumn", "companyColumn", "yearColumn", "startingRow"]);
 * ```
 * @see `checkVariablesTruthy()`
 */
function checkObjectPropertiesTruthy(obj, properties) {
  properties.forEach(property => {
    if (obj[property] === undefined || obj[property] === null || obj[property] === "" || Number.isNaN(obj[property]) || (Array.isArray(obj[property]) && obj[property].length === 0)) {
      throw `${property} must be defined! Recieved: ${obj[property]}.`;
    }
  });
}

/**
 * Checks if any given variable is less than or equal to 0.
 * 
 * This function takes an object where the keys are variable names and the values are the variables themselves.
 * It throws an error if any of the variables are less than or equal to 0, with a message indicating which variable is less than or equal to zero (along with its value).
 * @param {Object} variables An object containing variables to check.
 * @throws Throws an error if any variables specified are less than or equal to 0.
 * 
 * Example:
 * ```javascript
 * checkVariablesGreaterThanZero({ startingRow, startingColumn });
 * ```
 */
function checkVariablesGreaterThanZero(variables) {
  for (const key in variables) {
    if (variables[key] === undefined) {
      throw `${key} cannot be undefined!`;
    } else if (Number.isNaN(variables[key]) || variables[key] < 1) {
      throw `${key} must be greater than or equal to 1! Recieved: ${variables[key]}.`;
    }
  }
}

/**
 * Checks if the given array `arr` contains all the elements in `elms`.
 * 
 * This function takes in an array of objects and checks if it contains every element in `elms`.
 * @param {Object[]} arr The array to check for elements in.
 * @param {...Object} elms The elements to check for in `arr`.
 * @throws Throws an error if any of the elements in `elms` is not within `arr`.
 * 
 * Example:
 * ```javascript
 * checkArrayContainsElements(includeColumns, "Company");
 * checkArrayContainsElements(includeColumns, "Company", "Year");
 * checkArrayContainsElements(includeColumns, ["Company", "Year"]);
 * ```
 */
function checkArrayContainsElements(arr, ...elms) {
  checkVariablesDefined(arr);
  if (!Array.isArray(arr)) {
    throw new TypeError("Parameter 'arr' must be an array! Instead recieved " + typeof arr + ".");
  }
  if (elms === undefined || elms.length === 0) return;

  elms = elms.flat();
  // Use set instead of looping through all elements to allow for filter and outputting ALL the missing elements
  const arrSet = new Set(arr);

  /**
   * All of the elements missing from arr
   * @type {Object[]}
   */
  const missingElements = elms.filter(elm => !arrSet.has(elm));
  if (missingElements.length > 0) {
    throw new Error(`Array must include required elements: ${missingElements.join(', ')}!`);
  }
}