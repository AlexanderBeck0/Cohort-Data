// #####
// The entry point of the program. This is where new data comes in from the form. 
// This is probably the least readable portion of the code, so good luck. I tried my best to add comments, but the overall structure is difficult to understand without tracing through it.
// Future developer note: This is extremely fragile code. Especially the dataToAdd, since the values get moved around far too much.
// #####
/**
 * @type {RawDataColumns}
 * @satisfies {RawDataColumns}
 * @readonly
 * @global
 */
const rawDataColumns = Object.freeze({
  Company: 0,
  CohortYear: 1,
  Year: 2,
  Revenue: 3,
  Expenses: 4,
  Difference: 5,
  CompanyID: 6,
  CleanedCompany: 7,
  Funded: 8,
});

/**
 * @static
 * @const
 * @type {number}
 */
const NUMBER_OF_YEARS = 5;

/**
 * The event handler triggered when a form is submitted. Also the entry point of the program
 * @param {Event} e The onFormSubmit event
 */
function onFormSubmit(e) {
  /**
   * @readonly
   */
  const responseColumns = Object.freeze({
    Timestamp: 0,
    WhosEntering: 1,
    CompanyName: 2,
    CohortYear: 3,
    Year: 4,
    Revenue: 5,
    Expenses: 6,
    NewCompanyName: 7,
    HasBeenFunded: 8,
    FundingAmount: 9,
  });

  /**
   * The responses from the form
   * @type {Array}
   */
  const responses = e.values;

  Logger.log(responses);

  // Clip to CompanyName and HasBeenFunded
  // This is because it is zero indexed, not because it knows where CompanyName is. It is based entirely on the fact that responseColumns > 0.
  const dataToCopy = responses.slice(responseColumns.CompanyName, responseColumns.FundingAmount + 1);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = spreadsheet.getSheetByName("Raw Data");
  const companyListSheet = spreadsheet.getSheetByName("Company List");

  // A new company was added
  const isNewCompany = responses[responseColumns.NewCompanyName] !== "";
  if (isNewCompany) {
    // Last item of responses is New Company Name
    dataToCopy[rawDataColumns.Company] = responses[responseColumns.NewCompanyName];
    // Clean the company name
    dataToCopy[rawDataColumns.CleanedCompany] = cleanCompanyName(dataToCopy[rawDataColumns.Company]);

    // Autocorrect companies
    doAutocorrect(companyListSheet, dataToCopy, rawDataColumns, isNewCompany); // TODO: Might want to adjust tolerances to be very high to reduce likelyhood of a false positive
    addCompanyToListOnForm(dataToCopy[rawDataColumns.CleanedCompany]);
  } else {
    // No need to run this clean twice
    // Clean the company name
    dataToCopy[rawDataColumns.CleanedCompany] = cleanCompanyName(dataToCopy[rawDataColumns.Company]);
  }

  // Check that year is valid
  const isValidYear = +dataToCopy[rawDataColumns.Year] >= +dataToCopy[rawDataColumns.CohortYear] - 1 && +dataToCopy[rawDataColumns.Year] <= +dataToCopy[rawDataColumns.CohortYear] + (NUMBER_OF_YEARS - 1);
  // If year is invalid, do not continue processing the data
  if (!isValidYear) {
    Logger.log("Invalid year given! Not adding data to Raw Data. Year: " + dataToCopy[rawDataColumns.Year] + ", Cohort Year: " + dataToCopy[rawDataColumns.CohortYear]);
    return;
  }

  let indexOfCompany = rowIndexOfValue(dataToCopy[rawDataColumns.CleanedCompany], 1, companyListSheet, 1);

  // Calculate Difference field
  dataToCopy[rawDataColumns.Difference] = calculateDifference(dataToCopy[rawDataColumns.Revenue], dataToCopy[rawDataColumns.Expenses], "");

  if (indexOfCompany !== -1) {
    // 4 is the Funded column
    const funded = companyListSheet.getRange(indexOfCompany, 4).getValue();
    if (dataToCopy[rawDataColumns.Funded] === "Yes" && !funded) {
      setAllFunded(rawDataSheet, dataToCopy[rawDataColumns.CleanedCompany], true);
    } else {
      dataToCopy[rawDataColumns.Funded] = String(funded);
    }
  } else {
    dataToCopy[rawDataColumns.Funded] = String(responses[responseColumns.HasBeenFunded] === "Yes");
  }

  // Add new data or update existing data
  // Existing meaning it has the same CompanyID and Year
  updateOrAddData(rawDataSheet, dataToCopy);

  // Add or Update the cohort years within Company List
  // let indexOfCompany = rowIndexOfValue(dataToCopy[rawDataColumns.CleanedCompany], rawDataColumns.CleanedCompany, companyListSheet, 1);
  if (indexOfCompany !== -1) {
    Logger.log("Company exists! Updating Company List...");
    updateOrAddCohortYearList(companyListSheet, dataToCopy[rawDataColumns.CohortYear], indexOfCompany + 1, 3, ", ");
  } else {
    // addToCompanyList returns the index that it was added to
    // I was about to go make this change, but I realized I had already done this ahead of time
    // Yay me
    Logger.log("New Company being added to Company List! New Company: " + dataToCopy[rawDataColumns.CleanedCompany] + ".");
    indexOfCompany = addToCompanyList(companyListSheet, dataToCopy[rawDataColumns.CleanedCompany], dataToCopy[rawDataColumns.CohortYear]);
  }

  // The company has been marked as funded
  if (responses[responseColumns.HasBeenFunded] !== "") {
    // Sends the index if it was found, or sends undefined if it was not
    setCompanyAsFunded(companyListSheet, dataToCopy[rawDataColumns.CleanedCompany], responses[responseColumns.FundingAmount], (indexOfCompany !== -1 ? indexOfCompany : undefined));
  }

  // Update other sheet fields
  updateYearValues();
}

/**
 * Must use `cleanCompanyName()` before running. Will not run if `CleanedCompany` in `dataToCopy` is empty.
 * NOTE: Modifies `dataToCopy`. Changes `CompanyID` and changes `Company`
 * @param {SpreadsheetApp.Sheet} companyListSheet The sheet containing a list of companies
 * @param {Array} dataToCopy The new data to be added. Is mutated
 * @param {RawDataColumns} columnIndices The indices of the raw data sheet and dataToCopy
 * @param {number} columnIndices.Company `dataToCopy[columnIndices.Company]` is mutated. The non-negative index of the Company column in rawDataSheet
 * @param {number} columnIndices.CohortYear The non-negative index of the Cohort Year column in Raw Data
 * @param {number} columnIndices.CompanyID `dataToCopy[columnIndices.CompanyID]` is mutated. The non-negative index of the Company ID column in rawDataSheet. CompanyID data in the form Company-$-CohortYear
 * @param {number} columnIndices.CleanedCompany The company name, cleaned
 * @param {boolean} isNewCompany A boolean representing if the user ticked off Yes to is a new company
 */
function doAutocorrect(companyListSheet, dataToCopy, columnIndices, isNewCompany) {
  if (!dataToCopy[columnIndices.CleanedCompany] || dataToCopy[columnIndices.CleanedCompany] === "") {
    throw "Company name must be cleaned before doing autocorrect!";
  }

  let indexOfCompany = rowIndexOfValue(dataToCopy[columnIndices.CleanedCompany], columnIndices.CleanedCompany, companyListSheet, 1);
  if (isNewCompany) {
    // Check if the company already exists in Company List
    const companyExists = indexOfCompany !== -1;
    if (companyExists) {
      Logger.log("Company " + dataToCopy[columnIndices.CleanedCompany] + " already exists! Canceling...");
      return;
    }

    // Add the new CompanyID and Company name (the cleaned one) to the dataToCopy
    dataToCopy[columnIndices.CompanyID] = `${dataToCopy[columnIndices.CleanedCompany]}-\$-${dataToCopy[columnIndices.CohortYear]}`;
    dataToCopy[columnIndices.Company] = dataToCopy[columnIndices.CleanedCompany];
    return;
  }

  // Not labeled as a new company
  // const companies = companyListSheet.getDataRange().getValues().flat().filter(String);
  const companies = companyListSheet.getRange(1, 1, companyListSheet.getLastRow()).getValues().flat().filter(String);
  const closestCompany = autocorrectCompany(dataToCopy[columnIndices.CleanedCompany], companies);

  if (closestCompany !== "Not Found") {
    // Found an autocorrect match
    // Just hope it's not something totally wrong...
    dataToCopy[columnIndices.CompanyID] = `${closestCompany}-\$-${dataToCopy[columnIndices.CohortYear]}`;
    dataToCopy[columnIndices.Company] = closestCompany;
  } else {
    // New data or something so horribly misspelled it can be considered new data
    Logger.log("Failed to find company name.");
    dataToCopy[columnIndices.CompanyID] = `${dataToCopy[columnIndices.CleanedCompany]}-\$-${dataToCopy[columnIndices.CohortYear]}`;
  }

}

/**
 * Updates the data inside `rawDataSheet` if within the sheet, there exists a row with the same CompanyID and Year. 
 * If there isn't a row where both match, add the value to the bottom of the sheet
 * @param {SpreadsheetApp.Sheet} rawDataSheet The sheet containing the raw data. The columns must line up with `columnIndices`
 * @param {Array} dataToCopy The new data to be added. Is mutated
 * @param {RawDataColumns} rawDataColumns The indices of the raw data sheet and dataToCopy
 * @param {number} columnIndices.Company The non-negative index of the Company column in rawDataSheet
 * @param {number} columnIndices.Year The non-negative index of the Year column in rawDataSheet 
 * @param {number} columnIndices.CompanyID The non-negative index of the Company ID column in rawDataSheet. CompanyID data in the form Company-$-CohortYear
 * @returns {boolean} Whether the data already existed or not. `true` if existing data was edited, `false` if data was added to the bottom
 */
function updateOrAddData(rawDataSheet, dataToCopy) {
  if (!dataToCopy[rawDataColumns.CompanyID] || dataToCopy[rawDataColumns.CompanyID] === "") {
    // throw "Must have Company ID before adding or updating data!";
    dataToCopy[rawDataColumns.CompanyID] = `${dataToCopy[rawDataColumns.Company]}-\$-${dataToCopy[rawDataColumns.CohortYear]}`;
  }

  const rawData = rawDataSheet.getDataRange().getValues();
  const newCompanyID = dataToCopy[rawDataColumns.CompanyID];
  // Loop through all data (and stop if the data was edited, no need to waste resources after finding the first instance)
  const rawDataLength = rawData.length; // I googled it and apparently it is faster to precalculate the length before the loop
  for (let i = 0; i < rawDataLength; i++) {
    const companyID = rawData[i][rawDataColumns.CompanyID];
    const year = rawData[i][rawDataColumns.Year];
    if (companyID === newCompanyID && +year === +dataToCopy[rawDataColumns.Year]) {
      // Already exists, copy it to another sheet for record purposes and replace it in rawDataSheet
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const editedDataSheet = spreadsheet.getSheetByName("Edited Data");
      copyRowsToSheet(rawDataSheet, i + 1, 1, editedDataSheet);

      // i + 1 because it gets the wrong row otherwise
      rawDataSheet.getRange(i + 1, 1, 1, dataToCopy.length).setValues([dataToCopy]);
      return true;
    }
  }

  // Data did not already exist, add new data
  const firstEmptyRow = rawDataSheet.getLastRow() + 1;
  rawDataSheet.getRange(firstEmptyRow, 1, 1, dataToCopy.length).setValues([dataToCopy]);
  return false;
}

/**
 * Sets all the values in `rawDataSheet` in `rawDataColumns.Funded` to `isFunded`
 * @param {SpreadsheetApp.Sheet} rawDataSheet The Raw Data sheet
 * @param {string} company The company to change the `Funded` column in
 * @param {boolean} isFunded A boolean representing if the company should be funded or not
 */
function setAllFunded(rawDataSheet, company, isFunded) {
  const companyRange = rawDataSheet.getRange(String.fromCharCode(rawDataColumns.Company + 65) + "2:" + rawDataSheet.getLastRow());
  const fundedRange = rawDataSheet.getRange(String.fromCharCode(rawDataColumns.Funded + 65) + "2:" + rawDataSheet.getLastRow());

  const companyValues = companyRange.getValues();
  const fundedValues = fundedRange.getValues();

  companyValues.forEach((companyRow, index) => {
    if (companyRow[0] === company) {
      fundedValues[index][0] = String(isFunded);
    }
  });

  fundedRange.setValues(fundedValues);
}