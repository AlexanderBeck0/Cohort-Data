// #####
// Used for updating the input form
// #####
/**
 * @returns {FormApp.Form} The Form used to input data
 */
function getForm() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const formURL = spreadsheet.getFormUrl();
  const form = FormApp.openByUrl(formURL);
  return form;
}

/**
 * On a trigger to run every month on the first. Will update the Company List on the form to be the companies marked Active on `Company List`.
 */
function updateCompanyListOnForm() {
  const form = getForm();
  const items = form.getItems();
  const companyListIndex = items.findIndex(item => item.getTitle() === "Company Name");
  const companyListItem = items[companyListIndex].asListItem();
  const activeCompanies = getActiveCompanies().sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  const pageBreakItems = form.getItems(FormApp.ItemType.PAGE_BREAK).map(pageBreak => pageBreak.asPageBreakItem());

  const activeCompanysAsChoices = activeCompanies.map(company => companyListItem.createChoice(company, pageBreakItems[1]));
  const newCompanyChoice = companyListItem.createChoice("Not listed/New Company", pageBreakItems[0]);
  companyListItem.setChoices([newCompanyChoice].concat(activeCompanysAsChoices));
}

/**
 * @param {string} company The company to add to the choice list
 */
function addCompanyToListOnForm(company) {
  const form = getForm();
  const items = form.getItems();
  const companyListIndex = items.findIndex(item => item.getTitle() === "Company Name");
  const companyListItem = items[companyListIndex].asListItem();
  const activeCompanies = companyListItem.getChoices();
  // Don't add the company if it is already on the list
  if (activeCompanies.filter(activeCompany => activeCompany.getValue() === company).length > 0) return;

  const pageBreakItems = form.getItems(FormApp.ItemType.PAGE_BREAK).map(pageBreak => pageBreak.asPageBreakItem());
  const newCompany = companyListItem.createChoice(company, pageBreakItems[1]);
  const newCompanyChoice = activeCompanies.shift();
  activeCompanies.push(newCompany);
  const newActiveCompanies = activeCompanies.sort((a, b) => a.getValue().toLowerCase().localeCompare(b.getValue().toLowerCase()));
  companyListItem.setChoices([newCompanyChoice].concat(newActiveCompanies));
}

function updateYearAndCohortYearOnForm() {
  throw "Not implemented";
}

/**
 * @param {string} name The name of the person to delete all the responses from
 * @param {Date} [timestamp] The earliest date and time for which form responses should be deleted 
 */
function deleteAllResponsesByEnterName(name, timestamp) {
  if (!name || name.length === 0) {
    throw "name cannot be undefined!";
  }

  const form = getForm();
  const nameItem = form.getItems(FormApp.ItemType.TEXT).find(item => item.getTitle() === "Who's entering the data?");
  let responseIDs = [];
  if (timestamp === undefined) {
    // Get all responses from name
    responseIDs = form.getResponses().filter(response => response.getResponseForItem(nameItem).getResponse() === name);
  } else {
    // Get only responses after timestamp from name
    responseIDs = form.getResponses(timestamp).filter(response => response.getResponseForItem(nameItem).getResponse() === name);
  }

  responseIDs.forEach(response => form.deleteResponse(response.getId()));

}
