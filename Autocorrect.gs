// ######
// This file is mostly depricated, but is still technically used. It is for checking inputted company names and comparing them to already listed ones.
// This has changed since now it is very unlikely for there be a typo due to how a user can click on a dropdown (courtesy of Form Updater).
// It would not surprise me in the slightest if this were to cause issues with it having a false positive with a company name. That is why the Cleaned Company field is in Raw Data.
// ######
/**
 * Computes the Jaro-Winkler distance between two strings
 * @param {string} s1 The first string
 * @param {string} s2 The second string
 * @returns {number} The Jaro-Winkler distance. 0 if `s1` and `s2` do not match at all and 1 if they are an exact match
 * @see https://en.wikipedia.org/wiki/Jaro-Winkler_distance
 */
function jaroWinkler(s1, s2) {
  const s1_len = s1.length;
  const s2_len = s2.length;

  if (s1_len === 0 && s2_len === 0) return 1.0;
  if (s1_len === 0 || s2_len === 0) return 0.0;

  const matchWindow = Math.floor(Math.max(s1_len, s2_len) / 2) - 1;

  let matches = 0;
  let transpositions = 0;
  const s1_matches = Array(s1_len).fill(false);
  const s2_matches = Array(s2_len).fill(false);

  for (let i = 0; i < s1_len; i++) {
    const start = Math.max(0, i - matchWindow);
    const end = Math.min(i + matchWindow + 1, s2_len);

    for (let j = start; j < end; j++) {
      if (s2_matches[j]) continue;
      if (s1[i] !== s2[j]) continue;
      s1_matches[i] = true;
      s2_matches[j] = true;
      matches++;
      break;
    }
  }

  if (matches === 0) return 0.0;

  let k = 0;
  for (let i = 0; i < s1_len; i++) {
    if (!s1_matches[i]) continue;
    while (!s2_matches[k]) k++;
    if (s1[i] !== s2[k]) transpositions++;
    k++;
  }

  const jaro = (matches / s1_len + matches / s2_len + (matches - transpositions / 2) / matches) / 3.0;
  let prefix = 0;
  const prefixScale = 0.1;

  for (let i = 0; i < Math.min(4, s1_len, s2_len); i++) {
    if (s1[i] === s2[i]) {
      prefix++;
    } else {
      break;
    }
  }

  return jaro + prefix * prefixScale * (1 - jaro);
}

/**
 * Calculate the threshold based on the string's length.
 * @param {string} str A string to calculate the threshold for
 * @returns {number} The calculated threshold
 */
function calculateThreshold(str) {
  const length = str.length;
  if (length <= 5) return 0.8;
  if (length <= 10) return 0.7;
  if (length <= 15) return 0.6;
  return 0.5;
}


/**
 * @param {string} company The company to find the closest one to
 * @param {string[]} companies A list of companies to use as the autocorrect list
 * @param {number} threshold The threshold tolerance to set the autocorrect to.
 * @returns {string} The closest company to company within companies, or `Not Found` if there is no match
 */
function findClosestCompany(company, companies, threshold) {
  if (threshold < 0) {
    throw "Threshold must be greater than 0!";
  }

  let maxDistance = 0;
  let closestCompany = "Not Found";

  companies.forEach(validCompany => {
    let distance = jaroWinkler(company.toLowerCase(), validCompany.toLowerCase());
    if (distance > maxDistance) {
      maxDistance = distance;
      closestCompany = validCompany;
    }
  });

  return maxDistance >= threshold ? closestCompany : "Not Found";
}



/**
 * @param {string} company The company to find the closest one to
 * @param {string[]} companies A list of companies to use as the autocorrect list
 * @param {number} threshold The threshold tolerance to set the autocorrect to
 * @returns {string} The autocorrected company, `""` if the company is empty, or `Not Found` if there is not a match
 */
function autocorrectCompanyWithThreshold(company, companies, threshold) {
  if (!company) {
    return "";
  }
  if (threshold < 0) {
    throw "Threshold should be greater than 0!";
  }

  let closestCompany = findClosestCompany(company, companies, threshold);
  return closestCompany;
}

/**
 * Computes the autocorrect with the threshold being determined based on the string's length
 * @param {string} company The company to find the closest one to
 * @param {string[]} companies A list of companies to use as the autocorrect list
 * @returns {string} The autocorrected company, `""` if the company is empty, or `Not Found` if there is not a match
 * @see autocorrectCompany(company, companies, threshold)
 */
function autocorrectCompany(company, companies) {
  return autocorrectCompanyWithThreshold(company, companies, calculateThreshold(company));
}


/**
 * Cleans up the company name. Cleaning entails:
 * - Removing "Inc." and all its variants
 * - Any tabs/line breaks replaced with spaces
 * - Replace all spaces greater than 1 with a single space
 * - Removing leading and trailing whitespace
 * @param {string} company The company to clean
 * @returns {string} A String of the cleaned company name
 */
function cleanCompanyName(company) {
  let cleanedCompany = company;

  // Remove "Inc." and all its varients
  const incRegex = /\s*inc\.?$/i;
  cleanedCompany = cleanedCompany.replace(incRegex, "");

  // Any tabs/line breaks replaced with spaces
  cleanedCompany = cleanedCompany.replace("\t", " ");
  cleanedCompany = cleanedCompany.replace("\n", " ");
  cleanedCompany = cleanedCompany.replace("\r", " ");

  // Replace all spaces greater than 1 with a single space
  const multiSpaceRegex = /\s{2,}/;
  cleanedCompany = cleanedCompany.replace(multiSpaceRegex, " ");

  // Remove leading and trailing whitespace
  cleanedCompany = cleanedCompany.trim();

  return cleanedCompany;
}