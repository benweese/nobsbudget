/**
 * Takes a raw transaction description from a bank import and attempts to clean it up.
 * It uses a set of rules from the "Vendors" sheet to perform a partial, case-insensitive match.
 * If a matching rule is found, it returns the standardized vendor name.
 * If a rule is found but the standardized name is blank, it returns a flag.
 * If no rule is found, it adds the new, unknown description to the "Vendors" sheet for later review
 * and returns the original description.
 *
 * @param {string} description The raw transaction description from the bank.
 * @return {string} The standardized vendor name, a flag, or the original description.
 */
function getRenamedVendor(description) {
  // --- INITIAL SETUP ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Get the "Vendors" sheet using the constant from utils.gs to prevent typos.
  const vendorSheet = ss.getSheetByName(SHEET_NAMES.VENDORS);
  // If the sheet doesn't exist, stop the script and throw an error. This is a critical failure.
  if (!vendorSheet) {
    Logger.log(`Error: Missing sheet - ${SHEET_NAMES.VENDORS}`);
    throw new Error(`Missing '${SHEET_NAMES.VENDORS}' tab`);
  }

  // --- EFFICIENT DATA READ ---
  // Read all the vendor rules (Column A: Keyword, Column B: Rename) into a 2D array in memory.
  // This is highly efficient because we only read from the spreadsheet once, not inside the loop.
  // It starts from row 2 to skip the header.
  const vendorData = vendorSheet.getRange(2, 1, vendorSheet.getLastRow() - 1, 2).getValues();
  // Convert the incoming bank description to lowercase once to be used for all comparisons.
  const lowerDesc = description.toLowerCase();

  // --- MATCHING LOGIC ---
  // Loop through each rule (row) from the in-memory 'vendorData' array.
  for (let [matchText, renamed] of vendorData) {
    // Basic validation: ensure the keyword cell is not empty or just whitespace.
    if (matchText && typeof matchText === 'string' && matchText.trim() !== '') {
      // Get the keyword from the sheet, trim whitespace, and convert to lowercase for matching.
      const keyword = matchText.trim().toLowerCase();
      // Check if the raw bank description *includes* the keyword. This allows for partial matches.
      const doesInclude = lowerDesc.includes(keyword);

      // This is a special log for debugging a specific, problematic vendor name.
      if (lowerDesc.includes("children's hospita")) {
        Logger.log(`DEBUG Vendor: Checking if "${lowerDesc}" includes "${keyword}": ${doesInclude}`);
      }

      // If a match is found...
      if (doesInclude) {
        // ...log the success for debugging purposes.
        Logger.log(`DEBUG Vendor: MATCH FOUND! Desc: "${lowerDesc}" | Keyword: "${keyword}" | Rename: "${renamed}"`);
        // Return the result. If the 'renamed' cell (Column B) has text, return it.
        // If the 'renamed' cell is empty, return a flag to signal that this vendor needs a name.
        return renamed ? String(renamed) : "Flag: Please name Vendor";
      }
    }
  }
  // If the loop completes without finding any matches...
  Logger.log(`DEBUG Vendor: No match found for "${description}". Adding to sheet and returning original.`);


  // --- HANDLING UNKNOWN VENDORS ---
  // If no match was found, add the new, unknown description to the "Vendors" sheet.
  // This allows you to categorize it later.
  if (description && typeof description === 'string' && description.trim() !== '') {
    // Get the next available empty row and write the description into Column A.
    const lastRow = vendorSheet.getLastRow() + 1;
    vendorSheet.getRange(lastRow, 1).setValue(description);
    Logger.log(`Added new vendor description to ${SHEET_NAMES.VENDORS}: ${description}`);
  } else {
    // Log if the description was empty, to avoid adding blank rows to the Vendors sheet.
    Logger.log(`Skipped adding empty or invalid description to ${SHEET_NAMES.VENDORS}.`);
  }

  // Since no rule was matched, return the original, unchanged description.
  return description;
}