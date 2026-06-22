/**
 * This script file handles all aspects of importing bank transactions.
 * It's designed with two workflows: a simple one-click import for normal months,
 * and a two-step process for reconciliation months that provides more control.
 */

/**
 * HELPER FUNCTION: Cleans the raw data on the "BankImport" sheet.
 * It takes a raw CSV paste, standardizes it to three columns, and writes it back.
 * @return {Array<Array<String>>|null} A 2D array of the cleaned data (if successful), or null on failure.
 */
function cleanBankImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName(SHEET_NAMES.BANK_IMPORT);
  const ui = SpreadsheetApp.getUi();

  // Validate that there is data to clean.
  if (!importSheet || importSheet.getLastRow() < 2) {
    ui.alert('Nothing to clean. Paste raw CSV data into the "BankImport" sheet first.');
    return null;
  }
  Logger.log("Cleaning bank import sheet...");
  const data = importSheet.getDataRange().getValues();
  const RAW_DESC_INDEX = 4; // Assumes Description is in Column 5 of the raw CSV.

  // Loop through the raw data and map it to a clean, 3-column format.
  const cleaned = data.map((row, index) => {
    if (index === 0) return ["Date", "Amount", "Description"]; // Set standard headers.
    if (row.length > RAW_DESC_INDEX) {
      return [row[0], row[1], row[RAW_DESC_INDEX]]; // Extract the three key columns.
    }
    return [null, null, null]; // Return a blank row if the source row is too short.
  });

  // Clear the import sheet and write the newly cleaned data back.
  importSheet.getRange("A2:Z").clearContent();
  const cleanedDataRows = cleaned.slice(1).filter(r => r[0]); // Filter out any empty rows.
  if (cleanedDataRows.length > 0) {
    importSheet.getRange(2, 1, cleanedDataRows.length, 3).setValues(cleanedDataRows);
    Logger.log("BankImport sheet has been cleaned successfully.");
    // Return the cleaned data array so it can be passed directly to other functions.
    return cleanedDataRows;
  }
  return null; // Return null if no valid data was found after cleaning.
}

/**
 * HELPER FUNCTION: Reads the cleaned "BankImport" sheet and auto-populates the Vendors sheet.
 * This function's only purpose is to find new, unknown vendors for categorization.
 */
function populateVendorsFromCleanImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName(SHEET_NAMES.BANK_IMPORT);
  if (!importSheet || importSheet.getLastRow() < 2) return;

  Logger.log("Populating vendor list...");
  // Read only the descriptions from Column C.
  const descriptions = importSheet.getRange(2, 3, importSheet.getLastRow() - 1, 1).getValues();
  // Call getRenamedVendor for each description. We don't use the return value; we are just
  // using the function for its side-effect of adding unknown vendors to the "Vendors" sheet.
  descriptions.forEach(row => {
    if (row[0]) getRenamedVendor(row[0]);
  });
  Logger.log("Vendor list population complete.");
}


// --- USER-FACING FUNCTIONS (CALLED BY THE MENU) ---

/**
 * WORKFLOW 1: The standard, all-in-one import process for normal months.
 * It cleans the sheet, then immediately injects the data into the monthly budget.
 */
function importBankTransactions() {
  // First, clean the sheet and get the cleaned data back.
  const cleanedData = cleanBankImportSheet();
  // The 'if (cleanedData)' check works because a non-empty array is "truthy" and null is "falsy".
  if (cleanedData) {
    // If cleaning was successful, pass the cleaned data directly to the injection function.
    cleanAndInjectBankCSV(cleanedData);
  }
}

/**
 * WORKFLOW 2: The preparation step for reconciliation months.
 * This cleans the sheet, populates the vendor list, and then stops, allowing the user to categorize.
 */
function prepareForReconciliation() {
  if (cleanBankImportSheet()) {
    populateVendorsFromCleanImport();
    SpreadsheetApp.getUi().alert('✅ Preparation complete. Please go to your "Vendors" sheet to fill in renames before injecting.');
  }
}


// --- MAIN PROCESSING FUNCTION (REFACTORED) ---
/**
 * This is the core data processing engine, refactored for high efficiency.
 * It now accepts the cleaned data as an argument, making it more focused and reusable.
 * It uses a multi-pass, cache-based approach to minimize slow spreadsheet operations.
 * It injects transactions, de-duplicates them against manual entries, and moves
 * any unmatched manual entries.
 * @param {Array<Array<String>>} importRows The cleaned transaction data to be injected.
 */
function cleanAndInjectBankCSV(importRows) {
  // --- INITIAL SETUP ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const todayCol = new Date().getDate();
  const numTransactionRows = LAYOUT.TRANSACTION_END_ROW - LAYOUT.TRANSACTION_START_ROW + 1;

  Logger.log("Starting injection process...");

  // Validate the data passed into the function.
  if (!importRows || importRows.length === 0) {
    Logger.log("No valid data rows received for injection.");
    return;
  }

  // Because the data is already cleaned, we know the column indexes are fixed.
  const dateColIdx = 0;
  const amountColIdx = 1;
  const descColIdx = 2;

  // --- STEP 1: FIRST PASS - IDENTIFY NEEDED DATA ---
  // Loop through the import data to see which sheets/columns we need to read from.
  const sheetDataCache = {};
  const sheetsToAccess = {};
  const touchedTabs = new Set();

  importRows.forEach(row => {
    try {
      const date = new Date(row[dateColIdx]);
      const sheetName = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMMM YYYY");
      const col = date.getDate();
      touchedTabs.add(sheetName);
      if (!sheetsToAccess[sheetName]) {
        sheetsToAccess[sheetName] = new Set();
      }
      sheetsToAccess[sheetName].add(col);
      sheetsToAccess[sheetName].add(todayCol);
    } catch (e) {
      Logger.log(`Error during sheet/col identification: ${e}`);
    }
  });

  // --- STEP 2: BATCH READ - PULL ALL DATA INTO CACHE ---
  // Now, read all required data from the identified sheets/columns into an in-memory cache.
  for (const sheetName in sheetsToAccess) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    if (!sheetDataCache[sheetName]) sheetDataCache[sheetName] = {};
    const colsToRead = Array.from(sheetsToAccess[sheetName]);

    colsToRead.forEach(col => {
      try {
        if (!col || col < 1 || col > sheet.getMaxColumns()) return;
        const range = sheet.getRange(LAYOUT.TRANSACTION_START_ROW, col, numTransactionRows, 1);
        // Store all cell properties in our cache object for fast access.
        sheetDataCache[sheetName][col] = {
          values: range.getValues(),
          notes: range.getNotes(),
          backgrounds: range.getBackgrounds(),
          formats: range.getNumberFormats(),
          fontStyles: range.getFontStyles(),
          fontWeights: range.getFontWeights(),
          fontColors: range.getFontColors()
        };
      } catch (e) {
        Logger.log(`Error reading data for ${sheetName}, column ${col}: ${e}`);
      }
    });
  }

  // --- STEP 3: PROCESS IN MEMORY - MAKE DECISIONS USING CACHE ---
  // Loop through import data again, making all decisions using the fast in-memory cache.
  const changesToApply = {
    updates: [],
    clears: []
  };
  let transactionsProcessed = 0,
    transactionsPlaced = 0,
    transactionsMatched = 0;

  importRows.forEach(row => {
    let sheetName, col, amount, renamedVendor, note, date;
    try {
      date = new Date(row[dateColIdx]);
      sheetName = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMMM YYYY");
      col = date.getDate();
      amount = parseFloat(row[amountColIdx]);
      if (isNaN(amount)) return;
      renamedVendor = getRenamedVendor(row[descColIdx]);
      note = `Imported: ${renamedVendor} | Original: ${row[descColIdx]}`;
      transactionsProcessed++;
    } catch (e) {
      Logger.log(`Error preparing import row for processing: ${e}.`);
      return;
    }

    const dayData = sheetDataCache[sheetName] ?.[col];
    if (!dayData) return;

    let placed = false;

    // A) De-duplication Logic: Check for a match against manual entries.
    for (let i = 0; i < numTransactionRows; i++) {
      const cellValue = dayData.values[i][0];
      const cellNote = dayData.notes[i][0];
      if (typeof cellValue === 'number' && cellValue === amount && typeof cellNote === 'string' && cellNote.trim().toLowerCase() === renamedVendor.toLowerCase()) {
        const targetRow = i + LAYOUT.TRANSACTION_START_ROW;
        changesToApply.updates.push({
          sheetName, row: targetRow, col, note: note, background: "#d9d9d9"
        });
        dayData.notes[i][0] = note;
        dayData.backgrounds[i][0] = "#d9d9d9"; // Update cache immediately.
        transactionsMatched++;
        break;
      }
    }

    // B) Placement Logic: If no match, find the first empty row.
    if (!placed) {
      for (let i = 0; i < numTransactionRows; i++) {
        if (dayData.values[i][0] === '' || dayData.values[i][0] === null) {
          const targetRow = i + LAYOUT.TRANSACTION_START_ROW;
          changesToApply.updates.push({
            sheetName, row: targetRow, col, value: amount, note: note, background: "#d9d9d9", format: '$#,##0.00'
          });
          dayData.values[i][0] = amount; // Update cache immediately.
          dayData.notes[i][0] = note;
          dayData.backgrounds[i][0] = "#d9d9d9";
          placed = true;
          transactionsPlaced++;
          break;
        }
      }
    }
  });

  // --- STEP 3b: PROCESS MOVES IN MEMORY ---
  // Find any manual entries in touched columns that were NOT matched and move them.
  let entriesMoved = 0;
  touchedTabs.forEach(sheetName => {
    const processedCols = Array.from(sheetsToAccess[sheetName] || []).filter(c => c !== todayCol);
    const destDayData = sheetDataCache[sheetName] ?.[todayCol];
    if (!destDayData) return;

    processedCols.forEach(col => {
      const sourceDayData = sheetDataCache[sheetName] ?.[col];
      if (!sourceDayData) return;
      for (let i = 0; i < numTransactionRows; i++) {
        // Find a manual entry (value exists and background is not gray).
        if ((sourceDayData.values[i][0] || sourceDayData.values[i][0] === 0) && sourceDayData.backgrounds[i][0] !== "#d9d9d9") {
          // Find the first empty slot in today's column to move it to.
          for (let iMove = 0; iMove < numTransactionRows; iMove++) {
            if (destDayData.values[iMove][0] === '' || destDayData.values[iMove][0] === null) {
              const sourceRow = i + LAYOUT.TRANSACTION_START_ROW;
              const destRow = iMove + LAYOUT.TRANSACTION_START_ROW;
              // Record the clear and update operations.
              changesToApply.clears.push({
                sheetName, row: sourceRow, col
              });
              changesToApply.updates.push({
                sheetName, row: destRow, col: todayCol, value: sourceDayData.values[i][0],
                note: `Moved: ${sourceDayData.notes[i][0]}`,
                format: sourceDayData.formats[i][0], fontColor: sourceDayData.fontColors[i][0],
                fontStyle: sourceDayData.fontStyles[i][0], fontWeight: sourceDayData.fontWeights[i][0]
              });
              // Update the cache immediately to reflect the move.
              destDayData.values[iMove][0] = sourceDayData.values[i][0];
              sourceDayData.values[i][0] = '';
              entriesMoved++;
              break;
            }
          }
        }
      }
    });
  });

  // --- STEP 4: BATCH WRITE - APPLY ALL CHANGES TO THE SPREADSHEET ---
  // Group clear operations by sheet for efficiency.
  const rangesToClearBySheet = {};
  changesToApply.clears.forEach(c => {
    if (!rangesToClearBySheet[c.sheetName]) rangesToClearBySheet[c.sheetName] = [];
    rangesToClearBySheet[c.sheetName].push(`R${c.row}C${c.col}`);
  });
  for (const sheetName in rangesToClearBySheet) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.getRangeList(rangesToClearBySheet[sheetName]).clear({
        contentsOnly: false, formatOnly: true
      });
    }
  }

  // Apply all update operations.
  changesToApply.updates.forEach(u => {
    const sheet = ss.getSheetByName(u.sheetName);
    if (sheet) {
      try {
        const cell = sheet.getRange(u.row, u.col);
        if (u.value !== undefined) cell.setValue(u.value);
        if (u.note !== undefined) cell.setNote(u.note);
        if (u.background !== undefined) cell.setBackground(u.background);
        if (u.format !== undefined) cell.setNumberFormat(u.format);
        if (u.fontColor) cell.setFontColor(u.fontColor);
        if (u.fontStyle) cell.setFontStyle(u.fontStyle);
        if (u.fontWeight) cell.setFontWeight(u.fontWeight);
      } catch (e) {
        Logger.log(`Error applying update to ${u.sheetName}!R${u.row}C${u.col}: ${e}`);
      }
    }
  });

  // --- FINALIZATION ---
  SpreadsheetApp.flush();
  Logger.log("cleanAndInjectBankCSV completed.");
  ui.alert(`✅ Import Complete! Processed: ${transactionsProcessed}, Matched: ${transactionsMatched}, Placed: ${transactionsPlaced}, Moved: ${entriesMoved}.`);
}