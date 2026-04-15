/**
 * Creates and populates the budget sheet for the next calendar month.
 * This function handles headers, recurring items, carry-over balance,
 * and summary formulas, all in one go.
 */
function populateNextMonthBudget() {
  // --- INITIAL SETUP ---
  // Get the main spreadsheet file, the currently active sheet, and its name.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  const currentName = currentSheet.getName();
  // Get the UI environment to show alerts and messages to the user.
  const ui = SpreadsheetApp.getUi();

  Logger.log("Starting populateNextMonthBudget...");

  // --- CALCULATE NEXT MONTH'S NAME AND YEAR ---
  // Use a regular expression to parse the current sheet name (e.g., "September 2025").
  const nameParts = currentName.match(/^(\w+)\s+(\d{4})$/);
  if (!nameParts) {
    ui.alert(`Current sheet name "${currentName}" is not in the expected "Month YYYY" format. Aborting.`);
    return;
  }
  const [, monthName, yearStr] = nameParts;
  const months = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const monthIndex = months.indexOf(monthName);
  if (monthIndex === -1) {
    ui.alert(`Invalid month name "${monthName}" in current sheet name. Aborting.`);
    return;
  }
  const year = parseInt(yearStr);
  const nextMonthIndex = (monthIndex + 1) % 12;
  const nextYear = monthIndex === 11 ? year + 1 : year;
  const nextMonthName = months[nextMonthIndex];
  const newSheetName = `${nextMonthName} ${nextYear}`;
  Logger.log(`Target new sheet: ${newSheetName}`);

  // --- VALIDATE AND CREATE NEW SHEET ---
  if (ss.getSheetByName(newSheetName)) {
    ui.alert(`Sheet "${newSheetName}" already exists. Aborting to avoid overwriting.`);
    return;
  }
  const newSheet = ss.insertSheet(newSheetName);
  ss.setActiveSheet(newSheet);
  const daysInMonth = new Date(nextYear, nextMonthIndex + 1, 0).getDate();
  Logger.log(`Created sheet "${newSheetName}" with ${daysInMonth} days.`);

  // --- STRUCTURE THE NEW SHEET ---
  const headerValues = [];
  for (let d = 1; d <= daysInMonth; d++) headerValues.push(d);
  const headerRange = newSheet.getRange(LAYOUT.HEADER_ROW, 1, 1, daysInMonth);
  headerRange.setValues([headerValues]);
  headerRange.setBackground('#e0e0e0').setFontWeight('bold').setFontColor('#356854');

  const targetRowCount = LAYOUT.TOTAL_ROWS;
  const currentRows = newSheet.getMaxRows();
  if (currentRows < targetRowCount) {
    newSheet.insertRowsAfter(currentRows, targetRowCount - currentRows);
  } else if (currentRows > targetRowCount) {
    newSheet.deleteRows(targetRowCount + 1, currentRows - targetRowCount);
  }
  const requiredCols = daysInMonth;
  const currentCols = newSheet.getMaxColumns();
  if (currentCols < requiredCols) {
    newSheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
  Logger.log(`Sheet structure set: ${targetRowCount} rows, ${daysInMonth} day columns.`);

  // --- CARRY OVER PREVIOUS MONTH'S FINAL BALANCE ---
  let prevBalance = 0;
  try {
    const lastRow = currentSheet.getLastRow();
    const lastCol = currentSheet.getLastColumn();
    if (lastRow > 0 && lastCol > 0) {
      prevBalance = currentSheet.getRange(lastRow, lastCol).getValue();
      if (isNaN(parseFloat(prevBalance))) {
        prevBalance = 0;
      }
    }
  } catch (e) {
    Logger.log(`Error reading previous balance: ${e}. Setting carry-over to 0.`);
    prevBalance = 0;
  }
  const carryCell = newSheet.getRange(LAYOUT.CARRY_OVER_ROW, 1);
  carryCell.setValue(prevBalance).setNumberFormat('$#,##0.00').setFontStyle('italic').setFontColor('#888888');
  Logger.log(`Carry-over balance set to: ${prevBalance}`);

  // --- EFFICIENTLY POPULATE RECURRING ITEMS ---
  const numTransactionRows = LAYOUT.TRANSACTION_END_ROW - LAYOUT.TRANSACTION_START_ROW + 1;
  const transactionRange = newSheet.getRange(LAYOUT.TRANSACTION_START_ROW, 1, numTransactionRows, daysInMonth);
  let sheetValues = transactionRange.getValues();

  const cellsToUpdate = [];
  const recurringSheet = ss.getSheetByName(SHEET_NAMES.RECURRING);

  if (!recurringSheet) {
    ui.alert(`Error: Recurring items sheet "${SHEET_NAMES.RECURRING}" not found.`);
  } else {
    const recurringData = recurringSheet.getRange(2, 1, recurringSheet.getLastRow() - 1, 5).getValues();

    recurringData.forEach(([dayOfMonthRaw, name, amount, frequency, startDateRaw]) => {
      if (!name || !amount || !frequency) return;
      const numericAmount = parseFloat(amount);
      if (isNaN(numericAmount)) return;

      // --- CENTRALIZED FREQUENCY LOGIC ---
      // This block determines which days of the month the item should be placed on.
      const targetDays = [];
      for (let day = 1; day <= daysInMonth; day++) {
        const thisDate = new Date(nextYear, nextMonthIndex, day);

        // All the complex frequency logic has been moved to a single, central function
        // in utils.gs for maintainability (Don't Repeat Yourself principle).
        // We create a simple 'rule' object to pass all the necessary data to it.
        const rule = {
            dayOfMonthRaw: dayOfMonthRaw,
            frequency: frequency,
            startDate: startDateRaw instanceof Date ? startDateRaw : (startDateRaw ? new Date(startDateRaw) : null)
        };
        const isMatch = isRecurringDateMatch(thisDate, rule);

        if (isMatch) {
          targetDays.push(day);
        }
      } // End of day loop

      // Place the recurring item into the first empty slot for each target day.
      targetDays.forEach(targetDay => {
        const colIndex = targetDay - 1;
        for (let r = 0; r < sheetValues.length; r++) {
          if (sheetValues[r][colIndex] === '' || sheetValues[r][colIndex] === null) {
            cellsToUpdate.push({
              row: r + LAYOUT.TRANSACTION_START_ROW,
              col: targetDay,
              value: numericAmount,
              note: String(name),
              format: '$#,##0.00',
              fontColor: numericAmount >= 0 ? '#3366cc' : '#e67e22',
              fontStyle: 'italic'
            });
            sheetValues[r][colIndex] = numericAmount;
            break;
          }
        }
      });
    });
  }

  // --- Step 3: Batch Write - Apply all collected changes to the spreadsheet at once. ---
  if (cellsToUpdate.length > 0) {
    cellsToUpdate.forEach(change => {
      const cell = newSheet.getRange(change.row, change.col);
      cell.setValue(change.value);
      if (change.note) cell.setNote(change.note);
      if (change.format) cell.setNumberFormat(change.format);
      if (change.fontColor) cell.setFontColor(change.fontColor);
      if (change.fontStyle) cell.setFontStyle(change.fontStyle);
    });
  }

  // --- SET SUMMARY FORMULAS (ROWS 18-20) ---
  for (let col = 1; col <= daysInMonth; col++) {
    const colLetter = columnToLetter(col);
    const rangeA1Notation = `${colLetter}${LAYOUT.TRANSACTION_START_ROW}:${colLetter}${LAYOUT.TRANSACTION_END_ROW}`;

    // Row 19: Expenses (SUM of all negative numbers in the column)
    const expenseFormula = `=IFERROR(SUM(FILTER(${rangeA1Notation}, ${rangeA1Notation}<0)), 0)`;
    newSheet.getRange(LAYOUT.EXPENSE_TOTAL_ROW, col).setFormula(expenseFormula);

    // Row 18: Previous Balance + Income (SUM of all positive numbers)
    let incomeFormula;
    if (col === 1) {
      const day1IncomeRangeA1 = `${colLetter}${LAYOUT.TRANSACTION_START_ROW + 1}:${colLetter}${LAYOUT.TRANSACTION_END_ROW}`;
      incomeFormula = `=A${LAYOUT.CARRY_OVER_ROW} + IFERROR(SUM(FILTER(${day1IncomeRangeA1}, ${day1IncomeRangeA1}>0)), 0)`;
    } else {
      const prevBalanceCell = `INDIRECT(ADDRESS(${LAYOUT.EOD_BALANCE_ROW}, COLUMN()-1))`;
      incomeFormula = `=${prevBalanceCell} + IFERROR(SUM(FILTER(${rangeA1Notation}, ${rangeA1Notation}>0)), 0)`;
    }
    newSheet.getRange(LAYOUT.INCOME_TOTAL_ROW, col).setFormula(incomeFormula);

    // Row 20: End of Day Balance (Row 18 + Row 19)
    const balanceFormula = `=${colLetter}${LAYOUT.INCOME_TOTAL_ROW}+${colLetter}${LAYOUT.EXPENSE_TOTAL_ROW}`;
    newSheet.getRange(LAYOUT.EOD_BALANCE_ROW, col).setFormula(balanceFormula);
  }

  // --- APPLY FORMATTING TO SUMMARY ROWS AND CELLS ---
  const incomeTotalRange = newSheet.getRange(LAYOUT.INCOME_TOTAL_ROW, 1, 1, daysInMonth);
  const expenseTotalRange = newSheet.getRange(LAYOUT.EXPENSE_TOTAL_ROW, 1, 1, daysInMonth);
  const eodBalanceRange = newSheet.getRange(LAYOUT.EOD_BALANCE_ROW, 1, 1, daysInMonth);
  incomeTotalRange.setNumberFormat('$#,##0.00');
  expenseTotalRange.setNumberFormat('$#,##0.00');
  eodBalanceRange.setNumberFormat('$#,##0.00');
  incomeTotalRange.setFontWeight('bold').setBorder(true, null, null, null, null, null, '#999999', SpreadsheetApp.BorderStyle.SOLID_THICK);
  expenseTotalRange.setFontColor('#c0392b');
  eodBalanceRange.setFontWeight('bold').setBorder(true, null, null, null, null, null, '#356854', SpreadsheetApp.BorderStyle.SOLID_THICK);

  const rules = newSheet.getConditionalFormatRules();
  const newRules = [];
  const transactionRangeCF = newSheet.getRange(LAYOUT.TRANSACTION_START_ROW, 1, numTransactionRows, daysInMonth);
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#008000').setRanges([transactionRangeCF]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#FF0000').setRanges([transactionRangeCF]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor('#356854').setBackground('#eaffea').setBold(true).setRanges([eodBalanceRange]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#b00020').setBackground('#ffeaea').setBold(true).setRanges([eodBalanceRange]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#228B22').setRanges([incomeTotalRange]).build());
  newSheet.setConditionalFormatRules(rules.concat(newRules));

  // --- FINALIZATION ---
  SpreadsheetApp.flush();
  Logger.log("populateNextMonthBudget completed successfully.");
  ui.alert(`✅ "${newSheetName}" created & populated.`);
}