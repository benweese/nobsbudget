/**
 * This script file is the engine for the "Am I Screwed?" dashboard.
 * It reads data from monthly sheets, performs calculations, and updates the dashboard tab.
 */

/**
 * Reads all transaction data from a specified monthly sheet and returns it as an array of objects.
 * This is the primary data source for all other dashboard calculations.
 * @param {string} sheetName The name of the monthly sheet to read (e.g., "September 2025").
 * @return {Array<Object>} An array of structured transaction objects.
 */
function getAllMonthTransactions(sheetName) {
  // --- INITIAL SETUP AND VALIDATION ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`getAllMonthTransactions: Sheet "${sheetName}" not found.`);
    return [];
  }

  // Define the transaction area using constants from utils.gs.
  const START_ROW = LAYOUT.TRANSACTION_START_ROW;
  const END_ROW = LAYOUT.TRANSACTION_END_ROW;
  const numRows = END_ROW - START_ROW + 1;

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0 || numRows <= 0) {
    Logger.log(`getAllMonthTransactions: Sheet "${sheetName}" appears empty.`);
    return [];
  }

  // --- PARSE DATE FROM SHEET NAME ---
  // We derive the month and year from the sheet's name for reliability.
  let year, monthIndex;
  try {
    const nameParts = sheetName.match(/^(\w+)\s+(\d{4})$/);
    if (!nameParts) throw new Error("Invalid sheet name format.");
    const monthName = nameParts[1];
    year = parseInt(nameParts[2]);
    const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    monthIndex = months.indexOf(monthName); // JS months are 0-indexed.
    if (monthIndex === -1) throw new Error(`Invalid month name "${monthName}".`);
  } catch (e) {
    Logger.log(`getAllMonthTransactions: Error parsing sheet name "${sheetName}": ${e}.`);
    return [];
  }

  // --- BATCH READ DATA FROM SPREADSHEET ---
  // Read all values and notes in two efficient operations.
  let values = [];
  let notes = [];
  try {
    const range = sheet.getRange(START_ROW, 1, numRows, lastCol);
    values = range.getValues();
    notes = range.getNotes();
  } catch (e) {
    Logger.log(`getAllMonthTransactions: Error reading data range in "${sheetName}": ${e}.`);
    return [];
  }

  // --- PROCESS DATA IN MEMORY ---
  // Loop through the in-memory arrays to build the transaction list.
  const transactions = [];
  for (let j = 0; j < lastCol; j++) {
    const colNum = j + 1;
    for (let i = 0; i < numRows; i++) {
      const value = values[i][j];
      const rowNum = i + START_ROW;
      // Skip empty cells and the A2 carry-over balance.
      if (value === '' || value === null) continue;
      if (rowNum === LAYOUT.CARRY_OVER_ROW && colNum === 1) continue;

      // Create a structured object for each transaction.
      try {
        const date = new Date(year, monthIndex, colNum);
        const amount = parseFloat(value);
        transactions.push({
          date: date,
          amount: isNaN(amount) ? 0 : amount,
          note: notes[i][j] || "",
          column: colNum,
          row: rowNum
        });
      } catch (e) {
        Logger.log(`Error processing cell ${sheetName}!R${rowNum}C${colNum}: ${e}`);
      }
    }
  }
  return transactions;
}

/**
 * Calculates total cash income and expenses for a given month from its transaction list.
 * @param {string} sheetName The name of the monthly sheet (e.g., "September 2025").
 * @return {object} An object with { income: number, expenses: number }.
 */
function getMonthlyCashFlowSummary(sheetName) {
  // Get all transactions for the month.
  const transactions = getAllMonthTransactions(sheetName);
  if (!transactions) {
    return { income: 0, expenses: 0 };
  }
  // Sum up positive (income) and negative (expense) amounts.
  let totalIncome = 0;
  let totalExpenses = 0;
  transactions.forEach(transaction => {
    if (transaction.amount > 0) {
      totalIncome += transaction.amount;
    } else if (transaction.amount < 0) {
      totalExpenses += transaction.amount;
    }
  });
  return {
    income: totalIncome,
    expenses: totalExpenses
  };
}

/**
 * Reads the 'Recurring' sheet and calculates the total estimated amount of all
 * predictable, fixed expenses for a given month.
 * @param {string} sheetName The name of the month to calculate for (e.g., "September 2025").
 * @return {number} The total amount of fixed expenses for the month (as a negative number).
 */
function getMonthlyRecurringTotal(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recurringSheet = ss.getSheetByName(SHEET_NAMES.RECURRING);
  if (!recurringSheet) return 0;

  // Parse the month and year from the sheet name.
  const nameParts = sheetName.match(/^(\w+)\s+(\d{4})$/);
  if (!nameParts) return 0;
  const monthName = nameParts[1];
  const year = parseInt(nameParts[2]);
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthIndex = months.indexOf(monthName);
  if (monthIndex === -1) return 0;
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

  let totalRecurringExpenses = 0;
  const recurringData = recurringSheet.getRange(2, 1, recurringSheet.getLastRow() - 1, 5).getValues();

  // Loop through each rule to see if it applies to any day in the target month.
  recurringData.forEach(([dayOfMonthRaw, name, amount, frequency, startDateRaw]) => {
    // Only sum up expenses (negative amounts).
    if (!amount || isNaN(parseFloat(amount)) || parseFloat(amount) >= 0) {
      return; 
    }
    const numericAmount = parseFloat(amount);
    
    for (let day = 1; day <= daysInMonth; day++) {
      const thisDate = new Date(year, monthIndex, day);
      // Create a rule object to pass to our central logic function.
      const rule = { 
        dayOfMonthRaw: dayOfMonthRaw, 
        frequency: frequency, 
        startDate: startDateRaw instanceof Date ? startDateRaw : (startDateRaw ? new Date(startDateRaw) : null)
      };
      // Use the central, authoritative function from utils.gs to check for a match.
      const isMatch = isRecurringDateMatch(thisDate, rule);
      if (isMatch) {
        totalRecurringExpenses += numericAmount;
      }
    }
  });
  return totalRecurringExpenses;
}


/**
 * MASTER DASHBOARD UPDATE FUNCTION
 * This is the single function to call to refresh all metrics on the dashboard.
 * It calculates and displays the monthly summary and the daily spend limit.
 */
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("Dashboard");
  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('The "Dashboard" sheet is missing. Please create it.');
    return;
  }

  // --- Date & Sheet Setup ---
  const now = new Date();
  const today = now.getDate();
  const currentMonthSheetName = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MMMM yyyy");
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const daysRemaining = daysInMonth - today + 1;

  // --- Calculations ---
  // 1. Get all cash flow transactions that have happened so far this month.
  const summary = getMonthlyCashFlowSummary(currentMonthSheetName);
  
  // 2. Get the total of all PREDICTED fixed bills for the entire month.
  const totalBills = getMonthlyRecurringTotal(currentMonthSheetName);
  
  // 3. Calculate how much money is left for non-fixed spending.
  // Formula: (Income To Date) + (Total Bills for Month) + (Expenses To Date)
  const discretionaryBudgetLeft = summary.income + totalBills + summary.expenses;

  // 4. Calculate the daily limit for remaining days.
  let dailyLimit = 0;
  if (daysRemaining > 0 && discretionaryBudgetLeft > 0) {
    dailyLimit = discretionaryBudgetLeft / daysRemaining;
  }

  // --- Display on Dashboard ---
  dashboardSheet.getRange("A2").setValue("Total Income:");
  dashboardSheet.getRange("B2").setValue(summary.income).setNumberFormat('$#,##0.00');
  
  dashboardSheet.getRange("A3").setValue("Total Expenses:");
  dashboardSheet.getRange("B3").setValue(summary.expenses).setNumberFormat('$#,##0.00;(#,##0.00)');
  
  dashboardSheet.getRange("A4").setValue("Daily Spend Limit (Non-Fixed):");
  dashboardSheet.getRange("B4").setValue(dailyLimit).setNumberFormat('$#,##0.00');
  
  SpreadsheetApp.getUi().alert('✅ Dashboard has been refreshed!');
  Logger.log(`Dashboard Refreshed. Daily Spend Limit: ${dailyLimit}`);
}