/**
 * This script file serves as a central library of shared constants and helper functions
 * for the entire project. It does not contain functions that are run directly by the user.
 * Centralizing these utilities makes the code cleaner and much easier to maintain.
 */

// An object holding the exact names of key sheets. Using these constants (e.g., SHEET_NAMES.VENDORS)
// prevents bugs from typos and makes renaming a sheet a one-line change.
const SHEET_NAMES = {
  BANK_IMPORT: "BankImport",
  VENDORS: "Vendors",
  RECURRING: "Recurring",
};

// An object defining the fixed row structure of the monthly budget sheets. Using these constants
// (e.g., LAYOUT.INCOME_TOTAL_ROW) makes other scripts more readable and robust against layout changes.
const LAYOUT = {
  HEADER_ROW: 1,
  CARRY_OVER_ROW: 2, // A2 specifically
  TRANSACTION_START_ROW: 2,
  TRANSACTION_END_ROW: 17,
  INCOME_TOTAL_ROW: 18,    // Row for "Previous Balance + Income"
  EXPENSE_TOTAL_ROW: 19,   // Row for "Daily Expenses"
  EOD_BALANCE_ROW: 20,     // Row for "End of Day Balance"
  TOTAL_ROWS: 20           // The total number of rows in the fixed sheet structure
};

/**
 * Converts a column number (e.g., 3) into its letter representation (e.g., "C").
 * Necessary for building A1-style cell references dynamically.
 * @param {number} column The column number (1-based).
 * @return {string} The column letter(s).
 */
function columnToLetter(column) {
  let temp = '', letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * The single source of truth for determining if a recurring item rule matches a given date.
 * Centralizing this logic prevents bugs and ensures consistency across the application.
 * @param {Date} dateToCheck The date we are checking.
 * @param {Object} rule An object with the rule's properties: {dayOfMonthRaw, frequency, startDate}.
 * @return {boolean} True if the rule matches the date.
 */
function isRecurringDateMatch(dateToCheck, rule) {
  const { dayOfMonthRaw, frequency, startDate } = rule;
  const freq = String(frequency).toLowerCase();

  if (freq.startsWith('monthly')) {
    const dayOfMonth = String(dayOfMonthRaw).toLowerCase();
    if (dayOfMonth === 'last') {
      const lastDayOfMonth = new Date(dateToCheck.getFullYear(), dateToCheck.getMonth() + 1, 0).getDate();
      return dateToCheck.getDate() === lastDayOfMonth;
    }
    return parseInt(dayOfMonthRaw) === dateToCheck.getDate();
  }

  const targetWeekday = freq.match(/\((\w+)\)/i)?.[1]?.toLowerCase();
  const weekDayMap = { 'sunday': 0, 'monday': 1, 'tuesday': 2, 'wednesday': 3, 'thursday': 4, 'friday': 5, 'saturday': 6 };
  
  if (freq.startsWith('weekly')) {
    if (!targetWeekday) return false;
    return dateToCheck.getDay() === weekDayMap[targetWeekday];
  }
  
  if (freq.startsWith('biweekly')) {
    if (!startDate || !targetWeekday || dateToCheck < startDate) return false;
    if (dateToCheck.getDay() !== weekDayMap[targetWeekday]) return false;
    // Zero out the time to prevent daylight saving issues from affecting week calculation.
    const diffTime = Math.abs(dateToCheck.setHours(0,0,0,0) - startDate.setHours(0,0,0,0));
    const diffWeeks = Math.floor(diffTime / (1000 * 60 * 60 * 24 * 7));
    return diffWeeks % 2 === 0; // It's a match if it's an even number of weeks from the start.
  }
  
  if (freq.startsWith('bimonthly')) {
    if (!startDate || dateToCheck < startDate) return false;
    // Bimonthly is based on the day of the month of the start date.
    if (dateToCheck.getDate() !== startDate.getDate()) return false;
    const monthDiff = (dateToCheck.getMonth() - startDate.getMonth()) + (12 * (dateToCheck.getFullYear() - startDate.getFullYear()));
    return monthDiff >= 0 && monthDiff % 2 === 0; // It's a match if it's an even number of months from the start.
  }
  
  if (freq.startsWith('annually')) {
    if (!startDate) return false;
    return dateToCheck.getMonth() === startDate.getMonth() && dateToCheck.getDate() === startDate.getDate();
  }
  
  return false;
}