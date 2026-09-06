/* exported SHEET_NAMES, LAYOUT, columnToLetter, isRecurringDateMatch */

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
  let temp, letter = '';          // ← was: let temp = '', letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * The single source of truth for determining if a recurring item rule matches a given date.
 * Supports week-based cadences (weekly/biweekly/triweekly or "N Weeks(Weekday)") and
 * month-based cadences (monthly/bimonthly/quarterly/semiannually/annually or "N Months"/"N Years").
 * @param {Date} dateToCheck The date being tested.
 * @param {Object} rule The rule: { dayOfMonthRaw, frequency, startDate }.
 * @return {boolean} True if the rule fires on dateToCheck.
 */
function isRecurringDateMatch(dateToCheck, rule) {
  const { dayOfMonthRaw, frequency, startDate } = rule;
  const freq = String(frequency).toLowerCase();

  const targetWeekday = freq.match(/\((\w+)\)/i)?.[1]?.toLowerCase();
  const weekDayMap = { sunday: 0, monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6 };

  // --- WEEK-INTERVAL FREQUENCIES ---
  const WEEK_ALIASES = { weekly: 1, biweekly: 2, triweekly: 3 };
  let interval = null;

  for (const alias in WEEK_ALIASES) {
    if (freq.startsWith(alias)) { interval = WEEK_ALIASES[alias]; break; }
  }
  if (interval === null) {
    const numericMatch = freq.match(/(\d+)\s*week/);
    if (numericMatch && parseInt(numericMatch[1], 10) >= 1) {
      interval = parseInt(numericMatch[1], 10);
    }
  }

  if (interval !== null) {
    if (!targetWeekday || dateToCheck.getDay() !== weekDayMap[targetWeekday]) return false;
    if (interval === 1) return true;
    if (!startDate || dateToCheck < startDate) return false;

    const MS_PER_DAY = 86400000;
    const toDayNumber = (d) => Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / MS_PER_DAY);

    // Snap to the first target-weekday on or after the start date. This way a start
    // date that isn't itself on the anchor weekday still yields clean weekly intervals
    // instead of silently never matching.
    const startOffset = (weekDayMap[targetWeekday] - startDate.getDay() + 7) % 7;
    const anchorDayNumber = toDayNumber(startDate) + startOffset;

    const weeksApart = (toDayNumber(dateToCheck) - anchorDayNumber) / 7;
    return weeksApart >= 0 && weeksApart % interval === 0;
  }

  // --- MONTH-INTERVAL FREQUENCIES ---
  const MONTH_ALIASES = { monthly: 1, bimonthly: 2, quarterly: 3, semiannually: 6, annually: 12 };
  let monthInterval = null;

  for (const alias in MONTH_ALIASES) {
    if (freq.startsWith(alias)) { monthInterval = MONTH_ALIASES[alias]; break; }
  }
  if (monthInterval === null) {
    const yearMatch = freq.match(/(\d+)\s*year/);
    const monthMatch = freq.match(/(\d+)\s*month/);
    if (yearMatch && parseInt(yearMatch[1], 10) >= 1) {
      monthInterval = parseInt(yearMatch[1], 10) * 12;
    } else if (monthMatch && parseInt(monthMatch[1], 10) >= 1) {
      monthInterval = parseInt(monthMatch[1], 10);
    }
  }

  if (monthInterval !== null) {
    const lastDayOfMonth = new Date(dateToCheck.getFullYear(), dateToCheck.getMonth() + 1, 0).getDate();
    const rawDay = String(dayOfMonthRaw).toLowerCase().trim();
    let targetDay;

    if (rawDay === 'last') {
      targetDay = lastDayOfMonth;
    } else if (rawDay && !isNaN(parseInt(rawDay, 10))) {
      targetDay = parseInt(rawDay, 10);
    } else if (startDate) {
      targetDay = startDate.getDate();
    } else {
      return false;
    }

    // Clamp overflow so a "31" rule fires on the last day of short months instead of vanishing.
    if (targetDay > lastDayOfMonth) targetDay = lastDayOfMonth;
    if (dateToCheck.getDate() !== targetDay) return false;
    if (monthInterval === 1) return true;
    if (!startDate || dateToCheck < startDate) return false;

    const monthDiff = (dateToCheck.getFullYear() - startDate.getFullYear()) * 12
                    + (dateToCheck.getMonth() - startDate.getMonth());

    return monthDiff >= 0 && monthDiff % monthInterval === 0;
  }

  return false;
}
