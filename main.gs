/**
 * @OnlyCurrentDoc
 * This is a special function that runs automatically every time the spreadsheet is opened.
 * Its purpose is to create a custom menu in the spreadsheet's user interface,
 * providing easy-to-use buttons for all the main script functions.
 */
function onOpen() {
  // Get the user interface environment for the active spreadsheet.
  const ui = SpreadsheetApp.getUi();

  // Create a new menu with the name "💸 Budget Tools".
  ui.createMenu("💸 Budget Tools")
    
    // WORKFLOW 1: The standard one-click import for normal months.
    // Also serves as the final injection step after manual reconciliation.
    .addItem("🚀 Import & Inject Transactions", "importBankTransactions")
    
    // WORKFLOW 2: The optional preparation step for reconciliation months.
    // This cleans the import sheet and populates the vendor list, then stops.
    .addItem("⚙️ Prepare for Reconciliation", "prepareForReconciliation")
    
    // Adds a visual dividing line to keep the menu organized.
    .addSeparator()
    
    // Adds the button to create the next calendar month's sheet.
    .addItem("➕ Create Next Month", "populateNextMonthBudget")
    
    // Adds the button to refresh all calculations on the Dashboard.
    .addItem("🔄 Refresh Dashboard", "updateDashboard")
    
    // Add the newly created menu and all its items to the spreadsheet's UI.
    .addToUi();
}
