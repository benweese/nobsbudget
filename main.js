/**
 * @OnlyCurrentDoc
 * This is a special function that runs automatically every time the spreadsheet is opened.
 * Its purpose is to create a custom menu in the spreadsheet's user interface,
 * providing easy-to-use buttons for all the main script functions.
 */
// eslint-disable-next-line no-unused-vars
function onOpen() {
  // Get the user interface environment for the active spreadsheet.
  const ui = SpreadsheetApp.getUi();

  // Create a new menu with the name "💸 Budget Tools".
  ui.createMenu("💸 Budget Tools")
    
    // Adds the button to create the next calendar month's sheet.
    .addItem("➕ Create Next Month", "populateNextMonthBudget")
    
    // Adds the button to refresh all calculations on the Dashboard.
    .addItem("🔄 Refresh Dashboard", "updateDashboard")
    
    // Add the newly created menu and all its items to the spreadsheet's UI.
    .addToUi();
}