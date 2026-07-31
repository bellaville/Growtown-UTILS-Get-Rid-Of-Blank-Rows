/**
 * Function to delete blank rows in a specific google sheet (should ensure there arent many blank rows)
 */
function removeBlankRows(sheetName) {
    sheetName = sheetName.trim();

    const workLogSheet = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties().getProperty('WORK_LOG_SHEET_ID')
    );
    const sheet = workLogSheet.getSheetByName(sheetName);
    if (!sheet) {
        Logger.log("Sheet not found: " + sheetName);
        return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // Nothing to check

    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) return; // No columns

    // Read all data at once
    const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues(); // skip header row

    // Collect indices of blank rows (1-based row numbers)
    const rowsToDelete = [];
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row.every(cell => cell === "" || cell === null)) {
            rowsToDelete.push(i + 2); // +2 because data starts at row 2
        }
    }

    // Delete from bottom to top so row numbers don't shift
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
        Logger.log("Deleted row: " + rowsToDelete[i] + " in table: " + sheetName);
    }
}

function runAllSheets(sheetNames) {

    Logger.log(sheetNames);

    for (let i = 0; i < sheetNames.length; i++) {
        removeBlankRows(sheetNames[i]);
    }
}

function TEST() {
    removeBlankRows("Copy of Outputs");
}