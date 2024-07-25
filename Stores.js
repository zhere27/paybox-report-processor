function refreshStores() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("Kiosk %");

  // Clear the range from A2:B to the last row, as well as any formatting
  var lastRow = sheet.getLastRow();
  sheet.getRange("A2:B" + lastRow).clear().setFontFamily('Century Gothic').setFontSize(9);

  try {
    const rows = getStores();

    if (!rows || rows.length === 0) {
      Logger.log("No data to populate the new sheet.");
      return;
    }

    // Prepare the data rows for insertion
    const data = rows.map(row => row.f.map(cell => cell.v));

    // Define the range where new data will be appended
    var range = sheet.getRange(2, 1, data.length, data[0].length);

    // Set the new data values
    range.setValues(data);

    // Activate the sheet
    sheet.activate();

  } catch (e) {
    Logger.log("Error populating sheet: " + e.message);
  }

  sortRange();
}