function refresh() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");

  var firstFile = GDriveFilesAPI.getFirstFileInFolder('1cAFjrqlAWh4Di_6MflkEdL4KdVveStG1');
  if (firstFile !== null) {
    var numbers = extractNumbersFromText(firstFile.getName());
    var dateParts = breakdownDateTime(numbers.toString());
    Logger.log('First file name: ' + firstFile.getName());
  }

  PayboxReportProcessor.processMonitoringHourly();

  let now = new Date(dateParts.year, dateParts.month-1, dateParts.day, dateParts.hour, dateParts.minute);
  const lastRow = sheet.getLastRow();
  const sheetName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMdd HH00yy");

  // Define formulas
  sheet.insertColumnsAfter(18, 1);
  sheet.deleteColumns(3, 1);

  const formula = `=IFNA(TEXT(QUERY('${sheetName}'!$A:$D,"select D where A='" & TRIM(A2) & "'",0),"0.00%")/100,0)`;
  const formulaCurrentAmount = `=IFNA(QUERY('${sheetName}'!$A:$D,"select B where A='" & TRIM(A2) & "'",0),0)`;
  const formulaSparkline = `=SPARKLINE(INDIRECT("C"&ROW()&":P"&ROW()))`;
  const formulaNoMovement = '=IF(INDIRECT("G"&ROW())<>"",IF(COUNTIF(INDIRECT("G"&ROW()&":Q"&ROW()),"<>"& INDIRECT("G"&ROW()))=0, "No changes", "Values changed"),"")';
  const formulaCollectedStores = '=IFNA(VLOOKUP(A2,\'Collected last 3 days\'!B:B,1,false))';


  // Set formulas with formatting
  const formulaRange = sheet.getRange(`Q2:Q${lastRow}`)
    .setFontFamily('Century Gothic')
    .setFontSize(9)
    .setHorizontalAlignment("Center")
    .setFormula(formula)
    .setNumberFormat("0.00%");

  const formulaCurrentAmountRange = sheet.getRange(`R2:R${lastRow}`)
    .clear()
    .setFontFamily('Century Gothic')
    .setFontSize(9)
    .setFormula(formulaCurrentAmount)
    .setNumberFormat("###,###,##0");

  const formulaSparklineRange = sheet.getRange(`U2:U${lastRow}`)
    .clear()
    .setFontFamily('Century Gothic')
    .setFontSize(9)
    .setFormula(formulaSparkline);

  const formulaNoMovementRange = sheet.getRange(`V2:V${lastRow}`)
    .clear()
    .setFontFamily('Century Gothic')
    .setFontSize(9)
    .setFormula(formulaNoMovement);

  const formulaCollectedStoresRange = sheet.getRange(`W2:W${lastRow}`)
    .clear()
    .setFontFamily('Century Gothic')
    .setFontSize(9)
    .setFormula(formulaCollectedStores);

  // Manage sheets
  const existingSheet = spreadSheet.getSheetByName(sheetName);
  if (existingSheet) {
    spreadSheet.deleteSheet(existingSheet);
    Logger.log(`Deleted existing sheet: ${sheetName}`);
  }

  const newSheet = spreadSheet.insertSheet(sheetName, 3);
  if (!newSheet) {
    Logger.log(`Failed to create new sheet: ${sheetName}`);
    return;
  }

  // Populate the new sheet
  try {
    const trnDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const hours = String(now.getHours()).padStart(2, '0');
    const rows = populateSheet(trnDate, `${hours}:00:00`);

    if (!rows || rows.length === 0) {
      Logger.log("No data to populate the new sheet.");
      return;
    }

    // Check the structure of the rows array
    if (rows.length > 0 && rows[0].f && rows[0].f.length > 0) {
      // Populate the sheet with the mapped values
      newSheet.getRange(1, 1, rows.length, rows[0].f.length).setValues(rows.map(row => row.f.map(cell => cell.v)));
    } else {
      Logger.log("Error: rows array is not in the expected format.");
    }
  } catch (e) {
    Logger.log(`Error populating sheet: ${e.message}`);
  }


  // Update main sheet
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:00:00");
  sheet.getRange('Q1').setValue(formattedDate).setNumberFormat("MM/dd HH:00AM/PM").setBackground('#d9d9d9');
  sheet.getRange('R1').setValue('Current Cash Amount').setBackground('#d9d9d9');
  sheet.getRange('S1').setValue('Remarks').setBackground('#d9d9d9');
  sheet.getRange('T1').setValue('Last Requested for Collection').setBackground('#d9d9d9');
  sheet.getRange('U1').setValue('Movement').setBackground('#d9d9d9');
  sheet.getRange('V1').setValue('No Movement for 3 days').setBackground('#d9d9d9');
  sheet.getRange('W1').setValue('Collected Stores').setBackground('#d9d9d9');
  sheet.activate();

  getCollectedStores();

  //Sort
  sortRange();

  // Apply conditional formatting
  applyConditionalFormatting();

  deleteOldSheet();
}

function getCollectedStores() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("Collected last 3 days");

  sheet.getRange('A2:E').clear();

  var query = `
 select created_at
 ,machine_name
 ,amount
 ,count
 ,collector_name 
from ms-paybox-prod-1.pldtsmart.collections 
where date(created_at) >= DATE(CURRENT_DATE())-3 and date(created_at) <= DATE(CURRENT_DATE())
order by created_at desc
    `;

  // Logger.log(query);

  var queryResult = executeQueryAndWait(query);

  var rows = queryResult.rows;
  if (rows) {
    var ctr = 1;
    for (var i = 0; i < rows.length; i++) {
      ctr++;

      let timestamp = rows[i].f[0].v * 1000; // Multiply by 1000 to convert to milliseconds
      let dateValue = new Date(timestamp); // Create Date object from timestamp
      let formattedDate = Utilities.formatDate(dateValue, 'UTC', 'yyyy-MM-dd HH:mm:ss');

      sheet.getRange('A' + ctr).setValue(formattedDate).setFontFamily("Century Gothic").setFontSize(9).setHorizontalAlignment("Left");  //Machine Name
      sheet.getRange('B' + ctr).setValue(rows[i].f[1].v).setFontFamily("Century Gothic").setFontSize(9).setHorizontalAlignment("Left"); //Account No.
      sheet.getRange('C' + ctr).setValue(rows[i].f[2].v).setFontFamily("Century Gothic").setFontSize(9).setHorizontalAlignment("Right").setNumberFormat("###,###,##0.00"); //Trans. Ref. No.
      sheet.getRange('D' + ctr).setValue(rows[i].f[3].v).setFontFamily("Century Gothic").setFontSize(9).setHorizontalAlignment("Center").setNumberFormat("###,###,##0");
      sheet.getRange('E' + ctr).setValue(rows[i].f[4].v).setFontFamily("Century Gothic").setFontSize(9).setHorizontalAlignment("Left"); //Status

    }
  }

}


function sortRange() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");

  const range = sheet.getRange("A2:T");
  const columnToSortBy = 17;

  range.sort({ column: columnToSortBy, ascending: false });
}

function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kiosk %");

  if (!sheet) {
    Logger.log("Sheet 'Kiosk %' not found.");
    return;
  }

  // Clear existing conditional formatting rules
  sheet.setConditionalFormatRules([]);

  // Define ranges
  const range = sheet.getRange("C2:Q1000"); // Adjust the range as needed
  const rangeNoMovement = sheet.getRange("V2:V1000"); // Adjust the range as needed
  const rangeCollected = sheet.getRange("A2:A1000");

  // Define conditional formatting rules for the main range
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.9)
      .setBackground('#cc0000')
      .setFontColor('white')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.8)
      .setBackground('#e69138')
      .setFontColor('white')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.7)
      .setBackground('#ffe599')
      .setFontColor('black')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.6)
      .setBackground('#fff2cc')
      .setFontColor('black')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.5)
      .setBackground('#d9ead3')
      .setFontColor('black')
      .setRanges([range])
      .build()      
  ];

  // Define conditional formatting rules for the no movement range
  const rulesNoMovement = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('No changes')
      .setFontColor('#85200c')
      .setRanges([rangeNoMovement])
      .build()
  ];

const rulesCollected = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=A2=W2')
      .setBackground('#ff9900')
      .setRanges([rangeCollected])
      .build()
  ];

  // Combine all rules
  const allRules = rules.concat(rulesNoMovement).concat(rulesCollected);

  // Set combined conditional formatting rules
  sheet.setConditionalFormatRules(allRules);
  Logger.log("Conditional formatting applied successfully.");
}


function deleteOldSheet() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  //delete sheet
  var sheets = spreadSheet.getSheets();
  var sheetIndex = 18; //index to delete

  if (sheetIndex >= 0 && sheetIndex < sheets.length) {
    var sheetToDelete = sheets[sheetIndex];
    spreadSheet.deleteSheet(sheetToDelete);
    Logger.log("Sheet at index " + sheetIndex + " deleted.");
  } else {
    Logger.log("Sheet index " + sheetIndex + " is out of range.");
  }
}
