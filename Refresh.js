function test() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");

  sheet.insertColumnsAfter(18, 1);
  sheet.deleteColumns(3, 1);

  const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:00:00");
  sheet.getRange('Q1').setValue(formattedDate).setNumberFormat("MM/dd HH:00AM/PM").setBackground('#d9d9d9');
  sheet.getRange('R1').setValue('Current Cash Amount').setBackground('#d9d9d9');
  sheet.getRange('S1').setValue('Remarks').setBackground('#d9d9d9');
  sheet.getRange('T1').setValue('Last Requested for Collection').setBackground('#d9d9d9');
  sheet.getRange('U1').setValue('Movement').setBackground('#d9d9d9');
}

function refresh() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("Kiosk %");
  const configSheet = spreadSheet.getSheetByName('Config');

  // Ensure required sheets are present
  if (!sheet || !configSheet) {
    Logger.log("Required sheets not found.");
    return;
  }

  // Call process monitoring if B1 is false
  if (!configSheet.getRange('B1').getValue()) {
    processMonitoringHourly();
  }

  let now = new Date();

  // Late Processing if B2 is true
  if (configSheet.getRange('B2').getValue()) {
    const procDate = new Date(configSheet.getRange('B3').getValue());
    const formattedDate = Utilities.formatDate(procDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const [year, month, day] = formattedDate.split("-").map(Number);
    const hours = parseInt(configSheet.getRange('B4').getValue().substring(0, 2));
    now = new Date(year, month - 1, day, hours, 0, 0, 0);
  }

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

  const range = sheet.getRange("A2:S");
  const columnToSortBy = 17;

  range.sort({ column: columnToSortBy, ascending: false });
}

function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kiosk %");
  const lastRow = sheet.getLastRow();

  if (!sheet) {
    Logger.log("Sheet 'Kiosk %' not found.");
    return;
  }

  // Clear existing conditional formatting rules
  sheet.setConditionalFormatRules([]);

  // Define ranges
  const range = sheet.getRange(`C2:Q${lastRow}`); // Adjust the range as needed
  const rangeNoMovement = sheet.getRange(`V2:V${lastRow}`); // Adjust the range as needed
  const rangeCollectedStore = sheet.getRange(`A2:A${lastRow}`); // Adjust the range as needed

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
      .setBackground('#f1c232')
      .setFontColor('black')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.7)
      .setBackground('#c4ecc4')
      .setFontColor('black')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.6)
      .setBackground('#fff2cc')
      .setFontColor('black')
      .setRanges([range])
      .build()
  ];

  // Define conditional formatting rules for the no movement range
  const rulesNoMovement = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('No changes')
      .setFontColor('#a61c00')
      .setRanges([rangeNoMovement])
      .build()
  ];

  const rulesCollectedStores = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=A2=W2')
      .setBackground('#ff9900')
      .setFontColor('#black')
      .setRanges([rangeCollectedStore])
      .build()
  ];

  // Combine all rules
//  const allRules = rules.concat(rulesNoMovement).concat(rulesCollectedStores);
  const allRules = [...rules, ...rulesNoMovement, ...rulesCollectedStores];

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


// function refresh() {
//   var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = spreadSheet.getSheetByName("Kiosk %");
//   var configSheet = spreadSheet.getSheetByName('Config');

//   //call process monitoring
//   if (configSheet.getRange('B1').getValue() === false) {
//     processMonitoringHourly();
//   }

//   var now = new Date();
//   //Set to True if late processing
//   if (configSheet.getRange('B2').getValue() === true) {
//     var procDate = new Date(configSheet.getRange('B3').getValue());
//     var formattedDate = Utilities.formatDate(procDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//     var dateParts = formattedDate.split("-");
//     var year = parseInt(dateParts[0]);
//     var month = parseInt(dateParts[1]) - 1; // Months are zero-based in JavaScript Date
//     var day = parseInt(dateParts[2]);
//     var hours = parseInt(configSheet.getRange('B4').getValue().substring(0, 2));

//     now = new Date(year, month, day, hours, 0, 0, 0);
//   }

//   var lastRow = sheet.getLastRow();
//   var sheetName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMdd HH00yy");
//   var formula = "=IFNA(TEXT(QUERY('" + sheetName + "'!$A:$D,\"select D where A='\" & TRIM(A2) & \"'\",0),\"0.00%\")/100,0)";
//   var formulaRange = sheet.getRange("Q2:Q" + lastRow).setFontFamily('Century Gothic').setFontSize(9).setHorizontalAlignment("Center");

//   var formulaCurrentAmount = "=IFNA(QUERY('" + sheetName + "'!$A:$D,\"select B where A='\" & TRIM(A2) & \"'\",0),0)";
//   var formulaCurrentAmountRange = sheet.getRange("R2:R" + lastRow).clear().setFontFamily('Century Gothic').setFontSize(9);

//   var existingSheet = spreadSheet.getSheetByName(sheetName);
//   if (existingSheet) {
//     spreadSheet.deleteSheet(existingSheet);
//     Logger.log("Deleted existing sheet: " + sheetName);
//   }

//   var newSheet = spreadSheet.insertSheet(sheetName, 1);
//   if (!newSheet) {
//     Logger.log("Failed to create new sheet: " + sheetName);
//     return;
//   }

//   try {
//     const trnDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
//     const hours = String(now.getHours()).padStart(2, '0');
//     const rows = populateSheet(trnDate, hours + ":00:00");

//     if (!rows || rows.length === 0) {
//       Logger.log("No data to populate the new sheet.");
//       return;
//     }

//     // Write the data rows
//     rows.forEach(row => {
//       const data = row.f.map(cell => cell.v);
//       newSheet.appendRow(data);
//     });

//     newSheet.activate();
//     Logger.log("New sheet created and populated successfully: " + sheetName);
//   } catch (e) {
//     Logger.log("Error populating sheet: " + e.message);
//   }




//   var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:00:00");

//   sheet.deleteColumns(3, 1);
//   sheet.getRange('Q1').setValue(formattedDate).setNumberFormat("MM/dd HH:00AM/PM");
//   sheet.getRange('R1').setValue('Current Cash Amount');
//   sheet.getRange('S1').setValue('Remarks');

//   if (!sheet) {
//     Logger.log("Sheet 'Kiosk %' not found.");
//     return;
//   }


//   try {
//     formulaRange.setFormula(formula).setNumberFormat("0.00%");
//     formulaCurrentAmountRange.setFormula(formulaCurrentAmount).setNumberFormat("###,###,##0");
//   } catch (e) {
//     Logger.log("Error setting formula: " + e.message);
//     return;
//   }




//   applyConditionalFormatting();
// }

// function sortRange() {
//   var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = spreadSheet.getSheetByName("Kiosk %");
//   var range = sheet.getRange("A2:S"); // Adjust the range as needed
//   var columnToSortBy = 18; // Column AE is the 31st column (A=1, B=2, ..., AE=31)

//   range.sort({
//     column: columnToSortBy,
//     ascending: false // Sort in descending order
//   });
// }

// function applyConditionalFormatting() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kiosk %"); // Replace "Kiosk %" with your sheet name

//   sheet.setConditionalFormatRules([]);

//   // Define the range to apply conditional formatting
//   var range = sheet.getRange("C2:Q"); // Change to the range where you want to apply the formatting

//   // Get the existing conditional format rules
//   var rules = sheet.getConditionalFormatRules();

//   // Create a new conditional format rule
//   var rule1 = SpreadsheetApp.newConditionalFormatRule()
//     .whenNumberGreaterThanOrEqualTo(0.9) // Checks if the value is greater than 90%
//     .setBackground('#cc0000')   // Red background color
//     .setFontColor('White')    // White font color
//     .setRanges([range])         // Apply the rule to the specified range
//     .build();

//   var rule2 = SpreadsheetApp.newConditionalFormatRule()
//     .whenNumberGreaterThan(0.8) // Checks if the value is greater than 90%
//     .setBackground('#f1c232')   // Red background color
//     .setFontColor('Black')    // White font color
//     .setRanges([range])         // Apply the rule to the specified range
//     .build();

//   var rule3 = SpreadsheetApp.newConditionalFormatRule()
//     .whenNumberGreaterThan(0.7) // Checks if the value is greater than 90%
//     .setBackground('#c4ecc4')   // Red background color
//     .setFontColor('Black')    // White font color
//     .setRanges([range])         // Apply the rule to the specified range
//     .build();

//   // Add the new rule to the existing rules
//   rules.push(rule1);
//   rules.push(rule2);
//   rules.push(rule3);

//   // Set the updated rules to the sheet
//   sheet.setConditionalFormatRules(rules);

//   Logger.log("Conditional formatting applied successfully.");
// }


// // function deleteOldSheet(){
// //   var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

// //   //delete sheet
// //   var sheets = spreadSheet.getSheets();
// //   var sheetIndex = 17; //index to delete

// //   if (sheetIndex >= 0 && sheetIndex < sheets.length) {
// //     var sheetToDelete = sheets[sheetIndex];
// //     spreadSheet.deleteSheet(sheetToDelete);
// //     Logger.log("Sheet at index " + sheetIndex + " deleted.");
// //   } else {
// //     Logger.log("Sheet index " + sheetIndex + " is out of range.");
// //   }
// // }

