function collectedStores() {
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
