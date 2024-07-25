function processMonitoringHourly() {
  var folderCollections = DriveApp.getFolderById('1cAFjrqlAWh4Di_6MflkEdL4KdVveStG1');

  // Get all CSV files in the folder
  var files = folderCollections.getFilesByType(MimeType.CSV);

  // Loop through CSV files and import them into separate sheets
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    runLoadCsvFromDrive(file.getId());
    GDriveFilesAPI.deleteFileByFileId(fileId);
  }
  
  function runLoadCsvFromDrive(fileId) {
    // Fetch the CSV file content from Google Drive
    const file = DriveApp.getFileById(fileId);
    const csvContent = file.getBlob().getDataAsString();

    var csvString = validateLine(csvContent, file.getName());

    const csvData = Utilities.parseCsv(csvString)
      .map(
        row => [
          row[0],
          row[1],
          row[2],
          parseFloat(row[3]),
          row[4],
          parseFloat(row[5])
        ]
      ); // Retrieve only specific columns

    // Insert data into BigQuery table
    uploadToBQ(csvData);
  }

  function validateLine(data, fileName) {
    var match = fileName.match(/_(\d{12})/);
    const dateStr = match[1];
    const year = dateStr.substring(0, 4);
    const month = dateStr.substring(4, 6);
    const day = dateStr.substring(6, 8);
    const hours = dateStr.substring(8, 10);
    const mins = dateStr.substring(10, 12);

    const trnDate = new Date(`${year}-${month}-${day} ${hours}:${mins}:00`);
    trnDate.setDate(trnDate.getDate());
    const formattedDate = trnDate.toISOString().split('T')[0];

    var validLines = [];

    var lines = data.split("\n").filter(function (line) {
      return line.trim() !== ""; // Remove empty lines
    });

    lines.forEach(function (line) {
      if (line.startsWith("Partner Name")) {
        return; // Skip this iteration of the loop
      }

      var regex = /^(?<partner_name>[^,]+),(?<machine_name>[^,]+),(?<current_cash_amount>[1-9]\d*)?,(?<machine_status>[^,]+),(?<bv_health>0|[1-9]\d*)(?<bv_healthDecimalPlace>\.\d*)?$/gm;

      var matches = line.matchAll(regex);
      for (const match of matches) {

        var bill_validator = match.groups.bv_healthDecimalPlace ? parseFloat(match.groups.bv_health + match.groups.bv_healthDecimalPlace) : parseFloat(match.groups.bv_health);
        // var date = new Date(trnDate);
        const year = trnDate.getFullYear();
        const month = String(trnDate.getMonth() + 1).padStart(2, '0'); // Months are zero-indexed
        const day = String(trnDate.getDate()).padStart(2, '0');
        const hours = String(trnDate.getHours()).padStart(2, '0');
        const minutes = String(trnDate.getMinutes()).padStart(2, '0');
        const seconds = String(trnDate.getSeconds()).padStart(2, '0');
        const formattedDateTime = isNaN(month) ? '1970-01-01 00:00:00' : `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;

        var current_cash_amount = isNaN(match[3]) ? 0 : match[3];

        var pushLine = formattedDateTime + ',' + match[1] + ',' + match[2] + ',' + current_cash_amount + ',' + match[4] + ',' + bill_validator;

        validLines.push(pushLine);
      }
    });

    return validLines.join("\n"); // Join the valid lines into a single string
  }
}

function uploadToBQ(dataArray) {
  // Replace 'ms-paybox-prod-1', 'pldtsmart', and 'monitoring' with your actual values
  var projectId = 'ms-paybox-prod-1';
  var datasetId = 'pldtsmart';
  var tableId = 'monitoring_hourly';

  // Define the schema for the BigQuery table
  var schema = [
    { name: 'created_at', type: 'TIMESTAMP' },
    { name: 'partner_name', type: 'STRING' },
    { name: 'machine_name', type: 'STRING' },
    { name: 'current_amount', type: 'FLOAT' },
    { name: 'machine_status', type: 'STRING' },
    { name: 'bill_validator_health', type: 'FLOAT' }
  ];

  try {
    // Prepare the BigQuery job configuration with schema and autodetection enabled
    var jobConfig = {
      configuration: {
        load: {
          destinationTable: {
            projectId: projectId,
            datasetId: datasetId,
            tableId: tableId
          },
          schema: { fields: schema },
          sourceFormat: 'CSV',
          writeDisposition: 'WRITE_APPEND',
          autodetect: false // Disable schema autodetection since we provide schema explicitly
        }
      }
    };

    // Insert data into BigQuery
    var job = BigQuery.Jobs.insert(jobConfig, projectId, Utilities.newBlob(dataArray.join('\n')));
    var jobId = job.jobReference.jobId;

    var sleepTimeMs = 500;
    while (job.status.state !== "DONE") {
      Utilities.sleep(sleepTimeMs);
      job = BigQuery.Jobs.get(projectId, jobId);
      // Logger.log('Job Status: ' + job.status.state);
    }
    if (job.status.errors != null && job.status.errors.length > 0) {
      Logger.log("Job Status: FAILED - " + job.status.errors);
    } else {
      Logger.log("BigQuery job completed successfully.")
    }
  } catch (error) {
    Logger.log("Error: " + error);
  }
}




