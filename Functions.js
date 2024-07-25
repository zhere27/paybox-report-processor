function executeQueryAndWait(query) {
  const projectId = 'ms-paybox-prod-1';
  const request = {
    query: query,
    useLegacySql: false
  };

  const queryResults = BigQuery.Jobs.query(request, projectId);
  const jobId = queryResults.jobReference.jobId;

  // Wait for the query to complete with a timeout of 60 seconds
  let status;
  const maxRetries = 60;
  let retries = 0;
  do {
    Utilities.sleep(1000);
    status = BigQuery.Jobs.get(projectId, jobId);
    retries++;
  } while (status.status.state !== 'DONE' && retries < maxRetries);

  if (status.status.state !== 'DONE') {
    throw new Error("Query did not complete within the timeout period.");
  }

  return BigQuery.Jobs.getQueryResults(projectId, jobId);
}


function populateSheet(trnDate, trnTime) {
  const query = `
  SELECT machine_name, current_amount, machine_status, bill_validator_health 
  FROM \`ms-paybox-prod-1.pldtsmart.monitoring_hourly\`
  WHERE created_at >= '${trnDate} ${trnTime.substring(0, 2)}:00:00' 
  AND created_at <= '${trnDate} ${trnTime.substring(0, 2)}:59:59';
`;


  Logger.log(query);

  const queryResults = executeQueryAndWait(query);
  const rows = queryResults.rows;

  if (!rows || rows.length === 0) {
    throw new Error("No data found");
  }

  return rows;
}

function getStores() {
  const query = `
    SELECT machine_name, CONCAT(servicing_bank, ' - ', dpu_frequency) AS servicing_bank
FROM \`ms-paybox-prod-1.pldtsmart.machines\`
WHERE status = TRUE
  `;

  Logger.log(query);

  const queryResults = executeQueryAndWait(query);
  const rows = queryResults.rows;

  if (!rows || rows.length === 0) {
    throw new Error("No data found");
  }

  return rows;
}

function formatDateToYYYYMMDD(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDate(dateString) {
  let date;

  // Check if the dateString is a Unix timestamp
  if (/^\d+(\.\d+)?$/.test(dateString)) {
    date = new Date(Number(dateString) * 1000); // Convert Unix timestamp to milliseconds
  } else {
    date = new Date(dateString);
  }

  if (isNaN(date)) {
    // If the date is invalid, handle the error or return an empty string
    return '';
  }

  const year = date.getFullYear();
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);
  return `${year}-${month}-${day}`;
}
