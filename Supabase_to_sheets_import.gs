// A script to use in Google Sheets that will import from a Supabase table. Can be run on a trigger to happen automatically.
// Batched sheet clear, batched data import to allow scaling beyong ~20k rows
// Written with AI tools in Aug 2025, tested on a project with 3,700 rows.

const SUPABASE_URL = 'https://abcdefghijklmnopqrst.supabase.co';
const SUPABASE_ANON_KEY = 'YOUR_KEY_HERE';
const SUPABASE_TABLE_NAME = 'your_table_01';
const GOOGLE_SHEET_NAME = 'SupabaseImport';
const CLEAR_BATCH_SIZE = 1000;
const MAX_ROWS_PER_BATCH = 1000;

/**
 * Main function
 */
function mainFunction() {
  console.log("=== STARTING MAIN FUNCTION ===");
  const start = new Date();
  console.log("Start time: " + start.toISOString());

  console.log("Step 1: Accessing or creating sheet...");
  const sheet = getOrCreateSheet(GOOGLE_SHEET_NAME);

  console.log("Step 2: Clearing sheet and verifying...");
  clearSheet(sheet);

  console.log("Step 3: Fetching data from Supabase...");
  const rawData = fetchSupabaseData();
  console.log(`Total records fetched: ${rawData.length}`);

  if (rawData.length === 0) {
    console.log("No data returned. Exiting without writing.");
    sheet.getRange("A1").setValue(`No new data - ${new Date().toISOString()}`);
    return;
  }

  console.log("Step 4: Processing data...");
  const processedData = processData(rawData);
  console.log(`Prepared ${processedData.length - 1} valid rows out of ${rawData.length}`);

  console.log("Step 5: Writing data to sheet in batches...");
  writeSheet(sheet, processedData);

  console.log("=== FINISHED MAIN FUNCTION ===");
  const end = new Date();
  console.log("End time: " + end.toISOString());
  console.log(`Duration: ${(end - start) / 1000}s`);
}

/**
 * Ensure sheet exists
 */
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    console.log(`Created new sheet: ${sheetName}`);
  }
  return sheet;
}

/**
 * Clear sheet in chunks and confirm it cleared
 */
function clearSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const batchSize = CLEAR_BATCH_SIZE; // rows per chunk

  if (lastRow === 0 || lastCol === 0) return; // nothing to clear

  for (let startRow = 1; startRow <= lastRow; startRow += batchSize) {
    const numRows = Math.min(batchSize, lastRow - startRow + 1);
    sheet.getRange(startRow, 1, numRows, lastCol).clearContent();
  }

  SpreadsheetApp.flush();
}


/**
 * Fetch data from Supabase
 */
function fetchSupabaseData() {
  const url = `${SUPABASE_URL}/rest/v1/${SUPABASE_TABLE_NAME}?select=*`;

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
    },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("Failed to fetch data from Supabase: " + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * Turn array of objects into [header, rows]
 */
function processData(records) {
  if (!records || records.length === 0) return [];
  const headers = Object.keys(records[0]);
  const rows = records.map(obj => headers.map(h => obj[h] ?? ""));
  return [headers, ...rows];
}

/**
 * Write data to sheet in batches with final check
 */
function writeSheet(sheet, data) {
  if (!data || data.length === 0) return;

  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const sheetName = sheet.getName();
  const totalRows = data.length;
  let startIdx = 0;
  let batchNumber = 1;

  while (startIdx < totalRows) {
    const endIdx = Math.min(startIdx + MAX_ROWS_PER_BATCH, totalRows);
    const batch = data.slice(startIdx, endIdx);

    Sheets.Spreadsheets.Values.update(
      { values: batch },
      ssId,
      `${sheetName}!A${startIdx + 1}`,
      { valueInputOption: "USER_ENTERED" }
    );

    console.log(`Batch #${batchNumber}: wrote ${batch.length} rows`);
    startIdx = endIdx;
    batchNumber++;
  }
}
