// ============================================================
//  T-29 Details Dashboard — Google Apps Script Web App
// ============================================================

const SOURCE_SPREADSHEET_ID = '1yvC466_OqOeUotT8okfbKFE3QR028VzxCmsZL5GLGR8';
const SECRET_TOKEN = 'MASAI@2019';

function doGet(e) {
  const params = e.parameter || {};

  if (params.token !== SECRET_TOKEN) {
    return jsonOut({ error: 'Unauthorized' });
  }

  const sheetName = params.sheet || 'Master Data';
  const action    = params.action || 'index';

  try {
    const ss    = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return jsonOut({ error: 'Sheet not found: ' + sheetName });

    if (action === 'index') return getIndex(sheet);
    if (action === 'row')   return getRow(sheet, parseInt(params.row));
    if (action === 'rows')  return getRows(sheet, parseInt(params.start), parseInt(params.end));

    return jsonOut({ error: 'Invalid action' });
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

// ─────────────────────────────────────────────────────────────
//  1. Fast Index Fetch (Single Bulk Data Pull)
// ─────────────────────────────────────────────────────────────
function getIndex(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return jsonOut({ columns: {} });

  // 1. Fetch headers to map columns
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Helper to safely match headers regardless of case/spaces
  const norm = (str) => (str || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '');

  const columnsToFetch = [
    'emailid', 'studentname', 'mobile', 'alternatenumber', 'coursename',
    'paymentmode', 'overallamountpaid', 'nbfcstatus', 'commonname', 'batchstartdate'
  ];

  const colMap = {};
  let maxColIndex = 0;
  headers.forEach((h, i) => {
    const normalizedHeader = norm(h);
    if (columnsToFetch.includes(normalizedHeader)) {
       colMap[normalizedHeader] = i; // 0-indexed column position
       if (i > maxColIndex) maxColIndex = i;
    }
  });

  // 2. ONE fast backend call to get all required data
  // Limit the fetching to only the maximum required column (1-indexed) to avoid loading heavy unneeded columns
  const colsToFetch = maxColIndex + 1;
  const allData = sheet.getRange(2, 1, lastRow - 1, colsToFetch).getValues();
  
  const columns = {
    email: [], name: [], mobile: [], altphone: [], course: [],
    payment: [], paid: [], nbfc: [], common: [], batch: []
  };

  const eIdx = colMap['emailid'];
  const nIdx = colMap['studentname'];
  const mIdx = colMap['mobile'];
  const aIdx = colMap['alternatenumber'];
  const cIdx = colMap['coursename'];
  const pIdx = colMap['paymentmode'];
  const paIdx = colMap['overallamountpaid'];
  const nbIdx = colMap['nbfcstatus'];
  const cmIdx = colMap['commonname'];
  const bIdx = colMap['batchstartdate'];

  // 3. Find true last row to avoid processing thousands of empty formatted rows at the bottom
  let actualLastRowIdx = allData.length - 1;
  while (actualLastRowIdx >= 0) {
    if (allData[actualLastRowIdx].some(cell => cell !== '' && cell !== null)) {
      break; // Found a row with actual data
    }
    actualLastRowIdx--;
  }
  const maxRowsToProcess = actualLastRowIdx + 1;

  // 4. Fast native V8 loop to extract needed columns
  for (let i = 0; i < maxRowsToProcess; i++) {
    const row = allData[i];
    
    // We MUST push an item for every row so the index lines up correctly with the sheet row numbers
    if (eIdx !== undefined) columns.email.push(row[eIdx] != null ? row[eIdx].toString() : '');
    if (nIdx !== undefined) columns.name.push(row[nIdx] != null ? row[nIdx].toString() : '');
    if (mIdx !== undefined) columns.mobile.push(row[mIdx] != null ? row[mIdx].toString() : '');
    if (aIdx !== undefined) columns.altphone.push(row[aIdx] != null ? row[aIdx].toString() : '');
    if (cIdx !== undefined) columns.course.push(row[cIdx] != null ? row[cIdx].toString() : '');
    if (pIdx !== undefined) columns.payment.push(row[pIdx] != null ? row[pIdx].toString() : '');
    if (paIdx !== undefined) columns.paid.push(row[paIdx] != null ? row[paIdx].toString() : '');
    if (nbIdx !== undefined) columns.nbfc.push(row[nbIdx] != null ? row[nbIdx].toString() : '');
    if (cmIdx !== undefined) columns.common.push(row[cmIdx] != null ? row[cmIdx].toString() : '');
    if (bIdx !== undefined) columns.batch.push(row[bIdx] != null ? row[bIdx].toString() : '');
  }

  return jsonOut({ headers, columns });
}

// ─────────────────────────────────────────────────────────────
//  2. Fetch a single row by index
// ─────────────────────────────────────────────────────────────
function getRow(sheet, rowIdx) {
  const lastCol = sheet.getLastColumn();
  const raw = sheet.getRange(rowIdx, 1, 1, lastCol).getValues()[0];
  const row = raw.map(v => (v !== null && v !== undefined ? v.toString() : ''));
  return jsonOut({ row });
}

// ─────────────────────────────────────────────────────────────
//  3. Fetch multiple rows (Pagination/Prefetching)
// ─────────────────────────────────────────────────────────────
function getRows(sheet, start, end) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (start < 2) start = 2;
  if (end > lastRow) end = lastRow;
  
  const count = end - start + 1;
  if (count <= 0) return jsonOut({ rows: [], startRow: start });

  const data = sheet.getRange(start, 1, count, lastCol).getValues();
  const rows = data.map(row =>
    row.map(v => (v !== null && v !== undefined ? v.toString() : ''))
  );

  return jsonOut({ rows, startRow: start });
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
