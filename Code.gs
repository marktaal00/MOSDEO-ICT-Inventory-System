// =========================================================================
// 1. WEB APP INITIALIZATION
// =========================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ICT Equipment Inventory')
      .setFaviconUrl('https://drive.google.com/thumbnail?id=1eWolVlTlkM3Ac7pUHOHxUGE3WfI-p8wW&sz=w64#.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =========================================================================
// 2. FETCHING DATA FROM GOOGLE SHEETS
// =========================================================================

function getInventoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  let sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || allSheets[0];

  const data = sheet.getDataRange().getValues();
  const backgrounds = sheet.getDataRange().getBackgrounds();
  
  if (data.length <= 1) return [];

  const rows = data.slice(1); 
  
  const safeDate = (val) => (val instanceof Date ? val.toLocaleDateString() : val);

  return rows.map((row, index) => {
    const rowBgColor = backgrounds[index + 1][0];
    return {
      rowIndex: index + 2, 
      rowColor: rowBgColor,
      type: row[1],
      acquisitionType: row[2],
      processor: row[3],
      memory: row[4],
      disk: row[5],
      os: row[6],
      officeInstalled: row[7],
      otherSoftware: row[8],
      status: row[9],
      par: row[10],
      serial: row[11],
      propertyNo: row[12],
      desc: row[13],
      model: row[14],
      brand: row[15],
      cost: row[16],
      user: row[17],
      designation: row[18],
      section: row[19],
      district: row[20],
      office: row[21],
      assetOwner: row[22],
      dateReceived: safeDate(row[23]),
      receivedFrom: row[24],
      supplier: row[25],
      dateAcquired: safeDate(row[26]),
      remarks: row[27],
      ictName: row[28],
      officeLic: row[29],
      fromComputerSerial: row[30],
      warrantyEndDate: safeDate(row[33]) // col AH [33] — Warranty End Date
    };
  });
}

// =========================================================================
// 3. LOGGING & ACTIVITY
// =========================================================================

function getActivityLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("ActivityLog");
  
  if (!logSheet) return [];
  
  const data = logSheet.getDataRange().getValues();
  return data.slice(1).reverse().map(row => ({
    timestamp: row[0] instanceof Date ? row[0].toLocaleString() : String(row[0] || ''),
    action: String(row[1] || '')
  }));
}

function logActivity(actionDetail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("ActivityLog");
  
  if (!logSheet) {
    logSheet = ss.insertSheet("ActivityLog");
    logSheet.appendRow(["Timestamp", "Activity / Action Details"]);
    logSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#f1f5f9");
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 180);
    logSheet.setColumnWidth(2, 600);
    logSheet.hideSheet(); 
  }
  
  logSheet.appendRow([new Date().toLocaleString(), actionDetail]);
}

// =========================================================================
// 4. LOGIN & AUTHENTICATION
// =========================================================================

// =========================================================================
// 4b. USER MANAGEMENT (Admin only)
// =========================================================================

// Returns all users from the Users sheet (passwords masked with ***).
// Never sends raw passwords to the frontend.
function getUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return [];
  const data = userSheet.getDataRange().getValues();
  // Skip header row, return username + role only (never the real password)
  return data.slice(1).map((row, i) => ({
    rowIndex: i + 2, // 1-based sheet row (row 1 = header)
    username: row[0],
    role:     row[2]
  }));
}

// Adds a new user. Returns 'exists' if the username is already taken.
function addUser(username, password, role) {
  if (!username || !password || !role) throw new Error('Username, password, and role are required.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['Username', 'Password', 'Role']);
  }
  const data = userSheet.getDataRange().getValues();
  // Check for duplicate username (case-insensitive)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username.toLowerCase()) return 'exists';
  }
  userSheet.appendRow([username.trim(), password, role]);
  logActivity(`USER ADDED: Account '${username}' (${role}) created by Admin.`);
  return 'success';
}

// Updates username and/or role of an existing user (by sheet row).
// Does NOT change the password — use resetPassword for that.
function updateUser(rowIndex, newUsername, newRole) {
  if (!rowIndex) throw new Error('Missing rowIndex.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) throw new Error('Users sheet not found.');
  const oldUsername = userSheet.getRange(rowIndex, 1).getValue();
  userSheet.getRange(rowIndex, 1).setValue(newUsername.trim());
  userSheet.getRange(rowIndex, 3).setValue(newRole);
  logActivity(`USER UPDATED: Account '${oldUsername}' renamed to '${newUsername}' / role set to ${newRole}.`);
  return 'success';
}

// Resets a user's password to the new value provided.
function resetPassword(rowIndex, newPassword) {
  if (!rowIndex || !newPassword) throw new Error('Missing rowIndex or new password.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) throw new Error('Users sheet not found.');
  const username = userSheet.getRange(rowIndex, 1).getValue();
  userSheet.getRange(rowIndex, 2).setValue(newPassword);
  logActivity(`PASSWORD RESET: Password changed for account '${username}' by Admin.`);
  return 'success';
}

// Deletes a user row. Prevents deleting the last remaining Admin account.
function deleteUser(rowIndex) {
  if (!rowIndex) throw new Error('Missing rowIndex.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) throw new Error('Users sheet not found.');
  const data = userSheet.getDataRange().getValues();
  const username = userSheet.getRange(rowIndex, 1).getValue();
  const role     = userSheet.getRange(rowIndex, 3).getValue();
  // Count remaining Admins — never delete the last one
  const adminCount = data.slice(1).filter(row => String(row[2]).toLowerCase() === 'admin').length;
  if (adminCount <= 1 && role === 'Admin') {
    return 'last_admin'; // signal to frontend to show a friendly error
  }
  userSheet.deleteRow(rowIndex);
  logActivity(`USER DELETED: Account '${username}' removed by Admin.`);
  return 'success';
}

// =========================================================================
// 4c. LOGIN (original)
// =========================================================================

function checkLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");

  // No hardcoded fallback account here on purpose — see setupInitialAdmin()
  // below, which you run ONCE manually from the Apps Script editor to create
  // the very first Admin account with a password YOU choose.
  if (!userSheet) {
    throw new Error('Users sheet not found. Run setupInitialAdmin() once from the Apps Script editor first.');
  }

  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      logActivity(`LOGIN: User '${username}' signed in as ${data[i][2]}.`);
      return { success: true, role: data[i][2] };
    }
  }
  
  logActivity(`FAILED LOGIN: Failed attempt for user '${username}'.`);
  return { success: false };
}

// ── ONE-TIME SETUP — run this yourself, once, from the Apps Script editor ──
// Select this function in the toolbar dropdown and click Run. It creates the
// Users sheet and your first Admin account, using the username/password you
// type in right here (change these two values before running!).
// This intentionally does NOT run automatically on login, so no default
// credentials ever ship in the code itself — safe to make this repo public.
function setupInitialAdmin() {
  const USERNAME = 'CHANGE_ME';       // <-- set your desired admin username
  const PASSWORD = 'CHANGE_ME_TOO';   // <-- set your desired admin password

  if (USERNAME === 'CHANGE_ME' || PASSWORD === 'CHANGE_ME_TOO') {
    throw new Error('Edit USERNAME and PASSWORD in setupInitialAdmin() before running this.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userSheet = ss.getSheetByName("Users");

  if (!userSheet) {
    userSheet = ss.insertSheet("Users");
    userSheet.appendRow(["Username", "Password", "Role"]);
  }

  userSheet.appendRow([USERNAME, PASSWORD, "Admin"]);
  Logger.log(`Admin account '${USERNAME}' created. You can now log in with it.`);
}

// =========================================================================
// 5. DATA MANAGEMENT FUNCTIONS
// =========================================================================

// Helper: maps an item object to the sheet row array.
// Column A (index 0) = auto row number written by this function.
// Columns B-AE (indices 1-30) = the 30 inventory fields.
function itemToRowArray(item, rowNumber) {
  return [
    rowNumber,              // col A  — row ID / sequence number
    item.type,              // col B  [1]
    item.acquisitionType,   // col C  [2]
    item.processor,         // col D  [3]
    item.memory,            // col E  [4]
    item.disk,              // col F  [5]
    item.os,                // col G  [6]
    item.officeInstalled,   // col H  [7]
    item.otherSoftware,     // col I  [8]
    item.status,            // col J  [9]
    item.par,               // col K  [10]
    item.serial,            // col L  [11]
    item.propertyNo,        // col M  [12]
    item.desc,              // col N  [13]
    item.model,             // col O  [14]
    item.brand,             // col P  [15]
    item.cost,              // col Q  [16]
    item.user,              // col R  [17]
    item.designation,       // col S  [18]
    item.section,           // col T  [19]
    item.district,          // col U  [20]
    item.office,            // col V  [21]
    item.assetOwner,        // col W  [22]
    item.dateReceived,      // col X  [23]
    item.receivedFrom,      // col Y  [24]
    item.supplier,          // col Z  [25]
    item.dateAcquired,      // col AA [26]
    item.remarks,           // col AB [27]
    item.ictName,           // col AC [28]
    item.officeLic,         // col AD [29]
    item.fromComputerSerial, // col AE [30]
    item._preserveAF || '',    // col AF [31] — preserved as-is; not managed by this app
    item._preserveAG || '',    // col AG [32] — preserved as-is; not managed by this app
    item.warrantyEndDate || '' // col AH [33] — Warranty End Date
  ];
}

function addInventoryItem(item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  const lastRow = sheet.getLastRow();
  const newRowIndex = lastRow + 1;

  let colAValue;
  if (item.colAValue !== undefined && item.colAValue !== null && item.colAValue !== '') {
    colAValue = item.colAValue;
  } else {
    const lastColA = sheet.getRange(lastRow, 1).getValue();
    const lastNumber = parseInt(lastColA, 10);
    colAValue = isNaN(lastNumber) ? 1 : lastNumber + 1;
  }

  // Read AF and AG from the last data row so we know how many columns the sheet uses,
  // and to safely pass empty strings (new rows have no prior AF/AG data to preserve).
  item._preserveAF = '';
  item._preserveAG = '';
  const rowData = itemToRowArray(item, colAValue);
  sheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);

  logActivity(`ADDED: New ${item.type || 'item'} — Col A: ${colAValue}, ICT Name: ${item.ictName || 'N/A'}, Serial: ${item.serial || 'N/A'}.`);
  return `Success:${colAValue}`;
}

// IMPLEMENTED: Updates an existing row identified by item.rowIndex.
// After saving the target row it also:
//   1. ALWAYS syncs 9 shared fields to all rows with the same ICT Name prefix.
//   2. Optionally cascades PAR/Property # when cascadeParProp === true.
function updateInventoryItem(item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  if (!item.rowIndex) throw new Error("Missing rowIndex — cannot update without knowing which row to target.");

  // BUG FIX: Read the ORIGINAL col A value from the sheet before overwriting.
  // rowIndex is the sheet row number (e.g. 1126), NOT the inventory sequence number.
  // The old code wrote (rowIndex - 1) into col A, which corrupted sequence numbers
  // like 348 into 1125. We now read whatever is already in col A and preserve it.
  //
  // TRANSFER HISTORY: we also read the original End User (col R) and Section (col T)
  // here, BEFORE the row gets overwritten, so we can compare old vs new values further
  // down and log a transfer entry if either one changed.
  const originalRowValues = sheet.getRange(item.rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const originalColA = originalRowValues[0];
  const originalUser = String(originalRowValues[17] || '').trim();    // col R - End User
  const originalSection = String(originalRowValues[19] || '').trim(); // col T - Section
  // Preserve whatever is in AF and AG — this app does not manage those columns.
  item._preserveAF = originalRowValues[30] !== undefined ? originalRowValues[30] : '';
  item._preserveAG = originalRowValues[31] !== undefined ? originalRowValues[31] : '';

  // Write all 31 columns (A through AE) preserving the original col A value.
  const rowData = itemToRowArray(item, originalColA);
  sheet.getRange(item.rowIndex, 1, 1, rowData.length).setValues([rowData]);

  logActivity(`UPDATED: Row ${item.rowIndex} — ICT Name: ${item.ictName || 'N/A'}.`);

  // TRANSFER HISTORY: compare old vs new End User / Section. If either changed,
  // record a chain-of-custody entry. This catches both standalone edits AND
  // edits made through the modal's normal save flow — no special UI action needed.
  logTransferIfChanged(item, originalUser, originalSection);

  // -------------------------------------------------------------------------
  // FAST BATCH SYNC
  //
  // Performance fix: old code called setValue() once per cell per linked row
  // — up to 36 individual Sheets API calls, each a network round-trip.
  //
  // New approach:
  //   1. Read allData ONCE (one API call).
  //   2. Find all linked rows in a single pass.
  //   3. Clone each linked row, patch only the needed columns, write the
  //      entire row back in ONE setValues() call.
  //   4. PAR/Property cascade patches in the same single write.
  //
  // Result: 1 read + N writes instead of 1 read + 9N writes.
  // Desktop with 4 peripherals: ~4 API calls instead of ~36.
  // -------------------------------------------------------------------------
  if (item.ictName && (item.skipSharedSync !== true || item.cascadeParProp === true)) {
    var basePrefix = _getIctNamePrefix(item.ictName);
    var allData = sheet.getDataRange().getValues(); // single read

    // Find all linked rows in one pass.
    var linkedRows = [];
    for (var i = 1; i < allData.length; i++) {
      var sheetRowIndex = i + 1;
      if (sheetRowIndex === item.rowIndex) continue;
      var rowIctName = String(allData[i][28] || '');
      if (rowIctName && _getIctNamePrefix(rowIctName) === basePrefix) {
        linkedRows.push({ sheetRow: sheetRowIndex, dataIndex: i });
      }
    }

    if (linkedRows.length > 0) {
      // Map of 0-based column index -> new value.
      // Shared field indices (0-based): J=9 R=17 S=18 T=19 W=22 X=23 Z=25 AA=26 AD=29
      var syncFields = {};
      if (item.skipSharedSync !== true) {
        syncFields[9]  = item.status       || '';  // col J  - Status
        syncFields[13] = item.desc         || '';  // col N  - Description
        syncFields[17] = item.user         || '';  // col R  - End User
        syncFields[18] = item.designation  || '';  // col S  - Designation
        syncFields[19] = item.section      || '';  // col T  - Section
        syncFields[22] = item.assetOwner   || '';  // col W  - Asset Owner
        syncFields[23] = item.dateReceived || '';  // col X  - Date Received
        syncFields[25] = item.supplier     || '';  // col Z  - Supplier
        syncFields[26] = item.dateAcquired || '';  // col AA - Date Acquired
        syncFields[29] = item.officeLic    || '';  // col AD - MS Office Licence
        syncFields[27] = item.remarks      || '';  // col AB - Remarks
      }
      // PAR/Property cascade — patched in the same write (col K=10, col M=12, 0-based).
      if (item.cascadeParProp === true) {
        syncFields[10] = item.par        || '';  // col K - PAR/ICS #
        syncFields[12] = item.propertyNo || '';  // col M - Property #
      }

      var colIdxList = Object.keys(syncFields);

      // One setValues() per linked row instead of one setValue() per cell.
      linkedRows.forEach(function(link) {
        var updatedRow = allData[link.dataIndex].slice(); // clone to avoid mutating allData
        colIdxList.forEach(function(idx) { updatedRow[idx] = syncFields[idx]; });
        sheet.getRange(link.sheetRow, 1, 1, updatedRow.length).setValues([updatedRow]);
      });

      var logParts = [];
      if (item.skipSharedSync !== true) logParts.push('shared fields (Description/User/Designation/Section/AssetOwner/Status/DateReceived/Supplier/DateAcquired/OfficeLic)');
      if (item.cascadeParProp === true)  logParts.push('PAR/Property #');
      logActivity(
        'SYNC: Updated ' + logParts.join(' + ') + ' on ' + linkedRows.length + ' linked row(s) ' +
        "under ICT prefix '" + basePrefix + "' (rows: " + linkedRows.map(function(l){ return l.sheetRow; }).join(', ') + ').'
      );
    }
  }

  return "Success";
}

// =========================================================================
// 6B. TRANSFER HISTORY (Chain of Custody)
// =========================================================================

// Compares the OLD End User / Section (read before the row was overwritten)
// against the NEW values on `item`. If either changed, appends an entry to
// the hidden "TransferHistory" sheet so there's a full audit trail of every
// unit's custody chain over time.
function logTransferIfChanged(item, originalUser, originalSection) {
  const newUser = String(item.user || '').trim();
  const newSection = String(item.section || '').trim();

  const userChanged = originalUser !== newUser;
  const sectionChanged = originalSection !== newSection;

  if (!userChanged && !sectionChanged) return; // nothing to log

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("TransferHistory");

  if (!logSheet) {
    logSheet = ss.insertSheet("TransferHistory");
    logSheet.appendRow([
      "Timestamp", "Row", "ICT Name", "Serial #", "Property #",
      "Field Changed", "From", "To", "Changed By"
    ]);
    logSheet.getRange("A1:I1").setFontWeight("bold").setBackground("#f1f5f9");
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidths(1, 9, 150);
    logSheet.hideSheet();
  }

  // Whoever is logged in client-side gets passed through item.changedBy from the
  // frontend (set from currentUserRole/username at save time). Falls back to
  // the script's effective user email if not provided.
  const changedBy = item.changedBy || Session.getEffectiveUser().getEmail() || 'Unknown';
  // Always write timestamp as a locale string so it round-trips through
  // google.script.run cleanly — Date objects cannot be serialized.
  const timestamp = new Date().toLocaleString();
  const rowsToAppend = [];

  if (userChanged) {
    rowsToAppend.push([
      timestamp, item.rowIndex, item.ictName || '', item.serial || '', item.propertyNo || '',
      'End User', originalUser || '(blank)', newUser || '(blank)', changedBy
    ]);
  }
  if (sectionChanged) {
    rowsToAppend.push([
      timestamp, item.rowIndex, item.ictName || '', item.serial || '', item.propertyNo || '',
      'Section', originalSection || '(blank)', newSection || '(blank)', changedBy
    ]);
  }

  // Append all rows in one call instead of one appendRow() per change.
  if (rowsToAppend.length > 0) {
    const startRow = logSheet.getLastRow() + 1;
    logSheet.getRange(startRow, 1, rowsToAppend.length, 9).setValues(rowsToAppend);
  }

  const summary = [];
  if (userChanged) summary.push(`End User: '${originalUser || '(blank)'}' -> '${newUser || '(blank)'}'`);
  if (sectionChanged) summary.push(`Section: '${originalSection || '(blank)'}' -> '${newSection || '(blank)'}'`);
  logActivity(`TRANSFER: Row ${item.rowIndex} (${item.ictName || 'N/A'}) — ${summary.join(', ')}.`);
}

// Returns the full transfer history, newest first, for ALL units. Used by the
// frontend's "Transfer History" admin view.
function getAllTransferHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("TransferHistory");

  if (!logSheet) return [];

  const data = logSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).reverse().map(row => ({
    timestamp: row[0] instanceof Date ? row[0].toLocaleString() : String(row[0] || ''),
    rowIndex: row[1],
    ictName: row[2],
    serial: String(row[3] || ''),
    propertyNo: String(row[4] || ''),
    fieldChanged: row[5],
    from: row[6],
    to: row[7],
    changedBy: row[8]
  }));
}

// Returns the transfer history for a SINGLE unit, identified by ICT Name.
// Used by the modal's "Transfer History" tab so a user can see one unit's
// full chain of custody without scrolling through every unit's history.
function getTransferHistoryForUnit(ictName) {
  if (!ictName) return [];
  const all = getAllTransferHistory();
  return all.filter(entry => entry.ictName === ictName);
}


// and its own peripherals match each other — but NOT other computers.
//
// Examples:
//   "R10-MISOR2-045"     -> prefix "R10-MISOR2-045"  (Desktop itself)
//   "R10-MISOR2-045-KB"  -> prefix "R10-MISOR2-045"  (matches Desktop above)
//   "R10-MISOR2-046"     -> prefix "R10-MISOR2-046"  (different computer, no match)
//   "MISOR2-PRT-003"     -> prefix "MISOR2-PRT-003"  (standalone, matches only itself)
//
// Rule: take the first 3 hyphen-separated segments.
// The old code took only 2 segments ("R10-MISOR2"), which matched EVERY
// computer in the inventory and caused all rows to update together.
function _getIctNamePrefix(ictName) {
  const parts = ictName.trim().split('-');
  return parts.length >= 3 ? parts.slice(0, 3).join('-') : ictName.trim();
}

// FIX: Uses sheet.getLastColumn() instead of hardcoded 31.
// This prevents under-coloring if a new column is ever added to the sheet.
function colorRowInSheet(rowIndex, hexColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || ss.getSheets()[0];
  const lastCol = sheet.getLastColumn();
  sheet.getRange(rowIndex, 1, 1, lastCol).setBackground(hexColor);
  
  logActivity(`COLOR UPDATE: Row ${rowIndex} painted to ${hexColor}.`);
  return "Success";
}

// =========================================================================
// 7. PER-UNIT PAR REPORT
// =========================================================================

// Generates a Property Acknowledgement Receipt (PAR) PDF for a SINGLE
// inventory row, formatted like the standard government PAR form.
//
// Approach: builds a temporary Google Doc from a template-style layout,
// converts it to PDF, saves to Drive, and returns the download link.
// A Doc (not the Sheet) is used here because a PAR has a fixed letter-size
// layout with signature lines — very different from a spreadsheet export.
function exportSinglePAR(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  if (!rowIndex) throw new Error("Missing rowIndex — cannot generate PAR without knowing which row.");

  const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Map row values to named fields using the same column order as itemToRowArray.
  const item = {
    colA: rowValues[0],
    type: rowValues[1],
    acquisitionType: rowValues[2],
    processor: rowValues[3],
    memory: rowValues[4],
    disk: rowValues[5],
    os: rowValues[6],
    officeInstalled: rowValues[7],
    otherSoftware: rowValues[8],
    status: rowValues[9],
    par: rowValues[10],
    serial: rowValues[11],
    propertyNo: rowValues[12],
    desc: rowValues[13],
    model: rowValues[14],
    brand: rowValues[15],
    cost: rowValues[16],
    user: rowValues[17],
    designation: rowValues[18],
    section: rowValues[19],
    district: rowValues[20],
    office: rowValues[21],
    assetOwner: rowValues[22],
    dateReceived: rowValues[23] instanceof Date ? rowValues[23].toLocaleDateString() : rowValues[23],
    receivedFrom: rowValues[24],
    supplier: rowValues[25],
    dateAcquired: rowValues[26] instanceof Date ? rowValues[26].toLocaleDateString() : rowValues[26],
    remarks: rowValues[27],
    ictName: rowValues[28],
    officeLic: rowValues[29],
    fromComputerSerial: rowValues[30],
    warrantyEndDate: safeDate(rowValues[33]) // col AH
  };

  // Build a new Google Doc with the PAR layout.
  const docName = 'PAR_' + (item.ictName || 'Unit_' + item.colA) + '_' + new Date().toISOString().slice(0, 10);
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);

  // --- Header ---
  const title = body.appendParagraph('PROPERTY ACKNOWLEDGEMENT RECEIPT');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const subtitle = body.appendParagraph(item.office || 'DPWH');
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subtitle.editAsText().setBold(true);

  body.appendParagraph(''); // spacer

  // --- PAR / Property Numbers ---
  const refTable = body.appendTable([
    ['PAR/ICS No.:', String(item.par || ''), 'Property No.:', String(item.propertyNo || '')],
    ['Date Acquired:', String(item.dateAcquired || ''), 'Date Received:', String(item.dateReceived || '')]
  ]);
  refTable.setBorderWidth(0);

  body.appendParagraph('');

  // --- Equipment Details Table ---
  const detailsHeader = body.appendParagraph('EQUIPMENT DETAILS');
  detailsHeader.editAsText().setBold(true);

  const detailsRows = [
    ['ICT Name', String(item.ictName || '')],
    ['Type', String(item.type || '')],
    ['Description', String(item.desc || '')],
    ['Brand / Model', (item.brand || '') + ' / ' + (item.model || '')],
    ['Serial No.', String(item.serial || '')],
    ['Processor', String(item.processor || '')],
    ['Memory', String(item.memory || '')],
    ['Disk', String(item.disk || '')],
    ['Operating System', String(item.os || '')],
    ['Unit Cost', String(item.cost || '')],
    ['Status', String(item.status || '')]
  ];
  const detailsTable = body.appendTable(detailsRows);
  detailsTable.setBorderWidth(1);
  // Bold the left column (labels)
  for (let r = 0; r < detailsTable.getNumRows(); r++) {
    detailsTable.getRow(r).getCell(0).editAsText().setBold(true);
  }

  body.appendParagraph('');

  // --- Accountability / Assignment ---
  const acctHeader = body.appendParagraph('ACCOUNTABILITY');
  acctHeader.editAsText().setBold(true);

  const acctRows = [
    ['End User', String(item.user || '')],
    ['Designation', String(item.designation || '')],
    ['Section', String(item.section || '')],
    ['District', String(item.district || '')],
    ['Asset Owner / Accountable Person', String(item.assetOwner || '')]
  ];
  const acctTable = body.appendTable(acctRows);
  acctTable.setBorderWidth(1);
  for (let r = 0; r < acctTable.getNumRows(); r++) {
    acctTable.getRow(r).getCell(0).editAsText().setBold(true);
  }

  body.appendParagraph('');
  if (item.remarks) {
    const remarksPara = body.appendParagraph('Remarks: ' + item.remarks);
    remarksPara.editAsText().setItalic(true);
  }

  body.appendParagraph('');
  body.appendParagraph('');

  // --- Signature Block ---
  const sigTable = body.appendTable([
    ['Received by:', 'Issued by:'],
    ['', ''],
    ['', ''],
    ['_______________________________', '_______________________________'],
    [String(item.user || '(End User Name)'), '(Authorized ICT Personnel)'],
    ['Signature over Printed Name', 'Signature over Printed Name'],
    ['Date: ______________________', 'Date: ______________________']
  ]);
  sigTable.setBorderWidth(0);

  doc.saveAndClose();

  // Convert the Doc to PDF.
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs('application/pdf').setName(docName + '.pdf');

  // Save PDF in the same folder as the spreadsheet (or root), then delete the temp Doc.
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Clean up — remove the intermediate Google Doc, we only need the PDF.
  docFile.setTrashed(true);

  logActivity(`PAR GENERATED: ${pdfFile.getName()} for row ${rowIndex} (ICT Name: ${item.ictName || 'N/A'}).`);

  return {
    success: true,
    url: pdfFile.getUrl(),
    fileName: pdfFile.getName()
  };
}

// =========================================================================
// 8. FILTERED EXPORT (Excel/CSV)
// =========================================================================

// Exports a FILTERED subset of the inventory (rows matching whatever the
// frontend currently has filtered/searched) as an .xlsx file.
// rowIndexes: array of sheet row numbers to include (sent from frontend,
// based on whatever rows are currently visible in the filtered table).
function exportFilteredToExcel(rowIndexes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  if (!rowIndexes || rowIndexes.length === 0) {
    throw new Error("No rows to export — the current filter returned zero results.");
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Create a temporary spreadsheet containing only the header + filtered rows.
  const tempName = 'ICT_Inventory_Filtered_Export_' + new Date().toISOString().slice(0, 10);
  const tempSS = SpreadsheetApp.create(tempName);
  const tempSheet = tempSS.getSheets()[0];

  tempSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  tempSheet.getRange('1:1').setFontWeight('bold').setBackground('#f1f5f9');
  tempSheet.setFrozenRows(1);

  // Pull each requested row's full values and write them into the temp sheet.
  const dataRows = rowIndexes.map(function(rIdx) {
    return sheet.getRange(rIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
  });
  tempSheet.getRange(2, 1, dataRows.length, headerRow.length).setValues(dataRows);
  tempSheet.autoResizeColumns(1, headerRow.length);

  // Convert the temp spreadsheet to .xlsx via the export endpoint.
  const tempSSId = tempSS.getId();
  const url = 'https://docs.google.com/spreadsheets/d/' + tempSSId
    + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  const xlsxBlob = response.getBlob().setName(tempName + '.xlsx');

  // Save the .xlsx into the same folder as the main spreadsheet.
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const xlsxFile = folder.createFile(xlsxBlob);
  xlsxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // The temporary Google Sheet is no longer needed once converted — trash it.
  DriveApp.getFileById(tempSSId).setTrashed(true);

  logActivity(`EXPORTED: Filtered Excel report generated (${xlsxFile.getName()}, ${rowIndexes.length} row(s)).`);

  return {
    success: true,
    url: xlsxFile.getUrl(),
    fileName: xlsxFile.getName(),
    rowCount: rowIndexes.length
  };
}

// =========================================================================
// 9. DASHBOARD SUMMARY COUNTS
// =========================================================================

// Returns aggregate counts of inventory items grouped by Status, and also
// broken down by Office/Section, so the frontend can render summary cards
// without having to recompute this from the full dataset every time.
function getDashboardSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { totalUnits: 0, statusCounts: {}, officeCounts: {} };
  }

  const rows = data.slice(1);
  const statusCounts = {};
  const officeCounts = {};

  rows.forEach(function(row) {
    const status = String(row[9] || 'Unspecified').trim();
    const office = String(row[21] || 'Unspecified').trim();

    statusCounts[status] = (statusCounts[status] || 0) + 1;
    if (!officeCounts[office]) {
      officeCounts[office] = { total: 0, Operational: 0, 'For Repair': 0, Decommission: 0 };
    }
    officeCounts[office].total++;

    // Bucket common status variants — adjust these match strings if your
    // sheet uses different exact wording for status values.
    const statusLower = status.toLowerCase();
    if (statusLower.includes('operational')) officeCounts[office].Operational++;
    else if (statusLower.includes('repair')) officeCounts[office]['For Repair']++;
    else if (statusLower.includes('decommission')) officeCounts[office].Decommission++;
  });

  return {
    totalUnits: rows.length,
    statusCounts: statusCounts,
    officeCounts: officeCounts
  };
}

// Exports the ENTIRE inventory sheet (all rows, all 31 columns) as a
// landscape PDF, saves it to Google Drive, and returns a download link.
//
// How it works:
//   1. Uses Sheets' native export-to-PDF endpoint via UrlFetchApp — this
//      renders the actual sheet (same data, same columns) directly to PDF
//      without rebuilding the table from scratch.
//   2. Landscape + small font + fit-to-width so all 31 columns fit on
//      each printed page instead of being cut off.
//   3. Saves the PDF into the same Drive folder as the spreadsheet (or
//      Drive root if the file isn't in a folder), named with today's date.
//   4. Returns the file's direct download URL so the frontend can open it
//      in a new tab — the browser handles the actual download/print.
function exportInventoryToPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();

  // Build the export URL with landscape + fit-to-width + all columns options.
  // Reference: Google Sheets supports these query params on the export endpoint.
  const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export'
    + '?format=pdf'
    + '&gid=' + sheetId
    + '&size=A3'           // A3 gives more horizontal room for 31 columns than A4/Letter
    + '&portrait=false'    // landscape orientation
    + '&fitw=true'         // fit columns to page width
    + '&top_margin=0.3'
    + '&bottom_margin=0.3'
    + '&left_margin=0.3'
    + '&right_margin=0.3'
    + '&sheetnames=false'
    + '&printtitle=false'
    + '&pagenumbers=true'
    + '&gridlines=true'
    + '&fzr=true';         // repeat frozen header row on every printed page

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  const pdfBlob = response.getBlob().setName(
    'ICT_Inventory_Full_Report_' + new Date().toISOString().slice(0, 10) + '.pdf'
  );

  // Save into the same folder as the spreadsheet so it's easy to find later.
  const ssFile = DriveApp.getFileById(ssId);
  const parents = ssFile.getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const pdfFile = folder.createFile(pdfBlob);

  // Make it viewable by anyone with the link so the frontend can open it
  // directly without requiring a separate Drive permission prompt.
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  logActivity(`EXPORTED: Full inventory PDF report generated (${pdfFile.getName()}).`);

  return {
    success: true,
    url: pdfFile.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId(),
    fileName: pdfFile.getName()
  };
}
// =========================================================================
// BULK FIND & REPLACE
// =========================================================================

// Updates a specific field to replaceValue on all rows identified by rowIndexes.
// rowIndexes is an array of sheet row numbers (1-based) from the frontend.
// field is the item property key (e.g. "assetOwner", "user", "section").
// Returns count of rows updated.
function bulkFindReplace(field, rowIndexes, replaceValue) {
  if (!rowIndexes || rowIndexes.length === 0) return 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mis Or 2nd ICT inventory") || ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  // Map field key to sheet column number (1-based, col A = 1).
  // Array index in itemToRowArray + 1 = sheet column number
  // e.g. type is at array index 1 → sheet col B = 2
  const FIELD_COL = {
    type:         2,  // col B  [array 1]
    status:       10, // col J  [array 9]
    user:         18, // col R  [array 17]
    designation:  19, // col S  [array 18]
    section:      20, // col T  [array 19]
    district:     21, // col U  [array 20]
    office:       22, // col V  [array 21]
    assetOwner:   23, // col W  [array 22]
    supplier:     26, // col Z  [array 25]
    remarks:      28, // col AB [array 27]
    officeLic:    30, // col AD [array 29]
  };

  const colIndex = FIELD_COL[field];
  if (!colIndex) throw new Error('Field "' + field + '" is not supported for bulk replace.');

  // colIndex is already the 1-based sheet column number — use directly.
  rowIndexes.forEach(function(rowIndex) {
    sheet.getRange(rowIndex, colIndex).setValue(replaceValue);
  });

  logActivity(
    'BULK REPLACE: Field "' + field + '" set to "' + replaceValue + '" on ' +
    rowIndexes.length + ' row(s): [' + rowIndexes.join(', ') + '].'
  );

  return rowIndexes.length;
}