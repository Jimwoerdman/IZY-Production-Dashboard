/***** CONFIG ROW LOCKS *****/
const CONFIG = {
  WATCH_SHEETS: ["Workfile"],
  KEY_COLUMN_INDEX: 3,        // C
  FIXED_LAST_COLUMN: 23,      // A:T
  ALWAYS_ALLOWED: [
    "daan@izybottles.com",
    "jim@izybottles.com",
    "biessenlevi@gmail.com",
    "sharon@orderchamp.com"
  ],
  HEADER_ROWS: 1,
  TAG: "AUTOLOCK_BY_COLC_STRICT_DATE",
  DEBUG: false
};

/***** MAIN (edit trigger row locks) *****/
function onEditRowLockByC(e) {
  return; // tijdelijk uitgeschakeld

  if (!e) return;
  const sheet = e.range.getSheet();

  if (CONFIG.WATCH_SHEETS.length && !CONFIG.WATCH_SHEETS.includes(sheet.getName())) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const row = e.range.getRow();
  if (row <= CONFIG.HEADER_ROWS) return;

  const keyCell = sheet.getRange(row, CONFIG.KEY_COLUMN_INDEX);
  const shouldLock = isPastDateInCell_(keyCell);

  if (CONFIG.DEBUG) {
    SpreadsheetApp.getActive().toast(
      `Row ${row}: ${shouldLock ? "LOCK A:T" : "UNLOCK"} (C="${keyCell.getDisplayValue()}")`
    );
  }

  if (shouldLock) lockRowStrict_(sheet, row);
  else unlockRowStrict_(sheet, row);
}

/***** RETROACTIVE APPLY *****/
function applyLocksNow() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const sheets = ss.getSheets().filter(sh =>
    CONFIG.WATCH_SHEETS.length ? CONFIG.WATCH_SHEETS.includes(sh.getName()) : true
  );

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROWS) return;

    const startRow = CONFIG.HEADER_ROWS + 1;
    const numRows = lastRow - CONFIG.HEADER_ROWS;
    const values = sheet.getRange(startRow, CONFIG.KEY_COLUMN_INDEX, numRows, 1).getValues();

    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      const v = values[i][0];
      const shouldLock = isPastDateValue_(v, tz);

      if (shouldLock) lockRowStrict_(sheet, row);
      else unlockRowStrict_(sheet, row);
    }
  });

  SpreadsheetApp.getActive().toast("Applied lock rule to all existing rows ✅");
}

/***** HELPERS ROW LOCKS *****/
function lockRowStrict_(sheet, row) {
  try {
    const ss = sheet.getParent();
    const lastCol = CONFIG.FIXED_LAST_COLUMN || sheet.getLastColumn();
    const rowRange = sheet.getRange(row, 1, 1, lastCol);

    const allowed = new Set(
      (CONFIG.ALWAYS_ALLOWED || []).map(e => String(e).trim().toLowerCase())
    );

    try {
      const owner = ss.getOwner();
      if (owner && owner.getEmail) {
        const ownerEmail = owner.getEmail();
        if (ownerEmail) allowed.add(ownerEmail.trim().toLowerCase());
      }
    } catch (err) {
      Logger.log("Could not read owner: " + err);
    }

    unlockRowStrict_(sheet, row);

    const p = rowRange.protect();
    p.setDescription(`${CONFIG.TAG}:${sheet.getName()}:row${row}`);
    p.setWarningOnly(false);

    try {
      p.setDomainEdit(false);
    } catch (err) {
      Logger.log("setDomainEdit failed: " + err);
    }

    try {
      const curEditors = p.getEditors();
      if (curEditors && curEditors.length) p.removeEditors(curEditors);
    } catch (err) {
      Logger.log("removeEditors failed: " + err);
    }

    const allowedArr = Array.from(allowed).filter(Boolean);
    if (allowedArr.length) {
      try {
        p.addEditors(allowedArr);
      } catch (err) {
        Logger.log("addEditors failed: " + err);
      }
    }
  } catch (err) {
    Logger.log("lockRowStrict_ failed on row " + row + ": " + err);
    throw err;
  }
}

function unlockRowStrict_(sheet, row) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const tag = `${CONFIG.TAG}:${sheet.getName()}:row${row}`;

  protections.forEach(p => {
    try {
      if (p && p.getDescription && p.getDescription() === tag) p.remove();
    } catch (_) {}
  });
}

function isPastDateInCell_(cell) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return isPastDateValue_(cell.getValue(), tz);
}

function isPastDateValue_(v, tz) {
  if (v === "" || v === null || typeof v === "undefined") return false;

  let d = v;
  if (!(d instanceof Date)) {
    const s = String(v).trim();
    if (!s) return false;
    const parsed = new Date(s);
    if (isNaN(parsed)) return false;
    d = parsed;
  }

  const todayMid = new Date(Utilities.formatDate(new Date(), tz, "yyyy-MM-dd") + " 00:00:00");
  const cellMid = new Date(Utilities.formatDate(d, tz, "yyyy-MM-dd") + " 00:00:00");
  return cellMid.getTime() < todayMid.getTime();
}

/***** UX: add menu *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Row Locks")
    .addItem("Apply to all rows now", "applyLocksNow")
    .addToUi();
}

/***** INSTALL TRIGGER ROW LOCKS *****/
function installTrigger() {
  const ssId = SpreadsheetApp.getActive().getId();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === "onEditRowLockByC") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onEditRowLockByC")
    .forSpreadsheet(ssId)
    .onEdit()
    .create();

  SpreadsheetApp.getActive().toast("Installed trigger for onEditRowLockByC ✅");
}

/***** CONFIG WORKFILE READY FLOW *****/
const WORKFILE_READY_FLOW_CONFIG = {
  WATCH_SHEETS: ["Workfile"],
  HEADER_ROWS: 1,

  STILL_TO_PRINT_COLUMN: 20,  // T
  READY_COLUMN: 23,           // W

  DEBUG: true
};

/***** MAIN (validate only when W = Yes) *****/
function onEditWorkfileReadyFlow(e) {
  if (!e) return;

  const sheet = e.range.getSheet();
  if (!WORKFILE_READY_FLOW_CONFIG.WATCH_SHEETS.includes(sheet.getName())) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row <= WORKFILE_READY_FLOW_CONFIG.HEADER_ROWS) return;
  if (col !== WORKFILE_READY_FLOW_CONFIG.READY_COLUMN) return;

  const newValue = String(e.value || "").trim().toLowerCase();

  // Alleen reageren als W op Yes wordt gezet
  if (newValue !== "yes") return;

  const stillToPrintRaw = sheet.getRange(row, WORKFILE_READY_FLOW_CONFIG.STILL_TO_PRINT_COLUMN).getValue();
  const stillToPrint = Number(stillToPrintRaw || 0);

  if (stillToPrint > 0) {
    sheet.getRange(row, WORKFILE_READY_FLOW_CONFIG.READY_COLUMN).setValue("No");

    SpreadsheetApp.getActive().toast(
      `Printing ready kan niet op Yes worden gezet. Er moeten nog ${stillToPrint} items geprint worden.`
    );
    return;
  }

  if (WORKFILE_READY_FLOW_CONFIG.DEBUG) {
    SpreadsheetApp.getActive().toast("Printing ready succesvol op Yes gezet ✅");
  }
}

/***** INSTALL TRIGGER READY FLOW *****/
function installWorkfileReadyFlowTrigger() {
  const ssId = SpreadsheetApp.getActive().getId();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === "onEditWorkfileReadyFlow") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onEditWorkfileReadyFlow")
    .forSpreadsheet(ssId)
    .onEdit()
    .create();

  SpreadsheetApp.getActive().toast("Installed trigger for onEditWorkfileReadyFlow ✅");
}

/***** CONFIG PRINT DASHBOARD *****/
// Create a folder in Google Drive for product photos, then paste its ID here:
const DRIVE_FOLDER_ID = '1mcZ2zLKtAR02jgxhLb6l3XE20108hkbi';

/***** PRINT DASHBOARD — serve sheet data to the dashboard *****/
function doGet(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const tz    = ss.getSpreadsheetTimeZone();
  const sheet = e.parameter.sheet === 'shipping'
    ? ss.getSheetByName('ShippingHistory')
    : e.parameter.sheet === 'sleeves'
      ? ss.getSheetByName('Sleeves')
      : ss.getSheetByName('Workfile');

  // Phone photo poll
  if (e.parameter.photosession) {
    const key   = e.parameter.photosession;
    const props = PropertiesService.getScriptProperties();
    const url   = props.getProperty('photo_' + key);
    if (url) props.deleteProperty('photo_' + key); // clean up after delivery
    return respondGet({ photoUrl: url || null });
  }

  if (!sheet) return respondGet({ error: 'Sheet not found', availableSheets: ss.getSheets().map(s => s.getName()) });
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1).map((row, idx) => {
    const obj = Object.fromEntries(headers.map((h, i) => {
      const v = row[i];
      if (v instanceof Date && !isNaN(v)) {
        return [h, v.getTime() === 0 ? '' : Utilities.formatDate(v, tz, 'dd/MM/yyyy')];
      }
      return [h, String(v ?? '').trim()];
    }));
    obj['_sheetRow'] = idx + 2; // 1-based row number (row 1 = headers)
    return obj;
  });

  return respondGet({ rows });
}

function respondGet(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/***** PRINT DASHBOARD — receive updates from the dashboard *****/
function doPost(e) {
  Logger.log('doPost called');
  try {
    const raw  = e.postData ? e.postData.contents : null;
    Logger.log('raw body: ' + raw);
    const data = JSON.parse(raw);
    Logger.log('action=' + data.action + ' sheetRow=' + data.sheetRow);

    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheet  = ss.getSheetByName('Workfile');
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());

    // Find columns by header name
    const col = name => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
    const priorityCol        = col('Priority');
    const soortCol           = col('Soort');
    const quantityPrintedCol = headers.findIndex(h => h.toLowerCase().includes('quantity printed'));
    const faultyCol          = col('Faulty prints');

    // Fixed column positions per user specification
    const photoCol   = 8;  // Column H
    const statusCol  = 9;  // Column I ("Status")
    const printerCol = 35; // Column AI

    if (priorityCol === -1) return respond({ error: 'Priority column not found' });

    // Phone photo upload — save to Drive, store URL in ScriptProperties for polling
    if (data.action === 'upload_photo') {
      try {
        const base64Data = data.imageBase64.includes(',') ? data.imageBase64.split(',')[1] : data.imageBase64;
        const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        const blob     = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'phone_' + data.sessionKey + '.jpg');
        const file     = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const url = 'https://drive.google.com/uc?id=' + file.getId();
        PropertiesService.getScriptProperties().setProperty('photo_' + data.sessionKey, url);
        return respond({ success: true });
      } catch (err) {
        return respond({ error: err.message });
      }
    }

    // Add new job (must be before rowIndex check — no sheetRow for new jobs)
    if (data.action === 'add_job') {
      // Find last row where both Name_Company (D) and Name_Print (F) are filled
      const colDF = sheet.getRange(1, 4, sheet.getMaxRows(), 3).getValues();
      let lastDataRow = 1;
      for (let i = colDF.length - 1; i >= 1; i--) {
        if (String(colDF[i][0]).trim() !== '' && String(colDF[i][2]).trim() !== '') {
          lastDataRow = i + 1; break;
        }
      }
      const newRow = lastDataRow + 1;

      // Calculate priority: count how many existing rows have the same Soort
      const soortVals = sheet.getRange(2, 2, lastDataRow - 1, 1).getValues();
      const priority  = soortVals.filter(r => String(r[0]).trim() === String(data.soort).trim()).length + 1;

      const vals   = new Array(35).fill('');
      const tz     = ss.getSpreadsheetTimeZone();
      vals[0]  = priority;
      vals[1]  = data.soort     || '';
      vals[2]  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
      vals[3]  = data.company   || '';
      vals[5]  = data.printName || '';
      // vals[6]  = col G  — set via Drive upload below
      // vals[10] = col K  — set via setFormula below
      vals[8]  = data.status    || 'To Print';
      vals[11] = data.owner     || '';
      vals[12] = data.deadline  || '';
      // vals[13] = col N  — set via setFormula below
      // vals[14] = col O  — set via setFormula below
      vals[15] = data.color     || '';
      vals[16] = data.lid       || '';
      vals[17] = data.quantity  ? parseInt(data.quantity) : '';
      // vals[18] = col S  — set via setFormula below
      // vals[19] = col T  — set via setFormula below
      vals[21] = data.tosleeve  || '';
      vals[25] = data.notes     || '';
      vals[35] = data.changedBy || '';
      sheet.getRange(newRow, 1, 1, vals.length).setValues([vals]);

      // Find last row with a valid (non-#ERROR!) formula in col K to copy from
      const kFormulas = sheet.getRange(2, 11, lastDataRow - 1, 1).getFormulas();
      const kValues   = sheet.getRange(2, 11, lastDataRow - 1, 1).getValues();
      let formulaSourceRow = null;
      for (let i = lastDataRow - 2; i >= 0; i--) {
        if (kFormulas[i][0] && String(kValues[i][0]) !== '#ERROR!') {
          formulaSourceRow = i + 2; break;
        }
      }
      if (formulaSourceRow) {
        [11, 14, 15, 19, 20].forEach(function(c) {
          const src = sheet.getRange(formulaSourceRow, c);
          if (src.getFormula()) {
            src.copyTo(sheet.getRange(newRow, c), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
          }
        });
      }

      // Upload mockup image to Drive and store URL in col G
      if (data.mockupBase64) {
        try {
          const base64Data = data.mockupBase64.includes(',')
            ? data.mockupBase64.split(',')[1]
            : data.mockupBase64;
          const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const fileName = 'mockup_' + Date.now() + '.jpg';
          const blob     = Utilities.newBlob(
            Utilities.base64Decode(base64Data), 'image/jpeg', fileName
          );
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          sheet.getRange(newRow, 7).setValue(
            'https://drive.google.com/uc?id=' + file.getId()
          );
        } catch (imgErr) {
          Logger.log('Mockup upload failed: ' + imgErr.message);
        }
      }

      // Upload design file to Drive and store URL in 'File' column
      if (data.designFileBase64) {
        try {
          const raw      = data.designFileBase64.includes(',') ? data.designFileBase64.split(',')[1] : data.designFileBase64;
          const mime     = data.designFileMime || 'application/octet-stream';
          const fname    = data.designFileName || ('design_file_' + Date.now());
          const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const blob     = Utilities.newBlob(Utilities.base64Decode(raw), mime, fname);
          const file     = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          const fileUrl  = 'https://drive.google.com/file/d/' + file.getId() + '/view';
          const fileCol  = headers.findIndex(h => h.toLowerCase() === 'file');
          if (fileCol >= 0) sheet.getRange(newRow, fileCol + 1).setValue(fileUrl);
        } catch (fileErr) {
          Logger.log('Design file upload failed: ' + fileErr.message);
        }
      }

      Logger.log('add_job: wrote row ' + newRow);
      return respond({ success: true, newRow });
    }

    // Update sleeve job (Sleeves sheet)
    if (data.action === 'update_sleeve') {
      const svSheet = ss.getSheetByName('Sleeves');
      const svHeaders = svSheet.getRange(1, 1, 1, svSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const rowIndex  = data.sheetRow ? parseInt(data.sheetRow) : -1;
      if (rowIndex < 2) return respond({ error: 'Invalid sheet row: ' + data.sheetRow });

      const findCol = kw => svHeaders.findIndex(h => h.toLowerCase().includes(kw.toLowerCase())) + 1;
      const exactCol = name => svHeaders.findIndex(h => h.toLowerCase() === name.toLowerCase()) + 1;

      if (data.status) {
        const c = exactCol('status');
        if (c > 0) svSheet.getRange(rowIndex, c).setValue(data.status);
      }
      if (data.changedBy) {
        const c = findCol('changed');
        if (c > 0) svSheet.getRange(rowIndex, c).setValue(data.changedBy);
      }
      // Upload multiple files to Drive and append URLs (newline-separated)
      if (data.sleeveFiles && data.sleeveFiles.length > 0) {
        try {
          const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const fc       = findCol('file');
          const newUrls  = [];
          data.sleeveFiles.forEach(function(sf) {
            try {
              const raw   = sf.base64.includes(',') ? sf.base64.split(',')[1] : sf.base64;
              const mime  = sf.mime || 'application/octet-stream';
              const fname = sf.name || ('sleeve_file_' + Date.now());
              const blob  = Utilities.newBlob(Utilities.base64Decode(raw), mime, fname);
              const file  = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              newUrls.push('https://drive.google.com/file/d/' + file.getId() + '/view');
            } catch (e) { Logger.log('File upload error: ' + e.message); }
          });
          if (fc > 0 && newUrls.length > 0) {
            const existing = String(svSheet.getRange(rowIndex, fc).getValue() || '').trim();
            const combined = [existing, ...newUrls].filter(Boolean).join('\n');
            svSheet.getRange(rowIndex, fc).setValue(combined);
          }
        } catch (fileErr) {
          Logger.log('Sleeve files upload failed: ' + fileErr.message);
        }
      }
      return respond({ success: true });
    }

    // Add sleeve job (Sleeves sheet)
    if (data.action === 'add_sleeve_job') {
      const svSheet   = ss.getSheetByName('Sleeves');
      const lastRow   = svSheet.getLastRow();
      const svHeaders = svSheet.getRange(1, 1, 1, Math.max(svSheet.getLastColumn(), 12)).getValues()[0].map(h => String(h).trim());
      const newRow    = lastRow + 1;
      const tz        = ss.getSpreadsheetTimeZone();

      const findIdx = kw => svHeaders.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));

      // Use passed priority if provided (e.g. from add_job auto-create), otherwise calculate
      let priority;
      if (data.priority !== undefined && data.priority !== null && data.priority !== '') {
        priority = parseInt(data.priority);
      } else {
        priority = 1;
        const soortIdx = findIdx('soort');
        if (lastRow > 1 && soortIdx >= 0) {
          const soortVals = svSheet.getRange(2, soortIdx + 1, lastRow - 1, 1).getValues();
          priority = soortVals.filter(r => String(r[0]).trim() === String(data.soort).trim()).length + 1;
        }
      }

      const vals = new Array(svHeaders.length).fill('');
      const set  = (kw, value) => { const i = findIdx(kw); if (i >= 0) vals[i] = value; };

      set('priority', priority);
      set('soort',    data.soort     || '');
      set('date',     Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy'));
      set('company',  data.company   || '');
      set('print',    data.printName || '');
      // exact match for 'Quantity' to avoid hitting 'Quantity sleeved'
      const qIdx = svHeaders.findIndex(h => h.toLowerCase() === 'quantity');
      if (qIdx >= 0) vals[qIdx] = data.quantity ? parseInt(data.quantity) : '';
      set('status',   'To Sleeve');
      set('owner',    data.owner    || '');
      set('deadline', data.deadline || '');
      set('notes',    data.notes    || '');
      set('changed',  data.changedBy || '');

      svSheet.getRange(newRow, 1, 1, vals.length).setValues([vals]);

      // Upload attached file to Drive and store URL
      if (data.sleeveFileBase64) {
        try {
          const raw      = data.sleeveFileBase64.includes(',') ? data.sleeveFileBase64.split(',')[1] : data.sleeveFileBase64;
          const mime     = data.sleeveFileMime || 'application/octet-stream';
          const fname    = data.sleeveFileName || ('sleeve_file_' + Date.now());
          const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const blob     = Utilities.newBlob(Utilities.base64Decode(raw), mime, fname);
          const file     = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          const fileUrl  = 'https://drive.google.com/file/d/' + file.getId() + '/view';
          const fc = findIdx('file');
          if (fc >= 0) svSheet.getRange(newRow, fc + 1).setValue(fileUrl);
        } catch (fileErr) {
          Logger.log('Sleeve file upload failed: ' + fileErr.message);
        }
      }

      Logger.log('add_sleeve_job: wrote row ' + newRow);
      return respond({ success: true, newRow });
    }

    // Use the exact sheet row number sent by the dashboard (most reliable — no search needed)
    const rowIndex = data.sheetRow ? parseInt(data.sheetRow) : -1;
    Logger.log('rowIndex=' + rowIndex);
    if (rowIndex < 2) return respond({ error: 'Invalid sheet row: ' + data.sheetRow });

    // Update Quantity Printed
    if (quantityPrintedCol >= 0 && data.quantityPrinted !== undefined) {
      sheet.getRange(rowIndex, quantityPrintedCol + 1).setValue(data.quantityPrinted);
    }

    // Update Faulty Prints
    if (faultyCol >= 0 && data.faultyPrints !== undefined) {
      sheet.getRange(rowIndex, faultyCol + 1).setValue(data.faultyPrints);
    }

    // Update Status (column Q = 17)
    if (data.status) {
      Logger.log('Updating status col Q=' + statusCol + ' to: ' + data.status);
      sheet.getRange(rowIndex, statusCol).setValue(data.status);
    }

    // Update Printer Used (column AI = 35)
    if (data.printer !== undefined) {
      sheet.getRange(rowIndex, printerCol).setValue(data.printer);
    }

    // Log who made the change (column AJ = 36)
    if (data.changedBy) {
      sheet.getRange(rowIndex, 36).setValue(data.changedBy);
    }

    // Phone photo already uploaded — just save the URL to column H
    if (data.phonePhotoUrl) {
      sheet.getRange(rowIndex, photoCol).setValue(data.phonePhotoUrl);
    }

    // Upload photo to Google Drive and store URL in column H
    if (data.imageBase64) {
      try {
        const base64Data = data.imageBase64.includes(',')
          ? data.imageBase64.split(',')[1]
          : data.imageBase64;

        const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        const fileName = 'job_' + data.priority + '_' + Date.now() + '.jpg';
        const blob     = Utilities.newBlob(
          Utilities.base64Decode(base64Data),
          'image/jpeg',
          fileName
        );
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        const photoUrl = 'https://drive.google.com/uc?id=' + file.getId();
        sheet.getRange(rowIndex, photoCol).setValue(photoUrl);
      } catch (imgErr) {
        Logger.log('Photo upload failed: ' + imgErr.message);
      }
    }

    return respond({ success: true });

  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/***** DEBUG — run this to test add_job directly without HTTP *****/
function testAddJob() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Workfile');
  const colDF = sheet.getRange(1, 4, sheet.getMaxRows(), 3).getValues();
  let lastDataRow = 1;
  for (let i = colDF.length - 1; i >= 1; i--) {
    if (String(colDF[i][0]).trim() !== '' && String(colDF[i][2]).trim() !== '') {
      lastDataRow = i + 1; break;
    }
  }
  const newRow = lastDataRow + 1;
  const soortVals = sheet.getRange(2, 2, lastDataRow - 1, 1).getValues();
  const priority  = soortVals.filter(r => String(r[0]).trim() === 'Bottle').length + 1;
  const vals   = new Array(35).fill('');
  const tz     = ss.getSpreadsheetTimeZone();
  vals[0]  = priority;
  vals[1]  = 'Bottle';
  vals[2]  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  vals[3]  = 'TEST COMPANY';
  vals[5]  = 'Test Print';
  vals[8]  = 'To Print';
  vals[17] = 10;
  sheet.getRange(newRow, 1, 1, vals.length).setValues([vals]);

  // Find last row with a valid (non-#ERROR!) formula in col K to copy from
  const kFormulas = sheet.getRange(2, 11, lastDataRow - 1, 1).getFormulas();
  const kValues   = sheet.getRange(2, 11, lastDataRow - 1, 1).getValues();
  let formulaSourceRow = null;
  for (let i = lastDataRow - 2; i >= 0; i--) {
    if (kFormulas[i][0] && String(kValues[i][0]) !== '#ERROR!') {
      formulaSourceRow = i + 2; break;
    }
  }
  if (formulaSourceRow) {
    [11, 14, 15, 19, 20].forEach(function(c) {
      const src = sheet.getRange(formulaSourceRow, c);
      if (src.getFormula()) {
        src.copyTo(sheet.getRange(newRow, c), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      }
    });
  }

  Logger.log('testAddJob: wrote row ' + newRow);
  SpreadsheetApp.getActive().toast('Test row added at row ' + newRow + ' ✅');
}

/***** DEBUG — run this manually in Apps Script to check column mapping *****/
function debugStatusColumn() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('Workfile');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  headers.forEach((h, i) => {
    const letter = columnToLetter_(i + 1);
    Logger.log(`Col ${letter} (${i + 1}): "${h}"`);
  });

  // Also show what's currently in row 2 col Q
  const valQ2 = sheet.getRange(2, 17).getValue();
  Logger.log('Value at Q2 (row 2, col 17): ' + valQ2);
}

function columnToLetter_(col) {
  let s = '', n = col;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}
