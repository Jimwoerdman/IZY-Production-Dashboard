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
      : e.parameter.sheet === 'mockups'
        ? ss.getSheetByName('Mockups')
        : e.parameter.sheet === 'stock'
          ? ss.getSheetByName('Stock')
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

  // Also return PrintLog when fetching main Workfile
  let printLog = [];
  if (!e.parameter.sheet || e.parameter.sheet === '') {
    const logSheet = ss.getSheetByName('PrintLog');
    if (logSheet && logSheet.getLastRow() > 1) {
      const logVals    = logSheet.getDataRange().getValues();
      const logHeaders = logVals[0].map(h => String(h).trim());
      printLog = logVals.slice(1).map(row => {
        const obj = Object.fromEntries(logHeaders.map((h, i) => {
          const v = row[i];
          if (v instanceof Date && !isNaN(v)) {
            return [h, v.getTime() === 0 ? '' : Utilities.formatDate(v, tz, 'dd/MM/yyyy')];
          }
          return [h, String(v ?? '').trim()];
        }));
        return obj;
      });
    }
  }

  return respondGet({ rows, printLog });
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
    const tz     = ss.getSpreadsheetTimeZone();
    const sheet  = ss.getSheetByName('Workfile');
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());

    // Find columns by header name
    const col = name => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
    const priorityCol        = col('Priority');
    const soortCol           = col('Soort');
    const quantityPrintedCol = headers.findIndex(h => h.toLowerCase().includes('quantity printed'));
    const faultyCol          = col('Faulty prints');
    const datePrintedCol     = headers.findIndex(h => h.toLowerCase() === 'date printed');

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

      const wfLen  = Math.max(headers.length, 36);
      const vals   = new Array(wfLen).fill('');
      const setW   = (kw, value) => { const i = headers.findIndex(h => h.toLowerCase().includes(kw.toLowerCase())); if (i >= 0) vals[i] = value; };
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
      { const i = headers.findIndex(h => h.toLowerCase() === 'bottle color'); if (i >= 0) vals[i] = data.color || ''; }
      { const i = headers.findIndex(h => h.toLowerCase().includes('lid'));   if (i >= 0) vals[i] = data.lid   || ''; }
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
      let mockupUrl = null;
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
          mockupUrl = 'https://drive.google.com/uc?id=' + file.getId();
          sheet.getRange(newRow, 7).setValue(mockupUrl);
        } catch (imgErr) {
          Logger.log('Mockup upload failed: ' + imgErr.message);
        }
      }

      // Upload design files to Drive and store URLs (newline-separated) in 'File' column
      let designFileUrls = [];
      if (data.designFiles && data.designFiles.length > 0) {
        try {
          const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const fileCol  = headers.findIndex(h => h.toLowerCase() === 'file');
          data.designFiles.forEach(function(df) {
            try {
              const raw   = df.base64.includes(',') ? df.base64.split(',')[1] : df.base64;
              const mime  = df.mime || 'application/octet-stream';
              const fname = df.name || ('design_file_' + Date.now());
              const blob  = Utilities.newBlob(Utilities.base64Decode(raw), mime, fname);
              const file  = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              designFileUrls.push('https://drive.google.com/file/d/' + file.getId() + '/view');
            } catch (e) { Logger.log('Design file upload error: ' + e.message); }
          });
          if (fileCol >= 0 && designFileUrls.length > 0) {
            sheet.getRange(newRow, fileCol + 1).setValue(designFileUrls.join('\n'));
          }
        } catch (fileErr) {
          Logger.log('Design files upload failed: ' + fileErr.message);
        }
      }

      // Auto-create sleeve job in Sleeves sheet (backend handles it so Drive URLs are available)
      if (data.tosleeve === 'Yes') {
        try {
          const svSheet   = ss.getSheetByName('Sleeves');
          const lastSvRow = svSheet.getLastRow();
          const svHeaders = svSheet.getRange(1, 1, 1, Math.max(svSheet.getLastColumn(), 20)).getValues()[0].map(h => String(h).trim());
          const newSvRow  = lastSvRow + 1;
          const findSvIdx = kw => svHeaders.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
          const svVals    = new Array(svHeaders.length).fill('');
          const setS      = (kw, value) => { const i = findSvIdx(kw); if (i >= 0) svVals[i] = value; };

          setS('priority', priority);
          setS('soort',    data.soort     || '');
          setS('date',     Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy'));
          setS('company',  data.company   || '');
          setS('print',    data.printName || '');
          const svQIdx = svHeaders.findIndex(h => h.toLowerCase() === 'quantity');
          if (svQIdx >= 0) svVals[svQIdx] = data.quantity ? parseInt(data.quantity) : '';
          setS('status',   'To make');
          setS('owner',    data.owner    || '');
          setS('deadline', data.deadline || '');
          setS('notes',    data.notes    || '');
          setS('changed',  data.changedBy || '');
          setS('bottle',   data.color    || '');
          setS('lid',      data.lid      || '');
          setS('mockup',   mockupUrl     || '');
          const svFileIdx = findSvIdx('file');
          if (svFileIdx >= 0 && designFileUrls.length > 0) {
            svVals[svFileIdx] = designFileUrls.join('\n');
          }

          svSheet.getRange(newSvRow, 1, 1, svVals.length).setValues([svVals]);
          Logger.log('add_job: auto-created sleeve row ' + newSvRow);
        } catch (svErr) {
          Logger.log('Auto sleeve job creation failed: ' + svErr.message);
        }
      }

      // Auto-create mockup job in Mockups sheet when needmockup = Yes
      if (data.needmockup === 'Yes') {
        try {
          const mkSheet   = ss.getSheetByName('Mockups');
          const lastMkRow = mkSheet.getLastRow();
          const mkHeaders = mkSheet.getRange(1, 1, 1, Math.max(mkSheet.getLastColumn(), 20)).getValues()[0].map(h => String(h).trim());
          const newMkRow  = lastMkRow + 1;
          const findMkIdx = kw => mkHeaders.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
          const mkVals    = new Array(mkHeaders.length).fill('');
          const setMk     = (kw, value) => { const i = findMkIdx(kw); if (i >= 0) mkVals[i] = value; };

          setMk('priority', priority);
          setMk('soort',    data.soort     || '');
          setMk('date',     Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy'));
          setMk('company',  data.company   || '');
          setMk('print',    data.printName || '');
          const mkQIdx = mkHeaders.findIndex(h => h.toLowerCase() === 'quantity');
          if (mkQIdx >= 0) mkVals[mkQIdx] = data.quantity ? parseInt(data.quantity) : '';
          setMk('status',   'To make');
          setMk('owner',    data.owner    || '');
          setMk('deadline', data.deadline || '');
          setMk('notes',    data.notes    || '');
          setMk('changed',  data.changedBy || '');
          setMk('bottle',   data.color    || '');
          setMk('lid',      data.lid      || '');
          setMk('mockup',   mockupUrl     || '');
          const mkFileIdx = findMkIdx('file');
          if (mkFileIdx >= 0 && designFileUrls.length > 0) {
            mkVals[mkFileIdx] = designFileUrls.join('\n');
          }

          mkSheet.getRange(newMkRow, 1, 1, mkVals.length).setValues([mkVals]);
          Logger.log('add_job: auto-created mockup row ' + newMkRow);
        } catch (mkErr) {
          Logger.log('Auto mockup job creation failed: ' + mkErr.message);
        }
      }

      Logger.log('add_job: wrote row ' + newRow);
      sendJobNotification(
        data.changedBy, data.owner,
        '🆕 Nieuwe print job: ' + (data.company || '') + ' – ' + (data.printName || ''),
        [
          ['Actie',    'Nieuwe print job toegevoegd'],
          ['Bedrijf',  data.company  || '—'],
          ['Product',  data.printName || '—'],
          ['Type',     data.soort    || '—'],
          ['Aantal',   data.quantity || '—'],
          ['Deadline', data.deadline || '—']
        ]
      );
      return respond({ success: true, newRow });
    }

    // Update sleeve job (Sleeves sheet)
    if (data.action === 'update_sleeve') {
      const svSheet = ss.getSheetByName('Sleeves');
      const svHeaders = svSheet.getRange(1, 1, 1, Math.max(svSheet.getLastColumn(), 20)).getValues()[0].map(h => String(h).trim());
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
      set('status',   data.status || 'To Sleeve');
      set('owner',    data.owner    || '');
      set('deadline', data.deadline || '');
      const svBottleIdx = svHeaders.findIndex(h => h.toLowerCase() === 'bottle color');
      if (svBottleIdx >= 0) vals[svBottleIdx] = data.bottleColor || '';
      const svLidIdx = svHeaders.findIndex(h => h.toLowerCase() === 'lid');
      if (svLidIdx >= 0) vals[svLidIdx] = data.lidColor || '';
      set('notes',    data.notes    || '');
      set('changed',  data.changedBy || '');

      svSheet.getRange(newRow, 1, 1, vals.length).setValues([vals]);

      // Upload design files to Drive and store URLs (newline-separated) in 'File' column
      if (data.designFiles && data.designFiles.length > 0) {
        try {
          const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const fileCol = findIdx('file');
          const urls    = [];
          data.designFiles.forEach(function(df) {
            try {
              const raw   = df.base64.includes(',') ? df.base64.split(',')[1] : df.base64;
              const mime  = df.mime || 'application/octet-stream';
              const fname = df.name || ('sleeve_file_' + Date.now());
              const blob  = Utilities.newBlob(Utilities.base64Decode(raw), mime, fname);
              const file  = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              urls.push('https://drive.google.com/file/d/' + file.getId() + '/view');
            } catch (e) { Logger.log('Sleeve file upload error: ' + e.message); }
          });
          if (fileCol >= 0 && urls.length > 0) svSheet.getRange(newRow, fileCol + 1).setValue(urls.join('\n'));
        } catch (fileErr) {
          Logger.log('Sleeve files upload failed: ' + fileErr.message);
        }
      }

      Logger.log('add_sleeve_job: wrote row ' + newRow);
      sendJobNotification(
        data.changedBy, data.owner,
        '🆕 Nieuwe sleeve job: ' + (data.company || '') + ' – ' + (data.printName || ''),
        [
          ['Actie',    'Nieuwe sleeve job toegevoegd'],
          ['Bedrijf',  data.company  || '—'],
          ['Product',  data.printName || '—'],
          ['Type',     data.soort    || '—'],
          ['Aantal',   data.quantity || '—'],
          ['Deadline', data.deadline || '—']
        ]
      );
      return respond({ success: true, newRow });
    }

    // Update mockup job (Mockups sheet)
    if (data.action === 'update_mockup') {
      const mkSheet   = ss.getSheetByName('Mockups');
      const mkHeaders = mkSheet.getRange(1, 1, 1, Math.max(mkSheet.getLastColumn(), 30)).getValues()[0].map(h => String(h).trim());
      const rowIndex  = data.sheetRow ? parseInt(data.sheetRow) : -1;
      if (rowIndex < 2) return respond({ error: 'Invalid sheet row: ' + data.sheetRow });

      const findCol  = kw => mkHeaders.findIndex(h => h.toLowerCase().includes(kw.toLowerCase())) + 1;
      const exactCol = name => mkHeaders.findIndex(h => h.toLowerCase() === name.toLowerCase()) + 1;

      if (data.status) {
        const c = exactCol('status');
        if (c > 0) mkSheet.getRange(rowIndex, c).setValue(data.status);
      }
      if (data.changedBy) {
        const c = findCol('changed');
        if (c > 0) mkSheet.getRange(rowIndex, c).setValue(data.changedBy);
      }
      if (data.mockupFiles && data.mockupFiles.length > 0) {
        try {
          const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const fc      = findCol('file');
          const newUrls = [];
          data.mockupFiles.forEach(function(mf) {
            try {
              const raw   = mf.base64.includes(',') ? mf.base64.split(',')[1] : mf.base64;
              const mime  = mf.mime || 'application/octet-stream';
              const fname = mf.name || ('mockup_file_' + Date.now());
              const blob  = Utilities.newBlob(Utilities.base64Decode(raw), mime, fname);
              const file  = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              newUrls.push('https://drive.google.com/file/d/' + file.getId() + '/view');
            } catch (e) { Logger.log('Mockup file upload error: ' + e.message); }
          });
          if (fc > 0 && newUrls.length > 0) {
            const existing = String(mkSheet.getRange(rowIndex, fc).getValue() || '').trim();
            const combined = [existing, ...newUrls].filter(Boolean).join('\n');
            mkSheet.getRange(rowIndex, fc).setValue(combined);
          }
        } catch (fileErr) {
          Logger.log('Mockup files upload failed: ' + fileErr.message);
        }
      }
      return respond({ success: true });
    }

    // Add mockup job (Mockups sheet)
    if (data.action === 'add_mockup_job') {
      const mkSheet   = ss.getSheetByName('Mockups');
      const lastRow   = mkSheet.getLastRow();
      const mkHeaders = mkSheet.getRange(1, 1, 1, Math.max(mkSheet.getLastColumn(), 20)).getValues()[0].map(h => String(h).trim());
      const newRow    = lastRow + 1;
      const findIdx   = kw => mkHeaders.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));

      let priority = 1;
      const soortIdx = findIdx('soort');
      if (lastRow > 1 && soortIdx >= 0) {
        const soortVals = mkSheet.getRange(2, soortIdx + 1, lastRow - 1, 1).getValues();
        priority = soortVals.filter(r => String(r[0]).trim() === String(data.soort).trim()).length + 1;
      }

      const vals = new Array(mkHeaders.length).fill('');
      const set  = (kw, value) => { const i = findIdx(kw); if (i >= 0) vals[i] = value; };

      set('priority', priority);
      set('soort',    data.soort     || '');
      set('date',     Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy'));
      set('company',  data.company   || '');
      set('print',    data.printName || '');
      const qIdx = mkHeaders.findIndex(h => h.toLowerCase() === 'quantity');
      if (qIdx >= 0) vals[qIdx] = data.quantity ? parseInt(data.quantity) : '';
      set('status',   data.status || 'To make');
      set('owner',    data.owner    || '');
      set('deadline', data.deadline || '');
      const mkBottleIdx = mkHeaders.findIndex(h => h.toLowerCase() === 'bottle color');
      if (mkBottleIdx >= 0) vals[mkBottleIdx] = data.bottleColor || '';
      const mkLidIdx = mkHeaders.findIndex(h => h.toLowerCase() === 'lid');
      if (mkLidIdx >= 0) vals[mkLidIdx] = data.lidColor || '';
      set('notes',    data.notes    || '');
      set('changed',  data.changedBy || '');

      mkSheet.getRange(newRow, 1, 1, vals.length).setValues([vals]);

      // Upload design files to Drive and store URLs (newline-separated) in 'File' column
      if (data.designFiles && data.designFiles.length > 0) {
        try {
          const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const fileCol = findIdx('file');
          const urls    = [];
          data.designFiles.forEach(function(df) {
            try {
              const raw   = df.base64.includes(',') ? df.base64.split(',')[1] : df.base64;
              const mime  = df.mime || 'application/octet-stream';
              const fname = df.name || ('mockup_file_' + Date.now());
              const blob  = Utilities.newBlob(Utilities.base64Decode(raw), mime, fname);
              const file  = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              urls.push('https://drive.google.com/file/d/' + file.getId() + '/view');
            } catch (e) { Logger.log('Mockup file upload error: ' + e.message); }
          });
          if (fileCol >= 0 && urls.length > 0) mkSheet.getRange(newRow, fileCol + 1).setValue(urls.join('\n'));
        } catch (fileErr) {
          Logger.log('Mockup files upload failed: ' + fileErr.message);
        }
      }

      Logger.log('add_mockup_job: wrote row ' + newRow);
      sendJobNotification(
        data.changedBy, data.owner,
        '🆕 Nieuwe mockup job: ' + (data.company || '') + ' – ' + (data.printName || ''),
        [
          ['Actie',    'Nieuwe mockup job toegevoegd'],
          ['Bedrijf',  data.company  || '—'],
          ['Product',  data.printName || '—'],
          ['Type',     data.soort    || '—'],
          ['Aantal',   data.quantity || '—'],
          ['Deadline', data.deadline || '—']
        ]
      );
      return respond({ success: true, newRow });
    }

    // Edit sleeve job
    if (data.action === 'edit_sleeve_job') {
      const svSheet   = ss.getSheetByName('Sleeves');
      const svHeaders = svSheet.getRange(1, 1, 1, svSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const rowIdx    = parseInt(data.sheetRow);
      if (rowIdx < 2) return respond({ error: 'Invalid sheet row' });
      const findSv = kw => svHeaders.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
      const setSv  = (kw, val) => { const c = findSv(kw); if (c >= 0) svSheet.getRange(rowIdx, c + 1).setValue(val); };
      setSv('name_company',  data.company    || '');
      setSv('name_print',    data.printName  || '');
      setSv('soort',         data.soort      || '');
      setSv('bottle color',  data.bottleColor|| '');
      setSv('lid',           data.lidColor   || '');
      setSv('quantity',      data.quantity   || '');
      setSv('deadline',      data.deadline   || '');
      setSv('owner',         data.owner      || '');
      setSv('notes',         data.notes      || '');
      if (data.designFiles && data.designFiles.length > 0) {
        try {
          const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const fileCol = findSv('file');
          const newUrls = [];
          data.designFiles.forEach(function(df) {
            try {
              const raw   = df.base64.includes(',') ? df.base64.split(',')[1] : df.base64;
              const blob  = Utilities.newBlob(Utilities.base64Decode(raw), df.mime || 'application/octet-stream', df.name || ('file_' + Date.now()));
              const file  = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              newUrls.push('https://drive.google.com/file/d/' + file.getId() + '/view');
            } catch(e) { Logger.log('edit_sleeve file error: ' + e.message); }
          });
          if (fileCol >= 0 && newUrls.length > 0) {
            const existing = String(svSheet.getRange(rowIdx, fileCol + 1).getValue() || '').trim();
            svSheet.getRange(rowIdx, fileCol + 1).setValue([existing, ...newUrls].filter(Boolean).join('\n'));
          }
        } catch(fileErr) { Logger.log('edit_sleeve files failed: ' + fileErr.message); }
      }
      Logger.log('edit_sleeve_job: updated row ' + rowIdx);
      sendJobNotification(
        data.changedBy, data.owner,
        '✏️ Sleeve job gewijzigd: ' + (data.company || ''),
        [
          ['Actie',    'Sleeve job aangepast'],
          ['Bedrijf',  data.company  || '—'],
          ['Type',     data.soort    || '—'],
          ['Deadline', data.deadline || '—']
        ]
      );
      return respond({ success: true });
    }

    // Edit mockup job
    if (data.action === 'edit_mockup_job') {
      const mkSheet   = ss.getSheetByName('Mockups');
      const mkHeaders = mkSheet.getRange(1, 1, 1, mkSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const rowIdx    = parseInt(data.sheetRow);
      if (rowIdx < 2) return respond({ error: 'Invalid sheet row' });
      const findMk = kw => mkHeaders.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
      const setMk  = (kw, val) => { const c = findMk(kw); if (c >= 0) mkSheet.getRange(rowIdx, c + 1).setValue(val); };
      setMk('name_company',  data.company    || '');
      setMk('name_print',    data.printName  || '');
      setMk('soort',         data.soort      || '');
      setMk('bottle color',  data.bottleColor|| '');
      setMk('lid',           data.lidColor   || '');
      setMk('quantity',      data.quantity   || '');
      setMk('deadline',      data.deadline   || '');
      setMk('owner',         data.owner      || '');
      setMk('notes',         data.notes      || '');
      if (data.designFiles && data.designFiles.length > 0) {
        try {
          const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const fileCol = findMk('file');
          const newUrls = [];
          data.designFiles.forEach(function(df) {
            try {
              const raw   = df.base64.includes(',') ? df.base64.split(',')[1] : df.base64;
              const blob  = Utilities.newBlob(Utilities.base64Decode(raw), df.mime || 'application/octet-stream', df.name || ('file_' + Date.now()));
              const file  = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              newUrls.push('https://drive.google.com/file/d/' + file.getId() + '/view');
            } catch(e) { Logger.log('edit_mockup file error: ' + e.message); }
          });
          if (fileCol >= 0 && newUrls.length > 0) {
            const existing = String(mkSheet.getRange(rowIdx, fileCol + 1).getValue() || '').trim();
            mkSheet.getRange(rowIdx, fileCol + 1).setValue([existing, ...newUrls].filter(Boolean).join('\n'));
          }
        } catch(fileErr) { Logger.log('edit_mockup files failed: ' + fileErr.message); }
      }
      Logger.log('edit_mockup_job: updated row ' + rowIdx);
      sendJobNotification(
        data.changedBy, data.owner,
        '✏️ Mockup job gewijzigd: ' + (data.company || ''),
        [
          ['Actie',    'Mockup job aangepast'],
          ['Bedrijf',  data.company  || '—'],
          ['Type',     data.soort    || '—'],
          ['Deadline', data.deadline || '—']
        ]
      );
      return respond({ success: true });
    }

    // Add stock delivery
    if (data.action === 'add_stock') {
      const stSheet = ss.getSheetByName('Stock');
      if (!stSheet) return respond({ error: 'Stock sheet not found' });
      const stVals    = stSheet.getDataRange().getValues();
      const stHeaders = stVals[0].map(h => String(h).trim().toLowerCase());
      const stTypeCol  = stHeaders.indexOf('type');
      const stColorCol = stHeaders.indexOf('color');
      const stQtyCol   = stHeaders.indexOf('quantity');
      if (stTypeCol < 0 || stColorCol < 0 || stQtyCol < 0) return respond({ error: 'Stock sheet missing columns' });

      const tl = (data.type  || '').trim().toLowerCase();
      const cl = (data.color || '').trim().toLowerCase();
      const amount = parseInt(data.quantity) || 0;
      if (!tl || !cl || amount <= 0) return respond({ error: 'Invalid type, color or quantity' });

      let matched = false;
      for (var si = 1; si < stVals.length; si++) {
        if (String(stVals[si][stTypeCol]  ?? '').trim().toLowerCase() === tl &&
            String(stVals[si][stColorCol] ?? '').trim().toLowerCase() === cl) {
          const current = parseInt(stVals[si][stQtyCol]) || 0;
          const newQty  = current + amount;
          stSheet.getRange(si + 1, stQtyCol + 1).setValue(newQty);

          // Log to StockLog
          let slSheet = ss.getSheetByName('StockLog');
          if (!slSheet) {
            slSheet = ss.insertSheet('StockLog');
            slSheet.appendRow(['Date', 'Type', 'Color', 'Deducted', 'Result', 'Job Row', 'Logged By', 'Note']);
          }
          const tzz = ss.getSpreadsheetTimeZone();
          slSheet.appendRow([
            Utilities.formatDate(new Date(), tzz, 'dd/MM/yyyy'),
            data.type, data.color, amount,
            'delivery (' + current + '→' + newQty + ')',
            '', data.changedBy || '', data.note || ''
          ]);
          matched = true;
          break;
        }
      }
      if (!matched) return respond({ error: 'No Stock row found for ' + data.type + ' / ' + data.color });
      return respond({ success: true });
    }

    // Use the exact sheet row number sent by the dashboard (most reliable — no search needed)
    const rowIndex = data.sheetRow ? parseInt(data.sheetRow) : -1;
    Logger.log('rowIndex=' + rowIndex);
    if (rowIndex < 2) return respond({ error: 'Invalid sheet row: ' + data.sheetRow });

    // Update Quantity Printed
    if (quantityPrintedCol >= 0 && data.quantityPrinted !== undefined) {
      sheet.getRange(rowIndex, quantityPrintedCol + 1).setValue(data.quantityPrinted);
    }

    // Append to PrintLog (always, independent of quantityPrintedCol)
    const sessionQty = parseInt(data.sessionPrinted);
    const isReset    = data.quantityPrinted === 0 && !data.sessionPrinted;
    Logger.log('sessionQty=' + sessionQty + ' isReset=' + isReset);

    const rowData = values[rowIndex - 1]; // 0-based: row 2 → index 1
    const g = (kw) => { const i = headers.findIndex(h => h.toLowerCase().includes(kw.toLowerCase())); return i >= 0 ? String(rowData[i] ?? '').trim() : ''; };

    // Get or create PrintLog sheet
    const ensureLogSheet = () => {
      let ls = ss.getSheetByName('PrintLog');
      if (!ls) {
        ls = ss.insertSheet('PrintLog');
        ls.appendRow(['Date', 'Company', 'Print Name', 'Owner', 'Type', 'Quantity', 'Priority', 'Logged By', 'SheetRow']);
      }
      return ls;
    };

    if (sessionQty > 0) {
      const logSheet = ensureLogSheet();
      logSheet.appendRow([
        Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy'),
        g('company') || g('name_company'),
        g('name_print') || g('print'),
        g('owner'),
        g('soort'),
        sessionQty,
        g('priority'),
        data.changedBy || '',
        rowIndex  // SheetRow — used to delete entries on reset
      ]);
      Logger.log('PrintLog row appended for sessionQty=' + sessionQty);
    }

    // On reset: delete PrintLog rows and refund stock
    if (isReset) {
      const logSheet = ss.getSheetByName('PrintLog');
      let refundQty = 0;
      if (logSheet && logSheet.getLastRow() > 1) {
        const logData = logSheet.getDataRange().getValues();
        // Sum quantities (index 5) before deleting; SheetRow is index 8
        for (var li = 1; li < logData.length; li++) {
          if (parseInt(logData[li][8]) === rowIndex) {
            refundQty += parseInt(logData[li][5]) || 0;
          }
        }
        // Delete rows bottom-up
        for (var li = logData.length - 1; li >= 1; li--) {
          if (parseInt(logData[li][8]) === rowIndex) logSheet.deleteRow(li + 1);
        }
        Logger.log('PrintLog rows deleted for sheetRow=' + rowIndex + ', refundQty=' + refundQty);
      }

      // Refund stock
      if (refundQty > 0) {
        const stockSheet = ss.getSheetByName('Stock');
        if (stockSheet && stockSheet.getLastRow() >= 2) {
          const stockVals    = stockSheet.getDataRange().getValues();
          const stockHeaders = stockVals[0].map(h => String(h).trim().toLowerCase());
          const stTypeCol    = stockHeaders.indexOf('type');
          const stColorCol   = stockHeaders.indexOf('color');
          const stQtyCol     = stockHeaders.indexOf('quantity');

          if (stTypeCol >= 0 && stColorCol >= 0 && stQtyCol >= 0) {
            const exactCol = (name) => {
              const i = headers.findIndex(h => h.trim().toLowerCase() === name.toLowerCase());
              return i >= 0 ? String(rowData[i] ?? '').trim() : '';
            };
            const normalizeType = (s) => {
              const sl = (s || '').toLowerCase().trim();
              if (sl === 'bottle' || sl.startsWith('bottle sample')) return 'Bottle';
              if (sl === 'mug'    || sl.startsWith('mug sample'))    return 'Mug';
              if (sl.startsWith('travel bottle'))                     return 'Travel Bottle';
              if (sl === 'tumbler' || sl.startsWith('tumbler sample'))return 'Tumbler';
              return (s || '').trim().replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
            };

            const bottleColor = exactCol('Bottle color');
            const lidColor    = exactCol('Lid');
            const soortRaw    = data.soort || exactCol('Soort');
            const stockType   = normalizeType(soortRaw);

            let stockLog = ss.getSheetByName('StockLog');
            if (!stockLog) {
              stockLog = ss.insertSheet('StockLog');
              stockLog.appendRow(['Date', 'Type', 'Color', 'Deducted', 'Result', 'Job Row', 'Logged By', 'Note']);
            }
            const logDate = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');

            const adjustReset = (typeName, colorName, delta, note) => {
              if (!typeName || !colorName) return;
              const tl = typeName.trim().toLowerCase();
              const cl = colorName.trim().toLowerCase();
              for (var si = 1; si < stockVals.length; si++) {
                if (String(stockVals[si][stTypeCol] ?? '').trim().toLowerCase() === tl &&
                    String(stockVals[si][stColorCol] ?? '').trim().toLowerCase() === cl) {
                  const current = parseInt(stockVals[si][stQtyCol]) || 0;
                  const newQty  = Math.max(0, current + delta);
                  stockSheet.getRange(si + 1, stQtyCol + 1).setValue(newQty);
                  Logger.log('Stock reset ' + (delta > 0 ? 'refund' : 're-deduct') + ': ' + typeName + '/' + colorName + ' ' + (delta > 0 ? '+' : '') + delta + ' (' + current + '→' + newQty + ')');
                  stockLog.appendRow([logDate, typeName, colorName, delta, 'reset (' + current + '→' + newQty + ')', rowIndex, data.changedBy || '', note || '']);
                  return;
                }
              }
              Logger.log('Stock reset: no match for ' + typeName + '/' + colorName);
              stockLog.appendRow([logDate, typeName, colorName, delta, 'no match', rowIndex, data.changedBy || '', note || '']);
            };

            // Refund product stock
            adjustReset(stockType, bottleColor, refundQty, 'reset — product refund');

            // Reverse lid swap if lid color ≠ bottle color
            if (lidColor && lidColor.trim().toLowerCase() !== (bottleColor || '').trim().toLowerCase()) {
              const soortLower = (soortRaw || '').toLowerCase();
              const lidType = (soortLower.includes('mug') || soortLower.includes('tumbler')) ? 'Mug lids' : 'Bottle lids';
              adjustReset(lidType, lidColor,    +refundQty, 'reset — spare lid refund');
              adjustReset(lidType, bottleColor, -refundQty, 'reset — matching lid re-deducted');
            }
          }
        }
      }
    }

    // Deduct from Stock sheet when prints are logged
    if (sessionQty > 0) {
      const stockSheet = ss.getSheetByName('Stock');

      // Ensure StockLog sheet exists for audit trail
      const ensureStockLog = () => {
        let sl = ss.getSheetByName('StockLog');
        if (!sl) {
          sl = ss.insertSheet('StockLog');
          sl.appendRow(['Date', 'Type', 'Color', 'Deducted', 'Result', 'Job Row', 'Logged By', 'Note']);
        }
        return sl;
      };

      if (!stockSheet) {
        Logger.log('Stock: sheet "Stock" not found — skipping deduction');
      } else if (stockSheet.getLastRow() < 2) {
        Logger.log('Stock: sheet is empty — skipping deduction');
      } else {
        const stockVals    = stockSheet.getDataRange().getValues();
        const stockHeaders = stockVals[0].map(h => String(h).trim().toLowerCase());
        const stTypeCol    = stockHeaders.indexOf('type');
        const stColorCol   = stockHeaders.indexOf('color');
        const stQtyCol     = stockHeaders.indexOf('quantity');

        if (stTypeCol < 0 || stColorCol < 0 || stQtyCol < 0) {
          Logger.log('Stock: missing required column(s). Found headers: ' + stockHeaders.join(', '));
        } else {
          // Use exact header match (not includes) to avoid ambiguous column lookups
          const exactCol = (name) => {
            const i = headers.findIndex(h => h.trim().toLowerCase() === name.toLowerCase());
            return i >= 0 ? String(rowData[i] ?? '').trim() : '';
          };

          const bottleColor = exactCol('Bottle color');
          const lidColor    = exactCol('Lid');
          const soortRaw    = data.soort || exactCol('Soort');
          const totalDeduct = sessionQty + (parseInt(data.faultyPrints) || 0);

          // Map Workfile Soort values → Stock sheet Type names
          // Handles sample variants, capitalisation differences, etc.
          const normalizeType = (s) => {
            const sl = (s || '').toLowerCase().trim();
            if (sl === 'bottle' || sl.startsWith('bottle sample')) return 'Bottle';
            if (sl === 'mug'    || sl.startsWith('mug sample'))    return 'Mug';
            if (sl.startsWith('travel bottle'))                     return 'Travel Bottle';
            if (sl === 'tumbler' || sl.startsWith('tumbler sample'))return 'Tumbler';
            // Fallback: title-case the raw value
            return (s || '').trim().replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
          };

          const stockType = normalizeType(soortRaw);
          const stockLog  = ensureStockLog();
          const logDate   = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');

          // Adjust a stock row by delta (negative = deduct, positive = add back)
          const adjustStock = (typeName, colorName, delta, note) => {
            if (!typeName || !colorName) {
              const msg = 'skipped — missing type or color (type="' + typeName + '" color="' + colorName + '")';
              Logger.log('Stock ' + (note || '') + ': ' + msg);
              stockLog.appendRow([logDate, typeName, colorName, delta, 'skipped', rowIndex, data.changedBy || '', msg]);
              return;
            }
            const tl = typeName.trim().toLowerCase();
            const cl = colorName.trim().toLowerCase();
            let matched = false;
            for (var si = 1; si < stockVals.length; si++) {
              const rowType  = String(stockVals[si][stTypeCol]  ?? '').trim().toLowerCase();
              const rowColor = String(stockVals[si][stColorCol] ?? '').trim().toLowerCase();
              if (rowType === tl && rowColor === cl) {
                const current = parseInt(stockVals[si][stQtyCol]) || 0;
                const newQty  = Math.max(0, current + delta);
                stockSheet.getRange(si + 1, stQtyCol + 1).setValue(newQty);
                const direction = delta < 0 ? 'deducted' : 'added';
                Logger.log('Stock ' + direction + ': ' + typeName + '/' + colorName + ' ' + delta + ' (' + current + '→' + newQty + ')');
                stockLog.appendRow([logDate, typeName, colorName, delta, 'ok (' + current + '→' + newQty + ')', rowIndex, data.changedBy || '', note || '']);
                matched = true;
                break;
              }
            }
            if (!matched) {
              const msg = 'no match in Stock sheet for type="' + typeName + '" color="' + colorName + '" (soortRaw="' + soortRaw + '")';
              Logger.log('Stock ' + (note || '') + ': ' + msg);
              stockLog.appendRow([logDate, typeName, colorName, delta, 'no match', rowIndex, data.changedBy || '', msg]);
            }
          };

          // Deduct product stock
          adjustStock(stockType, bottleColor, -totalDeduct, 'product');

          // Lid logic:
          // Every bottle/mug comes with a matching lid included in the bottle stock.
          // Spare lids only need adjusting when the lid color differs from the bottle color:
          //   - Deduct the chosen lid color from spare lids (you're using a spare)
          //   - Add the original matching lid color back to spare lids (it's now unused/spare)
          if (lidColor) {
            const soortLower = (soortRaw || '').toLowerCase();
            const lidType = (soortLower.includes('mug') || soortLower.includes('tumbler')) ? 'Mug lids' : 'Bottle lids';
            if (lidColor.trim().toLowerCase() !== (bottleColor || '').trim().toLowerCase()) {
              adjustStock(lidType, lidColor,    -totalDeduct, 'lids — spare used');
              adjustStock(lidType, bottleColor,  totalDeduct, 'lids — original returned to spare');
            }
            // If lid color === bottle color: lid comes with the bottle, no spare lid movement needed
          }
        }
      }
    }

    // Update Faulty Prints
    if (faultyCol >= 0 && data.faultyPrints !== undefined) {
      sheet.getRange(rowIndex, faultyCol + 1).setValue(data.faultyPrints);
    }

    // Update Status (column Q = 17)
    if (data.status) {
      Logger.log('Updating status col Q=' + statusCol + ' to: ' + data.status);
      sheet.getRange(rowIndex, statusCol).setValue(data.status);

      // Write today as Date Printed when status is set to a printed variant
      const PRINTED_VARIANTS = ['printed', 'print ready', 'printing ready'];
      if (datePrintedCol >= 0 && PRINTED_VARIANTS.some(v => data.status.toLowerCase().includes(v))) {
        sheet.getRange(rowIndex, datePrintedCol + 1).setValue(
          Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy')
        );
      }
      // Clear Date Printed if status is moved back to To Print / Waiting
      const UNPRINT_VARIANTS = ['to print', 'waiting'];
      if (datePrintedCol >= 0 && UNPRINT_VARIANTS.some(v => data.status.toLowerCase().includes(v))) {
        sheet.getRange(rowIndex, datePrintedCol + 1).setValue('');
      }
    }

    // Write shipping date to column J when status is set to Shipped
    if (data.shippingDate) {
      sheet.getRange(rowIndex, 10).setValue(data.shippingDate);
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

/***** JOB CHANGE NOTIFICATIONS *****/
const GEERTJAN = 'geertjan@izybottles.com';
const JIM      = 'jim@izybottles.com';
const SHARON   = 'sharon@orderchamp.com';

function buildHtmlEmail(subject, lines, changedBy) {
  var dashUrl = 'https://jimwoerdman.github.io/IZY-Production-Dashboard/';
  var rows = lines.map(function(l) {
    return '<tr><td style="padding:6px 0;color:#6b7a99;font-size:13px;white-space:nowrap;padding-right:16px;">' +
           l[0] + '</td><td style="padding:6px 0;color:#0f1629;font-size:13px;font-weight:500;">' + l[1] + '</td></tr>';
  }).join('');
  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f7;font-family:\'Segoe UI\',Arial,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f2f7;padding:32px 16px;">' +
    '<tr><td align="center">' +
    '<table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10);">' +
    // Header
    '<tr><td style="background:#0f1629;padding:28px 32px;">' +
    '<div style="color:#ffffff;font-size:20px;font-weight:700;letter-spacing:0.5px;">IZY Production Dashboard</div>' +
    '<div style="color:#8892aa;font-size:13px;margin-top:4px;">Job notificatie</div>' +
    '</td></tr>' +
    // Body
    '<tr><td style="background:#ffffff;padding:28px 32px;">' +
    '<div style="font-size:16px;font-weight:600;color:#0f1629;margin-bottom:20px;">' + subject + '</div>' +
    '<table cellpadding="0" cellspacing="0" style="width:100%;border-top:1px solid #e8ebf2;">' + rows + '</table>' +
    '</td></tr>' +
    // Footer
    '<tr><td style="background:#f7f8fb;padding:20px 32px;border-top:1px solid #e8ebf2;">' +
    '<div style="font-size:12px;color:#6b7a99;margin-bottom:12px;">Gedaan door: <strong style="color:#0f1629;">' + changedBy + '</strong></div>' +
    '<a href="' + dashUrl + '" style="display:inline-block;background:#2563eb;color:#ffffff;font-size:13px;font-weight:600;padding:10px 20px;border-radius:6px;text-decoration:none;">Bekijk dashboard →</a>' +
    '</td></tr>' +
    '</table></td></tr></table>' +
    '</body></html>';
}

function sendJobNotification(changedBy, owner, subject, lines) {
  if (!changedBy) return;
  if (changedBy === SHARON) return; // Sharon is volledig uitgesloten

  const recipients = new Set();

  // Notify the job owner if someone else made the change (Sharon nooit notificeren)
  if (owner && owner !== changedBy && owner !== SHARON) recipients.add(owner);

  // Geertjan's changes always notify Jim
  if (changedBy === GEERTJAN) recipients.add(JIM);

  const htmlBody = buildHtmlEmail(subject, lines, changedBy);
  const plainBody = lines.map(function(l){ return l[0] + ': ' + l[1]; }).join('\n') +
                    '\n\nGedaan door: ' + changedBy +
                    '\n\nDashboard: https://jimwoerdman.github.io/IZY-Production-Dashboard/';

  recipients.forEach(function(email) {
    try {
      MailApp.sendEmail({
        to:       email,
        subject:  subject,
        body:     plainBody,
        htmlBody: htmlBody
      });
    } catch(e) { Logger.log('Mail failed to ' + email + ': ' + e.message); }
  });
}

/***** TEST — run this to verify mail notifications work *****/
function testMail() {
  var lines = [
    ['Bedrijf',    'Test Company BV'],
    ['Actie',      'Nieuwe print job toegevoegd'],
    ['Job type',   'Bottle'],
    ['Deadline',   '01/04/2026'],
    ['Bestelling', '#12345']
  ];
  var htmlBody  = buildHtmlEmail('Nieuwe job: Test Company BV', lines, 'geertjan@izybottles.com');
  var plainBody = lines.map(function(l){ return l[0] + ': ' + l[1]; }).join('\n') +
                  '\n\nGedaan door: geertjan@izybottles.com' +
                  '\n\nDashboard: https://jimwoerdman.github.io/IZY-Production-Dashboard/';
  MailApp.sendEmail({
    to:       'jim@izybottles.com',
    subject:  '🧪 Testmelding IZY Dashboard — Nieuwe job: Test Company BV',
    body:     plainBody,
    htmlBody: htmlBody
  });
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
