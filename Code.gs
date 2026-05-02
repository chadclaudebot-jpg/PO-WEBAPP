/**
 * PO WEBAPP V1.3
 * Backend for the multi-module PO webapp.
 *
 * V1.3 changes:
 *   - Generic row API: getRows / addRow / updateRow / deleteRow.
 *     The first argument is always the tab name (sheetName).
 *   - Multi-module support: one Apps Script project bound to one
 *     "PO WEBAPP BACKEND" spreadsheet that contains a tab per module:
 *       PURCHASED_ORDER, INCOMING_SHIPMENT, SAMPLES, NOTES.
 *   - listSheets() returns every tab in the active spreadsheet so the
 *     Settings panel can show them as suggestions.
 *
 * Setup:
 *   1. Open your "PO WEBAPP BACKEND" Google Sheet
 *      (rename the existing PURCHASED ORDERS sheet for clarity, optional)
 *   2. Add 4 tabs: PURCHASED_ORDER, INCOMING_SHIPMENT, SAMPLES, NOTES
 *      (or whatever names you want — set them in the in-app Settings later)
 *      Add column headers in row 1 of each tab.
 *   3. Extensions -> Apps Script
 *   4. Replace Code.gs with this file.
 *   5. Replace the Index HTML file with the V1.3 Index.html.
 *   6. Deploy -> Manage deployments -> Edit -> New version -> Deploy.
 */

// ===================== CONFIG =====================
// Default tab name used when the frontend does not pass one. The frontend
// stores its own per-module tab names in localStorage (Settings panel).
const SHEET_NAME = 'PURCHASED_ORDER';
// Server-side CacheService TTL for getRows() results.
// Reads within this window return the cached blob (~10x faster than re-reading
// the sheet). Writes (addRow / updateRow / deleteRow) invalidate the cache
// for that tab so users see their own edits immediately.
const CACHE_TTL_SECONDS = 60;
// ===================================================

/**
 * Serves the web app HTML.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PO WEBAPP V1.3')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Resolve a tab name to a Sheet object.
 * If `sheetName` is empty/undefined, falls back to SHEET_NAME.
 * If still not found, throws a descriptive error so the frontend can
 * surface it (instead of silently using the wrong tab).
 */
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = (sheetName && String(sheetName).trim()) || SHEET_NAME;
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet tab "' + name + '" not found in "' + ss.getName() +
                    '". Available tabs: ' +
                    ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  }
  return sheet;
}

/**
 * Lists every tab in the active spreadsheet (used by the Settings panel).
 */
function listSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const names = ss.getSheets().map(function (s) { return s.getName(); });
    return { ok: true, sheets: names, defaultSheet: SHEET_NAME, spreadsheetName: ss.getName() };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Reads all rows from a tab and returns { headers, rows, sheetName }.
 * Each row object has a __rowIndex pointing to its 1-based sheet row,
 * which the frontend uses for edit / delete.
 *
 * Results are cached in CacheService for CACHE_TTL_SECONDS so subsequent
 * reads are very fast. The cache is invalidated by addRow/updateRow/deleteRow.
 */
function getRows(sheetName) {
  try {
    const sheet = getSheet_(sheetName);
    const actualName = sheet.getName();
    const cache = CacheService.getScriptCache();
    const cacheKey = 'rows_' + actualName;

    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) { /* fall through */ }
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    let result;
    if (lastRow < 1 || lastCol < 1) {
      result = { headers: [], rows: [], sheetName: actualName };
    } else {
      const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const headers = values[0].map(function (h) { return String(h || '').trim(); });
      const tz = Session.getScriptTimeZone();
      const rows = values.slice(1).map(function (row, idx) {
        const obj = { __rowIndex: idx + 2 };
        headers.forEach(function (h, i) {
          let val = row[i];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
          }
          obj[h] = val;
        });
        return obj;
      });
      result = { headers: headers, rows: rows, sheetName: actualName };
    }

    // CacheService values are capped at 100KB — large sheets silently skip caching.
    try {
      cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_SECONDS);
    } catch (e) { /* payload too large; serve uncached */ }

    return result;
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Invalidates the CacheService entry for a tab so the next getRows() call
 * fetches fresh from the spreadsheet. Called after every write.
 */
function invalidateCache_(sheetName) {
  try {
    const sheet = getSheet_(sheetName);
    CacheService.getScriptCache().remove('rows_' + sheet.getName());
  } catch (e) { /* ignore */ }
}

/**
 * Appends a row to a tab. `data` is keyed by column header.
 */
function addRow(sheetName, data) {
  try {
    const sheet = getSheet_(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    sheet.appendRow(newRow);
    invalidateCache_(sheetName);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Updates a row in place. `rowIndex` is the 1-based sheet row.
 */
function updateRow(sheetName, rowIndex, data) {
  try {
    const sheet = getSheet_(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([newRow]);
    invalidateCache_(sheetName);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Moves a row from sourceTab to targetTab (called when STATUS → "Received").
 * 1. Writes the updated data back to the source row.
 * 2. Appends a copy to targetTab, mapped by that sheet's column headers.
 * 3. Deletes the source row so it no longer appears in PURCHASED_ORDER.
 * Both caches are invalidated so the next getRows() call is fresh.
 */
function moveToReceived(sourceTab, rowIndex, data, targetTab) {
  try {
    const srcSheet = getSheet_(sourceTab);
    const tgtSheet = getSheet_(targetTab);

    // Write the updated row back to the source sheet first.
    const srcHeaders = srcSheet.getRange(1, 1, 1, srcSheet.getLastColumn()).getValues()[0];
    const srcRow = srcHeaders.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    srcSheet.getRange(rowIndex, 1, 1, srcHeaders.length).setValues([srcRow]);

    // Append a copy to the target (RECEIVED) sheet, mapped by its own headers.
    // If the target sheet has no headers yet, seed it with the source headers first.
    let tgtHeaders;
    const tgtLastCol = tgtSheet.getLastColumn();
    if (tgtLastCol < 1) {
      tgtHeaders = srcHeaders;
      tgtSheet.appendRow(tgtHeaders);
    } else {
      tgtHeaders = tgtSheet.getRange(1, 1, 1, tgtLastCol).getValues()[0];
    }
    const tgtRow = tgtHeaders.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    tgtSheet.appendRow(tgtRow);

    invalidateCache_(sourceTab);
    invalidateCache_(targetTab);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Deletes a row from a tab. `rowIndex` is the 1-based sheet row.
 */
function deleteRow(sheetName, rowIndex) {
  try {
    const sheet = getSheet_(sheetName);
    sheet.deleteRow(rowIndex);
    invalidateCache_(sheetName);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Returns all item code → description pairs from the ITEMCODES tab.
 * Expects row 1 to be a header row; data starts at row 2.
 * Col A = Item Code, Col B = Description.
 * Result is cached in CacheService for CACHE_TTL_SECONDS.
 */
function getItemCodes() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'itemcodes_map';
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) { /* fall through */ }
    }
    const sheet = getSheet_('ITEMCODES');
    const lastRow = sheet.getLastRow();
    const result = { ok: true, map: {} };
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach(function (row) {
        const code = String(row[0] || '').trim();
        const desc = String(row[1] || '').trim();
        if (code) result.map[code.toUpperCase()] = desc;
      });
    }
    try { cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_SECONDS); } catch (e) {}
    return result;
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}
