/**
 * Code4.gs
 * NeoChrono receiver for the pharmacist-facing Google Sheet.
 *
 * Add this file to the Apps Script project deployed at:
 * https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec
 *
 * This file intentionally defines doPost() only, so it can coexist with the
 * existing doGet() in Code.gs.
 */

const NEOCHRONO_RECEIVER_SHEET_NAME = 'Schedule';

// Leave blank when this Apps Script project is bound to the correct spreadsheet.
// If it is a standalone Apps Script project, paste the destination spreadsheet ID here.
const NEOCHRONO_RECEIVER_SPREADSHEET_ID = '';

const NEOCHRONO_SCHEDULE_HEADERS = Object.freeze([
  'Generation ID',
  'Assignment ID',
  'Date',
  'Day',
  'Shift',
  'Slot',
  'Assigned Pharmacist',
  'Username',
  'Hours',
  'Credited Hours',
  'Required Skill',
  'Coverage For Pharmacist',
  'Coverage For Username',
  'Coverage Reason',
  'Shift Type',
  'Weekend',
  'Weekend Group',
  'Holiday',
  'Locked',
  'Manual',
  'Status',
  'Warning',
  'Updated At',
  'Updated By',
  'Finalized At'
]);

function doPost(e) {
  var transferId = '';

  try {
    if (!e || !e.parameter || !e.parameter.payload) {
      throw new Error('No schedule payload was received.');
    }

    var payload = JSON.parse(String(e.parameter.payload || ''));
    transferId = String(payload.transferId || '').trim();

    if (!transferId) {
      throw new Error('Transfer ID is missing.');
    }

    if (String(payload.action || '') !== 'replaceSchedule') {
      throw new Error('Invalid transfer action.');
    }

    if (String(payload.sheetName || '') !== NEOCHRONO_RECEIVER_SHEET_NAME) {
      throw new Error(
        'NeoChrono attempted to write to "' +
        String(payload.sheetName || '') +
        '" instead of "' +
        NEOCHRONO_RECEIVER_SHEET_NAME +
        '".'
      );
    }

    if (!Array.isArray(payload.headers)) {
      throw new Error('Schedule headers are missing.');
    }

    if (!neoChronoHeadersMatch_(payload.headers, NEOCHRONO_SCHEDULE_HEADERS)) {
      throw new Error(
        'The incoming schedule columns do not match the required Schedule column order.'
      );
    }

    if (!Array.isArray(payload.rows)) {
      throw new Error('Schedule rows are missing.');
    }

    var ss = neoChronoReceiverSpreadsheet_();
    var sheet = ss.getSheetByName(NEOCHRONO_RECEIVER_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(NEOCHRONO_RECEIVER_SHEET_NAME);
    }

    var cleanRows = payload.rows.map(function(row) {
      var source = Array.isArray(row) ? row : [];
      return NEOCHRONO_SCHEDULE_HEADERS.map(function(_header, index) {
        var value = source[index];
        if (value === null || value === undefined) return '';
        return value;
      });
    });

    // Replace the Schedule tab only. Other tabs in the spreadsheet are untouched.
    sheet.clearContents();

    // Always write the exact 25-column header row, even if there are no data rows.
    sheet
      .getRange(1, 1, 1, NEOCHRONO_SCHEDULE_HEADERS.length)
      .setValues([NEOCHRONO_SCHEDULE_HEADERS.slice()]);

    if (cleanRows.length) {
      sheet
        .getRange(2, 1, cleanRows.length, NEOCHRONO_SCHEDULE_HEADERS.length)
        .setValues(cleanRows);
    }

    // Keep the receiver easy to read.
    sheet.setFrozenRows(1);
    sheet
      .getRange(1, 1, 1, NEOCHRONO_SCHEDULE_HEADERS.length)
      .setFontWeight('bold');

    // Column C = Date.
    if (cleanRows.length) {
      sheet
        .getRange(2, 3, cleanRows.length, 1)
        .setNumberFormat('yyyy-mm-dd');
    }

    SpreadsheetApp.flush();

    // Verify that the sheet really contains what NeoChrono sent.
    var verifyHeaders = sheet
      .getRange(1, 1, 1, NEOCHRONO_SCHEDULE_HEADERS.length)
      .getValues()[0]
      .map(function(v) { return String(v || '').trim(); });

    if (!neoChronoHeadersMatch_(verifyHeaders, NEOCHRONO_SCHEDULE_HEADERS)) {
      throw new Error('Header verification failed after writing the Schedule sheet.');
    }

    var expectedLastRow = cleanRows.length + 1;
    if (sheet.getLastRow() !== expectedLastRow) {
      throw new Error(
        'Row verification failed. Expected ' +
        expectedLastRow +
        ' total rows but the Schedule sheet has ' +
        sheet.getLastRow() +
        '.'
      );
    }

    return neoChronoReceiverResponse_({
      type: 'NEOCHRONO_SCHEDULE_RECEIVER',
      ok: true,
      transferId: transferId,
      transferredRows: cleanRows.length,
      transferredColumns: NEOCHRONO_SCHEDULE_HEADERS.length,
      receiverSpreadsheetId: ss.getId(),
      receiverSpreadsheetName: ss.getName(),
      receiverSheetName: NEOCHRONO_RECEIVER_SHEET_NAME,
      receiverUrl:
        ss.getUrl() +
        '#gid=' +
        sheet.getSheetId(),
      message:
        'Schedule received successfully. ' +
        cleanRows.length +
        ' row(s) were written to the Schedule sheet.'
    });

  } catch (error) {
    return neoChronoReceiverResponse_({
      type: 'NEOCHRONO_SCHEDULE_RECEIVER',
      ok: false,
      transferId: transferId,
      message:
        error && error.message
          ? error.message
          : String(error)
    });
  }
}

function neoChronoReceiverSpreadsheet_() {
  if (NEOCHRONO_RECEIVER_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(
      NEOCHRONO_RECEIVER_SPREADSHEET_ID
    );
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'This Apps Script project is not bound to a spreadsheet. ' +
      'Paste the destination spreadsheet ID into NEOCHRONO_RECEIVER_SPREADSHEET_ID in Code4.gs.'
    );
  }

  return ss;
}

function neoChronoHeadersMatch_(actual, expected) {
  if (!Array.isArray(actual) || actual.length !== expected.length) {
    return false;
  }

  for (var i = 0; i < expected.length; i++) {
    if (
      String(actual[i] || '').trim() !==
      String(expected[i] || '').trim()
    ) {
      return false;
    }
  }

  return true;
}

/**
 * The response is HTML because the NeoChrono GitHub page submits through a
 * hidden iframe. postMessage lets the parent page verify the transfer without
 * relying on cross-origin fetch/CORS.
 */
function neoChronoReceiverResponse_(result) {
  var json = JSON.stringify(result)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  var html =
    '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>' +
    'try{' +
    'window.parent.postMessage(' + json + ',"*");' +
    '}catch(e){}' +
    '</script>' +
    '<div style="font-family:Arial,sans-serif;padding:20px">' +
    (result.ok
      ? '<h3 style="color:#13795b">Schedule received</h3>'
      : '<h3 style="color:#b42318">Schedule transfer failed</h3>') +
    '<p>' +
    neoChronoHtmlEscape_(result.message || '') +
    '</p>' +
    '</div>' +
    '</body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}

function neoChronoHtmlEscape_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Optional manual test from the Apps Script editor.
 */
function testNeoChronoScheduleReceiver() {
  var ss = neoChronoReceiverSpreadsheet_();
  var sheet = ss.getSheetByName(
    NEOCHRONO_RECEIVER_SHEET_NAME
  );

  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    scheduleSheetExists: !!sheet,
    scheduleSheetName: sheet ? sheet.getName() : '',
    scheduleLastRow: sheet ? sheet.getLastRow() : 0,
    scheduleLastColumn: sheet ? sheet.getLastColumn() : 0,
    spreadsheetUrl: ss.getUrl()
  };
}
