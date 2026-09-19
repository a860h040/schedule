/**
 * Code5.gs
 * NeoChrono -> Pharmacists Schedule receiver
 *
 * Receives the schedule from the GitHub NeoChrono app and replaces the
 * Google Sheet tab named exactly "Schedule".
 *
 * IMPORTANT:
 * - Keep your existing Code.gs / other .gs files.
 * - If another file already contains doPost(e), remove or rename that old
 *   doPost so there is only ONE doPost(e) in the Apps Script project.
 * - After saving this file, update the EXISTING web-app deployment:
 *   Deploy -> Manage deployments -> Edit -> New version -> Deploy.
 */

const NEOCHRONO_RECEIVER = Object.freeze({
  SPREADSHEET_ID: '1flTBzOIM_dbDODjC-S-DHoViAhHASEyNWInZBxab5To',
  SHEET_NAME: 'Schedule',
  MESSAGE_TYPE: 'NEOCHRONO_SCHEDULE_RECEIVER'
});

function doPost(e) {
  var transferId = '';

  try {
    if (!e || !e.parameter || !e.parameter.payload) {
      throw new Error('Missing NeoChrono payload.');
    }

    var payload = JSON.parse(e.parameter.payload);
    transferId = String(payload.transferId || '').trim();

    if (String(payload.action || '') !== 'replaceSchedule') {
      throw new Error('Unsupported action.');
    }

    if (String(payload.sheetName || '') !== NEOCHRONO_RECEIVER.SHEET_NAME) {
      throw new Error(
        'NeoChrono attempted to write to an unexpected sheet. Expected "' +
        NEOCHRONO_RECEIVER.SHEET_NAME + '".'
      );
    }

    if (!Array.isArray(payload.headers) || !payload.headers.length) {
      throw new Error('Schedule headers are missing.');
    }

    if (!Array.isArray(payload.rows)) {
      throw new Error('Schedule rows are missing.');
    }

    var headers = payload.headers.map(function(value) {
      return value === null || value === undefined ? '' : String(value);
    });

    var width = headers.length;

    var rows = payload.rows.map(function(row, rowIndex) {
      if (!Array.isArray(row)) {
        throw new Error('Schedule row ' + (rowIndex + 1) + ' is invalid.');
      }

      var normalized = row.slice(0, width);

      while (normalized.length < width) {
        normalized.push('');
      }

      return normalized.map(function(value) {
        if (value === null || value === undefined) return '';
        return value;
      });
    });

    var ss = SpreadsheetApp.openById(
      NEOCHRONO_RECEIVER.SPREADSHEET_ID
    );

    var sheet = ss.getSheetByName(
      NEOCHRONO_RECEIVER.SHEET_NAME
    );

    if (!sheet) {
      sheet = ss.insertSheet(
        NEOCHRONO_RECEIVER.SHEET_NAME
      );
    }

    sheet.clearContents();

    var matrix = [headers].concat(rows);

    if (sheet.getMaxRows() < matrix.length) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        matrix.length - sheet.getMaxRows()
      );
    }

    if (sheet.getMaxColumns() < width) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        width - sheet.getMaxColumns()
      );
    }

    sheet
      .getRange(1, 1, matrix.length, width)
      .setValues(matrix);

    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();

    var writtenRows = Math.max(0, sheet.getLastRow() - 1);
    var writtenColumns = sheet.getLastColumn();

    if (writtenRows !== rows.length) {
      throw new Error(
        'Google verification failed. Expected ' +
        rows.length +
        ' data rows but found ' +
        writtenRows +
        ' in the Schedule sheet.'
      );
    }

    if (writtenColumns < width) {
      throw new Error(
        'Google verification failed. Expected at least ' +
        width +
        ' columns but found ' +
        writtenColumns +
        '.'
      );
    }

    var receiverUrl =
      ss.getUrl() +
      '#gid=' +
      sheet.getSheetId();

    return neoChronoReceiverResponse_({
      type: NEOCHRONO_RECEIVER.MESSAGE_TYPE,
      ok: true,
      transferId: transferId,
      receiverSpreadsheetId: ss.getId(),
      receiverSpreadsheetName: ss.getName(),
      receiverSheetName: sheet.getName(),
      receiverUrl: receiverUrl,
      writtenRows: writtenRows,
      writtenColumns: width,
      message:
        'Google confirmed that the Schedule sheet was replaced successfully.'
    });

  } catch (error) {
    return neoChronoReceiverResponse_({
      type: NEOCHRONO_RECEIVER.MESSAGE_TYPE,
      ok: false,
      transferId: transferId,
      receiverSheetName: NEOCHRONO_RECEIVER.SHEET_NAME,
      message:
        error && error.message
          ? error.message
          : String(error)
    });
  }
}

function neoChronoReceiverResponse_(result) {
  var json = JSON.stringify(result)
    .replace(/</g, '\\u003c');

  var html =
    '<!doctype html>' +
    '<html><head><meta charset="utf-8"></head><body>' +
    '<script>' +
    'window.parent.postMessage(' + json + ', "*");' +
    '</script>' +
    '<div style="font-family:Arial,sans-serif;padding:16px">' +
    (result.ok
      ? 'Schedule received successfully.'
      : 'Schedule transfer failed: ' +
        String(result.message || 'Unknown error')) +
    '</div>' +
    '</body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}
