/**
 * Code4.gs — NeoChrono pharmacist schedule receiver
 *
 * Add this file to the Google Apps Script project deployed at:
 * https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec
 *
 * This does NOT replace your existing doGet() / pharmacist-facing web app.
 * It only adds doPost(e) so NeoChrono can replace the sheet named "Schedule".
 */

const NEOCHRONO_RECEIVER = Object.freeze({
  SHEET_NAME: 'Schedule',
  SPREADSHEET_ID_PROPERTY: 'NEOCHRONO_RECEIVER_SPREADSHEET_ID',
  LAST_PUBLISH_ID_PROPERTY: 'NEOCHRONO_LAST_PUBLISH_ID'
});

/**
 * Run this ONCE manually from the bound Apps Script editor.
 * It remembers which spreadsheet this receiver belongs to and makes sure the
 * "Schedule" sheet exists.
 */
function setupNeoChronoScheduleReceiver() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'This Apps Script project must be bound to the pharmacists spreadsheet. Open the spreadsheet, then Extensions → Apps Script, and run setupNeoChronoScheduleReceiver() there.'
    );
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      NEOCHRONO_RECEIVER.SPREADSHEET_ID_PROPERTY,
      ss.getId()
    );

  let sh = ss.getSheetByName(NEOCHRONO_RECEIVER.SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(NEOCHRONO_RECEIVER.SHEET_NAME);
  }

  sh.setFrozenRows(1);

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheetName: sh.getName()
  };
}


/**
 * Called automatically by NeoChrono's Finalize & Send button.
 */
function doPost(e) {
  let publishId = '';

  try {
    const action = String(
      e && e.parameter && e.parameter.action
        ? e.parameter.action
        : ''
    ).trim();

    publishId = String(
      e && e.parameter && e.parameter.publishId
        ? e.parameter.publishId
        : ''
    ).trim();

    if (action !== 'replaceSchedule') {
      throw new Error('Unsupported receiver action.');
    }

    const raw = String(
      e && e.parameter && e.parameter.payload
        ? e.parameter.payload
        : ''
    );

    if (!raw) {
      throw new Error('Schedule payload was empty.');
    }

    const payload = JSON.parse(raw);

    if (String(payload.sheet || '') !== NEOCHRONO_RECEIVER.SHEET_NAME) {
      throw new Error(
        'NeoChrono may publish only to the sheet named "' +
        NEOCHRONO_RECEIVER.SHEET_NAME +
        '".'
      );
    }

    publishId = String(payload.publishId || publishId || '').trim();

    if (!publishId) {
      throw new Error('Publish ID is missing.');
    }

    const result = neoChronoReplaceSchedule_(payload);

    return neoChronoReceiverHtmlResponse_({
      type: 'NEOCHRONO_SCHEDULE_PUBLISH_RESULT',
      ok: true,
      publishId: publishId,
      transferredRows: result.transferredRows,
      transferredColumns: result.transferredColumns,
      sheetName: result.sheetName,
      spreadsheetName: result.spreadsheetName,
      message:
        'Schedule was uploaded successfully to the Google Sheet named Schedule.'
    });

  } catch (err) {
    return neoChronoReceiverHtmlResponse_({
      type: 'NEOCHRONO_SCHEDULE_PUBLISH_RESULT',
      ok: false,
      publishId: publishId,
      message:
        err && err.message
          ? err.message
          : String(err)
    });
  }
}


function neoChronoReplaceSchedule_(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = neoChronoReceiverSpreadsheet_();
    let sh = ss.getSheetByName(NEOCHRONO_RECEIVER.SHEET_NAME);

    if (!sh) {
      sh = ss.insertSheet(NEOCHRONO_RECEIVER.SHEET_NAME);
    }

    const values = payload.values;

    if (!Array.isArray(values) || !values.length) {
      throw new Error('Schedule payload does not contain rows.');
    }

    const headers = values[0].map(function(v) {
      return String(v == null ? '' : v).trim();
    });

    if (!headers.length) {
      throw new Error('Schedule header row is empty.');
    }

    const requiredHeaders = [
      'Date',
      'Shift',
      'Assigned Pharmacist',
      'Status'
    ];

    requiredHeaders.forEach(function(header) {
      if (headers.indexOf(header) < 0) {
        throw new Error(
          'Schedule payload is missing required column: ' + header
        );
      }
    });

    const normalized = values.map(function(row, rowIndex) {
      return headers.map(function(header, colIndex) {
        return neoChronoReceiverValue_(
          header,
          row && row[colIndex],
          rowIndex
        );
      });
    });

    // Idempotency: if a browser repeats the same POST because of a redirect or
    // network retry, do not rewrite the sheet twice.
    const props = PropertiesService.getScriptProperties();
    const lastPublishId = props.getProperty(
      NEOCHRONO_RECEIVER.LAST_PUBLISH_ID_PROPERTY
    );

    if (
      lastPublishId &&
      lastPublishId === String(payload.publishId || '')
    ) {
      return {
        transferredRows: Math.max(0, normalized.length - 1),
        transferredColumns: headers.length,
        sheetName: sh.getName(),
        spreadsheetName: ss.getName()
      };
    }

    neoChronoEnsureSheetSize_(
      sh,
      normalized.length,
      headers.length
    );

    // The receiver sheet must contain only the newly published schedule.
    sh.clearContents();

    sh
      .getRange(
        1,
        1,
        normalized.length,
        headers.length
      )
      .setValues(normalized);

    sh.setFrozenRows(1);

    const dateCol = headers.indexOf('Date') + 1;

    if (dateCol > 0 && normalized.length > 1) {
      sh
        .getRange(
          2,
          dateCol,
          normalized.length - 1,
          1
        )
        .setNumberFormat('m/d/yyyy');
    }

    const finalizedCol = headers.indexOf('Finalized At') + 1;

    if (finalizedCol > 0 && normalized.length > 1) {
      sh
        .getRange(
          2,
          finalizedCol,
          normalized.length - 1,
          1
        )
        .setNumberFormat('m/d/yyyy h:mm AM/PM');
    }

    // Simple header formatting without changing the existing spreadsheet theme.
    sh
      .getRange(1,1,1,headers.length)
      .setFontWeight('bold');

    SpreadsheetApp.flush();

    props.setProperty(
      NEOCHRONO_RECEIVER.LAST_PUBLISH_ID_PROPERTY,
      String(payload.publishId || '')
    );

    props.setProperty(
      'NEOCHRONO_LAST_PUBLISHED_AT',
      new Date().toISOString()
    );

    props.setProperty(
      'NEOCHRONO_LAST_PUBLISHED_RANGE',
      String(payload.startDate || '') +
      ' through ' +
      String(payload.endDate || '')
    );

    return {
      transferredRows: Math.max(0, normalized.length - 1),
      transferredColumns: headers.length,
      sheetName: sh.getName(),
      spreadsheetName: ss.getName()
    };

  } finally {
    lock.releaseLock();
  }
}


function neoChronoReceiverSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();

  if (active) {
    return active;
  }

  const id = PropertiesService
    .getScriptProperties()
    .getProperty(
      NEOCHRONO_RECEIVER.SPREADSHEET_ID_PROPERTY
    );

  if (!id) {
    throw new Error(
      'Receiver spreadsheet is not configured. Run setupNeoChronoScheduleReceiver() once from the bound Apps Script editor, then deploy a new web-app version.'
    );
  }

  return SpreadsheetApp.openById(id);
}


function neoChronoEnsureSheetSize_(sheet, rows, cols) {
  const needRows = Math.max(1, Number(rows || 1));
  const needCols = Math.max(1, Number(cols || 1));

  if (sheet.getMaxRows() < needRows) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      needRows - sheet.getMaxRows()
    );
  }

  if (sheet.getMaxColumns() < needCols) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      needCols - sheet.getMaxColumns()
    );
  }
}


function neoChronoReceiverValue_(header, value, rowIndex) {
  if (rowIndex === 0) {
    return String(value == null ? '' : value);
  }

  if (value === null || value === undefined || value === '') {
    return '';
  }

  const dateOnlyHeaders = {
    'Date': true,
    'Start Date': true,
    'End Date': true,
    'Weekend Saturday': true,
    'Weekend Sunday': true,
    'Effective Start': true,
    'Effective End': true
  };

  const dateTimeHeaders = {
    'Updated At': true,
    'Finalized At': true,
    'Submitted At': true,
    'Reviewed At': true
  };

  if (dateOnlyHeaders[header]) {
    const d = neoChronoReceiverDate_(value, true);
    return d || value;
  }

  if (dateTimeHeaders[header]) {
    const d = neoChronoReceiverDate_(value, false);
    return d || value;
  }

  return value;
}


function neoChronoReceiverDate_(value, dateOnly) {
  if (value instanceof Date) {
    return value;
  }

  const text = String(value || '').trim();

  if (!text) {
    return '';
  }

  if (
    dateOnly &&
    /^\d{4}-\d{2}-\d{2}$/.test(text)
  ) {
    const parts = text.split('-').map(Number);
    const d = new Date(
      parts[0],
      parts[1] - 1,
      parts[2],
      12,0,0,0
    );

    return isNaN(d.getTime()) ? '' : d;
  }

  const d = new Date(text);

  return isNaN(d.getTime()) ? '' : d;
}


/**
 * Returns a tiny HTML page inside NeoChrono's hidden iframe.
 * postMessage works cross-origin, so NeoChrono receives a verified success or
 * failure without relying on fetch/CORS.
 */
function neoChronoReceiverHtmlResponse_(payload) {
  const json = JSON
    .stringify(payload || {})
    .replace(/</g,'\\u003c');

  return HtmlService
    .createHtmlOutput(
      '<!doctype html><html><body>' +
      '<script>' +
      'try{' +
      'window.parent.postMessage(' + json + ', "*");' +
      '}catch(e){}' +
      '<\/script>' +
      '</body></html>'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}
