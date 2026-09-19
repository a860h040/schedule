/**
 * NeoChrono Pharmacists Schedule Receiver
 * File name: Code4.gs
 *
 * Receives the finalized schedule from the GitHub NeoChrono app and replaces
 * the contents of the Google Sheet tab named exactly "Schedule".
 *
 * IMPORTANT:
 * - Add this file to the Apps Script project deployed at:
 *   https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec
 * - Then deploy a NEW VERSION of the EXISTING web app deployment.
 * - If another file already contains function doPost(e), do not keep two
 *   doPost functions. Route that existing doPost to neoChronoReceiveSchedule_(e).
 */

const NEOCHRONO_SCHEDULE_RECEIVER_V4 = Object.freeze({
  SPREADSHEET_ID: '1flTBzOIM_dbDODjC-S-DHoViAhHASEyNWInZBxab5To',
  SHEET_NAME: 'Schedule',
  ACTION: 'replaceSchedule'
});


function doPost(e) {
  return neoChronoReceiveSchedule_(e);
}


function neoChronoReceiveSchedule_(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!e || !e.parameter) {
      throw new Error('No POST data was received.');
    }

    var action = String(e.parameter.action || '').trim();
    if (action !== NEOCHRONO_SCHEDULE_RECEIVER_V4.ACTION) {
      throw new Error('Unsupported receiver action: ' + action);
    }

    var publishId = String(e.parameter.publishId || '').trim();
    var payloadText = String(e.parameter.payload || '').trim();

    if (!publishId) {
      throw new Error('Publish ID is missing.');
    }

    if (!payloadText) {
      throw new Error('Schedule payload is empty.');
    }

    var payload;
    try {
      payload = JSON.parse(payloadText);
    } catch (parseError) {
      throw new Error('The schedule payload is not valid JSON.');
    }

    if (
      String(payload.sheet || '') !==
      NEOCHRONO_SCHEDULE_RECEIVER_V4.SHEET_NAME
    ) {
      throw new Error(
        'NeoChrono attempted to publish to an unexpected sheet. Expected: ' +
        NEOCHRONO_SCHEDULE_RECEIVER_V4.SHEET_NAME
      );
    }

    var values = payload.values;

    if (!Array.isArray(values) || values.length < 1) {
      throw new Error('No schedule rows were supplied.');
    }

    if (!Array.isArray(values[0]) || values[0].length < 1) {
      throw new Error('The schedule header row is missing.');
    }

    var columnCount = values[0].length;

    for (var r = 0; r < values.length; r++) {
      if (!Array.isArray(values[r])) {
        throw new Error('Schedule row ' + (r + 1) + ' is invalid.');
      }

      while (values[r].length < columnCount) {
        values[r].push('');
      }

      if (values[r].length > columnCount) {
        values[r] = values[r].slice(0, columnCount);
      }
    }

    values = neoChronoNormalizeScheduleValues_(values);

    var ss = SpreadsheetApp.openById(
      NEOCHRONO_SCHEDULE_RECEIVER_V4.SPREADSHEET_ID
    );

    var sheet = ss.getSheetByName(
      NEOCHRONO_SCHEDULE_RECEIVER_V4.SHEET_NAME
    );

    if (!sheet) {
      sheet = ss.insertSheet(
        NEOCHRONO_SCHEDULE_RECEIVER_V4.SHEET_NAME
      );
    }

    neoChronoEnsureSheetSize_(
      sheet,
      values.length,
      columnCount
    );

    // Replace the existing Schedule data completely.
    sheet.clearContents();

    sheet
      .getRange(1, 1, values.length, columnCount)
      .setValues(values);

    sheet.setFrozenRows(1);

    try {
      sheet
        .getRange(1, 1, 1, columnCount)
        .setFontWeight('bold');
    } catch (ignore) {}

    SpreadsheetApp.flush();

    // Verify the write BEFORE telling NeoChrono it succeeded.
    var actualRows = sheet.getLastRow();
    var actualColumns = sheet.getLastColumn();

    if (actualRows !== values.length) {
      throw new Error(
        'Write verification failed. Expected ' +
        values.length +
        ' rows but Google Sheets reports ' +
        actualRows +
        '.'
      );
    }

    if (actualColumns < columnCount) {
      throw new Error(
        'Write verification failed. Expected at least ' +
        columnCount +
        ' columns but Google Sheets reports ' +
        actualColumns +
        '.'
      );
    }

    var properties = PropertiesService.getScriptProperties();
    properties.setProperty(
      'NEOCHRONO_LAST_PUBLISH_ID',
      publishId
    );
    properties.setProperty(
      'NEOCHRONO_LAST_PUBLISH_AT',
      new Date().toISOString()
    );
    properties.setProperty(
      'NEOCHRONO_LAST_PUBLISH_ROWS',
      String(Math.max(0, values.length - 1))
    );

    var receiverUrl =
      'https://docs.google.com/spreadsheets/d/' +
      NEOCHRONO_SCHEDULE_RECEIVER_V4.SPREADSHEET_ID +
      '/edit#gid=' +
      sheet.getSheetId();

    return neoChronoReceiverReply_({
      type: 'NEOCHRONO_SCHEDULE_PUBLISH_RESULT',
      ok: true,
      publishId: publishId,
      transferredRows: Math.max(0, values.length - 1),
      transferredColumns: columnCount,
      spreadsheetName: ss.getName(),
      sheetName: NEOCHRONO_SCHEDULE_RECEIVER_V4.SHEET_NAME,
      receiverUrl: receiverUrl,
      message:
        'Schedule was written successfully to the Google Sheet tab named Schedule.'
    });

  } catch (error) {
    var failedPublishId = '';

    try {
      failedPublishId = String(
        e && e.parameter
          ? e.parameter.publishId || ''
          : ''
      );
    } catch (ignore) {}

    return neoChronoReceiverReply_({
      type: 'NEOCHRONO_SCHEDULE_PUBLISH_RESULT',
      ok: false,
      publishId: failedPublishId,
      message:
        error && error.message
          ? error.message
          : String(error)
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


function neoChronoEnsureSheetSize_(sheet, rows, columns) {
  var maxRows = sheet.getMaxRows();
  var maxColumns = sheet.getMaxColumns();

  if (maxRows < rows) {
    sheet.insertRowsAfter(
      maxRows,
      rows - maxRows
    );
  }

  if (maxColumns < columns) {
    sheet.insertColumnsAfter(
      maxColumns,
      columns - maxColumns
    );
  }
}


function neoChronoNormalizeScheduleValues_(values) {
  if (!values.length) {
    return values;
  }

  var headers = values[0].map(function(value) {
    return String(value || '').trim();
  });

  var dateOnlyHeaders = {
    'Date': true,
    'Start Date': true,
    'End Date': true,
    'Weekend Saturday': true,
    'Weekend Sunday': true,
    'Effective Start': true,
    'Effective End': true
  };

  var dateTimeHeaders = {
    'Updated At': true,
    'Submitted At': true,
    'Reviewed At': true,
    'Finalized At': true,
    'Created At': true
  };

  var output = [headers];

  for (var r = 1; r < values.length; r++) {
    var row = [];

    for (var c = 0; c < headers.length; c++) {
      var header = headers[c];
      var value = values[r][c];

      if (
        value !== '' &&
        value !== null &&
        value !== undefined &&
        (
          dateOnlyHeaders[header] ||
          dateTimeHeaders[header]
        )
      ) {
        var parsed = new Date(value);

        if (!isNaN(parsed.getTime())) {
          value = parsed;
        }
      }

      if (value === null || value === undefined) {
        value = '';
      }

      row.push(value);
    }

    output.push(row);
  }

  return output;
}


function neoChronoReceiverReply_(result) {
  var json = JSON.stringify(result)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  var html =
    '<!doctype html>' +
    '<html><head><meta charset="utf-8"></head>' +
    '<body>' +
    '<script>' +
    'window.parent.postMessage(' +
    json +
    ',"*");' +
    '<\/script>' +
    '</body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


/**
 * Optional manual test from the Apps Script editor.
 * Running this confirms this project can access the target spreadsheet
 * and the tab named "Schedule".
 */
function testNeoChronoScheduleReceiverV4() {
  var ss = SpreadsheetApp.openById(
    NEOCHRONO_SCHEDULE_RECEIVER_V4.SPREADSHEET_ID
  );

  var sheet = ss.getSheetByName(
    NEOCHRONO_SCHEDULE_RECEIVER_V4.SHEET_NAME
  );

  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    scheduleSheetExists: !!sheet,
    scheduleSheetId: sheet
      ? sheet.getSheetId()
      : null,
    url: sheet
      ? (
          'https://docs.google.com/spreadsheets/d/' +
          ss.getId() +
          '/edit#gid=' +
          sheet.getSheetId()
        )
      : ss.getUrl()
  };
}
