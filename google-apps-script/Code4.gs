/**
 * Code4.gs — NeoChrono Pharmacists Schedule Receiver
 *
 * Destination spreadsheet:
 * 1flTBzOIM_dbDODjC-S-DHoViAhHASEyNWInZBxab5To
 *
 * Destination sheet:
 * Schedule
 *
 * Paste this file into the Apps Script project deployed at:
 * https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec
 *
 * Then:
 * Deploy -> Manage deployments -> Edit -> New version -> Deploy
 *
 * IMPORTANT:
 * This code supports BOTH:
 * 1) the current NeoChrono raw JSON POST, and
 * 2) the older form-encoded payload POST.
 */

const NEOCHRONO_SCHEDULE_RECEIVER_V5 = Object.freeze({
  SPREADSHEET_ID: '1flTBzOIM_dbDODjC-S-DHoViAhHASEyNWInZBxab5To',
  SHEET_NAME: 'Schedule',
  ACTION: 'replaceSchedule',
  LAST_RECEIVED_AT: 'NEOCHRONO_LAST_RECEIVED_AT',
  LAST_SENT_AT: 'NEOCHRONO_LAST_SENT_AT',
  LAST_ROW_COUNT: 'NEOCHRONO_LAST_ROW_COUNT'
});


function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    var payload = neoChronoReadIncomingPayload_(e);

    if (!payload) {
      throw new Error('No schedule payload was received.');
    }

    var action = String(payload.action || '').trim();

    if (
      action &&
      action !== NEOCHRONO_SCHEDULE_RECEIVER_V5.ACTION
    ) {
      throw new Error(
        'Unsupported receiver action: ' + action
      );
    }

    var requestedSheet = String(
      payload.sheet ||
      NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME
    ).trim();

    if (
      requestedSheet !==
      NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME
    ) {
      throw new Error(
        'Destination must be the sheet named "' +
        NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME +
        '".'
      );
    }

    var values = neoChronoPayloadToValues_(payload);

    if (!values.length) {
      throw new Error('No schedule rows were supplied.');
    }

    if (
      !Array.isArray(values[0]) ||
      !values[0].length
    ) {
      throw new Error('The schedule header row is missing.');
    }

    var columnCount = values[0].length;

    for (var r = 0; r < values.length; r++) {
      if (!Array.isArray(values[r])) {
        throw new Error(
          'Schedule row ' + (r + 1) + ' is invalid.'
        );
      }

      while (values[r].length < columnCount) {
        values[r].push('');
      }

      if (values[r].length > columnCount) {
        values[r] = values[r].slice(
          0,
          columnCount
        );
      }
    }

    values = neoChronoNormalizeScheduleValues_(
      values
    );

    var ss = SpreadsheetApp.openById(
      NEOCHRONO_SCHEDULE_RECEIVER_V5.SPREADSHEET_ID
    );

    var sheet = ss.getSheetByName(
      NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME
    );

    if (!sheet) {
      sheet = ss.insertSheet(
        NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME
      );
    }

    neoChronoEnsureSheetSize_(
      sheet,
      values.length,
      columnCount
    );

    // Replace Schedule completely.
    sheet.clearContents();

    sheet
      .getRange(
        1,
        1,
        values.length,
        columnCount
      )
      .setValues(values);

    sheet.setFrozenRows(1);

    try {
      sheet
        .getRange(1, 1, 1, columnCount)
        .setFontWeight('bold')
        .setBackground('#1f4e78')
        .setFontColor('#ffffff');
    } catch (ignore) {}

    try {
      sheet.autoResizeColumns(
        1,
        columnCount
      );
    } catch (ignore) {}

    SpreadsheetApp.flush();

    var actualRows = sheet.getLastRow();
    var actualColumns = sheet.getLastColumn();

    if (actualRows !== values.length) {
      throw new Error(
        'Write verification failed. Expected ' +
        values.length +
        ' total row(s), but Schedule contains ' +
        actualRows + '.'
      );
    }

    if (actualColumns < columnCount) {
      throw new Error(
        'Write verification failed. Expected ' +
        columnCount +
        ' column(s), but Schedule contains ' +
        actualColumns + '.'
      );
    }

    var dataRows = Math.max(
      0,
      values.length - 1
    );

    var now = new Date();
    var props =
      PropertiesService.getScriptProperties();

    props.setProperties({
      [NEOCHRONO_SCHEDULE_RECEIVER_V5.LAST_RECEIVED_AT]:
        now.toISOString(),

      [NEOCHRONO_SCHEDULE_RECEIVER_V5.LAST_SENT_AT]:
        String(payload.sentAt || ''),

      [NEOCHRONO_SCHEDULE_RECEIVER_V5.LAST_ROW_COUNT]:
        String(dataRows)
    });

    return neoChronoJsonReply_({
      success: true,
      ok: true,
      sheet:
        NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME,
      dataRows: dataRows,
      totalRows: values.length,
      columns: columnCount,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      lastReceivedAt: now.toISOString(),
      sentAt: String(payload.sentAt || ''),
      message:
        'Schedule was written successfully to the Google Sheet tab named Schedule.'
    });

  } catch (error) {

    return neoChronoJsonReply_({
      success: false,
      ok: false,
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


/**
 * Reads BOTH the new raw-JSON request and the old form-style request.
 */
function neoChronoReadIncomingPayload_(e) {
  // New NeoChrono sender:
  // Content-Type: text/plain
  // Body: JSON
  if (
    e &&
    e.postData &&
    e.postData.contents
  ) {
    var raw = String(
      e.postData.contents || ''
    ).trim();

    if (raw) {
      try {
        return JSON.parse(raw);
      } catch (ignore) {
        // Continue to legacy form parser.
      }
    }
  }

  // Legacy sender.
  if (e && e.parameter) {
    var action = String(
      e.parameter.action || ''
    ).trim();

    var payloadText = String(
      e.parameter.payload || ''
    ).trim();

    if (payloadText) {
      var legacy = JSON.parse(
        payloadText
      );

      if (!legacy.action) {
        legacy.action = action;
      }

      if (!legacy.sentAt) {
        legacy.sentAt =
          String(
            e.parameter.sentAt || ''
          );
      }

      return legacy;
    }

    if (action) {
      return {
        action: action,
        sheet:
          String(
            e.parameter.sheet || ''
          ),
        sentAt:
          String(
            e.parameter.sentAt || ''
          )
      };
    }
  }

  return null;
}


/**
 * Accepts either:
 *
 * New format:
 * {
 *   headers: [...],
 *   rows: [[...],[...]]
 * }
 *
 * Old format:
 * {
 *   values: [[header...],[row...]]
 * }
 */
function neoChronoPayloadToValues_(payload) {
  if (
    Array.isArray(payload.values) &&
    payload.values.length
  ) {
    return payload.values;
  }

  var headers =
    Array.isArray(payload.headers)
      ? payload.headers
      : [];

  var rows =
    Array.isArray(payload.rows)
      ? payload.rows
      : [];

  if (!headers.length) {
    throw new Error(
      'No schedule headers were received.'
    );
  }

  return [headers].concat(rows);
}


function doGet(e) {
  try {
    var action = String(
      e &&
      e.parameter &&
      e.parameter.action
        ? e.parameter.action
        : ''
    ).trim();

    if (action === 'status') {
      return neoChronoReceiverStatus_(e);
    }

    return neoChronoReceiverStatusPage_();

  } catch (error) {
    return HtmlService
      .createHtmlOutput(
        '<h2>Pharmacists Schedule Receiver</h2>' +
        '<p style="color:#b91c1c">' +
        neoChronoEscapeHtml_(
          error && error.message
            ? error.message
            : String(error)
        ) +
        '</p>'
      )
      .setTitle(
        'Pharmacists Schedule Receiver'
      );
  }
}


function neoChronoReceiverStatus_(e) {
  var ss = SpreadsheetApp.openById(
    NEOCHRONO_SCHEDULE_RECEIVER_V5.SPREADSHEET_ID
  );

  var sheet = ss.getSheetByName(
    NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME
  );

  var props =
    PropertiesService.getScriptProperties();

  var payload = {
    success: true,
    sheet:
      NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME,

    rows:
      sheet
        ? sheet.getLastRow()
        : 0,

    dataRows:
      sheet
        ? Math.max(
            0,
            sheet.getLastRow() - 1
          )
        : 0,

    columns:
      sheet
        ? sheet.getLastColumn()
        : 0,

    lastReceivedAt:
      props.getProperty(
        NEOCHRONO_SCHEDULE_RECEIVER_V5.LAST_RECEIVED_AT
      ) || '',

    lastSentAt:
      props.getProperty(
        NEOCHRONO_SCHEDULE_RECEIVER_V5.LAST_SENT_AT
      ) || '',

    lastRowCount:
      Number(
        props.getProperty(
          NEOCHRONO_SCHEDULE_RECEIVER_V5.LAST_ROW_COUNT
        ) || 0
      ),

    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl()
  };

  var callback = String(
    e &&
    e.parameter &&
    e.parameter.callback
      ? e.parameter.callback
      : ''
  ).trim();

  if (
    callback &&
    /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(
      callback
    )
  ) {
    return ContentService
      .createTextOutput(
        callback +
        '(' +
        JSON.stringify(payload) +
        ');'
      )
      .setMimeType(
        ContentService.MimeType.JAVASCRIPT
      );
  }

  return neoChronoJsonReply_(
    payload
  );
}


function neoChronoReceiverStatusPage_() {
  var ss = SpreadsheetApp.openById(
    NEOCHRONO_SCHEDULE_RECEIVER_V5.SPREADSHEET_ID
  );

  var sheet = ss.getSheetByName(
    NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME
  );

  var props =
    PropertiesService.getScriptProperties();

  var dataRows =
    sheet
      ? Math.max(
          0,
          sheet.getLastRow() - 1
        )
      : 0;

  var columns =
    sheet
      ? sheet.getLastColumn()
      : 0;

  var lastReceivedAt =
    props.getProperty(
      NEOCHRONO_SCHEDULE_RECEIVER_V5.LAST_RECEIVED_AT
    ) || 'Never';

  var html =
    '<div style="' +
      'font-family:Arial,sans-serif;' +
      'max-width:560px;' +
      'margin:40px auto;' +
      'padding:24px;' +
      'border:1px solid #ddd;' +
      'border-radius:16px;' +
    '">' +

      '<h2 style="margin-top:0">' +
        'Pharmacists Schedule Receiver' +
      '</h2>' +

      '<p><b>Destination sheet:</b> Schedule</p>' +

      '<p><b>Schedule rows:</b> ' +
        neoChronoEscapeHtml_(
          String(dataRows)
        ) +
      '</p>' +

      '<p><b>Columns:</b> ' +
        neoChronoEscapeHtml_(
          String(columns)
        ) +
      '</p>' +

      '<p><b>Last successful upload:</b> ' +
        neoChronoEscapeHtml_(
          String(lastReceivedAt)
        ) +
      '</p>' +

      '<p><a href="' +
        neoChronoEscapeHtml_(
          ss.getUrl()
        ) +
        '" target="_blank">' +
        'Open Google Sheet' +
      '</a></p>' +

    '</div>';

  return HtmlService
    .createHtmlOutput(html)
    .setTitle(
      'Pharmacists Schedule Receiver'
    );
}


function neoChronoEnsureSheetSize_(
  sheet,
  rows,
  columns
) {
  var maxRows =
    sheet.getMaxRows();

  var maxColumns =
    sheet.getMaxColumns();

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


function neoChronoNormalizeScheduleValues_(
  values
) {
  if (!values.length) {
    return values;
  }

  var headers =
    values[0].map(function(value) {
      return String(
        value || ''
      ).trim();
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

  for (
    var r = 1;
    r < values.length;
    r++
  ) {
    var row = [];

    for (
      var c = 0;
      c < headers.length;
      c++
    ) {
      var header =
        headers[c];

      var value =
        values[r][c];

      if (
        value !== '' &&
        value !== null &&
        value !== undefined &&
        (
          dateOnlyHeaders[header] ||
          dateTimeHeaders[header]
        )
      ) {
        var parsed =
          new Date(value);

        if (
          !isNaN(
            parsed.getTime()
          )
        ) {
          value = parsed;
        }
      }

      if (
        value === null ||
        value === undefined
      ) {
        value = '';
      }

      row.push(value);
    }

    output.push(row);
  }

  return output;
}


function neoChronoJsonReply_(payload) {
  return ContentService
    .createTextOutput(
      JSON.stringify(payload)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


function neoChronoEscapeHtml_(value) {
  return String(
    value == null
      ? ''
      : value
  )
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}


/**
 * Manual test from Apps Script editor.
 */
function testNeoChronoScheduleReceiverV5() {
  var ss = SpreadsheetApp.openById(
    NEOCHRONO_SCHEDULE_RECEIVER_V5.SPREADSHEET_ID
  );

  var sheet = ss.getSheetByName(
    NEOCHRONO_SCHEDULE_RECEIVER_V5.SHEET_NAME
  );

  return {
    spreadsheetId:
      ss.getId(),

    spreadsheetName:
      ss.getName(),

    scheduleSheetExists:
      !!sheet,

    scheduleSheetId:
      sheet
        ? sheet.getSheetId()
        : null,

    dataRows:
      sheet
        ? Math.max(
            0,
            sheet.getLastRow() - 1
          )
        : 0,

    url:
      ss.getUrl()
  };
}
