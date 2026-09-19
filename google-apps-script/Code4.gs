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

// Destination is the Google Spreadsheet this Apps Script project is bound to.
// The schedule is written only to the tab named exactly "Schedule".
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

    var timezone =
      ss.getSpreadsheetTimeZone() ||
      Session.getScriptTimeZone() ||
      'America/New_York';

    var cleanRows = payload.rows.map(function(row) {
      var source = Array.isArray(row) ? row : [];

      return NEOCHRONO_SCHEDULE_HEADERS.map(function(_header, index) {
        var value = source[index];

        if (value === null || value === undefined) {
          return '';
        }

        return value;
      });
    });

    /*
     * Determine exactly which date range NeoChrono is replacing.
     *
     * Example:
     *   Existing Google data = September + October
     *   Incoming range        = October 1 - October 31
     *
     * Result:
     *   Keep every September row.
     *   Remove the old October rows.
     *   Add the newly finalized October rows.
     */
    var startKey =
      neoChronoScheduleDateKey_(
        payload.startDate,
        timezone
      );

    var endKey =
      neoChronoScheduleDateKey_(
        payload.endDate,
        timezone
      );

    if (!startKey || !endKey) {
      var incomingDateKeys = cleanRows
        .map(function(row) {
          return neoChronoScheduleDateKey_(
            row[2],
            timezone
          );
        })
        .filter(function(key) {
          return !!key;
        })
        .sort();

      if (!startKey && incomingDateKeys.length) {
        startKey = incomingDateKeys[0];
      }

      if (!endKey && incomingDateKeys.length) {
        endKey = incomingDateKeys[incomingDateKeys.length - 1];
      }
    }

    if (!startKey || !endKey) {
      throw new Error(
        'The schedule transfer is missing a valid Start Date / End Date.'
      );
    }

    if (startKey > endKey) {
      throw new Error(
        'Schedule Start Date must be before or equal to End Date.'
      );
    }

    /*
     * Validate every incoming row belongs to the selected finalized range.
     */
    cleanRows.forEach(function(row, index) {
      var rowDate =
        neoChronoScheduleDateKey_(
          row[2],
          timezone
        );

      if (!rowDate) {
        throw new Error(
          'Incoming schedule row ' +
          (index + 1) +
          ' is missing a valid Date.'
        );
      }

      if (
        rowDate < startKey ||
        rowDate > endKey
      ) {
        throw new Error(
          'Incoming schedule row ' +
          (index + 1) +
          ' has date ' +
          rowDate +
          ', which is outside the finalized range ' +
          startKey +
          ' through ' +
          endKey +
          '.'
        );
      }
    });

    /*
     * Read the existing Schedule sheet BEFORE writing anything.
     */
    var existingRows = [];
    var oldLastRow = sheet.getLastRow();
    var oldLastColumn = sheet.getLastColumn();

    if (oldLastRow >= 1 && oldLastColumn > 0) {
      var existingHeaderWidth =
        Math.min(
          oldLastColumn,
          NEOCHRONO_SCHEDULE_HEADERS.length
        );

      var existingHeaders =
        sheet
          .getRange(
            1,
            1,
            1,
            existingHeaderWidth
          )
          .getValues()[0]
          .map(function(value) {
            return String(value || '').trim();
          });

      var hasExistingHeader =
        existingHeaders.some(function(value) {
          return !!value;
        });

      if (
        hasExistingHeader &&
        (
          existingHeaderWidth !== NEOCHRONO_SCHEDULE_HEADERS.length ||
          !neoChronoHeadersMatch_(
            existingHeaders,
            NEOCHRONO_SCHEDULE_HEADERS
          )
        )
      ) {
        throw new Error(
          'The existing Schedule sheet headers do not match NeoChrono. ' +
          'The receiver stopped before changing any schedule data.'
        );
      }

      if (oldLastRow > 1) {
        existingRows =
          sheet
            .getRange(
              2,
              1,
              oldLastRow - 1,
              NEOCHRONO_SCHEDULE_HEADERS.length
            )
            .getValues()
            .filter(function(row) {
              return row.some(function(value) {
                return (
                  value !== '' &&
                  value !== null &&
                  value !== undefined
                );
              });
            });
      }
    }

    /*
     * Preserve ALL existing rows outside the newly finalized range.
     *
     * Rows without a usable Date are also preserved so the receiver never
     * silently destroys old/manual information.
     */
    var preservedRows =
      existingRows.filter(function(row) {
        var dateKey =
          neoChronoScheduleDateKey_(
            row[2],
            timezone
          );

        if (!dateKey) {
          return true;
        }

        return (
          dateKey < startKey ||
          dateKey > endKey
        );
      });

    var replacedExistingRows =
      existingRows.length -
      preservedRows.length;

    var mergedRows =
      preservedRows.concat(cleanRows);

    /*
     * Keep the full historical Schedule sheet in chronological order.
     * Sorting is stable for rows on the same date.
     */
    mergedRows =
      neoChronoSortScheduleRows_(
        mergedRows,
        timezone
      );

    /*
     * Make sure the sheet is large enough.
     */
    var requiredRows =
      Math.max(
        1,
        mergedRows.length + 1
      );

    if (
      sheet.getMaxRows() <
      requiredRows
    ) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        requiredRows -
        sheet.getMaxRows()
      );
    }

    if (
      sheet.getMaxColumns() <
      NEOCHRONO_SCHEDULE_HEADERS.length
    ) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        NEOCHRONO_SCHEDULE_HEADERS.length -
        sheet.getMaxColumns()
      );
    }

    /*
     * Write header + the COMPLETE merged data set.
     *
     * Do NOT clear the whole Schedule tab.
     */
    sheet
      .getRange(
        1,
        1,
        1,
        NEOCHRONO_SCHEDULE_HEADERS.length
      )
      .setValues([
        NEOCHRONO_SCHEDULE_HEADERS.slice()
      ]);

    if (mergedRows.length) {
      sheet
        .getRange(
          2,
          1,
          mergedRows.length,
          NEOCHRONO_SCHEDULE_HEADERS.length
        )
        .setValues(
          mergedRows
        );
    }

    /*
     * If the new merged result is shorter than the old sheet, clear only
     * leftover trailing rows. Historical rows already included in mergedRows
     * are never deleted.
     */
    var finalLastRow =
      mergedRows.length + 1;

    if (
      oldLastRow >
      finalLastRow
    ) {
      sheet
        .getRange(
          finalLastRow + 1,
          1,
          oldLastRow - finalLastRow,
          NEOCHRONO_SCHEDULE_HEADERS.length
        )
        .clearContent();
    }

    sheet.setFrozenRows(1);

    sheet
      .getRange(
        1,
        1,
        1,
        NEOCHRONO_SCHEDULE_HEADERS.length
      )
      .setFontWeight(
        'bold'
      );

    if (mergedRows.length) {
      sheet
        .getRange(
          2,
          3,
          mergedRows.length,
          1
        )
        .setNumberFormat(
          'yyyy-mm-dd'
        );
    }

    SpreadsheetApp.flush();

    /*
     * Verify headers.
     */
    var verifyHeaders =
      sheet
        .getRange(
          1,
          1,
          1,
          NEOCHRONO_SCHEDULE_HEADERS.length
        )
        .getValues()[0]
        .map(function(value) {
          return String(value || '').trim();
        });

    if (
      !neoChronoHeadersMatch_(
        verifyHeaders,
        NEOCHRONO_SCHEDULE_HEADERS
      )
    ) {
      throw new Error(
        'Header verification failed after writing the Schedule sheet.'
      );
    }

    /*
     * Verify total row count.
     */
    var actualLastRow =
      sheet.getLastRow();

    if (
      actualLastRow !==
      finalLastRow
    ) {
      throw new Error(
        'Row verification failed. Expected ' +
        finalLastRow +
        ' total rows but the Schedule sheet has ' +
        actualLastRow +
        '.'
      );
    }

    /*
     * Verify that the finalized range contains exactly the rows NeoChrono sent.
     */
    var verifyRows =
      mergedRows.length
        ? sheet
            .getRange(
              2,
              1,
              mergedRows.length,
              NEOCHRONO_SCHEDULE_HEADERS.length
            )
            .getValues()
        : [];

    var verifyIncomingCount =
      verifyRows.filter(function(row) {
        var dateKey =
          neoChronoScheduleDateKey_(
            row[2],
            timezone
          );

        return (
          dateKey &&
          dateKey >= startKey &&
          dateKey <= endKey
        );
      }).length;

    if (
      verifyIncomingCount !==
      cleanRows.length
    ) {
      throw new Error(
        'Date-range verification failed. NeoChrono sent ' +
        cleanRows.length +
        ' row(s) for ' +
        startKey +
        ' through ' +
        endKey +
        ', but Google contains ' +
        verifyIncomingCount +
        ' row(s) in that range.'
      );
    }

    return neoChronoReceiverResponse_({
      type:
        'NEOCHRONO_SCHEDULE_RECEIVER',

      ok:
        true,

      transferId:
        transferId,

      /*
       * Keep transferredRows equal to ONLY the rows sent in this transfer.
       * The NeoChrono browser uses this number to verify the submission.
       */
      transferredRows:
        cleanRows.length,

      writtenRows:
        cleanRows.length,

      transferredColumns:
        NEOCHRONO_SCHEDULE_HEADERS.length,

      writtenColumns:
        NEOCHRONO_SCHEDULE_HEADERS.length,

      preservedRows:
        preservedRows.length,

      replacedExistingRows:
        replacedExistingRows,

      totalScheduleRows:
        mergedRows.length,

      startDate:
        startKey,

      endDate:
        endKey,

      receiverSpreadsheetId:
        ss.getId(),

      receiverSpreadsheetName:
        ss.getName(),

      receiverSheetName:
        NEOCHRONO_RECEIVER_SHEET_NAME,

      receiverUrl:
        ss.getUrl() +
        '#gid=' +
        sheet.getSheetId(),

      message:
        'Schedule merged successfully. ' +
        cleanRows.length +
        ' row(s) were written for ' +
        startKey +
        ' through ' +
        endKey +
        '. ' +
        preservedRows.length +
        ' existing row(s) outside that date range were preserved. ' +
        mergedRows.length +
        ' total schedule row(s) are now stored.'
    });

  } catch (error) {

    return neoChronoReceiverResponse_({
      type:
        'NEOCHRONO_SCHEDULE_RECEIVER',

      ok:
        false,

      transferId:
        transferId,

      message:
        error && error.message
          ? error.message
          : String(error)
    });
  }
}


/**
 * Converts a Schedule date to YYYY-MM-DD without changing the calendar date.
 */
function neoChronoScheduleDateKey_(value, timezone) {
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return '';
  }

  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {
    return Utilities.formatDate(
      value,
      timezone ||
      Session.getScriptTimeZone() ||
      'America/New_York',
      'yyyy-MM-dd'
    );
  }

  var raw =
    String(value).trim();

  var iso =
    raw.match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );

  if (iso) {
    return (
      iso[1] +
      '-' +
      iso[2] +
      '-' +
      iso[3]
    );
  }

  var parsed =
    new Date(raw);

  if (
    isNaN(parsed.getTime())
  ) {
    return '';
  }

  return Utilities.formatDate(
    parsed,
    timezone ||
    Session.getScriptTimeZone() ||
    'America/New_York',
    'yyyy-MM-dd'
  );
}


/**
 * Sort the combined historical schedule by Date.
 * Rows with no valid date are kept at the end and never discarded.
 */
function neoChronoSortScheduleRows_(rows, timezone) {
  return (rows || [])
    .map(function(row, index) {
      return {
        row:
          row,

        index:
          index,

        date:
          neoChronoScheduleDateKey_(
            row[2],
            timezone
          )
      };
    })
    .sort(function(a, b) {
      if (a.date && b.date) {
        if (a.date < b.date) return -1;
        if (a.date > b.date) return 1;
      } else if (a.date) {
        return -1;
      } else if (b.date) {
        return 1;
      }

      return (
        a.index -
        b.index
      );
    })
    .map(function(item) {
      return item.row;
    });
}

function neoChronoReceiverSpreadsheet_() {
  // If you intentionally configure an ID above, use it.
  if (String(NEOCHRONO_RECEIVER_SPREADSHEET_ID || '').trim()) {
    return SpreadsheetApp.openById(
      String(NEOCHRONO_RECEIVER_SPREADSHEET_ID).trim()
    );
  }

  // Normal setup: Code4.gs lives inside the pharmacists spreadsheet's
  // Extensions -> Apps Script project.
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Code4.gs is not running from a spreadsheet-bound Apps Script project. ' +
      'Open the pharmacists Google Sheet, choose Extensions > Apps Script, ' +
      'add Code4.gs there, then redeploy the existing web app. ' +
      'If the project must remain standalone, set NEOCHRONO_RECEIVER_SPREADSHEET_ID ' +
      'to the pharmacists spreadsheet ID.'
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
