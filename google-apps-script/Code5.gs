/**
 * Code5.gs
 * NeoChrono Google bridge
 *
 * One POST endpoint handles:
 *   1) Schedule replacement
 *   2) PTO read
 *   3) PTO save / review / delete
 *
 * IMPORTANT:
 * - Keep only ONE doPost(e) in the Apps Script project.
 * - This file does NOT define doGet(e), so your existing Code.gs doGet can stay.
 * - After saving this file, update the EXISTING web-app deployment:
 *   Deploy -> Manage deployments -> Edit -> New version -> Deploy.
 */

const NEOCHRONO_RECEIVER = Object.freeze({
  SPREADSHEET_ID: '1flTBzOIM_dbDODjC-S-DHoViAhHASEyNWInZBxab5To',
  SCHEDULE_SHEET: 'Schedule',
  PTO_SHEET: 'PTO / Availability Requests',
  PRN_AVAILABILITY_SHEET: 'My Availability',
  SCHEDULE_MESSAGE_TYPE: 'NEOCHRONO_SCHEDULE_RECEIVER',
  PTO_MESSAGE_TYPE: 'NEOCHRONO_PTO_RECEIVER',
  PRN_AVAILABILITY_MESSAGE_TYPE: 'NEOCHRONO_PRN_AVAILABILITY',
  PTO_AUTO_APPROVE_LIMIT: 2
});

const NEOCHRONO_PTO_HEADERS = Object.freeze([
  'Record Type',
  'Record ID',
  'Pharmacist',
  'Username',
  'Date',
  'Start Date',
  'End Date',
  'Weekend Saturday',
  'Weekend Sunday',
  'Available',
  'Status',
  'Comment',
  'Submitted At',
  'Reviewed By',
  'Reviewed At',
  'Updated At',
  'Updated By'
]);

function doPost(e) {
  var requestId = '';
  var transferId = '';

  try {
    var parsed = neoChronoParsePost_(e);
    var action = String(parsed.action || '').trim();
    var payload = parsed.payload || {};

    requestId = String(parsed.requestId || '').trim();
    transferId = String(
      payload.transferId ||
      parsed.transferId ||
      ''
    ).trim();

    if (action === 'replaceSchedule' || String(payload.action || '') === 'replaceSchedule') {
      return neoChronoHandleSchedule_(payload, transferId);
    }

    if (action === 'neochronoPto') {
      return neoChronoHandlePtoRead_(requestId, parsed.sheet);
    }

    // Used by the neochrono-data GitHub Action. Returns plain JSON rather
    // than iframe HTML so the workflow can import Google PTO into GitHub.
    if (action === 'neochronoPtoExport') {
      return neoChronoHandlePtoExport_(parsed.sheet);
    }

    if (action === 'neochronoPtoWrite') {
      return neoChronoHandlePtoWrite_(requestId, parsed.sheet, payload);
    }

    if (action === 'neochronoPrnAvailability') {
      return neoChronoHandlePrnAvailabilityRead_(requestId, parsed.sheet);
    }

    if (action === 'neochronoPrnAvailabilityExport') {
      return neoChronoHandlePrnAvailabilityExport_(parsed.sheet);
    }

    throw new Error('Unsupported NeoChrono action: ' + action);

  } catch (error) {
    var message = error && error.message
      ? error.message
      : String(error);

    if (requestId) {
      var isPrnAvailabilityAction =
        String(action || '') === 'neochronoPrnAvailability';

      return neoChronoBridgeResponse_({
        type: isPrnAvailabilityAction
          ? NEOCHRONO_RECEIVER.PRN_AVAILABILITY_MESSAGE_TYPE
          : NEOCHRONO_RECEIVER.PTO_MESSAGE_TYPE,
        ok: false,
        success: false,
        requestId: requestId,
        sheet: isPrnAvailabilityAction
          ? NEOCHRONO_RECEIVER.PRN_AVAILABILITY_SHEET
          : NEOCHRONO_RECEIVER.PTO_SHEET,
        message: message
      });
    }

    return neoChronoBridgeResponse_({
      type: NEOCHRONO_RECEIVER.SCHEDULE_MESSAGE_TYPE,
      ok: false,
      success: false,
      transferId: transferId,
      receiverSheetName: NEOCHRONO_RECEIVER.SCHEDULE_SHEET,
      message: message
    });
  }
}

function neoChronoParsePost_(e) {
  var params = e && e.parameter ? e.parameter : {};
  var rawBody = e && e.postData && e.postData.contents
    ? String(e.postData.contents)
    : '';

  var body = {};
  if (rawBody) {
    try {
      body = JSON.parse(rawBody);
    } catch (_ignored) {
      body = {};
    }
  }

  var payloadRaw =
    params.payload !== undefined
      ? params.payload
      : body.payload;

  var payload = {};
  if (typeof payloadRaw === 'string' && payloadRaw.trim()) {
    payload = JSON.parse(payloadRaw);
  } else if (payloadRaw && typeof payloadRaw === 'object') {
    payload = payloadRaw;
  } else if (body && typeof body === 'object' && body.action) {
    payload = body.payload || {};
  }

  return {
    action: String(params.action || body.action || payload.action || '').trim(),
    sheet: String(params.sheet || body.sheet || payload.sheetName || '').trim(),
    requestId: String(params.requestId || body.requestId || '').trim(),
    transferId: String(params.transferId || body.transferId || payload.transferId || '').trim(),
    payload: payload
  };
}

/* =========================
   SCHEDULE RECEIVER
   ========================= */

function neoChronoHandleSchedule_(payload, transferId) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Missing NeoChrono schedule payload.');
  }

  if (String(payload.sheetName || '') !== NEOCHRONO_RECEIVER.SCHEDULE_SHEET) {
    throw new Error(
      'NeoChrono attempted to write to an unexpected sheet. Expected "' +
      NEOCHRONO_RECEIVER.SCHEDULE_SHEET +
      '".'
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
  var dateColumnIndex = headers.indexOf('Date');

  if (dateColumnIndex < 0) {
    throw new Error('The incoming Schedule payload does not contain a Date column.');
  }

  var rows = payload.rows.map(function(row, rowIndex) {
    if (!Array.isArray(row)) {
      throw new Error(
        'Schedule row ' +
        (rowIndex + 1) +
        ' is invalid.'
      );
    }

    var normalized = row.slice(0, width);

    while (normalized.length < width) {
      normalized.push('');
    }

    return normalized.map(function(value) {
      return value === null || value === undefined ? '' : value;
    });
  });

  var ss = neoChronoSpreadsheet_();
  var sheet = ss.getSheetByName(NEOCHRONO_RECEIVER.SCHEDULE_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(NEOCHRONO_RECEIVER.SCHEDULE_SHEET);
  }

  var timezone =
    ss.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'America/New_York';

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
    var incomingDates =
      rows
        .map(function(row) {
          return neoChronoScheduleDateKey_(
            row[dateColumnIndex],
            timezone
          );
        })
        .filter(function(value) {
          return !!value;
        })
        .sort();

    if (!startKey && incomingDates.length) {
      startKey = incomingDates[0];
    }

    if (!endKey && incomingDates.length) {
      endKey = incomingDates[incomingDates.length - 1];
    }
  }

  if (!startKey || !endKey) {
    throw new Error(
      'The Schedule transfer is missing a valid Start Date / End Date.'
    );
  }

  if (startKey > endKey) {
    throw new Error(
      'Schedule Start Date must be before or equal to End Date.'
    );
  }

  rows.forEach(function(row, index) {
    var rowDate =
      neoChronoScheduleDateKey_(
        row[dateColumnIndex],
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
        ', outside the finalized range ' +
        startKey +
        ' through ' +
        endKey +
        '.'
      );
    }
  });

  var oldLastRow = sheet.getLastRow();
  var oldLastColumn = sheet.getLastColumn();
  var existingRows = [];

  if (oldLastRow >= 1 && oldLastColumn > 0) {
    var existingHeaderWidth =
      Math.min(
        oldLastColumn,
        width
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
        existingHeaderWidth !== width ||
        existingHeaders.join('\u001f') !== headers.join('\u001f')
      )
    ) {
      throw new Error(
        'The existing Schedule sheet headers do not match the incoming NeoChrono schedule. ' +
        'No schedule data was changed.'
      );
    }

    if (oldLastRow > 1) {
      existingRows =
        sheet
          .getRange(
            2,
            1,
            oldLastRow - 1,
            width
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

  var preservedRows =
    existingRows.filter(function(row) {
      var dateKey =
        neoChronoScheduleDateKey_(
          row[dateColumnIndex],
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
    neoChronoSortScheduleRows_(
      preservedRows.concat(rows),
      dateColumnIndex,
      timezone
    );

  var requiredRows =
    Math.max(
      1,
      mergedRows.length + 1
    );

  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      requiredRows - sheet.getMaxRows()
    );
  }

  if (sheet.getMaxColumns() < width) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      width - sheet.getMaxColumns()
    );
  }

  sheet
    .getRange(
      1,
      1,
      1,
      width
    )
    .setValues([
      headers
    ]);

  if (mergedRows.length) {
    sheet
      .getRange(
        2,
        1,
        mergedRows.length,
        width
      )
      .setValues(
        mergedRows
      );
  }

  var finalLastRow =
    mergedRows.length + 1;

  if (oldLastRow > finalLastRow) {
    sheet
      .getRange(
        finalLastRow + 1,
        1,
        oldLastRow - finalLastRow,
        width
      )
      .clearContent();
  }

  sheet.setFrozenRows(1);

  sheet
    .getRange(
      1,
      1,
      1,
      width
    )
    .setFontWeight('bold');

  if (mergedRows.length) {
    sheet
      .getRange(
        2,
        dateColumnIndex + 1,
        mergedRows.length,
        1
      )
      .setNumberFormat('yyyy-mm-dd');
  }

  SpreadsheetApp.flush();

  var writtenColumns =
    sheet.getLastColumn();

  if (writtenColumns < width) {
    throw new Error(
      'Google verification failed. Expected at least ' +
      width +
      ' columns but found ' +
      writtenColumns +
      '.'
    );
  }

  var actualLastRow =
    sheet.getLastRow();

  if (actualLastRow !== finalLastRow) {
    throw new Error(
      'Google row verification failed. Expected ' +
      finalLastRow +
      ' total rows but found ' +
      actualLastRow +
      '.'
    );
  }

  var verifyRows =
    mergedRows.length
      ? sheet
          .getRange(
            2,
            1,
            mergedRows.length,
            width
          )
          .getValues()
      : [];

  var rangeCount =
    verifyRows.filter(function(row) {
      var dateKey =
        neoChronoScheduleDateKey_(
          row[dateColumnIndex],
          timezone
        );

      return (
        dateKey &&
        dateKey >= startKey &&
        dateKey <= endKey
      );
    }).length;

  if (rangeCount !== rows.length) {
    throw new Error(
      'Google date-range verification failed. NeoChrono sent ' +
      rows.length +
      ' row(s) for ' +
      startKey +
      ' through ' +
      endKey +
      ', but Google contains ' +
      rangeCount +
      ' row(s) in that range.'
    );
  }

  var receiverUrl =
    ss.getUrl() +
    '#gid=' +
    sheet.getSheetId();

  return neoChronoBridgeResponse_({
    type:
      NEOCHRONO_RECEIVER.SCHEDULE_MESSAGE_TYPE,

    ok:
      true,

    success:
      true,

    transferId:
      transferId,

    receiverSpreadsheetId:
      ss.getId(),

    receiverSpreadsheetName:
      ss.getName(),

    receiverSheetName:
      sheet.getName(),

    receiverUrl:
      receiverUrl,

    /*
     * These represent THIS transfer only so the browser's verification
     * continues to work.
     */
    writtenRows:
      rows.length,

    writtenColumns:
      width,

    transferredRows:
      rows.length,

    transferredColumns:
      width,

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

    message:
      'Schedule merged successfully. ' +
      rows.length +
      ' row(s) were written for ' +
      startKey +
      ' through ' +
      endKey +
      '. ' +
      preservedRows.length +
      ' existing row(s) outside that range were preserved. ' +
      mergedRows.length +
      ' total schedule row(s) are now stored.'
  });
}


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

  if (isNaN(parsed.getTime())) {
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


function neoChronoSortScheduleRows_(
  rows,
  dateColumnIndex,
  timezone
) {
  return (rows || [])
    .map(function(row, index) {
      return {
        row:
          row,

        index:
          index,

        date:
          neoChronoScheduleDateKey_(
            row[dateColumnIndex],
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

      return a.index - b.index;
    })
    .map(function(item) {
      return item.row;
    });
}

/* =========================
   PRN "MY AVAILABILITY" READ
   ========================= */

function neoChronoValidatePrnAvailabilitySheetName_(requestedSheet) {
  var requested = String(
    requestedSheet ||
    NEOCHRONO_RECEIVER.PRN_AVAILABILITY_SHEET
  ).trim();

  if (requested !== NEOCHRONO_RECEIVER.PRN_AVAILABILITY_SHEET) {
    throw new Error(
      'Invalid PRN availability sheet. Expected "' +
      NEOCHRONO_RECEIVER.PRN_AVAILABILITY_SHEET +
      '".'
    );
  }
}

function neoChronoHandlePrnAvailabilityRead_(requestId, requestedSheet) {
  neoChronoValidatePrnAvailabilitySheetName_(requestedSheet);

  var ss = neoChronoSpreadsheet_();
  var snapshot = neoChronoPrnAvailabilitySnapshot_(ss);

  return neoChronoBridgeResponse_({
    type: NEOCHRONO_RECEIVER.PRN_AVAILABILITY_MESSAGE_TYPE,
    ok: true,
    success: true,
    requestId: requestId,
    sheet: NEOCHRONO_RECEIVER.PRN_AVAILABILITY_SHEET,
    headers: snapshot.headers,
    rows: snapshot.rows,
    rowCount: snapshot.rowCount,
    generatedAt: snapshot.generatedAt,
    message: 'PRN availability loaded from My Availability.'
  });
}

function neoChronoHandlePrnAvailabilityExport_(requestedSheet) {
  neoChronoValidatePrnAvailabilitySheetName_(requestedSheet);

  var ss = neoChronoSpreadsheet_();
  var snapshot = neoChronoPrnAvailabilitySnapshot_(ss);

  return neoChronoJsonResponse_({
    ok: true,
    success: true,
    sheet: NEOCHRONO_RECEIVER.PRN_AVAILABILITY_SHEET,
    headers: snapshot.headers,
    rows: snapshot.rows,
    rowCount: snapshot.rowCount,
    generatedAt: snapshot.generatedAt,
    message: 'PRN availability export generated.'
  });
}

function neoChronoPrnAvailabilitySnapshot_(ss) {
  var sheet = ss.getSheetByName(
    NEOCHRONO_RECEIVER.PRN_AVAILABILITY_SHEET
  );

  if (!sheet) {
    return {
      headers: [],
      rows: [],
      rowCount: 0,
      generatedAt: new Date().toISOString()
    };
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 1 || lastColumn < 1) {
    return {
      headers: [],
      rows: [],
      rowCount: 0,
      generatedAt: new Date().toISOString()
    };
  }

  var values = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();

  var headers = values[0].map(function(value) {
    return String(value || '').trim();
  });

  var timezone =
    ss.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'America/New_York';

  var rows = [];

  for (var r = 1; r < values.length; r++) {
    var source = values[r];
    var object = {};
    var hasValue = false;

    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      if (!header) continue;

      var value = source[col];

      if (value !== '' && value !== null && value !== undefined) {
        hasValue = true;
      }

      if (value instanceof Date && !isNaN(value.getTime())) {
        if (/time/i.test(header) && !/date/i.test(header)) {
          object[header] = Utilities.formatDate(value, timezone, 'HH:mm');
        } else if (/date|day/i.test(header)) {
          object[header] = Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
        } else {
          object[header] = value.toISOString();
        }
      } else {
        object[header] =
          value === null || value === undefined
            ? ''
            : value;
      }
    }

    if (hasValue) rows.push(object);
  }

  return {
    headers: headers,
    rows: rows,
    rowCount: rows.length,
    generatedAt: new Date().toISOString()
  };
}

/* =========================
   PTO READ / WRITE
   ========================= */

function neoChronoHandlePtoRead_(requestId, requestedSheet) {
  neoChronoValidatePtoSheetName_(requestedSheet);

  var ss = neoChronoSpreadsheet_();
  var sheet = neoChronoPreparePtoSheet_(ss);
  var lock = LockService.getScriptLock();

  lock.waitLock(30000);
  try {
    neoChronoReconcilePtoAutoApprovals_(sheet, 'SYSTEM');
    var snapshot = neoChronoPtoSnapshot_(sheet, ss);

    return neoChronoBridgeResponse_({
      type: NEOCHRONO_RECEIVER.PTO_MESSAGE_TYPE,
      ok: true,
      success: true,
      requestId: requestId,
      sheet: NEOCHRONO_RECEIVER.PTO_SHEET,
      headers: snapshot.headers,
      rows: snapshot.rows,
      rowCount: snapshot.rowCount,
      generatedAt: snapshot.generatedAt,
      message: 'PTO / Regular Off data loaded from Google Sheets.'
    });
  } finally {
    lock.releaseLock();
  }
}

function neoChronoHandlePtoExport_(requestedSheet) {
  neoChronoValidatePtoSheetName_(requestedSheet);

  var ss = neoChronoSpreadsheet_();
  var sheet = neoChronoPreparePtoSheet_(ss);
  var lock = LockService.getScriptLock();

  lock.waitLock(30000);
  try {
    neoChronoReconcilePtoAutoApprovals_(sheet, 'SYSTEM');
    var snapshot = neoChronoPtoSnapshot_(sheet, ss);

    return neoChronoJsonResponse_({
      ok: true,
      success: true,
      sheet: NEOCHRONO_RECEIVER.PTO_SHEET,
      headers: snapshot.headers,
      rows: snapshot.rows,
      rowCount: snapshot.rowCount,
      generatedAt: snapshot.generatedAt,
      message: 'PTO / Regular Off export generated for neochrono-data.'
    });
  } catch (error) {
    return neoChronoJsonResponse_({
      ok: false,
      success: false,
      sheet: NEOCHRONO_RECEIVER.PTO_SHEET,
      message: error && error.message ? error.message : String(error)
    });
  } finally {
    lock.releaseLock();
  }
}

function neoChronoPtoSnapshot_(sheet, ss) {
  var table = neoChronoReadPtoTable_(sheet);
  var rows = table.rows.map(function(item) {
    return neoChronoSerializePtoObject_(item.object, ss);
  });

  return {
    headers: table.headers,
    rows: rows,
    rowCount: rows.length,
    generatedAt: new Date().toISOString()
  };
}

function neoChronoHandlePtoWrite_(requestId, requestedSheet, payload) {
  neoChronoValidatePtoSheetName_(requestedSheet);

  if (!payload || typeof payload !== 'object') {
    throw new Error('Missing PTO write payload.');
  }

  var operation = String(payload.operation || '').trim().toLowerCase();
  var actor = String(
    payload.actor ||
    (payload.row && payload.row['Updated By']) ||
    'NeoChrono'
  ).trim() || 'NeoChrono';

  var ss = neoChronoSpreadsheet_();
  var sheet = neoChronoPreparePtoSheet_(ss);
  var lock = LockService.getScriptLock();

  lock.waitLock(30000);
  try {
    var result;

    if (operation === 'save') {
      result = neoChronoSavePto_(sheet, payload.row || {}, actor);
    } else if (operation === 'review') {
      result = neoChronoReviewPto_(
        sheet,
        payload.recordId,
        payload.status,
        payload.comment,
        actor
      );
    } else if (operation === 'delete') {
      result = neoChronoDeletePto_(
        sheet,
        payload.recordId,
        actor
      );
    } else {
      throw new Error('Unsupported PTO operation: ' + operation);
    }

    SpreadsheetApp.flush();

    // Return the entire confirmed Google PTO snapshot. The NeoChrono browser
    // commits this snapshot into neochrono-data immediately, then the scheduler
    // reads only from GitHub.
    var snapshot = neoChronoPtoSnapshot_(sheet, ss);

    return neoChronoBridgeResponse_({
      type: NEOCHRONO_RECEIVER.PTO_MESSAGE_TYPE,
      ok: true,
      success: true,
      requestId: requestId,
      sheet: NEOCHRONO_RECEIVER.PTO_SHEET,
      operation: operation,
      recordId: result.recordId || '',
      status: result.status || '',
      autoApproved: !!result.autoApproved,
      blockedDates: result.blockedDates || [],
      headers: snapshot.headers,
      rows: snapshot.rows,
      rowCount: snapshot.rowCount,
      generatedAt: snapshot.generatedAt,
      message: result.message || 'Time-off change saved to Google Sheets.'
    });

  } finally {
    lock.releaseLock();
  }
}

function neoChronoSavePto_(sheet, incoming, actor) {
  var table = neoChronoReadPtoTable_(sheet);
  var now = new Date();
  var row = {};

  NEOCHRONO_PTO_HEADERS.forEach(function(header) {
    var value = incoming && incoming[header] !== undefined
      ? incoming[header]
      : '';
    row[header] = value === null || value === undefined ? '' : value;
  });

  row['Record Type'] = neoChronoQueuedTimeOffType_(row['Record Type'] || 'PTO');
  if (!row['Record Type']) {
    throw new Error('This endpoint accepts only PTO or Regular Off requests.');
  }

  row['Record ID'] = String(row['Record ID'] || '').trim();
  if (!row['Record ID']) {
    row['Record ID'] = neoChronoMakePtoRecordId_(now);
  }

  row.Pharmacist = String(row.Pharmacist || '').trim();
  row.Username = String(row.Username || '').trim();

  if (!row.Pharmacist && !row.Username) {
    throw new Error('Time-off request is missing Pharmacist / Username.');
  }

  var requestedDates = neoChronoPtoDateKeys_(row);
  if (!requestedDates.length) {
    throw new Error('PTO / Regular Off requires a Single Date or a Start Date / End Date.');
  }

  var existing = neoChronoFindPtoRow_(table, row['Record ID']);
  var existingObj = existing ? existing.object : null;

  var duplicate = neoChronoFindDuplicatePto_(table, row, row['Record ID']);
  if (duplicate) {
    throw new Error(
      'This pharmacist already has a PTO / Regular Off request that overlaps ' +
      duplicate.date +
      ' (Record ' +
      duplicate.recordId +
      ').'
    );
  }

  var wasRejected =
    existingObj &&
    String(existingObj.Status || '').trim().toLowerCase() === 'rejected';

  row['Submitted At'] =
    existingObj && existingObj['Submitted At'] && !wasRejected
      ? existingObj['Submitted At']
      : now.toISOString();

  row['Updated At'] = now.toISOString();
  row['Updated By'] = actor;

  var manualApproved =
    existingObj &&
    String(existingObj.Status || '').trim().toLowerCase() === 'approved' &&
    String(existingObj['Reviewed By'] || '').trim().toUpperCase() !== 'SYSTEM AUTO-APPROVAL';

  if (manualApproved) {
    row.Status = 'Approved';
    row['Reviewed By'] = existingObj['Reviewed By'];
    row['Reviewed At'] = existingObj['Reviewed At'];
  } else {
    row.Status = 'Pending';
    row['Reviewed By'] = '';
    row['Reviewed At'] = '';
  }

  if (existing) {
    neoChronoWritePtoObject_(sheet, table.headers, existing.rowNumber, row);
  } else {
    neoChronoAppendPtoObject_(sheet, table.headers, row);
  }

  var reconcile = neoChronoReconcilePtoAutoApprovals_(sheet, 'SYSTEM');
  var refreshed = neoChronoReadPtoTable_(sheet);
  var saved = neoChronoFindPtoRow_(refreshed, row['Record ID']);

  if (!saved) {
    throw new Error('Time-off verification failed after saving.');
  }

  var queue = neoChronoBuildPtoQueueState_(
    refreshed.rows.map(function(x) { return x.object; }),
    NEOCHRONO_RECEIVER.PTO_AUTO_APPROVE_LIMIT
  );

  var q = queue[row['Record ID']] || {
    autoEligible: false,
    blockedDates: requestedDates
  };

  var savedStatus = String(saved.object.Status || '').trim();
  var autoApproved =
    savedStatus.toLowerCase() === 'approved' &&
    String(saved.object['Reviewed By'] || '').trim().toUpperCase() === 'SYSTEM AUTO-APPROVAL';

  var message;
  if (autoApproved) {
    message =
      'AUTO APPROVED: this request is within the first ' +
      NEOCHRONO_RECEIVER.PTO_AUTO_APPROVE_LIMIT +
      ' combined PTO / Regular Off request(s) on every requested OFF date.';
  } else if (savedStatus.toLowerCase() === 'approved') {
    message = row['Record Type'] + ' request is approved by an administrator.';
  } else {
    message =
      'PENDING ADMIN REVIEW: this request is number 3 or later on ' +
      ((q.blockedDates && q.blockedDates.length)
        ? q.blockedDates.join(', ')
        : 'one or more covered dates') +
      '.';
  }

  return {
    recordId: row['Record ID'],
    status: savedStatus,
    autoApproved: autoApproved,
    blockedDates: q.blockedDates || [],
    message: message,
    ptoAutoApprovalsUpdated: reconcile.updated
  };
}

function neoChronoReviewPto_(sheet, recordId, status, comment, actor) {
  var id = String(recordId || '').trim();
  if (!id) throw new Error('Request ID is required.');

  var normalizedStatus = String(status || '').trim();
  if (!normalizedStatus) throw new Error('Review status is required.');

  var allowed = {
    'approved': 'Approved',
    'rejected': 'Rejected',
    'pending': 'Pending'
  };

  var key = normalizedStatus.toLowerCase();
  if (!allowed[key]) {
    throw new Error('Review status must be Approved, Rejected, or Pending.');
  }

  var table = neoChronoReadPtoTable_(sheet);
  var found = neoChronoFindPtoRow_(table, id);
  if (!found) throw new Error('Request not found: ' + id);

  var row = found.object;

  if (
    String(row.Status || '').trim().toLowerCase() === 'approved' &&
    key === 'pending'
  ) {
    throw new Error('An approved PTO / Regular Off request cannot be returned to Pending. Use Reject or Delete if the approval must be removed.');
  }

  row.Status = allowed[key];
  row.Comment = comment === undefined
    ? row.Comment
    : String(comment);

  if (key === 'pending') {
    row['Reviewed By'] = '';
    row['Reviewed At'] = '';
  } else {
    row['Reviewed By'] = actor;
    row['Reviewed At'] = new Date().toISOString();
  }

  row['Updated At'] = new Date().toISOString();
  row['Updated By'] = actor;

  neoChronoWritePtoObject_(sheet, table.headers, found.rowNumber, row);
  neoChronoReconcilePtoAutoApprovals_(sheet, 'SYSTEM');

  var refreshed = neoChronoReadPtoTable_(sheet);
  var saved = neoChronoFindPtoRow_(refreshed, id);

  return {
    recordId: id,
    status: saved ? String(saved.object.Status || '') : allowed[key],
    message: 'Time-off review saved to Google Sheets.'
  };
}

function neoChronoDeletePto_(sheet, recordId, actor) {
  var id = String(recordId || '').trim();
  if (!id) throw new Error('Request ID is required.');

  var table = neoChronoReadPtoTable_(sheet);
  var found = neoChronoFindPtoRow_(table, id);
  if (!found) throw new Error('Request not found: ' + id);

  sheet.deleteRow(found.rowNumber);
  neoChronoReconcilePtoAutoApprovals_(sheet, actor || 'SYSTEM');

  return {
    recordId: id,
    status: 'Deleted',
    message: 'Time-off request deleted from Google Sheets.'
  };
}

/* =========================
   PTO TABLE HELPERS
   ========================= */

function neoChronoPreparePtoSheet_(ss) {
  var sheet = ss.getSheetByName(NEOCHRONO_RECEIVER.PTO_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(NEOCHRONO_RECEIVER.PTO_SHEET);
  }

  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function(v) { return String(v || '').trim(); });

  var hasAnyHeader = existing.some(function(v) { return !!v; });

  if (!hasAnyHeader) {
    sheet
      .getRange(1, 1, 1, NEOCHRONO_PTO_HEADERS.length)
      .setValues([NEOCHRONO_PTO_HEADERS.slice()]);
  } else {
    var headers = existing.slice();
    NEOCHRONO_PTO_HEADERS.forEach(function(required) {
      if (headers.indexOf(required) === -1) {
        headers.push(required);
      }
    });

    if (headers.length !== existing.length) {
      if (sheet.getMaxColumns() < headers.length) {
        sheet.insertColumnsAfter(
          sheet.getMaxColumns(),
          headers.length - sheet.getMaxColumns()
        );
      }

      sheet
        .getRange(1, 1, 1, headers.length)
        .setValues([headers]);
    }
  }

  sheet.setFrozenRows(1);
  sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), NEOCHRONO_PTO_HEADERS.length))
    .setFontWeight('bold');

  SpreadsheetApp.flush();
  return sheet;
}

function neoChronoReadPtoTable_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastColumn = Math.max(sheet.getLastColumn(), NEOCHRONO_PTO_HEADERS.length);
  var values = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();

  var headers = values[0].map(function(v) {
    return String(v || '').trim();
  });

  var rows = [];

  for (var r = 1; r < values.length; r++) {
    var source = values[r];
    var object = {};
    var hasValue = false;

    for (var c = 0; c < headers.length; c++) {
      var header = headers[c];
      if (!header) continue;

      var value = source[c];
      object[header] = value;

      if (value !== '' && value !== null && value !== undefined) {
        hasValue = true;
      }
    }

    if (hasValue) {
      rows.push({
        rowNumber: r + 1,
        object: object
      });
    }
  }

  return {
    headers: headers,
    rows: rows
  };
}

function neoChronoFindPtoRow_(table, recordId) {
  var id = String(recordId || '').trim();

  for (var i = 0; i < table.rows.length; i++) {
    var item = table.rows[i];
    if (String(item.object['Record ID'] || '').trim() === id) {
      return item;
    }
  }

  return null;
}

function neoChronoWritePtoObject_(sheet, headers, rowNumber, object) {
  var values = headers.map(function(header) {
    var value = object[header];
    return value === null || value === undefined ? '' : value;
  });

  sheet
    .getRange(rowNumber, 1, 1, headers.length)
    .setValues([values]);
}

function neoChronoAppendPtoObject_(sheet, headers, object) {
  var rowNumber = Math.max(sheet.getLastRow() + 1, 2);
  neoChronoWritePtoObject_(sheet, headers, rowNumber, object);
}

function neoChronoSerializePtoObject_(object, ss) {
  var timezone =
    ss.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'America/New_York';

  var out = {};

  Object.keys(object || {}).forEach(function(header) {
    var value = object[header];

    if (!(value instanceof Date)) {
      out[header] =
        value === null || value === undefined
          ? ''
          : value;
      return;
    }

    if (neoChronoIsPtoDateOnlyHeader_(header)) {
      out[header] = Utilities.formatDate(
        value,
        timezone,
        'yyyy-MM-dd'
      );
    } else {
      out[header] = value.toISOString();
    }
  });

  return out;
}

function neoChronoIsPtoDateOnlyHeader_(header) {
  return {
    'Date': true,
    'Start Date': true,
    'End Date': true,
    'Weekend Saturday': true,
    'Weekend Sunday': true
  }[header] === true;
}

/* =========================
   TIME-OFF BUSINESS RULES
   ========================= */

function neoChronoQueuedTimeOffType_(value) {
  var type = String(value || '')
    .trim()
    .toUpperCase()
    .replace(/[\-_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (type === 'PTO') return 'PTO';
  if (type === 'REGULAR OFF' || type === 'REGULAR OFF REQUEST' || type === 'REGULAROFF') {
    return 'REGULAR OFF';
  }
  return '';
}

function neoChronoPtoDateKeys_(row) {
  var start = neoChronoDateKey_(row['Start Date']);
  var end = neoChronoDateKey_(row['End Date']);

  if (start || end) {
    if (start && !end) end = start;
    if (!start && end) start = end;

    if (!start || !end) return [];
    if (end < start) return [];

    var out = [];
    var current = neoChronoDateFromKey_(start);
    var finalDate = neoChronoDateFromKey_(end);

    while (current.getTime() <= finalDate.getTime()) {
      out.push(neoChronoUtcDateKey_(current));
      if (out.length > 370) {
        throw new Error('Time-off request is too long.');
      }
      current.setUTCDate(current.getUTCDate() + 1);
    }

    return out;
  }

  var single = neoChronoDateKey_(row.Date);
  return single ? [single] : [];
}

function neoChronoDateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      'UTC',
      'yyyy-MM-dd'
    );
  }

  var raw = String(value || '').trim();
  if (!raw) return '';

  var match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return match[1] + '-' + match[2] + '-' + match[3];
  }

  var parsed = new Date(raw);
  if (isNaN(parsed.getTime())) return '';

  return Utilities.formatDate(
    parsed,
    'UTC',
    'yyyy-MM-dd'
  );
}

function neoChronoDateFromKey_(key) {
  var parts = String(key || '').split('-');
  return new Date(Date.UTC(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  ));
}

function neoChronoUtcDateKey_(date) {
  return Utilities.formatDate(
    date,
    'UTC',
    'yyyy-MM-dd'
  );
}

function neoChronoFindDuplicatePto_(table, row, excludeRecordId) {
  var requested = {};
  neoChronoPtoDateKeys_(row).forEach(function(k) {
    requested[k] = true;
  });

  var username = String(row.Username || '').trim();
  var pharmacist = String(row.Pharmacist || '').trim();

  for (var i = 0; i < table.rows.length; i++) {
    var existing = table.rows[i].object;

    if (
      String(existing['Record ID'] || '').trim() ===
      String(excludeRecordId || '').trim()
    ) continue;

    if (!neoChronoQueuedTimeOffType_(existing['Record Type'])) continue;

    if (
      String(existing.Status || '').trim().toLowerCase() ===
      'rejected'
    ) continue;

    var samePerson =
      (username &&
       String(existing.Username || '').trim() === username) ||
      (pharmacist &&
       String(existing.Pharmacist || '').trim() === pharmacist);

    if (!samePerson) continue;

    var dates = neoChronoPtoDateKeys_(existing);

    for (var d = 0; d < dates.length; d++) {
      if (requested[dates[d]]) {
        return {
          recordId: String(existing['Record ID'] || '').trim(),
          date: dates[d]
        };
      }
    }
  }

  return null;
}

function neoChronoBuildPtoQueueState_(rows, limit) {
  limit = Math.max(
    0,
    Math.floor(
      Number(limit || NEOCHRONO_RECEIVER.PTO_AUTO_APPROVE_LIMIT)
    )
  );

  var items = (rows || []).map(function(row, index) {
    return {
      row: row,
      index: index,
      id: String(row['Record ID'] || '').trim(),
      submitted: neoChronoSubmittedMs_(row, index),
      dates: neoChronoPtoDateKeys_(row)
    };
  }).filter(function(item) {
    return (
      item.id &&
      !!neoChronoQueuedTimeOffType_(item.row['Record Type']) &&
      String(item.row.Status || '').trim().toLowerCase() !== 'rejected' &&
      item.dates.length > 0
    );
  });

  var byDate = {};

  items.forEach(function(item) {
    item.dates.forEach(function(dateKey) {
      if (!byDate[dateKey]) byDate[dateKey] = [];
      byDate[dateKey].push(item);
    });
  });

  var rankByRecord = {};

  items.forEach(function(item) {
    rankByRecord[item.id] = {};
  });

  Object.keys(byDate).forEach(function(dateKey) {
    var queue = byDate[dateKey].slice().sort(function(a, b) {
      return (
        a.submitted - b.submitted ||
        a.index - b.index ||
        a.id.localeCompare(b.id)
      );
    });

    queue.forEach(function(item, index) {
      rankByRecord[item.id][dateKey] = index + 1;
    });
  });

  var state = {};

  items.forEach(function(item) {
    var positions = rankByRecord[item.id] || {};
    var blockedDates = item.dates.filter(function(dateKey) {
      return Number(positions[dateKey] || 999999) > limit;
    });

    state[item.id] = {
      autoEligible:
        item.dates.length > 0 &&
        blockedDates.length === 0,
      blockedDates: blockedDates,
      positions: positions,
      dates: item.dates.slice(),
      limit: limit
    };
  });

  return state;
}

function neoChronoSubmittedMs_(row, index) {
  var raw =
    row['Submitted At'] ||
    row['Updated At'] ||
    '';

  var d = raw instanceof Date
    ? raw
    : new Date(String(raw || ''));

  return d && !isNaN(d.getTime())
    ? d.getTime()
    : Number(index || 0);
}

function neoChronoReconcilePtoAutoApprovals_(sheet, updatedBy) {
  var table = neoChronoReadPtoTable_(sheet);
  var objects = table.rows.map(function(item) {
    return item.object;
  });

  var state = neoChronoBuildPtoQueueState_(
    objects,
    NEOCHRONO_RECEIVER.PTO_AUTO_APPROVE_LIMIT
  );

  var now = new Date().toISOString();
  var updated = 0;
  var promoted = 0;
  var demoted = 0;

  table.rows.forEach(function(item) {
    var row = item.object;

    if (!neoChronoQueuedTimeOffType_(row['Record Type'])) return;

    var id = String(row['Record ID'] || '').trim();
    if (!id) return;

    var status = String(row.Status || '').trim().toLowerCase();
    if (status === 'rejected') return;

    var q = state[id];
    if (!q) return;

    var reviewer = String(row['Reviewed By'] || '').trim().toUpperCase();
    var systemAuto = reviewer === 'SYSTEM AUTO-APPROVAL';
    var changed = false;

    if (q.autoEligible) {
      if (status !== 'approved') {
        row.Status = 'Approved';
        row['Reviewed By'] = 'SYSTEM AUTO-APPROVAL';
        row['Reviewed At'] = now;
        row['Updated At'] = now;
        row['Updated By'] = updatedBy || 'SYSTEM';
        changed = true;
        promoted++;
      }
    }

    /*
     * Approval is sticky.
     * Once a PTO / Regular Off request reaches Approved status (whether automatically or by
     * an administrator), queue reconciliation must NEVER move it back to
     * Pending. Only an explicit administrator Reject/Delete action may remove
     * an approval. This also means approving request #3+ never steals approval
     * from an already-approved request.
     *
     * Rows entered directly in Google Sheets can have a blank Status. If they
     * are outside the first-two queue and are not already Approved, normalize
     * them to Pending so the admin sees the correct review state.
     */
    if (!q.autoEligible && status !== 'approved' && status !== 'pending') {
      row.Status = 'Pending';
      row['Reviewed By'] = '';
      row['Reviewed At'] = '';
      row['Updated At'] = now;
      row['Updated By'] = updatedBy || 'SYSTEM';
      changed = true;
    }

    if (changed) {
      neoChronoWritePtoObject_(
        sheet,
        table.headers,
        item.rowNumber,
        row
      );
      updated++;
    }
  });

  if (updated) SpreadsheetApp.flush();

  return {
    ok: true,
    updated: updated,
    promoted: promoted,
    demoted: demoted,
    limit: NEOCHRONO_RECEIVER.PTO_AUTO_APPROVE_LIMIT
  };
}

function neoChronoMakePtoRecordId_(date) {
  return (
    'REQ-' +
    Utilities.formatDate(
      date || new Date(),
      Session.getScriptTimeZone() || 'America/New_York',
      'yyyyMMddHHmmss'
    ) +
    '-' +
    Math.floor(Math.random() * 900 + 100)
  );
}

/* =========================
   SHARED HELPERS
   ========================= */

function neoChronoValidatePtoSheetName_(requestedSheet) {
  var requested = String(
    requestedSheet ||
    NEOCHRONO_RECEIVER.PTO_SHEET
  ).trim();

  if (requested !== NEOCHRONO_RECEIVER.PTO_SHEET) {
    throw new Error(
      'Invalid PTO sheet. Expected "' +
      NEOCHRONO_RECEIVER.PTO_SHEET +
      '".'
    );
  }
}

function neoChronoSpreadsheet_() {
  return SpreadsheetApp.openById(
    NEOCHRONO_RECEIVER.SPREADSHEET_ID
  );
}

function neoChronoJsonResponse_(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function neoChronoBridgeResponse_(result) {
  var json = JSON.stringify(result)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');

  var message = String(
    result && result.message
      ? result.message
      : ''
  );

  var html =
    '<!doctype html>' +
    '<html><head><meta charset="utf-8"></head><body>' +
    '<script>' +
    'try{' +
    'var m=' + json + ';' +
    'if(window.top){window.top.postMessage(m, "*");}' +
    'if(window.parent&&window.parent!==window.top){window.parent.postMessage(m, "*");}' +
    '}catch(e){}' +
    '</script>' +
    '<div style="font-family:Arial,sans-serif;padding:16px">' +
    neoChronoHtmlEscape_(message) +
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
 * Manual test from Apps Script editor.
 */
function testNeoChronoCode5() {
  var ss = neoChronoSpreadsheet_();
  var pto = neoChronoPreparePtoSheet_(ss);
  var schedule = ss.getSheetByName(
    NEOCHRONO_RECEIVER.SCHEDULE_SHEET
  );

  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    ptoSheetExists: !!pto,
    ptoLastRow: pto ? pto.getLastRow() : 0,
    scheduleSheetExists: !!schedule,
    scheduleLastRow: schedule ? schedule.getLastRow() : 0,
    spreadsheetUrl: ss.getUrl()
  };
}
