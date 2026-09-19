/**
 * Code5.gs — NeoChrono PTO / Availability Requests API
 *
 * PURPOSE
 * -------
 * Makes the Google Sheet tab named exactly:
 *   PTO / Availability Requests
 * the live read/write source for NeoChrono.
 *
 * IMPORTANT
 * ---------
 * 1) Add this file to the SAME Apps Script project behind:
 *    https://script.google.com/macros/s/AKfycbxBTnuzFvNXOROz4fpGni_exZap2YfSTm5aNWw4eAwnyGe-m9jU3SZriupyFf3yDP-r/exec
 *
 * 2) DO NOT create a second doGet(). Your project already has doGet().
 *    Add this near the TOP of your existing doGet(e):
 *
 *      var action = String(e && e.parameter && e.parameter.action || '').trim();
 *      if (action === 'neochronoPto') {
 *        return neoChronoPtoReadApiV5_(e);
 *      }
 *
 * 3) If your project does NOT already have doPost(e), keep the doPost(e)
 *    function included at the bottom of this file.
 *
 *    If the project ALREADY has doPost(e), do not keep two doPost functions.
 *    Instead add this near the TOP of the existing doPost(e):
 *
 *      var body = neoChronoParsePostV5_(e);
 *      if (body && body.action === 'neochronoPtoWrite') {
 *        return neoChronoPtoWriteApiV5_(e, body);
 *      }
 *
 * 4) Deploy a NEW VERSION of the EXISTING web-app deployment.
 *    Do not create a different deployment URL.
 */

var NEOCHRONO_PTO_V5 = Object.freeze({
  SHEET_NAME: 'PTO / Availability Requests',
  AUTO_APPROVE_PER_DATE: 2,
  HEADERS: [
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
  ]
});


function neoChronoPtoSpreadsheetV5_() {
  // Prefer the helper already used by the existing PTO app, if present.
  if (typeof getPtoRequestSpreadsheet_ === 'function') {
    var existing = getPtoRequestSpreadsheet_();
    if (existing) return existing;
  }

  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  var id = PropertiesService
    .getScriptProperties()
    .getProperty('NEOCHRONO_PTO_SPREADSHEET_ID');

  if (id) return SpreadsheetApp.openById(id);

  throw new Error(
    'PTO spreadsheet could not be resolved. ' +
    'Bind this Apps Script project to the spreadsheet, keep the existing ' +
    'getPtoRequestSpreadsheet_() helper, or set script property ' +
    'NEOCHRONO_PTO_SPREADSHEET_ID.'
  );
}


function neoChronoPtoSheetV5_() {
  var ss = neoChronoPtoSpreadsheetV5_();
  var sh = ss.getSheetByName(NEOCHRONO_PTO_V5.SHEET_NAME);

  if (!sh) {
    throw new Error(
      'Sheet not found: ' + NEOCHRONO_PTO_V5.SHEET_NAME
    );
  }

  neoChronoEnsurePtoHeadersV5_(sh);
  return sh;
}


function neoChronoEnsurePtoHeadersV5_(sh) {
  var wanted = NEOCHRONO_PTO_V5.HEADERS.slice();
  var lastCol = Math.max(sh.getLastColumn(), wanted.length);
  var current = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  var headers = current
    .map(function(v){ return String(v || '').trim(); })
    .filter(function(v){ return !!v; });

  if (!headers.length) {
    sh.getRange(1, 1, 1, wanted.length).setValues([wanted]);
    sh.setFrozenRows(1);
    return;
  }

  wanted.forEach(function(h){
    if (headers.indexOf(h) < 0) headers.push(h);
  });

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
}


function neoChronoHeadersV5_(sh) {
  return sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(v){ return String(v || '').trim(); });
}


function neoChronoRowsV5_(sh) {
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  var headers = values[0].map(function(v){
    return String(v || '').trim();
  });

  var rows = [];

  for (var r = 1; r < values.length; r++) {
    var obj = {};
    var hasValue = false;

    for (var c = 0; c < headers.length; c++) {
      var value = values[r][c];

      if (value !== '' && value !== null && value !== undefined) {
        hasValue = true;
      }

      obj[headers[c]] = value;
    }

    if (hasValue) {
      obj.__rowNumber = r + 1;
      rows.push(obj);
    }
  }

  return rows;
}


function neoChronoFindRowV5_(rows, recordId) {
  var wanted = String(recordId || '').trim();

  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['Record ID'] || '').trim() === wanted) {
      return rows[i];
    }
  }

  return null;
}


function neoChronoWriteObjectV5_(sh, rowNumber, obj) {
  var headers = neoChronoHeadersV5_(sh);
  var values = headers.map(function(h){
    var value = obj[h];

    if (
      value === null ||
      value === undefined
    ) {
      return '';
    }

    return value;
  });

  if (rowNumber) {
    sh.getRange(rowNumber, 1, 1, headers.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }
}


function neoChronoDateKeyV5_(value, timezone) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      timezone || Session.getScriptTimeZone() || 'America/New_York',
      'yyyy-MM-dd'
    );
  }

  var text = String(value).trim();
  var m = text.match(/^(\d{4})-(\d{2})-(\d{2})/);

  if (m) {
    return m[1] + '-' + m[2] + '-' + m[3];
  }

  var parsed = new Date(text);

  if (isNaN(parsed.getTime())) return '';

  return Utilities.formatDate(
    parsed,
    timezone || Session.getScriptTimeZone() || 'America/New_York',
    'yyyy-MM-dd'
  );
}


function neoChronoRequestDatesV5_(row, timezone) {
  var start = neoChronoDateKeyV5_(row['Start Date'], timezone);
  var end = neoChronoDateKeyV5_(row['End Date'], timezone);

  if (start || end) {
    if (!start) start = end;
    if (!end) end = start;

    if (end < start) {
      throw new Error('PTO End Date cannot be before Start Date.');
    }

    var out = [];
    var d = new Date(start + 'T12:00:00');
    var stop = new Date(end + 'T12:00:00');

    while (d.getTime() <= stop.getTime()) {
      out.push(
        Utilities.formatDate(
          d,
          timezone || 'America/New_York',
          'yyyy-MM-dd'
        )
      );

      d.setDate(d.getDate() + 1);

      if (out.length > 370) {
        throw new Error('PTO request is too long.');
      }
    }

    return out;
  }

  var single = neoChronoDateKeyV5_(row.Date, timezone);
  return single ? [single] : [];
}


function neoChronoSubmittedMsV5_(row, fallback) {
  var v = row['Submitted At'] || row['Updated At'];

  if (v instanceof Date) {
    return v.getTime();
  }

  var parsed = new Date(v);

  if (!isNaN(parsed.getTime())) {
    return parsed.getTime();
  }

  return fallback || 0;
}


function neoChronoRebalancePtoV5_(sh) {
  var ss = sh.getParent();
  var timezone =
    ss.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'America/New_York';

  var rows = neoChronoRowsV5_(sh);
  var items = [];

  rows.forEach(function(row, index){
    var type = String(row['Record Type'] || '').trim().toUpperCase();
    var status = String(row.Status || '').trim().toLowerCase();
    var id = String(row['Record ID'] || '').trim();

    if (
      type !== 'PTO' ||
      !id ||
      status === 'rejected'
    ) {
      return;
    }

    var dates = neoChronoRequestDatesV5_(row, timezone);

    if (!dates.length) return;

    items.push({
      row: row,
      index: index,
      id: id,
      dates: dates,
      submitted: neoChronoSubmittedMsV5_(row, index)
    });
  });

  var byDate = {};

  items.forEach(function(item){
    item.dates.forEach(function(dateKey){
      if (!byDate[dateKey]) byDate[dateKey] = [];
      byDate[dateKey].push(item);
    });
  });

  var rankByRecord = {};

  Object.keys(byDate).forEach(function(dateKey){
    var queue = byDate[dateKey].slice().sort(function(a,b){
      return (
        a.submitted - b.submitted ||
        a.index - b.index ||
        a.id.localeCompare(b.id)
      );
    });

    queue.forEach(function(item, index){
      if (!rankByRecord[item.id]) rankByRecord[item.id] = {};
      rankByRecord[item.id][dateKey] = index + 1;
    });
  });

  var now = new Date();

  items.forEach(function(item){
    var row = item.row;
    var reviewedBy = String(row['Reviewed By'] || '').trim();
    var currentStatus = String(row.Status || '').trim().toLowerCase();

    // Preserve an explicit administrator approval.
    var manualApproval =
      currentStatus === 'approved' &&
      reviewedBy &&
      reviewedBy.toUpperCase() !== 'SYSTEM AUTO-APPROVAL';

    if (manualApproval) return;

    var eligible = item.dates.every(function(dateKey){
      var rank =
        rankByRecord[item.id] &&
        rankByRecord[item.id][dateKey];

      return (
        rank &&
        rank <= NEOCHRONO_PTO_V5.AUTO_APPROVE_PER_DATE
      );
    });

    var desiredStatus = eligible ? 'Approved' : 'Pending';
    var changed = String(row.Status || '') !== desiredStatus;

    if (eligible) {
      row.Status = 'Approved';
      row['Reviewed By'] = 'SYSTEM AUTO-APPROVAL';
      row['Reviewed At'] = now;
    } else {
      row.Status = 'Pending';

      if (
        reviewedBy.toUpperCase() === 'SYSTEM AUTO-APPROVAL'
      ) {
        row['Reviewed By'] = '';
        row['Reviewed At'] = '';
      }
    }

    if (changed) {
      row['Updated At'] = now;
      row['Updated By'] = 'NeoChrono';
    }

    neoChronoWriteObjectV5_(sh, row.__rowNumber, row);
  });

  SpreadsheetApp.flush();
}


function neoChronoSaveRequestV5_(payload) {
  var sh = neoChronoPtoSheetV5_();
  var ss = sh.getParent();
  var timezone =
    ss.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'America/New_York';

  var incoming = payload && payload.row
    ? payload.row
    : {};

  var rows = neoChronoRowsV5_(sh);
  var id = String(incoming['Record ID'] || '').trim();

  if (!id) {
    id =
      'REQ-' +
      Utilities.formatDate(
        new Date(),
        timezone,
        'yyyyMMddHHmmss'
      ) +
      '-' +
      Math.floor(Math.random() * 900 + 100);
  }

  var old = neoChronoFindRowV5_(rows, id);
  var now = new Date();

  var row = {};
  NEOCHRONO_PTO_V5.HEADERS.forEach(function(h){
    row[h] = incoming[h] !== undefined
      ? incoming[h]
      : (
          old && old[h] !== undefined
            ? old[h]
            : ''
        );
  });

  row['Record ID'] = id;
  row['Record Type'] =
    String(row['Record Type'] || 'PTO')
      .trim()
      .toUpperCase();

  if (!row['Submitted At']) {
    row['Submitted At'] = now;
  }

  row['Updated At'] = now;
  row['Updated By'] = 'NeoChrono';

  if (row['Record Type'] === 'PTO') {
    var dates = neoChronoRequestDatesV5_(row, timezone);

    if (!dates.length) {
      throw new Error(
        'PTO requires a Single Date or a Start Date / End Date.'
      );
    }

    // Prevent overlapping duplicate PTO for the same pharmacist.
    var pharmacist = String(row.Pharmacist || '').trim();
    var username = String(row.Username || '').trim();

    for (var i = 0; i < rows.length; i++) {
      var other = rows[i];

      if (
        String(other['Record ID'] || '').trim() === id ||
        String(other['Record Type'] || '').trim().toUpperCase() !== 'PTO' ||
        String(other.Status || '').trim().toLowerCase() === 'rejected'
      ) {
        continue;
      }

      var samePerson =
        (
          username &&
          String(other.Username || '').trim() === username
        ) ||
        (
          pharmacist &&
          String(other.Pharmacist || '').trim() === pharmacist
        );

      if (!samePerson) continue;

      var otherDates = neoChronoRequestDatesV5_(other, timezone);
      var overlap = dates.some(function(d){
        return otherDates.indexOf(d) >= 0;
      });

      if (overlap) {
        throw new Error(
          'This pharmacist already has a PTO request overlapping an existing request.'
        );
      }
    }

    var preserveManualApproval =
      old &&
      String(old.Status || '').trim().toLowerCase() === 'approved' &&
      String(old['Reviewed By'] || '')
        .trim()
        .toUpperCase() !== 'SYSTEM AUTO-APPROVAL';

    if (preserveManualApproval) {
      row.Status = 'Approved';
      row['Reviewed By'] = old['Reviewed By'];
      row['Reviewed At'] = old['Reviewed At'];
    } else {
      row.Status = 'Pending';
      row['Reviewed By'] = '';
      row['Reviewed At'] = '';
    }
  } else {
    row.Status =
      String(row.Status || (old ? old.Status : 'Approved') || 'Approved');

    row['Reviewed By'] =
      row['Reviewed By'] || 'NeoChrono';

    row['Reviewed At'] =
      row['Reviewed At'] || now;
  }

  neoChronoWriteObjectV5_(
    sh,
    old ? old.__rowNumber : null,
    row
  );

  neoChronoRebalancePtoV5_(sh);

  return {
    success: true,
    operation: 'save',
    recordId: id
  };
}


function neoChronoReviewRequestV5_(payload) {
  var sh = neoChronoPtoSheetV5_();
  var rows = neoChronoRowsV5_(sh);
  var id = String(payload.recordId || '').trim();
  var row = neoChronoFindRowV5_(rows, id);

  if (!row) {
    throw new Error('Request not found: ' + id);
  }

  var status = String(payload.status || '').trim();

  if (
    ['Approved','Rejected','Pending'].indexOf(status) < 0
  ) {
    throw new Error(
      'Status must be Approved, Rejected, or Pending.'
    );
  }

  var now = new Date();

  row.Status = status;

  if (payload.comment !== undefined) {
    row.Comment = String(payload.comment || '');
  }

  row['Reviewed By'] = 'NeoChrono Admin';
  row['Reviewed At'] = now;
  row['Updated At'] = now;
  row['Updated By'] = 'NeoChrono Admin';

  neoChronoWriteObjectV5_(sh, row.__rowNumber, row);

  // Rejected requests stop consuming one of the first-two PTO slots.
  neoChronoRebalancePtoV5_(sh);

  return {
    success: true,
    operation: 'review',
    recordId: id,
    status: status
  };
}


function neoChronoDeleteRequestV5_(payload) {
  var sh = neoChronoPtoSheetV5_();
  var rows = neoChronoRowsV5_(sh);
  var id = String(payload.recordId || '').trim();
  var row = neoChronoFindRowV5_(rows, id);

  if (!row) {
    return {
      success: true,
      operation: 'delete',
      recordId: id,
      deleted: false
    };
  }

  sh.deleteRow(row.__rowNumber);
  neoChronoRebalancePtoV5_(sh);
  SpreadsheetApp.flush();

  return {
    success: true,
    operation: 'delete',
    recordId: id,
    deleted: true
  };
}


function neoChronoPtoReadApiV5_(e) {
  try {
    var sh = neoChronoPtoSheetV5_();
    var ss = sh.getParent();
    var timezone =
      ss.getSpreadsheetTimeZone() ||
      Session.getScriptTimeZone() ||
      'America/New_York';

    var headers = neoChronoHeadersV5_(sh);
    var rows = neoChronoRowsV5_(sh).map(function(row){
      var clean = {};

      headers.forEach(function(header){
        var value = row[header];

        if (value instanceof Date) {
          var dateOnly = {
            'Date': true,
            'Start Date': true,
            'End Date': true,
            'Weekend Saturday': true,
            'Weekend Sunday': true
          };

          clean[header] = dateOnly[header]
            ? Utilities.formatDate(
                value,
                timezone,
                'yyyy-MM-dd'
              )
            : value.toISOString();
        } else {
          clean[header] =
            value === null || value === undefined
              ? ''
              : value;
        }
      });

      return clean;
    });

    return neoChronoResponseV5_(
      {
        success: true,
        sheet: NEOCHRONO_PTO_V5.SHEET_NAME,
        headers: headers,
        rows: rows,
        rowCount: rows.length,
        generatedAt: new Date().toISOString()
      },
      e && e.parameter
        ? e.parameter.callback
        : ''
    );

  } catch (error) {
    return neoChronoResponseV5_(
      {
        success: false,
        sheet: NEOCHRONO_PTO_V5.SHEET_NAME,
        message:
          error && error.message
            ? error.message
            : String(error)
      },
      e && e.parameter
        ? e.parameter.callback
        : ''
    );
  }
}


function neoChronoParsePostV5_(e) {
  try {
    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {
      return {};
    }

    return JSON.parse(e.postData.contents);
  } catch (error) {
    return {};
  }
}


function neoChronoPtoWriteApiV5_(e, alreadyParsed) {
  try {
    var body =
      alreadyParsed ||
      neoChronoParsePostV5_(e);

    if (String(body.action || '') !== 'neochronoPtoWrite') {
      throw new Error('Invalid write action.');
    }

    var payload = body.payload || {};
    var operation = String(payload.operation || '').trim();

    var result;

    if (operation === 'save') {
      result = neoChronoSaveRequestV5_(payload);
    } else if (operation === 'review') {
      result = neoChronoReviewRequestV5_(payload);
    } else if (operation === 'delete') {
      result = neoChronoDeleteRequestV5_(payload);
    } else {
      throw new Error(
        'Unknown PTO write operation: ' + operation
      );
    }

    return ContentService
      .createTextOutput(
        JSON.stringify(result)
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  } catch (error) {
    return ContentService
      .createTextOutput(
        JSON.stringify({
          success: false,
          message:
            error && error.message
              ? error.message
              : String(error)
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }
}


function neoChronoResponseV5_(payload, callback) {
  var json = JSON.stringify(payload);
  var cb = String(callback || '').trim();

  if (
    cb &&
    /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(cb)
  ) {
    return ContentService
      .createTextOutput(
        cb + '(' + json + ');'
      )
      .setMimeType(
        ContentService.MimeType.JAVASCRIPT
      );
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/**
 * KEEP THIS doPost(e) only when the project does not already have doPost(e).
 * If there is already a doPost(e), merge the route from the file header into it.
 */
function doPost(e) {
  var body = neoChronoParsePostV5_(e);

  if (
    body &&
    body.action === 'neochronoPtoWrite'
  ) {
    return neoChronoPtoWriteApiV5_(e, body);
  }

  return ContentService
    .createTextOutput(
      JSON.stringify({
        success: false,
        message: 'Unknown POST action.'
      })
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
