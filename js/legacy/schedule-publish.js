/**
 * NeoChrono -> Pharmacists Google Sheet publisher
 *
 * Sends the selected schedule period from the GitHub-backed NeoChrono app
 * to the Google Apps Script receiver. The receiver writes the data into the
 * sheet named exactly "Schedule".
 */

const PHARMACIST_SCHEDULE_RECEIVER = Object.freeze({
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec',
  SHEET_NAME: 'Schedule',
  HEADERS: Object.freeze([
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
  ])
});

function pharmacistSchedulePublishDateKey_(value) {
  if (value === null || value === undefined || value === '') return '';

  if (value instanceof Date && !isNaN(value.getTime())) {
    return [
      value.getFullYear(),
      String(value.getMonth() + 1).padStart(2,'0'),
      String(value.getDate()).padStart(2,'0')
    ].join('-');
  }

  var raw = String(value).trim();
  var iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];

  try {
    var d = asDate_(value);
    return d ? formatDateKey_(startOfDay_(d)) : '';
  } catch (_e) {
    return '';
  }
}

function pharmacistSchedulePublishCell_(header, value) {
  if (value === null || value === undefined) return '';

  if (header === 'Date') {
    return pharmacistSchedulePublishDateKey_(value);
  }

  if (
    header === 'Updated At' ||
    header === 'Finalized At'
  ) {
    if (value instanceof Date && !isNaN(value.getTime())) {
      return value.toISOString();
    }
    return String(value || '');
  }

  return value;
}

function pharmacistSchedulePayloadRows_(startDate, endDate) {
  var startKey = pharmacistSchedulePublishDateKey_(startDate);
  var endKey = pharmacistSchedulePublishDateKey_(endDate);

  if (!startKey || !endKey) {
    throw new Error('Start Date and End Date are required.');
  }
  if (startKey > endKey) {
    throw new Error('Start Date must be before or equal to End Date.');
  }

  var headers = PHARMACIST_SCHEDULE_RECEIVER.HEADERS.slice();
  var rows = readTable_(APP.SHEETS.SCHEDULE)
    .filter(function(r) {
      var dk = pharmacistSchedulePublishDateKey_(r.Date);
      return dk && dk >= startKey && dk <= endKey;
    })
    .map(function(r) {
      return headers.map(function(h) {
        return pharmacistSchedulePublishCell_(h, r[h]);
      });
    });

  return {
    startDate: startKey,
    endDate: endKey,
    headers: headers,
    rows: rows
  };
}

/**
 * Cross-origin POST using a normal HTML form targeted at a hidden iframe.
 * Code4.gs posts a verified result back to the parent window with postMessage.
 */
function pharmacistSchedulePostToGoogle_(payload) {
  return new Promise(function(resolve, reject) {
    if (typeof document === 'undefined' || typeof window === 'undefined') {
      reject(new Error('Google Sheet transfer requires the NeoChrono browser app.'));
      return;
    }

    var transferId = String(payload.transferId || '');
    var frameName =
      'neoScheduleReceiver_' +
      transferId.replace(/[^A-Za-z0-9_]/g,'_');

    var iframe = document.createElement('iframe');
    var form = document.createElement('form');
    var input = document.createElement('input');

    iframe.name = frameName;
    iframe.style.display = 'none';
    iframe.setAttribute('aria-hidden','true');

    form.method = 'POST';
    form.action = PHARMACIST_SCHEDULE_RECEIVER.WEB_APP_URL;
    form.target = frameName;
    form.style.display = 'none';

    input.type = 'hidden';
    input.name = 'payload';
    input.value = JSON.stringify(payload);
    form.appendChild(input);

    document.body.appendChild(iframe);
    document.body.appendChild(form);

    try {
      form.submit();

      // Fire-and-forget by design.
      // The user does not want NeoChrono to wait for or require a receipt
      // from Google Apps Script. A successful browser form submission is
      // treated as a successful send.
      resolve({
        ok: true,
        submitted: true,
        confirmationReceived: false,
        transferId: transferId,
        receiverSheetName: PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME
      });

      // Leave the iframe/form in place briefly so the browser has time to
      // complete the POST before cleanup.
      setTimeout(function() {
        try { form.remove(); } catch (_e) {}
        try { iframe.remove(); } catch (_e) {}
      }, 8000);

    } catch (e) {
      try { form.remove(); } catch (_e) {}
      try { iframe.remove(); } catch (_e) {}
      reject(e);
    }
  });
}

/**
 * Called by the dashboard button "Finalize & Send".
 */
async function finalizeAndSendToPharmacistsSchedule(token, startDate, endDate) {
  var ctx = requireAdmin_(token);
  var prepared = pharmacistSchedulePayloadRows_(startDate, endDate);

  if (!prepared.rows.length) {
    throw new Error(
      'There are no Schedule rows from ' +
      prepared.startDate + ' through ' + prepared.endDate + '.'
    );
  }

  var transferId =
    'NEO-' +
    Date.now() +
    '-' +
    String(Utilities.getUuid()).slice(0,8).toUpperCase();

  var payload = {
    action: 'replaceSchedule',
    transferId: transferId,
    sheetName: PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME,
    headers: prepared.headers,
    rows: prepared.rows,
    startDate: prepared.startDate,
    endDate: prepared.endDate,
    sentAt: new Date().toISOString(),
    sentBy: ctx.username || ''
  };

  var receipt = await pharmacistSchedulePostToGoogle_(payload);

  if (typeof audit_ === 'function') {
    audit_(
      'SEND_SCHEDULE_TO_GOOGLE',
      prepared.startDate + ' through ' + prepared.endDate,
      '', '', '', '', '',
      'No', '',
      'Submitted ' + prepared.rows.length +
      ' Schedule row(s) to Google receiver sheet "' +
      PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME +
      '". Transfer ID: ' + transferId + '.',
      ctx.username
    );
  }

  return {
    ok: true,
    transferId: transferId,
    transferredRows: prepared.rows.length,
    transferredColumns: prepared.headers.length,
    receiverSpreadsheetId: receipt.receiverSpreadsheetId || '',
    receiverSpreadsheetName: receipt.receiverSpreadsheetName || 'Pharmacists Schedule',
    receiverSheetName: receipt.receiverSheetName || PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME,
    receiverUrl: receipt.receiverUrl || '',
    confirmationReceived: receipt.confirmationReceived !== false,
    message:
      'Schedule was submitted to the Google Sheet receiver for tab "' +
      PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME + '".'
  };
}
