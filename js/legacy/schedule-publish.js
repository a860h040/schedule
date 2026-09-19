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
    var frameName = 'neoScheduleReceiver_' + transferId.replace(/[^A-Za-z0-9_]/g,'_');
    var iframe = document.createElement('iframe');
    var form = document.createElement('form');
    var input = document.createElement('input');
    var finished = false;

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

    function cleanup() {
      window.removeEventListener('message', onMessage);
      try { form.remove(); } catch (_e) {}
      setTimeout(function() {
        try { iframe.remove(); } catch (_e) {}
      }, 250);
    }

    function finishError(err) {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      cleanup();
      reject(err instanceof Error ? err : new Error(String(err)));
    }

    function onMessage(event) {
      var data = event && event.data;
      if (!data || data.type !== 'NEOCHRONO_SCHEDULE_RECEIVER') return;
      if (String(data.transferId || '') !== transferId) return;

      if (finished) return;
      finished = true;
      clearTimeout(timer);
      cleanup();

      if (!data.ok) {
        reject(new Error(data.message || 'Google Sheet receiver reported an error.'));
        return;
      }

      resolve(data);
    }

    window.addEventListener('message', onMessage);

    var timer = setTimeout(function() {
      if (finished) return;
      finished = true;
      cleanup();

      // The Apps Script receiver can successfully write the Schedule sheet
      // even when a cross-origin iframe cannot post its receipt back to
      // GitHub Pages. Do not turn that browser limitation into a false error.
      resolve({
        ok: true,
        submitted: true,
        confirmationReceived: false,
        transferId: transferId,
        receiverSheetName: PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME
      });
    }, 3000);

    document.body.appendChild(iframe);
    document.body.appendChild(form);

    try {
      form.submit();
    } catch (e) {
      finishError(e);
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

  if (
    receipt &&
    receipt.confirmationReceived !== false &&
    (
      Number(receipt.transferredRows) !== prepared.rows.length ||
      String(receipt.receiverSheetName || '') !== PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME
    )
  ) {
    throw new Error(
      'Google returned a transfer receipt that did not match the schedule sent. ' +
      'NeoChrono sent ' + prepared.rows.length +
      ' rows and Google reported ' + Number(receipt.transferredRows || 0) + '.'
    );
  }

  if (typeof audit_ === 'function') {
    audit_(
      'SEND_SCHEDULE_TO_GOOGLE',
      prepared.startDate + ' through ' + prepared.endDate,
      '', '', '', '', '',
      'No', '',
      'Sent ' + prepared.rows.length +
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
      'Schedule was sent to the Google Sheet tab "' +
      PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME + '".'
  };
}
