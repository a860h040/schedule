/**
 * DONOR ADD-ON v2: Publish the current donor Schedule sheet AS-IS
 * to the pharmacist-facing receiver spreadsheet.
 *
 * IMPORTANT:
 * - This file does NOT call finalizeSchedule().
 * - Validation warnings / 5-day warnings / unfilled shifts do NOT block transfer.
 * - The receiver's previous "Schedule" sheet is replaced only after the new copy succeeds.
 * - Keep your existing main Code.gs and HTML. Replace only the previous ScheduleTransfer.gs
 *   with this file (or paste this over it).
 */

const PHARMACIST_SCHEDULE_RECEIVER = Object.freeze({
  SPREADSHEET_ID: '1flTBzOIM_dbDODjC-S-DHoViAhHASEyNWInZBxab5To',
  SHEET_NAME: 'Schedule'
});

/**
 * Called by the existing donor button:
 * "Finalize & Send to Pharmacists Schedule"
 *
 * The function name is intentionally unchanged so the current HTML keeps working.
 */
function finalizeAndSendToPharmacistsSchedule(token, startDate, endDate) {
  const ctx = requireAdmin_(token);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const donorSs = getDb_();
    const donorSheetName =
      (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.SCHEDULE)
        ? APP.SHEETS.SCHEDULE
        : 'Schedule';

    const donorSheet = donorSs.getSheetByName(donorSheetName);
    if (!donorSheet) {
      throw new Error('Donor Schedule sheet was not found.');
    }

    const donorLastRow = donorSheet.getLastRow();
    const donorLastColumn = donorSheet.getLastColumn();
    if (donorLastRow < 1 || donorLastColumn < 1) {
      throw new Error('The donor Schedule sheet is empty. Nothing was sent.');
    }

    const receiverSs = SpreadsheetApp.openById(
      PHARMACIST_SCHEDULE_RECEIVER.SPREADSHEET_ID
    );

    if (donorSs.getId() === receiverSs.getId()) {
      throw new Error(
        'The donor and receiver spreadsheet IDs are the same. The receiver must be a different Google Sheet.'
      );
    }

    // Copy FIRST. This preserves the old published schedule if the copy fails.
    const newSheet = donorSheet.copyTo(receiverSs);
    const temporaryName =
      '__NEW_SCHEDULE_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/New_York', 'yyyyMMdd_HHmmss');
    newSheet.setName(temporaryName);
    SpreadsheetApp.flush();

    // Basic verification before removing the old receiver schedule.
    if (newSheet.getLastRow() !== donorLastRow || newSheet.getLastColumn() !== donorLastColumn) {
      try { receiverSs.deleteSheet(newSheet); } catch (ignore) {}
      throw new Error(
        'Transfer verification failed before replacing the old receiver schedule. The old receiver Schedule was kept.'
      );
    }

    // Remove the PREVIOUS receiver Schedule only after the new copy exists and is verified.
    const oldSheet = receiverSs.getSheetByName(
      PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME
    );

    if (oldSheet && oldSheet.getSheetId() !== newSheet.getSheetId()) {
      receiverSs.deleteSheet(oldSheet);
    }

    // Rename the new copy to the official receiver sheet name.
    newSheet.setName(PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME);
    receiverSs.setActiveSheet(newSheet);
    SpreadsheetApp.flush();

    // Final verification.
    const verifySheet = receiverSs.getSheetByName(
      PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME
    );

    if (!verifySheet || verifySheet.getSheetId() !== newSheet.getSheetId()) {
      throw new Error(
        'Transfer verification failed. The receiver Schedule sheet was not created correctly.'
      );
    }

    if (
      verifySheet.getLastRow() !== donorLastRow ||
      verifySheet.getLastColumn() !== donorLastColumn
    ) {
      throw new Error(
        'Transfer completed, but row/column verification did not match the donor Schedule sheet.'
      );
    }

    const receiverUrl =
      'https://docs.google.com/spreadsheets/d/' +
      PHARMACIST_SCHEDULE_RECEIVER.SPREADSHEET_ID +
      '/edit#gid=' +
      verifySheet.getSheetId();

    const transferredRows = Math.max(0, donorLastRow - 1);

    // Audit in donor if the main scheduler exposes audit_().
    if (typeof audit_ === 'function') {
      let dateText = '';
      try {
        const s = startDate ? formatDateKey_(startOfDay_(asDate_(startDate))) : '';
        const e = endDate ? formatDateKey_(startOfDay_(asDate_(endDate))) : '';
        dateText = s && e ? s + ' through ' + e : '';
      } catch (ignore) {}

      audit_(
        'SEND_SCHEDULE_AS_IS',
        dateText,
        '', '', '', '', '',
        'No', '',
        'Published donor Schedule sheet AS-IS to receiver spreadsheet ' +
          PHARMACIST_SCHEDULE_RECEIVER.SPREADSHEET_ID +
          '. Validation warnings did not block transfer. Rows sent: ' +
          transferredRows + '.',
        ctx.username
      );
    }

    return {
      ok: true,
      transferredRows: transferredRows,
      transferredColumns: donorLastColumn,
      receiverSpreadsheetId: PHARMACIST_SCHEDULE_RECEIVER.SPREADSHEET_ID,
      receiverSpreadsheetName: receiverSs.getName(),
      receiverSheetName: PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME,
      receiverUrl: receiverUrl,
      validationWasRequired: false,
      message:
        'Current Schedule sheet was sent as-is to the pharmacists schedule. Existing validation warnings did not block the transfer.'
    };
  } catch (err) {
    throw new Error(
      'Send failed: ' + (err && err.message ? err.message : String(err))
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Optional donor-side test. Run manually if you want to verify receiver access.
 */
function testPharmacistScheduleReceiverConnection() {
  const ss = SpreadsheetApp.openById(
    PHARMACIST_SCHEDULE_RECEIVER.SPREADSHEET_ID
  );

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    hasScheduleSheet: !!ss.getSheetByName(
      PHARMACIST_SCHEDULE_RECEIVER.SHEET_NAME
    ),
    url: ss.getUrl()
  };
}