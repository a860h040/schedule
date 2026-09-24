/**
 * Code11.gs — Date-aware Preceptor Calendar
 *
 * ADD this as a NEW .gs file in the SAME Apps Script project.
 * Do not replace Code.gs.
 *
 * What this adds
 * 1) A new Google Sheet named "Preceptor Calendar".
 * 2) 12 month rows per preceptor/year with Week 1–Week 6 Yes/No switches.
 * 3) Admin APIs used by the new Preceptor Calendar page in Index.html.
 * 4) A date-aware replacement for preceptorEveningBlocked_().
 *
 * IMPORTANT DESIGN
 * - Users -> Preceptor = Yes means the pharmacist IS CAPABLE of precepting.
 * - Preceptor Calendar controls WHEN that pharmacist is ACTIVELY precepting.
 * - Preceptor=Yes is a weekday teaching-coverage rule: the pharmacist must
 *   be scheduled on a Day/Morning shift when working Monday-Friday.
 * - If a Preferred / Home Unit is configured, weekday assignments stay there.
 *   Without a preferred unit, any qualified Day/Morning shift may be used.
 * - Preceptor Calendar ON/OFF remains visible for planning/history, but OFF no
 *   longer rotates a preceptor into weekday evening shifts.
 * - New calendar rows default to Yes for backward-compatible display behavior.
 *
 * Week boundaries follow the scheduler's Sunday–Saturday week. The first and
 * last week displayed in a month may therefore be partial calendar weeks.
 */

var PRECEPTOR_CALENDAR_SHEET_ = 'Preceptor Calendar';
var PRECEPTOR_CALENDAR_HEADERS_ = [
  'Calendar ID','Employee ID','Pharmacist Name','Username','Year','Month','Month Number',
  'Week 1','Week 2','Week 3','Week 4','Week 5','Week 6',
  'Active','Notes','Updated At','Updated By'
];
var _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = null;

/** Optional manual setup/repair function. The UI also creates the sheet automatically. */
function setupPreceptorCalendar(year) {
  var y = preceptorCalendarNormalizeYear_(year || new Date().getFullYear());
  var sh = preceptorCalendarEnsureSheet_();
  var sync = preceptorCalendarEnsureRowsForYear_(y, 'SYSTEM');
  return {
    ok:true,
    sheet:PRECEPTOR_CALENDAR_SHEET_,
    year:y,
    rowsAdded:sync.rowsAdded,
    rowsRepaired:sync.rowsRepaired,
    spreadsheetId:getDb_().getId(),
    message:'Preceptor Calendar is ready for '+y+'. New preceptor weeks default to Yes.'
  };
}

/** Admin page data. */
function getPreceptorCalendarData(token, year) {
  var ctx = requireAdmin_(token);
  var y = preceptorCalendarNormalizeYear_(year || new Date().getFullYear());
  preceptorCalendarEnsureSheet_();
  preceptorCalendarEnsureRowsForYear_(y, ctx.username);

  var users = readTable_(APP.SHEETS.USERS)
    .filter(function(u){ return clean_(u.Role).toLowerCase() !== 'administrator' && yes_(u.Preceptor); })
    .sort(function(a,b){ return clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name'])); });

  var rows = preceptorCalendarReadRows_().filter(function(r){
    return Number(r.Year) === y && yesDefault_(r.Active,true);
  });

  var byKey = {};
  rows.forEach(function(r){
    byKey[preceptorCalendarEmployeeKeyFromRow_(r)+'|'+Number(r['Month Number'])] = r;
  });

  var calendar = [];
  users.forEach(function(u){
    var employeeKey = preceptorCalendarEmployeeKeyFromUser_(u);
    for (var month=1; month<=12; month++) {
      var r = byKey[employeeKey+'|'+month] || {};
      var segments = preceptorCalendarMonthSegments_(y,month);
      calendar.push({
        employeeId:clean_(u['Employee ID']),
        pharmacistName:clean_(u['Pharmacist Name']),
        username:clean_(u.Username),
        year:y,
        monthNumber:month,
        month:preceptorCalendarMonthName_(month),
        weeks:segments.map(function(seg){
          var raw = r['Week '+seg.weekNumber];
          // Blank values preserve the legacy/static Preceptor=Yes behavior.
          var on = clean_(raw)==='' ? true : yes_(raw);
          return {
            weekNumber:seg.weekNumber,
            start:preceptorCalendarDateKey_(seg.start),
            end:preceptorCalendarDateKey_(seg.end),
            label:preceptorCalendarRangeLabel_(seg.start,seg.end),
            precepting:on
          };
        })
      });
    }
  });

  return {
    ok:true,
    year:y,
    sheetName:PRECEPTOR_CALENDAR_SHEET_,
    spreadsheetUrl:getDb_().getUrl(),
    preceptors:users.map(function(u){
      return {
        employeeId:clean_(u['Employee ID']),
        pharmacistName:clean_(u['Pharmacist Name']),
        username:clean_(u.Username),
        active:yes_(u.Active),
        resident:yes_(u.Resident),
        weekendGroup:clean_(u['Weekend Group'])
      };
    }),
    calendar:calendar,
    weekDefinition:'Sunday–Saturday. Partial first/last calendar weeks are shown separately.',
    defaultBehavior:'If a calendar row is missing or blank, Preceptor=Yes keeps the legacy preceptor restriction ON.'
  };
}

/** Save all 12 months for one pharmacist/year in one request. */
function savePreceptorCalendarYear(token, employeeId, year, monthStates) {
  var ctx = requireAdmin_(token);
  var y = preceptorCalendarNormalizeYear_(year);
  var eid = clean_(employeeId);
  if (!eid) throw new Error('Employee ID is required.');

  var user = readTable_(APP.SHEETS.USERS).find(function(u){
    return clean_(u['Employee ID']) === eid && clean_(u.Role).toLowerCase() !== 'administrator';
  });
  if (!user) throw new Error('Pharmacist was not found.');
  if (!yes_(user.Preceptor)) throw new Error(clean_(user['Pharmacist Name'])+' is not marked Preceptor = Yes in Users.');

  preceptorCalendarEnsureSheet_();
  preceptorCalendarEnsureRowsForYear_(y, ctx.username);

  var statesByMonth = {};
  (Array.isArray(monthStates) ? monthStates : []).forEach(function(m){
    var month = Number(m.monthNumber);
    if (month < 1 || month > 12) return;
    statesByMonth[month] = Array.isArray(m.weeks) ? m.weeks : [];
  });

  var sh = getDb_().getSheetByName(PRECEPTOR_CALENDAR_SHEET_);
  var values = sh.getDataRange().getValues();
  if (!values.length) throw new Error('Preceptor Calendar sheet is empty.');
  var h = {};
  values[0].forEach(function(name,i){ h[clean_(name)] = i; });
  var changed = 0;

  for (var r=1; r<values.length; r++) {
    var row = values[r];
    if (clean_(row[h['Employee ID']]) !== eid) continue;
    if (Number(row[h['Year']]) !== y) continue;
    if (!yesDefault_(row[h['Active']],true)) continue;

    var monthNo = Number(row[h['Month Number']]);
    var states = statesByMonth[monthNo];
    if (!states) continue;

    var validSegments = preceptorCalendarMonthSegments_(y,monthNo);
    var validWeekCount = validSegments.length;
    for (var w=1; w<=6; w++) {
      var col = h['Week '+w];
      if (col === undefined) continue;
      if (w > validWeekCount) {
        row[col] = '';
      } else {
        var stateObj = states.find(function(x){ return Number(x.weekNumber) === w; });
        row[col] = stateObj && stateObj.precepting === false ? 'No' : 'Yes';
      }
    }
    row[h['Pharmacist Name']] = clean_(user['Pharmacist Name']);
    row[h['Username']] = clean_(user.Username);
    row[h['Updated At']] = new Date();
    row[h['Updated By']] = ctx.username;
    changed++;
  }

  sh.getDataRange().setValues(values);
  SpreadsheetApp.flush();
  _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = null;

  audit_(
    'PRECEPTOR_CALENDAR_UPDATE',
    String(y),
    clean_(user['Pharmacist Name']),
    '', '', '',
    changed+' month row(s) updated',
    'No','',
    'Updated date-aware preceptor calendar for '+clean_(user['Pharmacist Name'])+' ('+eid+').',
    ctx.username
  );

  return {
    ok:true,
    employeeId:eid,
    pharmacistName:clean_(user['Pharmacist Name']),
    year:y,
    updatedMonths:changed,
    message:'Preceptor calendar saved.'
  };
}

/** Adds newly-marked preceptors and repairs missing month rows for a year. */
function syncPreceptorCalendar(token, year) {
  var ctx = requireAdmin_(token);
  var y = preceptorCalendarNormalizeYear_(year || new Date().getFullYear());
  preceptorCalendarEnsureSheet_();
  var result = preceptorCalendarEnsureRowsForYear_(y,ctx.username);
  _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = null;
  return {
    ok:true,
    year:y,
    rowsAdded:result.rowsAdded,
    rowsRepaired:result.rowsRepaired,
    message:'Preceptor Calendar synchronized with Users.'
  };
}

/**
 * DATE-AWARE PRECEPTOR MODE ALGORITHM
 *
 * Correct behavior:
 *
 *   Calendar ON  = pharmacist IS precepting
 *                  -> weekday work is restricted to the pharmacist's
 *                     exact Preferred Shift Type / preferred unit.
 *
 *   Calendar OFF = pharmacist is NOT precepting
 *                  -> weekday work may be ANY active Evening-type shift
 *                     OR the pharmacist's Skills → Preferred / Home Unit.
 *
 * Saturday/Sunday keep the existing weekend rotation logic.
 *
 * Code.gs is not edited. Code11 wraps the existing scheduler helpers.
 */

var PRECEPTOR_CALENDAR_EVENING_CODES_ = ['E1','E2'];

function preceptorCalendarEveningCode_(code) {
  return PRECEPTOR_CALENDAR_EVENING_CODES_.indexOf(clean_(code).toUpperCase()) >= 0;
}

function preceptorCalendarPreferredUnitCodes_(u, model) {
  if (!u || !model || !model.shiftMap) return [];

  var username = clean_(u.Username);
  var employeeId = clean_(u['Employee ID']);
  var out = [];

  // New source of truth: Employee Skills -> Preferred = Yes.
  (model.skills || []).forEach(function(r) {
    if (!yesDefault_(r.Active,true)) return;
    if (!yes_(r.Preferred)) return;

    var sameEmployee =
      (employeeId && clean_(r['Employee ID']) === employeeId) ||
      (username && clean_(r.Username) === username);

    if (!sameEmployee) return;

    var code = clean_(r.Skill).toUpperCase();
    if (
      code &&
      model.shiftMap[code] &&
      yesDefault_(model.shiftMap[code].Active,true) &&
      out.indexOf(code) < 0
    ) {
      out.push(code);
    }
  });

  if (out.length) return out;

  // Backward-compatible fallback for older records not migrated yet.
  var raw = clean_(u['Preferred Shift Type']).toUpperCase();
  if (!raw) return [];

  raw
    .split(/[\\/,;|]+/)
    .map(function(x){ return clean_(x).toUpperCase(); })
    .filter(Boolean)
    .forEach(function(code) {
      if (
        model.shiftMap[code] &&
        yesDefault_(model.shiftMap[code].Active,true) &&
        out.indexOf(code) < 0
      ) {
        out.push(code);
      }
    });

  return out;
}

function preceptorCalendarPreferredUnitMatches_(u, shiftOrSlot, model) {
  var codes = preceptorCalendarPreferredUnitCodes_(u,model);
  if (!codes.length) return false;

  var code = clean_(
    shiftOrSlot && (
      shiftOrSlot.shiftCode ||
      shiftOrSlot.Shift ||
      (shiftOrSlot.shift && shiftOrSlot.shift.Shift)
    )
  ).toUpperCase();

  return codes.indexOf(code) >= 0;
}

function preceptorCalendarIsManagedWeekday_(u,date) {
  if (!u || !yes_(u.Preceptor) || isSevenOn_(u)) return false;
  var d = startOfDay_(asDate_(date));
  if (!d || isWeekendDate_(d)) return false;
  return true;
}

function preceptorCalendarModeForDate_(u,date) {
  if (!preceptorCalendarIsManagedWeekday_(u,date)) return 'NORMAL';
  return preceptorCalendarIsActiveOnDate_(u,date) ? 'ON' : 'OFF';
}

/**
 * Returns true when this pharmacist is actively precepting during the
 * Sunday-Saturday week containing date. We look at the weekdays in that week
 * because precepting is a weekday activity, but the resulting evening block
 * applies to the entire week, including Saturday/Sunday.
 */
function preceptorCalendarIsPreceptingWeek_(u,date) {
  if (!u || !yes_(u.Preceptor) || isSevenOn_(u)) return false;

  var d = startOfDay_(asDate_(date));
  if (!d) return false;

  var sunday = addDays_(d,-d.getDay());

  for (var i=1;i<=5;i++) {
    var weekday = addDays_(sunday,i);
    if (preceptorCalendarIsActiveOnDate_(u,weekday)) return true;
  }

  return false;
}

function preceptorCalendarIsEveningSlot_(slot,model) {
  if (!slot) return false;

  var code = clean_(
    slot.shiftCode ||
    slot.Shift ||
    (slot.shift && slot.shift.Shift)
  ).toUpperCase();

  if (preceptorCalendarEveningCode_(code)) return true;

  var shift =
    slot.shift ||
    (model && model.shiftMap && model.shiftMap[code]) ||
    null;

  return !!(
    shift &&
    clean_(shift.Type).toLowerCase() === 'evening'
  );
}

function preceptorCalendarIsDaySlot_(slot,model) {
  if (!slot) return false;

  var code = clean_(
    slot.shiftCode ||
    slot.Shift ||
    (slot.shift && slot.shift.Shift)
  ).toUpperCase();

  var shift =
    slot.shift ||
    (model && model.shiftMap && model.shiftMap[code]) ||
    null;

  if (!shift) return false;

  var type = clean_(shift.Type).toLowerCase();
  return type === 'day' || type === 'morning';
}

/**
 * Preceptor evening restriction:
 * ON/precepting week -> every Evening-type shift is blocked.
 * OFF/non-precepting week -> Evening-type shifts may be used, subject to
 * the hard monthly maximum of 5 total Evening shifts.
 */
function preceptorCalendarAwarePreceptorEveningBlocked_(u,slot,model) {
  if (!u || !slot || !yes_(u.Preceptor)) return false;
  if (!preceptorCalendarIsEveningSlot_(slot,model)) return false;

  // v7: Preceptor=Yes pharmacists teach students during weekday work.
  // Therefore Monday-Friday evening shifts are never automatic options for
  // a preceptor. Weekend evening eligibility continues to follow the existing
  // weekend rule because the student/precepting requirement is weekday based.
  if (!isWeekendDate_(slot.date)) return true;

  return !model.settings.preceptorWeekendEveningAllowed;
}

/**
 * Hard eligibility:
 *
 * ON weekday:
 *   pharmacist may work only the exact preferred/home-unit shift code(s).
 *
 * OFF weekday:
 *   pharmacist may work any active Evening-type shift or the Preferred/Home Unit.
 *
 * If ON but Preferred Shift Type does not contain an exact active shift code,
 * we do not invent a unit. Existing eligibility remains in place and
 * validation reports the missing setup.
 */
function preceptorCalendarAwareEligibilityReason_(u,shift,date,model) {
  var baseReason = _PC11_BASE_RESIDENT_ELIGIBILITY_REASON_
    ? _PC11_BASE_RESIDENT_ELIGIBILITY_REASON_(u,shift,date,model)
    : '';

  if (baseReason) return baseReason;
  if (!preceptorCalendarIsManagedWeekday_(u,date)) return '';

  var code = clean_(shift && shift.Shift).toUpperCase();
  var preferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);
  var slot = {shiftCode:code,shift:shift};

  // Every weekday assignment for a preceptor must be a morning/day shift.
  if (!preceptorCalendarIsDaySlot_(slot,model)) {
    return 'PRECEPTOR_WEEKDAY_DAY_SHIFT_ONLY';
  }

  // If a Preferred / Home Unit is configured, keep the preceptor there so the
  // student stays with the pharmacist in the intended teaching unit.
  if (preferredCodes.length && preferredCodes.indexOf(code) < 0) {
    return 'PRECEPTOR_PREFERRED_UNIT_ONLY';
  }

  // No preferred unit configured: any qualified Day/Morning shift is allowed.
  return '';
}

/**
 * Very strong score support in case a preferred slot reaches the general
 * allocator rather than being handled by the preassignment pass.
 */
function preceptorCalendarAwareScoreCandidate_(u,slot,model,state,elig) {
  var score = _PC11_BASE_SCORE_CANDIDATE_
    ? _PC11_BASE_SCORE_CANDIDATE_(u,slot,model,state,elig)
    : 0;

  if (!preceptorCalendarIsManagedWeekday_(u,slot && slot.date)) return score;

  var preferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);

  if (preferredCodes.length && preceptorCalendarPreferredUnitMatches_(u,slot,model)) {
    score += 300000;
  } else if (!preferredCodes.length && preceptorCalendarIsDaySlot_(slot,model)) {
    // A preceptor without a configured home unit still needs a weekday
    // morning/day assignment. Normal skill priority decides which one.
    score += 250000;
  }

  return score;
}

/**
 * Assign actively-precepting pharmacists into their preferred/home unit
 * before the general allocator runs.
 */
function preceptorCalendarPreassignOnPreferredUnits_(
  slots, results, model, state, actor, reserved, start, end
) {
  if (!slots || !slots.length || !model || !state || !reserved) return;

  var users = (model.users || [])
    .filter(function(u){
      return yes_(u.Active) &&
             yes_(u.Preceptor) &&
             !isSevenOn_(u);
    })
    .sort(function(a,b){
      return clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name']));
    });

  if (!users.length) return;

  var rangeStart = startOfDay_(asDate_(start));
  var rangeEnd = startOfDay_(asDate_(end));
  if (!rangeStart || !rangeEnd) return;

  for (
    var day = new Date(rangeStart);
    day.getTime() <= rangeEnd.getTime();
    day = addDays_(day,1)
  ) {
    if (isWeekendDate_(day)) continue;

    var dk = formatDateKey_(day);

    users.forEach(function(u){
      var username = clean_(u.Username);
      if (state.byEmployeeDate[username+'|'+dk]) return;

      // PTO / approved Regular Off remain legitimate protected off days.
      if (isBlockedByPto_(u,day,model) || isBlockedByRegularOff_(u,day,model)) return;

      var preferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);

      var candidates = slots
        .filter(function(s){
          if (!s || s.dateKey !== dk || reserved.has(s.slotKey)) return false;
          if (!preceptorCalendarIsDaySlot_(s,model)) return false;

          var code = clean_(s.shiftCode).toUpperCase();
          if (preferredCodes.length && preferredCodes.indexOf(code) < 0) return false;

          var e = eligibility_(u,s,model,state,false);
          return !!(e && e.ok);
        })
        .sort(function(a,b){
          var ap = num_(a.shift && a.shift.Priority,50);
          var bp = num_(b.shift && b.shift.Priority,50);
          return ap-bp ||
                 clean_(a.shiftCode).localeCompare(clean_(b.shiftCode)) ||
                 a.slot-b.slot;
        });

      if (!candidates.length) return;

      var assigned = assignSpecificUser_(
        candidates[0],
        u,
        model,
        state,
        actor,
        preferredCodes.length
          ? 'PRECEPTOR WEEKDAY DAY — PREFERRED UNIT'
          : 'PRECEPTOR WEEKDAY DAY — MORNING/DAY SHIFT'
      );

      if (assigned) {
        results.push(assigned);
        reserved.add(candidates[0].slotKey);
      }
    });
  }
}


/* =====================================================================
   v6 — MONTHLY PRECEPTOR EVENING ROTATION + HOME-UNIT COVERAGE
   =====================================================================

   Monthly behavior for Preceptor=Yes pharmacists:

   1. If every weekday in the calendar month is Preceptor Calendar ON:
        evening rotation target = 0.

   2. If one or more weekdays in the calendar month are OFF AND belong to a
      fully non-precepting Sunday-Saturday week:
        target = 5 Evening-type shifts for that calendar month.

   3. A preceptor may work a maximum of 5 total Evening-type shifts in that month.
      Shift #6 and higher are not allowed, even as fallback coverage.

   4. A pharmacist cannot work ANY evening shift during a Sunday-Saturday
      week in which they are actively precepting. Evening rotation shifts can
      be placed only during non-precepting weeks and only on calendar-OFF
      weekdays.

   5. On each planned Evening rotation date, another pharmacist with the NORMAL
      Employee Skill for the preceptor's Preferred/Home Unit is assigned to
      cover that home-unit shift when the unit is required that day.

   6. Existing hard rules still apply:
      PTO, weekly availability, one shift/day, five-day week, hours,
      evening-to-morning transition, resident rules, weekend rules, etc.
*/

var PRECEPTOR_MONTHLY_EVENING_TARGET_ = 5;
var PRECEPTOR_MONTHLY_EVENING_MAX_ = 5;
var PRECEPTOR_ROTATION_COVERAGE_REASON_ = 'PRECEPTOR EVENING ROTATION';

function preceptorCalendarMonthHasOffWeekday_(u, monthDate) {
  if (!u || !yes_(u.Preceptor)) return false;

  var first = firstOfMonth_(monthDate);
  var last = lastOfMonth_(monthDate);

  for (var d = new Date(first); d.getTime() <= last.getTime(); d = addDays_(d,1)) {
    if (isWeekendDate_(d)) continue;

    // Only an OFF weekday in a fully non-precepting Sunday-Saturday week can
    // support the monthly Evening-shift rotation target.
    if (
      !preceptorCalendarIsActiveOnDate_(u,d) &&
      !preceptorCalendarIsPreceptingWeek_(u,d)
    ) {
      return true;
    }
  }
  return false;
}

function preceptorCalendarMonthlyEveningMaximum_(u, monthDate, model) {
  if (!u || !model || !yes_(u.Preceptor) || isSevenOn_(u)) return 0;
  if (!preceptorCalendarMonthHasOffWeekday_(u,monthDate)) return 0;

  // Five is the absolute preceptor Evening-shift ceiling. A lower pharmacist-specific
  // Maximum Evening Shifts Per Month remains a stricter hard limit.
  var employeeMax = Math.floor(employeeEveningMax_(u,model));
  if (employeeMax <= 0) return 0;
  return Math.min(PRECEPTOR_MONTHLY_EVENING_MAX_, employeeMax);
}

function preceptorCalendarMonthlyEveningTarget_(u, monthDate, model) {
  var maxAllowed = preceptorCalendarMonthlyEveningMaximum_(u,monthDate,model);
  if (maxAllowed <= 0) return 0;

  // Normal monthly goal is five, not seven.
  return Math.min(PRECEPTOR_MONTHLY_EVENING_TARGET_, maxAllowed);
}

function preceptorCalendarCountEveningInMonth_(username, monthDate, state, model) {
  var mk = monthKey_(monthDate);
  return (state.assignments || []).filter(function(a) {
    if (
      !a ||
      a.status === 'UNFILLED' ||
      clean_(a.username) !== clean_(username) ||
      monthKey_(a.date) !== mk
    ) {
      return false;
    }

    return preceptorCalendarIsEveningSlot_(
      {
        shiftCode:a.shiftCode,
        shift:model && model.shiftMap
          ? model.shiftMap[clean_(a.shiftCode).toUpperCase()]
          : null
      },
      model
    );
  }).length;
}

// Backward-compatible alias for older helper calls.
function preceptorCalendarCountE1E2InMonth_(username, monthDate, state, model) {
  return preceptorCalendarCountEveningInMonth_(username,monthDate,state,model);
}

function preceptorCalendarOtherEligibleEveningCovererExists_(owner, slot, model, state) {
  if (!owner || !slot || !model || !state) return false;

  return (model.users || []).some(function(other) {
    if (!other || !yes_(other.Active)) return false;
    if (clean_(other.Username) === clean_(owner.Username)) return false;

    var e = eligibility_(other,slot,model,state,false);
    return !!(e && e.ok);
  });
}

function preceptorCalendarPreferredUnitSlotsForDate_(u, dateKey, slots, model) {
  var preferred = preceptorCalendarPreferredUnitCodes_(u,model);
  if (!preferred.length) return [];

  return (slots || [])
    .filter(function(s) {
      return s &&
        s.dateKey === dateKey &&
        preferred.indexOf(clean_(s.shiftCode).toUpperCase()) >= 0;
    })
    .sort(function(a,b) {
      return preferred.indexOf(clean_(a.shiftCode).toUpperCase()) -
             preferred.indexOf(clean_(b.shiftCode).toUpperCase()) ||
             a.slot-b.slot;
    });
}

function preceptorCalendarExistingHomeCoverage_(results, u, dateKey, model) {
  var preferred = preceptorCalendarPreferredUnitCodes_(u,model);
  if (!preferred.length) return null;

  return (results || []).find(function(a) {
    return a &&
      a.status !== 'UNFILLED' &&
      a.dateKey === dateKey &&
      clean_(a.username) !== clean_(u.Username) &&
      preferred.indexOf(clean_(a.shiftCode).toUpperCase()) >= 0;
  }) || null;
}

function preceptorCalendarHomeCoverageCandidates_(owner, slot, model, state) {
  var candidates = [];

  model.users.forEach(function(u) {
    if (!yes_(u.Active)) return;
    if (clean_(u.Username) === clean_(owner.Username)) return;

    // User specifically asked for pharmacists who "has the skill" to cover.
    // Use normal Employee Skills here, not coverage-only skills.
    if (!hasRequiredSkillForSlot_(u,slot,model)) return;

    var e = eligibility_(u,slot,model,state,false);
    if (!e.ok) return;

    candidates.push({
      user:u,
      elig:e,
      score:scoreCandidate_(u,slot,model,state,e) + 5000
    });
  });

  candidates.sort(function(a,b) {
    return b.score-a.score ||
      clean_(a.user['Pharmacist Name']).localeCompare(clean_(b.user['Pharmacist Name']));
  });

  return candidates;
}

function preceptorCalendarAssignHomeCoverage_(owner, unitSlot, results, model, state, actor, reserved) {
  if (!owner || !unitSlot) return null;

  var existing = preceptorCalendarExistingHomeCoverage_(
    results,
    owner,
    unitSlot.dateKey,
    model
  );

  if (existing) {
    existing.coverageForPharmacist = clean_(owner['Pharmacist Name']);
    existing.coverageForUsername = clean_(owner.Username);
    existing.coverageReason = PRECEPTOR_ROTATION_COVERAGE_REASON_;
    existing.warning = [
      clean_(existing.warning),
      'COVERING PRECEPTOR HOME UNIT WHILE PRECEPTOR WORKS EVENING'
    ].filter(Boolean).join(' | ');
    return existing;
  }

  if (reserved.has(unitSlot.slotKey)) {
    // Slot is already protected by another mandatory pass. Treat it as covered.
    var reservedAssignment = (results || []).find(function(a) {
      return a &&
        a.status !== 'UNFILLED' &&
        a.dateKey === unitSlot.dateKey &&
        clean_(a.shiftCode) === clean_(unitSlot.shiftCode) &&
        num_(a.slot,1) === num_(unitSlot.slot,1);
    });

    if (reservedAssignment && clean_(reservedAssignment.username) !== clean_(owner.Username)) {
      reservedAssignment.coverageForPharmacist = clean_(owner['Pharmacist Name']);
      reservedAssignment.coverageForUsername = clean_(owner.Username);
      reservedAssignment.coverageReason = PRECEPTOR_ROTATION_COVERAGE_REASON_;
      return reservedAssignment;
    }

    return null;
  }

  var candidates = preceptorCalendarHomeCoverageCandidates_(owner,unitSlot,model,state);
  if (!candidates.length) return null;

  var pick = candidates[0];
  var a = assignSpecificUser_(
    unitSlot,
    pick.user,
    model,
    state,
    actor,
    'PRECEPTOR EVENING ROTATION COVERAGE for '+clean_(owner['Pharmacist Name'])
  );

  if (!a) return null;

  a.coverageForPharmacist = clean_(owner['Pharmacist Name']);
  a.coverageForUsername = clean_(owner.Username);
  a.coverageReason = PRECEPTOR_ROTATION_COVERAGE_REASON_;

  results.push(a);
  reserved.add(unitSlot.slotKey);
  return a;
}

function preceptorCalendarRemoveNewCoverage_(coverage, results, model, state, reserved) {
  if (!coverage) return;

  // Only rollback coverage created by this pass. Do not remove a mandatory
  // assignment that merely received a coverage label.
  var idx = (results || []).indexOf(coverage);
  var createdByPass =
    idx >= 0 &&
    clean_(coverage.warning).indexOf('PRECEPTOR EVENING ROTATION COVERAGE') >= 0;

  if (!createdByPass) return;

  removeAssignmentFromState_(state,coverage,model);
  reserved.delete(
    slotKey_(coverage.dateKey,coverage.shiftCode,num_(coverage.slot,1))
  );
  results.splice(idx,1);
}


/**
 * Assign non-precepting pharmacists to Evening-type shifts before the general allocator.
 * Mandatory resident E2 is handled first by the original scheduler.
 */
function preceptorCalendarPreassignOffEvenings_(
  slots, results, model, state, actor, reserved, start, end
) {
  // v7: disabled. Preceptor=Yes pharmacists are no longer automatically
  // rotated into weekday evening shifts when their calendar is OFF.
  // They remain weekday Day/Morning pharmacists so students are not left
  // without a preceptor.
  return;
}

/**
 * Existing generation order calls preassignResidentE2_ after weekend/ED
 * priority work. Keep resident E2 first, then apply both preceptor modes.
 */
function preceptorCalendarAwareResidentE2Pass_(
  slots, results, model, state, actor, reserved, start, end
) {
  if (_PC11_BASE_PREASSIGN_RESIDENT_E2_) {
    _PC11_BASE_PREASSIGN_RESIDENT_E2_(
      slots, results, model, state, actor, reserved, start, end
    );
  }

  preceptorCalendarPreassignOnPreferredUnits_(
    slots, results, model, state, actor, reserved, start, end
  );

  // If the pharmacist has any OFF weekday in the month, fill the monthly
  // five-total-evening target on OFF weekdays and simultaneously cover
  // the pharmacist's preferred/home unit with another normally-skilled pharmacist.
  preceptorCalendarPreassignOffEvenings_(
    slots, results, model, state, actor, reserved, start, end
  );
}

/**
 * Full-period validation:
 * ON  -> preferred unit
 * OFF -> Evening-type shift or Preferred/Home Unit
 */
function preceptorCalendarAwareValidateGeneratedAssignments_(
  assignments, model, start, end, validationOptions
) {
  var report = _PC11_BASE_VALIDATE_GENERATED_ASSIGNMENTS_
    ? _PC11_BASE_VALIDATE_GENERATED_ASSIGNMENTS_(
        assignments, model, start, end, validationOptions
      )
    : {errors:[],warnings:[]};

  report = report || {errors:[],warnings:[]};
  report.errors = Array.isArray(report.errors) ? report.errors.slice() : [];
  report.warnings = Array.isArray(report.warnings) ? report.warnings.slice() : [];

  (assignments || [])
    .filter(function(a){
      return a && a.status !== 'UNFILLED' && a.username;
    })
    .forEach(function(a){
      var u = model && model.usersByUsername
        ? model.usersByUsername[clean_(a.username)]
        : null;

      if (!u || !yes_(u.Preceptor)) return;
      if (!preceptorCalendarIsManagedWeekday_(u,a.date)) return;

      var code = clean_(a.shiftCode).toUpperCase();
      var slot = {
        shiftCode:code,
        shift:model && model.shiftMap ? model.shiftMap[code] : null
      };
      var preferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);

      if (!preceptorCalendarIsDaySlot_(slot,model)) {
        report.errors.push(
          a.dateKey+' '+a.shiftCode+': '+
          clean_(u['Pharmacist Name'])+
          ' is a preceptor and must work a weekday Day/Morning shift so the student has preceptor coverage.'
        );
        return;
      }

      if (preferredCodes.length && preferredCodes.indexOf(code) < 0) {
        report.errors.push(
          a.dateKey+' '+a.shiftCode+': '+
          clean_(u['Pharmacist Name'])+
          ' is a preceptor with Preferred / Home Unit '+
          preferredCodes.join('/')+
          ' and must remain in that unit on weekdays.'
        );
      }
    });

  if (typeof unique_ === 'function') {
    report.errors = unique_(report.errors);
    report.warnings = unique_(report.warnings);
  }

  return report;
}

function preceptorCalendarAwareEligibility_(u,slot,model,state,manualMode) {
  var e = _PC11_BASE_ELIGIBILITY_
    ? _PC11_BASE_ELIGIBILITY_(u,slot,model,state,manualMode)
    : {ok:true,reasons:[],warnings:[]};

  if (!e) e = {ok:true,reasons:[],warnings:[]};
  e.reasons = Array.isArray(e.reasons) ? e.reasons.slice() : [];
  e.warnings = Array.isArray(e.warnings) ? e.warnings.slice() : [];

  // The residentEligibilityReason_ wrapper already enforces the weekday
  // preferred-unit/day-shift rule. Keep this additional guard so an older base
  // eligibility implementation cannot accidentally allow a weekday evening.
  if (
    e.ok &&
    u &&
    slot &&
    yes_(u.Preceptor) &&
    preceptorCalendarIsManagedWeekday_(u,slot.date) &&
    !preceptorCalendarIsDaySlot_(slot,model)
  ) {
    e.reasons.push('PRECEPTOR_WEEKDAY_DAY_SHIFT_ONLY');
    e.ok = false;
  }

  return e;
}

function preceptorCalendarAwareReasonToWarning_(reason) {
  if (reason === 'PRECEPTOR_WEEKDAY_DAY_SHIFT_ONLY') {
    return 'Preceptor weekday rule: pharmacist must work a Day/Morning shift so the student has preceptor coverage.';
  }

  if (reason === 'PRECEPTOR_PREFERRED_UNIT_ONLY') {
    return 'Preceptor weekday rule: pharmacist must work the Skills → Preferred / Home Unit.';
  }

  // Backward-compatible messages for older saved warnings.
  if (reason === 'PRECEPTOR_OFF_EVENING_OR_PREFERRED_ONLY') {
    return 'Preceptor weekday rule updated: weekday evening rotation is no longer used; schedule a Day/Morning shift instead.';
  }

  if (reason === 'PRECEPTOR_ON_PREFERRED_UNIT_ONLY') {
    return 'Preceptor weekday rule: pharmacist must work the Skills → Preferred / Home Unit.';
  }

  if (reason === 'PRECEPTOR_MONTHLY_EVENING_MAX') {
    return 'Preceptor weekday evening rotation is disabled under the current teaching-coverage rule.';
  }

  if (reason === 'PRECEPTOR_PRECEPTING_WEEK_NO_EVENING') {
    return 'Preceptor weekday rule: pharmacist cannot work an evening shift while providing student coverage.';
  }

  return _PC11_BASE_REASON_TO_WARNING_
    ? _PC11_BASE_REASON_TO_WARNING_(reason)
    : reason;
}

/* Capture the current Code.gs functions, then install Code11 wrappers. */
var _PC11_BASE_ELIGIBILITY_ =
  typeof eligibility_ === 'function'
    ? eligibility_
    : null;

var _PC11_BASE_RESIDENT_ELIGIBILITY_REASON_ =
  typeof residentEligibilityReason_ === 'function'
    ? residentEligibilityReason_
    : null;

var _PC11_BASE_SCORE_CANDIDATE_ =
  typeof scoreCandidate_ === 'function'
    ? scoreCandidate_
    : null;

var _PC11_BASE_PREASSIGN_RESIDENT_E2_ =
  typeof preassignResidentE2_ === 'function'
    ? preassignResidentE2_
    : null;

var _PC11_BASE_VALIDATE_GENERATED_ASSIGNMENTS_ =
  typeof validateGeneratedAssignments_ === 'function'
    ? validateGeneratedAssignments_
    : null;

var _PC11_BASE_REASON_TO_WARNING_ =
  typeof reasonToWarning_ === 'function'
    ? reasonToWarning_
    : null;

preceptorEveningBlocked_ = preceptorCalendarAwarePreceptorEveningBlocked_;

if (_PC11_BASE_ELIGIBILITY_) {
  eligibility_ = preceptorCalendarAwareEligibility_;
}

if (_PC11_BASE_RESIDENT_ELIGIBILITY_REASON_) {
  residentEligibilityReason_ = preceptorCalendarAwareEligibilityReason_;
}

if (_PC11_BASE_SCORE_CANDIDATE_) {
  scoreCandidate_ = preceptorCalendarAwareScoreCandidate_;
}

if (_PC11_BASE_PREASSIGN_RESIDENT_E2_) {
  preassignResidentE2_ = preceptorCalendarAwareResidentE2Pass_;
}

if (_PC11_BASE_VALIDATE_GENERATED_ASSIGNMENTS_) {
  validateGeneratedAssignments_ = preceptorCalendarAwareValidateGeneratedAssignments_;
}

if (_PC11_BASE_REASON_TO_WARNING_) {
  reasonToWarning_ = preceptorCalendarAwareReasonToWarning_;
}

function verifyPreceptorCalendarIntegration(token) {
  requireAdmin_(token);

  var connected =
    String(preceptorEveningBlocked_ || '').indexOf('preceptorCalendarAwarePreceptorEveningBlocked_') >= 0 &&
    String(residentEligibilityReason_ || '').indexOf('preceptorCalendarAwareEligibilityReason_') >= 0 &&
    String(scoreCandidate_ || '').indexOf('preceptorCalendarAwareScoreCandidate_') >= 0 &&
    String(preassignResidentE2_ || '').indexOf('preceptorCalendarAwareResidentE2Pass_') >= 0 &&
    String(validateGeneratedAssignments_ || '').indexOf('preceptorCalendarAwareValidateGeneratedAssignments_') >= 0;

  return {
    ok:true,
    connected:connected,
    sheetExists:!!getDb_().getSheetByName(PRECEPTOR_CALENDAR_SHEET_),
    behavior:'Preceptor=Yes requires weekday Day/Morning work. Preferred/Home Unit is mandatory when configured; otherwise any qualified Day/Morning shift may be used. Preceptor Calendar OFF no longer triggers weekday evening rotation.',
    message:connected
      ? 'Code11 is connected. Preceptors stay on weekday Day/Morning shifts; configured Preferred/Home Units are enforced, and OFF weeks no longer create weekday evening rotations.'
      : 'Code11 loaded, but one or more scheduling hooks are not connected. Replace the old Code11.gs with this version and redeploy.'
  };
}

/** Returns true only when a Preceptor=Yes pharmacist is actively precepting on this date. */
function preceptorCalendarIsActiveOnDate_(u,date) {
  if (!u || !yes_(u.Preceptor)) return false;
  var d = startOfDay_(asDate_(date));
  if (!d) return true;

  var cache = preceptorCalendarRuntimeCache_();
  if (!cache.sheetExists) return true;

  var key = preceptorCalendarEmployeeKeyFromUser_(u)+'|'+d.getFullYear()+'|'+(d.getMonth()+1);
  var row = cache.rowsByKey[key];
  if (!row || !yesDefault_(row.Active,true)) return true;

  var weekNo = preceptorCalendarWeekNumber_(d);
  var value = row['Week '+weekNo];
  if (clean_(value) === '') return true;
  return yes_(value);
}

/** Public diagnostic helper for the admin page / testing. */
function getPreceptorStatusForDate(token, employeeId, date) {
  requireAdmin_(token);

  var eid = clean_(employeeId);
  var u = readTable_(APP.SHEETS.USERS).find(function(x){
    return clean_(x['Employee ID']) === eid;
  });

  if (!u) throw new Error('Pharmacist not found.');

  var d = startOfDay_(asDate_(date));
  if (!d) throw new Error('A valid date is required.');

  var mode = preceptorCalendarModeForDate_(u,d);
  var model = null;
  try { model = loadSchedulingModel_(); } catch(e) {}

  return {
    ok:true,
    employeeId:eid,
    pharmacistName:clean_(u['Pharmacist Name']),
    date:preceptorCalendarDateKey_(d),
    preceptorCapable:yes_(u.Preceptor),
    activelyPrecepting:preceptorCalendarIsActiveOnDate_(u,d),
    mode:mode,
    preferredUnits:model ? preceptorCalendarPreferredUnitCodes_(u,model) : [],
    expectedAssignment:
      preceptorCalendarIsPreceptingWeek_(u,d)
        ? 'No evening shifts this Sunday-Saturday week; ON weekdays stay in Skills → Preferred / Home Unit'
        : mode === 'ON'
          ? 'Skills → Preferred / Home Unit'
          : mode === 'OFF'
            ? 'Skills → Preferred / Home Unit, E1, or E2'
            : 'Existing scheduler rules',
    weekNumber:preceptorCalendarWeekNumber_(d)
  };
}


/* ------------------------- INTERNAL HELPERS ------------------------- */

function preceptorCalendarEnsureSheet_() {
  var ss = getDb_();
  var sh = ss.getSheetByName(PRECEPTOR_CALENDAR_SHEET_);
  if (!sh) sh = ss.insertSheet(PRECEPTOR_CALENDAR_SHEET_);

  var needsHeader = sh.getLastRow() === 0;
  if (!needsHeader) {
    var existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),PRECEPTOR_CALENDAR_HEADERS_.length)).getValues()[0];
    needsHeader = PRECEPTOR_CALENDAR_HEADERS_.some(function(h,i){ return clean_(existing[i]) !== h; });
  }
  if (needsHeader) {
    sh.clear();
    sh.getRange(1,1,1,PRECEPTOR_CALENDAR_HEADERS_.length).setValues([PRECEPTOR_CALENDAR_HEADERS_]);
  }

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,PRECEPTOR_CALENDAR_HEADERS_.length)
    .setFontWeight('bold')
    .setBackground('#16324a')
    .setFontColor('#ffffff');
  sh.setColumnWidth(1,190);
  sh.setColumnWidth(2,110);
  sh.setColumnWidth(3,190);
  sh.setColumnWidth(4,110);
  sh.setColumnWidth(5,70);
  sh.setColumnWidth(6,110);
  sh.setColumnWidth(7,90);
  for (var c=8;c<=13;c++) sh.setColumnWidth(c,82);
  sh.setColumnWidth(14,70);
  sh.setColumnWidth(15,250);

  var maxRows = Math.max(2,sh.getMaxRows());
  var validation = SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'],true).setAllowInvalid(false).build();
  sh.getRange(2,8,maxRows-1,6).setDataValidation(validation);
  sh.getRange(2,14,maxRows-1,1).setDataValidation(validation);
  return sh;
}

function preceptorCalendarEnsureRowsForYear_(year, actor) {
  var y = preceptorCalendarNormalizeYear_(year);
  var sh = preceptorCalendarEnsureSheet_();
  var users = readTable_(APP.SHEETS.USERS).filter(function(u){
    return clean_(u.Role).toLowerCase() !== 'administrator' && yes_(u.Preceptor);
  });

  var values = sh.getDataRange().getValues();
  var headers = values[0] || PRECEPTOR_CALENDAR_HEADERS_;
  var h = {};
  headers.forEach(function(name,i){ h[clean_(name)] = i; });
  var existing = {};
  for (var i=1;i<values.length;i++) {
    var row = values[i];
    if (!row.some(function(v){ return clean_(v) !== ''; })) continue;
    var key = clean_(row[h['Employee ID']])+'|'+Number(row[h['Year']])+'|'+Number(row[h['Month Number']]);
    existing[key] = {rowIndex:i,row:row};
  }

  var added = [], repaired = 0;
  users.forEach(function(u){
    var eid = clean_(u['Employee ID']);
    for (var month=1;month<=12;month++) {
      var key = eid+'|'+y+'|'+month;
      var segments = preceptorCalendarMonthSegments_(y,month);
      if (!existing[key]) {
        var rowObj = {};
        PRECEPTOR_CALENDAR_HEADERS_.forEach(function(name){ rowObj[name]=''; });
        rowObj['Calendar ID'] = 'PC-'+eid+'-'+y+'-'+String(month).padStart(2,'0');
        rowObj['Employee ID'] = eid;
        rowObj['Pharmacist Name'] = clean_(u['Pharmacist Name']);
        rowObj.Username = clean_(u.Username);
        rowObj.Year = y;
        rowObj.Month = preceptorCalendarMonthName_(month);
        rowObj['Month Number'] = month;
        for (var w=1;w<=6;w++) rowObj['Week '+w] = w<=segments.length ? 'Yes' : '';
        rowObj.Active = 'Yes';
        rowObj.Notes = 'Preceptor=Yes means capable; this row controls active precepting weeks.';
        rowObj['Updated At'] = new Date();
        rowObj['Updated By'] = actor || 'SYSTEM';
        added.push(PRECEPTOR_CALENDAR_HEADERS_.map(function(name){ return rowObj[name]; }));
      } else {
        var ex = existing[key].row;
        var dirty = false;
        if (clean_(ex[h['Pharmacist Name']]) !== clean_(u['Pharmacist Name'])) { ex[h['Pharmacist Name']] = clean_(u['Pharmacist Name']); dirty=true; }
        if (clean_(ex[h['Username']]) !== clean_(u.Username)) { ex[h['Username']] = clean_(u.Username); dirty=true; }
        if (clean_(ex[h['Month']]) !== preceptorCalendarMonthName_(month)) { ex[h['Month']] = preceptorCalendarMonthName_(month); dirty=true; }
        for (var w2=1;w2<=6;w2++) {
          var ci = h['Week '+w2];
          if (ci===undefined) continue;
          if (w2>segments.length && clean_(ex[ci])!=='') { ex[ci]=''; dirty=true; }
          if (w2<=segments.length && clean_(ex[ci])==='') { ex[ci]='Yes'; dirty=true; }
        }
        if (dirty) {
          ex[h['Updated At']] = new Date();
          ex[h['Updated By']] = actor || 'SYSTEM';
          repaired++;
        }
      }
    }
  });

  if (values.length>1) sh.getRange(1,1,values.length,headers.length).setValues(values);
  if (added.length) sh.getRange(sh.getLastRow()+1,1,added.length,PRECEPTOR_CALENDAR_HEADERS_.length).setValues(added);
  SpreadsheetApp.flush();
  _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = null;
  return {rowsAdded:added.length,rowsRepaired:repaired};
}

function preceptorCalendarReadRows_() {
  var sh = getDb_().getSheetByName(PRECEPTOR_CALENDAR_SHEET_);
  if (!sh || sh.getLastRow()<2) return [];
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function(x){ return clean_(x); });
  return values.slice(1).filter(function(row){ return row.some(function(v){ return clean_(v)!==''; }); }).map(function(row){
    var obj = {};
    headers.forEach(function(h,i){ obj[h]=row[i]; });
    return obj;
  });
}

function preceptorCalendarRuntimeCache_() {
  if (_PRECEPTOR_CALENDAR_RUNTIME_CACHE_) return _PRECEPTOR_CALENDAR_RUNTIME_CACHE_;
  var sh = getDb_().getSheetByName(PRECEPTOR_CALENDAR_SHEET_);
  if (!sh || sh.getLastRow()<2) {
    _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = {sheetExists:!!sh,rowsByKey:{}};
    return _PRECEPTOR_CALENDAR_RUNTIME_CACHE_;
  }
  var rows = preceptorCalendarReadRows_();
  var byKey = {};
  rows.forEach(function(r){
    var y = Number(r.Year), m = Number(r['Month Number']);
    if (!y || !m) return;
    byKey[preceptorCalendarEmployeeKeyFromRow_(r)+'|'+y+'|'+m] = r;
  });
  _PRECEPTOR_CALENDAR_RUNTIME_CACHE_ = {sheetExists:true,rowsByKey:byKey};
  return _PRECEPTOR_CALENDAR_RUNTIME_CACHE_;
}

function preceptorCalendarEmployeeKeyFromUser_(u) {
  return clean_(u['Employee ID']) || clean_(u.Username) || clean_(u['Pharmacist Name']);
}
function preceptorCalendarEmployeeKeyFromRow_(r) {
  return clean_(r['Employee ID']) || clean_(r.Username) || clean_(r['Pharmacist Name']);
}

function preceptorCalendarNormalizeYear_(year) {
  var y = Number(year);
  if (!Number.isFinite(y) || y<2020 || y>2100) throw new Error('Year must be between 2020 and 2100.');
  return Math.floor(y);
}

function preceptorCalendarMonthName_(month) {
  return ['January','February','March','April','May','June','July','August','September','October','November','December'][Number(month)-1] || '';
}

function preceptorCalendarWeekNumber_(date) {
  var d = startOfDay_(asDate_(date));
  var first = new Date(d.getFullYear(),d.getMonth(),1);
  return Math.floor((d.getDate()+first.getDay()-1)/7)+1;
}

function preceptorCalendarMonthSegments_(year,month) {
  var y = Number(year), m = Number(month);
  var first = new Date(y,m-1,1);
  var last = new Date(y,m,0);
  var out = [], cursor = new Date(first), week = 1;
  while (cursor.getTime() <= last.getTime() && week <= 6) {
    var daysToSaturday = (6-cursor.getDay()+7)%7;
    var end = new Date(cursor);
    end.setDate(end.getDate()+daysToSaturday);
    if (end.getTime()>last.getTime()) end = new Date(last);
    out.push({weekNumber:week,start:new Date(cursor),end:new Date(end)});
    cursor = new Date(end);
    cursor.setDate(cursor.getDate()+1);
    week++;
  }
  return out;
}

function preceptorCalendarDateKey_(date) {
  var d = startOfDay_(asDate_(date));
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}

function preceptorCalendarRangeLabel_(start,end) {
  var opts = {month:'short',day:'numeric'};
  var a = start.toLocaleDateString('en-US',opts);
  var b = end.toLocaleDateString('en-US',opts);
  return a===b ? a : a+' – '+b;
}