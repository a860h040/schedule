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
 * - During an OFF calendar week, the scheduler treats that pharmacist as a
 *   normal pharmacist for preceptor-specific scheduling restrictions.
 * - New calendar rows default to Yes so installing this add-on does not
 *   unexpectedly change the behavior of existing preceptors.
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
 *                  -> weekday work may be E1, E2, OR the pharmacist's
 *                     Skills → Preferred / Home Unit.
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
 * Existing preceptor evening restriction:
 * ON  -> evening remains blocked unless the legacy settings permit it.
 * OFF -> E1/E2 must be allowed so the pharmacist can move to evening staffing.
 */
function preceptorCalendarAwarePreceptorEveningBlocked_(u,slot,model) {
  if (!u || !slot || !slot.shift || !yes_(u.Preceptor)) return false;

  if (!preceptorCalendarIsActiveOnDate_(u,slot.date)) {
    return false;
  }

  if (clean_(slot.shift.Type).toLowerCase() !== 'evening') return false;
  if (isResidentMandatoryE2_(u,slot,model)) return false;

  if (isWeekendDate_(slot.date)) {
    return !model.settings.preceptorWeekendEveningAllowed;
  }

  return !model.settings.preceptorWeekdayEveningAllowed;
}

/**
 * Hard eligibility:
 *
 * ON weekday:
 *   pharmacist may work only the exact preferred/home-unit shift code(s).
 *
 * OFF weekday:
 *   pharmacist may work only E1 or E2.
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

  var mode = preceptorCalendarModeForDate_(u,date);
  var code = clean_(shift && shift.Shift).toUpperCase();

  if (mode === 'OFF') {
    var offPreferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);
    var isEveningOption = preceptorCalendarEveningCode_(code);
    var isPreferredOption = offPreferredCodes.indexOf(code) >= 0;

    if (!isEveningOption && !isPreferredOption) {
      return 'PRECEPTOR_OFF_E1_E2_OR_PREFERRED_ONLY';
    }
    return '';
  }

  if (mode === 'ON') {
    var preferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);

    // Exact unit configured -> enforce it.
    if (preferredCodes.length && preferredCodes.indexOf(code) < 0) {
      return 'PRECEPTOR_ON_PREFERRED_UNIT_ONLY';
    }
  }

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

  var mode = preceptorCalendarModeForDate_(u,slot.date);

  if (
    mode === 'ON' &&
    preceptorCalendarPreferredUnitMatches_(u,slot,model)
  ) {
    score += 300000;
  }

  if (mode === 'OFF') {
    var offCode = clean_(slot && slot.shiftCode).toUpperCase();
    var offPreferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);

    if (preceptorCalendarEveningCode_(offCode)) {
      var monthDate = firstOfMonth_(slot.date);
      var target = preceptorCalendarMonthlyEveningTarget_(u,monthDate,model);
      var maxAllowed = preceptorCalendarMonthlyEveningMaximum_(u,monthDate,model);
      var have = preceptorCalendarCountE1E2InMonth_(clean_(u.Username),monthDate,state);

      if (have < target) {
        // Strongly prioritize the normal target of five.
        score += 300000;
      } else if (have < maxAllowed) {
        // Shifts #6 and #7 are fallback-only. If another eligible pharmacist
        // can cover this E1/E2 slot, strongly prefer that pharmacist instead.
        if (preceptorCalendarOtherEligibleEveningCovererExists_(u,slot,model,state)) {
          score -= 1000000;
        } else {
          // Nobody else can cover: allow this preceptor to be the fallback.
          score += 5000;
        }
      } else {
        // The eligibility wrapper below enforces the hard ceiling. Keep this
        // score penalty as an additional safety measure.
        score -= 1000000;
      }
    } else if (offPreferredCodes.indexOf(offCode) >= 0) {
      score += 300000;
    }
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
             !isSevenOn_(u) &&
             preceptorCalendarPreferredUnitCodes_(u,model).length > 0;
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

    var onToday = users.filter(function(u){
      return preceptorCalendarModeForDate_(u,day) === 'ON';
    });

    onToday.forEach(function(u){
      var username = clean_(u.Username);
      if (state.byEmployeeDate[username+'|'+dk]) return;

      var preferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);

      var candidates = slots
        .filter(function(s){
          return s &&
                 s.dateKey === dk &&
                 !reserved.has(s.slotKey) &&
                 preferredCodes.indexOf(clean_(s.shiftCode).toUpperCase()) >= 0;
        })
        .sort(function(a,b){
          return a.slot-b.slot ||
                 clean_(a.shiftCode).localeCompare(clean_(b.shiftCode));
        });

      for (var i=0;i<candidates.length;i++) {
        var slot = candidates[i];

        var assigned = assignSpecificUser_(
          slot,
          u,
          model,
          state,
          actor,
          'PRECEPTOR CALENDAR ON — PREFERRED UNIT'
        );

        if (assigned) {
          results.push(assigned);
          reserved.add(slot.slotKey);
          break;
        }
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

   2. If one or more weekdays in the calendar month are OFF:
        target = 5 E1/E2 shifts for that calendar month.

   3. A preceptor may work up to 7 E1/E2 shifts in that month, but shifts
      above the target of 5 are fallback coverage only. The scheduler should
      use shift #6 or #7 only when no other eligible pharmacist can cover
      that E1/E2 shift.

   4. E1/E2 rotation shifts can be placed ONLY on calendar-OFF weekdays.

   5. On each planned E1/E2 rotation date, another pharmacist with the NORMAL
      Employee Skill for the preceptor's Preferred/Home Unit is assigned to
      cover that home-unit shift when the unit is required that day.

   6. Existing hard rules still apply:
      PTO, weekly availability, one shift/day, five-day week, hours,
      evening-to-morning transition, resident rules, weekend rules, etc.
*/

var PRECEPTOR_MONTHLY_EVENING_TARGET_ = 5;
var PRECEPTOR_MONTHLY_EVENING_MAX_ = 7;
var PRECEPTOR_ROTATION_COVERAGE_REASON_ = 'PRECEPTOR EVENING ROTATION';

function preceptorCalendarMonthHasOffWeekday_(u, monthDate) {
  if (!u || !yes_(u.Preceptor)) return false;

  var first = firstOfMonth_(monthDate);
  var last = lastOfMonth_(monthDate);

  for (var d = new Date(first); d.getTime() <= last.getTime(); d = addDays_(d,1)) {
    if (isWeekendDate_(d)) continue;
    if (!preceptorCalendarIsActiveOnDate_(u,d)) return true;
  }
  return false;
}

function preceptorCalendarMonthlyEveningMaximum_(u, monthDate, model) {
  if (!u || !model || !yes_(u.Preceptor) || isSevenOn_(u)) return 0;
  if (!preceptorCalendarMonthHasOffWeekday_(u,monthDate)) return 0;

  // Seven is the absolute preceptor E1/E2 ceiling. A lower pharmacist-specific
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

function preceptorCalendarCountE1E2InMonth_(username, monthDate, state) {
  var mk = monthKey_(monthDate);
  return (state.assignments || []).filter(function(a) {
    return a &&
      a.status !== 'UNFILLED' &&
      clean_(a.username) === clean_(username) &&
      monthKey_(a.date) === mk &&
      preceptorCalendarEveningCode_(a.shiftCode);
  }).length;
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
      'COVERING PRECEPTOR HOME UNIT WHILE PRECEPTOR WORKS E1/E2'
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
 * Assign non-precepting pharmacists to E1/E2 before the general allocator.
 * Mandatory resident E2 is handled first by the original scheduler.
 */
function preceptorCalendarPreassignOffEvenings_(
  slots, results, model, state, actor, reserved, start, end
) {
  if (!slots || !slots.length || !model || !state || !reserved) return;

  var users = (model.users || [])
    .filter(function(u) {
      return yes_(u.Active) &&
             yes_(u.Preceptor) &&
             !isSevenOn_(u) &&
             preceptorCalendarPreferredUnitCodes_(u,model).length > 0;
    })
    .sort(function(a,b) {
      return clean_(a['Pharmacist Name']).localeCompare(clean_(b['Pharmacist Name']));
    });

  if (!users.length) return;

  var rangeStart = startOfDay_(asDate_(start));
  var rangeEnd = startOfDay_(asDate_(end));
  if (!rangeStart || !rangeEnd) return;

  users.forEach(function(u) {
    var username = clean_(u.Username);

    // Find every calendar month touched by this generation chunk.
    var months = {};
    for (
      var md = new Date(rangeStart);
      md.getTime() <= rangeEnd.getTime();
      md = addDays_(md,1)
    ) {
      months[monthKey_(md)] = firstOfMonth_(md);
    }

    Object.keys(months).sort().forEach(function(monthKey) {
      var monthDate = months[monthKey];
      var target = preceptorCalendarMonthlyEveningTarget_(u,monthDate,model);

      // Full month ON -> target 0. Nothing to rotate.
      if (target < 1) return;

      var have = preceptorCalendarCountE1E2InMonth_(username,monthDate,state);
      if (have >= target) return;

      var eligibleDays = [];
      for (
        var d = new Date(rangeStart);
        d.getTime() <= rangeEnd.getTime();
        d = addDays_(d,1)
      ) {
        if (monthKey_(d) !== monthKey) continue;
        if (isWeekendDate_(d)) continue;
        if (preceptorCalendarModeForDate_(u,d) !== 'OFF') continue;
        eligibleDays.push(new Date(d));
      }

      eligibleDays.sort(function(a,b){ return a.getTime()-b.getTime(); });

      for (var di=0; di<eligibleDays.length && have<target; di++) {
        var day = eligibleDays[di];
        var dk = formatDateKey_(day);

        if (state.byEmployeeDate[username+'|'+dk]) continue;

        var eveningSlots = (slots || [])
          .filter(function(s) {
            return s &&
              s.dateKey === dk &&
              !reserved.has(s.slotKey) &&
              preceptorCalendarEveningCode_(s.shiftCode);
          })
          .sort(function(a,b) {
            // Keep the distribution deterministic: E1 before E2.
            var ac = clean_(a.shiftCode).toUpperCase();
            var bc = clean_(b.shiftCode).toUpperCase();
            var ar = ac === 'E1' ? 0 : 1;
            var br = bc === 'E1' ? 0 : 1;
            return ar-br || a.slot-b.slot;
          });

        if (!eveningSlots.length) continue;

        var preferredSlots = preceptorCalendarPreferredUnitSlotsForDate_(
          u,dk,slots,model
        );

        var freePreferred = preferredSlots.filter(function(s) {
          return !reserved.has(s.slotKey);
        });

        var assignedEvening = null;

        for (var ei=0; ei<eveningSlots.length; ei++) {
          var eveningSlot = eveningSlots[ei];

          // Confirm the preceptor can take this E1/E2 under all hard rules.
          var e = eligibility_(u,eveningSlot,model,state,false);
          if (!e.ok) continue;

          var coverage = null;

          // If the preferred/home unit is required and currently free,
          // secure a qualified coverer before moving the preceptor to evening.
          if (freePreferred.length) {
            var coverageSlot = freePreferred[0];
            coverage = preceptorCalendarAssignHomeCoverage_(
              u,coverageSlot,results,model,state,actor,reserved
            );

            if (!coverage) {
              // Do not pull the preceptor away from the home unit when nobody
              // with the normal skill can cover that required unit.
              continue;
            }
          }

          assignedEvening = assignSpecificUser_(
            eveningSlot,
            u,
            model,
            state,
            actor,
            'PRECEPTOR CALENDAR OFF — MONTHLY E1/E2 ROTATION'
          );

          if (!assignedEvening) {
            preceptorCalendarRemoveNewCoverage_(
              coverage,results,model,state,reserved
            );
            continue;
          }

          results.push(assignedEvening);
          reserved.add(eveningSlot.slotKey);
          have++;
          break;
        }
      }
    });
  });
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
  // five-evening rotation target on OFF weekdays and simultaneously cover
  // the pharmacist's preferred/home unit with another normally-skilled pharmacist.
  preceptorCalendarPreassignOffEvenings_(
    slots, results, model, state, actor, reserved, start, end
  );
}

/**
 * Full-period validation:
 * ON  -> preferred unit
 * OFF -> E1/E2
 */
function preceptorCalendarAwareValidateGeneratedAssignments_(
  assignments, model, start, end, validationOptions
) {
  var report = _PC11_BASE_VALIDATE_GENERATED_ASSIGNMENTS_
    ? _PC11_BASE_VALIDATE_GENERATED_ASSIGNMENTS_(
        assignments, model, start, end, validationOptions
      )
    : {errors:[],warnings:[]};

  report = report || {errors:[],warnings:[]};  report.errors = Array.isArray(report.errors) ? report.errors.slice() : [];
  report.warnings = Array.isArray(report.warnings) ? report.warnings.slice() : [];

  var warnedMissingPreferred = {};

  (assignments || [])
    .filter(function(a){
      return a && a.status !== 'UNFILLED' && a.username;
    })
    .forEach(function(a){
      var u = model && model.usersByUsername
        ? model.usersByUsername[clean_(a.username)]
        : null;

      if (!u || !preceptorCalendarIsManagedWeekday_(u,a.date)) return;

      var mode = preceptorCalendarModeForDate_(u,a.date);

      if (mode === 'OFF') {
        var offPreferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);
        var offCode = clean_(a.shiftCode).toUpperCase();

        if (
          !preceptorCalendarEveningCode_(offCode) &&
          offPreferredCodes.indexOf(offCode) < 0
        ) {
          report.errors.push(
            a.dateKey+' '+a.shiftCode+': '+
            clean_(u['Pharmacist Name'])+
            ' has Preceptor Calendar OFF and may work only E1, E2, or preferred unit '+
            (offPreferredCodes.length ? offPreferredCodes.join('/') : '(not configured)')+
            ' on this weekday.'
          );
        }
        return;
      }

      if (mode === 'ON') {
        var preferredCodes = preceptorCalendarPreferredUnitCodes_(u,model);

        if (!preferredCodes.length) {
          var userKey = clean_(u.Username);
          if (!warnedMissingPreferred[userKey]) {
            report.warnings.push(
              clean_(u['Pharmacist Name'])+
              ' is Preceptor Calendar ON but Skills tab does not have an active Preferred / Home Unit skill. Open Skills and set the pharmacist Preferred / Home Unit, for example IM, CC1, CC2, ONC, CARD, etc.'
            );
            warnedMissingPreferred[userKey] = true;
          }
          return;
        }

        if (preferredCodes.indexOf(clean_(a.shiftCode).toUpperCase()) < 0) {
          report.errors.push(
            a.dateKey+' '+a.shiftCode+': '+
            clean_(u['Pharmacist Name'])+
            ' has Preceptor Calendar ON and must work preferred unit '+
            preferredCodes.join('/')+'.'
          );
        }
      }
    });

  // Monthly E1/E2 target validation.
  // Build a combined month view from the assignments being validated plus
  // saved schedule rows outside the current validation range.
  var combined = [];
  var seenSlots = {};

  function addCombined_(a) {
    if (!a || a.status === 'UNFILLED') return;
    var key = a.dateKey+'|'+clean_(a.shiftCode)+'|'+num_(a.slot,1);
    if (seenSlots[key]) return;
    seenSlots[key] = true;
    combined.push(a);
  }

  (assignments || []).forEach(addCombined_);

  var rangeStart = startOfDay_(asDate_(start));
  var rangeEnd = startOfDay_(asDate_(end));

  if (rangeStart && rangeEnd) {
    readTable_(APP.SHEETS.SCHEDULE).forEach(function(r) {
      var d = startOfDay_(asDate_(r.Date));
      if (!d || inDateRange_(d,rangeStart,rangeEnd)) return;
      var a = scheduleRowToAssignment_(r,model);
      if (a) addCombined_(a);
    });

    (model.users || [])
      .filter(function(u) {
        return yes_(u.Active) && yes_(u.Preceptor) && !isSevenOn_(u);
      })
      .forEach(function(u) {
        var monthCursor = firstOfMonth_(rangeStart);
        var lastMonth = firstOfMonth_(rangeEnd);

        while (monthCursor.getTime() <= lastMonth.getTime()) {
          var target = preceptorCalendarMonthlyEveningTarget_(u,monthCursor,model);

          if (target > 0) {
            var mk = monthKey_(monthCursor);
            var count = combined.filter(function(a) {
              return clean_(a.username) === clean_(u.Username) &&
                monthKey_(a.date) === mk &&
                preceptorCalendarEveningCode_(a.shiftCode);
            }).length;

            var fullMonthInValidation =
              rangeStart.getTime() <= firstOfMonth_(monthCursor).getTime() &&
              rangeEnd.getTime() >= lastOfMonth_(monthCursor).getTime();

            var msg =
              clean_(u['Pharmacist Name'])+
              ' has Preceptor Calendar OFF during '+mk+
              ' and needs '+target+' E1/E2 rotation shift(s); currently '+count+'.';

            var maxAllowed = preceptorCalendarMonthlyEveningMaximum_(u,monthCursor,model);

            if (count < target) {
              if (fullMonthInValidation) report.errors.push(msg);
              else report.warnings.push(msg+' The validation range does not contain the full month.');
            } else if (count > maxAllowed) {
              report.errors.push(
                clean_(u['Pharmacist Name'])+' has '+count+
                ' E1/E2 shifts during '+mk+
                ', above the allowed maximum of '+maxAllowed+'.'
              );
            } else if (count > target) {
              report.warnings.push(
                clean_(u['Pharmacist Name'])+' has '+count+
                ' E1/E2 shifts during '+mk+
                '. Target is '+target+'; shifts above target are allowed only as fallback coverage when no other eligible pharmacist can cover E1/E2.'
              );
            }
          }

          monthCursor = firstOfMonth_(addDays_(lastOfMonth_(monthCursor),1));
        }
      });
  }

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

  if (
    e.ok &&
    u &&
    slot &&
    preceptorCalendarIsManagedWeekday_(u,slot.date) &&
    preceptorCalendarModeForDate_(u,slot.date) === 'OFF' &&
    preceptorCalendarEveningCode_(slot.shiftCode)
  ) {
    var monthDate = firstOfMonth_(slot.date);
    var maxAllowed = preceptorCalendarMonthlyEveningMaximum_(u,monthDate,model);
    var have = preceptorCalendarCountE1E2InMonth_(clean_(u.Username),monthDate,state);

    if (maxAllowed <= 0 || have + 1 > maxAllowed) {
      e.reasons.push('PRECEPTOR_MONTHLY_EVENING_MAX');
      e.ok = false;
    }
  }

  return e;
}

function preceptorCalendarAwareReasonToWarning_(reason) {
  if (reason === 'PRECEPTOR_OFF_E1_E2_OR_PREFERRED_ONLY') {
    return 'Preceptor Calendar OFF: pharmacist may work only E1, E2, or the Skills → Preferred / Home Unit on this weekday.';
  }

  if (reason === 'PRECEPTOR_ON_PREFERRED_UNIT_ONLY') {
    return 'Preceptor Calendar ON: pharmacist must remain in the Skills → Preferred / Home Unit.';
  }

  if (reason === 'PRECEPTOR_MONTHLY_EVENING_MAX') {
    return 'Preceptor monthly E1/E2 maximum reached: target is 5 and the hard maximum is 7.';
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
    behavior:'ON = Skills Preferred/Home Unit. Any month with OFF weekdays targets 5 E1/E2 shifts. Maximum is 7, and shifts 6-7 are fallback-only when no other eligible pharmacist can cover E1/E2. Normally-skilled pharmacists cover the home unit on planned rotation dates.',
    message:connected
      ? 'Code11 is connected. ON weeks stay in the preferred unit. Months with OFF weekdays target 5 E1/E2 shifts, with a fallback maximum of 7 only when no other eligible pharmacist can cover the E1/E2 shift, with normal-skill home-unit coverage on those dates.'
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
      mode === 'ON'
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