export const APP_VERSION = '2026.09.18-github-full-1';

export const DEFAULT_SETTINGS = {
  appName: 'NeoChrono',
  scheduleDays: 56,
  weekStart: 'Sunday',
  requiredHoursPerSchedulePeriod: 320,
  maximumHoursPerSchedulePeriod: 320,
  weeklyHoursLimit: 40,
  regularWorkdaysPerWeek: 5,
  maximumConsecutiveWorkdays: 5,
  maximumEveningShifts: 7,
  weekendRotation: ['A','B','C'],
  weekendAnchorDate: '',
  weekendAnchorGroup: 'A',
  allowWeekendFallback: false,
  requiredWeekendAssignment: true,
  residentE2ShiftCode: 'E2',
  residentE2ShiftsPerWeek: 1,
  residentsRequiredAssignedWeekend: true,
  sevenOnBlockScheduling: true,
  blockEveningToMorningTransition: true,
  ptoAutoApprovalLimitPerDate: 2,
  allowAdminRuleOverride: true,
  showFullScheduleToPharmacists: true,
  preceptorEveningTargetPerMonthWhenOff: 7,
  preceptorHomeUnitRequiredWhenOn: true,
  preceptorOffAllowsEvenings: true,
  eddEdePairRequired: true,
  edDayShiftCode: 'EDD',
  edEveningShiftCode: 'EDE',
  minimumRestHours: 8,
  scheduleTimezone: 'America/New_York'
};

export function blankDatabase() {
  return {
    meta: {
      schemaVersion: 1,
      appVersion: APP_VERSION,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    },
    settings: structuredClone(DEFAULT_SETTINGS),
    admins: [],
    users: [],
    skills: [],
    shifts: [],
    staffingRequirements: [],
    requests: [],
    weeklyAvailability: [],
    preceptorCalendar: [],
    schedule: [],
    schedulePeriods: [],
    swaps: [],
    audit: [],
    sessions: [],
    publishedSchedule: []
  };
}

export const USER_DEFAULTS = {
  active: true,
  role: 'Pharmacist',
  scheduleType: 'Regular',
  preceptor: false,
  resident: false,
  weekendGroup: '',
  weekendRotationAnchorDate: '',
  weeklyHourMaximum: 40,
  targetWeeklyHours: 40,
  maximumEveningShiftsPerMonth: 7,
  preferredStartTime: '',
  preferredEndTime: '',
  customHoursEnabled: false,
  customHoursMode: 'SOFT',
  preferredShiftType: '',
  offDayCoverageSkills: [],
  weekendEligible: true,
  eveningEligible: true,
  nightEligible: true,
  residentCoversRegular: false,
  residentWeekends: true,
  residentEvenings: true,
  residentNights: false,
  rotationAnchorDate: '',
  sevenOnWeeklyHandling: 'ENFORCE_MAX',
  notes: '',
  mustChangePassword: false
};

export const SHIFT_DEFAULTS = {
  meaning: '',
  start: '',
  end: '',
  hours: 0,
  creditedHours: 0,
  type: 'Day',
  skill: '',
  weekend: false,
  active: true,
  priority: 50
};

export function normalizeDatabase(db) {
  const base = blankDatabase();
  const out = Object.assign(base, db || {});
  out.meta = Object.assign(base.meta, (db || {}).meta || {});
  out.settings = Object.assign({}, DEFAULT_SETTINGS, (db || {}).settings || {});
  ['admins','users','skills','shifts','staffingRequirements','requests','weeklyAvailability','preceptorCalendar','schedule','schedulePeriods','swaps','audit','sessions','publishedSchedule']
    .forEach(k => { if (!Array.isArray(out[k])) out[k] = []; });
  out.users = out.users.map(u => Object.assign({}, USER_DEFAULTS, u));
  out.shifts = out.shifts.map(s => Object.assign({}, SHIFT_DEFAULTS, s));
  return out;
}

export function uid(prefix='ID') {
  const raw = (crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`)
    .replace(/[^a-zA-Z0-9]/g,'').toUpperCase();
  return `${prefix}-${raw.slice(0,12)}`;
}

export function clean(v) { return String(v ?? '').trim(); }
export function yes(v) { return v === true || /^(yes|true|1|y)$/i.test(clean(v)); }
export function num(v, fallback=0) { const n = Number(v); return Number.isFinite(n) ? n : fallback; }
export function dateKey(d) {
  const x = d instanceof Date ? new Date(d) : new Date(`${d}T12:00:00`);
  if (Number.isNaN(x.getTime())) return '';
  const y=x.getFullYear(), m=String(x.getMonth()+1).padStart(2,'0'), day=String(x.getDate()).padStart(2,'0');
  return `${y}-${m}-${day}`;
}
export function addDays(value, days) { const d = value instanceof Date ? new Date(value) : new Date(`${value}T12:00:00`); d.setDate(d.getDate()+days); return d; }
export function dayName(value) { return ['SUN','MON','TUE','WED','THU','FRI','SAT'][(value instanceof Date ? value : new Date(`${value}T12:00:00`)).getDay()]; }
export function isWeekend(value) { const d=(value instanceof Date?value:new Date(`${value}T12:00:00`)).getDay(); return d===0||d===6; }
export function startOfWeek(value, weekStart='Sunday') {
  const d=value instanceof Date?new Date(value):new Date(`${value}T12:00:00`); d.setHours(12,0,0,0);
  const target=/monday/i.test(weekStart)?1:0; const diff=(d.getDay()-target+7)%7; d.setDate(d.getDate()-diff); return d;
}
export function monthKey(value) { const d=value instanceof Date?value:new Date(`${value}T12:00:00`); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; }
export function minutes(t) { if(!/^\d{1,2}:\d{2}$/.test(clean(t))) return null; const [h,m]=t.split(':').map(Number); return h*60+m; }
export function escapeHtml(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));}
