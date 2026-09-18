import { addDays, clean, dateKey, dayName, daysBetween, eachDate, isWeekend, monthKey, num, parseDate, uuid, weekStartSunday, yes } from './utils.js';
import { buildModel, creditedHours, dateAvailabilityStatus, isDayShift, isEveningShift, isNightShift, isOnPto, isPrnUser, isRegularUser, isSevenOnDate, isSevenUser, preceptorModeForDate, recurringAvailabilityStatus, scheduleHealth, weekendGroupForDate } from './rules.js';

const DAY_FIELDS=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const dk=d=>dateKey(d);

function createSlots(model,start,end,generationId){
  const rows=[]; const reqMap={}; (model.data.staffing||[]).filter(r=>yes(r.Active??'Yes')).forEach(r=>reqMap[clean(r.Shift).toUpperCase()]=r);
  for(const date of eachDate(start,end)){
    const dow=date.getDay();
    for(const [code,req] of Object.entries(reqMap)){
      const shift=model.shifts[code]; if(!shift||!yes(shift.Active))continue;
      const count=Math.max(0,Math.floor(num(req[DAY_FIELDS[dow]],0)));
      for(let slot=1;slot<=count;slot++) rows.push({generationId,date,shift,slot,slotKey:`${dk(date)}|${code}|${slot}`});
    }
  }
  return rows.sort((a,b)=>a.date-b.date||num(a.shift.Priority,50)-num(b.shift.Priority,50)||clean(a.shift.Shift).localeCompare(clean(b.shift.Shift))||a.slot-b.slot);
}

function makeState(model,fixed,start,end){
  const s={byUserDate:{},assignments:[],weeklyDays:{},weeklyHours:{},totalHours:{},monthEvenings:{},shiftCounts:{},workedDates:{},periodStart:start,periodEnd:end};
  (fixed||[]).forEach(r=>{const u=model.byUsername[clean(r.Username)];const sh=model.shifts[clean(r.Shift).toUpperCase()];const d=parseDate(r.Date);if(u&&sh&&d)register(s,u,{date:d,shift:sh,slot:num(r.Slot,1)},r);}); return s;
}
function wkKey(user,date){return `${clean(user.Username)}|${dk(weekStartSunday(date))}`;}
function register(state,user,slot,row){
  const uk=clean(user.Username),date=slot.date,dkv=dk(date),wk=wkKey(user,date),mh=`${uk}|${monthKey(date)}`;
  state.byUserDate[`${uk}|${dkv}`]=row||slot; (state.workedDates[uk]||(state.workedDates[uk]=new Set())).add(dkv);
  state.weeklyDays[wk]=(state.weeklyDays[wk]||0)+1; const ch=creditedHours(slot.shift);state.weeklyHours[wk]=(state.weeklyHours[wk]||0)+ch;state.totalHours[uk]=(state.totalHours[uk]||0)+ch;
  if(isEveningShift(slot.shift))state.monthEvenings[mh]=(state.monthEvenings[mh]||0)+1; const sk=`${uk}|${clean(slot.shift.Shift)}`;state.shiftCounts[sk]=(state.shiftCounts[sk]||0)+1;
  state.assignments.push(row||slot);
}
function unregister(state,user,slot){/* generation never backtracks in GitHub v1 */}
function hasDate(state,user,date){return !!state.byUserDate[`${clean(user.Username)}|${dk(date)}`];}
function previousAssignment(state,user,date){return state.byUserDate[`${clean(user.Username)}|${dk(addDays(date,-1))}`]||null;}
function nextAssignment(state,user,date){return state.byUserDate[`${clean(user.Username)}|${dk(addDays(date,1))}`]||null;}
function rowShift(model,row){return row?.shift||model.shifts[clean(row?.Shift).toUpperCase()]||null;}
function wouldExceedConsecutive(state,user,date,max){
  const set=new Set(state.workedDates[clean(user.Username)]||[]);set.add(dk(date));let run=1;for(let d=addDays(date,-1);set.has(dk(d));d=addDays(d,-1))run++;for(let d=addDays(date,1);set.has(dk(d));d=addDays(d,1))run++;return run>max;
}
function hasSkill(model,user,skill){return !clean(skill)||!!model.skills[clean(user.Username)]?.has(clean(skill).toUpperCase());}
function prnAvailable(model,user,date){const sat=addDays(date,-((parseDate(date).getDay()+1)%7));const satKey=dk(sat);const rows=(model.data.requests||[]).filter(r=>clean(r['Record Type']).toUpperCase()==='PRN_AVAILABILITY'&&(clean(r.Username)===clean(user.Username)||clean(r.Pharmacist)===clean(user['Pharmacist Name'])));if(!rows.length)return false;return rows.some(r=>dk(r['Weekend Saturday']||r.Date)===satKey&&yes(r.Available));}
function preferredCodes(user){return new Set(clean(user['Preferred Shift Type']).toUpperCase().split(/[,;|/]+/).map(x=>x.trim()).filter(Boolean));}
function preceptorRule(model,user,slot){
  if(!yes(user.Preceptor)||isWeekend(slot.date))return {ok:true}; const mode=preceptorModeForDate(user,slot.date,model.data.preceptorCalendar||[])||'ON'; const prefs=preferredCodes(user),code=clean(slot.shift.Shift).toUpperCase();
  if(mode==='ON'){if(prefs.size&& !prefs.has(code))return {ok:false,reason:'PRECEPTOR_HOME_UNIT'};if(isEveningShift(slot.shift))return {ok:false,reason:'PRECEPTOR_WEEKDAY_EVENING'};}
  else if(mode==='OFF'){if(prefs.size&&!prefs.has(code)&&!['E1','E2'].includes(code))return {ok:false,reason:'PRECEPTOR_OFF_ALLOWED_ONLY'};}
  return {ok:true,mode};
}
function userEligibleForType(user,shift){if(isWeekend(new Date())&&false)return true; if(isEveningShift(shift)&&!yes(user['Evening Eligible']??'Yes'))return false;if(isNightShift(shift)&&!yes(user['Night Eligible']??'Yes'))return false;return true;}

function eligibility(model,state,user,slot,{ignoreWeekendGroup=false,allowOffDayCoverageSkill=false}={}){
  const reasons=[]; if(!yes(user.Active))reasons.push('INACTIVE'); const normalSkill=hasSkill(model,user,slot.shift.Skill); const coverSet=new Set(clean(user['Off-Day Coverage Skills']).toUpperCase().split(/[,;|/]+/).map(x=>x.trim()).filter(Boolean)); const coverSkill=coverSet.has(clean(slot.shift.Skill).toUpperCase())||coverSet.has(clean(slot.shift.Shift).toUpperCase()); if(!normalSkill&&!(allowOffDayCoverageSkill&&coverSkill))reasons.push('MISSING_SKILL');
  if(isOnPto(user,slot.date,model.pto))reasons.push('PTO'); if(!dateAvailabilityStatus(user,slot.date,model.data.requests).ok)reasons.push('UNAVAILABLE');
  const rav=recurringAvailabilityStatus(user,slot.date,slot.shift,model.data.weeklyAvailability); if(!rav.ok)reasons.push(rav.reason);
  if(hasDate(state,user,slot.date))reasons.push('ALREADY_SCHEDULED'); if(!userEligibleForType(user,slot.shift))reasons.push('SHIFT_TYPE_INELIGIBLE');
  if(isPrnUser(user)){if(!isWeekend(slot.date))reasons.push('PRN_WEEKDAY');else if(!prnAvailable(model,user,slot.date))reasons.push('PRN_NOT_AVAILABLE');}
  if(isWeekend(slot.date)&&!yes(user['Weekend Eligible']??'Yes'))reasons.push('WEEKEND_NOT_ELIGIBLE');
  if(isSevenUser(user)){
    if(!isSevenOnDate(user,slot.date))reasons.push('SEVEN_OFF'); const prefs=preferredCodes(user); if(prefs.size&&!prefs.has(clean(slot.shift.Shift).toUpperCase()))reasons.push('SEVEN_WRONG_SHIFT');
  }
  if(isRegularUser(user)){
    const wk=wkKey(user,slot.date),maxDays=num(model.settings['Regular Workdays Per Week'],5),maxCon=num(model.settings['Maximum Consecutive Workdays'],5); if((state.weeklyDays[wk]||0)>=maxDays)reasons.push('WEEKLY_DAYS');if(wouldExceedConsecutive(state,user,slot.date,maxCon))reasons.push('CONSECUTIVE_DAYS');
    const weeklyMax=num(user['Weekly Hour Maximum'],num(model.settings['Weekly Hours Limit'],40));if((state.weeklyHours[wk]||0)+creditedHours(slot.shift)>weeklyMax+0.001)reasons.push('WEEKLY_HOURS');
    const target=num(model.settings['Required Hours Per Schedule Period'],320);if(daysBetween(state.periodStart,state.periodEnd)+1===56&&(state.totalHours[clean(user.Username)]||0)+creditedHours(slot.shift)>target+0.001)reasons.push('TWO_MONTH_HOURS');
  }
  if(isWeekend(slot.date)&&!isPrnUser(user)&&!isSevenUser(user)&&yes(model.settings['Required Weekend Assignment'])&&!ignoreWeekendGroup){const expected=weekendGroupForDate(slot.date,model.settings);if(expected&&clean(user['Weekend Group']).toUpperCase()!==expected)reasons.push(yes(user.Resident)?'RESIDENT_WRONG_WEEKEND_GROUP':'WRONG_WEEKEND_GROUP');}
  if(yes(user.Resident)){if(isWeekend(slot.date)&&!yes(user['Resident Weekends']??'Yes'))reasons.push('RESIDENT_RESTRICTION');if(isEveningShift(slot.shift)&&!yes(user['Resident Evenings']??'Yes'))reasons.push('RESIDENT_RESTRICTION');if(isNightShift(slot.shift)&&!yes(user['Resident Nights']??'No'))reasons.push('RESIDENT_RESTRICTION');}
  const pr=preceptorRule(model,user,slot);if(!pr.ok)reasons.push(pr.reason);
  if(yes(model.settings['Block Evening To Morning Transition'])){const prev=previousAssignment(state,user,slot.date),next=nextAssignment(state,user,slot.date);if(prev&&isEveningShift(rowShift(model,prev))&&isDayShift(slot.shift))reasons.push('EVENING_TO_MORNING');if(next&&isEveningShift(slot.shift)&&isDayShift(rowShift(model,next)))reasons.push('EVENING_TO_MORNING');}
  if(isEveningShift(slot.shift)){const cap=num(user['Maximum Evening Shifts Per Month'],num(model.settings['Maximum Evening Shifts'],7));if(cap>0&&(state.monthEvenings[`${clean(user.Username)}|${monthKey(slot.date)}`]||0)>=cap)reasons.push('EVENING_LIMIT');}
  return {ok:reasons.length===0,reasons};
}

function score(model,state,user,slot){
  const uk=clean(user.Username),target=num(model.settings['Required Hours Per Schedule Period'],320),current=state.totalHours[uk]||0;let s=(target-current)*10;
  const prefs=preferredCodes(user);if(prefs.has(clean(slot.shift.Shift).toUpperCase()))s+=500;
  const count=state.shiftCounts[`${uk}|${clean(slot.shift.Shift)}`]||0;s-=count*15;
  if(isWeekend(slot.date))s-=Object.keys(state.byUserDate).filter(k=>k.startsWith(`${uk}|`)).length*0.2;
  if(yes(user.Preceptor)&&preceptorModeForDate(user,slot.date,model.data.preceptorCalendar||[])==='OFF'&&['E1','E2'].includes(clean(slot.shift.Shift).toUpperCase())){const tgt=num(model.settings['Preceptor Off Evening Target Per Month'],7),cur=state.monthEvenings[`${uk}|${monthKey(slot.date)}`]||0;if(cur<tgt)s+=250;}
  if(yes(user.Resident)&&clean(slot.shift.Shift).toUpperCase()===clean(model.settings['Resident E2 Shift Code']||'E2').toUpperCase())s+=100;
  return s;
}
function pick(model,state,slot,filter=()=>true,opts={}){const candidates=model.data.users.filter(u=>filter(u)&&eligibility(model,state,u,slot,opts).ok).map(u=>({u,score:score(model,state,u,slot)})).sort((a,b)=>b.score-a.score||clean(a.u['Pharmacist Name']).localeCompare(clean(b.u['Pharmacist Name'])));return candidates[0]?.u||null;}

function primaryHomeOwners(model,slot){const code=clean(slot.shift.Shift).toUpperCase();return model.data.users.filter(u=>yes(u.Active)&&isRegularUser(u)&&preferredCodes(u).has(code));}
function ownerHasGeneratedOffDay(model,state,owner,slot){if(hasDate(state,owner,slot.date))return false;if(isOnPto(owner,slot.date,model.pto))return false;if(!dateAvailabilityStatus(owner,slot.date,model.data.requests).ok)return false;if(!recurringAvailabilityStatus(owner,slot.date,slot.shift,model.data.weeklyAvailability).ok)return false;const wk=wkKey(owner,slot.date);return (state.weeklyDays[wk]||0)>=num(model.settings['Regular Workdays Per Week'],5);}
function pickOffDayCoverage(model,state,slot){const owners=primaryHomeOwners(model,slot).filter(o=>ownerHasGeneratedOffDay(model,state,o,slot));if(owners.length!==1)return null;const owner=owners[0];const candidates=model.data.users.filter(u=>clean(u.Username)!==clean(owner.Username)&&eligibility(model,state,u,slot,{allowOffDayCoverageSkill:true}).ok&&!hasSkill(model,u,slot.shift.Skill)).map(u=>({u,score:score(model,state,u,slot)})).sort((a,b)=>b.score-a.score||clean(a.u['Pharmacist Name']).localeCompare(clean(b.u['Pharmacist Name'])));return candidates.length?{user:candidates[0].u,owner}:null;}
function assignmentFrom(slot,user,status='ASSIGNED',warning=''){
  const sh=slot.shift;return {'Generation ID':slot.generationId,'Assignment ID':uuid(),'Date':dk(slot.date),'Day':dayName(slot.date),'Shift':sh.Shift,'Slot':slot.slot,'Assigned Pharmacist':user?user['Pharmacist Name']:'UNFILLED',Username:user?user.Username:'',Hours:num(sh.Hours,0),'Credited Hours':creditedHours(sh),'Required Skill':clean(sh.Skill),'Coverage For Pharmacist':'','Coverage For Username':'','Coverage Reason':'','Shift Type':sh.Type,Weekend:isWeekend(slot.date)?'Yes':'No','Weekend Group':isWeekend(slot.date)?'':'','Holiday':'','Locked':'No','Manual':'No',Status:user?status:'UNFILLED',Warning:warning,'Updated At':new Date().toISOString(),'Updated By':'GITHUB SCHEDULER','Finalized At':''};
}
function assign(model,state,slot,user,out,status='ASSIGNED',warning='',coverage=null){const row=assignmentFrom(slot,user,status,warning);if(coverage){row['Coverage For Pharmacist']=coverage.owner['Pharmacist Name'];row['Coverage For Username']=coverage.owner.Username;row['Coverage Reason']='GENERATED OFF DAY';row.Warning=[row.Warning,`OFF-DAY COVERAGE ONLY: covering ${coverage.owner['Pharmacist Name']}`].filter(Boolean).join(' | ');}if(user){row['Weekend Group']=isWeekend(slot.date)?weekendGroupForDate(slot.date,model.settings):'';register(state,user,slot,row);}out.set(slot.slotKey,row);return row;}

function assignSeven(model,state,slots,out){for(const u of model.data.users.filter(x=>yes(x.Active)&&isSevenUser(x))){const prefs=preferredCodes(u);if(!prefs.size)continue;for(const slot of slots){if(out.has(slot.slotKey)||!isSevenOnDate(u,slot.date)||!prefs.has(clean(slot.shift.Shift).toUpperCase()))continue;if(eligibility(model,state,u,slot).ok)assign(model,state,slot,u,out,'ASSIGNED','7-on/7-off ON block');}}}
function assignResidentE2(model,state,slots,out){const code=clean(model.settings['Resident E2 Shift Code']||'E2').toUpperCase(),per=Math.max(0,Math.floor(num(model.settings['Resident E2 Shifts Per Week'],1)));if(!per)return;const weeks=[...new Set(slots.map(s=>dk(weekStartSunday(s.date))))];for(const u of model.data.users.filter(x=>yes(x.Active)&&yes(x.Resident)&&isRegularUser(x))){for(const wk of weeks){let assigned=0;for(const slot of slots.filter(s=>!out.has(s.slotKey)&&dk(weekStartSunday(s.date))===wk&&clean(s.shift.Shift).toUpperCase()===code)){if(assigned>=per)break;if(eligibility(model,state,u,slot).ok){assign(model,state,slot,u,out,'ASSIGNED','Resident weekly E2');assigned++;}}}}}
function assignEdPairs(model,state,slots,out){if(!yes(model.settings['Require Paired ED Coverage']))return;const dayCode=clean(model.settings['ED Day Shift Code']||'EDD').toUpperCase(),eveCode=clean(model.settings['ED Evening Shift Code']||'EDE').toUpperCase();const dates=[...new Set(slots.map(s=>dk(s.date)))];for(const date of dates){const a=slots.find(s=>!out.has(s.slotKey)&&dk(s.date)===date&&clean(s.shift.Shift).toUpperCase()===dayCode),b=slots.find(s=>!out.has(s.slotKey)&&dk(s.date)===date&&clean(s.shift.Shift).toUpperCase()===eveCode);if(!a||!b)continue;const u1=pick(model,state,a);if(u1)assign(model,state,a,u1,out,'ASSIGNED','Paired ED coverage');const u2=pick(model,state,b,u=>!u1||clean(u.Username)!==clean(u1.Username));if(u2)assign(model,state,b,u2,out,'ASSIGNED','Paired ED coverage');}}
function assignWeekend(model,state,slots,out){for(const slot of slots.filter(s=>isWeekend(s.date)&&!out.has(s.slotKey))){let u=pick(model,state,slot);if(!u&&yes(model.settings['Allow Weekend Fallback']))u=pick(model,state,slot,()=>true,{ignoreWeekendGroup:true});if(u)assign(model,state,slot,u,out);}}

export function generateSchedule(data,{startDate,endDate,mode='NEW'}={}){
  const model=buildModel(data),start=parseDate(startDate),end=parseDate(endDate);if(!start||!end||end<start)throw new Error('Valid start and end dates are required.');const days=daysBetween(start,end)+1;if((start.getDay()!==0||end.getDay()!==6||days%7!==0)&&model.data.users.some(u=>yes(u.Active)&&isRegularUser(u)))throw new Error('Generation must start Sunday, end Saturday, and contain complete weeks.');
  const generationId=`GEN_${dk(start).replaceAll('-','')}_${dk(end).replaceAll('-','')}_${Date.now()}`;const allExisting=data.schedule||[];const overlap=allExisting.filter(r=>{const d=parseDate(r.Date);return d&&d>=start&&d<=end;});if(mode==='NEW'&&overlap.length)throw new Error('Selected dates overlap an existing schedule. Use REGENERATE.');if(overlap.some(r=>clean(r['Finalized At'])) )throw new Error('Selected dates overlap a finalized schedule. Unfinalize it first.');
  const preserved=mode==='REGENERATE'?overlap.filter(r=>yes(r.Locked)):[];const outside=allExisting.filter(r=>{const d=parseDate(r.Date);return !d||d<start||d>end;});const slots=createSlots(model,start,end,generationId),out=new Map(),state=makeState(model,outside.concat(preserved),start,end);
  preserved.forEach(r=>out.set(`${dk(r.Date)}|${clean(r.Shift).toUpperCase()}|${num(r.Slot,1)}`,r));
  assignSeven(model,state,slots,out);assignWeekend(model,state,slots,out);assignEdPairs(model,state,slots,out);assignResidentE2(model,state,slots,out);
  for(const slot of slots){if(out.has(slot.slotKey))continue;const u=pick(model,state,slot);if(u)assign(model,state,slot,u,out);else{const cover=pickOffDayCoverage(model,state,slot);if(cover)assign(model,state,slot,cover.user,out,'ASSIGNED','',cover);else assign(model,state,slot,null,out,'UNFILLED','No eligible pharmacist found.');}}
  const generated=slots.map(s=>out.get(s.slotKey));const schedule=outside.concat(generated).sort((a,b)=>clean(a.Date).localeCompare(clean(b.Date))||clean(a.Shift).localeCompare(clean(b.Shift))||num(a.Slot,1)-num(b.Slot,1));const report=validateAssignments(schedule,data,start,end);return {schedule,generationId,report,health:scheduleHealth(generated)};
}

export function validateAssignments(schedule,data,startDate,endDate){
  const model=buildModel(data),start=parseDate(startDate),end=parseDate(endDate),errors=[],warnings=[];const rows=(schedule||[]).filter(r=>{const d=parseDate(r.Date);return d&&d>=start&&d<=end;});const seen={};
  rows.forEach(r=>{if(clean(r.Status).toUpperCase()==='UNFILLED'||!clean(r.Username)){errors.push(`${r.Date} ${r.Shift} slot ${r.Slot} is unfilled.`);return;}const u=model.byUsername[clean(r.Username)],sh=model.shifts[clean(r.Shift).toUpperCase()];if(!u||!sh){errors.push(`${r.Date} ${r.Shift}: employee or shift record is missing.`);return;}const k=`${clean(r.Username)}|${r.Date}`;if(seen[k])errors.push(`${u['Pharmacist Name']} has more than one shift on ${r.Date}.`);seen[k]=r;if(!model.skills[clean(u.Username)]?.has(clean(r['Required Skill']).toUpperCase()))errors.push(`${u['Pharmacist Name']} lacks ${r['Required Skill']} on ${r.Date}.`);if(isOnPto(u,r.Date,model.pto))errors.push(`${u['Pharmacist Name']} is scheduled during PTO on ${r.Date}.`);if(isSevenUser(u)&&!isSevenOnDate(u,r.Date))errors.push(`${u['Pharmacist Name']} is scheduled during a 7-on/7-off OFF block on ${r.Date}.`);});
  const activeRegular=model.data.users.filter(u=>yes(u.Active)&&isRegularUser(u));const weeks=[...new Set(eachDate(start,end).map(d=>dk(weekStartSunday(d))))];
  activeRegular.forEach(u=>{const urows=rows.filter(r=>clean(r.Username)===clean(u.Username)&&clean(r.Status).toUpperCase()!=='UNFILLED');for(const wk of weeks){const wr=urows.filter(r=>dk(weekStartSunday(r.Date))===wk);const days=new Set(wr.map(r=>r.Date));if(days.size!==5)errors.push(`${u['Pharmacist Name']} works ${days.size} days in week ${wk}; exactly 5 are required.`);for(let i=1;i<wr.sort((a,b)=>clean(a.Date).localeCompare(clean(b.Date))).length;i++){const prev=wr[i-1],cur=wr[i];if(daysBetween(prev.Date,cur.Date)===1&&isEveningShift(model.shifts[clean(prev.Shift).toUpperCase()])&&isDayShift(model.shifts[clean(cur.Shift).toUpperCase()]))errors.push(`${u['Pharmacist Name']} has an evening-to-morning transition ${prev.Date} → ${cur.Date}.`);}}
    const target=num(model.settings['Required Hours Per Schedule Period'],320),total=urows.reduce((s,r)=>s+num(r['Credited Hours'],0),0);if(daysBetween(start,end)+1===56&&Math.abs(total-target)>0.001)errors.push(`${u['Pharmacist Name']} has ${total} credited hours; exactly ${target} are required for this 56-day period.`);
    const dates=[...new Set(urows.map(r=>r.Date))].sort();let run=1,maxRun=dates.length?1:0;for(let i=1;i<dates.length;i++){if(daysBetween(dates[i-1],dates[i])===1)run++;else run=1;maxRun=Math.max(maxRun,run);}if(maxRun>num(model.settings['Maximum Consecutive Workdays'],5))errors.push(`${u['Pharmacist Name']} has ${maxRun} consecutive workdays.`);
  });
  const code=clean(model.settings['Resident E2 Shift Code']||'E2').toUpperCase(),need=num(model.settings['Resident E2 Shifts Per Week'],1);model.data.users.filter(u=>yes(u.Active)&&yes(u.Resident)&&isRegularUser(u)).forEach(u=>weeks.forEach(wk=>{const n=rows.filter(r=>clean(r.Username)===clean(u.Username)&&dk(weekStartSunday(r.Date))===wk&&clean(r.Shift).toUpperCase()===code).length;if(n<need)errors.push(`${u['Pharmacist Name']} has ${n} ${code} shift(s) in week ${wk}; ${need} required.`);}));
  const preTarget=num(model.settings['Preceptor Off Evening Target Per Month'],7);if(preTarget>0){const months=[...new Set(eachDate(start,end).map(monthKey))];model.data.users.filter(u=>yes(u.Active)&&yes(u.Preceptor)).forEach(u=>months.forEach(m=>{const hasOff=(model.data.preceptorCalendar||[]).some(r=>clean(r.Username)===clean(u.Username)&&monthKey(r.Date)===m&&clean(r.Mode).toUpperCase()==='OFF'&&yes(r.Active??'Yes'));if(!hasOff)return;const n=rows.filter(r=>clean(r.Username)===clean(u.Username)&&monthKey(r.Date)===m&&['E1','E2'].includes(clean(r.Shift).toUpperCase())).length;if(n<preTarget)errors.push(`${u['Pharmacist Name']} has ${n} E1/E2 shifts in ${m}; ${preTarget} are required when preceptor mode has OFF dates.`);}));}
  return {ok:errors.length===0,errors:[...new Set(errors)],warnings:[...new Set(warnings)]};
}
