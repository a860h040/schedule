import { clean, yes, num, uid, dateKey, addDays, dayName, isWeekend, startOfWeek, monthKey, minutes } from './schema.js';

const DAY_FIELDS=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const DAY_INDEX={SUN:0,MON:1,TUE:2,WED:3,THU:4,FRI:5,SAT:6};

function parseDate(k){return new Date(`${dateKey(k)}T12:00:00`);}
function norm(v){return clean(v).toUpperCase();}
function tokens(v){return String(v||'').split(/[,;|/\\\n]+/).map(x=>norm(x)).filter(Boolean);}
function datesBetween(start,end){const out=[];for(let d=parseDate(start),e=parseDate(end);d<=e;d=addDays(d,1))out.push(dateKey(d));return out;}
function rangeContains(a,b,x){return (!a||x>=dateKey(a))&&(!b||x<=dateKey(b));}
function dayField(date){return DAY_FIELDS[parseDate(date).getDay()];}
function weekKey(date,settings){return dateKey(startOfWeek(parseDate(date),settings.weekStart||'Sunday'));}
function assignmentHours(r){return num(r.creditedHours,r.hours||0);}
function shiftType(s){return norm(s?.type||'DAY');}
function isEvening(s){return shiftType(s)==='EVENING'||['E1','E2','E'].includes(norm(s?.shift||s?.code));}
function isNight(s){return shiftType(s)==='NIGHT'||['N1','N2'].includes(norm(s?.shift||s?.code));}
function isMorningOrDay(s){return ['DAY','MORNING','TRAINING','OTHER'].includes(shiftType(s))&&!isEvening(s)&&!isNight(s);}
function weekendSaturday(date){const d=parseDate(date);const day=d.getDay();if(day===6)return dateKey(d);if(day===0)return dateKey(addDays(d,-1));return dateKey(addDays(d,6-day));}
function daysDiff(a,b){return Math.round((parseDate(b)-parseDate(a))/86400000);}

export function buildModel(db){
  const settings=db.settings||{};
  const users=(db.users||[]).filter(u=>u.active!==false).map(u=>({...u,_username:clean(u.username),_name:clean(u.pharmacistName||u.name),_preferred:new Set(tokens(u.preferredShiftType)),_offSkills:new Set(tokens(u.offDayCoverageSkills))}));
  const usersByUsername=Object.fromEntries(users.map(u=>[u._username,u]));
  const shifts=(db.shifts||[]).filter(s=>s.active!==false).map(s=>({...s,shift:clean(s.shift||s.code)}));
  const shiftMap=Object.fromEntries(shifts.map(s=>[norm(s.shift),s]));
  const skillsByUser={};
  (db.skills||[]).filter(x=>x.active!==false).forEach(x=>{const u=clean(x.username);if(!skillsByUser[u])skillsByUser[u]=new Set();skillsByUser[u].add(norm(x.skill));});
  const requirements=(db.staffingRequirements||[]).filter(r=>r.active!==false);
  const pto=(db.requests||[]).filter(r=>norm(r.recordType)==='PTO'&&norm(r.status)==='APPROVED');
  const weeklyAvailability=(db.weeklyAvailability||[]).filter(r=>r.active!==false);
  const preceptorCalendar=db.preceptorCalendar||[];
  return {db,settings,users,usersByUsername,shifts,shiftMap,skillsByUser,requirements,pto,weeklyAvailability,preceptorCalendar};
}

export function currentWeekendGroup(date,settings){
  const seq=(settings.weekendRotation||['A','B','C']).map(norm).filter(Boolean);if(!seq.length)return'';
  const anchor=settings.weekendAnchorDate;if(!anchor)return'';
  const anchorGroup=norm(settings.weekendAnchorGroup||seq[0]);let idx=seq.indexOf(anchorGroup);if(idx<0)idx=0;
  const diffWeeks=Math.floor(daysDiff(weekendSaturday(anchor),weekendSaturday(date))/7);
  const pos=((idx+diffWeeks)%seq.length+seq.length)%seq.length;return seq[pos];
}
function employeeWeekendIsOn(user,date,settings){
  const seq=(settings.weekendRotation||['A','B','C']).map(norm).filter(Boolean);
  const anchor=dateKey(user.weekendRotationAnchorDate||'');
  if(!anchor||!seq.length)return true;
  const diffWeeks=Math.floor(daysDiff(weekendSaturday(anchor),weekendSaturday(date))/7);
  return ((diffWeeks%seq.length)+seq.length)%seq.length===0;
}

export function isSevenOn(user,date){
  if(norm(user.scheduleType)!=='7ON7OFF')return null;
  const anchor=dateKey(user.rotationAnchorDate);if(!anchor)return false;
  const diff=daysDiff(anchor,date);return ((diff%14)+14)%14<7;
}

function isPto(user,date,model){
  return model.pto.some(r=>clean(r.username)===user._username&&rangeContains(r.startDate||r.date,r.endDate||r.startDate||r.date,date));
}
function oneDayAvailabilityStatus(user,date,model){
  const rows=(model.db.requests||[]).filter(r=>norm(r.recordType)==='AVAILABILITY'&&clean(r.username)===user._username&&dateKey(r.date)===dateKey(date)&&norm(r.status||'APPROVED')==='APPROVED');
  if(!rows.length)return {blocked:false};
  const last=[...rows].sort((a,b)=>String(a.updatedAt||a.submittedAt||'').localeCompare(String(b.updatedAt||b.submittedAt||''))).at(-1);
  return yes(last.available)?{blocked:false}:{blocked:true,reason:'UNAVAILABLE'};
}
function weeklyAvailabilityStatus(user,date,shift,model){
  const day=dayName(date);const rows=model.weeklyAvailability.filter(r=>clean(r.username)===user._username&&norm(r.day)===day&&rangeContains(r.effectiveStart,r.effectiveEnd,date));
  if(!rows.length)return {blocked:false};
  const usable=rows.filter(r=>r.available!==false&&norm(r.ruleType)!=='BLOCK'); if(!usable.length)return {blocked:true,reason:'WEEKLY_AVAILABILITY_DAY'};
  const sm=minutes(shift.start),em=minutes(shift.end);
  if(sm==null||em==null)return {blocked:false};
  const within=usable.some(r=>{const a=minutes(r.startTime),b=minutes(r.endTime);if(a==null||b==null)return true;return sm>=a&&em<=b;});
  return within?{blocked:false}:{blocked:true,reason:'WEEKLY_AVAILABILITY_TIME'};
}
function preceptorStatus(user,date,model){
  if(!yes(user.preceptor))return 'NA';
  const rows=model.preceptorCalendar.filter(r=>clean(r.username)===user._username&&dateKey(r.date)===dateKey(date));
  if(rows.length)return yes(rows[0].precepting)||norm(rows[0].status)==='ON'?'ON':'OFF';
  return 'OFF';
}
function userHasNormalSkill(user,slot,model){
  const req=norm(slot.requiredSkill||slot.shift); if(!req)return true;
  const code=norm(slot.shift);
  if(yes(user.resident)&&code===norm(model.settings.residentE2ShiftCode||'E2'))return true;
  if(norm(user.scheduleType)==='7ON7OFF'&&preferredCodeMatches(user,slot))return true;
  if(yes(user.preceptor)){
    const ps=preceptorStatus(user,slot.date,model);
    if(ps==='ON'&&preferredCodeMatches(user,slot))return true;
    if(ps==='OFF'&&(['E1','E2'].includes(code)||preferredCodeMatches(user,slot)))return true;
  }
  return model.skillsByUser[user._username]?.has(req)||false;
}
function preferredMatches(user,slot){return user._preferred.has(norm(slot.shift))||user._preferred.has(norm(slot.requiredSkill))||user._preferred.has(norm(slot.shiftType));}
function preferredCodeMatches(user,slot){return user._preferred.has(norm(slot.shift));}
function homeOwners(slot,model){return model.users.filter(u=>norm(u.scheduleType)!=='7ON7OFF'&&preferredMatches(u,slot));}
function isGeneratedOffDay(owner,slot,state,model){
  if(isPto(owner,slot.date,model))return false;
  if(state.byEmployeeDate[`${owner._username}|${slot.date}`])return false;
  const wk=`${owner._username}|${weekKey(slot.date,model.settings)}`;
  return num(state.weeklyDays[wk],0)>=num(model.settings.regularWorkdaysPerWeek,5);
}
function offDayCoverageApplies(user,slot,state,model){
  const req=norm(slot.requiredSkill||slot.shift);if(!user._offSkills.has(req))return false;
  const owners=homeOwners(slot,model);
  if(owners.length!==1)return false;
  const owner=owners[0];
  if(isGeneratedOffDay(owner,slot,state,model))return true;
  const sameDay=state.byEmployeeDate[`${owner._username}|${slot.date}`];
  if(sameDay&&yes(owner.preceptor)&&preceptorStatus(owner,slot.date,model)==='OFF'){
    const sh=model.shiftMap[norm(sameDay.shift)]||sameDay;
    return isEvening(sh);
  }
  return false;
}
function shiftEligibilityFlags(user,slot,model,state,{allowOffDayCoverage=true,manual=false}={}){
  const reasons=[]; const warnings=[]; const shift=model.shiftMap[norm(slot.shift)]||slot.shiftDef||{};
  if(user.active===false)reasons.push('INACTIVE');
  if(isPto(user,slot.date,model))reasons.push('PTO');
  const da=oneDayAvailabilityStatus(user,slot.date,model);if(da.blocked)reasons.push(da.reason);
  const wa=weeklyAvailabilityStatus(user,slot.date,shift,model);if(wa.blocked)reasons.push(wa.reason);
  if(state.byEmployeeDate[`${user._username}|${slot.date}`])reasons.push('ALREADY_SCHEDULED');
  let skillOk=userHasNormalSkill(user,slot,model); let coverage=false;
  if(!skillOk&&allowOffDayCoverage&&offDayCoverageApplies(user,slot,state,model)){skillOk=true;coverage=true;}
  if(!skillOk)reasons.push('MISSING_SKILL');
  const wk=`${user._username}|${weekKey(slot.date,model.settings)}`;
  const currentWeekDays=num(state.weeklyDays[wk],0), currentWeekHours=num(state.weeklyHours[wk],0), add=assignmentHours(slot);
  const is7=norm(user.scheduleType)==='7ON7OFF';
  if(!is7){
    if(currentWeekDays>=num(model.settings.regularWorkdaysPerWeek,5))reasons.push('WEEKLY_DAYS');
    const max=num(user.weeklyHourMaximum,model.settings.weeklyHoursLimit||40);if(max>0&&currentWeekHours+add>max+0.001)reasons.push('WEEKLY_HOURS');
  }
  const periodHours=num(state.totalHours[user._username],0);const pmax=num(model.settings.maximumHoursPerSchedulePeriod,320);if(!is7&&pmax>0&&periodHours+add>pmax+0.001)reasons.push('TWO_MONTH_HOURS');
  const prev=state.byEmployeeDate[`${user._username}|${dateKey(addDays(slot.date,-1))}`];
  if(prev){const ps=model.shiftMap[norm(prev.shift)]||prev;if(isEvening(ps)&&isMorningOrDay(shift)&&yes(model.settings.blockEveningToMorningTransition))reasons.push('EVENING_TO_MORNING');}
  const next=state.byEmployeeDate[`${user._username}|${dateKey(addDays(slot.date,1))}`];
  if(next){const ns=model.shiftMap[norm(next.shift)]||next;if(isEvening(shift)&&isMorningOrDay(ns)&&yes(model.settings.blockEveningToMorningTransition))reasons.push('EVENING_TO_MORNING');}
  const maxCon=num(model.settings.maximumConsecutiveWorkdays,5);if(maxCon>0&&wouldExceedConsecutive(user._username,slot.date,state,maxCon))reasons.push('CONSECUTIVE_DAYS');
  if(isWeekend(slot.date)&&!yes(user.weekendEligible))reasons.push('WEEKEND_NOT_ELIGIBLE');
  if(isEvening(shift)&&!yes(user.eveningEligible))reasons.push('EVENING_NOT_ELIGIBLE');
  if(isNight(shift)&&!yes(user.nightEligible))reasons.push('NIGHT_NOT_ELIGIBLE');
  if(is7){
    if(!user.rotationAnchorDate)reasons.push('SEVEN_MISSING_ANCHOR');
    else if(!isSevenOn(user,slot.date))reasons.push('SEVEN_OFF');
    const allowed=user._preferred.size?user._preferred:new Set(tokens(user.defaultShift));
    if(allowed.size&&!allowed.has(norm(slot.shift)))reasons.push('SEVEN_WRONG_SHIFT');
  }else if(isWeekend(slot.date)&&yes(model.settings.requiredWeekendAssignment)){
    const group=currentWeekendGroup(slot.date,model.settings);const ug=norm(user.weekendGroup);
    if(group&&ug&&group!==ug&&!yes(model.settings.allowWeekendFallback))reasons.push(yes(user.resident)?'RESIDENT_WRONG_WEEKEND_GROUP':'WRONG_WEEKEND_GROUP');
    if(ug&&group===ug&&!employeeWeekendIsOn(user,slot.date,model.settings))reasons.push('WEEKEND_ANCHOR_OFF');
  }
  const pstat=preceptorStatus(user,slot.date,model);
  if(yes(user.preceptor)&&!isWeekend(slot.date)){
    if(pstat==='ON'&&yes(model.settings.preceptorHomeUnitRequiredWhenOn)&&!preferredMatches(user,slot))reasons.push('PRECEPTOR_HOME_UNIT');
    if(pstat==='OFF'){
      const allowed=preferredMatches(user,slot)||(yes(model.settings.preceptorOffAllowsEvenings)&&['E1','E2'].includes(norm(slot.shift)));
      if(!allowed)reasons.push('PRECEPTOR_OFF_SHIFT');
    }
    if(pstat==='ON'&&isEvening(shift)&&!preferredMatches(user,slot))reasons.push('PRECEPTOR_WEEKDAY_EVENING');
  }
  if(yes(user.resident)){
    if(norm(slot.shift)===norm(model.settings.residentE2ShiftCode||'E2')&&num(state.weeklyE2[wk],0)>=num(model.settings.residentE2ShiftsPerWeek,1))reasons.push('RESIDENT_E2_WEEKLY_LIMIT');
    if(isWeekend(slot.date)&&!yes(user.residentWeekends))reasons.push('RESIDENT_RESTRICTION');
    if(isEvening(shift)&&!yes(user.residentEvenings))reasons.push('RESIDENT_RESTRICTION');
    if(isNight(shift)&&!yes(user.residentNights))reasons.push('RESIDENT_RESTRICTION');
    if(!isWeekend(slot.date)&&isMorningOrDay(shift)&&!yes(user.residentCoversRegular)&&!preferredMatches(user,slot))reasons.push('RESIDENT_RESTRICTION');
  }
  const mk=`${user._username}|${monthKey(slot.date)}`; const eveningCount=num(state.monthlyEvenings[mk],0); const emax=num(user.maximumEveningShiftsPerMonth,model.settings.maximumEveningShifts||7);
  if(isEvening(shift)&&emax>0&&eveningCount>=emax)reasons.push('EVENING_LIMIT');
  if(user.customHoursEnabled&&norm(user.customHoursMode)==='HARD'&&shift.start&&shift.end&&user.preferredStartTime&&user.preferredEndTime){const s=minutes(shift.start),e=minutes(shift.end),a=minutes(user.preferredStartTime),b=minutes(user.preferredEndTime);if(s!=null&&e!=null&&a!=null&&b!=null&&(s<a||e>b))reasons.push('CUSTOM_HOURS');}
  if(manual&&reasons.length)warnings.push(...reasons);
  return {ok:reasons.length===0,reasons,warnings,coverage};
}
function wouldExceedConsecutive(username,date,state,max){
  let streak=1;for(let i=1;i<=max;i++){if(state.byEmployeeDate[`${username}|${dateKey(addDays(date,-i))}`])streak++;else break;}for(let i=1;i<=max;i++){if(state.byEmployeeDate[`${username}|${dateKey(addDays(date,i))}`])streak++;else break;}return streak>max;
}
function makeState(seed=[]){
  const s={assignments:[],byEmployeeDate:{},weeklyDays:{},weeklyHours:{},totalHours:{},monthlyEvenings:{},weeklyE2:{},weekends:{},shiftCounts:{}};
  seed.forEach(r=>applyAssignment(r,s,null));return s;
}
function applyAssignment(row,state,model){
  state.assignments.push(row); if(!row.username||norm(row.status)==='UNFILLED')return;
  state.byEmployeeDate[`${row.username}|${row.date}`]=row;
  const wk=`${row.username}|${weekKey(row.date,model?.settings||{weekStart:'Sunday'})}`;state.weeklyDays[wk]=num(state.weeklyDays[wk],0)+1;state.weeklyHours[wk]=num(state.weeklyHours[wk],0)+assignmentHours(row);
  state.totalHours[row.username]=num(state.totalHours[row.username],0)+assignmentHours(row);
  const mk=`${row.username}|${monthKey(row.date)}`;if(model&&isEvening(model.shiftMap[norm(row.shift)]||row))state.monthlyEvenings[mk]=num(state.monthlyEvenings[mk],0)+1;
  if(norm(row.shift)==='E2')state.weeklyE2[wk]=num(state.weeklyE2[wk],0)+1;
  if(isWeekend(row.date)){const sat=weekendSaturday(row.date);const k=`${row.username}|${sat}`;state.weekends[k]=true;}
  const sk=`${row.username}|${norm(row.shift)}`;state.shiftCounts[sk]=num(state.shiftCounts[sk],0)+1;
}
function removeStateAssignment(row,state,model){
  const idx=state.assignments.indexOf(row);if(idx>=0)state.assignments.splice(idx,1);if(!row.username||norm(row.status)==='UNFILLED')return;
  delete state.byEmployeeDate[`${row.username}|${row.date}`];const wk=`${row.username}|${weekKey(row.date,model.settings)}`;state.weeklyDays[wk]=Math.max(0,num(state.weeklyDays[wk])-1);state.weeklyHours[wk]=Math.max(0,num(state.weeklyHours[wk])-assignmentHours(row));state.totalHours[row.username]=Math.max(0,num(state.totalHours[row.username])-assignmentHours(row));
}
function generateSlots(startDate,endDate,model){
  const slots=[];for(const dk of datesBetween(startDate,endDate)){const field=dayField(dk);for(const r of model.requirements){const shiftCode=clean(r.shift);const shift=model.shiftMap[norm(shiftCode)];if(!shift)continue;const count=Math.max(0,Math.floor(num(r[field]??r[field.toLowerCase()],0)));for(let slot=1;slot<=count;slot++)slots.push({assignmentId:uid('ASN'),generationId:'',date:dk,day:dayName(dk),shift:shiftCode,slot,hours:num(shift.hours,0),creditedHours:num(shift.creditedHours,shift.hours||0),requiredSkill:clean(shift.skill||shiftCode),coverageForPharmacist:'',coverageForUsername:'',coverageReason:'',shiftType:clean(shift.type||'Day'),weekend:isWeekend(dk),weekendGroup:currentWeekendGroup(dk,model.settings),holiday:false,locked:false,manual:false,status:'UNFILLED',warning:'',updatedAt:new Date().toISOString(),updatedBy:'SYSTEM'});}}
  return slots.sort((a,b)=>a.date.localeCompare(b.date)||num(model.shiftMap[norm(a.shift)]?.priority,50)-num(model.shiftMap[norm(b.shift)]?.priority,50)||a.shift.localeCompare(b.shift)||a.slot-b.slot);
}
function scoreCandidate(user,slot,model,state){
  let score=1000;
  const target=num(user.targetWeeklyHours,model.settings.weeklyHoursLimit||40);
  const wk=`${user._username}|${weekKey(slot.date,model.settings)}`;
  const currentDays=num(state.weeklyDays[wk],0),wh=num(state.weeklyHours[wk],0);
  if(preferredMatches(user,slot))score+=40;
  if(preferredCodeMatches(user,slot))score+=1400;
  if(norm(user.scheduleType)!=='7ON7OFF'){
    const deficit=Math.max(0,num(model.settings.regularWorkdaysPerWeek,5)-currentDays);
    score+=deficit*2500;
    if(deficit===1)score+=1000;
  }
  score+=Math.max(-40,Math.min(60,(target-wh)*2));
  if(norm(user.scheduleType)!=='7ON7OFF'){
    const total=num(state.totalHours[user._username],0),periodTarget=num(model.settings.requiredHoursPerSchedulePeriod,320),credit=assignmentHours(slot);
    const remainingBefore=periodTarget-total,remainingAfter=periodTarget-(total+credit);
    score+=Math.max(0,remainingBefore)*8;
    if(Math.abs(remainingAfter)<0.0001)score+=10000;else score+=Math.max(0,500-Math.abs(remainingAfter)*4);
  }
  if(isWeekend(slot.date)&&norm(user.weekendGroup)===currentWeekendGroup(slot.date,model.settings))score+=2500;
  if(yes(user.resident)&&norm(slot.shift)===norm(model.settings.residentE2ShiftCode||'E2')&&num(state.weeklyE2[wk],0)<num(model.settings.residentE2ShiftsPerWeek,1))score+=5000;
  if(yes(user.preceptor)){const ps=preceptorStatus(user,slot.date,model);if(ps==='ON'&&preferredMatches(user,slot))score+=1400;if(ps==='OFF'&&isEvening(model.shiftMap[norm(slot.shift)]||slot))score+=300;}
  const totalAssignments=Object.keys(state.byEmployeeDate).filter(k=>k.startsWith(user._username+'|')).length;score-=totalAssignments*1.5;
  const mk=`${user._username}|${monthKey(slot.date)}`;if(isEvening(model.shiftMap[norm(slot.shift)]||slot))score-=num(state.monthlyEvenings[mk],0)*10;
  const sc=num(state.shiftCounts[`${user._username}|${norm(slot.shift)}`],0);score-=sc*2;
  if(norm(user.scheduleType)==='7ON7OFF'&&isSevenOn(user,slot.date))score+=5000;
  return score + (String(user._username+'|'+slot.date+'|'+slot.shift).split('').reduce((a,c)=>a+c.charCodeAt(0),0)%100)/10000;
}
function assignUserToSlot(user,slot,model,state,reason='AUTO'){
  const e=shiftEligibilityFlags(user,slot,model,state,{allowOffDayCoverage:true});if(!e.ok)return false;
  slot.assignedPharmacist=user._name;slot.username=user._username;slot.status='ASSIGNED';slot.updatedAt=new Date().toISOString();slot.updatedBy='SCHEDULER';
  if(e.coverage){const owner=homeOwners(slot,model)[0];slot.coverageForPharmacist=owner?owner._name:'';slot.coverageForUsername=owner?owner._username:'';slot.coverageReason='OFF_DAY_COVERAGE';}
  slot.assignmentReason=reason;applyAssignment(slot,state,model);return true;
}
function preassignPreceptorEvenings(slots,model,state,start,end){
  const months=[...new Set(datesBetween(start,end).map(monthKey))];
  const targetDefault=Math.max(0,num(model.settings.preceptorEveningTargetPerMonthWhenOff,7));
  for(const u of model.users.filter(x=>yes(x.preceptor)&&targetDefault>0)){
    for(const mk of months){
      const periodDates=datesBetween(start,end).filter(d=>monthKey(d)===mk);
      const weekdays=periodDates.filter(d=>![0,6].includes(parseDate(d).getDay()));
      if(!weekdays.length)continue;
      const fullPrecepting=weekdays.every(d=>preceptorStatus(u,d,model)==='ON');
      if(fullPrecepting)continue;
      const key=`${u._username}|${mk}`;
      let current=num(state.monthlyEvenings[key],0);
      const target=Math.min(num(u.maximumEveningShiftsPerMonth,targetDefault),targetDefault);
      if(current>=target)continue;
      const candidates=slots.filter(s=>monthKey(s.date)===mk&&norm(s.status)==='UNFILLED'&&['E1','E2'].includes(norm(s.shift))&&preceptorStatus(u,s.date,model)==='OFF').sort((a,b)=>a.date.localeCompare(b.date)||a.shift.localeCompare(b.shift));
      for(const slot of candidates){
        if(current>=target)break;
        if(assignUserToSlot(u,slot,model,state,'PRECEPTOR_OFF_EVENING'))current=num(state.monthlyEvenings[key],0);
      }
    }
  }
}
function clonePlanningState(state){
  return {
    assignments:[...state.assignments],
    byEmployeeDate:{...state.byEmployeeDate},
    weeklyDays:{...state.weeklyDays},
    weeklyHours:{...state.weeklyHours},
    totalHours:{...state.totalHours},
    monthlyEvenings:{...state.monthlyEvenings},
    weeklyE2:{...state.weeklyE2},
    weekends:{...state.weekends},
    shiftCounts:{...state.shiftCounts}
  };
}
function weekendPairDays(date){
  const sat=weekendSaturday(date);
  return {sat,sun:dateKey(addDays(sat,1))};
}
function bestWeekendPairForUser(user,sat,sun,slots,model,state){
  const satExisting=state.byEmployeeDate[`${user._username}|${sat}`];
  const sunExisting=state.byEmployeeDate[`${user._username}|${sun}`];
  if(satExisting&&sunExisting)return null;

  // If one day is already assigned (including a locked/manual assignment),
  // only complete the missing half of that same weekend.
  if(satExisting||sunExisting){
    const missingDate=satExisting?sun:sat;
    const ranked=slots
      .filter(s=>s.date===missingDate&&norm(s.status)==='UNFILLED')
      .map(s=>({s,e:shiftEligibilityFlags(user,s,model,state,{allowOffDayCoverage:true})}))
      .filter(x=>x.e.ok)
      .map(x=>({s:x.s,score:scoreCandidate(user,x.s,model,state)+(preferredMatches(user,x.s)?80:0)}))
      .sort((a,b)=>b.score-a.score||a.s.shift.localeCompare(b.s.shift));
    if(!ranked.length)return null;
    return {
      satSlot:satExisting?null:ranked[0].s,
      sunSlot:sunExisting?null:ranked[0].s,
      score:100000+ranked[0].score,
      completion:true
    };
  }

  const satSlots=slots.filter(s=>s.date===sat&&norm(s.status)==='UNFILLED');
  const sunSlots=slots.filter(s=>s.date===sun&&norm(s.status)==='UNFILLED');
  let best=null;

  for(const satSlot of satSlots){
    const satEligibility=shiftEligibilityFlags(user,satSlot,model,state,{allowOffDayCoverage:true});
    if(!satEligibility.ok)continue;

    // Evaluate Sunday after hypothetically assigning Saturday so consecutive-day,
    // hour, and transition rules are evaluated exactly as they will be committed.
    const temp=clonePlanningState(state);
    applyAssignment({
      ...satSlot,
      assignedPharmacist:user._name,
      username:user._username,
      status:'ASSIGNED'
    },temp,model);

    for(const sunSlot of sunSlots){
      const sunEligibility=shiftEligibilityFlags(user,sunSlot,model,temp,{allowOffDayCoverage:true});
      if(!sunEligibility.ok)continue;
      const score=
        scoreCandidate(user,satSlot,model,state)+
        scoreCandidate(user,sunSlot,model,temp)+
        (preferredMatches(user,satSlot)?80:0)+
        (preferredMatches(user,sunSlot)?80:0);
      if(!best||score>best.score)best={satSlot,sunSlot,score,completion:false};
    }
  }
  return best;
}
function commitWeekendPair(user,pair,model,state,reason){
  if(!pair)return false;
  if(pair.satSlot&&!assignUserToSlot(user,pair.satSlot,model,state,reason))return false;
  if(pair.sunSlot&&!assignUserToSlot(user,pair.sunSlot,model,state,reason))return false;
  return true;
}
function preassignWeekendResidentsAndTeams(slots,model,state){
  const dateSet=new Set(slots.map(s=>s.date));
  const weekendSaturdays=[...new Set(slots.filter(s=>isWeekend(s.date)).map(s=>weekendSaturday(s.date)))].sort();

  for(const sat of weekendSaturdays){
    const {sun}=weekendPairDays(sat);
    // Do not force a pair at a generation boundary where one half is outside
    // the selected schedule period.
    if(!dateSet.has(sat)||!dateSet.has(sun))continue;

    const group=currentWeekendGroup(sat,model.settings);
    const team=model.users
      .filter(u=>
        norm(u.scheduleType)!=='7ON7OFF'&&
        norm(u.weekendGroup)===group&&
        employeeWeekendIsOn(u,sat,model.settings)
      )
      .sort((a,b)=>(yes(b.resident)?1:0)-(yes(a.resident)?1:0)||a._name.localeCompare(b._name));

    for(const u of team){
      const pair=bestWeekendPairForUser(u,sat,sun,slots,model,state);
      if(!pair)continue;
      commitWeekendPair(
        u,
        pair,
        model,
        state,
        pair.completion?'WEEKEND_PAIR_COMPLETION':(yes(u.resident)?'REQUIRED_RESIDENT_WEEKEND_PAIR':'WEEKEND_TEAM_PAIR')
      );
    }
  }
}
function fillRemainingWeekendPairs(slots,model,state){
  if(!yes(model.settings.requiredWeekendAssignment))return;
  const dateSet=new Set(slots.map(s=>s.date));
  const weekendSaturdays=[...new Set(slots.filter(s=>isWeekend(s.date)).map(s=>weekendSaturday(s.date)))].sort();

  for(const sat of weekendSaturdays){
    const {sun}=weekendPairDays(sat);
    if(!dateSet.has(sat)||!dateSet.has(sun))continue;

    while(true){
      const satOpen=slots.some(s=>s.date===sat&&norm(s.status)==='UNFILLED');
      const sunOpen=slots.some(s=>s.date===sun&&norm(s.status)==='UNFILLED');
      if(!satOpen&&!sunOpen)break;

      let best=null;
      for(const u of model.users.filter(x=>norm(x.scheduleType)!=='7ON7OFF')){
        const pair=bestWeekendPairForUser(u,sat,sun,slots,model,state);
        if(!pair)continue;
        if(!best||pair.score>best.pair.score)best={u,pair};
      }
      if(!best)break;
      if(!commitWeekendPair(best.u,best.pair,model,state,best.pair.completion?'WEEKEND_PAIR_COMPLETION':'WEEKEND_PAIR_FILL'))break;
    }
  }
}
function preassignPairedEd(slots,model,state){
  if(!yes(model.settings.eddEdePairRequired))return;
  const dayCode=norm(model.settings.edDayShiftCode||'EDD'),eveCode=norm(model.settings.edEveningShiftCode||'EDE');
  const dates=[...new Set(slots.filter(s=>!isWeekend(s.date)&&[dayCode,eveCode].includes(norm(s.shift))).map(s=>s.date))].sort();
  for(const dk of dates){
    const ds=slots.find(s=>s.date===dk&&norm(s.shift)===dayCode&&norm(s.status)==='UNFILLED');
    const es=slots.find(s=>s.date===dk&&norm(s.shift)===eveCode&&norm(s.status)==='UNFILLED');
    if(!ds||!es)continue;
    let best=null;
    const dc=model.users.filter(u=>shiftEligibilityFlags(u,ds,model,state).ok);
    const ec=model.users.filter(u=>shiftEligibilityFlags(u,es,model,state).ok);
    for(const du of dc)for(const eu of ec){if(du._username===eu._username)continue;const score=scoreCandidate(du,ds,model,state)+scoreCandidate(eu,es,model,state)+(preferredMatches(du,ds)?500:0)+(preferredMatches(eu,es)?500:0);if(!best||score>best.score)best={du,eu,score};}
    if(best){assignUserToSlot(best.eu,es,model,state,'PAIRED_ED_EVENING');assignUserToSlot(best.du,ds,model,state,'PAIRED_ED_DAY');}
  }
}
function bestCandidate(slot,model,state,filterFn=null){
  const list=model.users.filter(u=>!filterFn||filterFn(u)).map(u=>({u,e:shiftEligibilityFlags(u,slot,model,state,{allowOffDayCoverage:true})})).filter(x=>x.e.ok).map(x=>({u:x.u,score:scoreCandidate(x.u,slot,model,state)})).sort((a,b)=>b.score-a.score||a.u._name.localeCompare(b.u._name));return list[0]?.u||null;
}

export function validateConfiguration(db){
  const model=buildModel(db),errors=[],warnings=[];
  if(norm(model.settings.weekStart)!=='SUNDAY')errors.push('Week Start must be Sunday.');
  if(num(model.settings.regularWorkdaysPerWeek,5)!==5)errors.push('Regular Workdays Per Week must be 5.');
  if(num(model.settings.maximumConsecutiveWorkdays,5)!==5)errors.push('Maximum Consecutive Workdays must be 5.');
  if(num(model.settings.requiredHoursPerSchedulePeriod,320)!==320)warnings.push('Required schedule-period hours is not 320.');
  const shiftCodes=new Set(model.shifts.map(s=>norm(s.shift)));
  for(const s of model.shifts){if(!s.start||!s.end)errors.push(`Active shift ${s.shift} is missing start/end time.`);if(num(s.creditedHours,s.hours)<=0)errors.push(`Active shift ${s.shift} has no credited hours.`);const req=norm(s.skill||s.shift);if(req){const qualified=model.users.filter(u=>model.skillsByUser[u._username]?.has(req));if(!qualified.length)warnings.push(`No active pharmacist has normal scheduling skill ${req} for ${s.shift}.`);}}
  for(const r of model.requirements)if(!shiftCodes.has(norm(r.shift)))errors.push(`Staffing requirement references unknown/inactive shift ${r.shift}.`);
  const groups=(model.settings.weekendRotation||['A','B','C']).map(norm);
  for(const u of model.users){if(norm(u.scheduleType)==='7ON7OFF'&&!u.rotationAnchorDate)errors.push(`${u._name}: 7-on/7-off anchor date is missing.`);if(norm(u.scheduleType)!=='7ON7OFF'&&u.weekendGroup&&!groups.includes(norm(u.weekendGroup)))errors.push(`${u._name}: invalid weekend group ${u.weekendGroup}.`);if(norm(u.scheduleType)!=='PRN'&&num(u.weeklyHourMaximum,0)<=0)warnings.push(`${u._name}: weekly hour maximum is missing.`);}
  const sevenGroups={};for(const u of model.users.filter(x=>norm(x.scheduleType)==='7ON7OFF')){const code=[...u._preferred].find(c=>model.shiftMap[c]);if(!code){errors.push(`${u._name}: 7-on/7-off employee needs a dedicated active shift in Preferred Shift Type.`);continue;}(sevenGroups[code]??=[]).push(u);}
  for(const [code,list] of Object.entries(sevenGroups)){for(let i=0;i<list.length;i++)for(let j=i+1;j<list.length;j++){if(!list[i].rotationAnchorDate||!list[j].rotationAnchorDate)continue;const mod=((daysDiff(list[i].rotationAnchorDate,list[j].rotationAnchorDate)%14)+14)%14;if(mod!==7)errors.push(`7-on/7-off ${code} employees ${list[i]._name} and ${list[j]._name} must have opposite rotations (anchor dates 7 days apart modulo 14).`);}}
  return {ok:errors.length===0,errors,warnings};
}

export function generateSchedule(db,startDate,{days,regenerate=true,actor='SYSTEM'}={}){
  const cfg=validateConfiguration(db);if(!cfg.ok)throw new Error('Configuration errors:\n'+cfg.errors.join('\n'));
  const model=buildModel(db);const start=dateKey(startDate);const count=Math.max(1,Math.min(75,Math.floor(num(days,model.settings.scheduleDays||60))));const end=dateKey(addDays(start,count-1));
  const generationId=`GEN_${start.replace(/-/g,'')}_${end.replace(/-/g,'')}_${Date.now().toString(36).toUpperCase()}`;
  const existing=(db.schedule||[]).filter(r=>r.date>=start&&r.date<=end);const locked=existing.filter(r=>yes(r.locked));
  const retained=(db.schedule||[]).filter(r=>r.date<start||r.date>end);const slots=generateSlots(start,end,model);slots.forEach(s=>s.generationId=generationId);
  const state=makeState([]);locked.forEach(r=>applyAssignment(r,state,model));
  const keyOf=r=>`${r.date}|${norm(r.shift)}|${Number(r.slot)||1}`;const lockedMap=new Map(locked.map(r=>[keyOf(r),r]));
  for(let i=slots.length-1;i>=0;i--){const l=lockedMap.get(keyOf(slots[i]));if(l){slots.splice(i,1);}}
  // Pass 1: 7-on/7-off employees get their dedicated shift throughout ON blocks.
  for(const u of model.users.filter(x=>norm(x.scheduleType)==='7ON7OFF'))for(const dk of datesBetween(start,end)){if(!isSevenOn(u,dk))continue;const candidates=slots.filter(s=>s.date===dk&&norm(s.status)==='UNFILLED'&&(u._preferred.size?u._preferred.has(norm(s.shift)):true));for(const s of candidates){if(assignUserToSlot(u,s,model,state,'7ON7OFF'))break;}}
  // Pass 2: preceptors marked ON are placed in their home/preferred unit when that slot exists.
  for(const u of model.users.filter(x=>yes(x.preceptor)))for(const dk of datesBetween(start,end)){if(preceptorStatus(u,dk,model)!=='ON')continue;if(state.byEmployeeDate[`${u._username}|${dk}`])continue;const home=slots.filter(s=>s.date===dk&&norm(s.status)==='UNFILLED'&&preferredMatches(u,s));for(const s of home){if(assignUserToSlot(u,s,model,state,'PRECEPTOR_HOME'))break;}}
  // Pass 3: preceptors who are OFF rotation fill their monthly E1/E2 target; a full precepting month has no evening target.
  preassignPreceptorEvenings(slots,model,state,start,end);
  // Pass 4: reserve assigned A/B/C weekend team members; residents are prioritized and required on their assigned weekend.
  preassignWeekendResidentsAndTeams(slots,model,state);
  // Pass 5: reserve paired weekday EDD/EDE coverage using two different pharmacists.
  preassignPairedEd(slots,model,state);
  // Pass 6: guarantee resident E2 target when a valid E2 slot exists.
  const weeks=[...new Set(datesBetween(start,end).map(d=>weekKey(d,model.settings)))];
  for(const u of model.users.filter(x=>yes(x.resident))){for(const wk of weeks){let need=Math.max(0,num(model.settings.residentE2ShiftsPerWeek,1)-num(state.weeklyE2[`${u._username}|${wk}`],0));if(!need)continue;const weekSlots=slots.filter(s=>weekKey(s.date,model.settings)===wk&&norm(s.status)==='UNFILLED'&&norm(s.shift)===norm(model.settings.residentE2ShiftCode||'E2'));for(const s of weekSlots){if(need<=0)break;if(assignUserToSlot(u,s,model,state,'RESIDENT_E2'))need--;}}}
  // Pass 7: fill remaining NON-WEEKEND slots by qualifications, constraints,
  // under-target hours, and fairness. Weekend assignments are handled only as
  // Saturday+Sunday pairs when requiredWeekendAssignment is enabled.
  for(const slot of slots){
    if(norm(slot.status)!=='UNFILLED')continue;
    if(yes(model.settings.requiredWeekendAssignment)&&isWeekend(slot.date))continue;
    const u=bestCandidate(slot,model,state);
    if(u)assignUserToSlot(u,slot,model,state,'BEST_FIT');
  }
  // Pass 8: repair/fill weekend coverage without ever creating a one-day
  // weekend assignment. A pharmacist is either OFF both days or works both.
  fillRemainingWeekendPairs(slots,model,state);
  // Attach useful warnings to unfilled slots.
  for(const slot of slots.filter(s=>norm(s.status)==='UNFILLED')){const reasonCounts={};for(const u of model.users){for(const r of shiftEligibilityFlags(u,slot,model,state).reasons)reasonCounts[r]=(reasonCounts[r]||0)+1;}slot.assignedPharmacist='UNFILLED';slot.username='';slot.warning=Object.entries(reasonCounts).sort((a,b)=>b[1]-a[1]).slice(0,3).map(([r,c])=>`${r}:${c}`).join('; ');}
  const periodRows=[...locked,...slots].sort((a,b)=>a.date.localeCompare(b.date)||a.shift.localeCompare(b.shift)||a.slot-b.slot);
  const nextSchedule=[...retained,...periodRows];
  const period={generationId,startDate:start,endDate:end,status:'DRAFT',createdBy:actor,createdAt:new Date().toISOString(),finalizedBy:'',finalizedAt:''};
  const periods=(db.schedulePeriods||[]).filter(p=>!(p.startDate===start&&p.endDate===end));periods.push(period);
  const validation=validateSchedule({...db,schedule:nextSchedule,schedulePeriods:periods},generationId);
  return {generationId,startDate:start,endDate:end,rows:periodRows,nextSchedule,periods,validation};
}

export function manualAssignmentWarnings(db,assignmentId,username){
  const model=buildModel(db);const row=db.schedule.find(r=>r.assignmentId===assignmentId);const user=model.usersByUsername[clean(username)];if(!row)throw new Error('Assignment not found.');if(!user)throw new Error('Pharmacist not found.');
  const state=makeState([]);for(const r of db.schedule){if(r.assignmentId===assignmentId||norm(r.status)==='UNFILLED')continue;applyAssignment(r,state,model);}const e=shiftEligibilityFlags(user,row,model,state,{allowOffDayCoverage:true,manual:true});
  const wk=weekKey(row.date,model.settings),month=monthKey(row.date);const weeklyBefore=num(state.weeklyHours[`${user._username}|${wk}`],0),monthlyBefore=db.schedule.filter(r=>r.username===user._username&&monthKey(r.date)===month&&norm(r.status)!=='UNFILLED').reduce((s,r)=>s+assignmentHours(r),0),periodBefore=num(state.totalHours[user._username],0);
  return {ok:e.ok,reasons:e.reasons,weeklyBefore,weeklyAfter:weeklyBefore+assignmentHours(row),monthlyBefore,monthlyAfter:monthlyBefore+assignmentHours(row),periodBefore,periodAfter:periodBefore+assignmentHours(row),coverage:e.coverage};
}

export function validateSchedule(db,generationId=''){
  const model=buildModel(db);const rows=(db.schedule||[]).filter(r=>!generationId||r.generationId===generationId);const warnings=[];
  const state=makeState([]);
  for(const r of rows.sort((a,b)=>a.date.localeCompare(b.date))){
    if(norm(r.status)==='UNFILLED'||!r.username){warnings.push({severity:'ERROR',type:'UNFILLED',date:r.date,shift:r.shift,assignmentId:r.assignmentId,message:`${r.date} ${r.shift} slot ${r.slot} is unfilled.`});continue;}
    const u=model.usersByUsername[r.username];if(!u){warnings.push({severity:'ERROR',type:'UNKNOWN_USER',date:r.date,shift:r.shift,assignmentId:r.assignmentId,message:`Unknown pharmacist ${r.username}.`});continue;}
    const e=shiftEligibilityFlags(u,r,model,state,{allowOffDayCoverage:true});for(const reason of e.reasons)warnings.push({severity:'ERROR',type:reason,date:r.date,shift:r.shift,pharmacist:u._name,assignmentId:r.assignmentId,message:reason});
    applyAssignment(r,state,model);
  }
  // exact five workdays per Sunday-Saturday week for regular active pharmacists when that week is fully inside the generated range.
  if(rows.length){const min=rows.map(r=>r.date).sort()[0],max=rows.map(r=>r.date).sort().at(-1);for(const u of model.users.filter(x=>norm(x.scheduleType)!=='7ON7OFF')){for(let w=startOfWeek(parseDate(min),model.settings.weekStart);w<=parseDate(max);w=addDays(w,7)){const ws=dateKey(w),we=dateKey(addDays(w,6));if(ws<min||we>max)continue;const days=num(state.weeklyDays[`${u._username}|${ws}`],0);const target=num(model.settings.regularWorkdaysPerWeek,5);if(days!==target)warnings.push({severity:'WARNING',type:'WEEKLY_DAYS_EXACT',date:ws,pharmacist:u._name,message:`${u._name} has ${days}/${target} workdays for week ${ws}.`});}}
    for(const u of model.users.filter(x=>norm(x.scheduleType)!=='7ON7OFF')){const h=num(state.totalHours[u._username],0),target=num(model.settings.requiredHoursPerSchedulePeriod,320);if(target>0&&Math.abs(h-target)>0.001)warnings.push({severity:'WARNING',type:'PERIOD_HOURS',pharmacist:u._name,message:`${u._name} has ${h} / ${target} credited hours in this schedule period.`});}
  }
  if(rows.length){
    const minDate=rows.map(r=>r.date).sort()[0],maxDate=rows.map(r=>r.date).sort().at(-1);
    for(const u of model.users.filter(x=>norm(x.scheduleType)==='7ON7OFF'&&x.rotationAnchorDate)){
      for(const dk of datesBetween(minDate,maxDate)){if(!isSevenOn(u,dk))continue;const assigned=rows.find(r=>r.date===dk&&r.username===u._username&&norm(r.status)!=='UNFILLED');if(!assigned)warnings.push({severity:'ERROR',type:'SEVEN_ON_MISSING_DAY',date:dk,pharmacist:u._name,message:`${u._name} is in a 7-on ON block but has no assignment.`});}
    }
    // Any regular/resident pharmacist who works one day of a complete weekend
    // must also work the other day. This is broader than the resident-only rule.
    const completeWeekendSats=[...new Set(rows.filter(r=>isWeekend(r.date)).map(r=>weekendSaturday(r.date)))].filter(sat=>sat>=minDate&&dateKey(addDays(sat,1))<=maxDate);
    for(const u of model.users.filter(x=>norm(x.scheduleType)!=='7ON7OFF')){
      for(const sat of completeWeekendSats){
        const sun=dateKey(addDays(sat,1));
        const hasSat=rows.some(r=>r.date===sat&&r.username===u._username&&norm(r.status)!=='UNFILLED');
        const hasSun=rows.some(r=>r.date===sun&&r.username===u._username&&norm(r.status)!=='UNFILLED');
        if(hasSat!==hasSun)warnings.push({
          severity:'ERROR',
          type:'WEEKEND_PAIR_REQUIRED',
          date:sat,
          pharmacist:u._name,
          message:`${u._name} is scheduled for only one day of weekend ${sat}/${sun}; Saturday and Sunday must be worked together.`
        });
      }
    }
    if(yes(model.settings.residentsRequiredAssignedWeekend)){for(const u of model.users.filter(x=>yes(x.resident)&&norm(x.scheduleType)!=='7ON7OFF'&&x.weekendGroup)){for(const sat of [...new Set(rows.filter(r=>isWeekend(r.date)).map(r=>weekendSaturday(r.date)))]){if(currentWeekendGroup(sat,model.settings)!==norm(u.weekendGroup)||!employeeWeekendIsOn(u,sat,model.settings))continue;for(const dk of [sat,dateKey(addDays(sat,1))]){if(dk<minDate||dk>maxDate)continue;if(!rows.some(r=>r.date===dk&&r.username===u._username&&norm(r.status)!=='UNFILLED'))warnings.push({severity:'ERROR',type:'RESIDENT_REQUIRED_WEEKEND',date:dk,pharmacist:u._name,message:`${u._name} must work both days of the assigned weekend.`});}}}}
    const e2Code=norm(model.settings.residentE2ShiftCode||'E2'),e2Need=num(model.settings.residentE2ShiftsPerWeek,1);const weekStarts=[...new Set(rows.map(r=>weekKey(r.date,model.settings)))];for(const wk of weekStarts){const weekRows=rows.filter(r=>weekKey(r.date,model.settings)===wk),hasE2=weekRows.some(r=>norm(r.shift)===e2Code);if(!hasE2)continue;for(const u of model.users.filter(x=>yes(x.resident))){const count=weekRows.filter(r=>r.username===u._username&&norm(r.shift)===e2Code&&norm(r.status)!=='UNFILLED').length;if(count!==e2Need)warnings.push({severity:'ERROR',type:'RESIDENT_E2_EXACT',date:wk,pharmacist:u._name,message:`${u._name} has ${count} ${e2Code} shift(s) in week ${wk}; exactly ${e2Need} required.`});}}
    const months=[...new Set(rows.map(r=>monthKey(r.date)))];for(const u of model.users.filter(x=>yes(x.preceptor))){for(const mk of months){const mdates=rows.map(r=>r.date).filter(d=>monthKey(d)===mk);const weekdays=[...new Set(mdates)].filter(d=>![0,6].includes(parseDate(d).getDay()));if(!weekdays.length||weekdays.every(d=>preceptorStatus(u,d,model)==='ON'))continue;const evenings=rows.filter(r=>r.username===u._username&&monthKey(r.date)===mk&&isEvening(model.shiftMap[norm(r.shift)]||r)&&norm(r.status)!=='UNFILLED').length;const target=Math.min(num(u.maximumEveningShiftsPerMonth,model.settings.preceptorEveningTargetPerMonthWhenOff||7),num(model.settings.preceptorEveningTargetPerMonthWhenOff,7));if(evenings<target)warnings.push({severity:'WARNING',type:'PRECEPTOR_EVENING_TARGET',pharmacist:u._name,message:`${u._name} has ${evenings}/${target} evening shifts in ${mk} while not precepting the full month.`});}}
  }
  if(yes(model.settings.eddEdePairRequired)){const dayCode=norm(model.settings.edDayShiftCode||'EDD'),eveCode=norm(model.settings.edEveningShiftCode||'EDE');const dates=[...new Set(rows.map(r=>r.date))];for(const dk of dates){const dr=rows.filter(r=>r.date===dk&&norm(r.status)!=='UNFILLED'&&norm(r.shift)===dayCode),er=rows.filter(r=>r.date===dk&&norm(r.status)!=='UNFILLED'&&norm(r.shift)===eveCode);if(dr.length&&er.length&&dr[0].username===er[0].username)warnings.push({severity:'ERROR',type:'PAIRED_ED_SAME_PHARMACIST',date:dk,message:`${dayCode}/${eveCode} paired coverage must use two different pharmacists.`});}}
  return {ok:!warnings.some(w=>w.severity==='ERROR'),warnings,errorCount:warnings.filter(w=>w.severity==='ERROR').length,warningCount:warnings.filter(w=>w.severity==='WARNING').length};
}

export function scheduleHealth(db){const rows=db.schedule||[],required=rows.length,unfilled=rows.filter(r=>norm(r.status)==='UNFILLED'||norm(r.assignedPharmacist)==='UNFILLED').length,filled=required-unfilled,warnings=rows.filter(r=>clean(r.warning)).length;return{required,filled,unfilled,coveragePercent:required?Math.round(filled*1000/required)/10:100,ruleWarnings:warnings};}

export function employeeStats(db){
  const model=buildModel(db),map={};for(const u of model.users)map[u._username]={username:u._username,name:u._name,totalHours:0,shifts:0,dayShifts:0,eveningShifts:0,nightShifts:0,weekendShifts:0,weeklyHours:{},scheduleType:u.scheduleType,weekendGroup:u.weekendGroup,preceptor:u.preceptor,resident:u.resident};
  for(const r of db.schedule||[]){if(!r.username||norm(r.status)==='UNFILLED')continue;const x=map[r.username]||(map[r.username]={username:r.username,name:r.assignedPharmacist,totalHours:0,shifts:0,dayShifts:0,eveningShifts:0,nightShifts:0,weekendShifts:0,weeklyHours:{}});const s=model.shiftMap[norm(r.shift)]||r;x.totalHours+=assignmentHours(r);x.shifts++;if(isEvening(s))x.eveningShifts++;else if(isNight(s))x.nightShifts++;else x.dayShifts++;if(isWeekend(r.date))x.weekendShifts++;const wk=weekKey(r.date,model.settings);x.weeklyHours[wk]=num(x.weeklyHours[wk],0)+assignmentHours(r);}
  return Object.values(map);
}
export function fairnessReport(db){const stats=employeeStats(db),avg=stats.length?stats.reduce((s,x)=>s+x.totalHours,0)/stats.length:0;return{averageHours:Math.round(avg*10)/10,employees:stats.map(x=>({...x,hoursVsAverage:Math.round((x.totalHours-avg)*10)/10,workloadFlag:x.totalHours>avg*1.25?'HIGH':x.totalHours<avg*0.75?'LOW':'NORMAL'}))};}

export function reconcilePto(db){
  const limit=Math.max(1,num(db.settings?.ptoAutoApprovalLimitPerDate,2));const pto=(db.requests||[]).filter(r=>norm(r.recordType)==='PTO'&&norm(r.status)!=='REJECTED');const byDate={};
  for(const r of pto){const start=dateKey(r.startDate||r.date),end=dateKey(r.endDate||r.startDate||r.date);for(const dk of datesBetween(start,end)){if(!byDate[dk])byDate[dk]=[];byDate[dk].push(r);}}
  Object.values(byDate).forEach(list=>list.sort((a,b)=>String(a.submittedAt||'').localeCompare(String(b.submittedAt||''))||String(a.recordId).localeCompare(String(b.recordId))));
  for(const r of pto){const ds=datesBetween(r.startDate||r.date,r.endDate||r.startDate||r.date);const auto=ds.length&&ds.every(d=>byDate[d].indexOf(r)<limit);const manuallyReviewed=clean(r.reviewedBy)&&norm(r.reviewedBy)!=='SYSTEM AUTO-APPROVAL';if(auto&&!manuallyReviewed){r.status='Approved';r.reviewedBy='SYSTEM AUTO-APPROVAL';r.reviewedAt=new Date().toISOString();}else if(!auto&&norm(r.status)==='APPROVED'&&norm(r.reviewedBy)==='SYSTEM AUTO-APPROVAL'){r.status='Pending';r.reviewedBy='';r.reviewedAt='';}}
  return db;
}

export function swapEligibleCounterShifts(db,posterUsername,targetUsername,offeredAssignmentId){
  const offered=db.schedule.find(r=>r.assignmentId===offeredAssignmentId);if(!offered)return[];
  const targetDates=new Set((db.schedule||[]).filter(r=>r.username===posterUsername).map(r=>r.date));
  return (db.schedule||[]).filter(r=>r.username===targetUsername&&norm(r.status)!=='UNFILLED'&&!targetDates.has(r.date)&&r.date>=new Date().toISOString().slice(0,10));
}
