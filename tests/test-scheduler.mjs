import { blankDatabase } from '../js/schema.js';
import { generateSchedule, validateConfiguration, validateSchedule, reconcilePto } from '../js/scheduler.js';
const db=blankDatabase();
db.settings.weekendAnchorDate='2026-09-19';
db.shifts=[
 {shift:'DAY',meaning:'Day',start:'07:00',end:'15:00',hours:8,creditedHours:8,type:'Day',skill:'DAY',weekend:true,active:true,priority:10},
 {shift:'E1',meaning:'Evening',start:'15:00',end:'23:00',hours:8,creditedHours:8,type:'Evening',skill:'E1',weekend:true,active:true,priority:20},
 {shift:'E2',meaning:'Evening 2',start:'15:00',end:'23:00',hours:8,creditedHours:8,type:'Evening',skill:'E2',weekend:true,active:true,priority:20},
 {shift:'EDD',meaning:'ED Day',start:'07:00',end:'15:00',hours:8,creditedHours:8,type:'Day',skill:'EDD',weekend:false,active:true,priority:15},
 {shift:'EDE',meaning:'ED Eve',start:'15:00',end:'23:00',hours:8,creditedHours:8,type:'Evening',skill:'EDE',weekend:false,active:true,priority:15}
];
db.staffingRequirements=db.shifts.map(s=>({shift:s.shift,Sunday:1,Monday:1,Tuesday:1,Wednesday:1,Thursday:1,Friday:1,Saturday:1,active:true}));
for(let i=0;i<15;i++){
 const u={id:'U'+i,pharmacistName:'P'+i,name:'P'+i,username:'u'+i,active:true,role:'Pharmacist',scheduleType:'Regular',preceptor:false,resident:i<2,weekendGroup:['A','B','C'][i%3],weekendRotationAnchorDate:['2026-09-19','2026-09-26','2026-10-03'][i%3],weeklyHourMaximum:40,targetWeeklyHours:40,maximumEveningShiftsPerMonth:7,preferredShiftType:i<2?'E2':'DAY',offDayCoverageSkills:[],weekendEligible:true,eveningEligible:true,nightEligible:true,residentCoversRegular:true,residentWeekends:true,residentEvenings:true,residentNights:false};db.users.push(u);
 for(const sk of ['DAY','E1','E2','EDD','EDE']) db.skills.push({id:`${i}-${sk}`,username:u.username,skill:sk,active:true});
}
console.log('config',validateConfiguration(db));
const out=generateSchedule(db,'2026-09-20',{days:14,actor:'test'});
db.schedule=out.nextSchedule;db.schedulePeriods=out.periods;
console.log('rows',out.rows.length,'unfilled',out.rows.filter(r=>r.status==='UNFILLED').length);
console.log('validation',validateSchedule(db,out.generationId).errorCount,validateSchedule(db,out.generationId).warningCount);
db.requests.push({recordType:'PTO',recordId:'a',username:'u1',pharmacist:'P1',startDate:'2026-10-01',endDate:'2026-10-01',status:'Pending',submittedAt:'2026-09-01T01:00:00Z'});
db.requests.push({recordType:'PTO',recordId:'b',username:'u2',pharmacist:'P2',startDate:'2026-10-01',endDate:'2026-10-01',status:'Pending',submittedAt:'2026-09-01T02:00:00Z'});
db.requests.push({recordType:'PTO',recordId:'c',username:'u3',pharmacist:'P3',startDate:'2026-10-01',endDate:'2026-10-01',status:'Pending',submittedAt:'2026-09-01T03:00:00Z'});
reconcilePto(db);console.log(db.requests.map(r=>r.status).join(','));
