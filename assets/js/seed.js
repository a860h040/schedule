import { nowIso, uuid } from './utils.js';

const PHARMACISTS = [
  ['Balolong, Marilyn','Regular','A',40,'TBA',false,false],
  ['Booth, Alyssa','Regular','C',40,'IM',false,false],
  ['Brewster, Clare','Regular','B',40,'TBA',false,false],
  ['Camp, Ken','Regular','B',40,'C7',false,false],
  ['Cahill, Karlena','Regular','A',40,'TBA',false,false],
  ['Cothran, Ashton','Regular','C',40,'TBA',false,false],
  ['Duck, Meagan','Regular','B',40,'ONC',false,false],
  ['Gignac, Lindsey','Regular','C',40,'CC2',false,false],
  ['Hayes, Sarah','Regular','B',40,'EDD',false,false],
  ['Ramos, Kasi','Regular','',40,'TBA',false,false],
  ['Reneau, Shannon','Regular','',40,'CC1',false,false],
  ['Rossi, Nicole','Regular','A',40,'TBA',false,false],
  ['Moncrieffe, Tiffany','Regular','',40,'E',false,false],
  ['Osman, Badawi','7ON7OFF','',40,'E',false,false],
  ['Parker, Matthew','7ON7OFF','',40,'',false,false],
  ['Howell, Stephen','7ON7OFF','',40,'',false,false],
  ['Zussy, Karina','7ON7OFF','',40,'',false,false],
  ['Raja, Nazish','7ON7OFF','',40,'',false,false],
  ['Finnick, Micheal','PRN','',0,'',false,false],
  ['Logue, Taylor','PRN','C',0,'',false,false],
  ['Simmons, Henry','PRN','',0,'TBA',false,false],
  ['Traugott, Matt','PRN','',0,'TBA',false,false],
  ['Usztok, Shannon','PRN','',0,'',false,false],
  ['Villanueva, Josiah','PRN','C',0,'',false,false],
  ['Hamza, Alhamza','Regular','',40,'',false,true],
  ['Harris, Taylor','Regular','',40,'',false,true],
  ['Kim, Erica','Regular','',40,'',false,true],
  ['Odei-Wontumi, Paa Kwesi','Regular','',40,'',false,true],
  ['Seely, Kirsten','Regular','',40,'',false,true]
];

export const DEFAULT_SETTINGS = [
  {Setting:'Default Schedule Days',Value:'56',Description:'Default generation length: 8 complete Sunday-Saturday weeks'},
  {Setting:'Weekly Hours Limit',Value:'40',Description:'Default weekly maximum if employee-specific value is blank'},
  {Setting:'Maximum Hours Per Schedule Period',Value:'320',Description:'Legacy hard maximum; kept for compatibility'},
  {Setting:'Required Hours Per Schedule Period',Value:'320',Description:'Regular pharmacists should finish an 8-week period with exactly this many credited hours'},
  {Setting:'Resident E2 Shift Code',Value:'E2',Description:'Residents receive this evening shift once per week'},
  {Setting:'Resident E2 Shifts Per Week',Value:'1',Description:'Required resident E2 shifts per week'},
  {Setting:'SevenOn Block Scheduling',Value:'Yes',Description:'Keep 7-on/7-off employees on their dedicated shift during ON blocks'},
  {Setting:'Week Start',Value:'Sunday',Description:'Scheduling week begins Sunday'},
  {Setting:'Regular Workdays Per Week',Value:'5',Description:'Every regular pharmacist works exactly five days in each complete week'},
  {Setting:'Maximum Consecutive Workdays',Value:'5',Description:'Hard maximum consecutive workdays for regular pharmacists'},
  {Setting:'Block Evening To Morning Transition',Value:'Yes',Description:'Evening cannot be followed by next-day Day/Morning shift'},
  {Setting:'Required Weekend Assignment',Value:'Yes',Description:'Use A/B/C weekend group rotation as a hard rule'},
  {Setting:'Residents Required Assigned Weekend',Value:'Yes',Description:'Residents cover their assigned A/B/C weekend'},
  {Setting:'Preceptor Weekday Evening Allowed',Value:'No',Description:'When precepting Monday-Friday, stay on home/preferred unit'},
  {Setting:'Preceptor Weekend Evening Allowed',Value:'Yes',Description:'Preceptors may work evening shifts on weekends'},
  {Setting:'Preceptor Off Evening Target Per Month',Value:'7',Description:'When a preceptor has OFF-preceptor dates in a month, target seven E1/E2 shifts for that month'},
  {Setting:'Require Paired ED Coverage',Value:'Yes',Description:'EDD and EDE should use two different qualified pharmacists when both are required'},
  {Setting:'ED Day Shift Code',Value:'EDD',Description:'ED day shift'},
  {Setting:'ED Evening Shift Code',Value:'EDE',Description:'ED evening shift'},
  {Setting:'Maximum Evening Shifts',Value:'7',Description:'Default monthly evening maximum'},
  {Setting:'Weekend Rotation',Value:'A,B,C',Description:'Weekend groups'},
  {Setting:'Weekend Anchor Date',Value:'2026-08-15',Description:'Anchor Saturday for Weekend Anchor Group'},
  {Setting:'Weekend Anchor Group',Value:'A',Description:'Group assigned to anchor Saturday'},
  {Setting:'Allow Weekend Fallback',Value:'No',Description:'Permit another weekend group only if manually overridden'},
  {Setting:'Allow Admin Rule Override',Value:'Yes',Description:'Administrators may manually override soft warnings'},
  {Setting:'PTO Auto Approval Limit',Value:'2',Description:'First two PTO requests for each calendar date are automatically approved'},
  {Setting:'Session Hours',Value:'8',Description:'Local browser app session length'}
];

export const DEFAULT_SHIFTS = [
  ['C7','Central Pharmacy 0700','07:00','15:30',8.5,8,'Day','C7','No','Yes',40],
  ['C8','Central Pharmacy 0800','08:00','16:30',8.5,8,'Day','C8','No','Yes',40],
  ['CARD','Cardiology','07:00','15:30',8.5,8,'Day','CARD','No','Yes',20],
  ['CC1','Critical Care','07:00','15:30',8.5,8,'Day','CC1','No','Yes',10],
  ['CC2','Critical Care','07:00','15:30',8.5,8,'Day','CC2','No','Yes',10],
  ['EDD','ED Day','07:00','15:30',8.5,8,'Day','EDD','No','Yes',15],
  ['IM','Internal Medicine','07:00','15:30',8.5,8,'Day','IM','No','Yes',20],
  ['ONC','Oncology','07:00','15:30',8.5,8,'Day','ONC','No','Yes',10],
  ['E1','Evenings','13:00','21:30',8.5,8,'Evening','E1','No','Yes',25],
  ['E2','Evenings','13:00','21:30',8.5,8,'Evening','E2','No','Yes',25],
  ['EDE','ED Evenings','14:00','22:30',8.5,8,'Evening','EDE','No','Yes',15],
  ['E','Evening Central','11:00','21:30',10.5,10,'Evening','E','Yes','Yes',15],
  ['WC7','Weekend Central','08:00','16:30',8.5,8,'Day','WC7','Yes','Yes',15],
  ['WD1','Weekend Clinical','07:00','15:30',8.5,8,'Day','WD1','Yes','Yes',10],
  ['WD2','Weekend Clinical','07:00','15:30',8.5,8,'Day','WD2','Yes','Yes',10],
  ['WE1','Evenings','13:00','21:30',8.5,8,'Evening','WE1','No','No',25],
  ['WEDE','Evenings ED','13:00','21:30',8.5,8,'Evening','WEDE','Yes','Yes',15],
  ['WMC','Weekend Midday Central Pharmacy','10:00','18:30',8.5,8,'Evening','WMC','Yes','Yes',20],
  ['N1','Nights','20:30','07:00',10.5,10,'Night','N1','Yes','Yes',5],
  ['N2','Nights','21:00','07:30',10.5,10,'Night','N2','Yes','Yes',5],
  ['SUP','Supervisor','','',0,0,'Other','SUP','No','No',50],
  ['RES','Resident','','',0,0,'Training','RES','No','No',50],
  ['ED11','ED Day/Evening Overlap','11:00','19:00',8,8,'Day','ED11','No','No',30],
  ['TDP','Training / TDP','','',0,0,'Training','TDP','No','No',50]
].map(r=>({Shift:r[0],Meaning:r[1],Start:r[2],End:r[3],Hours:r[4],'Credited Hours':r[5],Type:r[6],Skill:r[7],Weekend:r[8],Active:r[9],Priority:r[10],'Updated At':'','Updated By':'SYSTEM'}));

const STAFFING = {
  C7:[0,1,1,1,1,1,0], C8:[0,1,1,1,1,1,0], CARD:[0,1,1,1,1,1,0], CC1:[0,1,1,1,1,1,0], CC2:[0,1,1,1,1,1,0],
  EDD:[0,1,1,1,1,1,0], IM:[0,1,1,1,1,1,0], ONC:[0,1,1,1,1,1,0], E1:[0,1,1,1,1,1,0], E2:[0,1,1,1,1,1,0], EDE:[0,1,1,1,1,1,0],
  E:[1,1,1,1,1,1,1], WC7:[1,0,0,0,0,0,1], WD1:[1,0,0,0,0,0,1], WD2:[1,0,0,0,0,0,1], WE1:[0,0,0,0,0,0,0],
  WEDE:[1,0,0,0,0,0,1], WMC:[1,0,0,0,0,0,1], N1:[1,1,1,1,1,1,1], N2:[1,1,1,1,1,1,1], SUP:[0,0,0,0,0,0,0], RES:[0,0,0,0,0,0,0], ED11:[0,0,0,0,0,0,0], TDP:[0,0,0,0,0,0,0]
};
export const DEFAULT_STAFFING = Object.entries(STAFFING).map(([Shift,v])=>({Shift,Sunday:v[0],Monday:v[1],Tuesday:v[2],Wednesday:v[3],Thursday:v[4],Friday:v[5],Saturday:v[6],Active:'Yes','Updated At':'','Updated By':'SYSTEM'}));

export function defaultUsers(){
  return PHARMACISTS.map((p,i)=>({
    'Employee ID':`EMP-${String(i+1).padStart(3,'0')}`,'Pharmacist Name':p[0],Username:`EMPLOYEE::EMP-${String(i+1).padStart(3,'0')}`,
    Active:'Yes',Role:'Pharmacist','Schedule Type':p[1],Preceptor:p[5]?'Yes':'No',Resident:p[6]?'Yes':'No','Weekend Group':p[2],
    'Weekly Hour Maximum':p[3]||40,'Target Weekly Hours':p[3]||40,'Maximum Evening Shifts Per Month':7,'Preferred Start Time':'','Preferred End Time':'',
    'Custom Hours Enabled':'No','Custom Hours Mode':'SOFT','Preferred Shift Type':p[4]==='TBA'?'':p[4],'Off-Day Coverage Skills':'',
    'Weekend Eligible':'Yes','Evening Eligible':'Yes','Night Eligible':'Yes','Resident Covers Regular':'Yes','Resident Weekends':'Yes','Resident Evenings':'Yes','Resident Nights':'No',
    'Rotation Anchor Date':'','Weekend Rotation Anchor Date':'','SevenOn Weekly Handling':'ENFORCE_MAX',Notes:'','Updated At':'','Updated By':'SYSTEM'
  }));
}
export function defaultSkills(users){
  const valid=new Set(DEFAULT_SHIFTS.map(s=>s.Skill).filter(Boolean));
  const rows=[];
  users.forEach(u=>{const pref=String(u['Preferred Shift Type']||'').toUpperCase(); if(pref&&valid.has(pref))rows.push({'Employee ID':u['Employee ID'],'Pharmacist Name':u['Pharmacist Name'],Username:u.Username,Skill:pref,Active:'Yes','Updated At':'','Updated By':'SYSTEM'});});
  return rows;
}

export function emptyDatabase(){
  const users=defaultUsers(), now=nowIso();
  users.forEach(x=>x['Updated At']=now); DEFAULT_SHIFTS.forEach(x=>x['Updated At']=now); DEFAULT_STAFFING.forEach(x=>x['Updated At']=now); DEFAULT_SETTINGS.forEach(x=>{x['Updated At']=now;x['Updated By']='SYSTEM';});
  return {
    meta:{app:'NeoChrono',schemaVersion:1,createdAt:now,updatedAt:now},admins:[],users,skills:defaultSkills(users),shifts:DEFAULT_SHIFTS,staffing:DEFAULT_STAFFING,
    requests:[],weeklyAvailability:[],preceptorCalendar:[],schedule:[],settings:DEFAULT_SETTINGS,audit:[],swaps:[]
  };
}
