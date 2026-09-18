import { APP_NAME, APP_VERSION, DATA_FILES } from './config.js';
import { emptyDatabase } from './seed.js';
import { clean, nowIso, passwordHash, sha256, uuid, yes } from './utils.js';
import { employeeStats, reconcilePtoAutoApprovals, scheduleHealth, settingsMap, validateConfiguration } from './rules.js';
import { validateAssignments } from './scheduler.js';

export class DataService {
  constructor(store){this.store=store;this.currentUser=null;}
  async needsSetup(){return !(await this.store.getFile(DATA_FILES.meta));}
  async initialize({username,password,adminName='Schedule Administrator'}){
    if(await this.store.getFile(DATA_FILES.meta)) throw new Error('This data repository is already initialized.');
    if(clean(password).length<10) throw new Error('Initial administrator password must be at least 10 characters.');
    const db=emptyDatabase(),salt=uuid(),hash=await passwordHash(password,salt),now=nowIso();
    db.admins=[{'Admin ID':'ADMIN-001','Admin Name':adminName,Username:clean(username).toLowerCase(),'Password Hash':hash,'Password Salt':salt,Active:'Yes','Must Change Password':'No','Updated At':now,'Updated By':'SYSTEM'}];
    db.audit.push({Timestamp:now,User:clean(username).toLowerCase(),Action:'DATABASE_INITIALIZED',Date:'',Pharmacist:'','Old Shift':'','New Shift':'','Old Value':'','New Value':'','Override':'No','Override Reason':'',Details:'NeoChrono GitHub database initialized.'});
    for(const [key,file] of Object.entries(DATA_FILES)){const value=db[key];if(value===undefined)continue;await this.store.putJSON(file,value,{message:`Initialize NeoChrono ${file}`});}
    return {ok:true};
  }
  async login(username,password){
    const admins=await this.store.getJSON(DATA_FILES.admins,[]);const u=clean(username).toLowerCase(),admin=admins.find(a=>clean(a.Username).toLowerCase()===u&&yes(a.Active??'Yes'));if(!admin)throw new Error('Invalid administrator username or password.');
    const hash=await passwordHash(String(password),clean(admin['Password Salt']));if(hash!==clean(admin['Password Hash']))throw new Error('Invalid administrator username or password.');this.currentUser={Username:admin.Username,'Pharmacist Name':admin['Admin Name'],Role:'Administrator','Admin ID':admin['Admin ID']};await this.audit('LOGIN',{Details:'Administrator logged in.'});return {ok:true,user:this.currentUser,version:APP_VERSION,mustChangePassword:yes(admin['Must Change Password'])};
  }
  async loadAll(){
    const keys=['admins','users','skills','shifts','staffing','requests','weeklyAvailability','preceptorCalendar','schedule','settings','audit','swaps'];const values=await Promise.all(keys.map(k=>this.store.getJSON(DATA_FILES[k],[])));const d=Object.fromEntries(keys.map((k,i)=>[k,values[i]||[]]));
    const limit=Number(settingsMap(d.settings)['PTO Auto Approval Limit']||2);const reconciled=reconcilePtoAutoApprovals(d.requests,limit);if(JSON.stringify(reconciled)!==JSON.stringify(d.requests)){d.requests=reconciled;await this.store.putJSON(DATA_FILES.requests,reconciled,{message:'Reconcile PTO auto approvals'});}
    const repo=await this.store.repoInfo();
    return {appName:APP_NAME,version:APP_VERSION,database:{id:`${this.store.owner}/${this.store.repo}`,name:this.store.repo,url:repo.html_url,branch:this.store.branch,dataDir:this.store.dataDir},user:this.currentUser||{Username:'',Role:'Administrator','Pharmacist Name':'Administrator'},isAdmin:true,admins:d.admins.map(a=>({'Admin ID':a['Admin ID'],'Admin Name':a['Admin Name'],Username:a.Username,Active:a.Active,'Must Change Password':a['Must Change Password'],'Updated At':a['Updated At'],'Updated By':a['Updated By']})),users:d.users,skills:d.skills,shifts:d.shifts,requirements:d.staffing,requests:d.requests,weeklyAvailability:d.weeklyAvailability,preceptorCalendar:d.preceptorCalendar,schedule:d.schedule,settings:settingsMap(d.settings),settingsRows:d.settings,audit:d.audit,swaps:d.swaps,health:scheduleHealth(d.schedule),stats:employeeStats(d.schedule,d.users),configValidation:validateConfiguration(d)};
  }
  async getRaw(){const out={};for(const [key,file] of Object.entries(DATA_FILES))out[key]=await this.store.getJSON(file,key==='meta'?{}:[]);return out;}
  async save(key,value,message){if(!DATA_FILES[key])throw new Error(`Unknown data file ${key}.`);await this.store.mutateJSON(DATA_FILES[key],Array.isArray(value)?[]:{},()=>value,message||`Update ${key}`);return value;}
  async mutate(key,fallback,fn,message){return this.store.mutateJSON(DATA_FILES[key],fallback,fn,message);}
  async audit(action,details={}){const now=nowIso(),user=clean(this.currentUser?.Username)||'SYSTEM';const row={Timestamp:now,User:user,Action:action,Date:details.Date||'',Pharmacist:details.Pharmacist||'','Old Shift':details['Old Shift']||'','New Shift':details['New Shift']||'','Old Value':details['Old Value']||'','New Value':details['New Value']||'',Override:details.Override||'No','Override Reason':details['Override Reason']||'',Details:details.Details||''};await this.mutate('audit',[],rows=>{rows.push(row);return rows.slice(-5000);},`${action} by ${user}`);return row;}
  async changePassword(currentPassword,newPassword){if(!this.currentUser)throw new Error('Sign in first.');if(String(newPassword).length<10)throw new Error('New password must be at least 10 characters.');const user=this.currentUser;let changed=false;await this.mutate('admins',[],async rows=>{const a=rows.find(x=>clean(x.Username)===clean(user.Username));if(!a)throw new Error('Administrator account not found.');if(await passwordHash(currentPassword,clean(a['Password Salt']))!==clean(a['Password Hash']))throw new Error('Current password is incorrect.');const salt=uuid();a['Password Salt']=salt;a['Password Hash']=await passwordHash(newPassword,salt);a['Must Change Password']='No';a['Updated At']=nowIso();a['Updated By']=user.Username;changed=true;return rows;},`Change password for ${user.Username}`);if(changed)await this.audit('PASSWORD_CHANGED',{Details:'Administrator password changed.'});return {ok:true};}
  async validateSaved(start,end){const d=await this.getRaw();return validateAssignments(d.schedule,d,start,end);}
}
