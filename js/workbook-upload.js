import { applyNeoChronoWorkbookImport } from './workbook-import.js';

const ALIASES={
  users:['Users','Pharmacists','Employees'],
  skills:['Employee Skills','Skills'],
  shifts:['Shifts','Shift Definitions'],
  staffingRequirements:['Staffing Requirements','Staffing','Requirements'],
  requests:['PTO / Availability Requests','PTO Availability Requests','PTO','Requests'],
  weeklyAvailability:['Employee Weekly Availability','Weekly Availability'],
  preceptorCalendar:['Preceptor Calendar','Preceptor Schedule'],
  settingsRows:['Settings'],
  schedule:['Schedule'],
  audit:['Audit Log','Audit'],
  swaps:['Swap Market','Shift Swaps','Swaps']
};
const DEC=new TextDecoder();
const u16=(d,o)=>d.getUint16(o,true),u32=(d,o)=>d.getUint32(o,true);

async function unzip(buffer){
  const bytes=new Uint8Array(buffer),dv=new DataView(buffer);
  let e=-1;
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--){
    if(u32(dv,i)===0x06054b50){e=i;break;}
  }
  if(e<0)throw new Error('The selected file is not a valid .xlsx workbook.');
  const count=u16(dv,e+10),start=u32(dv,e+16),files={};let p=start;
  for(let i=0;i<count;i++){
    if(u32(dv,p)!==0x02014b50)throw new Error('Invalid Excel ZIP directory.');
    const method=u16(dv,p+10),size=u32(dv,p+20),nl=u16(dv,p+28),xl=u16(dv,p+30),cl=u16(dv,p+32),local=u32(dv,p+42);
    const name=DEC.decode(bytes.slice(p+46,p+46+nl));
    files[name]={method,size,local};
    p+=46+nl+xl+cl;
  }
  async function readBytes(name){
    const f=files[name];if(!f)return null;
    const o=f.local;
    if(u32(dv,o)!==0x04034b50)throw new Error('Invalid Excel ZIP entry.');
    const nl=u16(dv,o+26),xl=u16(dv,o+28),s=o+30+nl+xl,part=bytes.slice(s,s+f.size);
    if(f.method===0)return part;
    if(f.method!==8)throw new Error('Unsupported Excel compression method.');
    if(typeof DecompressionStream==='undefined')throw new Error('This browser cannot open .xlsx files. Use the latest Chrome browser.');
    const stream=new Blob([part]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  return {text:async name=>{const b=await readBytes(name);return b?DEC.decode(b):'';}};
}
const xml=s=>new DOMParser().parseFromString(s,'application/xml');
const nodeText=n=>n?n.textContent||'':'';
function col(ref){
  const m=String(ref||'').match(/^([A-Z]+)/i);if(!m)return 0;
  let x=0;for(const c of m[1].toUpperCase())x=x*26+c.charCodeAt(0)-64;return x-1;
}
function excelDate(v){
  const d=new Date((Number(v)-25569)*86400000);
  return Number.isNaN(d.getTime())?v:d;
}
function toDateString(v){
  const d=typeof v==='number'?excelDate(v):v instanceof Date?v:new Date(v);
  if(!(d instanceof Date)||Number.isNaN(d.getTime()))return v;
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`;
}
function toIso(v){
  const d=typeof v==='number'?excelDate(v):v instanceof Date?v:new Date(v);
  return d instanceof Date&&!Number.isNaN(d.getTime())?d.toISOString():v;
}
function toTime(v){
  if(typeof v!=='number')return v;
  const mins=Math.round((((v%1)+1)%1)*1440)%1440;
  return `${String(Math.floor(mins/60)).padStart(2,'0')}:${String(mins%60).padStart(2,'0')}`;
}
function normalizeRows(sheetName,rows){
  const dateCols=new Set(['date','start date','end date','weekend saturday','weekend sunday','rotation anchor date','weekend rotation anchor date','effective start','effective end']);
  const isoCols=new Set(['submitted at','reviewed at','updated at','finalized at','timestamp','last login']);
  const timeCols=new Set(['start','end','preferred start time','preferred end time','start time','end time']);
  return rows.map(r=>{
    const out={...r};
    for(const k of Object.keys(out)){
      const lk=k.trim().toLowerCase();
      if(dateCols.has(lk))out[k]=toDateString(out[k]);
      else if(isoCols.has(lk))out[k]=toIso(out[k]);
      else if(timeCols.has(lk))out[k]=toTime(out[k]);
    }
    if(sheetName.toLowerCase()==='settings'&&/date/i.test(String(out.Setting||'')))out.Value=toDateString(out.Value);
    return out;
  });
}
async function readWorkbook(buffer){
  const z=await unzip(buffer),wb=xml(await z.text('xl/workbook.xml')),relsDoc=xml(await z.text('xl/_rels/workbook.xml.rels'));
  const rels={};for(const r of relsDoc.getElementsByTagName('Relationship'))rels[r.getAttribute('Id')]=r.getAttribute('Target')||'';
  const shared=[];const ss=await z.text('xl/sharedStrings.xml');
  if(ss){const d=xml(ss);for(const si of d.getElementsByTagName('si'))shared.push([...si.getElementsByTagName('t')].map(nodeText).join(''));}
  const SheetNames=[],Sheets={};
  for(const s of wb.getElementsByTagName('sheet')){
    const name=s.getAttribute('name')||'',rid=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    let target=rels[rid]||'';if(target.startsWith('/'))target=target.slice(1);else if(!target.startsWith('xl/'))target='xl/'+target.replace(/^\.\//,'');
    const sx=await z.text(target);if(!sx)continue;
    const d=xml(sx),matrix=[];
    for(const row of d.getElementsByTagName('row')){
      const a=[];
      for(const c of row.getElementsByTagName('c')){
        const i=col(c.getAttribute('r')),t=c.getAttribute('t')||'',v=nodeText(c.getElementsByTagName('v')[0]);let val='';
        if(t==='s')val=shared[Number(v)]??'';
        else if(t==='inlineStr')val=[...c.getElementsByTagName('t')].map(nodeText).join('');
        else if(t==='b')val=v==='1';
        else if(v!=='')val=Number.isFinite(Number(v))?Number(v):v;
        a[i]=val;
      }
      matrix.push(a);
    }
    const headers=(matrix[0]||[]).map(v=>String(v??'').trim()),out=[];
    for(const a of matrix.slice(1)){
      if(!a.some(v=>String(v??'').trim()!==''))continue;
      const r={};headers.forEach((h,i)=>{if(h)r[h]=a[i]??'';});out.push(r);
    }
    SheetNames.push(name);Sheets[name]=normalizeRows(name,out);
  }
  return {SheetNames,Sheets};
}
function findName(wb,names){
  const m=Object.fromEntries(wb.SheetNames.map(n=>[n.trim().toLowerCase(),n]));
  for(const n of names)if(m[n.toLowerCase()])return m[n.toLowerCase()];
  return '';
}
export async function parseWorkbookUpload(file){
  if(!/\.xlsx$/i.test(file.name))throw new Error('Please upload the Google Sheet exported as a Microsoft Excel (.xlsx) file.');
  const wb=await readWorkbook(await file.arrayBuffer()),tables={},recognized=[],seen=new Set();
  for(const [key,names] of Object.entries(ALIASES)){
    const name=findName(wb,names);tables[key]=name?(wb.Sheets[name]||[]):[];
    if(name){recognized.push({key,sheet:name,rows:tables[key].length});seen.add(name);}
  }
  return {fileName:file.name,tables,recognized,unrecognized:wb.SheetNames.filter(n=>!seen.has(n))};
}
export { applyNeoChronoWorkbookImport };
