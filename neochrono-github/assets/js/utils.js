export const clean = v => String(v ?? '').trim();
export const yes = v => /^(yes|true|1|y)$/i.test(clean(v));
export const num = (v, fallback = 0) => Number.isFinite(Number(v)) ? Number(v) : fallback;
export const clamp = (v, min, max) => Math.max(min, Math.min(max, v));
export const uuid = () => crypto.randomUUID ? crypto.randomUUID() : 'id-'+Date.now()+'-'+Math.random().toString(16).slice(2);
export const nowIso = () => new Date().toISOString();
export const esc = v => clean(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
export const attr = esc;

export function parseDate(value){
  if(value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 12);
  const s=clean(value);
  if(!s) return null;
  const m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12);
  const d=new Date(s); return Number.isNaN(d.getTime())?null:new Date(d.getFullYear(),d.getMonth(),d.getDate(),12);
}
export function dateKey(v){const d=parseDate(v); if(!d)return ''; return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
export function addDays(v,n){const d=parseDate(v); d.setDate(d.getDate()+Number(n||0)); return d;}
export function daysBetween(a,b){return Math.round((parseDate(b)-parseDate(a))/86400000);}
export function dayName(v){return ['SUN','MON','TUE','WED','THU','FRI','SAT'][parseDate(v).getDay()];}
export function isWeekend(v){const d=parseDate(v).getDay(); return d===0||d===6;}
export function weekStartSunday(v){const d=parseDate(v); d.setDate(d.getDate()-d.getDay()); return d;}
export function monthKey(v){const d=parseDate(v); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;}
export function eachDate(start,end){const out=[]; for(let d=parseDate(start),e=parseDate(end);d<=e;d=addDays(d,1))out.push(new Date(d)); return out;}
export function rangesOverlap(aStart,aEnd,bStart,bEnd){return parseDate(aStart)<=parseDate(bEnd)&&parseDate(bStart)<=parseDate(aEnd);}
export function b64EncodeUtf8(text){const bytes=new TextEncoder().encode(text);let bin='';bytes.forEach(b=>bin+=String.fromCharCode(b));return btoa(bin);}
export function b64DecodeUtf8(text){const bin=atob(text.replace(/\n/g,''));const bytes=Uint8Array.from(bin,c=>c.charCodeAt(0));return new TextDecoder().decode(bytes);}
export async function sha256(text){const buf=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(text));return [...new Uint8Array(buf)].map(b=>b.toString(16).padStart(2,'0')).join('');}
export async function passwordHash(password,salt){return sha256(`${salt}::${password}`);}
export function downloadText(filename,text,type='text/plain'){const blob=new Blob([text],{type});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=filename;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},0);}
