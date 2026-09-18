(function(){
  'use strict';

  const $=id=>document.getElementById(id);
  const token=sessionStorage.getItem('rxSchedulerToken')||'';
  let previewData=null;

  const PROTECTED=new Set(['Admin Accounts','Sessions']);
  const NAME_MAP={
    'PTO  Availability Requests':'PTO / Availability Requests',
    'PTO Availability Requests':'PTO / Availability Requests'
  };

  function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
  function clean(v){return String(v??'').trim();}
  function yes(v){return /^(yes|true|1|y|on)$/i.test(clean(v));}
  function dateString(v){
    if(v===''||v==null)return '';
    const d=v instanceof Date?v:new Date(v);
    if(Number.isNaN(d.getTime()))return v;
    const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),day=String(d.getDate()).padStart(2,'0');
    return y+'-'+m+'-'+day;
  }
  function timeString(v){
    if(v===''||v==null)return '';
    if(v instanceof Date)return String(v.getHours()).padStart(2,'0')+':'+String(v.getMinutes()).padStart(2,'0');
    if(typeof v==='number'){
      const mins=Math.round((((v%1)+1)%1)*1440)%1440;
      return String(Math.floor(mins/60)).padStart(2,'0')+':'+String(mins%60).padStart(2,'0');
    }
    const s=clean(v);
    const m=s.match(/^(\d{1,2}):(\d{2})(?:\s*(AM|PM))?$/i);
    if(m){
      let h=Number(m[1]),min=m[2];const ap=(m[3]||'').toUpperCase();
      if(ap==='PM'&&h<12)h+=12;if(ap==='AM'&&h===12)h=0;
      return String(h).padStart(2,'0')+':'+min;
    }
    return s;
  }
  function isoString(v){
    if(v===''||v==null)return '';
    const d=v instanceof Date?v:new Date(v);
    return Number.isNaN(d.getTime())?String(v):d.toISOString();
  }

  const DATE_HEADERS=new Set(['Date','Start Date','End Date','Weekend Saturday','Weekend Sunday','Rotation Anchor Date','Weekend Rotation Anchor Date','Effective Start','Effective End']);
  const TIME_HEADERS=new Set(['Start','End','Preferred Start Time','Preferred End Time','Start Time','End Time']);
  const DATETIME_HEADERS=new Set(['Updated At','Submitted At','Reviewed At','Finalized At','Timestamp','Last Login','Created At','Expires At','Last Seen At']);

  function normalizeMatrix(sheetName,matrix){
    if(!matrix.length)return [];
    const headers=(matrix[0]||[]).map(clean);
    const out=[headers];
    for(const raw of matrix.slice(1)){
      if(!(raw||[]).some(v=>clean(v)!==''))continue;
      const row=headers.map((h,i)=>{
        const v=raw?.[i]??'';
        if(DATE_HEADERS.has(h))return dateString(v);
        if(TIME_HEADERS.has(h))return timeString(v);
        if(DATETIME_HEADERS.has(h))return isoString(v);
        if(sheetName==='Settings'&&h==='Value'){
          const setting=clean(raw?.[0]);
          if(/Date/i.test(setting))return dateString(v);
        }
        return v instanceof Date?isoString(v):v;
      });
      out.push(row);
    }
    return out;
  }

  function ensureSkillPreferred(importedSheets){
    const key='Employee Skills';
    const users=importedSheets['Users']?.values||[];
    const skills=importedSheets[key]?.values||[];
    if(!skills.length)return;

    const sh=skills;
    let headers=(sh[0]||[]).map(clean);
    let prefIdx=headers.indexOf('Preferred');
    if(prefIdx<0){
      headers=[...headers,'Preferred'];
      sh[0]=headers;
      prefIdx=headers.length-1;
      for(let i=1;i<sh.length;i++)sh[i][prefIdx]='';
    }
    const uh=(users[0]||[]).map(clean);
    const uEmp=uh.indexOf('Employee ID'),uUser=uh.indexOf('Username'),uPref=uh.indexOf('Preferred Shift Type');
    const shEmp=headers.indexOf('Employee ID'),shUser=headers.indexOf('Username'),shSkill=headers.indexOf('Skill'),shActive=headers.indexOf('Active');

    const desired=new Map();
    for(const r of users.slice(1)){
      const pref=clean(r[uPref]).toUpperCase(); if(!pref)continue;
      const k=clean(r[uEmp])||clean(r[uUser]);
      if(k)desired.set(k,pref);
    }
    const seenPreferred=new Set();
    for(let i=1;i<sh.length;i++){
      const r=sh[i],key=clean(r[shEmp])||clean(r[shUser]),code=clean(r[shSkill]).toUpperCase();
      if(!key||!code||!yes(r[shActive]??'Yes')){r[prefIdx]='No';continue;}
      const want=desired.get(key);
      const mark=!!want&&want===code&&!seenPreferred.has(key);
      r[prefIdx]=mark?'Yes':'No';
      if(mark)seenPreferred.add(key);
    }
  }

  function readXlsx(file){
    return file.arrayBuffer().then(buf=>{
      const wb=XLSX.read(buf,{type:'array',cellDates:true});
      const sheets={};const info=[];
      for(const sourceName of wb.SheetNames){
        const target=NAME_MAP[sourceName]||sourceName;
        if(PROTECTED.has(target)){info.push({source:sourceName,target,status:'Preserved current GitHub account/session data',rows:0,protected:true});continue;}
        const matrix=XLSX.utils.sheet_to_json(wb.Sheets[sourceName],{header:1,defval:'',raw:true,blankrows:false});
        const norm=normalizeMatrix(target,matrix);
        sheets[target]={values:norm};
        info.push({source:sourceName,target,status:'Ready',rows:Math.max(0,norm.length-1),protected:false});
      }
      ensureSkillPreferred(sheets);
      return {fileName:file.name,sheets,info};
    });
  }

  function mergeSheet(current,incoming){
    if(!current?.values?.length)return incoming;
    if(!incoming?.values?.length)return current;
    const h1=(current.values[0]||[]).map(clean),h2=(incoming.values[0]||[]).map(clean);
    const headers=[...h1];for(const h of h2)if(h&&!headers.includes(h))headers.push(h);
    const toObj=(row,heads)=>{const o={};heads.forEach((h,i)=>{if(h)o[h]=row?.[i]??'';});return o;};
    const rows=[];
    const seen=new Set();
    for(const [mat,heads] of [[current.values,h1],[incoming.values,h2]]){
      for(const row of mat.slice(1)){
        const o=toObj(row,heads);
        const k=JSON.stringify(o);
        if(seen.has(k))continue;seen.add(k);
        rows.push(headers.map(h=>o[h]??''));
      }
    }
    return {values:[headers,...rows]};
  }

  function countsHtml(info){
    return info.map(x=>'<tr><td><b>'+esc(x.source)+'</b></td><td>'+esc(x.target)+'</td><td>'+x.rows+'</td><td>'+esc(x.status)+'</td></tr>').join('');
  }

  async function guard(){
    if(!token){$('guard').className='alert danger';$('guard').innerHTML='Administrator session not found. <a href="index.html">Sign in first</a>.';return;}
    try{
      const data=await window.__neoRuntime.invoke('getAppData',[token]);
      if(!data||!data.isAdmin)throw new Error('Administrator access is required.');
      $('guard').className='alert';
      $('guard').textContent='Signed in as '+(data.user?.['Admin Name']||data.user?.Username||'Administrator')+'.';
      $('main').classList.remove('hidden');
    }catch(e){
      $('guard').className='alert danger';$('guard').innerHTML=esc(e.message)+' <a href="index.html">Return to login</a>.';
    }
  }

  $('preview').onclick=async()=>{
    const file=$('file').files?.[0];
    if(!file){$('result').innerHTML='<div class="alert danger">Choose the .xlsx workbook first.</div>';return;}
    $('preview').disabled=true;$('apply').disabled=true;$('result').innerHTML='<div class="alert warn">Reading workbook…</div>';
    try{
      if(typeof XLSX==='undefined')throw new Error('Excel reader did not load. Check your internet connection and refresh.');
      previewData=await readXlsx(file);
      const rows=previewData.info.filter(x=>!x.protected).reduce((n,x)=>n+x.rows,0);
      $('result').innerHTML=
        '<div class="kpis"><div class="kpi"><b>'+previewData.info.filter(x=>!x.protected).length+'</b><span>Sheets ready</span></div><div class="kpi"><b>'+rows+'</b><span>Rows ready</span></div><div class="kpi"><b>'+previewData.info.filter(x=>x.protected).length+'</b><span>Protected sheets</span></div></div>'+
        '<div class="card"><h3>Import preview</h3><table><thead><tr><th>Workbook tab</th><th>GitHub sheet</th><th>Rows</th><th>Action</th></tr></thead><tbody>'+countsHtml(previewData.info)+'</tbody></table></div>'+
        '<div class="alert warn"><b>Protected:</b> Admin Accounts and Sessions are not overwritten, so importing the workbook cannot replace your current NeoChrono administrator login.</div>';
      $('apply').disabled=false;
    }catch(e){previewData=null;$('result').innerHTML='<div class="alert danger">'+esc(e.message)+'</div>';}
    finally{$('preview').disabled=false;}
  };

  $('apply').onclick=async()=>{
    if(!previewData)return;
    if(!confirm('Import all scheduler data from '+previewData.fileName+' into the private GitHub database?'))return;
    $('apply').disabled=true;$('result').insertAdjacentHTML('afterbegin','<div class="alert warn" id="saving">Saving workbook to GitHub…</div>');
    try{
      const loaded=await window.__neoRuntime.loadWorkbook(true);
      const next=structuredClone(loaded.data);
      next.meta=next.meta||{};
      next.meta.lastExcelImportAt=new Date().toISOString();
      next.meta.lastExcelImportFile=previewData.fileName;
      next.meta.source='Uploaded NeoChrono workbook';
      next.sheets=next.sheets||{};

      for(const [name,sheet] of Object.entries(previewData.sheets)){
        if(PROTECTED.has(name))continue;
        if($('mode').value==='merge')next.sheets[name]=mergeSheet(next.sheets[name],sheet);
        else next.sheets[name]=sheet;
      }

      await window.__neoRuntime.saveWorkbook(next,loaded.sha,'Import complete NeoChrono workbook: '+previewData.fileName);
      await window.__neoRuntime.invoke('ensureSkillPreferenceSystem',[token]);
      let validation=null;
      try{validation=await window.__neoRuntime.invoke('validateConfiguration',[token]);}catch(e){validation={errors:[e.message],warnings:[]};}

      const errs=validation?.errors||[],warns=validation?.warnings||[];
      $('result').innerHTML=
        '<div class="alert"><b>Import complete.</b> '+esc(previewData.fileName)+' is now stored in the private GitHub database.</div>'+
        (errs.length?'<div class="alert danger"><b>Configuration errors:</b><br>'+errs.map(esc).join('<br>')+'</div>':'<div class="alert"><b>Configuration check:</b> No blocking setup errors were found.</div>')+
        (warns.length?'<div class="alert warn"><b>Warnings:</b><br>'+warns.map(esc).join('<br>')+'</div>':'')+
        '<div class="card"><h3>Imported workbook</h3><table><thead><tr><th>Workbook tab</th><th>GitHub sheet</th><th>Rows</th><th>Status</th></tr></thead><tbody>'+countsHtml(previewData.info)+'</tbody></table></div>';
      previewData=null;
    }catch(e){
      $('result').insertAdjacentHTML('afterbegin','<div class="alert danger">'+esc(e.message)+'</div>');
      $('apply').disabled=false;
    }
  };

  guard();
})();