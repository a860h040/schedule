/* NeoChrono GitHub compatibility runtime
 * Runs the uploaded Apps Script scheduler in the browser against a private
 * GitHub JSON workbook. The scheduling algorithms remain in the uploaded
 * legacy files; this file only supplies Apps Script/Spreadsheet shims and
 * persistence.
 */
(function(){
  'use strict';

  const CONFIG_KEY='neochronoGithubConfigV1';
  const WORKBOOK_PATH='data/workbook.json';
  const PUBLISHED_PATH='data/published-schedule.json';
  const SESSION_SHEET_KEY='neochronoLegacySessionsV1';
  let cache={data:null,sha:null,loadedAt:0};
  let photoCache=new Map();
  let invokeQueue=Promise.resolve();

  const enc=s=>new TextEncoder().encode(String(s));
  function b64encodeUtf8(s){let b='';for(const x of enc(s))b+=String.fromCharCode(x);return btoa(b);}
  function b64decodeUtf8(s){const b=atob(String(s||'').replace(/\n/g,''));return new TextDecoder().decode(Uint8Array.from(b,c=>c.charCodeAt(0)));}
  function cfg(){
    try{
      const c=JSON.parse(localStorage.getItem(CONFIG_KEY)||'null');
      if(!c||!c.owner||!c.repo||!c.token)throw new Error('This browser is not connected to the NeoChrono private data repository. Open Setup This Device first.');
      return {...c,branch:c.branch||'main',workbookPath:WORKBOOK_PATH,publishedPath:PUBLISHED_PATH};
    }catch(e){if(/not connected/.test(String(e.message)))throw e;throw new Error('This browser is not connected to NeoChrono. Open Setup This Device first.');}
  }
  function apiUrl(c,path){return 'https://api.github.com/repos/'+encodeURIComponent(c.owner)+'/'+encodeURIComponent(c.repo)+'/contents/'+path.split('/').map(encodeURIComponent).join('/');}
  function headers(c){return {'Accept':'application/vnd.github+json','Authorization':'Bearer '+c.token,'X-GitHub-Api-Version':'2022-11-28','Content-Type':'application/json'};}
  async function gh(url,opt,c){
    let res;
    const controller=new AbortController();
    const timer=setTimeout(()=>controller.abort(),20000);
    try{
      res=await fetch(url,{
        ...(opt||{}),
        cache:'no-store',
        signal:controller.signal,
        headers:{...headers(c),...((opt||{}).headers||{})}
      });
    }catch(fetchErr){
      if(fetchErr&&fetchErr.name==='AbortError'){
        throw new Error('GitHub took too long to respond. Please try again.');
      }
      throw new Error('Could not connect to GitHub. Check your internet connection and the saved GitHub token, then try Setup This Device again.');
    }finally{
      clearTimeout(timer);
    }

    const txt=await res.text();
    let body=null;
    try{ body=txt?JSON.parse(txt):null; }catch{ body=txt; }

    if(!res.ok){
      let msg=body&&body.message?body.message:'GitHub request failed ('+res.status+')';
      if(res.status===401)msg='GitHub connection failed. The saved token is missing, expired, or invalid. Open Setup This Device and reconnect.';
      if(res.status===403)msg='GitHub denied access. Confirm the fine-grained token can read and write the neochrono-data repository.';
      const e=new Error(msg);
      e.status=res.status;
      e.body=body;
      throw e;
    }
    return body;
  }
  async function loadWorkbook(force=false){
    const c=cfg();
    if(cache.data&&!force&&Date.now()-cache.loadedAt<10000)return cache;

    const f=await gh(
      apiUrl(c,c.workbookPath)+'?ref='+encodeURIComponent(c.branch),
      {},
      c
    );

    let encoded=String((f&&f.content)||'').replace(/\n/g,'');

    // GitHub's Contents API omits inline "content" for larger files.
    // When workbook.json grows beyond that threshold, load the same blob
    // through the Git Data API instead of attempting JSON.parse('').
    if(!encoded){
      const sha=String((f&&f.sha)||'').trim();
      if(!sha){
        throw new Error(
          'NeoChrono could not read workbook.json from GitHub because the file response did not include content or a blob SHA.'
        );
      }

      const blobUrl=
        'https://api.github.com/repos/'+
        encodeURIComponent(c.owner)+'/'+
        encodeURIComponent(c.repo)+
        '/git/blobs/'+encodeURIComponent(sha);

      const blob=await gh(blobUrl,{},c);
      encoded=String((blob&&blob.content)||'').replace(/\n/g,'');

      if(!encoded){
        throw new Error(
          'NeoChrono found workbook.json in GitHub but GitHub returned an empty blob.'
        );
      }
    }

    let text='';
    let data=null;
    try{
      text=b64decodeUtf8(encoded);
      data=JSON.parse(text);
    }catch(e){
      throw new Error(
        'NeoChrono could not parse workbook.json from GitHub. '+
        'The stored workbook may be incomplete or invalid JSON. '+e.message
      );
    }

    cache={data,sha:f.sha,loadedAt:Date.now()};
    return cache;
  }
  async function saveWorkbook(data,sha,message){
    const c=cfg();
    const body={message:message||'NeoChrono data update',content:b64encodeUtf8(JSON.stringify(data,null,2)),branch:c.branch};
    if(sha)body.sha=sha;
    const r=await gh(apiUrl(c,c.workbookPath),{method:'PUT',body:JSON.stringify(body)},c);
    cache={data,sha:r.content&&r.content.sha?r.content.sha:sha,loadedAt:Date.now()};
    return r;
  }
  async function savePublished(payload){
    const c=cfg();let old=null;
    try{old=await gh(apiUrl(c,c.publishedPath)+'?ref='+encodeURIComponent(c.branch),{},c)}catch(e){if(e.status!==404)throw e}
    const body={message:'Publish pharmacist schedule',content:b64encodeUtf8(JSON.stringify(payload,null,2)),branch:c.branch};
    if(old&&old.sha)body.sha=old.sha;
    return gh(apiUrl(c,c.publishedPath),{method:'PUT',body:JSON.stringify(body)},c);
  }


  function employeePhotoSafeId_(employeeId){
    const id=String(employeeId||'').trim();
    if(!id)throw new Error('Employee ID is required for a pharmacist photo.');
    return id.replace(/[^A-Za-z0-9._-]+/g,'_');
  }

  function employeePhotoPath_(employeeId){
    return 'employee-photos/'+employeePhotoSafeId_(employeeId)+'.jpg';
  }

  function base64ToDataUrl_(base64,mime){
    return 'data:'+(mime||'image/jpeg')+';base64,'+String(base64||'').replace(/\s+/g,'');
  }

  async function getEmployeePhotoDataUrl(employeeId,force=false){
    const id=employeePhotoSafeId_(employeeId);
    if(!force&&photoCache.has(id))return photoCache.get(id);

    const c=cfg();
    const path=employeePhotoPath_(id);
    try{
      const file=await gh(apiUrl(c,path)+'?ref='+encodeURIComponent(c.branch),{},c);
      const url=base64ToDataUrl_(file.content||'','image/jpeg');
      photoCache.set(id,url);
      return url;
    }catch(e){
      if(e&&e.status===404){
        photoCache.set(id,'');
        return '';
      }
      throw e;
    }
  }

  async function uploadEmployeePhoto(employeeId,base64Jpeg){
    const id=employeePhotoSafeId_(employeeId);
    const content=String(base64Jpeg||'').replace(/^data:image\/[A-Za-z0-9.+-]+;base64,/,'').replace(/\s+/g,'');
    if(!content)throw new Error('The selected image could not be prepared for upload.');

    // Keep profile photos compact. Base64 is ~4/3 the binary size.
    if(content.length>1400000){
      throw new Error('The prepared pharmacist photo is too large. Please use a smaller image.');
    }

    const c=cfg();
    const path=employeePhotoPath_(id);
    let old=null;
    try{
      old=await gh(apiUrl(c,path)+'?ref='+encodeURIComponent(c.branch),{},c);
    }catch(e){
      if(!e||e.status!==404)throw e;
    }

    const body={
      message:'Update pharmacist photo: '+id,
      content:content,
      branch:c.branch
    };
    if(old&&old.sha)body.sha=old.sha;

    const result=await gh(
      apiUrl(c,path),
      {method:'PUT',body:JSON.stringify(body)},
      c
    );

    const dataUrl=base64ToDataUrl_(content,'image/jpeg');
    photoCache.set(id,dataUrl);
    return {
      ok:true,
      employeeId:id,
      path:path,
      dataUrl:dataUrl,
      sha:result&&result.content?result.content.sha:''
    };
  }

  async function deleteEmployeePhoto(employeeId){
    const id=employeePhotoSafeId_(employeeId);
    const c=cfg();
    const path=employeePhotoPath_(id);
    let old=null;

    try{
      old=await gh(apiUrl(c,path)+'?ref='+encodeURIComponent(c.branch),{},c);
    }catch(e){
      if(e&&e.status===404){
        photoCache.set(id,'');
        return {ok:true,removed:false,path:path};
      }
      throw e;
    }

    await gh(
      apiUrl(c,path),
      {
        method:'DELETE',
        body:JSON.stringify({
          message:'Remove pharmacist photo: '+id,
          sha:old.sha,
          branch:c.branch
        })
      },
      c
    );

    photoCache.set(id,'');
    return {ok:true,removed:true,path:path};
  }

  function clone(v){return v instanceof Date?new Date(v):Array.isArray(v)?v.map(clone):(v&&typeof v==='object'?Object.fromEntries(Object.entries(v).map(([k,x])=>[k,clone(x)])):v);}
  function isBlank(v){return v===null||v===undefined||v==='';}
  const DATE_HEADERS=new Set(['Date','Start Date','End Date','Rotation Anchor Date','Weekend Rotation Anchor Date','Weekend Saturday','Weekend Sunday','Effective Start','Effective End','Finalized At']);
  const DATETIME_HEADERS=new Set(['Updated At','Submitted At','Reviewed At','Timestamp','Created At','Expires At','Last Seen At','Last Login']);
  function parseDateMaybe(v){
    if(v instanceof Date)return new Date(v);
    if(typeof v!=='string')return v;
    if(!/^\d{4}-\d{2}-\d{2}(?:T| |$)/.test(v))return v;
    const d=new Date(v.length===10?v+'T12:00:00':v);
    return isNaN(d)?v:d;
  }
  function serializeCell(v){return v instanceof Date?v.toISOString():v;}

  class VRange{
    constructor(sheet,row,col,nr=1,nc=1){this.sheet=sheet;this.row=row;this.col=col;this.nr=nr||1;this.nc=nc||1;}
    getValues(){
      const out=[],headers=this.sheet.rawHeaders();
      for(let r=0;r<this.nr;r++){
        const row=[];
        for(let c=0;c<this.nc;c++){
          const rr=this.row-1+r,cc=this.col-1+c;
          let v=this.sheet.valueAt(rr,cc);
          const h=headers[cc]||'';
          if(DATE_HEADERS.has(h)||DATETIME_HEADERS.has(h))v=parseDateMaybe(v);
          if(this.sheet.name==='Settings'&&h==='Value'){
            const setting=String(this.sheet.valueAt(rr,0)||'');
            if(/Date/i.test(setting))v=parseDateMaybe(v);
          }
          row.push(clone(v));
        }
        out.push(row);
      }
      return out;
    }
    setValues(vals){
      for(let r=0;r<this.nr;r++)for(let c=0;c<this.nc;c++)this.sheet.setValueAt(this.row-1+r,this.col-1+c,vals?.[r]?.[c]??'');
      return this;
    }
    setValue(v){this.sheet.setValueAt(this.row-1,this.col-1,v);return this;}
    clearContent(){for(let r=0;r<this.nr;r++)for(let c=0;c<this.nc;c++)this.sheet.setValueAt(this.row-1+r,this.col-1+c,'');return this;}
    clearContents(){return this.clearContent();}
    setFontWeight(){return this} setWrap(){return this} setBackground(){return this} setFontColor(){return this}
    setNumberFormat(){return this} setDataValidation(){return this}
  }

  class VSheet{
    constructor(book,name,local=false){this.book=book;this.name=name;this.local=local;}
    matrix(){
      if(this.local){
        try{return JSON.parse(localStorage.getItem(SESSION_SHEET_KEY)||'null')||[['Token','Username','Created At','Expires At','Last Seen At']]}catch{return [['Token','Username','Created At','Expires At','Last Seen At']]}
      }
      const sh=this.book.data.sheets[this.name]||(this.book.data.sheets[this.name]={values:[]});
      if(!Array.isArray(sh.values))sh.values=[];
      return sh.values;
    }
    persistLocal(m){if(this.local)localStorage.setItem(SESSION_SHEET_KEY,JSON.stringify(m.map(r=>r.map(serializeCell))));}
    rawHeaders(){return (this.matrix()[0]||[]).map(x=>String(x??'').trim());}
    valueAt(r,c){const m=this.matrix();return m[r]?.[c]??'';}
    setValueAt(r,c,v){
      const m=this.matrix();
      while(m.length<=r)m.push([]);
      while(m[r].length<=c)m[r].push('');
      const old=m[r][c];
      const oldCmp=old instanceof Date?old.toISOString():JSON.stringify(old??'');
      const newCmp=v instanceof Date?v.toISOString():JSON.stringify(v??'');
      if(oldCmp===newCmp)return;
      m[r][c]=v;
      if(this.local)this.persistLocal(m);
      else this.book.dirty=true;
    }
    getLastRow(){
      const m=this.matrix();for(let i=m.length-1;i>=0;i--)if((m[i]||[]).some(v=>!isBlank(v)))return i+1;return 0;
    }
    getLastColumn(){
      const m=this.matrix();let max=0;for(const r of m)for(let i=(r||[]).length-1;i>=0;i--)if(!isBlank(r[i])){max=Math.max(max,i+1);break;}return max;
    }
    getDataRange(){return new VRange(this,1,1,Math.max(1,this.getLastRow()),Math.max(1,this.getLastColumn()));}
    getRange(row,col,nr=1,nc=1){return new VRange(this,row,col,nr,nc);}
    appendRow(row){const m=this.matrix();m.push((row||[]).slice());if(this.local)this.persistLocal(m);else this.book.dirty=true;return this;}
    deleteRow(n){const m=this.matrix();if(n>=1&&n<=m.length)m.splice(n-1,1);if(this.local)this.persistLocal(m);else this.book.dirty=true;}
    clearContents(){if(this.local){this.persistLocal([])}else{this.book.data.sheets[this.name].values=[];this.book.dirty=true}return this;}
    clear(){return this.clearContents();}
    setFrozenRows(){return this} setColumnWidth(){return this} autoResizeColumns(){return this} hideColumns(){return this}
    getMaxRows(){return Math.max(1000,this.matrix().length);}
    getSheetId(){let h=0;for(const ch of this.name)h=((h<<5)-h+ch.charCodeAt(0))|0;return Math.abs(h);}
    getName(){return this.name;}
    setName(newName){
      if(this.local){this.name=newName;return this;}
      const sh=this.book.data.sheets[this.name];delete this.book.data.sheets[this.name];this.book.data.sheets[newName]=sh;this.name=newName;this.book.dirty=true;return this;
    }
    copyTo(targetBook){const newName=this.name+' Copy';targetBook.data.sheets[newName]={values:clone(this.matrix())};targetBook.dirty=true;return new VSheet(targetBook,newName,false);}
  }

  class VBook{
    constructor(data){this.data=data;this.dirty=false;this.active=null;if(!this.data.sheets)this.data.sheets={};}
    getSheetByName(name){if(name==='Sessions')return new VSheet(this,'Sessions',true);return this.data.sheets[name]?new VSheet(this,name,false):null;}
    insertSheet(name){if(name==='Sessions')return new VSheet(this,'Sessions',true);if(!this.data.sheets[name]){this.data.sheets[name]={values:[]};this.dirty=true;}return new VSheet(this,name,false);}
    deleteSheet(sh){if(sh&&this.data.sheets[sh.name]){delete this.data.sheets[sh.name];this.dirty=true;}}
    setActiveSheet(sh){this.active=sh;return sh;}
    getId(){return this.data.meta?.id||'github:a860h040/neochrono-data';}
    getName(){return this.data.meta?.name||'NeoChrono GitHub Workbook';}
    getUrl(){return this.data.meta?.url||'https://github.com/a860h040/neochrono-data';}
    getSpreadsheetTimeZone(){return this.data.meta?.timezone||'America/New_York';}
  }

  function currentBook(){if(!window.__neoVBook)throw new Error('NeoChrono database is not loaded.');return window.__neoVBook;}

  window.SpreadsheetApp={
    getActiveSpreadsheet:()=>currentBook(),
    openById:()=>currentBook(),
    flush:()=>{},
    newDataValidation:()=>{const b={requireValueInList(){return b},setAllowInvalid(){return b},build(){return {}}};return b;}
  };
  window.PropertiesService={getScriptProperties:()=>({
    getProperty:k=>localStorage.getItem('neoProp:'+k)||'',
    setProperty:(k,v)=>localStorage.setItem('neoProp:'+k,String(v))
  })};
  window.LockService={getScriptLock:()=>({waitLock(){return true},tryLock(){return true},releaseLock(){}})};
  window.Session={getScriptTimeZone:()=>currentBook()?.getSpreadsheetTimeZone?.()||'America/New_York'};
  window.Logger={log:(...a)=>console.log(...a)};
  window.HtmlService={createTemplateFromFile:()=>({evaluate:()=>({setTitle(){return this},setXFrameOptionsMode(){return this},addMetaTag(){return this}})}),createHtmlOutputFromFile:()=>({getContent:()=>''}),XFrameOptionsMode:{ALLOWALL:'ALLOWALL'}};

  function uuid(){return crypto.randomUUID?crypto.randomUUID():('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g,c=>{const r=Math.random()*16|0,v=c==='x'?r:(r&3|8);return v.toString(16)}));}
  function sha256bytes(ascii){
    function rightRotate(v,a){return(v>>>a)|(v<<(32-a))}
    const maxWord=Math.pow(2,32),lengthProperty='length';let i,j,result='',words=[],asciiBitLength=ascii[lengthProperty]*8;
    let hash=sha256bytes.h=sha256bytes.h||[],k=sha256bytes.k=sha256bytes.k||[],primeCounter=k[lengthProperty],isComposite={};
    for(let candidate=2;primeCounter<64;candidate++){if(!isComposite[candidate]){for(i=0;i<313;i+=candidate)isComposite[i]=candidate;hash[primeCounter]=(Math.pow(candidate,.5)*maxWord)|0;k[primeCounter++]=(Math.pow(candidate,1/3)*maxWord)|0}}
    ascii+='\x80';while(ascii[lengthProperty]%64-56)ascii+='\x00';
    for(i=0;i<ascii[lengthProperty];i++){j=ascii.charCodeAt(i);if(j>>8)return sha256bytes(unescape(encodeURIComponent(ascii)));words[i>>2]|=j<<((3-i)%4)*8}
    words[words[lengthProperty]]=((asciiBitLength/maxWord)|0);words[words[lengthProperty]]=asciiBitLength;
    for(j=0;j<words[lengthProperty];){
      const w=words.slice(j,j+=16),oldHash=hash;hash=hash.slice(0,8);
      for(i=0;i<64;i++){
        const w15=w[i-15],w2=w[i-2],a=hash[0],e=hash[4];
        const temp1=hash[7]+(rightRotate(e,6)^rightRotate(e,11)^rightRotate(e,25))+((e&hash[5])^((~e)&hash[6]))+k[i]+(w[i]=(i<16)?w[i]:((w[i-16]+(rightRotate(w15,7)^rightRotate(w15,18)^(w15>>>3))+w[i-7]+(rightRotate(w2,17)^rightRotate(w2,19)^(w2>>>10)))|0));
        const temp2=(rightRotate(a,2)^rightRotate(a,13)^rightRotate(a,22))+((a&hash[1])^(a&hash[2])^(hash[1]&hash[2]));
        hash=[(temp1+temp2)|0].concat(hash);hash[4]=(hash[4]+temp1)|0;hash.pop();
      }
      for(i=0;i<8;i++)hash[i]=(hash[i]+oldHash[i])|0;
    }
    const bytes=[];for(i=0;i<8;i++)for(j=3;j+1;j--)bytes.push((hash[i]>>(j*8))&255);return bytes;
  }
  function tzParts(date,tz){
    const fmt=new Intl.DateTimeFormat('en-US',{timeZone:tz||'America/New_York',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23',weekday:'long'});
    const p={};for(const x of fmt.formatToParts(date))p[x.type]=x.value;return p;
  }
  function offsetZ(date,tz){
    try{
      const s=new Intl.DateTimeFormat('en-US',{timeZone:tz,timeZoneName:'longOffset',hour:'2-digit'}).formatToParts(date).find(x=>x.type==='timeZoneName')?.value||'GMT+00:00';
      const m=s.match(/GMT([+-])(\d{2}):?(\d{2})?/);return m?(m[1]+m[2]+(m[3]||'00')):'+0000';
    }catch{return '+0000'}
  }
  function formatDate(date,tz,pattern){
    const d=date instanceof Date?date:new Date(date);if(isNaN(d))return '';
    const p=tzParts(d,tz),mon=Number(p.month),day=Number(p.day);
    const shortWd=new Intl.DateTimeFormat('en-US',{timeZone:tz,weekday:'short'}).format(d);
    const longWd=new Intl.DateTimeFormat('en-US',{timeZone:tz,weekday:'long'}).format(d);
    const rep={yyyy:p.year,MM:p.month,dd:p.day,HH:p.hour,mm:p.minute,ss:p.second,EEE:shortWd,EEEE:longWd,M:String(mon),d:String(day),Z:offsetZ(d,tz)};
    return String(pattern).replace(/yyyy|EEEE|EEE|MM|dd|HH|mm|ss|M|d|Z/g,x=>rep[x]);
  }
  window.Utilities={
    getUuid:uuid,
    DigestAlgorithm:{SHA_256:'SHA_256'},
    Charset:{UTF_8:'UTF_8'},
    computeDigest:(_alg,s)=>sha256bytes(String(s)),
    base64EncodeWebSafe:bytes=>{let b='';for(const x of bytes)b+=String.fromCharCode((x+256)%256);return btoa(b).replace(/\+/g,'-').replace(/\//g,'_');},
    formatDate,
    sleep:()=>{}
  };

  async function pbkdf2Hex(password,saltHex,iterations){
    const salt=new Uint8Array((saltHex.match(/../g)||[]).map(x=>parseInt(x,16)));
    const key=await crypto.subtle.importKey('raw',enc(password),'PBKDF2',false,['deriveBits']);
    const bits=await crypto.subtle.deriveBits({name:'PBKDF2',salt,iterations:Number(iterations||180000),hash:'SHA-256'},key,256);
    return [...new Uint8Array(bits)].map(x=>x.toString(16).padStart(2,'0')).join('');
  }


  const GOOGLE_PTO_SHEET='PTO / Availability Requests';
  const GOOGLE_PTO_AWARE_FUNCTIONS=new Set([
    'getAppData',
    'generateSchedule',
    'preflightScheduleGeneration',
    'validateSavedSchedule',
    'validateConfiguration',
    'repairPtoAutoApprovals',
    'reviewRequest',
    'saveEmployeeRequest',
    'submitRequest'
  ]);

  function requestMatrixToObjects_(matrix){
    if(!Array.isArray(matrix)||!matrix.length)return [];
    const headers=(matrix[0]||[]).map(x=>String(x??'').trim());
    return matrix.slice(1).map(row=>{
      const o={};
      headers.forEach((h,i)=>o[h]=row?.[i]??'');
      return o;
    });
  }

  function mergeExternalPtoMatrix_(localMatrix,externalMatrix){
    // Google is authoritative for PTO only. Keep UNAVAILABLE / AVAILABILITY /
    // PREFERRED and any other non-PTO rules in the NeoChrono GitHub workbook.
    const headers=(window.__neoGooglePtoSource&&window.__neoGooglePtoSource.headers)
      ? window.__neoGooglePtoSource.headers.slice()
      : ['Record Type','Record ID','Pharmacist','Username','Date','Start Date','End Date','Weekend Saturday','Weekend Sunday','Available','Status','Comment','Submitted At','Reviewed By','Reviewed At','Updated At','Updated By'];

    const local=requestMatrixToObjects_(localMatrix||[]);
    const external=requestMatrixToObjects_(externalMatrix||[]);

    const rows=[
      ...external.filter(x=>String(x['Record Type']||'').trim().toUpperCase()==='PTO'),
      ...local.filter(x=>String(x['Record Type']||'').trim().toUpperCase()!=='PTO')
    ];

    return [headers,...rows.map(x=>headers.map(h=>x[h]??''))];
  }

  async function overlayGooglePto_(fn,data){
    if(!GOOGLE_PTO_AWARE_FUNCTIONS.has(String(fn)))return null;
    if(!window.__neoGooglePtoSource||typeof window.__neoGooglePtoSource.load!=='function')return null;

    const ext=await window.__neoGooglePtoSource.load(false);
    if(!ext)return null; // migration-safe fallback until Apps Script API is deployed

    data.sheets=data.sheets||{};
    const original=clone(data.sheets[GOOGLE_PTO_SHEET]||{values:[]});
    const localMatrix=original&&Array.isArray(original.values)?original.values:[];
    const merged=mergeExternalPtoMatrix_(localMatrix,ext.matrix);

    data.sheets[GOOGLE_PTO_SHEET]={values:merged};
    return {original,source:ext.source,recordIds:ext.recordIds};
  }

  window.__neoRuntime={
    cfg,loadWorkbook,saveWorkbook,savePublished,pbkdf2Hex,getEmployeePhotoDataUrl,uploadEmployeePhoto,deleteEmployeePhoto,employeePhotoPath:employeePhotoPath_,
    currentBook:()=>currentBook(),
    markDirty:()=>{if(window.__neoVBook)window.__neoVBook.dirty=true;},
    invoke(fn,args){
      const work=async()=>{
        for(let attempt=0;attempt<4;attempt++){
          // Always start each server-style call from the newest GitHub file.
          // This prevents stale SHA values after workbook imports or another open tab.
          const loaded=await loadWorkbook(true);
          const data=clone(loaded.data);
          const ptoOverlay=await overlayGooglePto_(fn,data);
          const book=new VBook(data);window.__neoVBook=book;
          try{ if(typeof _DB_CACHE!=='undefined') _DB_CACHE=null; }catch(_e){}
          try{ if(typeof _TZ_CACHE!=='undefined') _TZ_CACHE=null; }catch(_e){}
          try{ if(typeof PRECEPTOR_CALENDAR_RUNTIME_CACHE_!=='undefined') PRECEPTOR_CALENDAR_RUNTIME_CACHE_=null; }catch(_e){}
          const callable=window[fn];
          if(typeof callable!=='function')throw new Error('Backend function not found: '+fn);
          try{
            // The Google Sheet is authoritative for the entire PTO / Availability
            // Requests section. Save/review there directly and verify by re-reading it.
            if(ptoOverlay&&String(fn)==='saveEmployeeRequest'){
              const payload=(args||[])[1]||{};
              if(String(payload['Record Type']||'PTO').trim().toUpperCase()==='PTO'){
                return await window.__neoGooglePtoSource.saveRequest(payload);
              }
            }
            if(ptoOverlay&&String(fn)==='submitRequest'){
              const payload=(args||[])[1]||{};
              if(String(payload['Record Type']||'PTO').trim().toUpperCase()==='PTO'){
                return await window.__neoGooglePtoSource.saveRequest(payload);
              }
            }
            if(ptoOverlay&&String(fn)==='reviewRequest'){
              const recordId=String((args||[])[1]||'');
              const status=String((args||[])[2]||'');
              const comment=(args||[])[3];
              return await window.__neoGooglePtoSource.reviewRequest(recordId,status,comment);
            }

            let result=callable.apply(window,args||[]);
            if(result&&typeof result.then==='function')result=await result;

            // Never persist Google PTO rows into neochrono-data. Google Sheets
            // is authoritative for PTO only; non-PTO availability stays in GitHub.
            if(ptoOverlay){
              book.data.sheets=book.data.sheets||{};
              book.data.sheets[GOOGLE_PTO_SHEET]=clone(ptoOverlay.original);
            }

            const actualChanged=book.dirty && JSON.stringify(book.data)!==JSON.stringify(loaded.data);
            if(actualChanged){
              try{
                await saveWorkbook(book.data,loaded.sha,'NeoChrono: '+fn);
              }catch(e){
                const conflict =
                  e.status===409 ||
                  e.status===422 ||
                  /does not match|sha mismatch|conflict/i.test(String(e&&e.message||''));
                if(conflict && attempt<3){
                  cache={data:null,sha:null,loadedAt:0};
                  await new Promise(resolve=>setTimeout(resolve,250*(attempt+1)));
                  continue;
                }
                throw e;
              }
            }else{
              cache={data:book.data,sha:loaded.sha,loadedAt:Date.now()};
            }
            return result;
          }finally{window.__neoVBook=null;}
        }
        throw new Error('NeoChrono detected another update to the GitHub database while saving. It retried automatically but the file kept changing. Close other NeoChrono tabs and try again.');
      };
      const p=invokeQueue.then(work,work);
      invokeQueue=p.catch(()=>{});
      return p;
    }
  };

  function chain(state){
    return new Proxy({},{
      get(_t,prop){
        if(prop==='withSuccessHandler')return fn=>chain({...state,success:fn});
        if(prop==='withFailureHandler')return fn=>chain({...state,failure:fn});
        return (...args)=>{window.__neoRuntime.invoke(String(prop),args).then(r=>state.success&&state.success(r)).catch(e=>state.failure?state.failure({message:e.message}):console.error(e));};
      }
    });
  }
  const run=new Proxy({},{
    get(_t,prop){
      if(prop==='withSuccessHandler')return fn=>chain({success:fn,failure:null});
      if(prop==='withFailureHandler')return fn=>chain({success:null,failure:fn});
      return (...args)=>window.__neoRuntime.invoke(String(prop),args);
    }
  });
  window.google={script:{run}};

  window.addEventListener('neo:publish',async e=>savePublished(e.detail));
})();