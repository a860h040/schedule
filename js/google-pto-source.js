(function(){
  'use strict';

  const ENDPOINT='https://script.google.com/macros/s/AKfycbzd-2TtruGPlfUbRlS7EDFlm7jgBlBUgfM66PJX03zMpZOHFoXBSXBd-rJMX7s71fXk/exec';
  const SHEET='PTO / Availability Requests';
  const READ_ACTION='neochronoPto';
  const WRITE_ACTION='neochronoPtoWrite';
  const MESSAGE_TYPE='NEOCHRONO_PTO_RECEIVER';
  const CACHE_KEY='neochrono_google_pto_cache_v4';
  const ACTIVE_KEY='neochrono_google_pto_active_v4';
  const CACHE_MS=15000;

  const HEADERS=[
    'Record Type','Record ID','Pharmacist','Username','Date','Start Date','End Date',
    'Weekend Saturday','Weekend Sunday','Available','Status','Comment','Submitted At',
    'Reviewed By','Reviewed At','Updated At','Updated By'
  ];

  let memory={matrix:null,loadedAt:0,lastError:'',source:'google',recordIds:new Set()};

  function clone(v){return JSON.parse(JSON.stringify(v));}

  function actor_(){
    try{
      const u=window.State&&State.data&&State.data.user?State.data.user:{};
      return String(u.Username||u.username||u['Pharmacist Name']||u.Name||'NeoChrono').trim()||'NeoChrono';
    }catch(_e){
      return 'NeoChrono';
    }
  }

  function objectRowsToMatrix(rows){
    return [
      HEADERS.slice(),
      ...(rows||[]).map(r=>HEADERS.map(h=>r&&r[h]!==undefined?r[h]:''))
    ];
  }

  function matrixToObjects(matrix){
    if(!Array.isArray(matrix)||!matrix.length)return [];
    const headers=(matrix[0]||[]).map(x=>String(x??'').trim());
    return matrix.slice(1).map(row=>{
      const o={};
      headers.forEach((h,i)=>o[h]=row?.[i]??'');
      return o;
    });
  }

  function normalizePayload(payload){
    if(!payload||payload.ok===false||payload.success===false){
      throw new Error(payload&&payload.message?payload.message:'Google PTO API returned no data.');
    }

    let matrix=null;
    if(Array.isArray(payload.values))matrix=payload.values;
    else if(Array.isArray(payload.matrix))matrix=payload.matrix;
    else if(Array.isArray(payload.rows)&&payload.rows.length&&Array.isArray(payload.rows[0])){
      matrix=[Array.isArray(payload.headers)?payload.headers:HEADERS].concat(payload.rows);
    }else if(Array.isArray(payload.rows))matrix=objectRowsToMatrix(payload.rows);
    else if(Array.isArray(payload.requests))matrix=objectRowsToMatrix(payload.requests);
    else if(Array.isArray(payload.data)){
      matrix=payload.data.length&&Array.isArray(payload.data[0])
        ? [Array.isArray(payload.headers)?payload.headers:HEADERS].concat(payload.data)
        : objectRowsToMatrix(payload.data);
    }

    if(!Array.isArray(matrix)||!matrix.length){
      matrix=[HEADERS.slice()];
    }

    const sourceHeaders=(matrix[0]||[]).map(x=>String(x??'').trim());
    const index={};
    sourceHeaders.forEach((h,i)=>index[h]=i);

    const normalized=[HEADERS.slice()];
    for(let r=1;r<matrix.length;r++){
      const row=matrix[r]||[];
      const out=HEADERS.map(h=>index[h]!==undefined?(row[index[h]]??''):'');
      const recordType=String(out[0]||'').trim().toUpperCase();
      const recordId=String(out[1]||'').trim();
      if(recordType!=='PTO')continue;
      if(!recordId)continue;
      out[0]='PTO';
      normalized.push(out);
    }
    return normalized;
  }

  function recordIdsFromMatrix(matrix){
    const out=new Set();
    for(const r of matrixToObjects(matrix)){
      const id=String(r['Record ID']||'').trim();
      if(id)out.add(id);
    }
    return out;
  }

  function saveCache(matrix){
    try{
      localStorage.setItem(CACHE_KEY,JSON.stringify({matrix,at:Date.now()}));
      localStorage.setItem(ACTIVE_KEY,'1');
    }catch(_e){}
  }

  function loadCache(){
    try{
      const raw=JSON.parse(localStorage.getItem(CACHE_KEY)||'null');
      if(raw&&Array.isArray(raw.matrix)&&raw.matrix.length)return raw;
    }catch(_e){}
    return null;
  }

  function wasActivated(){
    try{return localStorage.getItem(ACTIVE_KEY)==='1';}catch(_e){return false;}
  }

  function requestId_(){
    if(window.crypto&&crypto.randomUUID)return crypto.randomUUID();
    return 'PTO-'+Date.now()+'-'+Math.random().toString(36).slice(2);
  }

  function postBridge_(action,payload,timeoutMs){
    return new Promise((resolve,reject)=>{
      const requestId=requestId_();
      const frameName='neoPtoFrame_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const iframe=document.createElement('iframe');
      iframe.name=frameName;
      iframe.style.display='none';

      const form=document.createElement('form');
      form.method='POST';
      form.action=ENDPOINT;
      form.target=frameName;
      form.style.display='none';

      function addField(name,value){
        const input=document.createElement('input');
        input.type='hidden';
        input.name=name;
        input.value=String(value===null||value===undefined?'':value);
        form.appendChild(input);
      }

      addField('action',action);
      addField('sheet',SHEET);
      addField('requestId',requestId);
      if(payload!==undefined)addField('payload',JSON.stringify(payload));

      let finished=false;
      let timer=null;

      function cleanup(){
        if(finished)return;
        finished=true;
        if(timer)clearTimeout(timer);
        window.removeEventListener('message',onMessage);
        try{form.remove();}catch(_e){}
        try{iframe.remove();}catch(_e){}
      }

      function fail(message){
        cleanup();
        reject(new Error(message));
      }

      function onMessage(event){
        const data=event&&event.data;
        if(!data||data.type!==MESSAGE_TYPE)return;
        if(String(data.requestId||'')!==String(requestId))return;
        cleanup();
        if(data.ok===false||data.success===false){
          reject(new Error(data.message||'Google PTO receiver rejected the request.'));
          return;
        }
        resolve(data);
      }

      window.addEventListener('message',onMessage);
      timer=setTimeout(()=>{
        fail('The Google PTO receiver did not confirm the request. Make sure Code5.gs is saved and the Apps Script web app is redeployed as a new version.');
      },Number(timeoutMs||20000));

      iframe.onerror=()=>fail('The Google PTO receiver could not be reached.');

      document.body.appendChild(iframe);
      document.body.appendChild(form);

      try{form.submit();}
      catch(e){fail(e&&e.message?e.message:String(e));}
    });
  }

  async function load(force=false){
    if(memory.matrix&&!force&&Date.now()-memory.loadedAt<CACHE_MS){
      return {matrix:clone(memory.matrix),source:memory.source,recordIds:new Set(memory.recordIds)};
    }

    try{
      const response=await postBridge_(READ_ACTION,undefined,20000);
      const matrix=normalizePayload(response);
      memory={
        matrix,
        loadedAt:Date.now(),
        lastError:'',
        source:'google',
        recordIds:recordIdsFromMatrix(matrix)
      };
      saveCache(matrix);
      return {matrix:clone(matrix),source:'google',recordIds:new Set(memory.recordIds)};
    }catch(e){
      memory.lastError=String(e&&e.message||e);
      const cached=loadCache();
      if(wasActivated()&&cached&&Array.isArray(cached.matrix)){
        const ids=recordIdsFromMatrix(cached.matrix);
        memory={
          matrix:cached.matrix,
          loadedAt:Date.now(),
          lastError:memory.lastError,
          source:'google-cache',
          recordIds:ids
        };
        return {matrix:clone(cached.matrix),source:'google-cache',recordIds:new Set(ids)};
      }
      memory.source='github-fallback';
      return null;
    }
  }

  function makeRecordId(){
    const d=new Date();
    const pad=n=>String(n).padStart(2,'0');
    const stamp=''+d.getFullYear()+pad(d.getMonth()+1)+pad(d.getDate())+
      pad(d.getHours())+pad(d.getMinutes())+pad(d.getSeconds());
    return 'REQ-'+stamp+'-'+Math.floor(Math.random()*900+100);
  }

  async function postWrite(payload){
    const response=await postBridge_(WRITE_ACTION,payload,25000);
    memory.loadedAt=0;
    return response;
  }

  async function saveRequest(data){
    const row={...(data||{})};
    row['Record Type']=String(row['Record Type']||'PTO').trim().toUpperCase();
    if(row['Record Type']!=='PTO')throw new Error('Google PTO source only accepts PTO records.');
    row['Record ID']=String(row['Record ID']||'').trim()||makeRecordId();
    row['Updated By']=String(row['Updated By']||actor_()).trim()||actor_();

    const response=await postWrite({
      operation:'save',
      row:row,
      actor:actor_()
    });

    const refreshed=await load(true);
    if(!refreshed){
      throw new Error('The Google PTO API is not active yet. Confirm Code5.gs is deployed at the configured Apps Script web app URL.');
    }
    const rows=matrixToObjects(refreshed.matrix);
    const saved=rows.find(r=>String(r['Record ID']||'')===row['Record ID']);
    if(!saved){
      throw new Error('Google confirmed the save, but the PTO row was not found when NeoChrono re-read the Google Sheet.');
    }

    return {
      ok:true,
      recordId:row['Record ID'],
      status:String(saved.Status||response.status||''),
      autoApproved:!!response.autoApproved,
      blockedDates:Array.isArray(response.blockedDates)?response.blockedDates:[],
      message:response.message||('Saved to Google Sheet: '+SHEET)
    };
  }

  async function reviewRequest(recordId,status,comment){
    const id=String(recordId||'').trim();
    if(!id)throw new Error('Request ID is required.');

    const response=await postWrite({
      operation:'review',
      recordId:id,
      status:String(status||'').trim(),
      comment:comment===undefined?'':String(comment),
      actor:actor_()
    });

    const refreshed=await load(true);
    if(!refreshed){
      throw new Error('The Google PTO API is not active yet. Confirm Code5.gs is deployed at the configured Apps Script web app URL.');
    }
    const rows=matrixToObjects(refreshed.matrix);
    const saved=rows.find(r=>String(r['Record ID']||'')===id);
    if(!saved)throw new Error('The request could not be found in the Google Sheet after review.');
    if(String(saved.Status||'').trim().toLowerCase()!==String(status||'').trim().toLowerCase()){
      throw new Error('The Google Sheet did not confirm the requested status change. Check the Code5.gs deployment.');
    }

    return {ok:true,recordId:id,status:saved.Status,message:response.message||'Review saved to Google Sheet.'};
  }

  async function removeRequest(recordId){
    const id=String(recordId||'').trim();
    if(!id)throw new Error('Request ID is required.');

    await postWrite({
      operation:'delete',
      recordId:id,
      actor:actor_()
    });

    const refreshed=await load(true);
    if(!refreshed){
      throw new Error('The Google PTO API is not active yet. Confirm Code5.gs is deployed at the configured Apps Script web app URL.');
    }
    const rows=matrixToObjects(refreshed.matrix);
    if(rows.some(r=>String(r['Record ID']||'')===id)){
      throw new Error('The request still exists in the Google Sheet after delete.');
    }
    return {ok:true,recordId:id};
  }

  function status(){
    return {
      endpoint:ENDPOINT,
      sheet:SHEET,
      active:wasActivated(),
      source:memory.source,
      lastError:memory.lastError,
      recordIds:new Set(memory.recordIds)
    };
  }

  function isExternalRecord(recordId){
    return memory.recordIds.has(String(recordId||''));
  }

  window.__neoGooglePtoSource={
    endpoint:ENDPOINT,
    sheet:SHEET,
    headers:HEADERS.slice(),
    load,
    saveRequest,
    reviewRequest,
    removeRequest,
    status,
    isExternalRecord,
    isExternalPtoRecord:isExternalRecord
  };
})();
