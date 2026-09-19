
(function(){
  'use strict';

  const ENDPOINT='https://script.google.com/macros/s/AKfycbxBTnuzFvNXOROz4fpGni_exZap2YfSTm5aNWw4eAwnyGe-m9jU3SZriupyFf3yDP-r/exec';
  const SHEET='PTO / Availability Requests';
  const READ_ACTION='neochronoPto';
  const WRITE_ACTION='neochronoPtoWrite';
  const CACHE_KEY='neochrono_google_pto_cache_v3';
  const ACTIVE_KEY='neochrono_google_pto_active_v3';
  const CACHE_MS=15000;

  const HEADERS=[
    'Record Type','Record ID','Pharmacist','Username','Date','Start Date','End Date',
    'Weekend Saturday','Weekend Sunday','Available','Status','Comment','Submitted At',
    'Reviewed By','Reviewed At','Updated At','Updated By'
  ];

  let memory={matrix:null,loadedAt:0,lastError:'',source:'google',recordIds:new Set()};

  function clone(v){return JSON.parse(JSON.stringify(v));}

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
    if(!payload||payload.success===false){
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
      throw new Error('Google PTO API response does not contain request rows.');
    }

    const sourceHeaders=(matrix[0]||[]).map(x=>String(x??'').trim());
    const index={};
    sourceHeaders.forEach((h,i)=>index[h]=i);

    const normalized=[HEADERS.slice()];
    for(let r=1;r<matrix.length;r++){
      const row=matrix[r]||[];
      const out=HEADERS.map(h=>index[h]!==undefined?(row[index[h]]??''):'');
      const recordType=String(out[0]||'').trim();
      const recordId=String(out[1]||'').trim();
      if(!recordType&&!recordId)continue;
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

  function jsonp(force){
    return new Promise((resolve,reject)=>{
      const cb='__neoPtoCb_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const script=document.createElement('script');
      const timer=setTimeout(()=>{
        cleanup();
        reject(new Error('Google PTO source did not return data. Make sure the PTO read API is deployed at the configured Apps Script web app URL.'));
      },12000);

      function cleanup(){
        clearTimeout(timer);
        try{delete window[cb];}catch(_e){window[cb]=undefined;}
        if(script.parentNode)script.parentNode.removeChild(script);
      }

      window[cb]=payload=>{
        cleanup();
        try{resolve(normalizePayload(payload));}
        catch(e){reject(e);}
      };

      script.onerror=()=>{
        cleanup();
        reject(new Error('Could not load PTO / Availability Requests from Google Apps Script.'));
      };

      const qs=new URLSearchParams({
        action:READ_ACTION,
        sheet:SHEET,
        callback:cb,
        _:String(Date.now())
      });
      script.src=ENDPOINT+'?'+qs.toString();
      script.async=true;
      document.head.appendChild(script);
    });
  }

  async function load(force=false){
    if(memory.matrix&&!force&&Date.now()-memory.loadedAt<CACHE_MS){
      return {matrix:clone(memory.matrix),source:memory.source,recordIds:new Set(memory.recordIds)};
    }

    try{
      const matrix=await jsonp(force);
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
      // Before the first successful Code5 connection, keep NeoChrono usable
      // with its existing GitHub copy. Once Google has connected, use cache.
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
    const body=JSON.stringify({
      action:WRITE_ACTION,
      sheet:SHEET,
      payload:payload
    });

    try{
      await fetch(ENDPOINT,{
        method:'POST',
        mode:'no-cors',
        cache:'no-store',
        headers:{'Content-Type':'text/plain;charset=UTF-8'},
        body:body
      });
    }catch(e){
      throw new Error('Could not send the PTO change to Google Sheets. '+(e&&e.message?e.message:''));
    }

    // Apps Script writes are asynchronous from the browser's perspective.
    // Re-read the sheet after a short delay and verify the change actually landed.
    await new Promise(resolve=>setTimeout(resolve,900));
    memory.loadedAt=0;
  }

  async function saveRequest(data){
    const row={...(data||{})};
    row['Record Type']=String(row['Record Type']||'PTO').trim().toUpperCase();
    row['Record ID']=String(row['Record ID']||'').trim()||makeRecordId();

    await postWrite({operation:'save',row:row});

    const refreshed=await load(true);
    if(!refreshed){
      throw new Error('The Google PTO write API is not active yet. Confirm the PTO read/write API is deployed at the configured Apps Script web app URL.');
    }
    const rows=matrixToObjects(refreshed.matrix);
    const saved=rows.find(r=>String(r['Record ID']||'')===row['Record ID']);
    if(!saved){
      throw new Error('Google Apps Script accepted the request, but the row was not found in the Google Sheet after saving. Confirm the configured Apps Script deployment exposes the PTO API.');
    }

    return {
      ok:true,
      recordId:row['Record ID'],
      status:String(saved.Status||''),
      message:'Saved to Google Sheet: '+SHEET
    };
  }

  async function reviewRequest(recordId,status,comment){
    const id=String(recordId||'').trim();
    if(!id)throw new Error('Request ID is required.');

    await postWrite({
      operation:'review',
      recordId:id,
      status:String(status||'').trim(),
      comment:comment===undefined?'':String(comment)
    });

    const refreshed=await load(true);
    if(!refreshed){
      throw new Error('The Google PTO write API is not active yet. Confirm the PTO API is deployed at the configured Apps Script web app URL.');
    }
    const rows=matrixToObjects(refreshed.matrix);
    const saved=rows.find(r=>String(r['Record ID']||'')===id);
    if(!saved){
      throw new Error('The request could not be found in the Google Sheet after review.');
    }
    if(String(saved.Status||'').trim().toLowerCase()!==String(status||'').trim().toLowerCase()){
      throw new Error('The Google Sheet did not confirm the requested status change. Check the PTO Code4.gs deployment.');
    }

    return {ok:true,recordId:id,status:saved.Status,message:'Review saved to Google Sheet.'};
  }

  async function removeRequest(recordId){
    const id=String(recordId||'').trim();
    if(!id)throw new Error('Request ID is required.');
    await postWrite({operation:'delete',recordId:id});
    const refreshed=await load(true);
    if(!refreshed){
      throw new Error('The Google PTO write API is not active yet. Confirm the PTO API is deployed at the configured Apps Script web app URL.');
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
