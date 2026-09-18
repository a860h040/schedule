
(function(){
  'use strict';

  const ENDPOINT='https://script.google.com/macros/s/AKfycbxBTnuzFvNXOROz4fpGni_exZap2YfSTm5aNWw4eAwnyGe-m9jU3SZriupyFf3yDP-r/exec';
  const SHEET='PTO / Availability Requests';
  const ACTION='neochronoPto';
  const CACHE_KEY='neochrono_google_pto_cache_v1';
  const ACTIVE_KEY='neochrono_google_pto_active_v1';
  const CACHE_MS=30000;

  const HEADERS=[
    'Record Type','Record ID','Pharmacist','Username','Date','Start Date','End Date',
    'Weekend Saturday','Weekend Sunday','Available','Status','Comment','Submitted At',
    'Reviewed By','Reviewed At','Updated At','Updated By'
  ];

  let memory={matrix:null,loadedAt:0,lastError:'',source:'github',recordIds:new Set()};

  function clone(v){return JSON.parse(JSON.stringify(v));}

  function objectRowsToMatrix(rows){
    return [
      HEADERS.slice(),
      ...(rows||[]).map(r=>HEADERS.map(h=>r&&r[h]!==undefined?r[h]:''))
    ];
  }

  function normalizePayload(payload){
    if(!payload||payload.success===false){
      throw new Error(payload&&payload.message?payload.message:'Google PTO API returned no data.');
    }

    let matrix=null;

    if(Array.isArray(payload.values)){
      matrix=payload.values;
    }else if(Array.isArray(payload.matrix)){
      matrix=payload.matrix;
    }else if(Array.isArray(payload.rows)&&payload.rows.length&&Array.isArray(payload.rows[0])){
      matrix=[Array.isArray(payload.headers)?payload.headers:HEADERS].concat(payload.rows);
    }else if(Array.isArray(payload.rows)){
      matrix=objectRowsToMatrix(payload.rows);
    }else if(Array.isArray(payload.requests)){
      matrix=objectRowsToMatrix(payload.requests);
    }else if(Array.isArray(payload.data)){
      matrix=payload.data.length&&Array.isArray(payload.data[0])
        ? [Array.isArray(payload.headers)?payload.headers:HEADERS].concat(payload.data)
        : objectRowsToMatrix(payload.data);
    }

    if(!Array.isArray(matrix)||!matrix.length){
      throw new Error('Google PTO API response does not contain PTO rows.');
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
      if(!recordType&&!recordId)continue;
      normalized.push(out);
    }

    return normalized;
  }

  function recordIdsFromMatrix(matrix){
    const out=new Set();
    if(!matrix||matrix.length<2)return out;
    const headers=matrix[0].map(x=>String(x??'').trim());
    const idIdx=headers.indexOf('Record ID');
    const typeIdx=headers.indexOf('Record Type');
    if(idIdx<0)return out;
    for(let i=1;i<matrix.length;i++){
      const type=typeIdx>=0?String(matrix[i]?.[typeIdx]||'').trim().toUpperCase():'PTO';
      if(type!=='PTO')continue;
      const id=String(matrix[i]?.[idIdx]||'').trim();
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

  function jsonp(){
    return new Promise((resolve,reject)=>{
      const cb='__neoPtoCb_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const script=document.createElement('script');
      const timer=setTimeout(()=>{
        cleanup();
        reject(new Error('Google PTO source did not return data. The Apps Script PTO API may not be deployed yet.'));
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
        reject(new Error('Could not load PTO data from the Google Apps Script deployment.'));
      };

      const qs=new URLSearchParams({
        action:ACTION,
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
      const matrix=await jsonp();
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

      // Once Google PTO has successfully connected, never go back to GitHub PTO.
      // Use the last Google copy if the Google endpoint has a temporary outage.
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

      // Migration-safe state before the Apps Script API is deployed.
      // The legacy GitHub PTO source remains temporarily available so the
      // scheduling app does not break during deployment.
      memory.source='github-fallback';
      return null;
    }
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

  function isExternalPtoRecord(recordId){
    return memory.recordIds.has(String(recordId||''));
  }

  window.__neoGooglePtoSource={
    endpoint:ENDPOINT,
    sheet:SHEET,
    headers:HEADERS.slice(),
    load,
    status,
    isExternalPtoRecord
  };
})();
