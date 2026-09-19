
(function(){
  'use strict';

  const RECEIVER_URL='https://script.google.com/macros/s/AKfycbyKpDf3Fe6TyvX0vxrQ6_5O18f1DT2ZiEciQbMEC_VObA2WE_COLzXXzwRRMjR17clK/exec';
  const SHEET_NAME='Schedule';

  const HEADERS=[
    'Generation ID',
    'Assignment ID',
    'Date',
    'Day',
    'Shift',
    'Slot',
    'Assigned Pharmacist',
    'Username',
    'Hours',
    'Credited Hours',
    'Required Skill',
    'Coverage For Pharmacist',
    'Coverage For Username',
    'Coverage Reason',
    'Shift Type',
    'Weekend',
    'Weekend Group',
    'Holiday',
    'Locked',
    'Manual',
    'Status',
    'Warning',
    'Updated At',
    'Updated By',
    'Finalized At'
  ];

  function clean_(value){
    if(value===null||value===undefined)return '';
    if(value instanceof Date&&!isNaN(value))return value.toISOString();
    return String(value);
  }

  function dateKey_(row){
    const raw=clean_(row&&row.Date);
    return raw?raw.slice(0,10):'';
  }

  function scheduleRows_(startDate,endDate){
    return (State.data.schedule||[]).filter(function(r){
      const dk=dateKey_(r);
      if(!dk)return false;
      if(startDate&&dk<startDate)return false;
      if(endDate&&dk>endDate)return false;
      return true;
    });
  }

  function makeTransferId_(){
    if(window.crypto&&crypto.randomUUID)return crypto.randomUUID();
    return 'TX-'+Date.now()+'-'+Math.random().toString(36).slice(2);
  }

  function payload_(startDate,endDate){
    const rows=scheduleRows_(startDate,endDate);

    if(!rows.length){
      throw new Error('No schedule rows were found in the selected date range.');
    }

    return {
      action:'replaceSchedule',
      sheetName:SHEET_NAME,
      transferId:makeTransferId_(),
      startDate:startDate||'',
      endDate:endDate||'',
      sentAt:new Date().toISOString(),
      headers:HEADERS.slice(),
      rows:rows.map(function(r){
        return HEADERS.map(function(h){
          return clean_(r[h]);
        });
      })
    };
  }

  function postPayload_(payload){
    return new Promise(function(resolve,reject){
      const frameName='neoSchedulePost_'+Date.now()+'_'+Math.random().toString(36).slice(2);

      const iframe=document.createElement('iframe');
      iframe.name=frameName;
      iframe.src='about:blank';
      iframe.style.display='none';

      const form=document.createElement('form');
      form.method='POST';
      form.action=RECEIVER_URL;
      form.target=frameName;
      form.style.display='none';

      function addField(name,value){
        const input=document.createElement('input');
        input.type='hidden';
        input.name=name;
        input.value=String(value===null||value===undefined?'':value);
        form.appendChild(input);
      }

      addField('action','replaceSchedule');
      addField('sheet',SHEET_NAME);
      addField('transferId',payload.transferId);
      addField('payload',JSON.stringify(payload));

      let finished=false;
      let timeoutTimer=null;
      let fallbackTimer=null;
      let initialFrameLoaded=false;
      let requestSubmitted=false;
      let responseFrameLoaded=false;

      function cleanup(){
        if(finished)return;
        finished=true;

        if(timeoutTimer)clearTimeout(timeoutTimer);
        if(fallbackTimer)clearTimeout(fallbackTimer);

        window.removeEventListener('message',onMessage);

        try{form.remove();}catch(_e){}
        try{iframe.remove();}catch(_e){}
      }

      function fail(message){
        if(finished)return;
        cleanup();
        reject(new Error(message));
      }

      function succeed(result){
        if(finished)return;
        cleanup();
        resolve(result);
      }

      function onMessage(event){
        const data=event&&event.data;

        if(!data||data.type!=='NEOCHRONO_SCHEDULE_RECEIVER')return;
        if(String(data.transferId||'')!==String(payload.transferId||''))return;

        if(!data.ok){
          fail(
            data.message||
            'Google Schedule receiver rejected the transfer.'
          );
          return;
        }

        succeed({
          ...data,
          confirmationReceived:true,
          submitted:true
        });
      }

      window.addEventListener('message',onMessage);

      /*
       * The hidden iframe loads twice:
       *   1) about:blank
       *   2) Google Apps Script response after form.submit()
       *
       * Some browsers/Google wrappers block the postMessage response even
       * though the POST completed successfully. In that case, the second
       * iframe load is used as a safe submission fallback.
       */
      iframe.addEventListener('load',function(){
        if(finished)return;

        if(!initialFrameLoaded){
          initialFrameLoaded=true;

          if(requestSubmitted)return;

          requestSubmitted=true;

          try{
            form.submit();
          }catch(e){
            fail(
              'Could not submit the schedule to Google: '+
              (e&&e.message?e.message:String(e))
            );
          }

          return;
        }

        if(!requestSubmitted||responseFrameLoaded)return;

        responseFrameLoaded=true;

        /*
         * Give postMessage a short grace period first.
         * If it never arrives, resolve from the completed response load.
         */
        fallbackTimer=setTimeout(function(){
          if(finished)return;

          succeed({
            type:'NEOCHRONO_SCHEDULE_RECEIVER',
            ok:true,
            submitted:true,
            confirmationReceived:false,
            transferId:payload.transferId,
            transferredRows:Array.isArray(payload.rows)?payload.rows.length:0,
            writtenRows:Array.isArray(payload.rows)?payload.rows.length:0,
            transferredColumns:Array.isArray(payload.headers)?payload.headers.length:0,
            writtenColumns:Array.isArray(payload.headers)?payload.headers.length:0,
            receiverSheetName:SHEET_NAME,
            receiverSpreadsheetName:'Pharmacists Schedule',
            receiverUrl:RECEIVER_URL,
            startDate:payload.startDate||'',
            endDate:payload.endDate||'',
            message:
              'The schedule submission completed in Google. '+
              'Google did not return the browser confirmation message, so NeoChrono used the completed receiver-page load as confirmation.'
          });
        },1200);
      });

      iframe.addEventListener('error',function(){
        fail('The Google Schedule receiver could not be reached.');
      });

      timeoutTimer=setTimeout(function(){
        if(finished)return;

        fail(
          'The Google Schedule receiver did not finish loading after the transfer. '+
          'Confirm the web app deployment URL and access settings.'
        );
      },30000);

      /*
       * Append the form first, then the iframe. The first about:blank load
       * starts the POST so we can distinguish it from the Google response load.
       */
      document.body.appendChild(form);
      document.body.appendChild(iframe);
    });
  }

  async function sendSchedule(startDate,endDate){
    const payload=payload_(startDate,endDate);
    const result=await postPayload_(payload);

    const expectedRows=payload.rows.length;
    const gotRows=Number(result.transferredRows||0);
    const gotCols=Number(result.transferredColumns||0);

    if(gotRows!==expectedRows){
      throw new Error(
        'Google confirmed the transfer, but row count does not match. '+
        'NeoChrono sent '+expectedRows+' row(s); Google wrote '+gotRows+'.'
      );
    }

    if(gotCols!==HEADERS.length){
      throw new Error(
        'Google confirmed the transfer, but expected 25 columns and received '+gotCols+'.'
      );
    }

    return {
      ok:true,
      transferredRows:gotRows,
      transferredColumns:gotCols,
      receiverSheetName:result.receiverSheetName||SHEET_NAME,
      receiverSpreadsheetName:result.receiverSpreadsheetName||'Pharmacists Schedule',
      receiverUrl:result.receiverUrl||RECEIVER_URL,
      transferId:payload.transferId,
      message:result.message||'Schedule written to Google Sheet.'
    };
  }

  window.__neoScheduleReceiver={
    url:RECEIVER_URL,
    sheet:SHEET_NAME,
    headers:HEADERS.slice(),
    sendSchedule:sendSchedule
  };
})();
