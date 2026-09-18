
(function(){
  'use strict';

  function rows_(){
    return (State.data && State.data.schedule ? State.data.schedule : []).filter(function(r){
      var status=String(r.Status||'').toUpperCase();
      return status!=='CANCELLED' && String(r.Date||'').trim();
    });
  }

  function localDate_(value){
    var s=String(value||'').slice(0,10);
    var m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if(!m)return null;
    var d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]));
    return Number.isNaN(d.getTime())?null:d;
  }

  function defaults_(){
    var dates=rows_().map(function(r){return String(r.Date||'').slice(0,10);}).filter(Boolean).sort();
    if(dates.length)return {start:dates[0],end:dates[dates.length-1]};
    var now=new Date(),sun=addDays(now,-now.getDay());
    return {start:dateKey(sun),end:dateKey(addDays(sun,6))};
  }

  function state_(){
    if(!State.paperSchedule){
      var d=defaults_();
      State.paperSchedule={start:d.start,end:d.end,showUnfilled:true,showTimes:true,compact:false};
    }
    return State.paperSchedule;
  }

  function shiftOrder_(){
    var rows=(State.data.shifts||[]).slice(),meta={},i;
    rows.forEach(function(s,idx){
      var code=String(s.Shift||'').trim().toUpperCase();
      if(!code)return;
      meta[code]={
        priority:num(s.Priority,50),
        start:String(s.Start||''),
        index:idx
      };
    });
    var codes=Array.from(new Set(rows_().map(function(r){return String(r.Shift||'').trim().toUpperCase();}).filter(Boolean)));
    codes.sort(function(a,b){
      var A=meta[a]||{priority:50,start:'',index:9999};
      var B=meta[b]||{priority:50,start:'',index:9999};
      return A.priority-B.priority || A.start.localeCompare(B.start) || A.index-B.index || a.localeCompare(b);
    });
    return codes;
  }

  function dayHeader_(d){
    var names=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    return names[d.getDay()]+'<br><strong>'+(d.getMonth()+1)+'/'+d.getDate()+'</strong>';
  }

  function cell_(dateCode,shiftCode){
    var ps=state_();
    var list=rows_().filter(function(r){
      return String(r.Date||'').slice(0,10)===dateCode &&
        String(r.Shift||'').trim().toUpperCase()===shiftCode;
    }).sort(function(a,b){return num(a.Slot,1)-num(b.Slot,1);});

    var html=[];
    list.forEach(function(r){
      var status=String(r.Status||'').toUpperCase();
      var assigned=String(r['Assigned Pharmacist']||r.Username||'').trim();
      var unfilled=status==='UNFILLED' || assigned.toUpperCase()==='UNFILLED' || !assigned;

      if(unfilled){
        if(ps.showUnfilled)html.push('<div class="paper-person paper-unfilled"><span>UNFILLED</span></div>');
        return;
      }

      var flags=[];
      if(yes(r.Locked))flags.push('L');
      if(yes(r.Manual))flags.push('M');
      html.push(
        '<div class="paper-person"><span>'+esc(assigned)+'</span>'+
        (flags.length?'<small>'+esc(flags.join('·'))+'</small>':'')+
        '</div>'
      );
    });

    return html.length?html.join(''):'<span class="paper-empty">—</span>';
  }

  function weekBlock_(weekStart,index){
    var ps=state_();
    var days=[0,1,2,3,4,5,6].map(function(i){return addDays(weekStart,i);});
    var shiftsByCode={};
    (State.data.shifts||[]).forEach(function(s){
      shiftsByCode[String(s.Shift||'').trim().toUpperCase()]=s;
    });

    var head=days.map(function(d){
      return '<th class="'+(isWeekend(d)?'paper-weekend':'')+'">'+dayHeader_(d)+'</th>';
    }).join('');

    var body=shiftOrder_().map(function(code){
      var s=shiftsByCode[code]||{};
      var time='';
      if(ps.showTimes && (s.Start||s.End)){
        time='<small>'+esc(s.Start||'')+'–'+esc(s.End||'')+'</small>';
      }
      var meaning=s.Meaning?'<small>'+esc(s.Meaning)+'</small>':'';
      var cells=days.map(function(d){
        var dk=dateKey(d);
        var outside=dk<ps.start||dk>ps.end;
        return '<td class="'+(isWeekend(d)?'paper-weekend ':'')+(outside?'paper-outside':'')+'">'+
          (outside?'<span class="paper-empty">—</span>':cell_(dk,code))+
          '</td>';
      }).join('');
      return '<tr><th class="paper-shift"><strong>'+esc(code)+'</strong>'+time+meaning+'</th>'+cells+'</tr>';
    }).join('');

    return '<section class="paper-week">'+
      '<div class="paper-week-title"><div><strong>Week '+(index+1)+'</strong><span>'+
      esc(dateKey(days[0]))+' through '+esc(dateKey(days[6]))+
      '</span></div></div>'+
      '<table class="paper-table '+(ps.compact?'compact':'')+'">'+
      '<thead><tr><th class="paper-shift-head">Shift</th>'+head+'</tr></thead>'+
      '<tbody>'+(body||'<tr><td colspan="8" class="paper-empty-state">No scheduled shifts in this period.</td></tr>')+'</tbody>'+
      '</table></section>';
  }

  window.renderPaperSchedule=function(){
    var ps=state_();
    var start=localDate_(ps.start),end=localDate_(ps.end);
    if(!start||!end||end<start){
      var def=defaults_();
      ps.start=def.start; ps.end=def.end;
      start=localDate_(ps.start); end=localDate_(ps.end);
    }

    var firstWeek=addDays(start,-start.getDay()),weeks=[];
    for(var w=new Date(firstWeek),guard=0;w<=end&&guard<30;w=addDays(w,7),guard++){
      weeks.push(new Date(w));
    }

    var selected=rows_().filter(function(r){
      var dk=String(r.Date||'').slice(0,10);
      return dk>=ps.start&&dk<=ps.end;
    });
    var filled=selected.filter(function(r){
      var n=String(r['Assigned Pharmacist']||'').toUpperCase();
      return String(r.Status||'').toUpperCase()!=='UNFILLED'&&n!=='UNFILLED'&&n!=='';
    }).length;
    var unfilled=selected.length-filled;

    $('content').innerHTML=
      '<div class="paper-page-shell">'+
        '<div class="paper-screen-head">'+
          '<div><div class="clinical-eyebrow">PRINTABLE STAFFING VIEW</div><h1>Paper Schedule</h1>'+
          '<p>Weekly landscape schedule designed for printing, posting, or saving as a PDF.</p></div>'+
          '<div class="toolbar">'+
            '<button type="button" class="btn btn-secondary" onclick="paperScheduleShiftRange_(-7)">← Previous week</button>'+
            '<button type="button" class="btn btn-secondary" onclick="paperScheduleShiftRange_(7)">Next week →</button>'+
            '<button type="button" class="btn btn-primary" onclick="printPaperSchedule_()">Print / Save PDF</button>'+
          '</div>'+
        '</div>'+
        '<div class="card paper-controls">'+
          '<div class="toolbar">'+
            '<div class="field"><label>Start Date</label><input id="paperStart" type="date" class="control" value="'+attr(ps.start)+'" onchange="paperScheduleDateChanged_()"></div>'+
            '<div class="field"><label>End Date</label><input id="paperEnd" type="date" class="control" value="'+attr(ps.end)+'" onchange="paperScheduleDateChanged_()"></div>'+
            '<label class="control paper-check"><input type="checkbox" '+(ps.showUnfilled?'checked':'')+' onchange="State.paperSchedule.showUnfilled=this.checked;renderPaperSchedule()"> Show UNFILLED</label>'+
            '<label class="control paper-check"><input type="checkbox" '+(ps.showTimes?'checked':'')+' onchange="State.paperSchedule.showTimes=this.checked;renderPaperSchedule()"> Show shift times</label>'+
            '<label class="control paper-check"><input type="checkbox" '+(ps.compact?'checked':'')+' onchange="State.paperSchedule.compact=this.checked;renderPaperSchedule()"> Compact print</label>'+
          '</div>'+
          '<div class="paper-summary">'+
            '<span><b>'+selected.length+'</b> schedule rows</span>'+
            '<span><b>'+filled+'</b> filled</span>'+
            '<span class="'+(unfilled?'paper-summary-alert':'')+'"><b>'+unfilled+'</b> unfilled</span>'+
            '<span><b>'+weeks.length+'</b> week'+(weeks.length===1?'':'s')+'</span>'+
          '</div>'+
        '</div>'+
        '<div id="paperPrintArea" class="paper-print-area">'+
          '<div class="paper-document-title"><div><h2>Pharmacy Staff Schedule</h2><p>'+
            esc(ps.start)+' through '+esc(ps.end)+
          '</p></div><div class="paper-doc-brand">NeoChrono</div></div>'+
          weeks.map(function(w,i){return weekBlock_(w,i);}).join('')+
          '<div class="paper-legend"><span><b>L</b> Locked</span><span><b>M</b> Manual assignment</span>'+
          (ps.showUnfilled?'<span><b>UNFILLED</b> Open coverage</span>':'')+
          '</div>'+
        '</div>'+
      '</div>';
  };

  window.paperScheduleDateChanged_=function(){
    var ps=state_();
    var s=$('paperStart')?$('paperStart').value:ps.start;
    var e=$('paperEnd')?$('paperEnd').value:ps.end;
    if(!s||!e){toast('Choose both a start and end date.','error');return;}
    if(s>e){toast('End Date must be on or after Start Date.','error');return;}
    var sd=localDate_(s),ed=localDate_(e);
    if(sd&&ed&&Math.round((ed-sd)/86400000)>90){
      toast('Paper Schedule is limited to 91 days at one time.','error');return;
    }
    ps.start=s;ps.end=e;renderPaperSchedule();
  };

  window.paperScheduleShiftRange_=function(days){
    var ps=state_(),s=localDate_(ps.start),e=localDate_(ps.end);
    if(!s||!e)return;
    ps.start=dateKey(addDays(s,days));
    ps.end=dateKey(addDays(e,days));
    renderPaperSchedule();
  };

  window.printPaperSchedule_=function(){
    var ps=state_(),oldTitle=document.title;
    document.title='Pharmacy Schedule '+ps.start+' to '+ps.end;
    document.body.classList.add('printing-paper-schedule');
    window.print();
    setTimeout(function(){
      document.body.classList.remove('printing-paper-schedule');
      document.title=oldTitle;
    },300);
  };
})();
