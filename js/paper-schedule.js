
(function(){
  'use strict';

  function paperMonthDates_(month){
    var arr=[];
    var last=new Date(month.getFullYear(),month.getMonth()+1,0).getDate();
    for(var i=1;i<=last;i++)arr.push(new Date(month.getFullYear(),month.getMonth(),i));
    return arr;
  }

  function paperApprovedPto_(username,name,dk){
    return (State.data.requests||[]).some(function(r){
      if(String(r['Record Type']||'').toUpperCase()!=='PTO')return false;
      if(String(r.Status||'').toUpperCase()!=='APPROVED')return false;
      var sameUser=
        String(r.Username==null?'':r.Username)===String(username==null?'':username) ||
        String(r.Pharmacist||'')===String(name||'');
      if(!sameUser)return false;
      var start=dateKey(r['Start Date']||r['End Date']||r.Date);
      var end=dateKey(r['End Date']||r['Start Date']||r.Date);
      return !!start&&!!end&&dk>=start&&dk<=end;
    });
  }

  function paperWeekendGroupForDate_(dk){
    var rows=(State.data.schedule||[]).filter(function(r){
      return r.Date===dk && yes(r.Weekend) && String(r['Weekend Group']||'').trim();
    });
    return rows.length?String(rows[0]['Weekend Group']).trim():'';
  }

  function paperDateIsHoliday_(dk){
    return (State.data.schedule||[]).some(function(r){
      return r.Date===dk && yes(r.Holiday);
    });
  }

  function paperVisibleUsers_(){
    var f=State.calendarFilters||{};
    var selected=selectedCalendarEmployeeKeys();
    var users=(State.data.users||[]).filter(function(u){
      return yes(u.Active) && String(u.Role||'Pharmacist').toLowerCase()!=='administrator';
    });

    if(selected.length){
      users=users.filter(function(u){
        return selected.includes(String(u.Username==null?'':u.Username));
      });
    }
    if(f.resident){
      users=users.filter(function(u){return String(u.Resident||'')===f.resident;});
    }
    if(f.preceptor){
      users=users.filter(function(u){return String(u.Preceptor||'')===f.preceptor;});
    }

    return users.sort(function(a,b){
      return String(a['Pharmacist Name']||'').localeCompare(String(b['Pharmacist Name']||''));
    });
  }

  function paperCellData_(user,dk){
    var rows=calendarRowsForDate(dk).filter(function(r){
      return r.Status!=='UNFILLED' &&
        String(r.Username==null?'':r.Username)===String(user.Username==null?'':user.Username);
    });

    if(rows.length){
      var codes=rows.map(function(r){return String(r.Shift||'');}).filter(Boolean);
      var r=rows[0];
      var cls='has-assignment';

      if(String(r['Shift Type']||'')==='Evening')cls+=' paper-evening';
      if(String(r['Shift Type']||'')==='Night')cls+=' paper-night';
      if(r['Coverage For Pharmacist'])cls+=' paper-coverage';
      if(yes(r.Manual) || String(r.Status||'').toUpperCase()==='MANUAL' || String(r.Status||'').toUpperCase()==='LOCKED'){
        cls+=' paper-manual';
      }

      var title=rows.map(function(x){
        return String(x.Shift||'')+' — '+String(x['Assigned Pharmacist']||'')+
          (x['Coverage For Pharmacist']?' | covering '+x['Coverage For Pharmacist']:'')+
          (x.Warning?' | '+x.Warning:'');
      }).join('\n');

      return {
        text:codes.join('/'),
        cls:cls,
        title:title,
        id:r['Assignment ID']||''
      };
    }

    var filters=State.calendarFilters||{};
    var filtersActive=!!(
      filters.status || filters.shift || filters.type ||
      filters.skill || filters.weekendGroup
    );

    if(!filtersActive && paperApprovedPto_(
      user.Username,
      user['Pharmacist Name'],
      dk
    )){
      return {text:'P',cls:'paper-pto',title:'Approved PTO',id:''};
    }

    return {
      text:filtersActive?'':'X',
      cls:'paper-off',
      title:filtersActive?'No matching assignment':'OFF',
      id:''
    };
  }

  function paperUnfilledForDate_(dk){
    return calendarRowsForDate(dk).filter(function(r){
      return r.Status==='UNFILLED';
    });
  }

  function paperUnfilledCell_(rows){
    if(!rows.length)return '';
    return rows.map(function(r){
      return '<span class="paper-unfilled-chip" title="'+
        attr(String(r.Shift||'')+(r.Warning?' — '+r.Warning:''))+
        '" onclick="openAssignmentModal(\''+attr(r['Assignment ID']||'')+'\')">'+
        esc(r.Shift||'UNFILLED')+
        '</span>';
    }).join('');
  }

  window.renderPaperSchedule=function(){
    var m=State.month;
    var d=State.data;
    var filt=State.calendarFilters;
    var dates=paperMonthDates_(m);
    var users=paperVisibleUsers_();

    var html=
      '<div class="paper-page-shell">'+
      '<div class="calendar-head paper-page-head">'+
        '<div>'+
          '<div class="calendar-title">'+esc(monthTitle(m))+' — Paper Schedule</div>'+
          '<div class="muted small">Excel-style view: pharmacists down the left, dates across the top. X = OFF, P = approved PTO.</div>'+
        '</div>'+
        '<div class="toolbar">'+
          '<button class="btn btn-secondary btn-sm" onclick="paperScheduleChangeMonth_(-1)">← Previous</button>'+
          '<button class="btn btn-secondary btn-sm" onclick="State.month=firstOfMonth(new Date());renderPaperSchedule()">Today</button>'+
          '<button class="btn btn-secondary btn-sm" onclick="paperScheduleChangeMonth_(1)">Next →</button>'+
          '<button class="btn btn-primary btn-sm" onclick="printPaperSchedule_()">Print / Save PDF</button>'+
          (d.isAdmin?'<button class="btn btn-primary btn-sm" onclick="openAssignmentModal(null)">+ Manual assignment</button>':'')+
        '</div>'+
      '</div>';

    html+=
      '<div class="paper-filter-card">'+
        '<div class="paper-filter-row">'+
          selectFilter('weekendGroup','Weekend Group',['A','B','C'],filt.weekendGroup)+
          selectFilter('resident','Resident',['Yes','No'],filt.resident)+
          selectFilter('preceptor','Preceptor',['Yes','No'],filt.preceptor)+
          '<button class="btn btn-ghost btn-sm paper-clear-btn" onclick="clearPaperScheduleFilters_()">Clear filters</button>'+
        '</div>'+
      '</div>';


    html+=
      '<div class="paper-legend">'+
        '<span class="key"><span class="swatch" style="background:#fff"></span>X = OFF</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-pto)"></span>P = PTO</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-evening)"></span>Evening</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-night)"></span>Night</span>'+
        '<span class="key"><span class="swatch" style="background:var(--paper-coverage)"></span>OFF-day coverage</span>'+
        '<span class="key"><span class="swatch" style="background:#ffe6e6"></span>Unfilled</span>'+
      '</div>';

    html+='<div id="paperPrintArea" class="paper-shell"><table class="paper-table">'+
      '<colgroup><col class="paper-name-column">'+
      dates.map(function(){return '<col class="paper-date-column">';}).join('')+
      '</colgroup><thead>';

    html+=
      '<tr class="paper-title-row">'+
        '<th class="name-col">SCHEDULE PERIOD</th>'+
        '<th colspan="'+dates.length+'">'+esc(monthTitle(m))+'</th>'+
      '</tr>';

    html+=
      '<tr class="paper-group-row">'+
        '<th class="name-col">Weekend Group</th>'+
        dates.map(function(dt){
          var dk=dateKey(dt);
          var g=paperWeekendGroupForDate_(dk);
          return '<th class="date-col '+(isWeekend(dt)?'weekend-head ':'')+
            (paperDateIsHoliday_(dk)?'holiday-col':'')+'">'+esc(g)+'</th>';
        }).join('')+
      '</tr>';

    html+=
      '<tr class="paper-date-row">'+
        '<th class="name-col">Pharmacist</th>'+
        dates.map(function(dt){
          var dk=dateKey(dt);
          return '<th class="date-col '+(paperDateIsHoliday_(dk)?'holiday-col':'')+'">'+
            (dt.getMonth()+1)+'/'+dt.getDate()+
          '</th>';
        }).join('')+
      '</tr>';

    html+=
      '<tr class="paper-day-row">'+
        '<th class="name-col">Name</th>'+
        dates.map(function(dt){
          var dk=dateKey(dt);
          return '<th class="date-col '+(paperDateIsHoliday_(dk)?'holiday-col':'')+'">'+
            ['SUN','MON','TUE','WED','THU','FRI','SAT'][dt.getDay()]+
          '</th>';
        }).join('')+
      '</tr></thead><tbody>';

    if(State.calendarFilters.status!=='UNFILLED'){
      users.forEach(function(user){
        html+='<tr><td class="name-col">'+esc(user['Pharmacist Name']||'')+'</td>';

        dates.forEach(function(dt){
          var dk=dateKey(dt);
          var cell=paperCellData_(user,dk);
          var holiday=paperDateIsHoliday_(dk);
          var weekend=isWeekend(dt);
          var onclick=cell.id?
            ' onclick="openAssignmentModal(\''+attr(cell.id)+'\')"':'';

          html+='<td class="paper-cell date-col '+
            (weekend?'weekend ':'')+
            (holiday?'holiday-col ':'')+
            cell.cls+
            '" title="'+attr(cell.title||'')+'"'+onclick+'>'+
            esc(cell.text)+
          '</td>';
        });

        html+='</tr>';
      });
    }

    var selected=selectedCalendarEmployeeKeys();
    var showUnfilled=
      State.calendarFilters.status==='UNFILLED' ||
      (!selected.length && State.calendarFilters.status!=='FILLED') ||
      (selected.length && State.calendarFilters.showUnfilledWithSelected && State.calendarFilters.status!=='FILLED');

    if(showUnfilled){
      html+='<tr class="unfilled-row"><td class="name-col">UNFILLED</td>';

      dates.forEach(function(dt){
        var dk=dateKey(dt);
        var rows=paperUnfilledForDate_(dk);
        html+='<td class="paper-cell paper-unfilled date-col '+
          (isWeekend(dt)?'weekend ':'')+
          (paperDateIsHoliday_(dk)?'holiday-col':'')+
          '">'+paperUnfilledCell_(rows)+'</td>';
      });

      html+='</tr>';
    }

    html+=
      '</tbody></table></div>'+
      '<div class="paper-note">Tip: click any shift code to review or edit that assignment. The left pharmacist column stays fixed while you scroll horizontally, like a printed Excel schedule.</div>'+
      '</div>';

    $('content').innerHTML=html;
    bindPaperScheduleFilters_();
  };

  function bindPaperScheduleFilters_(){
    document.querySelectorAll('.cal-filter').forEach(function(el){
      el.onchange=function(){
        State.calendarFilters[el.dataset.filter]=el.value;
        renderPaperSchedule();
      };
    });

    document.querySelectorAll('.cal-employee-filter').forEach(function(el){
      el.onchange=function(){
        var slots=normalizedEmployeeFilterSlots();
        slots[num(el.dataset.employeeSlot,0)]=String(el.value||'');
        State.calendarFilters.employees=slots;
        renderPaperSchedule();
      };
    });
  }

  window.togglePaperShowUnfilled_=function(v){
    State.calendarFilters.showUnfilledWithSelected=!!v;
    renderPaperSchedule();
  };

  window.clearPaperScheduleFilters_=function(){
    State.calendarFilters.employees=['','','','','',''];
    State.calendarFilters.showUnfilledWithSelected=false;
    ['status','shift','type','skill','weekendGroup','resident','preceptor'].forEach(function(k){
      State.calendarFilters[k]='';
    });
    renderPaperSchedule();
  };

  window.paperScheduleChangeMonth_=function(n){
    State.month=addMonths(State.month,n);
    renderPaperSchedule();
  };

  window.printPaperSchedule_=function(){
    var oldTitle=document.title;
    document.title='Pharmacy Paper Schedule - '+monthTitle(State.month);
    document.body.classList.add('printing-paper-schedule');
    window.print();
    setTimeout(function(){
      document.body.classList.remove('printing-paper-schedule');
      document.title=oldTitle;
    },300);
  };
})();
