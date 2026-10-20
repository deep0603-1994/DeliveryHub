(() => {
  const headers = ['RequestID','Request Type','Subject','Request Status','Location','Location Code','Zone','Service Category','Category','Sub Category','Priority','Template','Status','Group','Technician','Engineer Name','Aging','Created Time','Last Updated Time','Requester','Requester Email','Requester Mobile Number','Contact No.'];
  const dimensions = ['Request Type','Status','Location','Zone','Category','Priority','Group','Engineer Name','Ageing Bucket'];
  const optionalHeaders = ['Department','Created Date','Closure Date'];
  const gridFirst = ['ID','Subject','Department','Group','Stage','Status','Subgroup','Created Date','Created Time','Requestor','Age in Days','Ageing Bucket'];
  const ageingBuckets = ['0\u20135 days','6\u201310 days','11\u201320 days','More than 20 days'];
  const state = { records: [], loaded: false, loading: false, error: '', user: '', query: '', filters: {}, page: 1, size: typeof DELIVERY_PAGE_SIZE === 'number' ? DELIVERY_PAGE_SIZE : 25, sort: 'ID', direction: 1, expanded: false, view: 'grid', summary: 'Request Type', busy: false, openCards: new Set(), preview: null };
  const root = () => document.getElementById('workspace-flat-page');
  const active = () => currentWorkspaceView === 'support';
  const editable = () => Boolean(currentUser && ['Admin','Level 1'].includes(currentUser.role));
  const escape = value => escapeHtml(String(value ?? ''));
  const fieldKey = name => String(name).trim().toLowerCase().replace(/[^a-z0-9]/g,'');
  function sourceField(record,names) {
    for(const name of names){
      const key=Object.keys(record.fields||{}).find(key=>fieldKey(key)===fieldKey(name));
      if(key!==undefined&&String(record.fields[key]??'').trim()!=='')return {key,text:String(record.fields[key])};
    }
    return {key:names[0],text:''};
  }
  function calendarDate(year,month,day,time='') {
    const instant=new Date(Date.UTC(year,month-1,day));
    if(instant.getUTCFullYear()!==year||instant.getUTCMonth()!==month-1||instant.getUTCDate()!==day)return null;
    return {day:instant.getTime()/86400000,date:`${String(year).padStart(4,'0')}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`,time};
  }
  function ticketDate(record,field) {
    const text=field.text.trim();if(!text)return null;
    const cell=(record.source_cells?.cells||[]).find(cell=>fieldKey(cell.header)===fieldKey(field.key));
    if(cell?.t==='n'&&typeof cell.v==='number'&&(text===String(cell.w??'').trim()||text===String(cell.v))&&typeof XLSX!=='undefined'){
      const parts=XLSX.SSF.parse_date_code(cell.v,{date1904:Boolean(record.source_cells?.date1904)});
      if(parts){const time=/[hs]:?|:m|am\/pm/i.test(cell.z||'')||/:/.test(cell.w||'')?[parts.H,parts.M,Math.floor(parts.S)].map(part=>String(part).padStart(2,'0')).join(':'):'';return calendarDate(parts.y,parts.m,parts.d,time);}
    }
    const clock=text.match(/(?:T|\s)(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?/i);
    let time='';
    if(clock){let hour=Number(clock[1]);if(clock[4])hour=hour%12+(/^pm$/i.test(clock[4])?12:0);if(hour>23||Number(clock[2])>59||Number(clock[3]||0)>59)return null;time=[hour,Number(clock[2]),Number(clock[3]||0)].map(part=>String(part).padStart(2,'0')).join(':');}
    const iso=text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\b|T)/);
    if(iso)return calendarDate(Number(iso[1]),Number(iso[2]),Number(iso[3]),time);
    const numeric=text.match(/^(\d{1,2})([-/])(\d{1,2})\2(\d{2}|\d{4})(?:\b|\s)/);
    if(numeric){let year=Number(numeric[4]);if(year<100)year+=year<70?2000:1900;const first=Number(numeric[1]),second=Number(numeric[3]),format=String(cell?.z||'').toLowerCase();const sourceMonthFirst=format.indexOf('m')>=0&&format.indexOf('d')>=0&&format.indexOf('m')<format.indexOf('d');const monthFirst=first<=12&&(second>12||sourceMonthFirst);return calendarDate(year,monthFirst?first:second,monthFirst?second:first,time);}
    const named=text.match(/^(\d{1,2})[ -]([a-z]{3,9})[ ,-]+(\d{4})/i);
    if(named){const month=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(named[2].slice(0,3).toLowerCase())+1;return month?calendarDate(Number(named[3]),month,Number(named[1]),time):null;}
    return null;
  }
  function deriveTicket(record,referenceDate) {
    const createdField=sourceField(record,['Created Date','Created Time','Created On','Creation Date']);
    const closureField=sourceField(record,['Closure Date','Closed Date','Closure Time','Closed Time','Closed On']);
    const created=ticketDate(record,createdField),closure=ticketDate(record,closureField),stage=closureField.text.trim()?'Closed':'Open';
    const now=new Date(),today=referenceDate|| (typeof localToday==='function'?localToday():`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`);
    const end=stage==='Closed'?closure:ticketDate({fields:{}},{key:'Today',text:today});
    const age=created&&end?Math.abs(end.day-created.day):null;
    const bucket=age===null?'':ageingBuckets[age<=5?0:age<=10?1:age<=20?2:3];
    const separateTime=sourceField(record,['Created Time']);
    const explicitTime=separateTime.text.trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i);
    const time=created?.time || ticketDate(record,separateTime)?.time || (explicitTime?ticketDate({fields:{}},{key:'Time',text:`2000-01-01 ${separateTime.text}`})?.time:'') || '';
    return {stage,age,bucket,createdDate:created?.date||'',createdTime:time,flag:age===null?'Check dates':bucket};
  }
  function value(record,field) {
    if(['Stage','Age in Days','Ageing Bucket','Created Date','Created Time','Ageing Flag'].includes(field)){
      const derived=deriveTicket(record);
      return String(({Stage:derived.stage,'Age in Days':derived.age??'','Ageing Bucket':derived.bucket,'Created Date':derived.createdDate,'Created Time':derived.createdTime,'Ageing Flag':derived.flag})[field]);
    }
    if(field==='ID')return record.request_id;
    if(field==='Subgroup')return sourceField(record,['Subgroup','Sub Group','Sub Category']).text;
    if(field==='Requestor')return sourceField(record,['Requestor','Requester']).text;
    return String(record.fields?.[field]??'');
  }
  async function load(force = false) {
    const user = String(currentUser?.email || currentUser?.id || '');
    if (state.user !== user) { state.records = []; state.loaded = false; state.user = user; state.query = ''; state.filters = {}; state.preview = null; }
    if (state.loading || state.loaded && !force) return;
    state.loading = true; state.error = '';
    try {
      const records = [];
      for (let start = 0; ; start += 500) {
        const { data, error } = await supabaseClient.from('support_tickets').select('*').order('request_id').range(start, start + 499);
        if (error) throw error;
        records.push(...data);
        if (data.length < 500) break;
      }
      if (state.user !== user) return;
      state.records = records; state.loaded = true;
    } catch (error) { state.error = error.message; }
    finally { state.loading = false; if (active()) render(); }
  }
  const icon = name => `<i data-lucide="${name}" class="w-4 h-4" aria-hidden="true"></i>`;
  const command = (action, name, label, extra = '') => `<button type="button" data-support-action="${action}" title="${label}" aria-label="${label}" class="support-icon-button border border-slate-300 rounded-lg bg-white text-slate-600 hover:bg-slate-50" ${extra}>${icon(name)}</button>`;
  const status = record => value(record, 'Status') || value(record, 'Request Status');
  const closed = record => deriveTicket(record).stage==='Closed';
  function agingDays(text) {
    const input = String(text).trim();
    if (/^\d+(?:\.\d+)?$/.test(input)) return Number(input);
    let days = 0, found = false;
    for (const match of input.matchAll(/(\d+(?:\.\d+)?)\s*(days?|d\b|hours?|hrs?|h\b|minutes?|mins?|m\b)/gi)) {
      days += Number(match[1]) / (/^d/i.test(match[2]) ? 1 : /^h/i.test(match[2]) ? 24 : 1440); found = true;
    }
    return found ? days : null;
  }
  function filtered() {
    const query = state.query.trim().toLowerCase();
    return state.records.filter(record => (!query || [...Object.values(record.fields),...gridFirst.map(field=>value(record,field))].some(item => String(item ?? '').toLowerCase().includes(query))) && Object.entries(state.filters).every(([field, selected]) => window.trackerMultiFilters ? trackerFilterMatches(`support:${field}`,value(record,field),selected) : !selected || value(record, field) === selected))
      .sort((left, right) => state.direction * (state.sort === 'Age in Days'
        ? (deriveTicket(left).age ?? -1) - (deriveTicket(right).age ?? -1)
        : state.sort==='Ageing Bucket'?ageingBuckets.indexOf(value(left,state.sort))-ageingBuckets.indexOf(value(right,state.sort))
        : value(left, state.sort).localeCompare(value(right, state.sort), undefined, { numeric: true, sensitivity: 'base' })) || left.request_id.localeCompare(right.request_id));
  }
  function summaries(records, dimension) {
    const groups = new Map();
    for (const record of records) {
      const name = value(record, dimension) || '(Not specified)';
      const entry = groups.get(name) || { name, total: 0, open: 0, closed: 0, ages: [] };
      entry.total++; closed(record) ? entry.closed++ : entry.open++;
      const days = deriveTicket(record).age; if (days !== null) entry.ages.push(days);
      groups.set(name, entry);
    }
    return [...groups.values()].sort((left, right) => right.total - left.total || left.name.localeCompare(right.name));
  }
  function filterControl(field) {
    const options = [...new Set(state.records.map(record => value(record, field)).filter(Boolean))].sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
    return `<label class="support-filter-control"><span class="sr-only">${escape(field)}</span><select data-support-filter="${escape(field)}" aria-label="${escape(field)}" class="w-full border border-slate-300 bg-slate-50 rounded-lg px-3 py-2 text-xs"><option value="">All ${escape(field)}</option>${options.map(option => `<option value="${escape(option)}" ${state.filters[field] === option ? 'selected' : ''}>${escape(option)}</option>`).join('')}</select></label>`;
  }
  function allHeaders(records = state.records) { return [...new Set([...headers,...optionalHeaders, ...records.flatMap(record => Object.keys(record.fields))])]; }
  function gridColumns(records=state.records) {
    const mapped=new Set([...gridFirst,'RequestID','Requester','Sub Category','Sub Group','Aging','Ageing','Ticket Age','Ageing Flag','Aging Flag'].map(fieldKey));
    return [...gridFirst,...allHeaders(records).filter(field=>!mapped.has(fieldKey(field)))];
  }
  function ageFlag(record) {
    const derived=deriveTicket(record),index=ageingBuckets.indexOf(derived.bucket);
    const color=['text-emerald-700','text-amber-700','text-orange-700','text-rose-700'][index]||'text-slate-500';
    return `<span class="inline-flex items-center gap-1.5 font-semibold ${color}" title="${escape(derived.flag)}" aria-label="${escape(derived.age===null?'Check Created Date or Closure Date':`${derived.age} days, ${derived.flag}`)}">${icon('flag')}<span>${derived.age??'-'}</span></span>`;
  }
  function badge(text, priority = false) {
    const color = priority && /critical|urgent|high/i.test(text) ? 'bg-rose-100 text-rose-700' : /closed|resolved|completed/i.test(text) ? 'bg-emerald-100 text-emerald-700' : 'bg-indigo-100 text-indigo-700';
    return `<span class="inline-flex px-2 py-1 text-[10px] font-semibold rounded border border-current/10 ${color}">${escape(text || 'Not specified')}</span>`;
  }
  function ticketCard(record) {
    const index = state.records.indexOf(record), open = state.openCards.has(record.request_id);
    return `<article class="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden"><div class="px-4 py-4 border-b border-slate-100 flex items-start justify-between gap-3"><div class="flex gap-3 min-w-0">${command('expand-card', open ? 'chevron-down' : 'chevron-right', open ? 'Collapse ticket' : 'Expand ticket', `data-index="${index}" aria-expanded="${open}"`)}<div class="min-w-0"><div class="flex flex-wrap items-center gap-2 mb-1"><span class="text-xs font-bold text-indigo-700">${escape(record.request_id)}</span>${badge(value(record,'Stage'))}${badge(status(record))}${badge(value(record, 'Priority'), true)}</div><h2 class="text-sm font-bold text-slate-800 break-words">${escape(value(record, 'Subject'))}</h2><div class="text-xs text-slate-500 mt-1 break-words">${escape(value(record, 'Request Type'))} | ${escape(value(record, 'Location'))} | ${escape(value(record, 'Zone'))}</div></div></div><div class="flex shrink-0 gap-1">${command('view', 'eye', 'View ticket', `data-index="${index}"`)}${editable() ? command('edit', 'pencil', 'Edit ticket', `data-index="${index}"`) : ''}</div></div><div class="grid grid-cols-2 md:grid-cols-4 gap-4 px-4 py-3 text-xs">${['Group','Engineer Name','Age in Days','Ageing Bucket'].map(field => `<div class="min-w-0"><div class="text-[10px] uppercase font-semibold text-slate-400">${escape(field)}</div><div class="mt-1 text-slate-700 break-words">${field==='Age in Days'?ageFlag(record):escape(value(record, field) || '-')}</div></div>`).join('')}</div>${open ? `<dl class="support-details border-t border-slate-100 p-4">${gridColumns([record]).map(field => `<div><dt class="text-[10px] uppercase text-slate-400 font-semibold">${escape(field)}</dt><dd class="mt-1 text-xs text-slate-700 whitespace-pre-wrap break-words">${field==='Age in Days'?ageFlag(record):escape(value(record, field) || '-')}</dd></div>`).join('')}</dl>` : ''}</article>`;
  }
  function ticketGrid(records) {
    const columns = gridColumns();
    return `<div class="bg-white rounded-xl border shadow-sm overflow-hidden"><div class="overflow-x-auto custom-scrollbar"><table class="min-w-full text-xs"><thead class="bg-slate-100"><tr>${columns.map(field => `<th class="px-3 py-3 text-left whitespace-nowrap" aria-sort="${state.sort === field ? state.direction === 1 ? 'ascending' : 'descending' : 'none'}"><button data-support-action="sort" data-field="${escape(field)}" class="inline-flex items-center gap-1 font-semibold">${escape(field)}${icon(state.sort === field ? state.direction === 1 ? 'arrow-up' : 'arrow-down' : 'arrow-up-down')}</button></th>`).join('')}<th class="px-3 py-3 text-left">Actions</th></tr></thead><tbody>${records.map(record => `<tr class="border-t hover:bg-slate-50">${columns.map(field => `<td class="px-3 py-3 ${field === 'Subject' ? 'min-w-[240px] max-w-[400px] whitespace-normal break-words' : 'whitespace-nowrap'}">${field === 'Status' ? badge(status(record)) : field==='Stage'?badge(value(record,field)):field==='Age in Days'?ageFlag(record):escape(value(record, field))}</td>`).join('')}<td class="px-3 py-3"><div class="flex gap-1">${command('view', 'eye', 'View ticket', `data-index="${state.records.indexOf(record)}"`)}${editable() ? command('edit', 'pencil', 'Edit ticket', `data-index="${state.records.indexOf(record)}"`) : ''}</div></td></tr>`).join('')}</tbody></table></div></div>`;
  }
  function renderResults() {
    if (!active()) return;
    const records = filtered(), ages = records.map(record => deriveTicket(record).age).filter(days => days !== null);
    const metrics = [['Total Tickets', records.length, 'text-slate-900'],['Open Tickets', records.filter(record => !closed(record)).length, 'text-indigo-600'],['Closed Tickets', records.filter(closed).length, 'text-emerald-600'],['High / Critical', records.filter(record => /high|critical|urgent/i.test(value(record, 'Priority'))).length, 'text-rose-600'],['Unassigned', records.filter(record => !value(record, 'Engineer Name') && !value(record, 'Technician')).length, 'text-amber-600'],['Average Age (Days)', ages.length ? (ages.reduce((total, days) => total + days, 0) / ages.length).toFixed(1) : '-', 'text-sky-600']];
    document.getElementById('support-metrics').innerHTML = metrics.map(([label, count, color]) => `<div class="bg-white p-4 rounded-xl border shadow-sm min-w-0"><div class="text-[10px] uppercase text-slate-500 font-semibold">${label}</div><div class="text-2xl font-bold ${color}">${count}</div></div>`).join('');
    const summaryGroups=state.view === 'summary' ? summaries(records,state.summary) : null;
    const totalCount=summaryGroups ? summaryGroups.length : records.length;
    const pages = Math.max(1, Math.ceil(totalCount / state.size)); state.page = Math.min(state.page, pages);
    const pageRecords = records.slice((state.page - 1) * state.size, state.page * state.size), host = document.getElementById('support-results');
    if (state.view === 'summary') {
      const groups = summaryGroups.slice((state.page - 1) * state.size,state.page * state.size);
      host.innerHTML = `<section class="space-y-3"><div class="flex flex-wrap items-center justify-between gap-2"><h2 class="text-sm font-bold text-slate-800">${escape(state.summary)} Summary</h2><select data-support-summary class="border rounded-lg px-3 py-2 text-xs" aria-label="Summary dimension">${dimensions.map(field => `<option ${field === state.summary ? 'selected' : ''}>${field}</option>`).join('')}</select></div><div class="overflow-x-auto bg-white border rounded-xl"><table class="w-full text-xs"><thead class="bg-slate-100 text-left"><tr>${[state.summary,'Total','Open','Closed','Average Aging (Days)'].map(field => `<th class="px-4 py-3">${escape(field)}</th>`).join('')}</tr></thead><tbody>${groups.map(group => `<tr class="border-t"><td class="px-4 py-3 whitespace-pre-wrap break-words">${escape(group.name)}</td><td class="px-4 py-3">${group.total}</td><td class="px-4 py-3">${group.open}</td><td class="px-4 py-3">${group.closed}</td><td class="px-4 py-3">${group.ages.length ? (group.ages.reduce((total, days) => total + days, 0) / group.ages.length).toFixed(1) : '-'}</td></tr>`).join('')}</tbody></table></div></section>`;
    } else host.innerHTML = pageRecords.length ? state.view === 'grid' ? ticketGrid(pageRecords) : pageRecords.map(ticketCard).join('') : `<div class="bg-white border rounded-xl p-12 text-center text-sm text-slate-500">${state.loading ? 'Loading support tickets...' : 'No tickets match the selected filters.'}</div>`;
    document.getElementById('support-pagination').innerHTML = `<div class="text-xs text-slate-500">${totalCount ? (state.page - 1) * state.size + 1 : 0}-${Math.min(state.page * state.size, totalCount)} of ${totalCount} ${summaryGroups ? 'groups' : 'tickets'}</div><div class="flex flex-wrap items-center gap-2"><select data-support-size aria-label="Tickets per page" class="text-xs border rounded-lg px-2 py-2">${[...new Set([10,25,50,100,state.size])].sort((left,right)=>left-right).map(size => `<option value="${size}" ${size === state.size ? 'selected' : ''}>${size} / page</option>`).join('')}</select>${command('previous', 'chevron-left', 'Previous page', state.page === 1 ? 'disabled' : '')}<span class="text-xs text-slate-600">${state.page} / ${pages}</span>${command('next', 'chevron-right', 'Next page', state.page === pages ? 'disabled' : '')}</div>`;
    const sortControl=document.querySelector('[data-support-sort]');if(sortControl)sortControl.value=state.sort;
    const direction=document.querySelector('[data-support-action="direction"]');if(direction)direction.innerHTML=icon(state.direction === 1 ? 'arrow-up' : 'arrow-down');
    const reset = document.querySelector('[data-support-action="reset"]'); if (reset) reset.disabled = !state.query && !Object.values(state.filters).some(Boolean);
    lucide.createIcons();
  }
  function render() {
    if (!active() || !root()) return;
    const user = String(currentUser?.email || currentUser?.id || '');
    if (state.user !== user) { state.user=user; state.records=[]; state.loaded=false; state.error=''; state.query=''; state.filters={}; state.openCards.clear(); state.preview=null; }
    document.getElementById('delivery-pagination-bottom')?.classList.add('hidden');
    const primary = ['Request Type','Status','Location','Priority'];
    root().innerHTML = `<div id="support-metrics" class="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-6 gap-3"></div><div id="support-filter-banner" class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm space-y-2"><div class="support-primary"><input id="support-search" type="search" value="${escape(state.query)}" placeholder="Search tickets..." aria-label="Search support tickets" class="w-full border border-slate-300 bg-slate-50 rounded-lg px-3 py-2 text-xs">${primary.map(filterControl).join('')}<div class="flex gap-2">${command('reset', 'rotate-ccw', 'Reset filters')}${command('filters', state.expanded ? 'chevron-up' : 'chevron-down', state.expanded ? 'Collapse filters' : 'Expand filters', `aria-expanded="${state.expanded}" aria-controls="support-secondary"`)}</div></div><div id="support-secondary" class="support-secondary" ${state.expanded ? '' : 'hidden'}>${[...new Set([...dimensions, 'Request Status','Service Category','Sub Category','Location Code','Technician'])].filter(field => !primary.includes(field)).map(filterControl).join('')}</div></div><div class="bg-white px-4 py-3 rounded-xl border border-slate-200 shadow-sm flex flex-wrap items-center justify-between gap-3"><div role="tablist" aria-label="Support ticket views" class="flex gap-1">${[['cards','layout-list','Ticket cards'],['grid','table-2','Ticket grid'],['summary','chart-no-axes-combined','Summary reporting']].map(([view,name,label]) => `<button type="button" role="tab" aria-selected="${state.view === view}" data-support-action="view-mode" data-view="${view}" title="${label}" aria-label="${label}" class="support-icon-button rounded-lg border ${state.view === view ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-white text-slate-600 border-slate-300'}">${icon(name)}</button>`).join('')}</div><div class="flex items-center gap-2"><select data-support-sort aria-label="Sort tickets by" class="text-xs border rounded-lg px-3 py-2">${headers.map(field => `<option ${field === state.sort ? 'selected' : ''}>${field}</option>`).join('')}</select>${command('direction', state.direction === 1 ? 'arrow-up' : 'arrow-down', 'Reverse sort direction')}${command('refresh', 'refresh-cw', 'Refresh tickets')}</div></div>${state.error ? `<div role="alert" class="border border-rose-200 bg-rose-50 text-rose-700 rounded-lg p-4 text-xs">${escape(state.error)}${/support_tickets|schema cache/i.test(state.error) ? '<br>Apply the Support Ticket Tracker database migration, then refresh.' : ''}</div>` : ''}<div id="support-results" class="space-y-4"></div><div id="support-pagination" class="flex flex-wrap items-center justify-between gap-3 bg-white border rounded-xl px-4 py-3"></div>`;
    renderResults();
    const sortSelect=root().querySelector('[data-support-sort]');
    sortSelect.replaceChildren(...gridColumns().map(field=>new Option(field,field,false,state.sort===field)));
    const tabular=root().querySelector('[data-view="grid"]');tabular.title='Tabular View';tabular.setAttribute('aria-label','Tabular View');
    document.getElementById('support-secondary').insertAdjacentHTML('afterbegin',filterControl('Stage'));
    if (!state.loaded && !state.loading && !state.error) load();
  }
  let previousFocus = null;
  function closeDialog() {
    const dialog = document.getElementById('support-ticket-dialog');
    if (dialog?.dataset.busy === 'true') return;
    dialog?.remove(); previousFocus?.isConnected && previousFocus.focus({ preventScroll: true });
  }
  function dialog(title, body) {
    closeDialog(); previousFocus = document.activeElement;
    const host = document.createElement('div'); host.id = 'support-ticket-dialog'; host.className = 'fixed inset-0 z-[120] bg-slate-900/60 flex items-center justify-center p-4';
    host.setAttribute('role','dialog'); host.setAttribute('aria-modal','true'); host.setAttribute('aria-labelledby','support-dialog-title');
    host.innerHTML = `<div class="bg-white rounded-xl shadow-2xl w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden"><header class="flex items-center justify-between gap-3 px-6 py-4 border-b bg-indigo-50"><h2 id="support-dialog-title" class="text-sm font-bold text-slate-800">${escape(title)}</h2>${command('close-dialog','x','Close dialog')}</header><div class="p-6 overflow-y-auto custom-scrollbar">${body}</div></div>`;
    document.body.append(host); lucide.createIcons(); host.querySelector('input:not([readonly]),button')?.focus();
  }
  function derivedPreview(record) {
    return `<div id="support-derived-preview" class="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-5 pb-4 border-b" aria-live="polite"><div><div class="text-[10px] uppercase text-slate-500">Stage</div>${badge(value(record,'Stage'))}</div><div><div class="text-[10px] uppercase text-slate-500">Age in Days</div>${ageFlag(record)}</div><div><div class="text-[10px] uppercase text-slate-500">Ageing Bucket</div><div class="text-xs font-semibold text-slate-700">${escape(value(record,'Ageing Bucket')||'Check dates')}</div></div></div>`;
  }
  function openTicket(index, mode) {
    if (mode !== 'view' && !editable()) return showToast('Permission denied.','error');
    const record = index === null ? null : state.records[index];
    if (index !== null && !record) return;
    const fields = allHeaders(record ? [record] : []).filter(field=>!['stage','ageindays','ageingbucket','ageingflag','agingflag'].includes(fieldKey(field))), readOnly = mode === 'view';
    dialog(readOnly ? `Ticket ${record.request_id}` : record ? `Edit Ticket ${record.request_id}` : 'New Support Ticket', `<form id="support-ticket-form" data-index="${index ?? ''}"><div class="support-details">${fields.map((field, position) => `<label class="min-w-0 text-xs font-semibold text-slate-600">${escape(field)}${['RequestID','Subject'].includes(field) ? ' *' : ''}${field === 'Subject' ? `<textarea data-ticket-field="${escape(field)}" rows="3" ${readOnly ? 'readonly' : 'required'} class="mt-1 w-full border rounded-lg px-3 py-2 font-normal whitespace-pre-wrap">${escape(record ? value(record,field) : '')}</textarea>` : `<input data-ticket-field="${escape(field)}" type="text" ${readOnly || record && field === 'RequestID' ? 'readonly' : ''} ${['RequestID','Subject'].includes(field) ? 'required' : ''} value="${escape(record ? value(record,field) : '')}" class="mt-1 w-full border rounded-lg px-3 py-2 font-normal">`}</label>`).join('')}</div>${record?.source_import_id ? `<div class="mt-4 text-xs text-slate-500">Source: ${escape(record.source_sheet || '')}, row ${record.source_row || '-'} ${command('source','file-down','Download original workbook', `data-index="${index}"`)}</div>` : ''}${readOnly ? '' : '<div class="mt-6 border-t pt-4 flex justify-end gap-2"><button type="button" data-support-action="close-dialog" class="border rounded-lg px-4 py-2 text-xs">Cancel</button><button type="submit" class="inline-flex items-center gap-2 bg-indigo-600 text-white rounded-lg px-4 py-2 text-xs font-semibold">'+icon('save')+'Save Ticket</button></div>'}</form>`);
    const form=document.getElementById('support-ticket-form');
    form.querySelectorAll('[data-ticket-field]').forEach(input=>{input.value=String(record?.fields?.[input.dataset.ticketField]??'');});
    form.insertAdjacentHTML('afterbegin',derivedPreview(record||{fields:{}}));
  }
  async function saveTicket(form) {
    if (!editable() || state.busy || !form.reportValidity()) return;
    const index = form.dataset.index === '' ? null : Number(form.dataset.index), record = index === null ? null : state.records[index];
    const fields = { ...(record?.fields || {}) }; form.querySelectorAll('[data-ticket-field]').forEach(input => { fields[input.dataset.ticketField] = input.value; });
    const requestId = String(fields.RequestID || '').trim();
    if (!requestId || !String(fields.Subject || '').trim()) return showToast('RequestID and Subject are required.','error');
    state.busy = true; const host = document.getElementById('support-ticket-dialog'); host.dataset.busy = 'true'; form.querySelector('button[type="submit"]').disabled = true;
    try {
      const query = record ? supabaseClient.from('support_tickets').update({ fields }).eq('request_id',record.request_id).eq('version',record.version) : supabaseClient.from('support_tickets').insert({ request_id: requestId, fields });
      const { data, error } = await query.select('*').single();
      if (error) throw new Error(error.code === '23505' ? 'A ticket with this RequestID already exists.' : error.code === 'PGRST116' ? 'This ticket changed elsewhere. Refresh before editing.' : error.message);
      if (record) state.records[index] = data; else state.records.push(data);
      host.dataset.busy = 'false'; closeDialog(); render(); showToast('Support ticket saved.','success');
    } catch (error) { showToast(error.message,'error'); }
    finally { state.busy = false; if (host.isConnected) { host.dataset.busy = 'false'; form.querySelector('button[type="submit"]').disabled = false; } }
  }
  const headerKey = input => String(input ?? '').replace(/^\uFEFF/,'').trim().toLowerCase().replace(/[^a-z0-9]/g,'');
  function parseWorkbook(workbook) {
    const known = new Map([...headers,...optionalHeaders].map(header => [headerKey(header),header])), rows = [], errors = [];
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName]; if (!sheet['!ref']) continue;
      const range = XLSX.utils.decode_range(sheet['!ref']); let headerRow = -1, names = [];
      for (let row = range.s.r; row <= Math.min(range.e.r,range.s.r + 100); row++) {
        const candidate = [];
        for (let column = range.s.c; column <= range.e.c; column++) candidate.push(String(sheet[XLSX.utils.encode_cell({r:row,c:column})]?.v ?? ''));
        const keys = candidate.map(headerKey);
        if (keys.includes('requestid') && keys.includes('subject')) { headerRow = row; names = candidate; break; }
      }
      if (headerRow < 0) continue;
      const counts = new Map(); names.forEach(name => { const key = headerKey(name); counts.set(key,(counts.get(key)||0)+1); });
      if (headers.some(header => counts.get(headerKey(header)) > 1)) { errors.push(`${sheetName}: repeated ticket column headers are ambiguous.`); continue; }
      for (let row = headerRow + 1; row <= range.e.r; row++) {
        const fields = Object.assign(Object.create(null),Object.fromEntries(headers.map(header => [header,'']))), cells = []; let populated = false;
        names.forEach((name, offset) => {
          const column = range.s.c + offset, cell = sheet[XLSX.utils.encode_cell({r:row,c:column})];
          if (cell?.v !== undefined && cell.v !== null && String(cell.v) !== '') populated = true;
          const field = known.get(headerKey(name)) || name || `Column ${column + 1}`;
          const uniqueField = Object.prototype.hasOwnProperty.call(fields,field) && !headers.includes(field) ? `${field} (Column ${column + 1})` : field;
          const text = cell ? cell.w !== undefined ? String(cell.w) : String(XLSX.utils.format_cell(cell)) : '';
          fields[uniqueField] = text;
          cells.push({ column: column + 1, header: name, ...(cell || {}) });
        });
        if (populated) rows.push({ fields, source_sheet:sheetName, source_row:row + 1, source_cells:{ headers:names, cells, date1904:Boolean(workbook.Workbook?.WBProps?.date1904) } });
      }
    }
    if (errors.length) throw new Error(errors.join('\n'));
    if (!rows.length) throw new Error('No ticket rows found. A ticket sheet must include RequestID and Subject headers.');
    return rows;
  }
  function classifyRows(rows) {
    const seen = new Set(state.records.map(record => record.request_id)), candidates = [], results = [];
    for (const row of rows) {
      const id = String(row.fields.RequestID || '').trim(), entry = {request_id:id,sheet:row.source_sheet,row:row.source_row};
      if (!id || !String(row.fields.Subject || '').trim()) results.push({...entry,status:'failed',message:'RequestID and Subject are required.'});
      else if (seen.has(id)) results.push({...entry,status:'duplicate',message:'RequestID already exists.'});
      else { seen.add(id); candidates.push(row); results.push({...entry,status:'ready',message:''}); }
    }
    return {candidates,results};
  }
  function importErrorMessage(error) {
    const message=String(error?.message || error || 'Unknown import error.');
    return ['PGRST202','PGRST204','PGRST205','42P01','42883','42703'].includes(error?.code) || /schema cache|could not find.*support_ticket_import|support_ticket_imports.*does not exist/i.test(message)
      ? "Support Ticket import database setup is missing or outdated. Run supabase-migration-support-ticket-tracker.sql in Supabase SQL Editor. It creates public.support_ticket_imports and refreshes the schema cache with NOTIFY pgrst, 'reload schema'; then reopen the importer. No ticket records were saved."
      : message;
  }
  async function importRpc(name,args={}) {
    const {data,error}=await supabaseClient.rpc(name,args);
    if(error)throw error;
    return data;
  }
  async function ensureImportSchema() {
    const schema=await importRpc('support_ticket_import_schema');
    if(!schema || schema.version<2 || !schema.atomic || !schema.imports_table || !schema.tickets_table)throw {code:'PGRST205',message:'Support Ticket import schema is outdated.'};
    for(const [table,columns] of [
      ['support_ticket_imports','id,file_name,file_size,file_base64,workbook,validation_summary,status,total_rows,processed_rows,staged_rows,error_message'],
      ['support_tickets','request_id,fields,source_cells,source_import_id,source_sheet,source_row,version']
    ]){
      const {error}=await supabaseClient.from(table).select(columns).limit(0);
      if(error)throw error;
    }
  }
  function progressMarkup() {
    const progress=state.preview?.progress || {processed:0,total:0,label:'Preparing import'},percent=progress.total?Math.round(progress.processed/progress.total*100):0;
    return `<div class="space-y-3 mb-5" aria-live="polite"><div class="import-progress-track" role="progressbar" aria-label="Support Ticket import progress" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${percent}"><div id="support-import-progress-bar" class="import-progress-bar" style="width:${percent}%"></div></div><div class="flex flex-wrap items-center justify-between gap-2"><span id="support-import-progress-detail" class="text-sm font-medium text-slate-700">${escape(progress.label)} ${progress.processed} of ${progress.total} rows</span><span id="support-import-progress-percent" class="text-lg font-bold text-indigo-600">${percent}%</span></div></div>`;
  }
  function updateSupportImportProgress(processed,total,label) {
    if(!state.preview)return;
    state.preview.progress={processed:Math.min(total,Math.max(0,processed)),total,label};
    const progress=state.preview.progress,percent=total?Math.round(progress.processed/total*100):0;
    const bar=document.getElementById('support-import-progress-bar');if(bar){bar.style.width=`${percent}%`;bar.parentElement.setAttribute('aria-valuenow',String(percent));}
    const detail=document.getElementById('support-import-progress-detail');if(detail)detail.textContent=`${label} ${progress.processed} of ${total} rows`;
    const percentage=document.getElementById('support-import-progress-percent');if(percentage)percentage.textContent=`${percent}%`;
  }
  function importReport() {
    const preview = state.preview, counts = {successful:0,updated:0,duplicate:0,failed:0,ready:0};
    preview.results.forEach(row => counts[row.status] = (counts[row.status] || 0) + 1);
    const metrics=Object.entries(counts).filter(([label])=>label!=='ready'||!preview.finished),canImport=!preview.finished && !counts.failed && preview.schemaReady && (!preview.error || preview.uncertain);
    const setupActions=!preview.schemaReady&&!preview.finished?`<div class="flex flex-wrap gap-2 mb-4"><button type="button" data-support-action="schema-sql" class="inline-flex items-center gap-2 border rounded-lg px-3 py-2 text-xs">${icon('file-down')}Download Database Migration</button><button type="button" data-support-action="schema-check" class="inline-flex items-center gap-2 border rounded-lg px-3 py-2 text-xs">${icon('refresh-cw')}Check Database Again</button></div>`:'';
    return `<div class="text-xs text-slate-500 mb-4 break-words">${escape(preview.name)}</div>${progressMarkup()}${preview.error?`<div role="alert" class="border border-rose-200 bg-rose-50 text-rose-700 rounded-lg p-3 text-xs mb-4 whitespace-pre-wrap">${escape(preview.error)}</div>`:''}${setupActions}<div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">${metrics.map(([label,count]) => `<div class="border rounded-lg p-3"><div class="uppercase text-[10px] text-slate-500">${label==='duplicate'?'Duplicate / Skipped':label==='ready'&&!preview.schemaReady?'Not processed':label}</div><div class="text-xl font-bold text-slate-800">${count}</div></div>`).join('')}</div><div class="overflow-x-auto border rounded-lg max-h-80"><table class="w-full text-xs"><thead class="bg-slate-100 sticky top-0 text-left"><tr>${['Sheet','Row','RequestID','Result','Details'].map(label=>`<th class="p-3">${label}</th>`).join('')}</tr></thead><tbody>${preview.results.slice(0,50).map(row => `<tr class="border-t"><td class="p-3">${escape(row.sheet)}</td><td class="p-3">${row.row ?? ''}</td><td class="p-3">${escape(row.request_id)}</td><td class="p-3">${!preview.schemaReady&&row.status==='ready'?'Not processed':escape(row.status)}</td><td class="p-3">${escape(row.message)}</td></tr>`).join('')}</tbody></table></div><div class="text-xs text-slate-500 mt-2">${preview.results.length} rows; first 50 displayed.</div><div class="flex flex-wrap justify-end gap-2 mt-5">${command('validation-export','download','Download validation report')}<button data-support-action="close-dialog" class="border rounded-lg px-4 py-2 text-xs">Close</button>${canImport ? `<button data-support-action="commit-import" class="bg-indigo-600 text-white rounded-lg px-4 py-2 text-xs font-semibold inline-flex items-center gap-2">${icon('upload')}${preview.uncertain?'Recover Import':`Import ${preview.rows.length} Rows`}</button>` : ''}</div>`;
  }
  function openImport() {
    if (!editable()) return showToast('Permission denied.','error');
    dialog('Import Support Tickets', `<div class="space-y-4"><input id="support-import-file" type="file" accept=".xls,.xlsx" aria-label="Support ticket workbook" class="w-full text-xs border rounded-lg p-3"><div class="flex justify-end"><button data-support-action="template" class="border rounded-lg px-4 py-2 text-xs inline-flex items-center gap-2">${icon('file-down')}Download Template</button></div><div id="support-import-preview" aria-live="polite">${state.preview?importReport():''}</div></div>`);
  }
  async function recheckImportSchema() {
    const preview=state.preview;if(!editable()||!preview||preview.finished||state.busy)return;
    const host=document.getElementById('support-ticket-dialog');state.busy=true;if(host)host.dataset.busy='true';
    try{
      await ensureImportSchema();await load(true);if(state.error)throw new Error(state.error);
      Object.assign(preview,classifyRows(preview.rows),{schemaReady:true,error:''});
      if(preview.results.some(row=>row.status==='failed'))preview.error='Validation failed. Correct the required fields and select the workbook again. No ticket records were saved.';
      updateSupportImportProgress(0,preview.rows.length,'Ready to validate RequestIDs');
    }catch(error){preview.schemaReady=false;preview.error=importErrorMessage(error);updateSupportImportProgress(0,preview.rows.length,'Database check blocked; not processed');}
    finally{
      state.busy=false;if(host)host.dataset.busy='false';
      const area=document.getElementById('support-import-preview');if(area)area.innerHTML=importReport();lucide.createIcons();
    }
  }
  async function previewFile(file) {
    if (!file || state.busy) return;
    const previewHost = document.getElementById('support-import-preview'); state.preview=null;previewHost.textContent = 'Reading workbook...';
    try {
      const bytes = new Uint8Array(await file.arrayBuffer()), workbook = XLSX.read(bytes,{type:'array',cellNF:true,cellStyles:true,cellFormula:true,cellText:true,cellDates:false}), rows = parseWorkbook(workbook);
      state.preview = {name:file.name,size:file.size,bytes,workbook,rows,...classifyRows(rows),finished:false,schemaReady:false,importId:crypto.randomUUID(),progress:{processed:0,total:rows.length,label:'Preparing import'}};
      await recheckImportSchema();
      previewHost.innerHTML = importReport(); lucide.createIcons();
    } catch (error) { state.preview = null; previewHost.textContent = error.message; }
  }
  function fileBase64(bytes) {
    const parts = [];
    for (let start = 0; start < bytes.length; start += 32768) parts.push(String.fromCharCode(...bytes.subarray(start,start + 32768)));
    return btoa(parts.join(''));
  }
  async function downloadMigration() {
    try {
      const response=await fetch('supabase-migration-support-ticket-tracker.sql');if(!response.ok)throw new Error('Database migration file is unavailable.');
      const url=URL.createObjectURL(new Blob([await response.text()],{type:'text/plain'})),link=document.createElement('a');
      link.href=url;link.download='supabase-migration-support-ticket-tracker.sql';link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
    }catch(error){showToast(error.message,'error');}
  }
  async function commitImport() {
    const preview = state.preview; if (!editable() || state.busy || !preview || preview.finished || !preview.schemaReady || preview.results.some(row=>row.status==='failed')) return;
    state.busy = true; const host = document.getElementById('support-ticket-dialog'); host.dataset.busy = 'true'; host.querySelectorAll('button,input').forEach(control=>control.disabled=true);
    let started=false,report=null;
    preview.error='';preview.uncertain=false;
    updateSupportImportProgress(preview.progress?.processed||0,preview.rows.length,'Validating RequestIDs');
    try {
      await ensureImportSchema();
      const job=await importRpc('begin_support_ticket_import',{p_client_id:preview.importId,p_source:{file_name:preview.name,file_size:preview.size,file_base64:fileBase64(preview.bytes),workbook:preview.workbook},p_total_rows:preview.rows.length});
      started=true;
      if(job.status==='completed'||job.status==='failed')report=job.report;
      else {
        for(let offset=job.processed;offset<preview.rows.length;){
          const rows=preview.rows.slice(offset,offset+200),result=await importRpc('stage_support_ticket_import',{p_import_id:preview.importId,p_offset:offset,p_rows:rows});
          if(result.processed!==offset+rows.length)throw new Error('Unexpected import progress. No ticket records have been committed.');
          preview.results.splice(offset,rows.length,...result.rows);offset=result.processed;
          updateSupportImportProgress(offset,preview.rows.length,'Validated RequestIDs');
        }
        updateSupportImportProgress(preview.rows.length,preview.rows.length,'Saving validated import atomically');
        report=await importRpc('complete_support_ticket_import',{p_import_id:preview.importId});
      }
    } catch (error) {
      preview.error=importErrorMessage(error);
      if(!started){preview.schemaReady=false;updateSupportImportProgress(0,preview.rows.length,'Database check blocked; not processed');showToast(preview.error,'error');}
      if(started){
        try{report=await importRpc('abort_support_ticket_import',{p_import_id:preview.importId,p_error:preview.error});}
        catch(recoveryError){preview.uncertain=true;preview.error='The server outcome could not be confirmed. No partial ticket import is possible. Use Recover Import with the same workbook to confirm whether the whole import committed. '+importErrorMessage(recoveryError);}
      }
      if(started && !preview.uncertain && (!report || report.status!=='completed')){
        report={status:'failed',rows:preview.rows.map(row=>({request_id:row.fields.RequestID,sheet:row.source_sheet,row:row.source_row,status:'failed',message:preview.error}))};
      }
    } finally {
      if(report){
        preview.finished=true;
        if(report.rows?.length)preview.results=report.rows;
        else preview.results=preview.rows.map(row=>({request_id:row.fields.RequestID,sheet:row.source_sheet,row:row.source_row,status:'failed',message:report.error || preview.error || 'Import failed; no ticket records were saved.'}));
        if(report.status==='completed'){
          preview.error='';updateSupportImportProgress(preview.rows.length,preview.rows.length,'Import complete');
          showToast(`Import complete: ${report.successful} successful, ${report.updated||0} updated, ${report.duplicate} duplicate/skipped, ${report.failed} failed.`,'success');
        }else {preview.error=preview.error || preview.results.find(row=>row.status==='failed')?.message || 'Import failed; no ticket records were saved.';updateSupportImportProgress(preview.progress.processed,preview.rows.length,'Import failed after validating');showToast(preview.error,'error');}
      }
      state.busy = false; host.dataset.busy = 'false'; host.querySelectorAll('button,input').forEach(control=>control.disabled=false);
      const area = document.getElementById('support-import-preview'); if (area) area.innerHTML = importReport();
      if(started)await load(true);lucide.createIcons();
    }
  }
  function exportWorkbook(template = false, validation = false) {
    const workbook = XLSX.utils.book_new();
    if (validation) {
      XLSX.utils.book_append_sheet(workbook,XLSX.utils.json_to_sheet(state.preview?.results || []),'Validation Results');
    } else {
      const records = template ? [] : filtered(), columns = template ? [...headers,...optionalHeaders] : gridColumns(records);
      const sheet = XLSX.utils.aoa_to_sheet([columns,...records.map(record=>columns.map(field=>value(record,field)))]);
      sheet['!cols'] = columns.map(field=>({wch:field === 'Subject' ? 55 : 24}));
      XLSX.utils.book_append_sheet(workbook,sheet,'Support Tickets');
      if (!template) for (const dimension of dimensions) XLSX.utils.book_append_sheet(workbook,XLSX.utils.json_to_sheet(summaries(records,dimension).map(group=>({[dimension]:group.name,Total:group.total,Open:group.open,Closed:group.closed,'Average Age (Days)':group.ages.length ? group.ages.reduce((total,days)=>total+days,0)/group.ages.length : ''}))),dimension.slice(0,31));
    }
    XLSX.writeFile(workbook,validation ? 'Support_Ticket_Validation.xlsx' : template ? 'Support_Ticket_Import_Template.xlsx' : `Support_Tickets_${new Date().toISOString().slice(0,10)}.xlsx`);
  }
  async function downloadSource(index) {
    const record = state.records[index]; if (!editable() || !record?.source_import_id) return;
    const result = await supabaseClient.from('support_ticket_imports').select('file_name,file_base64').eq('id',record.source_import_id).single();
    if (result.error) return showToast(result.error.message,'error');
    const bytes = Uint8Array.from(atob(result.data.file_base64),character=>character.charCodeAt(0)), url = URL.createObjectURL(new Blob([bytes]));
    const link = document.createElement('a'); link.href = url; link.download = result.data.file_name; link.click(); setTimeout(()=>URL.revokeObjectURL(url),1000);
  }
  document.addEventListener('input',event=>{
    if(active() && event.target.id === 'support-search') {state.query=event.target.value;state.page=1;renderResults();}
    const form=event.target.closest?.('#support-ticket-form');
    if(form&&event.target.hasAttribute('data-ticket-field')){
      const record=form.dataset.index===''?null:state.records[Number(form.dataset.index)],fields={...(record?.fields||{})};
      form.querySelectorAll('[data-ticket-field]').forEach(input=>{fields[input.dataset.ticketField]=input.value;});
      const preview=document.getElementById('support-derived-preview');if(preview){preview.outerHTML=derivedPreview({...record,fields});lucide.createIcons();}
    }
  });
  document.addEventListener('change',event=>{
    const control=event.target;
    if(control.id === 'support-import-file') {previewFile(control.files[0]);return;}
    if(!active())return;
    if(control.dataset.supportFilter) {state.filters[control.dataset.supportFilter]=control.value;state.page=1;renderResults();}
    if(control.hasAttribute('data-support-size')) {state.size=Number(control.value);state.page=1;renderResults();}
    if(control.hasAttribute('data-support-sort')) {state.sort=control.value;renderResults();}
    if(control.hasAttribute('data-support-summary')) {state.summary=control.value;state.page=1;renderResults();}
  });
  document.addEventListener('submit',event=>{if(event.target.id === 'support-ticket-form'){event.preventDefault();saveTicket(event.target);}});
  document.addEventListener('click',event=>{
    const button=event.target.closest('[data-support-action]');if(!button)return;
    const action=button.dataset.supportAction,index=Number(button.dataset.index);
    if(action === 'close-dialog')return closeDialog();
    if(action === 'template')return exportWorkbook(true);
    if(action === 'commit-import')return commitImport();
    if(action === 'validation-export')return exportWorkbook(false,true);
    if(action === 'schema-sql')return downloadMigration();
    if(action === 'schema-check')return recheckImportSchema();
    if(action === 'source')return downloadSource(index);
    if(!active())return;
    if(action === 'view' || action === 'edit')return openTicket(index,action);
    if(action === 'create')return openTicket(null,'create');
    if(action === 'refresh')return load(true);
    if(action === 'expand-card'){const id=state.records[index].request_id;state.openCards.has(id)?state.openCards.delete(id):state.openCards.add(id);return renderResults();}
    if(action === 'reset'){state.query='';state.filters={};state.page=1;return render();}
    if(action === 'filters'){state.expanded=!state.expanded;return render();}
    if(action === 'view-mode'){state.view=button.dataset.view;return render();}
    if(action === 'direction')state.direction*=-1;
    if(action === 'sort'){state.direction=state.sort === button.dataset.field ? -state.direction : 1;state.sort=button.dataset.field;}
    if(action === 'previous')state.page--;
    if(action === 'next')state.page++;
    renderResults();
  });
  document.addEventListener('keydown',event=>{
    const host=document.getElementById('support-ticket-dialog');if(!host)return;
    if(event.key==='Escape'){event.preventDefault();closeDialog();}
    if(event.key==='Tab'){const controls=[...host.querySelectorAll('button:not(:disabled),input:not(:disabled),select,textarea,a[href]')],first=controls[0],last=controls.at(-1);if(event.shiftKey&&document.activeElement===first){event.preventDefault();last?.focus();}else if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first?.focus();}}
  });
  const previousHeader=window.updateWorkspaceHeaderActions;
  window.updateWorkspaceHeaderActions=function(){
    const result=previousHeader.apply(this,arguments),isSupport=active(),importButton=document.getElementById('btn-page-import'),exportButton=document.getElementById('btn-page-export');
    const header=document.getElementById('workspace-page-title')?.closest('header');if(header)header.dataset.supportHeader=String(isSupport);
    for(const [id,label,name,action] of [['btn-support-template','Download Template','file-down','template'],['btn-support-create','New Ticket','ticket-plus','create']]){
      let button=document.getElementById(id);
      if(!button && exportButton){button=document.createElement('button');button.type='button';button.id=id;button.className=(id==='btn-support-create'?document.getElementById('btn-new-project'):exportButton)?.className || exportButton.className;button.dataset.supportAction=action;button.innerHTML=icon(name)+`<span>${label}</span>`;exportButton.parentElement.append(button);}
      if(button){button.style.display=isSupport && (action==='template'||editable())?'':'none';button.title=label;button.setAttribute('aria-label',label);}
    }
    if(isSupport){
      if(importButton){importButton.style.display=editable()?'':'none';importButton.title='Import Support Tickets';}
      if(exportButton){exportButton.style.display='';exportButton.title='Export Support Tickets';}
      for(const id of ['btn-new-program','btn-new-project']){const button=document.getElementById(id);if(button)button.style.display='none';}
    }
    return result;
  };
  const previousImport=window.openPageImportModal;
  window.openPageImportModal=function(){return active()?openImport():previousImport.apply(this,arguments);};
  const previousExport=window.exportCurrentWorkspacePage;
  window.exportCurrentWorkspacePage=function(){return active()?exportWorkbook():previousExport.apply(this,arguments);};
  const previousFlatRender=window.renderFlatWorkspacePage;
  window.renderFlatWorkspacePage=function(){return active()?render():previousFlatRender.apply(this,arguments);};
  const style=document.createElement('style');style.textContent=`
    #workspace-flat-page .support-primary{display:grid;grid-template-columns:minmax(150px,1.6fr) repeat(4,minmax(110px,1fr)) auto;gap:8px;align-items:center}
    #workspace-flat-page .support-primary input,#workspace-flat-page .support-primary select{min-height:38px;min-width:0;max-width:100%}
    .support-icon-button{display:inline-flex;align-items:center;justify-content:center;width:36px;height:36px;flex:0 0 36px}
    .support-icon-button:disabled{opacity:.4;cursor:not-allowed}
    .support-secondary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;padding-top:8px}
    .support-secondary[hidden]{display:none!important}
    .support-details{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px}
    .support-filter-control{min-width:0}
    #support-ticket-dialog input[readonly],#support-ticket-dialog textarea[readonly]{background:#f8fafc}
    @media(max-width:1100px){#workspace-flat-page .support-primary{grid-template-columns:minmax(0,1fr) minmax(0,1fr)}#support-search{grid-column:1/-1}.support-secondary{grid-template-columns:repeat(2,minmax(0,1fr))}}
    @media(max-width:640px){.support-details{grid-template-columns:minmax(0,1fr)}#support-ticket-dialog{padding:8px}.support-icon-button{width:34px;height:34px;flex-basis:34px}header[data-support-header="true"]>div{height:auto!important;min-height:64px;padding-top:8px;padding-bottom:8px}header[data-support-header="true"]>div>div{flex-wrap:wrap;gap:8px}header[data-support-header="true"] #btn-page-import,header[data-support-header="true"] #btn-page-export,header[data-support-header="true"] #btn-support-template,header[data-support-header="true"] #btn-support-create{font-size:0;width:36px;height:36px;min-width:36px;padding:0;justify-content:center}header[data-support-header="true"] button span{display:none}}
  `;document.head.append(style);
  document.getElementById('workspace-pages')?.append(document.getElementById('workspace-nav-support'));
  window.supportTickets = { headers, optionalHeaders, dimensions, state, load, render, filtered, summaries, agingDays, deriveTicket, gridColumns, value, parseWorkbook, classifyRows, openTicket, saveTicket, previewFile, commitImport, exportWorkbook, recheckImportSchema };
})();