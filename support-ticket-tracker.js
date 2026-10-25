(() => {
  const headers = ['RequestID','Request Type','Subject','Request Status','Location','Department','Service Category','Request Mode','Category','Sub Category','Type of issue','Priority','Template','Group','Status','Technician','Created Time','Last Updated Time','Resolution','Requester','Approval Status','Assigned Time','Completed Time','Created By','Due By Time','Item','Last Updated By','Overdue Status','Requester Email','Resolved By','Resolved Time','Site','Site Code','Type','Time Elapsed','Solution','Responded Date','Response Due By Time','Resolve Mode','Region'];
  const dimensions = ['Request Type','Request Status','Status','Location','Department','Service Category','Request Mode','Category','Sub Category','Priority','Group','Technician','Region'];
  const optionalHeaders = [];
  const gridFirst = ['RequestID','Subject','Request Status','Status','Requester','Priority','Group','Technician','Location','Created Time','Last Updated Time','Due By Time'];
  const dateHeaders = ['Created Time','Last Updated Time','Assigned Time','Completed Time','Due By Time','Resolved Time','Responded Date','Response Due By Time'];
  const ageingBuckets = ['0\u20135 days','6\u201310 days','11\u201320 days','More than 20 days'];
  const state = { records: [], loaded: false, loading: false, error: '', user: '', query: '', filters: {}, page: 1, size: typeof DELIVERY_PAGE_SIZE === 'number' ? DELIVERY_PAGE_SIZE : 25, sort: 'RequestID', direction: 1, expanded: false, view: 'grid', summary: 'Request Type', busy: false, openCards: new Set(), selected: new Set(), preview: null };
  const supportScriptUrl=document.currentScript?.src||'';
  let searchTimer = 0;
  let pageRequest=0,pageController;
  function pageArguments(){return {p_query:state.query,p_filters:Object.fromEntries(Object.entries(state.filters).map(([field,selected])=>[field,window.trackerMultiFilters?trackerMultiFilters.values(`support:${field}`,selected):selected?[selected]:[]])),p_sort:state.sort,p_direction:state.direction,p_offset:(state.page-1)*state.size,p_limit:state.size,p_summary:state.view==='summary'?state.summary:null};}
  const root = () => document.getElementById('workspace-flat-page');
  const active = () => currentWorkspaceView === 'support';
  const editable = () => Boolean(currentUser && ['Admin','Level 1'].includes(currentUser.role));
  const escape = value => escapeHtml(String(value ?? ''));
  function loadErrorHint(error) {
    const message=String(error?.message||error||''),code=String(error?.code||'');
    if(/support_ticket_page/i.test(message)&&(/^(PGRST202|42883)$/.test(code)||/schema cache|could not find.*function|does not exist/i.test(message)))return 'Apply supabase-migration-support-ticket-bulk-import.sql in Supabase, then refresh.';
    if(/support_tickets|schema cache/i.test(message))return 'Apply supabase-migration-support-ticket-tracker.sql followed by supabase-migration-support-ticket-tabular.sql in Supabase, then refresh.';
    return '';
  }
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
    const namedMonthFirst=text.match(/^([a-z]{3,9})[ -]+(\d{1,2}),?[ -]+(\d{4})/i);
    if(namedMonthFirst){const month=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(namedMonthFirst[1].slice(0,3).toLowerCase())+1;return month?calendarDate(Number(namedMonthFirst[3]),month,Number(namedMonthFirst[2]),time):null;}
    return null;
  }
  function normalizedId(value) { return String(value ?? '').trim().toLocaleLowerCase('en-US'); }
  function templateDateError(record,field) {
    const raw=String(record.fields[field]??'').trim();
    if(!raw||/^(?:-|not assigned)$/i.test(raw))return '';
    return ticketDate(record,{key:field,text:raw})?'':`${field} must contain a valid date/time or be blank, "-" or "Not Assigned".`;
  }
  function dateTimeRank(record,field) {
    const parsed=ticketDate(record,{key:field,text:String(record.fields[field]??'')});
    if(!parsed)return null;
    const parts=(parsed.time||'00:00:00').split(':').map(Number);
    const raw=String(record.fields[field]??''),zone=raw.match(/(Z|([+-])(\d{2}):?(\d{2}))$/i);
    const offset=zone?.[1].toUpperCase()==='Z'?0:zone?((Number(zone[3])*60+Number(zone[4]))*(zone[2]==='+'?1:-1)):0;
    return parsed.day*86400+(parts[0]||0)*3600+(parts[1]||0)*60+(parts[2]||0)-offset*60;
  }
  function normalizeExistingFields(record) {
    const fields=Object.assign(Object.create(null),record.fields||{});
    const aliases={'RequestID':['ID'],'Requester':['Requestor'],'Site Code':['Location Code'],'Region':['Zone'],'Technician':['Engineer Name'],'Time Elapsed':['Aging','Ageing','Ticket Age'],'Created Time':['Created Date'],'Completed Time':['Closure Date']};
    for(const [target,names] of Object.entries(aliases))if(!String(fields[target]??'').trim()){
      const source=names.find(name=>String(fields[name]??'').trim());if(source)fields[target]=fields[source];
    }
    if(!String(fields.RequestID??'').trim())fields.RequestID=record.request_id;
    for(const field of headers)if(fields[field]===undefined||fields[field]===null)fields[field]='';
    return {...record,fields};
  }
  function deriveTicket(record,referenceDate) {
    const createdField=sourceField(record,['Created Time']);
    const closureField=sourceField(record,['Completed Time','Resolved Time']);
    const created=ticketDate(record,createdField),closure=ticketDate(record,closureField);
    const statusText=[sourceField(record,['Request Status']).text,sourceField(record,['Status']).text].join(' ').toLowerCase();
    const stage=/closed|resolved|completed|cancelled|canceled/.test(statusText)||Boolean(closureField.text.trim())?'Closed':'Open';
    const now=new Date(),today=referenceDate|| (typeof localToday==='function'?localToday():`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`);
    const onHold=sourceField(record,['Request Status']).text.trim().toLowerCase()==='on hold';
    const end=stage==='Closed'?closure:ticketDate({fields:{}},{key:'Today',text:today});
    const age=!onHold&&created&&end?Math.abs(end.day-created.day):null;
    const bucket=age===null?'':ageingBuckets[age<=5?0:age<=10?1:age<=20?2:3];
    const time=created?.time || '';
    return {stage,age,bucket,createdDate:created?.date||'',createdTime:time,onHold,flag:onHold?'On Hold':age===null?'Check dates':bucket};
  }
  const derivedTickets = new WeakMap();
  function cachedTicket(record) {
    const now = new Date();
    const today = typeof localToday === 'function' ? localToday() : `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
    const cached = derivedTickets.get(record);
    if (cached && cached.fields === record.fields && cached.source === record.source_cells && cached.version === record.version && cached.today === today) return cached.result;
    const result = deriveTicket(record, today);
    derivedTickets.set(record, { fields: record.fields, source: record.source_cells, version: record.version, today, result });
    return result;
  }
  function value(record,field) {
    if(['Stage','Age in Days','Ageing Bucket','Ageing Flag'].includes(field)){
      const derived=cachedTicket(record);
      return String(({Stage:derived.stage,'Age in Days':derived.age??'','Ageing Bucket':derived.bucket,'Ageing Flag':derived.flag})[field]);
    }
    if(field==='RequestID')return String(record.fields?.RequestID??record.request_id??'');
    return String(record.fields?.[field]??'');
  }
  async function load(force = false) {
    const user = String(currentUser?.email || currentUser?.id || '');
    if (state.user !== user) { state.records = []; state.loaded = false; state.user = user; state.query = ''; state.filters = {}; state.preview = null;state.options=null;state.remote=false;state.selected.clear(); }
    if (state.loaded && !force) return;
    const request=++pageRequest;pageController?.abort();pageController=new AbortController();
    state.loading = true; state.error = '';
    try {
      const args=pageArguments();state.pendingKey=JSON.stringify(args);let query=supabaseClient.rpc('support_ticket_page',{...args,p_metadata:!state.options});
      if(query.abortSignal)query=query.abortSignal(pageController.signal);
      const {data,error}=await query;if(error)throw error;
      if(state.user!==user||request!==pageRequest)return;
      if(!data)throw new Error('Support Ticket paging setup is unavailable. Apply supabase-migration-support-ticket-bulk-import.sql.');
      state.records=(data.records||[]).map(normalizeExistingFields);state.remote=true;state.remoteKey=JSON.stringify(args);state.total=data.total;state.groups=data.groups||[];state.groupTotal=data.group_total;state.metrics=data.metrics;state.options=data.options||state.options||{};state.loaded=true;
    } catch (error) { if(request===pageRequest){const hint=loadErrorHint(error);state.error=hint?`Support Ticket paging RPC is unavailable. ${hint}`:error.message;} }
    finally { if(request===pageRequest){state.loading = false;state.pendingKey=null; if (active()) render();} }
  }
  const icon = name => `<i data-lucide="${name}" class="w-4 h-4" aria-hidden="true"></i>`;
  const command = (action, name, label, extra = '') => `<button type="button" data-support-action="${action}" title="${label}" aria-label="${label}" class="support-icon-button border border-slate-300 rounded-lg bg-white text-slate-600 hover:bg-slate-50" ${extra}>${icon(name)}</button>`;
  const status = record => value(record, 'Status') || value(record, 'Request Status');
  const closed = record => cachedTicket(record).stage==='Closed';
  function agingDays(text) {
    const input = String(text).trim();
    if (/^\d+(?:\.\d+)?$/.test(input)) return Number(input);
    let days = 0, found = false;
    for (const match of input.matchAll(/(\d+(?:\.\d+)?)\s*(days?|d\b|hours?|hrs?|h\b|minutes?|mins?|m\b)/gi)) {
      days += Number(match[1]) / (/^d/i.test(match[2]) ? 1 : /^h/i.test(match[2]) ? 24 : 1440); found = true;
    }
    return found ? days : null;
  }
  let filteredTickets = null;
  const ticketSortCollator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
  const ticketOptionCollator = new Intl.Collator(undefined, { numeric: true });
  function filtered() {
    if(state.remote)return state.records.slice();
    const query = state.query.trim().toLowerCase();
    const day = typeof localToday === 'function' ? localToday() : new Date().toDateString();
    const filters = Object.entries(state.filters).map(([field,selected]) => [field,window.trackerMultiFilters ? trackerMultiFilters.values(`support:${field}`,selected) : selected]);
    const key = JSON.stringify([query,filters,state.sort,state.direction,day]);
    if (filteredTickets && filteredTickets.records === state.records && filteredTickets.items.length === state.records.length && filteredTickets.key === key && filteredTickets.items.every((record,index) => record === state.records[index] && filteredTickets.versions[index] === record.version && filteredTickets.fields[index] === record.fields && filteredTickets.sources[index] === record.source_cells)) return filteredTickets.rows.slice();
    const rows = state.records.filter(record => (!query || [...Object.values(record.fields),...gridFirst.map(field=>value(record,field))].some(item => String(item ?? '').toLowerCase().includes(query))) && Object.entries(state.filters).every(([field, selected]) => window.trackerMultiFilters ? trackerFilterMatches(`support:${field}`,value(record,field),selected) : !selected || value(record, field) === selected))
      .sort((left, right) => state.direction * (state.sort === 'Age in Days'
        ? (cachedTicket(left).age ?? -1) - (cachedTicket(right).age ?? -1)
        : state.sort==='Ageing Bucket'?ageingBuckets.indexOf(value(left,state.sort))-ageingBuckets.indexOf(value(right,state.sort))
        : ticketSortCollator.compare(value(left, state.sort), value(right, state.sort))) || left.request_id.localeCompare(right.request_id));
    filteredTickets = { records: state.records, items: state.records.slice(), versions: state.records.map(record=>record.version), fields: state.records.map(record=>record.fields), sources: state.records.map(record=>record.source_cells), key, rows };
    return rows.slice();
  }
  function summaries(records, dimension) {
    const groups = new Map();
    for (const record of records) {
      const name = value(record, dimension) || '(Not specified)';
      const entry = groups.get(name) || { name, total: 0, open: 0, closed: 0, ages: [] };
      entry.total++; closed(record) ? entry.closed++ : entry.open++;
      const days = cachedTicket(record).age; if (days !== null) entry.ages.push(days);
      groups.set(name, entry);
    }
    return [...groups.values()].sort((left, right) => right.total - left.total || left.name.localeCompare(right.name));
  }
  function filterControl(field) {
    const options = [...new Set([...(state.remote?state.options?.[field]||[]:state.records.map(record=>value(record,field))),state.filters[field]].filter(Boolean))].sort(ticketOptionCollator.compare);
    return `<label class="support-filter-control"><span class="sr-only">${escape(field)}</span><select data-support-filter="${escape(field)}" aria-label="${escape(field)}" class="w-full border border-slate-300 bg-slate-50 rounded-lg px-3 py-2 text-xs"><option value="">All ${escape(field)}</option>${options.map(option => `<option value="${escape(option)}" ${state.filters[field] === option ? 'selected' : ''}>${escape(option)}</option>`).join('')}</select></label>`;
  }
  function allHeaders() { return headers; }
  function gridColumns() { return headers; }
  function ageDisplay(record) {
    const derived=cachedTicket(record);
    if(derived.onHold)return '';
    const color=derived.age===null?'text-slate-500':derived.age<=2?'text-emerald-700':derived.age<=5?'text-orange-700':derived.age<=10?'text-red-400':'text-rose-700';
    return `<span class="font-semibold ${color}" title="${escape(derived.age===null?'Check Created Time or Completed Time':`${derived.age} days`)}">${derived.age??'-'}</span>`;
  }
  function ageingBadge(record) {
    const derived=cachedTicket(record),index=ageingBuckets.indexOf(derived.bucket);
    if(index<0)return derived.onHold?'':'<span class="text-slate-500">-</span>';
    const color=['bg-emerald-100 text-emerald-700','bg-orange-100 text-orange-700','bg-red-100 text-red-600','bg-rose-100 text-rose-700'][index];
    return `<span class="inline-flex px-2 py-1 text-[10px] font-semibold rounded border border-current/10 ${color}">${escape(derived.bucket)}</span>`;
  }
  function badge(text, priority = false) {
    const color = priority && /critical|urgent|high/i.test(text) ? 'bg-rose-100 text-rose-700' : /closed|resolved|completed/i.test(text) ? 'bg-emerald-100 text-emerald-700' : 'bg-indigo-100 text-indigo-700';
    return `<span class="inline-flex px-2 py-1 text-[10px] font-semibold rounded border border-current/10 ${color}">${escape(text || 'Not specified')}</span>`;
  }
  function ticketCard(record) {
    const index = state.records.indexOf(record), open = state.openCards.has(record.request_id);
    return `<article class="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden"><div class="px-4 py-4 border-b border-slate-100 flex items-start justify-between gap-3"><div class="flex gap-3 min-w-0">${command('expand-card', open ? 'chevron-down' : 'chevron-right', open ? 'Collapse ticket' : 'Expand ticket', `data-index="${index}" aria-expanded="${open}"`)}<div class="min-w-0"><div class="flex flex-wrap items-center gap-2 mb-1"><span class="text-xs font-bold text-indigo-700">${escape(record.request_id)}</span>${badge(value(record,'Stage'))}${badge(status(record))}${badge(value(record, 'Priority'), true)}</div><h2 class="text-sm font-bold text-slate-800 break-words">${escape(value(record, 'Subject'))}</h2><div class="text-xs text-slate-500 mt-1 break-words">${escape(value(record, 'Request Type'))} | ${escape(value(record, 'Location'))} | ${escape(value(record, 'Region'))}</div></div></div><div class="flex shrink-0 gap-1">${command('view', 'eye', 'View ticket', `data-index="${index}"`)}${editable() ? command('edit', 'pencil', 'Edit ticket', `data-index="${index}"`) : ''}</div></div><div class="grid grid-cols-2 md:grid-cols-4 gap-4 px-4 py-3 text-xs">${['Group','Technician','Age in Days','Ageing Bucket'].map(field => `<div class="min-w-0"><div class="text-[10px] uppercase font-semibold text-slate-400">${escape(field)}</div><div class="mt-1 text-slate-700 break-words">${field==='Age in Days'?ageDisplay(record):field==='Ageing Bucket'?ageingBadge(record):escape(value(record, field) || '-')}</div></div>`).join('')}</div>${open ? `<dl class="support-details border-t border-slate-100 p-4">${gridColumns([record]).map(field => `<div><dt class="text-[10px] uppercase text-slate-400 font-semibold">${escape(field)}</dt><dd class="mt-1 text-xs text-slate-700 whitespace-pre-wrap break-words">${escape(value(record, field) || '-')}</dd></div>`).join('')}</dl>` : ''}</article>`;
  }
  function ticketGrid(records) {
    const columns = gridColumns();
    return `<div class="bg-white rounded-xl border shadow-sm overflow-hidden"><div class="overflow-x-auto custom-scrollbar"><table class="min-w-full text-xs"><thead class="bg-slate-100"><tr><th class="px-3 py-3"><input type="checkbox" data-support-select-page aria-label="Select all visible tickets" ${records.length&&records.every(record=>state.selected.has(record.request_id))?'checked':''}></th>${columns.map(field=>`<th class="px-3 py-3 text-left whitespace-nowrap" aria-sort="${state.sort===field?state.direction===1?'ascending':'descending':'none'}"><button data-support-action="sort" data-field="${escape(field)}" class="inline-flex items-center gap-1 font-semibold">${escape(field)}${icon(state.sort===field?state.direction===1?'arrow-up':'arrow-down':'arrow-up-down')}</button></th>`).join('')}<th class="px-3 py-3 text-left">Actions</th></tr></thead><tbody>${records.map(record=>`<tr class="border-t hover:bg-slate-50"><td class="px-3 py-3"><input type="checkbox" data-support-select="${escape(record.request_id)}" aria-label="Select ticket ${escape(record.request_id)}" ${state.selected.has(record.request_id)?'checked':''}></td>${columns.map(field=>`<td class="px-3 py-3 ${field==='Subject'?'min-w-[240px] max-w-[400px] whitespace-normal break-words':'whitespace-nowrap'}">${field==='Status'?badge(status(record)):escape(value(record,field))}</td>`).join('')}<td class="px-3 py-3"><div class="flex gap-1">${command('view','eye','View ticket',`data-index="${state.records.indexOf(record)}"`)}${editable()?command('edit','pencil','Edit ticket',`data-index="${state.records.indexOf(record)}"`):''}</div></td></tr>`).join('')}</tbody></table></div></div>`;
  }
  function renderResults() {
    if (!active()) return;
    if(state.remote&&state.remoteKey!==JSON.stringify(pageArguments())){if(state.pendingKey!==JSON.stringify(pageArguments()))load(true);return;}
    clearTimeout(searchTimer);
    const tableScroll = document.querySelector('#support-results .overflow-x-auto')?.scrollLeft || 0;
    const records = filtered(), ages = records.map(record => cachedTicket(record).age).filter(days => days !== null);
    const metrics = [['Total Tickets', records.length, 'text-slate-900'],['Open Tickets', records.filter(record => !closed(record)).length, 'text-indigo-600'],['Closed Tickets', records.filter(closed).length, 'text-emerald-600'],['High / Critical', records.filter(record => /high|critical|urgent/i.test(value(record, 'Priority'))).length, 'text-rose-600'],['Unassigned', records.filter(record => !value(record, 'Technician')).length, 'text-amber-600'],['Average Age (Days)', ages.length ? (ages.reduce((total, days) => total + days, 0) / ages.length).toFixed(1) : '-', 'text-sky-600']];
    if(state.remote&&state.metrics){const remote=state.metrics;[remote.total,remote.open,remote.closed,remote.high,remote.unassigned,remote.average_age===null?'-':Number(remote.average_age).toFixed(1)].forEach((count,index)=>metrics[index][1]=count);}
    document.getElementById('support-metrics').innerHTML = metrics.map(([label, count, color]) => `<div class="bg-white p-4 rounded-xl border shadow-sm min-w-0"><div class="text-[10px] uppercase text-slate-500 font-semibold">${label}</div><div class="text-2xl font-bold ${color}">${count}</div></div>`).join('');
    const summaryGroups=state.view === 'summary' ? state.remote?state.groups.map(group=>({...group,ages:group.average_age===null?[]:[Number(group.average_age)]})):summaries(records,state.summary) : null;
    const totalCount=state.remote?(summaryGroups?state.groupTotal:state.total):summaryGroups ? summaryGroups.length : records.length;
    const pages = Math.max(1, Math.ceil(totalCount / state.size)); state.page = Math.min(state.page, pages);
    const pageRecords = state.remote?records:records.slice((state.page - 1) * state.size, state.page * state.size), host = document.getElementById('support-results');
    if (state.view === 'summary') {
      const groups = state.remote?summaryGroups:summaryGroups.slice((state.page - 1) * state.size,state.page * state.size);
      host.innerHTML = `<section class="space-y-3"><div class="flex flex-wrap items-center justify-between gap-2"><h2 class="text-sm font-bold text-slate-800">${escape(state.summary)} Summary</h2><select data-support-summary class="border rounded-lg px-3 py-2 text-xs" aria-label="Summary dimension">${dimensions.map(field => `<option ${field === state.summary ? 'selected' : ''}>${field}</option>`).join('')}</select></div><div class="overflow-x-auto bg-white border rounded-xl"><table class="w-full text-xs"><thead class="bg-slate-100 text-left"><tr>${[state.summary,'Total','Open','Closed','Average Aging (Days)'].map(field => `<th class="px-4 py-3">${escape(field)}</th>`).join('')}</tr></thead><tbody>${groups.map(group => `<tr class="border-t"><td class="px-4 py-3 whitespace-pre-wrap break-words">${escape(group.name)}</td><td class="px-4 py-3">${group.total}</td><td class="px-4 py-3">${group.open}</td><td class="px-4 py-3">${group.closed}</td><td class="px-4 py-3">${group.ages.length ? (group.ages.reduce((total, days) => total + days, 0) / group.ages.length).toFixed(1) : '-'}</td></tr>`).join('')}</tbody></table></div></section>`;
    } else host.innerHTML = pageRecords.length ? state.view === 'grid' ? ticketGrid(pageRecords) : pageRecords.map(ticketCard).join('') : `<div class="bg-white border rounded-xl p-12 text-center text-sm text-slate-500">${state.loading ? 'Loading support tickets...' : 'No tickets match the selected filters.'}</div>`;
    document.getElementById('support-pagination').innerHTML = `<div class="text-xs text-slate-500">${totalCount ? (state.page - 1) * state.size + 1 : 0}-${Math.min(state.page * state.size, totalCount)} of ${totalCount} ${summaryGroups ? 'groups' : 'tickets'}</div><div class="flex flex-wrap items-center gap-2"><select data-support-size aria-label="Tickets per page" class="text-xs border rounded-lg px-2 py-2">${[...new Set([10,25,50,100,state.size])].sort((left,right)=>left-right).map(size => `<option value="${size}" ${size === state.size ? 'selected' : ''}>${size} / page</option>`).join('')}</select>${command('previous', 'chevron-left', 'Previous page', state.page === 1 ? 'disabled' : '')}<span class="text-xs text-slate-600">${state.page} / ${pages}</span>${command('next', 'chevron-right', 'Next page', state.page === pages ? 'disabled' : '')}</div>`;
    const sortControl=document.querySelector('[data-support-sort]');if(sortControl)sortControl.value=state.sort;
    const direction=document.querySelector('[data-support-action="direction"]');if(direction)direction.innerHTML=icon(state.direction === 1 ? 'arrow-up' : 'arrow-down');
    const reset = document.querySelector('[data-support-action="reset"]'); if (reset) reset.disabled = !state.query && !Object.values(state.filters).some(Boolean);
    const table = document.querySelector('#support-results .overflow-x-auto'); if (table) table.scrollLeft = tableScroll;
    lucide.createIcons();
  }
  function render() {
    if (!active() || !root()) return;
    const focused=document.activeElement,retainFocus=root().contains(focused);
    const focusAttributes=retainFocus?['id','data-support-filter','data-support-action','data-field','data-index','data-view','data-support-sort','data-support-size','data-support-summary'].filter(name=>focused.hasAttribute(name)).map(name=>[name,focused.getAttribute(name)]):[];
    const selection=retainFocus&&typeof focused.selectionStart==='number'?[focused.selectionStart,focused.selectionEnd,focused.selectionDirection]:null;
    const tableScroll=document.querySelector('#support-results .overflow-x-auto')?.scrollLeft||0;
    const scrollPositions=[];
    for(let element=root();element;element=element.parentElement)if(element.scrollLeft||element.scrollTop)scrollPositions.push([element,element.scrollLeft,element.scrollTop]);
    const scrollX=window.scrollX,scrollY=window.scrollY,previousUser=state.user;
    const user = String(currentUser?.email || currentUser?.id || '');
    if (state.user !== user) { state.user=user; state.records=[]; state.loaded=false; state.error=''; state.query=''; state.filters={}; state.openCards.clear(); state.preview=null;state.options=null;state.remote=false;state.selected.clear(); }
    document.getElementById('delivery-pagination-bottom')?.classList.add('hidden');
    const primary = ['Request Type','Request Status','Location','Priority'];
    root().innerHTML = `<div id="support-metrics" class="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-6 gap-3"></div><div id="support-filter-banner" class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm space-y-2"><div class="support-primary"><input id="support-search" type="search" value="${escape(state.query)}" placeholder="Search tickets..." aria-label="Search support tickets" class="w-full border border-slate-300 bg-slate-50 rounded-lg px-3 py-2 text-xs">${primary.map(filterControl).join('')}<div class="flex gap-2">${command('reset', 'rotate-ccw', 'Reset filters')}${command('filters', state.expanded ? 'chevron-up' : 'chevron-down', state.expanded ? 'Collapse filters' : 'Expand filters', `aria-expanded="${state.expanded}" aria-controls="support-secondary"`)}</div></div><div id="support-secondary" class="support-secondary" ${state.expanded ? '' : 'hidden'}>${headers.filter(field=>!primary.includes(field)).map(filterControl).join('')}</div></div><div class="bg-white px-4 py-3 rounded-xl border border-slate-200 shadow-sm flex flex-wrap items-center justify-between gap-3"><div role="tablist" aria-label="Support ticket views" class="flex gap-1">${[['grid','table-2','Table View'],['cards','layout-list','Card View'],['summary','chart-no-axes-combined','Summary Reporting']].map(([view,name,label]) => `<button type="button" role="tab" aria-selected="${state.view === view}" data-support-action="view-mode" data-view="${view}" title="${label}" aria-label="${label}" class="support-icon-button rounded-lg border ${state.view === view ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-white text-slate-600 border-slate-300'}">${icon(name)}</button>`).join('')}</div><div class="flex items-center gap-2"><select data-support-sort aria-label="Sort tickets by" class="text-xs border rounded-lg px-3 py-2">${headers.map(field => `<option ${field === state.sort ? 'selected' : ''}>${field}</option>`).join('')}</select>${command('direction', state.direction === 1 ? 'arrow-up' : 'arrow-down', 'Reverse sort direction')}${command('refresh', 'refresh-cw', 'Refresh tickets')}</div></div>${state.error ? `<div role="alert" class="border border-rose-200 bg-rose-50 text-rose-700 rounded-lg p-4 text-xs">${escape(state.error)}${/support_tickets|schema cache/i.test(state.error) ? '<br>Apply the Support Ticket Tracker database migration, then refresh.' : ''}</div>` : ''}<div id="support-results" class="space-y-4"></div><div id="support-pagination" class="flex flex-wrap items-center justify-between gap-3 bg-white border rounded-xl px-4 py-3"></div>`;
    renderResults();
    const sortSelect=root().querySelector('[data-support-sort]');
    sortSelect.replaceChildren(...gridColumns().map(field=>new Option(field,field,false,state.sort===field)));
    const tabular=root().querySelector('[data-view="grid"]');tabular.title='Table View';tabular.setAttribute('aria-label','Table View');
    document.getElementById('support-secondary').insertAdjacentHTML('afterbegin',filterControl('Stage'));
    if(previousUser===user){
      const target=focusAttributes.length?[...root().querySelectorAll('input,select,button,textarea')].find(control=>focusAttributes.every(([name,value])=>control.getAttribute(name)===value)):null;
      target?.focus({preventScroll:true});
      if(target&&selection&&typeof target.setSelectionRange==='function')target.setSelectionRange(...selection);
      const table=document.querySelector('#support-results .overflow-x-auto');if(table)table.scrollLeft=tableScroll;
      scrollPositions.forEach(([element,left,top])=>element.scrollTo(left,top));
      window.scrollTo(scrollX,scrollY);
    }
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
    return `<div id="support-derived-preview" class="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-5 pb-4 border-b" aria-live="polite"><div><div class="text-[10px] uppercase text-slate-500">Stage</div>${badge(value(record,'Stage'))}</div><div><div class="text-[10px] uppercase text-slate-500">Age in Days</div>${ageDisplay(record)}</div><div><div class="text-[10px] uppercase text-slate-500">Ageing Bucket</div>${ageingBadge(record)}</div></div>`;
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
    if (!requestId) return showToast('RequestID is required.','error');
    fields.RequestID=requestId;
    for(const field of headers)if(fields[field]===undefined||fields[field]===null)fields[field]='';
    state.busy = true; const host = document.getElementById('support-ticket-dialog'); host.dataset.busy = 'true'; form.querySelector('button[type="submit"]').disabled = true;
    try {
      const query = record ? supabaseClient.from('support_tickets').update({ fields }).eq('request_id',record.request_id).eq('version',record.version) : supabaseClient.from('support_tickets').insert({ request_id: requestId, fields });
      const { data, error } = await query.select('*').single();
      if (error) throw new Error(error.code === '23505' ? 'A ticket with this RequestID already exists.' : error.code === 'PGRST116' ? 'This ticket changed elsewhere. Refresh before editing.' : error.message);
      if (record) state.records[index] = data; else state.records.push(data);
      filteredTickets = null;
      host.dataset.busy = 'false'; closeDialog(); render(); showToast('Support ticket saved.','success');
    } catch (error) { showToast(error.message,'error'); }
    finally { state.busy = false; if (host.isConnected) { host.dataset.busy = 'false'; form.querySelector('button[type="submit"]').disabled = false; } }
  }
  const headerKey = input => String(input ?? '').replace(/^\uFEFF/,'').trim().toLowerCase().replace(/[^a-z0-9]/g,'');
  function parseWorkbook(workbook,sourceName='',sourceFileIndex=0) {
    const known = new Map(headers.map(header => [headerKey(header),header])), rows = [], errors = [];
    let foundTicketTable=false;
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName]; if (!sheet['!ref']) continue;
      const range = XLSX.utils.decode_range(sheet['!ref']); let headerRow = -1, names = [], bestMatchCount=1;
      for (let row = range.s.r; row <= Math.min(range.e.r,range.s.r + 100); row++) {
        const candidate = [];
        for (let column = range.s.c; column <= range.e.c; column++) candidate.push(String(sheet[XLSX.utils.encode_cell({r:row,c:column})]?.v ?? ''));
        const matchCount=candidate.filter(name=>known.has(headerKey(name))).length;
        if(matchCount>bestMatchCount){headerRow=row;names=candidate;bestMatchCount=matchCount;}
        if(matchCount===headers.length)break;
      }
      if (headerRow < 0||(!names.some(name=>headerKey(name)==='requestid')&&bestMatchCount<10)) continue;
      foundTicketTable=true;
      const counts = new Map(); names.forEach(name => { const key = headerKey(name); if(key)counts.set(key,(counts.get(key)||0)+1); });
      const missing=headers.filter(header=>!counts.has(headerKey(header)));
      const duplicated=[...counts].filter(([,count])=>count>1).map(([key])=>known.get(key)||key);
      const unrecognized=names.filter(name=>headerKey(name)&&!known.has(headerKey(name)));
      if(missing.length||duplicated.length||unrecognized.length){
        const details=[];
        if(missing.length)details.push(`missing required headers: ${missing.join(', ')}`);
        if(duplicated.length)details.push(`duplicated headers: ${duplicated.join(', ')}`);
        if(unrecognized.length)details.push(`unrecognized headers: ${unrecognized.join(', ')}`);
        errors.push(`${sourceName?`${sourceName} - `:''}${sheetName}: ${details.join('; ')}.`);
        continue;
      }
      for (let row = headerRow + 1; row <= range.e.r; row++) {
        const fields = Object.assign(Object.create(null),Object.fromEntries(headers.map(header => [header,'']))), cells = []; let populated = false;
        names.forEach((name, offset) => {
          const column = range.s.c + offset, cell = sheet[XLSX.utils.encode_cell({r:row,c:column})];
          if (cell?.v !== undefined && cell.v !== null && String(cell.v) !== '') populated = true;
          const field = known.get(headerKey(name));
          const text = cell ? cell.w !== undefined ? String(cell.w) : String(XLSX.utils.format_cell(cell)) : '';
          fields[field] = text;
          cells.push({ column: column + 1, header: field, ...(cell || {}) });
        });
        if (populated) {
          const rowErrors=[];
          if(!String(fields.RequestID??'').trim())rowErrors.push('RequestID is blank.');
          for(const field of dateHeaders){const error=templateDateError({fields,source_cells:{cells,date1904:Boolean(workbook.Workbook?.WBProps?.date1904)}},field);if(error)rowErrors.push(error);}
          rows.push({fields,source_sheet:sheetName,source_row:row+1,source_filename:sourceName,source_file_index:sourceFileIndex,row_errors:rowErrors,source_cells:{headers:headers.slice(),cells,date1904:Boolean(workbook.Workbook?.WBProps?.date1904)}});
        }
      }
    }
    if(errors.length)throw new Error(errors.join('\n'));
    if(!foundTicketTable)throw new Error(`${sourceName?`${sourceName}: `:''}no ticket table was found. The worksheet must contain all 40 required headers.`);
    if(!rows.length)throw new Error(`${sourceName?`${sourceName}: `:''}the ticket table contains no populated data rows.`);
    return rows;
  }
  async function parseFile(file,sourceFileIndex,onProgress=()=>{}) {
    onProgress(0,{phase:'Reading file'});
    const bytes=new Uint8Array(await file.arrayBuffer());
    const parseLocally=async()=>{
      onProgress(0,{phase:'Decoding workbook with browser parser'});
      await new Promise(resolve=>setTimeout(resolve,0));
      const workbook=XLSX.read(bytes,{type:'array',cellNF:true,cellStyles:true,cellFormula:true,cellText:true,cellDates:false});
      const parsedRows=parseWorkbook(workbook,file.name,sourceFileIndex);
      onProgress(parsedRows.length,{phase:'Workbook validated',total:parsedRows.length});
      return {bytes,rows:parsedRows};
    };
    if(typeof Worker==='undefined'||!supportScriptUrl)return parseLocally();
    let worker,timer;
    const rows=[];
    try{
      onProgress(0,{phase:'Starting workbook parser'});
      worker=new Worker(new URL('support-ticket-import-worker.js',supportScriptUrl));
      const result=await new Promise((resolve,reject)=>{
        const armTimeout=()=>{clearTimeout(timer);timer=setTimeout(()=>reject(new Error('Workbook parser stopped responding.')),15000);};
        armTimeout();
        worker.onmessage=event=>{
          armTimeout();
          if(event.data.type==='rows'){rows.push(...event.data.rows);onProgress(rows.length,{phase:'Validating rows',...event.data});}
          else if(event.data.type==='progress')onProgress(rows.length,event.data);
          else if(event.data.type==='error')reject(Object.assign(new Error(event.data.message),{validationError:true}));
          else if(event.data.type==='done')resolve();
        };
        worker.onerror=event=>reject(new Error(event.message||`Unable to parse ${file.name} in the import worker.`));
        const buffer=bytes.slice().buffer;
        worker.postMessage({buffer,name:file.name,fileIndex:sourceFileIndex},[buffer]);
      });
      void result;
      return {bytes,rows};
    }catch(error){
      if(error.validationError)throw error;
      worker?.terminate();clearTimeout(timer);
      return await parseLocally();
    }finally{clearTimeout(timer);worker?.terminate();}
  }
  function classifyRows(rows) {
    const results=rows.map(row=>({request_id:String(row.fields.RequestID??'').trim(),sheet:row.source_sheet,row:row.source_row,source_filename:row.source_filename,source_file_index:row.source_file_index,status:'new',message:(row.row_errors||[]).join(' ')}));
    results.forEach((result,index)=>{if(rows[index].row_errors?.length)result.status='rejected';});
    const groups=new Map();
    rows.forEach((row,index)=>{const id=normalizedId(row.fields.RequestID);if(!id)return;const group=groups.get(id)||[];group.push(index);groups.set(id,group);});
    for(const indexes of groups.values()){
      const byFile=new Map();
      for(const index of indexes){const fileIndex=rows[index].source_file_index??0;const group=byFile.get(fileIndex)||[];group.push(index);byFile.set(fileIndex,group);}
      for(const sameFile of byFile.values())if(sameFile.length>1)for(const index of sameFile)Object.assign(results[index],{status:'rejected',message:[results[index].message,'RequestID is duplicated within the uploaded file.'].filter(Boolean).join(' ')});
      const distinct=[...byFile.values()].filter(group=>group.length===1).flat().filter(index=>results[index].status!=='rejected');
      if(distinct.length>1){
        const ranked=distinct.map(index=>({index,date:dateTimeRank(rows[index],'Last Updated Time')}));
        const latest=Math.max(...ranked.map(item=>item.date??-Infinity));
        const winners=ranked.filter(item=>(item.date??-Infinity)===latest);
        if(winners.length!==1)for(const item of ranked)Object.assign(results[item.index],{status:'conflicted',message:'RequestID appears in multiple files with equal Last Updated Time; resolve the conflict and upload again.'});
        else for(const item of ranked)if(item.index!==winners[0].index)Object.assign(results[item.index],{status:'rejected',message:`Superseded by the row with the latest Last Updated Time in ${rows[winners[0].index].source_filename||'another uploaded file'}.`});
      }
    }
    const existing=new Map(state.records.map(record=>[normalizedId(record.request_id),record]));
    for(let index=0;index<rows.length;index++){
      const row=rows[index],result=results[index];
      if(result.status==='rejected'||result.status==='conflicted')continue;
      if(row.row_errors?.length){result.status='rejected';continue;}
      const current=existing.get(normalizedId(row.fields.RequestID));
      if(!current){result.status='new';continue;}
      const changed=headers.some(field=>field==='RequestID'
        ?normalizedId(current.fields?.[field])!==normalizedId(row.fields[field])
        :String(current.fields?.[field]??'')!==String(row.fields[field]??''));
      result.status=changed?'changed':'unchanged';
    }
    return {candidates:rows.filter((row,index)=>['new','changed','unchanged'].includes(results[index].status)),results};
  }
  function importErrorMessage(error) {
    const message=String(error?.message || error || 'Unknown import error.');
    const schemaVersion=Number(error?.schemaVersion);
    if(error?.code==='PGRST205'&&Number.isInteger(schemaVersion)&&schemaVersion<3)return `Connected Supabase reports Support Ticket schema version ${schemaVersion}; this importer requires version 3. Run the complete supabase-migration-support-ticket-tracker.sql through COMMIT in the same Supabase project, then choose Check Database Again. The file defines an older version earlier and upgrades it later.`;
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
    if(!schema || schema.version<3 || !schema.atomic || !schema.imports_table || !schema.tickets_table || !schema.normalized_request_id)throw {code:'PGRST205',schemaVersion:Number.isInteger(Number(schema?.version))?Number(schema.version):null,message:'Support Ticket import schema is outdated.'};
    for(const [table,columns] of [
      ['support_ticket_imports','id,file_name,file_size,file_base64,workbook,validation_summary,status,total_rows,processed_rows,staged_rows,error_message,inserted_rows,updated_rows,unchanged_rows,rejected_rows,conflicted_rows'],
      ['support_tickets','request_id,normalized_request_id,fields,source_cells,source_import_id,source_sheet,source_row,source_filename,source_file_index,version']
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
    const preview = state.preview, counts = {inserted:0,updated:0,unchanged:0,rejected:0,conflicted:0,new:0,changed:0};
    preview.results.forEach(row => counts[row.status] = (counts[row.status] || 0) + 1);
    const canImport=!preview.finished&&preview.schemaReady&&!preview.error&&!preview.uncertain&&preview.candidates.length>0;
    const setupActions=!preview.schemaReady&&!preview.finished?`<div class="flex flex-wrap gap-2 mb-4"><button type="button" data-support-action="schema-sql" class="inline-flex items-center gap-2 border rounded-lg px-3 py-2 text-xs">${icon('file-down')}Download Database Migration</button><button type="button" data-support-action="schema-check" class="inline-flex items-center gap-2 border rounded-lg px-3 py-2 text-xs">${icon('refresh-cw')}Check Database Again</button></div>`:'';
    const metrics=preview.finished
      ?[['Total',preview.results.length],['Inserted',counts.inserted],['Updated',counts.updated],['Unchanged',counts.unchanged],['Rejected',counts.rejected],['Conflicted',counts.conflicted]]
      :[['Total Valid Rows',counts.new+counts.changed+counts.unchanged],['New Tickets',counts.new],['Changed Tickets',counts.changed],['Unchanged Tickets',counts.unchanged],['Rejected Rows',counts.rejected],['Conflicted Rows',counts.conflicted],['Validation Errors',preview.results.filter(row=>row.message).length]];
    const metricTitle=preview.finished?'Total | Inserted | Updated | Unchanged | Rejected | Conflicted':'Import preview';
    return `<div class="text-xs text-slate-500 mb-4 break-words">${escape(preview.names.join(', '))}</div>${progressMarkup()}${preview.error?`<div role="alert" class="border border-rose-200 bg-rose-50 text-rose-700 rounded-lg p-3 text-xs mb-4 whitespace-pre-wrap">${escape(preview.error)}</div>`:''}${setupActions}<div class="text-[10px] font-semibold uppercase tracking-wide text-slate-500 mb-2">${metricTitle}</div><div class="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-7 gap-3 mb-4">${metrics.map(([label,count])=>`<div class="border rounded-lg p-3"><div class="uppercase text-[10px] text-slate-500">${label}</div><div class="text-xl font-bold text-slate-800">${count}</div></div>`).join('')}</div><div class="overflow-x-auto border rounded-lg max-h-80"><table class="w-full text-xs"><thead class="bg-slate-100 sticky top-0 text-left"><tr>${['Source file','Sheet','Row','RequestID','Result','Details'].map(label=>`<th class="p-3">${label}</th>`).join('')}</tr></thead><tbody>${preview.results.slice(0,100).map(row=>`<tr class="border-t"><td class="p-3">${escape(row.source_filename)}</td><td class="p-3">${escape(row.sheet)}</td><td class="p-3">${row.row??''}</td><td class="p-3">${escape(row.request_id)}</td><td class="p-3">${escape(row.status)}</td><td class="p-3">${escape(row.message)}</td></tr>`).join('')}</tbody></table></div><div class="text-xs text-slate-500 mt-2">${preview.results.length} rows; first 100 displayed.</div><div class="flex flex-wrap justify-end gap-2 mt-5">${command('validation-export','download','Download row-level validation report')}<button data-support-action="close-dialog" class="border rounded-lg px-4 py-2 text-xs">Close</button>${canImport?`<button data-support-action="commit-import" class="bg-indigo-600 text-white rounded-lg px-4 py-2 text-xs font-semibold inline-flex items-center gap-2">${icon('upload')}Confirm Import (${preview.candidates.length} rows)</button>`:''}</div>`;
  }
  function openImport() {
    if (!editable()) return showToast('Permission denied.','error');
    if(window.supportTicketBulkImport)return supportTicketBulkImport.open();
    dialog('Import Support Tickets', `<div class="space-y-4"><input id="support-import-file" type="file" accept=".xls,.xlsx" multiple aria-label="Support ticket workbooks" class="w-full text-xs border rounded-lg p-3"><p class="text-xs text-slate-500">Select either Open or Closed template, or both together. Ticket state is read from each row.</p><div class="flex justify-end"><button data-support-action="template" class="border rounded-lg px-4 py-2 text-xs inline-flex items-center gap-2">${icon('file-down')}Download Template</button></div><div id="support-import-preview" aria-live="polite">${state.preview?importReport():''}</div></div>`);
  }
  async function recheckImportSchema() {
    const preview=state.preview;if(!editable()||!preview||preview.finished||state.busy)return;
    const host=document.getElementById('support-ticket-dialog');state.busy=true;if(host)host.dataset.busy='true';
    try{
      await ensureImportSchema();await load(true);if(state.error)throw new Error(state.error);
      const classification=classifyRows(preview.rows);
      Object.assign(preview,classification,{schemaReady:true,error:''});
      updateSupportImportProgress(0,preview.rows.length,'Ready to validate RequestIDs');
    }catch(error){preview.schemaReady=false;preview.error=importErrorMessage(error);updateSupportImportProgress(0,preview.rows.length,'Database check blocked; not processed');}
    finally{
      state.busy=false;if(host)host.dataset.busy='false';
      const area=document.getElementById('support-import-preview');if(area)area.innerHTML=importReport();lucide.createIcons();
    }
  }
  async function previewFile(files) {
    const selected=Array.from(files?.length!==undefined&&typeof files!=='string'?files:[files]).filter(Boolean);
    if (!selected.length || state.busy) return;
    const previewHost = document.getElementById('support-import-preview'),host=document.getElementById('support-ticket-dialog'),input=document.getElementById('support-import-file');
    state.preview=null;state.busy=true;if(host)host.dataset.busy='true';if(input)input.disabled=true;
    previewHost.innerHTML='<div role="status" aria-live="polite"><div id="support-reading-detail" class="text-sm mb-2"></div><progress id="support-reading-progress" aria-label="Workbook reading and validation progress" class="w-full"></progress><div id="support-reading-elapsed" class="text-xs text-slate-500 mt-2"></div></div>';
    const started=Date.now(),elapsed=previewHost.querySelector('#support-reading-elapsed');
    const tick=()=>{elapsed.textContent=`Elapsed: ${Math.floor((Date.now()-started)/1000)} seconds. No ticket changes saved.`;};
    tick();const clock=setInterval(tick,1000);
    const readingProgress=(file,index,count,info={})=>{
      previewHost.querySelector('#support-reading-detail').textContent=`File ${index+1} of ${selected.length}: ${file.name} - ${info.phase||'Validating rows'}${info.total?`: ${info.processed??count} of ${info.total} rows (${Math.round((info.processed??count)/info.total*100)}%)`:count?`: ${count} rows parsed`:''}`;
      const progress=previewHost.querySelector('#support-reading-progress');
      if(info.total){progress.max=info.total;progress.value=info.processed??count;}else progress.removeAttribute('value');
    };
    try {
      const sources=[],rows=[];
      for(let index=0;index<selected.length;index++){
        const file=selected[index];
        if(!/\.xlsx?$/i.test(file.name))throw new Error(`${file.name}: only legacy .xls or template-compatible .xlsx files are supported.`);
        const parsed=await parseFile(file,index,(count,info)=>readingProgress(file,index,count,info));
        rows.push(...parsed.rows);
        sources.push({name:file.name,size:file.size,bytes:parsed.bytes});
        await new Promise(resolve=>setTimeout(resolve,0));
      }
      const classification=classifyRows(rows);
      state.preview = {names:sources.map(file=>file.name),size:sources.reduce((total,file)=>total+file.size,0),sources,rows,...classification,finished:false,schemaReady:false,uncertain:false,importId:crypto.randomUUID(),progress:{processed:0,total:rows.length,label:'Preparing import'}};
      clearInterval(clock);state.busy=false;
      readingProgress(selected[selected.length-1],selected.length-1,rows.length,{phase:'Checking database and preparing preview',total:rows.length});
      await recheckImportSchema();
      previewHost.innerHTML = importReport(); lucide.createIcons();
    } catch (error) { state.preview = null; previewHost.innerHTML=`<div role="alert" class="border border-rose-200 bg-rose-50 text-rose-700 rounded-lg p-3 text-xs whitespace-pre-wrap">${escape(error.message)}</div>`; }
    finally{clearInterval(clock);state.busy=false;if(host)host.dataset.busy='false';if(input)input.disabled=false;}
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
    if(!editable())return;
    if(!window.supportTicketBulkImport)return showToast('Reload the portal to load the batch import controller.','error');
    return supportTicketBulkImport.run();
  }
  function openExportDialog() {
    const selected=state.selected.size;
    const choices=[['filtered','Filtered tickets'],['selected',`Selected tickets (${selected})`],['open','Open tickets'],['closed','Closed tickets'],['all','All tickets']];
    dialog('Export Support Tickets',`<div class="space-y-3"><p class="text-xs text-slate-500">The export contains exactly the 40 template columns. Current filters are applied only to Filtered tickets.</p><div class="grid grid-cols-1 sm:grid-cols-2 gap-2">${choices.map(([scope,label])=>`<button type="button" data-support-action="export-scope" data-scope="${scope}" ${scope==='selected'&&!selected?'disabled':''} class="border rounded-lg px-4 py-3 text-left text-xs disabled:opacity-40">${escape(label)}</button>`).join('')}</div></div>`);
  }
  function exportWorkbook(template = false, validation = false, scope = 'filtered') {
    if(state.remote&&!template&&!validation)return exportRemote(scope).catch(error=>showToast(error.message,'error'));
    const workbook = XLSX.utils.book_new();
    if (validation) {
      const reportRows=(state.preview?.results||[]).map(({source_filename,sheet,row,request_id,status,message})=>({SourceFile:source_filename||'',Sheet:sheet||'',Row:row??'',RequestID:request_id||'',Result:status||'',Details:message||''}));
      XLSX.utils.book_append_sheet(workbook,XLSX.utils.json_to_sheet(reportRows),'Validation Results');
    } else {
      let records=[];
      if(!template){
        if(scope==='selected')records=state.records.filter(record=>state.selected.has(record.request_id));
        else if(scope==='open')records=state.records.filter(record=>!closed(record));
        else if(scope==='closed')records=state.records.filter(closed);
        else if(scope==='all')records=state.records.slice();
        else records=filtered();
      }
      const columns=headers;
      const sheet = XLSX.utils.aoa_to_sheet([columns,...records.map(record=>columns.map(field=>value(record,field)))]);
      sheet['!cols'] = columns.map(field=>({wch:field === 'Subject' ? 55 : 24}));
      XLSX.utils.book_append_sheet(workbook,sheet,'Support Tickets');
    }
    const suffix=scope==='filtered'?'Filtered':scope[0].toUpperCase()+scope.slice(1);
    XLSX.writeFile(workbook,validation?'Support_Ticket_Validation.xlsx':template?'Support_Ticket_Import_Template.xlsx':`Support_Tickets_${suffix}_${new Date().toISOString().slice(0,10)}.xlsx`);
  }
  async function exportRemote(scope) {
    const worker=new Worker(new URL('support-ticket-import-worker.js',supportScriptUrl));
    const send=message=>new Promise((resolve,reject)=>{
      const timer=setTimeout(()=>reject(new Error('Export worker timed out.')),120000);
      worker.onerror=event=>{clearTimeout(timer);reject(new Error(event.message||'Export worker failed.'));};
      worker.onmessage=event=>{if(['export-ready','export-file','error'].includes(event.data.type)){clearTimeout(timer);if(event.data.type==='error')reject(new Error(event.data.message));else resolve(event.data);}};
      worker.postMessage(message);
    });
    try{
      await send({type:'export-start',headers});let after='';
      for(;;){
        const {data,error}=await supabaseClient.rpc('support_ticket_export_page',{p_after:after,p_query:state.query,p_filters:pageArguments().p_filters,p_scope:scope,p_selected:[...state.selected]});
        if(error)throw error;if(!data?.length)break;
        await send({type:'export-rows',rows:data.map(record=>headers.map(header=>String(record.fields[header]??'')))});
        after=data.at(-1).request_id;await new Promise(resolve=>setTimeout(resolve,25));
      }
      const {buffer}=await send({type:'export-done'}),url=URL.createObjectURL(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const link=document.createElement('a');link.href=url;link.download=`Support_Tickets_${scope}_${new Date().toISOString().slice(0,10)}.xlsx`;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);showToast('Support ticket export completed.','success');
    }finally{worker.terminate();}
  }
  async function filterOptions(field,query){const {data,error}=await supabaseClient.rpc('support_ticket_options',{p_field:field,p_query:query});if(error)throw error;return data||[];}
  async function downloadSource(index) {
    const record = state.records[index]; if (!editable() || !record?.source_import_id) return;
    const result = await supabaseClient.from('support_ticket_imports').select('file_name,file_base64,workbook').eq('id',record.source_import_id).single();
    if (result.error) return showToast(result.error.message,'error');
    const stored=result.data.workbook?.storage_files?.[record.source_file_index||0];
    if(stored){const {data,error}=await supabaseClient.storage.from('support-ticket-sources').download(stored.path);if(error)return showToast(error.message,'error');const url=URL.createObjectURL(data),link=document.createElement('a');link.href=url;link.download=stored.name;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);return;}
    const sourceFile=result.data.file_base64?{name:record.source_filename||result.data.file_name,file_base64:result.data.file_base64}:result.data.workbook?.files?.find(file=>file.name===record.source_filename);
    if(!sourceFile?.file_base64)return showToast('The original source workbook is unavailable for this ticket.','error');
    const bytes=Uint8Array.from(atob(sourceFile.file_base64),character=>character.charCodeAt(0)),url=URL.createObjectURL(new Blob([bytes]));
    const link=document.createElement('a');link.href=url;link.download=sourceFile.name;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  }
  document.addEventListener('input',event=>{
    if(active() && event.target.id === 'support-search') {
      state.query=event.target.value;state.page=1;
      clearTimeout(searchTimer);
      const user = state.user;
      searchTimer=setTimeout(()=>{if(active() && state.user===user)renderResults();},160);
    }
    const form=event.target.closest?.('#support-ticket-form');
    if(form&&event.target.hasAttribute('data-ticket-field')){
      const record=form.dataset.index===''?null:state.records[Number(form.dataset.index)],fields={...(record?.fields||{})};
      form.querySelectorAll('[data-ticket-field]').forEach(input=>{fields[input.dataset.ticketField]=input.value;});
      const preview=document.getElementById('support-derived-preview');if(preview){preview.outerHTML=derivedPreview({...record,fields});lucide.createIcons();}
    }
  });
  document.addEventListener('change',event=>{
    const control=event.target;
    if(control.id === 'support-import-file') {previewFile(control.files);return;}
    if(!active())return;
    if(control.hasAttribute('data-support-select')){
      const id=control.dataset.supportSelect;
      if(control.checked)state.selected.add(id);else state.selected.delete(id);
      return;
    }
    if(control.hasAttribute('data-support-select-page')){
      const records=state.remote?state.records:filtered().slice((state.page-1)*state.size,state.page*state.size);
      for(const record of records)if(control.checked)state.selected.add(record.request_id);else state.selected.delete(record.request_id);
      renderResults();return;
    }
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
    if(action === 'export-scope'){closeDialog();return exportWorkbook(false,false,button.dataset.scope);}
    if(action === 'open-export')return openExportDialog();
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
  window.exportCurrentWorkspacePage=function(){return active()?openExportDialog():previousExport.apply(this,arguments);};
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
  window.supportTickets = { headers, optionalHeaders, dimensions, state, load, render, filtered, summaries, agingDays, deriveTicket, gridColumns, value, parseWorkbook, classifyRows, openTicket, saveTicket, previewFile, commitImport, exportWorkbook, recheckImportSchema,filterOptions,dateTimeRank,loadErrorHint };
})();