(() => {
  const pages = new Map();
  let account = '';
  function state() {
    const user = typeof currentUser === 'undefined' ? '' : String(currentUser?.id || currentUser?.email || '');
    if (user !== account) { pages.clear(); account = user; }
    const page = typeof currentWorkspaceView === 'undefined' ? 'overview' : currentWorkspaceView;
    if (!pages.has(page)) pages.set(page, new Map());
    return pages.get(page);
  }
  function values(key, fallback) {
    if (state().has(key)) return [...state().get(key)];
    const value = fallback === undefined ? document.getElementById(key)?.value || '' : fallback;
    return Array.isArray(value) ? value.map(String) : value === '' || value === 'all' ? [] : [String(value)];
  }
  function set(key, selected) { state().set(key, [...new Set(selected.map(String))]); }
  function some(key, predicate, fallback) {
    const selected = values(key, fallback);
    return !selected.length || selected.some(predicate);
  }
  function matches(key, actual, fallback) {
    const candidates = (Array.isArray(actual) ? actual : [actual]).map(value => String(value ?? '').trim());
    return some(key, selected => candidates.includes(selected), fallback);
  }
  function clear() { state().clear(); }
  const selectors = '#overview-filter-row select,#workspace-flat-page .flat-shared-filter-banner select,#daily-tracker-controls select,#support-filter-banner select,#workspace-flat-page [data-date-range-page] select';
  let menu = null, anchor = null, badges = [], scheduled = false;
  function keyFor(select) { return select.dataset.supportFilter ? `support:${select.dataset.supportFilter}` : select.id; }
  function eligible(select) {
    const page = typeof currentWorkspaceView === 'undefined' ? '' : currentWorkspaceView;
    return ['overview','programs','initiatives','tracks','milestones','daily','support'].includes(page) && select?.matches(selectors) && Boolean(keyFor(select)) && !select.disabled;
  }
  function allOption(option) { return option.value === '' || option.value === 'all' && /^all\b/i.test(option.textContent.trim()); }
  function text(select) { return select.getAttribute('aria-label') || select.title || select.options[0]?.textContent || 'Filter'; }
  function close() {
    menu?.remove();menu=null;
    if(anchor?.isConnected)anchor.setAttribute('aria-expanded','false');
    anchor=null;
  }
  function position() {
    if(!menu||!anchor?.isConnected){close();return;}
    const rect=anchor.getBoundingClientRect();
    if(!rect.width||!rect.height){close();return;}
    const width=Math.min(Math.max(rect.width,260),innerWidth-16),height=Math.min(360,innerHeight-24),below=innerHeight-rect.bottom-8;
    const down=below>=Math.min(height,220),available=down?below:rect.top-12;
    const limit=Math.max(72,Math.min(height,available));
    menu.style.width=`${width}px`;menu.style.maxHeight=`${limit}px`;
    menu.style.left=`${Math.max(8,Math.min(rect.left,innerWidth-width-8))}px`;
    menu.style.top=`${down?rect.bottom+4:Math.max(8,rect.top-Math.min(limit,menu.scrollHeight)-4)}px`;
  }
  function notify(select,selected) {
    const key=keyFor(select);set(key,selected);
    const first=[...select.options].find(option=>selected.includes(option.value));
    const reset=[...select.options].find(allOption);
    select.value=first?.value ?? reset?.value ?? '';
    select.dispatchEvent(new Event('change',{bubbles:true}));
    refresh();
    if(menu)renderOptions();
  }
  function renderOptions() {
    if(!menu||!anchor)return;
    const select=anchor,query=menu.querySelector('input[type="search"]').value.toLowerCase(),list=menu.querySelector('[role="listbox"]'),chosen=values(keyFor(select),select.value);
    const focused=menu.contains(document.activeElement)?document.activeElement.dataset.value:undefined;
    list.replaceChildren();
    const add=(label,value,checked,isAll=false)=>{
      const row=document.createElement('label');row.className='tracker-multi-option';row.setAttribute('role','option');row.setAttribute('aria-selected',String(checked));
      const box=document.createElement('input');box.type='checkbox';box.checked=checked;box.dataset.value=value;box.setAttribute('aria-label',label);
      box.addEventListener('change',()=>{const next=values(keyFor(select),select.value);notify(select,isAll?[]:box.checked?[...next,value]:next.filter(item=>item!==value));});
      const span=document.createElement('span');span.textContent=label;row.append(box,span);list.append(row);
    };
    add('All','',!chosen.length,true);
    [...select.options].filter(option=>!allOption(option)&&!option.disabled&&(!query||option.textContent.toLowerCase().includes(query))).forEach(option=>add(option.textContent,option.value,chosen.includes(option.value)));
    if(focused!==undefined)[...list.querySelectorAll('input')].find(input=>input.dataset.value===focused)?.focus({preventScroll:true});
    position();
  }
  function open(select) {
    if(anchor===select&&menu){close();return;}
    close();anchor=select;select.focus({preventScroll:true});
    menu=document.createElement('div');menu.className='tracker-multi-menu';menu.setAttribute('role','dialog');menu.setAttribute('aria-label',text(select));
    const search=document.createElement('input');search.type='search';search.placeholder='Search';search.setAttribute('aria-label','Search filter options');search.className='tracker-multi-search';
    const list=document.createElement('div');list.setAttribute('role','listbox');list.setAttribute('aria-multiselectable','true');list.setAttribute('aria-label',text(select));
    menu.append(search,list);document.body.append(menu);select.setAttribute('aria-expanded','true');
    search.addEventListener('input',renderOptions);
    renderOptions();search.focus();
  }
  function refresh() {
    scheduled=false;
    badges.forEach(badge=>badge.remove());badges=[];
    document.querySelectorAll(selectors).forEach(select=>{
      if(!eligible(select))return;
      const key=keyFor(select),chosen=values(key,select.value);
      if(state().has(key)){
        const option=[...select.options].find(option=>chosen.includes(option.value));
        const all=[...select.options].find(allOption);
        if(option)select.value=option.value;else if(!chosen.length&&all)select.value=all.value;
      }
      select.setAttribute('data-tracker-multi-filter','true');select.setAttribute('aria-haspopup','dialog');
      const label=text(select);select.setAttribute('aria-description',chosen.length?`${chosen.length} selected: ${chosen.join(', ')}`:'All values');
      const rect=select.getBoundingClientRect();
      if(chosen.length>1&&rect.width&&rect.height){
        const badge=document.createElement('span');badge.className='tracker-multi-count';badge.textContent=String(chosen.length);badge.title=`${label}: ${chosen.join(', ')}`;badge.setAttribute('aria-hidden','true');
        badge.style.left=`${rect.right-45}px`;badge.style.top=`${rect.top+Math.max(1,(rect.height-18)/2)}px`;document.body.append(badge);badges.push(badge);
      }
    });
    if(menu&&!anchor?.isConnected){
      const key=anchor?keyFor(anchor):null;
      anchor=[...document.querySelectorAll(selectors)].find(select=>eligible(select)&&keyFor(select)===key)||null;
      if(!anchor)close();else {anchor.setAttribute('aria-expanded','true');renderOptions();}
    }else position();
  }
  function schedule() { if(!scheduled){scheduled=true;queueMicrotask(refresh);} }
  function initialize() {
    const style=document.createElement('style');style.textContent=`
      .tracker-multi-menu{position:fixed;z-index:200000;box-sizing:border-box;background:#fff;border:1px solid #cbd5e1;border-radius:8px;box-shadow:0 8px 24px #0f172a26;padding:8px;overflow:auto;font-family:inherit;font-size:12px;line-height:1.4;color:#334155}
      .tracker-multi-search{width:100%;min-height:34px;border:1px solid #cbd5e1;border-radius:6px;padding:6px 8px;margin-bottom:6px;box-sizing:border-box}
      .tracker-multi-option{display:flex;align-items:flex-start;gap:8px;padding:8px;border-radius:4px;cursor:pointer;overflow-wrap:anywhere}
      .tracker-multi-option:hover,.tracker-multi-option:focus-within{background:#eef2ff}
      .tracker-multi-option input{margin-top:2px;accent-color:#4f46e5;flex:0 0 14px}
      .tracker-multi-count{position:fixed;z-index:45;pointer-events:none;background:#e0e7ff;color:#4338ca;border:1px solid #c7d2fe;border-radius:4px;min-width:18px;height:18px;padding:0 3px;text-align:center;font-size:10px;font-weight:700;line-height:16px;box-sizing:border-box}
    `;document.head.append(style);
    document.addEventListener('pointerdown',event=>{if(event.target instanceof HTMLSelectElement&&eligible(event.target)){event.preventDefault();event.stopImmediatePropagation();open(event.target);}else if(menu&&!menu.contains(event.target))close();},true);
    document.addEventListener('keydown',event=>{
      if(event.target instanceof HTMLSelectElement&&eligible(event.target)&&['Enter',' ','ArrowDown'].includes(event.key)){event.preventDefault();event.stopImmediatePropagation();open(event.target);return;}
      if(menu&&event.key==='Escape'){event.preventDefault();event.stopImmediatePropagation();const select=anchor;close();select?.focus();}
      if(menu&&['ArrowDown','ArrowUp'].includes(event.key)){event.preventDefault();const controls=[...menu.querySelectorAll('input')],index=controls.indexOf(document.activeElement);controls[(index+(event.key==='ArrowDown'?1:-1)+controls.length)%controls.length]?.focus();}
      if(menu&&event.key==='Tab'){close();}
    },true);
    document.addEventListener('click',event=>{if(event.target.closest('.unified-filter-reset,[data-support-action="reset"],[data-daily-action="reset"],[onclick*="resetCurrentPageFilters"],[onclick*="resetProjectTrackerFilters"]')){clear();close();schedule();}},true);
    document.addEventListener('change',event=>{if(event.target instanceof HTMLSelectElement&&eligible(event.target))schedule();},true);
    for(const name of ['resetCurrentPageFilters','resetProjectTrackerFilters']){
      const previous=window[name];if(typeof previous==='function')window[name]=function(){clear();close();const result=previous.apply(this,arguments);schedule();return result;};
    }
    const observer=new MutationObserver(records=>{
      if(records.some(record=>record.type==='childList'&&[...record.addedNodes,...record.removedNodes].some(node=>node.nodeType===1&&!node.matches?.('.tracker-multi-count,.tracker-multi-menu')&&(node.matches?.('select,option,.unified-filter-secondary,.flat-shared-filter-banner,#support-filter-banner')||node.querySelector?.('select,option')))))schedule();
    });observer.observe(document.body,{childList:true,subtree:true});
    window.addEventListener('resize',schedule,{passive:true});document.addEventListener('scroll',schedule,{passive:true,capture:true});
    refresh();
  }
  window.trackerMultiFilters = { values, set, some, matches, clear, refresh, open };
  window.trackerFilterMatches = matches;
  window.trackerFilterSome = some;
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',initialize,{once:true});else initialize();
})();