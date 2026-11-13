(() => {
  const active = item => !['deleted', 'archived'].includes(String(item?.recordStatus || item?.record_status || 'active').trim().toLowerCase());
  let context = null;
  let busy = false;
  let dialog = null;
  const pendingEchoes = new Map();
  const auditFields = new Set(['updated_at','modified_at','modified_time','modified_by','modified_by_name','updated_by','updated_by_name']);

  function trackName(project, id) {
    return (project.tracks || []).find(track => Number(track.id) === Number(id))?.name || 'Unassigned Track';
  }
  function milestoneName(milestone) {
    return milestone?.actions || milestone?.title || 'Milestone';
  }
  function dependencyLabel(project, milestone) {
    if (milestone.dependencyType === 'track') return `Depends on Track: ${trackName(project, milestone.dependencyTrackId)}`;
    if (milestone.dependencyType === 'milestone') {
      const source = (project.milestones || []).find(item => Number(item.id) === Number(milestone.dependencyMilestoneId));
      return source ? `Depends on: ${trackName(project, source.trackId)} \u2192 ${milestoneName(source)}` : `Dependency unavailable (ID ${milestone.dependencyMilestoneId || '-'})`;
    }
    return '';
  }
  function ensureDialog() {
    if (dialog) return dialog;
    dialog = document.createElement('dialog');
    dialog.className = 'w-full max-w-lg rounded-lg border border-slate-200 bg-white p-5 text-slate-800 shadow-xl';
    dialog.setAttribute('aria-labelledby', 'milestone-dependency-title');
    dialog.innerHTML = '<div class="flex items-center justify-between gap-3"><h2 id="milestone-dependency-title" class="text-base font-bold">Set Dependency</h2><button type="button" data-close class="p-2 rounded hover:bg-slate-100" aria-label="Close"><i data-lucide="x" class="w-4 h-4"></i></button></div><div class="mt-3 text-sm"><div><span class="text-slate-500">Milestone: </span><strong data-current></strong></div><div><span class="text-slate-500">Track: </span><span data-current-track></span></div><div class="mt-1 text-slate-500" data-current-dependency></div></div><fieldset class="mt-4" data-dependency-editor><legend class="text-xs font-semibold">Dependency Type</legend><div class="mt-2 flex gap-4 text-sm"><label class="flex items-center gap-2"><input type="radio" name="milestone-dependency-kind" value="milestone" checked> Milestone</label><label class="flex items-center gap-2"><input type="radio" name="milestone-dependency-kind" value="track"> Track</label></div></fieldset><label data-dependency-editor class="mt-4 block text-xs font-semibold" for="milestone-dependency-target">Dependency</label><select data-dependency-editor id="milestone-dependency-target" class="mt-1 w-full rounded border border-slate-300 bg-white px-3 py-2 text-sm"></select><div data-error role="alert" class="mt-2 text-xs text-red-700" hidden></div><div class="mt-5 flex flex-wrap justify-end gap-2"><button type="button" data-remove class="rounded border border-rose-300 px-3 py-2 text-xs font-semibold text-rose-700">Remove Dependency</button><button type="button" data-cancel class="rounded border border-slate-300 px-3 py-2 text-xs font-semibold">Cancel</button><button type="button" data-save class="rounded bg-indigo-600 px-3 py-2 text-xs font-semibold text-white">Save Dependency</button></div>';
    document.body.appendChild(dialog);
    dialog.querySelector('[data-close]').onclick = close;
    dialog.querySelector('[data-cancel]').onclick = close;
    dialog.querySelector('[data-save]').onclick = () => save(false);
    dialog.querySelector('[data-remove]').onclick = () => save(true);
    dialog.querySelectorAll('input[name="milestone-dependency-kind"]').forEach(input => input.addEventListener('change', () => populate()));
    dialog.addEventListener('cancel', event => { if (busy) event.preventDefault(); });
    return dialog;
  }
  function close() {
    if (busy || !dialog) return;
    dialog.close();
    context?.origin?.focus?.({preventScroll:true});
    context = null;
  }
  function populate(selectedId) {
    if (!context) return;
    const {project, milestone} = context;
    const kind = dialog.querySelector('input[name="milestone-dependency-kind"]:checked').value;
    const select = dialog.querySelector('#milestone-dependency-target');
    select.replaceChildren(new Option(kind === 'track' ? 'Select a Track' : 'Select a Milestone', ''));
    const tracks = (project.tracks || []).filter(active);
    if (kind === 'track') {
      tracks.filter(track => Number(track.id) !== Number(milestone.trackId)).forEach(track => select.add(new Option(track.name || `Track ${track.id}`, String(track.id))));
    } else {
      const activeTrackIds = new Set(tracks.map(track => Number(track.id)));
      const rows = (project.milestones || []).filter(item => active(item) && Number(item.id) !== Number(milestone.id) && (!item.trackId || activeTrackIds.has(Number(item.trackId))));
      for (const item of rows) select.add(new Option(`${trackName(project, item.trackId)} \u2192 ${milestoneName(item)}`, String(item.id)));
    }
    if (selectedId) select.value = String(selectedId);
  }
  function open(projectId, milestoneId, origin) {
    const project = projects.find(item => Number(item.id) === Number(projectId));
    const milestone = project?.milestones?.find(item => Number(item.id) === Number(milestoneId));
    const hasDependency = Boolean(milestone?.dependencyType);
    const canEdit = project && canEditInitiative(project);
    if (!project || !milestone || !active(milestone) || (!canEdit && !hasDependency)) return showToast('Milestone unavailable or permission denied.', 'error');
    ensureDialog();
    const readOnly = hasDependency || !canEdit;
    context = {project, milestone, origin, readOnly};
    dialog.querySelector('#milestone-dependency-title').textContent = readOnly ? 'View Dependency' : 'Set Dependency';
    dialog.querySelector('[data-current]').textContent = milestoneName(milestone);
    dialog.querySelector('[data-current-track]').textContent = trackName(project, milestone.trackId);
    dialog.querySelector('[data-current-dependency]').textContent = dependencyLabel(project, milestone) || 'No current dependency';
    dialog.querySelectorAll('[data-dependency-editor]').forEach(element => {element.hidden = readOnly;});
    dialog.querySelectorAll('input[name="milestone-dependency-kind"],#milestone-dependency-target').forEach(element => {element.disabled = readOnly;});
    dialog.querySelector('[data-remove]').hidden = readOnly;
    dialog.querySelector('[data-save]').hidden = readOnly;
    dialog.querySelector('[data-cancel]').textContent = readOnly ? 'Close' : 'Cancel';
    dialog.querySelector('[data-error]').hidden = true;
    const kind = milestone.dependencyType === 'track' ? 'track' : 'milestone';
    dialog.querySelector(`input[name="milestone-dependency-kind"][value="${kind}"]`).checked = true;
    populate(kind === 'track' ? milestone.dependencyTrackId : milestone.dependencyMilestoneId);
    dialog.showModal();
    lucide.createIcons({root:dialog});
  }
  async function snapshots(project) {
    const fields = ['id','track_id','record_status','sequence_no','schedule_mode','reference_milestone_id','reference_date_type','reference_offset_days','start_date','end_date','deadline','revised_end_date','completion_date','status','days','remarks','dependency_type','dependency_milestone_id','dependency_track_id'];
    const result = await fetchAllSupabaseRows('milestones','id',false,{columns:fields.join(','),equals:{initiative_id:Number(project.id)}});
    if(result.error)throw new Error(['42703','PGRST204'].includes(result.error.code)?'Run the dependency migration in Supabase, then refresh the portal.':result.error.message);
    const rows = (result.data||[]).filter(active),byId=new Map(rows.map(row=>[Number(row.id),row]));
    const local=(project.milestones||[]).filter(active);
    if(rows.length!==local.length)throw new Error('The Milestone list changed. Refresh and retry.');
    const same=(first,second)=>first==null||first===''?second==null||second==='':String(first)===String(second);
    for(const item of local){
      const row=byId.get(Number(item.id));
      if(!row||!Object.hasOwn(row,'dependency_type'))throw new Error('Run the dependency migration in Supabase, then refresh the portal.');
      const canonical={...item,startDate:row.start_date||'',endDate:row.end_date||'',deadline:row.deadline||'',days:row.days??deriveMilestoneDays(row.status,row.start_date,row.end_date),remarks:row.remarks||''};
      const schedule=calculateMilestoneSchedule(canonical);
      if(schedule.hasHistory)Object.assign(canonical,{startDate:schedule.startDate,endDate:schedule.endDate,deadline:schedule.deadline,days:schedule.days,remarks:schedule.remarks});
      if(!same(item.trackId,row.track_id)||Number(item.sequenceNo||0)!==Number(row.sequence_no||0)||!same(item.scheduleMode||'sequential',row.schedule_mode||'sequential')||!same(item.referenceMilestoneId,row.reference_milestone_id)||!same(item.referenceDateType||'start_date',row.reference_date_type||'start_date')||Number(item.referenceOffsetDays||0)!==Number(row.reference_offset_days||0)||!same(item.revisedEndDate,row.revised_end_date)||!same(item.completionDate,row.completion_date)||!same(item.dependencyType,row.dependency_type)||!same(item.dependencyMilestoneId,row.dependency_milestone_id)||!same(item.dependencyTrackId,row.dependency_track_id)||['startDate','endDate','deadline','days','remarks'].some(field=>!same(item[field],canonical[field])))throw new Error('Milestone dates or dependencies changed. Refresh and retry.');
    }
    return rows.map(row=>Object.fromEntries(fields.map(field=>[field,row[field]??null])));
  }
  function renderAffected(project, origin) {
    const old = document.querySelector(`#projects-container [data-overview-search-level="Initiative"][data-overview-search-id="${Number(project.id)}"]`);
    if (!old) return;
    const view = captureDeliveryViewState(origin, {layoutOnly:true});
    const host = document.createElement('div');
    renderInitiativeCards([project], host);
    const replacement = host.firstElementChild;
    if (replacement) {
      old.replaceWith(replacement);
      window.deliveryInitiativeReorder?.decorate(replacement.parentElement);
      decorateTrackReorderControls();
      decorateAll(replacement);
    }
    restoreDeliveryViewState(view,{restoreFocus:false,defer:false});
  }
  function applyRows(project, rows) {
    const byId = new Map((project.milestones || []).map(item => [Number(item.id), item]));
    const kpi = new Map(kpiMilestoneRows.map(row => [Number(row.id), row]));
    for (const row of rows) {
      const item = byId.get(Number(row.id));
      if (!item) continue;
      Object.assign(item,{startDate:row.start_date||'',endDate:row.end_date||'',deadline:row.deadline||'',status:row.status||'',persistedStatus:row.status||'',deviationDays:row.deviation_days??null,days:row.days??'',remarks:row.remarks||'',dependencyType:row.dependency_type||'',dependencyMilestoneId:Number(row.dependency_milestone_id)||null,dependencyTrackId:Number(row.dependency_track_id)||null});
      if (kpi.has(Number(row.id))) Object.assign(kpi.get(Number(row.id)),row);
    }
  }
  function prepareEchoes(changes, milestoneId, type, target) {
    for (const [id, echo] of pendingEchoes) if (echo.until < Date.now()) pendingEchoes.delete(id);
    const records = new Map(kpiMilestoneRows.map(row => [Number(row.id), row]));
    for (const change of changes) {
      const id=Number(change.id),baseline=records.get(id);
      if (!baseline) continue;
      const expected={...change};delete expected.id;
      if(id===Number(milestoneId))Object.assign(expected,{dependency_type:type,dependency_milestone_id:type==='milestone'?Number(target):null,dependency_track_id:type==='track'?Number(target):null});
      pendingEchoes.set(id,{baseline:{...baseline},expected,until:Date.now()+15000});
    }
  }
  function acceptRealtime(payload) {
    if(payload?.table!=='milestones'||payload.eventType!=='UPDATE')return false;
    const row=payload.new,id=Number(row?.id),echo=pendingEchoes.get(id);
    if(!echo)return false;
    if(echo.until<Date.now()){pendingEchoes.delete(id);return false;}
    const equal=(field,left,right)=>field==='days'&&left!=null&&right!=null?String(left)===String(right):JSON.stringify(left)===JSON.stringify(right);
    const matches=Object.keys(row).every(field=>auditFields.has(field)||equal(field,row[field],Object.hasOwn(echo.expected,field)?echo.expected[field]:echo.baseline[field]));
    if(matches){
      const kpi=kpiMilestoneRows.find(item=>Number(item.id)===id);
      if(kpi)for(const field of auditFields)if(Object.hasOwn(row,field))kpi[field]=row[field];
      pendingEchoes.delete(id);
    }
    return matches;
  }
  function verifyRows(rows, changes, milestoneId, type, target) {
    if(!Array.isArray(rows)||rows.length!==changes.size||new Set(rows.map(row=>Number(row.id))).size!==changes.size)return false;
    const equal=(left,right)=>left==null||right==null?left==null&&right==null:String(left)===String(right);
    for(const row of rows){
      const expected=changes.get(Number(row.id));
      if(!expected||Object.keys(expected).some(field=>field!=='id'&&!equal(row[field],expected[field])))return false;
      if(Number(row.id)===Number(milestoneId)&&(!equal(row.dependency_type,type)||!equal(row.dependency_milestone_id,type==='milestone'?target:null)||!equal(row.dependency_track_id,type==='track'?target:null)))return false;
    }
    return true;
  }
  async function propagate(project, options={}) {
    if (busy) return false;
    const dependencies=(project.milestones||[]).filter(item=>active(item)&&item.dependencyType);
    if (!dependencies.length) return false;
    busy=true;
    let echoIds=[];
    try {
      const expected=await snapshots(project);
      const plan=await planMilestoneDependencyChanges(project,(project.milestones||[]).filter(active).map(item=>Number(item.id)));
      if (!plan.changes.length) return false;
      const changes=new Map(plan.changes.map(item=>[Number(item.milestone.id),{id:Number(item.milestone.id),...item.payload}]));
      const owner=dependencies.find(item=>changes.has(Number(item.id)));
      if (!owner) throw new Error('Unable to identify the dependency that changed.');
      const target=owner.dependencyType==='milestone'?Number(owner.dependencyMilestoneId):Number(owner.dependencyTrackId);
      echoIds=[...changes.keys()];
      prepareEchoes([...changes.values()],owner.id,owner.dependencyType,target);
      const result=await supabaseClient.rpc('delivery_set_milestone_dependency',{p_initiative_id:Number(project.id),p_milestone_id:Number(owner.id),p_type:owner.dependencyType,p_target_id:target,p_expected:expected,p_changes:[...changes.values()]});
      if(result.error){changes.forEach((value,id)=>pendingEchoes.delete(id));throw new Error(result.error.message);}
      if(!verifyRows(result.data,changes,owner.id,owner.dependencyType,target))throw new Error('Supabase did not confirm every dependency update. Refresh before retrying.');
      applyRows(project,result.data);
      if(options.render!==false&&currentWorkspaceView==='overview')renderAffected(project,null);
      return true;
    } catch(error){echoIds.forEach(id=>pendingEchoes.delete(id));throw error;} finally {busy=false;}
  }
  async function save(remove) {
    if (busy || !context) return;
    const {project, milestone, origin} = context;
    if (context.readOnly || milestone.dependencyType) return showToast('An existing dependency is view-only and cannot be overridden.', 'info');
    const type = remove ? null : dialog.querySelector('input[name="milestone-dependency-kind"]:checked').value;
    const target = remove ? null : Number(dialog.querySelector('#milestone-dependency-target').value);
    const errorBox = dialog.querySelector('[data-error]');
    let echoIds=[];
    try {
      if (!remove) {
        const reason = validateMilestoneDependency(project, milestone.id, type, target);
        if (reason) throw new Error(reason);
      }
      busy = true;
      dialog.querySelectorAll('button,input,select').forEach(element => {element.disabled=true;});
      const expected = await snapshots(project);
      const staged = {...project,milestones:(project.milestones || []).map(item => Number(item.id) === Number(milestone.id) ? {...item,dependencyType:type || '',dependencyMilestoneId:type === 'milestone' ? target : null,dependencyTrackId:type === 'track' ? target : null} : {...item})};
      const sourceIds = remove ? [] : type === 'milestone' ? [target] : staged.milestones.filter(item => active(item) && Number(item.trackId) === target).map(item => Number(item.id));
      if (!remove && !hasMilestoneCompletion(milestone) && !sourceIds.some(id => {const source=staged.milestones.find(item=>Number(item.id)===id);return /^\d{4}-\d{2}-\d{2}$/.test(String(source?.revisedEndDate||source?.deadline||source?.endDate||''));})) throw new Error('The selected dependency has no valid end date.');
      const plan = remove ? {changes:[]} : await planMilestoneDependencyChanges(staged,sourceIds);
      const changes = new Map(plan.changes.map(item => [Number(item.milestone.id), {id:Number(item.milestone.id),...item.payload}]));
      if (!changes.has(Number(milestone.id))) changes.set(Number(milestone.id), {id:Number(milestone.id)});
      echoIds=[...changes.keys()];
      prepareEchoes([...changes.values()],milestone.id,type,target);
      const result = await supabaseClient.rpc('delivery_set_milestone_dependency',{p_initiative_id:Number(project.id),p_milestone_id:Number(milestone.id),p_type:type,p_target_id:target,p_expected:expected,p_changes:[...changes.values()]});
      if (result.error) {changes.forEach((value,id)=>pendingEchoes.delete(id));throw new Error(['PGRST202','42883'].includes(result.error.code) ? 'Run the dependency migration in Supabase, then refresh the portal.' : result.error.message);}
      if (!verifyRows(result.data,changes,milestone.id,type,target)) throw new Error('Supabase did not confirm every dependency update. Refresh before retrying.');
      applyRows(project,result.data);
      await synchronizeInitiativeRevisedEndDates([project]);
      busy = false;
      dialog.querySelectorAll('button,input,select').forEach(element => {element.disabled=false;});
      close();
      renderAffected(project,origin);
      showToast(remove ? 'Dependency removed.' : 'Dependency saved.', 'success');
    } catch (error) {
      echoIds.forEach(id=>pendingEchoes.delete(id));
      errorBox.textContent = error?.message || 'Unable to save dependency.';
      errorBox.hidden = false;
    } finally {
      busy = false;
      dialog.querySelectorAll('button,input,select').forEach(element => {element.disabled=false;});
    }
  }
  function decorate(row, project, milestone) {
    const cell = row.querySelector('.milestone-action-text');
    if (!cell) return;
    cell.querySelector('[data-milestone-dependency-indicator]')?.remove();
    const label = dependencyLabel(project,milestone);
    if (!label) return;
    const button = document.createElement('button');
    button.type = 'button';button.dataset.milestoneDependencyIndicator = '';
    button.className = 'mt-1 block text-left text-xs text-slate-500 underline';
    button.textContent = label;button.title = 'View dependency';
    button.onclick = event => {event.stopPropagation();open(project.id,milestone.id,button);};
    cell.appendChild(button);
  }
  function decorateAll(root=document.getElementById('projects-container')) {
    root?.querySelectorAll('tr[data-inline-milestone-id][data-inline-milestone-project-id]').forEach(row => {
      const project=projects.find(item=>Number(item.id)===Number(row.dataset.inlineMilestoneProjectId));
      const milestone=project?.milestones?.find(item=>Number(item.id)===Number(row.dataset.inlineMilestoneId));
      if(project&&milestone)decorate(row,project,milestone);
    });
  }
  window.deliveryMilestoneDependency = {open,decorate,decorateAll,propagate,renderAffected,acceptRealtime,get busy(){return busy;}};
  decorateAll();
})();