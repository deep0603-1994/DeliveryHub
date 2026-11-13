(() => {
  const key = value => String(value ?? '');
  const html = value => escapeHtml(key(value));
  let session = null;
  let opening = false;
  function locate(projectId, milestoneId) {
    const project = projects.find(row => key(row.id) === key(projectId));
    return { project, milestone: project?.milestones?.find(row => key(row.id) === key(milestoneId)) };
  }
  function permitted(project) {
    return Boolean(currentWorkspaceView === 'overview' && currentUser && project && canEditInitiative(project) && milestoneTypeIntegrity.active(project));
  }
  function modal() {
    let node = document.getElementById('fetch-planner-modal');
    if (node) return node;
    node = document.createElement('div');
    node.id = 'fetch-planner-modal';
    node.className = 'fixed inset-0 z-[120] hidden items-center justify-center bg-slate-900/60 p-3';
    node.setAttribute('role', 'dialog');
    node.setAttribute('aria-modal', 'true');
    node.setAttribute('aria-labelledby', 'fetch-planner-title');
    node.innerHTML = `<section class="flex w-full max-w-5xl flex-col overflow-hidden rounded-lg bg-white shadow-xl" style="max-height:calc(100dvh - 1.5rem)">
      <header class="flex items-center justify-between gap-3 border-b bg-indigo-50 px-4 py-3"><h2 id="fetch-planner-title" class="text-sm font-bold text-slate-800">Fetch Planner Type</h2><button type="button" data-fetch-close title="Close" aria-label="Close Fetch Planner Type" class="p-2 text-slate-500"><i data-lucide="x" class="h-5 w-5"></i></button></header>
      <div class="space-y-3 overflow-auto p-4 text-xs" style="min-height:0;overscroll-behavior:contain">
        <div id="fetch-planner-context" class="font-semibold text-slate-700"></div>
        <div class="flex flex-wrap items-end gap-3"><label class="font-semibold">Scope<select id="fetch-planner-scope" class="mt-1 block min-h-11 rounded border border-slate-300 bg-white px-3 py-2"><option value="milestone">This Milestone</option><option value="track">All Milestones in this Track</option><option value="initiative">All Milestones in this Initiative</option></select></label></div>
        <div id="fetch-planner-summary" class="font-semibold" role="status" aria-live="polite"></div>
        <div id="fetch-planner-message" role="alert" aria-live="assertive" class="whitespace-pre-wrap break-words text-rose-700"></div>
        <div class="overflow-auto rounded border border-slate-200"><table class="w-full text-left text-xs"><thead class="bg-slate-100"><tr>${['Milestone', 'Current Planner Type', 'Matched Planner Type', 'Match Source', 'Status'].map(label => `<th class="p-3">${label}</th>`).join('')}</tr></thead><tbody id="fetch-planner-rows"></tbody></table></div>
      </div>
      <footer class="flex flex-wrap justify-end gap-2 border-t px-4 py-3"><button type="button" data-fetch-close class="min-h-11 rounded border px-4 py-2 text-xs">Cancel</button><button id="fetch-planner-apply" type="button" class="min-h-11 rounded bg-indigo-600 px-4 py-2 text-xs font-semibold text-white disabled:opacity-50">Apply Mapping</button></footer>
    </section>`;
    document.body.appendChild(node);
    node.addEventListener('click', event => {
      event.stopPropagation();
      if(event.target.closest('button'))event.preventDefault();
      if (event.target === node || event.target.closest('[data-fetch-close]')) close();
      if (event.target.closest('#fetch-planner-apply')) apply();
    });
    node.addEventListener('change', event => {
      if (!session || session.saving) return;
      if(event.target.matches('#fetch-planner-scope'))preview();
    });
    window.lucide?.createIcons({ root: node });
    return node;
  }
  function close() {
    if (session?.saving || opening) return;
    const origin = session?.origin;
    session = null;
    document.getElementById('fetch-planner-modal')?.classList.add('hidden');
    origin?.isConnected && origin.focus({ preventScroll: true });
  }
  function candidates() {
    const scope = document.getElementById('fetch-planner-scope').value;
    return session.milestones.filter(row => milestoneTypeIntegrity.active(row) && Number(row.id) > 0 && (scope === 'initiative' || (scope === 'track'
      ? key(row.trackId) === key(session.source.trackId) : key(row.id) === key(session.source.id))));
  }
  function summary(updated=0){
    const count=status=>session.rows.filter(row=>row.status===status).length;
    return `Checked: ${session.rows.length} | Updated: ${updated} | Already Mapped: ${count('Already Mapped')} | No Match: ${count('No Match')} | Duplicate Prevented: ${count('Duplicate Planner Type')}`;
  }
  function preview() {
    if (!session) return;
    session.rows = milestoneTypeIntegrity.plannerMappingPreview(candidates(), session.assignments,
      session.rules, session.types);
    const count = status => session.rows.filter(row => row.status === status).length;
    document.getElementById('fetch-planner-summary').textContent = summary();
    const label = id => milestoneTypeEditing.label(id);
    document.getElementById('fetch-planner-rows').innerHTML = session.rows.map(row => {
      const matched = html(row.typeId ? label(row.typeId) : '-');
      return `<tr class="border-t"><td class="p-3">${html(row.milestone.actions || row.milestone.title)}</td><td class="p-3">${html(row.current ? label(row.current) : 'Not Mapped')}</td><td class="p-3">${matched}</td><td class="p-3">${html(row.source || '-')}</td><td class="p-3">${html(row.status)}${row.reason ? `<div class="mt-1 text-rose-700">${html(row.reason)}</div>` : ''}</td></tr>`;
    }).join('');
    document.getElementById('fetch-planner-apply').disabled = !count('Ready');
  }
  async function open(projectId, milestoneId, origin = document.activeElement) {
    if (opening || session?.saving) return;
    const { project, milestone } = locate(projectId, milestoneId);
    if (!permitted(project) || !milestone || !milestoneTypeIntegrity.active(milestone)) return showToast('Permission denied.', 'error');
    opening = true;
    try {
      const [loaded, typesLoaded, ruleResult, assignments] = await Promise.all([
        loadPlanners(), milestonePlannerTypes.load(),
        fetchAllSupabaseRows('project_planner_milestones', 'display_order', false, { equals: { is_active: true } }),
        fetchAllSupabaseRows('milestones', 'id', false, { equals: { initiative_id: project.id } })
      ]);
      if (!loaded || !typesLoaded) throw new Error(milestonePlannerTypes.state.error || 'Unable to load Project Planner.');
      if (ruleResult.error || assignments.error) throw ruleResult.error || assignments.error;
      if (!permitted(project)) throw new Error('Permission denied.');
      const templates = new Map(plannerTemplates.filter(row => row.is_active !== false).map(row => [key(row.id), row]));
      const types = milestonePlannerTypes.active('');
      const typeById = new Map(types.map(type => [key(type.id), type]));
      const rules = (ruleResult.data || []).filter(row => {
        const template = templates.get(key(row.template_id)), type = typeById.get(key(row.planner_type_id));
        return template && type && (!template.planner_type_group_id || key(template.planner_type_group_id) === key(type.group_id));
      });
      const milestones = (assignments.data || []).map(row => ({ ...row, trackId: row.track_id, actions: row.actions || row.milestone_title, plannerTypeId: row.planner_type_id, recordStatus: row.record_status }));
      const source = milestones.find(row => key(row.id) === key(milestoneId));
      if (!source) throw new Error('Milestone is no longer available. Refresh before mapping.');
      session = { projectId: project.id, source, milestones, rules, types, assignments: assignments.data || [], rows: [], saving: false, origin, actor: key(currentUser?.id || currentUser?.email), viewState:typeof captureDeliveryViewState==='function'?captureDeliveryViewState(origin):null };
      const node = modal();
      document.getElementById('fetch-planner-scope').value = 'milestone';
      document.getElementById('fetch-planner-context').textContent = project.name;
      document.getElementById('fetch-planner-message').textContent = '';
      preview();
      node.classList.remove('hidden');
      node.classList.add('flex');
      document.getElementById('fetch-planner-scope').focus();
    } catch (error) { showToast(error.message || 'Unable to fetch Planner Types.', 'error'); }
    finally { opening = false; }
  }
  async function apply() {
    if (!session || session.saving) return;
    const project = projects.find(row => key(row.id) === key(session.projectId));
    if (!permitted(project) || session.actor !== key(currentUser?.id || currentUser?.email)) return showToast('Permission denied.', 'error');
    preview();
    const changes = session.rows.filter(row => row.status === 'Ready' && !row.current).map(row => ({ id: row.milestone.id, planner_type_id: row.typeId, expected_planner_type_id: null }));
    if (!changes.length) return;
    session.saving = true;
    const node = modal();
    document.getElementById('fetch-planner-message').textContent = '';
    document.getElementById('fetch-planner-apply').textContent = 'Applying Mapping...';
    node.querySelectorAll('button,select,input').forEach(control => { control.disabled = true; });
    let saved = false;
    try {
      const patches = changes.map(change => ({ id: change.id, initiative_id: session.projectId, planner_type_id: change.planner_type_id }));
      await milestoneTypeIntegrity.validate(patches, session.assignments);
      const result = await supabaseClient.rpc('delivery_fetch_planner_types', { p_initiative_id: session.projectId, p_changes: changes });
      if (result.error) throw new Error(['PGRST202', '42883'].includes(result.error.code) ? 'Run supabase-migration-fetch-planner-types.sql in Supabase before applying mappings.' : result.error.message);
      const confirmed = new Map((Array.isArray(result.data)?result.data:[]).map(row => [key(row.id), row.planner_type_id]));
      if (!Array.isArray(result.data) || result.data.length !== changes.length || changes.some(change => key(confirmed.get(key(change.id)))!==key(change.planner_type_id))) throw new Error('Supabase did not confirm every mapping. Refresh before retrying.');
      saved = true;
      const currentProject=projects.find(row=>key(row.id)===key(session.projectId))||project;
      for (const milestone of currentProject.milestones || []) if (confirmed.has(key(milestone.id))) milestone.plannerTypeId = confirmed.get(key(milestone.id));
      for (const row of kpiMilestoneRows) if (confirmed.has(key(row.id))) row.planner_type_id = confirmed.get(key(row.id));
      const branch=document.querySelector(`#projects-container [data-overview-search-level="Initiative"][data-overview-search-id="${Number(project.id)}"]`)||document.getElementById('projects-container');
      if(branch)milestoneTypeEditing.decorate(branch);
      showToast(summary(changes.length), 'success');
    } catch (error) {
      const message=document.getElementById('fetch-planner-message');
      message.textContent=error.message;
      message.scrollIntoView({block:'nearest'});
    }
    finally {
      session.saving = false;
      node.querySelectorAll('button,select,input').forEach(control => { control.disabled = false; });
      document.getElementById('fetch-planner-apply').textContent = 'Apply Mapping';
      if(session.viewState&&currentWorkspaceView==='overview')restoreDeliveryViewState(session.viewState,{restoreFocus:false});
      if (saved) close();
    }
  }
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape' && session) { event.preventDefault(); close(); }
    if (event.key === 'Tab' && session && !session.saving) {
      const controls = [...modal().querySelectorAll('button:not(:disabled),select:not(:disabled),input:not(:disabled)')];
      const first = controls[0], last = controls.at(-1);
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
      if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
    }
  });
  window.deliveryPlannerFetch = { open, close, preview, apply, get busy(){return Boolean(session?.saving);} };
})();