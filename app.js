const APP_CONFIG = {
  clientId: 'e59797ac-6600-4848-b254-89781ae9dd27',
  tenantId: '554c414a-280b-4c81-97d2-f6c23aaefa88',
  siteHost: 'stratusdnd.sharepoint.com',
  sitePath: '/sites/SD-026UMassInfusionPumpHeuristicEval',
  studyId: 'SD-026',
  graphScopes: ['User.Read', 'Sites.ReadWrite.All']
};

const heuristics = [
  'Consistency: Consistency and standards',
  'Visibility: Visibility of system state',
  'Match: Match between system and world',
  'Minimalist: Minimize information to what is necessary',
  'Memory: Minimize memory load',
  'Feedback: Informative feedback',
  'Flexibility: Flexibility and efficiency',
  'Message: Good error messages',
  'Error: Prevent errors',
  'Closure: Clear closure',
  'Undo: Reversible actions',
  "Language: Use users’ language",
  'Control: Users in control',
  'Document: Help and documentation'
];

const defaultState = { session: null, issues: [], completedScenarios: [], auth: { account: null } };
const state = JSON.parse(localStorage.getItem('shc_state') || JSON.stringify(defaultState));
state.issues ||= [];
state.completedScenarios ||= [];
state.auth ||= { account: null };
let editingId = null;
let photos = [];
let graphContext = null;
let msalApp = null;

const $ = id => document.getElementById(id);
const screens = ['sessionScreen', 'homeScreen', 'issueScreen', 'exportScreen'];
const GRAPH = 'https://graph.microsoft.com/v1.0';

function saveLocal() { localStorage.setItem('shc_state', JSON.stringify(state)); }
function esc(s = '') { return String(s).replace(/[&<>"']/g, m => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[m])); }
function timestampFileName(prefix, ext = 'jpg') { return `${prefix}-${new Date().toISOString().replace(/[:.]/g, '-')}.${ext}`; }
function currentContextKey() { return `${state.session?.device || ''}||${state.session?.scenario || ''}`; }
function issueBelongsToCurrent(i) { return !!state.session && i.device === state.session.device && i.scenario === state.session.scenario; }
function numericSeverity(v = '') { return Number(String(v).trim().charAt(0)) || 0; }
function severityLabel(v = '') { return String(v || '—'); }
function pendingIssues() { return state.issues.filter(i => !i.remoteIssueId); }

function show(id) { screens.forEach(s => $(s).classList.toggle('active', s === id)); render(); }
function setText(id, value) { const el = $(id); if (el) el.textContent = value; }

function renderHeuristics() {
  $('heuristicChoices').innerHTML = heuristics.map(h => `<label><input type="checkbox" value="${esc(h)}" /> ${esc(h)}</label>`).join('');
}

function renderPhotos() {
  const c = $('photoThumbs');
  if (!photos.length) { c.innerHTML = '<p class="empty-small">No photos added.</p>'; return; }
  c.innerHTML = photos.map((p, idx) => `<div class="photo-card"><img src="${p.data}" alt="Evidence photo ${idx + 1}" /><div class="row between"><span>Photo ${idx + 1}</span><button type="button" class="remove-photo" data-idx="${idx}">Remove</button></div></div>`).join('');
  document.querySelectorAll('.remove-photo').forEach(btn => btn.onclick = () => { photos.splice(Number(btn.dataset.idx), 1); renderPhotos(); });
}

function render() {
  if (state.session) {
    setText('sessionSummary', `${state.session.device} · ${state.session.scenario}`);
    setText('currentDevice', state.session.device);
    setText('currentScenario', state.session.scenario);
  } else {
    setText('sessionSummary', 'No pump/scenario selected');
    setText('currentDevice', '—');
    setText('currentScenario', '—');
  }
  const account = state.auth?.account;
  setText('accountStatus', account ? `Signed in as ${account.name || account.username}` : 'Sign in to sync to SharePoint. You may still capture issues locally while offline.');
  $('signInBtn').classList.toggle('hidden', !!account);
  $('signOutBtn').classList.toggle('hidden', !account);

  const currentIssues = state.issues.filter(issueBelongsToCurrent);
  setText('issueCount', currentIssues.length);
  setText('highSeverityCount', currentIssues.filter(i => numericSeverity(i.severity) >= 3).length);
  setText('pendingCount', currentIssues.filter(i => !i.remoteIssueId).length);
  setText('syncStatus', account ? `${pendingIssues().length} issue(s) pending sync.` : 'Sign in to sync this evaluation.');

  if (!currentIssues.length) {
    $('issueList').className = 'issue-list empty';
    $('issueList').textContent = 'No issues recorded for this pump/scenario yet.';
  } else {
    $('issueList').className = 'issue-list';
    $('issueList').innerHTML = currentIssues.map(i => `<div class="issue" data-id="${esc(i.id)}"><div class="issue-title">${esc(i.title || 'Untitled issue')}</div><div class="issue-meta"><span class="badge">${esc(severityLabel(i.severity))}</span><span class="badge">${esc(i.locationOccurrence || '')}</span><span class="badge">${(i.photos || []).length} photo(s)</span>${i.remoteIssueId ? '<span class="badge">Synced</span>' : '<span class="badge">Pending</span>'}</div><div class="issue-meta">${esc((i.heuristics || []).join('; '))}</div></div>`).join('');
    document.querySelectorAll('.issue').forEach(el => el.onclick = () => editIssue(el.dataset.id));
  }
}

function setSessionFormFromCurrent() {
  if (!state.session) return;
  $('device').value = state.session.device || '';
  $('scenario').value = state.session.scenario || '';
  $('softwareVersion').value = state.session.softwareVersion || '';
  $('drugLibrary').value = state.session.drugLibrary || '';
}

function startSession() {
  if (!$('device').value || !$('scenario').value) { alert('Select a pump system and scenario.'); return; }
  state.session = {
    device: $('device').value,
    scenario: $('scenario').value,
    softwareVersion: $('softwareVersion').value.trim(),
    drugLibrary: $('drugLibrary').value.trim(),
    startedAt: state.session?.startedAt || new Date().toISOString(),
    remoteSessionId: state.session?.device === $('device').value && state.session?.scenario === $('scenario').value ? state.session.remoteSessionId || null : null
  };
  saveLocal();
  show('homeScreen');
}

function changeContext() { setSessionFormFromCurrent(); $('sessionScreenTitle').textContent = 'Change Pump / Scenario'; show('sessionScreen'); }
function finishScenario() {
  if (!state.session) return;
  const ok = confirm(`Finish Scenario?\n\n${state.session.device}\n${state.session.scenario}\n\nSelect OK to finish this scenario and change the evaluation context.`);
  if (!ok) return;
  state.completedScenarios.push({ ...state.session, finishedAt: new Date().toISOString(), issueCount: state.issues.filter(issueBelongsToCurrent).length });
  saveLocal();
  changeContext();
}

function clearIssueForm() {
  editingId = null; photos = [];
  $('issueFormTitle').textContent = 'New Issue';
  $('issueTitle').value = '';
  $('issueDescription').value = '';
  $('taskPhase').value = state.session?.scenario || 'Normal saline infusion at 125 mL/hr';
  $('locationOccurrence').value = 'Programming';
  $('photo').value = '';
  document.querySelectorAll('#heuristicChoices input').forEach(c => c.checked = false);
  document.querySelector('input[name="severity"][value="2: Minor"]').checked = true;
  renderPhotos();
}

function newIssue() { if (!state.session) { show('sessionScreen'); return; } clearIssueForm(); show('issueScreen'); }
function editIssue(id) {
  const issue = state.issues.find(i => i.id === id); if (!issue) return;
  editingId = id;
  photos = (issue.photos || []).map(p => ({ ...p }));
  $('issueFormTitle').textContent = 'Edit Issue';
  $('issueTitle').value = issue.title || '';
  $('issueDescription').value = issue.description || '';
  $('taskPhase').value = issue.taskPhase || state.session?.scenario || 'Normal saline infusion at 125 mL/hr';
  $('locationOccurrence').value = issue.locationOccurrence || 'Programming';
  document.querySelectorAll('#heuristicChoices input').forEach(c => c.checked = (issue.heuristics || []).includes(c.value));
  const sev = issue.severity || '2: Minor';
  const radio = document.querySelector(`input[name="severity"][value="${CSS.escape(sev)}"]`);
  if (radio) radio.checked = true;
  renderPhotos();
  show('issueScreen');
}

function saveIssue() {
  if (!state.session) { alert('Select a pump system and scenario before recording issues.'); return; }
  const title = $('issueTitle').value.trim();
  const description = $('issueDescription').value.trim();
  if (!title || !description) { alert('Enter an issue title and usability issue description.'); return; }
  const id = editingId || `ISS-${Date.now()}`;
  const prior = state.issues.find(i => i.id === id) || {};
  const item = {
    id,
    timestamp: prior.timestamp || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    device: prior.device || state.session.device,
    scenario: prior.scenario || state.session.scenario,
    softwareVersion: prior.softwareVersion || state.session.softwareVersion || '',
    drugLibrary: prior.drugLibrary || state.session.drugLibrary || '',
    title,
    description,
    taskPhase: $('taskPhase').value,
    locationOccurrence: $('locationOccurrence').value,
    heuristics: [...document.querySelectorAll('#heuristicChoices input:checked')].map(c => c.value),
    severity: document.querySelector('input[name="severity"]:checked').value,
    photos: photos.map((p, idx) => ({ ...p, fileName: p.fileName || `${id}_photo-${idx + 1}.${p.type?.includes('png') ? 'png' : 'jpg'}` })),
    remoteIssueId: prior.remoteIssueId || null,
    remotePhotoIds: prior.remotePhotoIds || []
  };
  if (editingId) state.issues = state.issues.map(i => i.id === id ? item : i); else state.issues.push(item);
  saveLocal();
  show('homeScreen');
}

async function compressImage(file) {
  const bitmap = await createImageBitmap(file);
  const max = 1600;
  const ratio = Math.min(1, max / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(bitmap.width * ratio));
  canvas.height = Math.max(1, Math.round(bitmap.height * ratio));
  canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.8));
  if (!blob) throw new Error('Photo compression failed.');
  const data = await blobToDataUrl(blob);
  return { data, type: 'image/jpeg', originalName: file.name || 'photo.jpg' };
}
function blobToDataUrl(blob) { return new Promise((resolve, reject) => { const r = new FileReader(); r.onload = () => resolve(r.result); r.onerror = reject; r.readAsDataURL(blob); }); }
function dataUriToBlob(dataUri) { const [head, b64] = dataUri.split(','); const type = (head.match(/data:(.*?);/) || [,'image/jpeg'])[1]; const binary = atob(b64); const bytes = new Uint8Array(binary.length); for (let i=0; i<binary.length; i++) bytes[i] = binary.charCodeAt(i); return new Blob([bytes], { type }); }

async function setupAuth() {
  if (!window.msal) { alert('Microsoft sign-in library did not load. Check your network connection.'); return; }
  msalApp = new msal.PublicClientApplication({ auth: { clientId: APP_CONFIG.clientId, authority: `https://login.microsoftonline.com/${APP_CONFIG.tenantId}`, redirectUri: window.location.origin + window.location.pathname }, cache: { cacheLocation: 'localStorage' } });
  await msalApp.initialize();
  const accounts = msalApp.getAllAccounts();
  if (accounts.length) setAccount(accounts[0]);
}
function setAccount(account) { state.auth.account = account ? { username: account.username, name: account.name, homeAccountId: account.homeAccountId } : null; saveLocal(); render(); }
async function signIn() {
  try { const result = await msalApp.loginPopup({ scopes: APP_CONFIG.graphScopes, prompt: 'select_account' }); setAccount(result.account); await getGraphContext(); }
  catch (err) { console.error(err); alert(`Sign-in was not completed. ${err.message || ''}`); }
}
async function signOut() {
  const accounts = msalApp.getAllAccounts();
  setAccount(null); graphContext = null;
  if (accounts.length) await msalApp.logoutPopup({ account: accounts[0], postLogoutRedirectUri: window.location.href });
}
async function token() {
  const account = msalApp.getAllAccounts()[0];
  if (!account) throw new Error('Sign in before syncing.');
  try { const r = await msalApp.acquireTokenSilent({ account, scopes: APP_CONFIG.graphScopes }); return r.accessToken; }
  catch { const r = await msalApp.acquireTokenPopup({ account, scopes: APP_CONFIG.graphScopes }); return r.accessToken; }
}
async function graph(path, options = {}) {
  const accessToken = await token();
  const res = await fetch(path.startsWith('http') ? path : `${GRAPH}${path}`, { ...options, headers: { Authorization: `Bearer ${accessToken}`, ...(options.headers || {}) } });
  if (!res.ok) { const body = await res.text(); throw new Error(`${res.status} ${res.statusText}: ${body}`); }
  if (res.status === 204) return null;
  const ct = res.headers.get('content-type') || '';
  return ct.includes('application/json') ? res.json() : res;
}
function listByName(lists, name) { const list = lists.find(x => x.displayName === name); if (!list) throw new Error(`SharePoint list not found: ${name}`); return list; }
function columnMap(cols) { const out = {}; cols.forEach(c => { if (c.displayName) out[c.displayName] = c.name; }); return out; }

async function getGraphContext() {
  if (graphContext) return graphContext;
  const site = await graph(`/sites/${APP_CONFIG.siteHost}:${APP_CONFIG.sitePath}`);
  const listsResponse = await graph(`/sites/${site.id}/lists?$select=id,displayName,list`);
  const lists = listsResponse.value || [];
  const names = ['Stratus Heuristic Studies', 'Stratus Heuristic Evaluators', 'Stratus Heuristic Sessions', 'Stratus Heuristic Issues', 'Stratus Heuristic Photos'];
  const resolved = {};
  for (const n of names) {
    const list = listByName(lists, n);
    const colResponse = await graph(`/sites/${site.id}/lists/${list.id}/columns?$select=id,name,displayName`);
    resolved[n] = { ...list, columns: columnMap(colResponse.value || []) };
  }
  const drive = await graph(`/sites/${site.id}/lists/${resolved['Stratus Heuristic Photos'].id}/drive`);
  graphContext = { site, lists: resolved, photoDriveId: drive.id };
  return graphContext;
}

async function getAllItems(siteId, listId) {
  let url = `/sites/${siteId}/lists/${listId}/items?expand=fields`;
  const values = [];
  while (url) { const page = await graph(url); values.push(...(page.value || [])); url = page['@odata.nextLink']?.replace(GRAPH, '') || null; }
  return values;
}
function fieldValue(item, fieldName) { return item.fields?.[fieldName]; }
async function findLookupItem(ctx, listName, possibleLabels, desired) {
  const list = ctx.lists[listName];
  const items = await getAllItems(ctx.site.id, list.id);
  const internalNames = possibleLabels.map(l => list.columns[l]).filter(Boolean);
  const norm = x => String(x || '').trim().toLowerCase();
  return items.find(item => internalNames.some(name => norm(fieldValue(item, name)) === norm(desired))) || null;
}
function lookupPayload(fieldInternalName, itemId) { return { [`${fieldInternalName}LookupId`]: Number(itemId) }; }

async function resolveEvaluator(ctx) {
  const profile = await graph('/me?$select=displayName,mail,userPrincipalName');
  const evaluator = await findLookupItem(ctx, 'Stratus Heuristic Evaluators', ['Evaluator Name', 'Title'], profile.displayName);
  if (!evaluator) throw new Error(`No matching record for ${profile.displayName} in Stratus Heuristic Evaluators. Add that evaluator to the list, then sync again.`);
  return { profile, evaluator };
}

async function ensureRemoteSession(ctx) {
  if (!state.session) throw new Error('Select a pump system and scenario first.');
  if (state.session.remoteSessionId) return state.session.remoteSessionId;
  const study = await findLookupItem(ctx, 'Stratus Heuristic Studies', ['Study ID', 'Title'], APP_CONFIG.studyId);
  if (!study) throw new Error(`Study ${APP_CONFIG.studyId} was not found in Stratus Heuristic Studies.`);
  const { evaluator } = await resolveEvaluator(ctx);
  const sessionList = ctx.lists['Stratus Heuristic Sessions'];
  const f = sessionList.columns;
  const fields = { Title: `${state.auth.account?.name || 'Evaluator'} – ${state.session.device} – ${state.session.scenario}` };
  if (f['Study ID']) Object.assign(fields, lookupPayload(f['Study ID'], study.id));
  if (f['Study Name']) Object.assign(fields, lookupPayload(f['Study Name'], study.id));
  if (f['Evaluator']) Object.assign(fields, lookupPayload(f['Evaluator'], evaluator.id));
  if (f['Device']) fields[f['Device']] = state.session.device;
  if (f['Scenario']) fields[f['Scenario']] = state.session.scenario;
  if (f['Software Version'] && state.session.softwareVersion) fields[f['Software Version']] = state.session.softwareVersion;
  const drugField = f['Drug Library/Profile'] || f['Drug Library / Profile'] || Object.entries(f).find(([label]) => label.startsWith('Drug Library'))?.[1];
  if (drugField && state.session.drugLibrary) fields[drugField] = state.session.drugLibrary;
  const created = await graph(`/sites/${ctx.site.id}/lists/${sessionList.id}/items`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
  state.session.remoteSessionId = created.id;
  saveLocal();
  return created.id;
}

async function createRemoteIssue(ctx, issue, sessionId) {
  const list = ctx.lists['Stratus Heuristic Issues'];
  const f = list.columns;
  const fields = { Title: issue.title };
  if (f['Session']) Object.assign(fields, lookupPayload(f['Session'], sessionId));
  if (f['Task Phase']) fields[f['Task Phase']] = issue.taskPhase;
  if (f['Location of Occurrence']) fields[f['Location of Occurrence']] = issue.locationOccurrence;
  if (f['Heuristic Violation Severity']) fields[f['Heuristic Violation Severity']] = issue.severity;
  if (f['Heuristics Violated']) {
    // Microsoft Graph requires an explicit OData type for a multi-select Choice field.
    fields[`${f['Heuristics Violated']}@odata.type`] = 'Collection(Edm.String)';
    fields[f['Heuristics Violated']] = issue.heuristics || [];
  }
  if (f['Usability Issue Description']) fields[f['Usability Issue Description']] = issue.description;
  return graph(`/sites/${ctx.site.id}/lists/${list.id}/items`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
}

async function uploadPhoto(ctx, issue, remoteIssueId, photo, index) {
  const safeTitle = (issue.title || 'issue').replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').slice(0, 50) || 'issue';
  const ext = photo.type?.includes('png') ? 'png' : 'jpg';
  const name = photo.fileName || timestampFileName(`${safeTitle}-${index + 1}`, ext);
  const blob = dataUriToBlob(photo.data);
  const encodedPath = encodeURIComponent(name).replace(/%2F/g, '/');
  const driveItem = await graph(`/drives/${ctx.photoDriveId}/root:/${encodedPath}:/content`, { method: 'PUT', headers: { 'Content-Type': blob.type }, body: blob });
  const photoList = ctx.lists['Stratus Heuristic Photos'];
  const f = photoList.columns;
  const fields = {};
  if (f['Issue']) Object.assign(fields, lookupPayload(f['Issue'], remoteIssueId));
  if (f['Photo Description']) fields[f['Photo Description']] = `Evidence photo ${index + 1} for: ${issue.title}`;
  if (Object.keys(fields).length) await graph(`/sites/${ctx.site.id}/lists/${photoList.id}/items/${driveItem.listItem.id}/fields`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(fields) });
  return driveItem.listItem.id;
}

async function syncAll() {
  if (!state.auth.account) { await signIn(); if (!state.auth.account) return; }
  const unsynced = pendingIssues();
  if (!unsynced.length) { alert('Everything is already synced to SharePoint.'); return; }
  $('syncBtn').disabled = true; $('syncBtn').textContent = 'Syncing…';
  let syncStage = 'connecting to SharePoint';
  try {
    const ctx = await getGraphContext();
    syncStage = 'creating or locating the evaluation session';
    const sessionId = await ensureRemoteSession(ctx);
    for (const issue of unsynced) {
      syncStage = `creating issue: ${issue.title || 'Untitled issue'}`;
      const remote = issue.remoteIssueId ? { id: issue.remoteIssueId } : await createRemoteIssue(ctx, issue, sessionId);
      issue.remoteIssueId = remote.id;
      issue.remotePhotoIds ||= [];
      for (let i = issue.remotePhotoIds.length; i < (issue.photos || []).length; i++) {
        syncStage = `uploading photo ${i + 1} for: ${issue.title || 'Untitled issue'}`;
        const remotePhotoId = await uploadPhoto(ctx, issue, remote.id, issue.photos[i], i);
        issue.remotePhotoIds.push(remotePhotoId);
      }
      saveLocal();
    }
    alert(`${unsynced.length} issue(s) and associated photos were synced to SharePoint.`);
  } catch (err) {
    console.error(err);
    alert(`Sync stopped during ${syncStage}. Previously synced records were preserved.\n\n${err.message || err}`);
  } finally {
    $('syncBtn').disabled = false; $('syncBtn').textContent = 'Sync'; render();
  }
}

function rows() {
  return state.issues.map(i => ({ issue_id: i.id, timestamp: i.timestamp, device: i.device, scenario: i.scenario, software_version: i.softwareVersion || '', drug_library_profile: i.drugLibrary || '', task_phase: i.taskPhase, location_of_occurrence: i.locationOccurrence, issue_title: i.title, usability_issue_description: i.description, heuristics: (i.heuristics || []).join('; '), severity: i.severity, photo_count: (i.photos || []).length, sync_status: i.remoteIssueId ? 'Synced' : 'Pending' }));
}
function toCsv(data) { const cols = Object.keys(data[0] || { issue_id:'', timestamp:'', device:'', scenario:'', software_version:'', drug_library_profile:'', task_phase:'', location_of_occurrence:'', issue_title:'', usability_issue_description:'', heuristics:'', severity:'', photo_count:'', sync_status:'' }); return [cols.join(','), ...data.map(r => cols.map(c => `"${String(r[c] ?? '').replace(/"/g, '""')}"`).join(','))].join('\n'); }
function download(name, content, type) { const blob = content instanceof Blob ? content : new Blob([content], { type }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
function openExport() { $('exportPreview').value = toCsv(rows()); show('exportScreen'); }

const crcTable = (() => { const table = new Uint32Array(256); for (let n=0;n<256;n++) { let c=n; for(let k=0;k<8;k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); table[n] = c >>> 0; } return table; })();
function crc32(bytes) { let c = 0xffffffff; for (const b of bytes) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function strBytes(s) { return new TextEncoder().encode(s); }
function dosDateTime(date = new Date()) { return { time: (date.getHours()<<11) | (date.getMinutes()<<5) | Math.floor(date.getSeconds()/2), date: ((date.getFullYear()-1980)<<9) | ((date.getMonth()+1)<<5) | date.getDate() }; }
function u16(n) { return [n & 255, (n>>>8)&255]; }
function u32(n) { return [n & 255, (n>>>8)&255, (n>>>16)&255, (n>>>24)&255]; }
function concatArrays(arrays) { const len = arrays.reduce((a,b) => a+b.length, 0); const out = new Uint8Array(len); let off = 0; arrays.forEach(a => { out.set(a,off); off += a.length; }); return out; }
function makeZip(files) { let offset=0; const locals=[]; const centrals=[]; const dt=dosDateTime(); for (const f of files) { const name=strBytes(f.name); const data=f.data instanceof Uint8Array ? f.data : strBytes(f.data); const crc=crc32(data); const local=concatArrays([Uint8Array.from([0x50,0x4b,0x03,0x04]),Uint8Array.from(u16(20)),Uint8Array.from(u16(0)),Uint8Array.from(u16(0)),Uint8Array.from(u16(dt.time)),Uint8Array.from(u16(dt.date)),Uint8Array.from(u32(crc)),Uint8Array.from(u32(data.length)),Uint8Array.from(u32(data.length)),Uint8Array.from(u16(name.length)),Uint8Array.from(u16(0)),name,data]); locals.push(local); const central=concatArrays([Uint8Array.from([0x50,0x4b,0x01,0x02]),Uint8Array.from(u16(20)),Uint8Array.from(u16(20)),Uint8Array.from(u16(0)),Uint8Array.from(u16(0)),Uint8Array.from(u16(dt.time)),Uint8Array.from(u16(dt.date)),Uint8Array.from(u32(crc)),Uint8Array.from(u32(data.length)),Uint8Array.from(u32(data.length)),Uint8Array.from(u16(name.length)),Uint8Array.from(u16(0)),Uint8Array.from(u16(0)),Uint8Array.from(u16(0)),Uint8Array.from(u16(0)),Uint8Array.from(u32(0)),Uint8Array.from(u32(offset)),name]); centrals.push(central); offset += local.length; } const centralBytes=concatArrays(centrals); const end=concatArrays([Uint8Array.from([0x50,0x4b,0x05,0x06]),Uint8Array.from(u16(0)),Uint8Array.from(u16(0)),Uint8Array.from(u16(files.length)),Uint8Array.from(u16(files.length)),Uint8Array.from(u32(centralBytes.length)),Uint8Array.from(u32(offset)),Uint8Array.from(u16(0))]); return concatArrays([...locals,centralBytes,end]); }
function dataUriToBytes(dataUri) { const [, b64] = dataUri.split(','); const bin=atob(b64); const bytes=new Uint8Array(bin.length); for(let i=0;i<bin.length;i++) bytes[i]=bin.charCodeAt(i); return bytes; }
function downloadZip() { const files=[{name:'issues.csv',data:toCsv(rows())},{name:'issues.json',data:JSON.stringify(state.issues,null,2)}]; state.issues.forEach(i => (i.photos || []).forEach((p,idx) => files.push({name:`photos/${p.fileName || `${i.id}_photo-${idx+1}.jpg`}`,data:dataUriToBytes(p.data)}))); download('stratus-heuristic-capture-export.zip',new Blob([makeZip(files)],{type:'application/zip'}),'application/zip'); }

function resetApp() { if (!confirm('Reset local capture data from this device? SharePoint records will not be deleted.')) return; localStorage.removeItem('shc_state'); location.reload(); }

function bind() {
  $('startSessionBtn').onclick = startSession;
  $('changeContextBtn').onclick = changeContext;
  $('cancelContextBtn').onclick = () => state.session ? show('homeScreen') : null;
  $('finishScenarioBtn').onclick = finishScenario;
  $('newIssueBtn').onclick = newIssue;
  $('cancelIssueBtn').onclick = () => show('homeScreen');
  $('saveIssueBtn').onclick = saveIssue;
  $('photo').onchange = async e => { const files = [...e.target.files]; try { for (const f of files) photos.push(await compressImage(f)); renderPhotos(); } catch (err) { alert(`Could not add photo: ${err.message || err}`); } finally { e.target.value = ''; } };
  $('takePhotoBtn').onclick = () => $('photo').click();
  $('exportBtn').onclick = openExport;
  $('backFromExportBtn').onclick = () => show('homeScreen');
  $('downloadCsvBtn').onclick = () => download('issues.csv', toCsv(rows()), 'text/csv');
  $('downloadJsonBtn').onclick = () => download('issues.json', JSON.stringify(state.issues, null, 2), 'application/json');
  $('downloadZipBtn').onclick = downloadZip;
  $('signInBtn').onclick = signIn;
  $('signOutBtn').onclick = signOut;
  $('syncBtn').onclick = syncAll;
  $('resetBtn').onclick = resetApp;
}

(async function init() { renderHeuristics(); bind(); render(); await setupAuth(); })();
