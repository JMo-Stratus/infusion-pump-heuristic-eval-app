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
  "Language: Use users' language",
  'Control: Users in control',
  'Document: Help and documentation'
];


const canonicalPumpNames = {
  'BD Alaris': 'BD Alaris',
  'BD Alaris – PCU and LVP': 'BD Alaris',
  'BD Alaris – PCU & Large Volume Pump': 'BD Alaris',
  'BD Alaris - PCU and Large Volume Pump': 'BD Alaris',
  'Baxter Sigma Spectrum IQ': 'Baxter Sigma Spectrum IQ',
  'Baxter Spectrum IQ': 'Baxter Sigma Spectrum IQ',
  'Baxter Sigma IQ': 'Baxter Sigma Spectrum IQ',
  'Ivenix': 'Ivenix',
  'Ivenix Infusion System': 'Ivenix',
  'Plum Duo': 'Plum Duo',
  'Plum Duo Infusion System': 'Plum Duo'
};
function canonicalPumpName(name=''){
  return canonicalPumpNames[String(name).trim()] || String(name).trim();
}
function normalizeStoredPumpNames(){
  if(state.session?.device) state.session.device = canonicalPumpName(state.session.device);
  state.issues.forEach(i => { if(i.device) i.device = canonicalPumpName(i.device); });
  state.completedScenarios.forEach(s => { if(s.device) s.device = canonicalPumpName(s.device); });
}

const defaultState = {session:null, issues:[], completedScenarios:[]};
const state = JSON.parse(localStorage.getItem('hef_state') || JSON.stringify(defaultState));
if(!state.completedScenarios) state.completedScenarios = [];
normalizeStoredPumpNames();
let editingId = null;
let photos = [];

const $ = id => document.getElementById(id);
const screens = ['sessionScreen','homeScreen','issueScreen','exportScreen'];
function save(){ localStorage.setItem('hef_state', JSON.stringify(state)); }
function show(id){ screens.forEach(s => $(s).classList.toggle('active', s === id)); render(); }
function severityLabel(v){ return ({2:'Minor',3:'Moderate',4:'Major',5:'Catastrophic'})[v] || '—'; }
function esc(s=''){ return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }
function currentContextKey(){ return `${state.session?.device || ''}||${state.session?.scenario || ''}`; }

function renderHeuristics(){
  $('heuristicChoices').innerHTML = heuristics.map(h => `<label><input type="checkbox" value="${esc(h)}" /> ${esc(h)}</label>`).join('');
}

function renderPhotos(){
  const c = $('photoThumbs');
  if(!photos.length){ c.innerHTML = '<p class="empty-small">No photos added.</p>'; return; }
  c.innerHTML = photos.map((p, idx) => `
    <div class="photo-card">
      <img src="${p.data}" alt="Evidence photo ${idx+1}" />
      <div class="row between"><span>Photo ${idx+1}</span><button type="button" class="remove-photo" data-idx="${idx}">Remove</button></div>
    </div>`).join('');
  document.querySelectorAll('.remove-photo').forEach(btn => btn.onclick = () => { photos.splice(Number(btn.dataset.idx), 1); renderPhotos(); });
}

function issueBelongsToCurrent(i){
  if(!state.session) return false;
  return canonicalPumpName(i.device || '') === canonicalPumpName(state.session.device || '') && String(i.scenario || '') === String(state.session.scenario || '');
}

function render(){
  if(state.session){
    $('sessionSummary').textContent = `${state.session.device} · ${state.session.scenario}`;
    $('currentDevice').textContent = state.session.device;
    $('currentScenario').textContent = state.session.scenario;
  } else {
    $('sessionSummary').textContent = 'No pump/scenario selected';
  }
  const currentIssues = state.issues.filter(issueBelongsToCurrent);
  $('issueCount').textContent = currentIssues.length;
  $('highSeverityCount').textContent = currentIssues.filter(i => Number(i.severity) >= 4).length;
  if(!currentIssues.length){ $('issueList').className='issue-list empty'; $('issueList').textContent='No issues recorded for this pump/scenario yet.'; }
  else {
    $('issueList').className='issue-list';
    $('issueList').innerHTML = currentIssues.map(i => `<div class="issue" data-id="${i.id}">
      <div class="issue-title">${esc(i.title || 'Untitled issue')}</div>
      <div class="issue-meta"><span class="badge sev${i.severity}">${severityLabel(i.severity)}</span><span class="badge">${esc(i.taskPhase)}</span><span class="badge">${(i.photos || []).length} photo(s)</span></div>
      <div class="issue-meta">${esc((i.heuristics || []).join('; '))}</div>
    </div>`).join('');
    document.querySelectorAll('.issue').forEach(el => el.onclick = () => editIssue(el.dataset.id));
  }
}

function setSessionFormFromCurrent(){
  if(state.session){
    $('evaluator').value = state.session.evaluator || '';
    $('device').value = state.session.device || '';
    $('scenario').value = state.session.scenario || '';
    $('softwareVersion').value = state.session.softwareVersion || '';
    $('drugLibrary').value = state.session.drugLibrary || '';
  }
}

function startSession(){
  if(!$('device').value || !$('scenario').value){ alert('Select a pump system and scenario.'); return; }
  state.session = {
    evaluator: $('evaluator').value,
    device: canonicalPumpName($('device').value),
    scenario: $('scenario').value,
    softwareVersion: $('softwareVersion').value,
    drugLibrary: $('drugLibrary').value,
    startedAt: new Date().toISOString()
  };
  save(); show('homeScreen');
}

function changeContext(){
  setSessionFormFromCurrent();
  $('sessionScreenTitle').textContent = 'Change Pump / Scenario';
  show('sessionScreen');
}

function finishScenario(){
  if(!state.session) return;
  const record = {
    device: canonicalPumpName(state.session.device),
    scenario: state.session.scenario,
    evaluator: state.session.evaluator || '',
    finishedAt: new Date().toISOString(),
    issueCount: state.issues.filter(issueBelongsToCurrent).length
  };
  state.completedScenarios.push(record);
  save();
  alert(`Scenario finished:\n${record.device}\n${record.scenario}\n\nYou can now select another pump system and/or scenario.`);
  changeContext();
}

function clearIssueForm(){
  editingId = null; photos = [];
  $('issueFormTitle').textContent = 'New Issue';
  ['issueTitle','issueDescription','evaluatorGoal','notes'].forEach(id => $(id).value = '');
  $('photo').value = '';
  $('taskPhase').value = 'Setup';
  document.querySelectorAll('#heuristicChoices input').forEach(c => c.checked = false);
  document.querySelector('input[name="severity"][value="3"]').checked = true;
  renderPhotos();
}
function newIssue(){ clearIssueForm(); show('issueScreen'); }
function editIssue(id){
  const issue = state.issues.find(i => i.id === id); if(!issue) return;
  editingId = id; photos = (issue.photos || []).map(p => ({...p}));
  $('issueFormTitle').textContent = 'Edit Issue';
  $('issueTitle').value = issue.title || '';
  $('issueDescription').value = issue.description || '';
  $('taskPhase').value = issue.taskPhase || 'Setup';
  $('evaluatorGoal').value = issue.evaluatorGoal || '';
  $('notes').value = issue.notes || '';
  document.querySelectorAll('#heuristicChoices input').forEach(c => c.checked = (issue.heuristics || []).includes(c.value));
  document.querySelector(`input[name="severity"][value="${issue.severity || 3}"]`).checked = true;
  renderPhotos();
  show('issueScreen');
}
function saveIssue(){
  if(!state.session){ alert('Select a pump system and scenario before recording issues.'); return; }
  const selected = [...document.querySelectorAll('#heuristicChoices input:checked')].map(c => c.value);
  const id = editingId || `ISS-${Date.now()}`;
  const prior = state.issues.find(i => i.id === id) || {};
  const item = {
    id,
    timestamp: editingId ? (prior.timestamp || new Date().toISOString()) : new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    evaluator: state.session.evaluator || '',
    device: editingId ? (prior.device || state.session.device) : state.session.device,
    scenario: editingId ? (prior.scenario || state.session.scenario) : state.session.scenario,
    softwareVersion: editingId ? (prior.softwareVersion || state.session.softwareVersion || '') : (state.session.softwareVersion || ''),
    drugLibrary: editingId ? (prior.drugLibrary || state.session.drugLibrary || '') : (state.session.drugLibrary || ''),
    title: $('issueTitle').value,
    description: $('issueDescription').value,
    taskPhase: $('taskPhase').value,
    evaluatorGoal: $('evaluatorGoal').value,
    heuristics: selected,
    severity: document.querySelector('input[name="severity"]:checked').value,
    notes: $('notes').value,
    photos: photos.map((p, idx) => ({...p, fileName: photoFileName(id, idx, p.type)}))
  };
  if(editingId){ state.issues = state.issues.map(i => i.id === editingId ? item : i); }
  else { state.issues.push(item); }
  save(); show('homeScreen');
}
function photoFileName(issueId, idx, type='image/jpeg'){
  const ext = type.includes('png') ? 'png' : type.includes('webp') ? 'webp' : 'jpg';
  return `${issueId}_photo-${idx+1}.${ext}`;
}
function rows(){ return state.issues.map(i => {
  const photoFiles = (i.photos || []).map((p, idx) => p.fileName || photoFileName(i.id, idx, p.type));
  return {
    issue_id:i.id,
    timestamp:i.timestamp,
    evaluator:i.evaluator || state.session?.evaluator || '',
    device:canonicalPumpName(i.device || ''),
    scenario:i.scenario || '',
    software_version:i.softwareVersion || '',
    drug_library_profile:i.drugLibrary || '',
    task_phase:i.taskPhase,
    issue_title:i.title,
    issue_description:i.description,
    evaluator_goal:i.evaluatorGoal,
    heuristics:(i.heuristics || []).join('; '),
    severity:i.severity,
    severity_label:severityLabel(i.severity),
    notes:i.notes,
    photo_count:photoFiles.length,
    photo_files:photoFiles.join('; ')
  };
}); }
function toCsv(data){
  const cols = Object.keys(data[0] || {issue_id:'',timestamp:'',evaluator:'',device:'',scenario:'',software_version:'',drug_library_profile:'',task_phase:'',issue_title:'',issue_description:'',evaluator_goal:'',heuristics:'',severity:'',severity_label:'',notes:'',photo_count:'',photo_files:''});
  return [cols.join(','), ...data.map(r => cols.map(c => `"${String(r[c] ?? '').replace(/"/g,'""')}"`).join(','))].join('\n');
}
function download(name, content, type){
  const blob = content instanceof Blob ? content : new Blob([content], {type}); const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(()=>URL.revokeObjectURL(url), 1000);
}
function openExport(){ $('exportPreview').value = toCsv(rows()); show('exportScreen'); }

function dataUriToBytes(dataUri){
  const [meta, b64] = dataUri.split(',');
  const type = (meta.match(/data:(.*?);/) || [,'application/octet-stream'])[1];
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);
  return {bytes, type};
}

// Minimal ZIP writer using STORE method, with CRC32.
const crcTable = (() => {
  const table = new Uint32Array(256);
  for(let n=0;n<256;n++){
    let c=n;
    for(let k=0;k<8;k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    table[n]=c>>>0;
  }
  return table;
})();
function crc32(bytes){
  let c = 0xffffffff;
  for(const b of bytes) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function strBytes(s){ return new TextEncoder().encode(s); }
function dosDateTime(date=new Date()){
  const time = (date.getHours()<<11) | (date.getMinutes()<<5) | Math.floor(date.getSeconds()/2);
  const d = ((date.getFullYear()-1980)<<9) | ((date.getMonth()+1)<<5) | date.getDate();
  return {time, date:d};
}
function u16(n){ return [n & 255, (n>>>8)&255]; }
function u32(n){ return [n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255]; }
function concatArrays(arrays){
  const len = arrays.reduce((a,b)=>a+b.length,0); const out = new Uint8Array(len); let off=0;
  arrays.forEach(a=>{ out.set(a, off); off += a.length; }); return out;
}
function makeZip(files){
  let offset=0; const locals=[]; const centrals=[]; const dt = dosDateTime();
  for(const f of files){
    const name = strBytes(f.name); const bytes = f.bytes instanceof Uint8Array ? f.bytes : strBytes(f.bytes);
    const crc = crc32(bytes); const size = bytes.length;
    const local = new Uint8Array([
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(dt.time), ...u16(dt.date),
      ...u32(crc), ...u32(size), ...u32(size), ...u16(name.length), ...u16(0), ...name, ...bytes
    ]);
    locals.push(local);
    const central = new Uint8Array([
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(dt.time), ...u16(dt.date),
      ...u32(crc), ...u32(size), ...u32(size), ...u16(name.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset), ...name
    ]);
    centrals.push(central); offset += local.length;
  }
  const centralSize = centrals.reduce((a,b)=>a+b.length,0);
  const end = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(centralSize), ...u32(offset), ...u16(0)]);
  return new Blob([concatArrays([...locals, ...centrals, end])], {type:'application/zip'});
}
function downloadZip(){
  const csv = toCsv(rows());
  const json = JSON.stringify({activeSession:state.session, completedScenarios:state.completedScenarios, issues:state.issues.map(i => ({...i, photos:(i.photos||[]).map(p => ({fileName:p.fileName, timestamp:p.timestamp, type:p.type}))}))}, null, 2);
  const files = [
    {name:'issues.csv', bytes:csv},
    {name:'issues.json', bytes:json}
  ];
  for(const issue of state.issues){
    (issue.photos || []).forEach((p, idx) => {
      const {bytes} = dataUriToBytes(p.data);
      files.push({name:`photos/${p.fileName || photoFileName(issue.id, idx, p.type)}`, bytes});
    });
  }
  download(`heuristic_eval_export_${new Date().toISOString().slice(0,10)}.zip`, makeZip(files), 'application/zip');
}

$('startSessionBtn').onclick = startSession;
$('cancelContextBtn').onclick = () => { if(state.session){ setSessionFormFromCurrent(); show('homeScreen'); } else { show('sessionScreen'); } };
$('newIssueBtn').onclick = newIssue;
$('cancelIssueBtn').onclick = () => show('homeScreen');
$('saveIssueBtn').onclick = saveIssue;
$('exportBtn').onclick = openExport;
$('backFromExportBtn').onclick = () => show('homeScreen');
$('downloadZipBtn').onclick = downloadZip;
$('downloadCsvBtn').onclick = () => download('heuristic_eval_issues.csv', toCsv(rows()), 'text/csv');
$('downloadJsonBtn').onclick = () => download('heuristic_eval_issues.json', JSON.stringify({activeSession:state.session, completedScenarios:state.completedScenarios, issues:state.issues}, null, 2), 'application/json');
$('resetBtn').onclick = () => { if(confirm('Clear the current session and all locally stored issues?')){ localStorage.removeItem('hef_state'); location.reload(); } };
$('takePhotoBtn').onclick = () => $('photo').click();
$('changeContextBtn').onclick = changeContext;
$('finishScenarioBtn').onclick = finishScenario;
$('photo').onchange = e => {
  const file = e.target.files[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    photos.push({data: reader.result, type: file.type || 'image/jpeg', originalName: file.name || '', timestamp: new Date().toISOString()});
    $('photo').value = '';
    renderPhotos();
  };
  reader.readAsDataURL(file);
};

renderHeuristics();
setSessionFormFromCurrent();
show(state.session ? 'homeScreen' : 'sessionScreen');
