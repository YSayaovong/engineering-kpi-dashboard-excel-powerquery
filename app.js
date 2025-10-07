// ===== Configuration =====
// If your CSV column names differ, set them here or via the UI mapping:
let CFG = {
  level: '',       // e.g., 'Level' or 'Area'
  status: '',      // e.g., 'Status'
  created: '',     // e.g., 'Created' or 'Start Date'
  due: '',         // e.g., 'Due Date'
  completed: ''    // e.g., 'Completed'
};

// ===== Utilities =====
function parseDate(s){
  if(!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function daysBetween(a,b){ if(!a||!b) return NaN; return Math.round((b - a)/(1000*60*60*24)); }
function fmt(n){ return isFinite(n) ? n.toLocaleString() : '—'; }
function pct(n){ return isFinite(n) ? (n*100).toFixed(1)+'%' : '—'; }
function toMonthKey(d){ if(!d) return ''; const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'); return `${y}-${m}`; }

function categorize(levelVal){
  const text = String(levelVal||'').toLowerCase();
  if (text.match(/\bdesign\b/)) return 'Design';
  if (text.includes('nameplate')) return 'Nameplates';
  if (text.match(/\b(ocr|o\.c\.r\.)\b/) || text.match(/\b(nc|n\.c\.)\b/) || text.includes('ocr/nc')) return 'OCR & NC';
  return 'Other Levels';
}

function inferStatus(row){
  const sKey = CFG.status && row.hasOwnProperty(CFG.status) ? CFG.status : null;
  if (sKey){
    const v = String(row[sKey]||'').trim();
    if (v) return v;
  }
  // fallback: use completed date
  const compKey = CFG.completed && row.hasOwnProperty(CFG.completed) ? CFG.completed : null;
  const comp = compKey ? row[compKey] : row['_Completed'];
  return comp ? 'Completed' : 'Open';
}

function barChart(canvasId, labels, datasets, stacked=false){
  const ctx = document.getElementById(canvasId).getContext('2d');
  return new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: { legend: { position: 'bottom' } },
      scales: { x: { stacked }, y: { stacked, beginAtZero: true } }
    }
  });
}

function hist(values, edges){
  const bins = new Array(edges.length-1).fill(0);
  values.forEach(v=>{
    const x = Number(v);
    if(!isFinite(x)) return;
    for(let i=0;i<edges.length-1;i++){
      if(x >= edges[i] && x < edges[i+1]){ bins[i]++; return; }
    }
    if(x >= edges[edges.length-1]) bins[bins.length-1]++;
  });
  return bins;
}

function tableFromData(el, rows, columns){
  if(!columns || !columns.length){
    const keys = new Set();
    rows.slice(0, 50).forEach(r => Object.keys(r).forEach(k=>{ if(!k.startsWith('_')) keys.add(k); }));
    columns = Array.from(keys);
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  columns.forEach(c=>{
    const th=document.createElement('th'); th.textContent=c; trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);
  const tbody=document.createElement('tbody');
  rows.forEach(r=>{
    const tr=document.createElement('tr');
    columns.forEach(c=>{
      const td=document.createElement('td');
      td.textContent = r[c] ?? '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  el.innerHTML='';
  el.appendChild(table);
}

// ===== Load CSV and render =====
async function loadCSV(path){
  return new Promise((resolve, reject) => {
    Papa.parse(path, {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: results => resolve(results.data),
      error: err => reject(err)
    });
  });
}

function applyMappingFromUI(){
  CFG.level = document.getElementById('cfg-level').value.trim();
  CFG.status = document.getElementById('cfg-status').value.trim();
  CFG.created = document.getElementById('cfg-created').value.trim();
  CFG.due = document.getElementById('cfg-due').value.trim();
  CFG.completed = document.getElementById('cfg-completed').value.trim();
}

document.getElementById('applyCfg').addEventListener('click', () => {
  render(); // re-render with new mapping
});

async function render(){
  const rows = await loadCSV('data/project-tracking.csv');

  // Infer columns if not provided
  function findCol(regexList){
    const cols = rows.length ? Object.keys(rows[0]) : [];
    for(const r of regexList){
      const re = new RegExp(r, 'i');
      const hit = cols.find(c => re.test(c));
      if(hit) return hit;
    }
    return '';
  }

  if(!CFG.level) CFG.level = findCol(['level','area','dept','department','station','work\s*center','process','stage']);
  if(!CFG.status) CFG.status = findCol(['status','state','phase','progress']);
  if(!CFG.created) CFG.created = findCol(['created','opened','start']);
  if(!CFG.due) CFG.due = findCol(['due','deadline']);
  if(!CFG.completed) CFG.completed = findCol(['completed','closed','finish']);

  // Enrich
  const now = new Date();
  const recs = rows.map(r => {
    const created = parseDate(r[CFG.created]);
    const due = parseDate(r[CFG.due]);
    const completed = parseDate(r[CFG.completed]);
    const status = inferStatus({...r, '_Completed': r[CFG.completed]});
    const age = created ? (completed ? daysBetween(created, completed) : daysBetween(created, now)) : NaN;
    const onTime = (due && completed) ? (completed <= due) : null;
    const category = categorize(r[CFG.level] || '');

    return {
      ...r,
      _Category: category,
      _Status: status,
      _AgeDays: isFinite(age) ? age : '',
      _Created: created ? created.toISOString() : '',
      _Due: due ? due.toISOString() : '',
      _Completed: completed ? completed.toISOString() : '',
      _Month: created ? toMonthKey(created) : ''
    };
  });

  // KPIs
  const total = recs.length;
  const open = recs.filter(r => String(r._Status).toLowerCase()!=='completed').length;
  const ages = recs.map(r=>Number(r._AgeDays)).filter(x=>isFinite(x));
  const avgAge = ages.length ? (ages.reduce((a,b)=>a+b,0)/ages.length) : NaN;
  const onTimeVals = recs.map(r=>r._Due && r._Completed ? (new Date(r._Completed) <= new Date(r._Due)) : null).filter(v=>v===true || v===false);
  const onTimePct = onTimeVals.length ? (onTimeVals.filter(Boolean).length / onTimeVals.length) : NaN;

  document.getElementById('kpi-total').textContent = fmt(total);
  document.getElementById('kpi-open').textContent = fmt(open);
  document.getElementById('kpi-age').textContent = isFinite(avgAge) ? Math.round(avgAge) : '—';
  document.getElementById('kpi-ontime').textContent = pct(onTimePct);

  // Grouping
  const design = recs.filter(r => r._Category === 'Design');
  const nameplates = recs.filter(r => r._Category === 'Nameplates');
  const ocrnc = recs.filter(r => r._Category === 'OCR & NC');
  const others = recs.filter(r => r._Category === 'Other Levels');

  // Tables
  const cols = rows.length ? Object.keys(rows[0]) : [];
  tableFromData(document.getElementById('design-table'), design, cols);
  tableFromData(document.getElementById('nameplates-table'), nameplates, cols);
  tableFromData(document.getElementById('ocrnc-table'), ocrnc, cols);
  tableFromData(document.getElementById('other-table'), others, cols);

  // Charts
  const cats = ['Design','Nameplates','OCR & NC','Other Levels'];
  const statuses = Array.from(new Set(recs.map(r=>r._Status)));
  const datasets = statuses.map(s => ({
    label: s,
    data: cats.map(cat => recs.filter(r => r._Category===cat && r._Status===s).length)
  }));
  barChart('statusByCategory', cats, datasets, true);

  const completed = recs.filter(r => r._Completed);
  const months = Array.from(new Set(completed.map(r=>r._Completed.slice(0,7)))).sort();
  const throughput = months.map(m => completed.filter(r => String(r._Completed).startsWith(m)).length);
  barChart('monthlyThroughput', months, [{ label: 'Completed', data: throughput }], false);

  const agesList = ages;
  const edges = [0,7,14,21,30,60,90,180,365];
  const bins = hist(agesList, edges);
  const labels = ['0-6','7-13','14-20','21-29','30-59','60-89','90-179','180-365+'];
  barChart('agingDist', labels, [{ label: 'Items', data: bins }], false);
}

// Initial render
render().catch(err => {
  console.error('Failed to load CSV:', err);
  alert('Could not load data/project-tracking.csv. Make sure the file exists in your repo.');
});
