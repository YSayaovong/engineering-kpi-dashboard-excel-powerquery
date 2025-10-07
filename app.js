/* ===========================================
   Engineering Project Tracker — Analytics
   Reads Excel in-browser (SheetJS) and charts with Chart.js
   =========================================== */

const state = {
  rawRows: [],
  mapped: {
    status: null, level: null, created: null, due: null, completed: null, title: null
  },
  charts: {},
  filteredRows: [],
  categories: ["Design", "Nameplates", "OCR & NC", "Other Levels"]
};

// Auto-load default file if present
window.addEventListener("DOMContentLoaded", async () => {
  wireUI();
  try {
    const resp = await fetch("./data/Project%20Tracking.xlsx", { cache: "no-store" });
    if (resp.ok) {
      const buf = await resp.arrayBuffer();
      await loadFromArrayBuffer(buf);
    }
  } catch {}
});

function wireUI(){
  const fileInput = document.getElementById("fileInput");
  const urlForm = document.getElementById("urlForm");
  fileInput.addEventListener("change", handleFile);
  urlForm.addEventListener("submit", handleUrlLoad);

  document.getElementById("applyMapping").addEventListener("click", applyMapping);
  document.getElementById("applyFilters").addEventListener("click", () => {
    applyFilters(); refreshAll();
  });
  document.getElementById("resetFilters").addEventListener("click", () => {
    resetFilters(); applyFilters(); refreshAll();
  });
  document.getElementById("downloadCsv").addEventListener("click", downloadFilteredCsv);
}

async function handleFile(e){
  const file = e.target.files?.[0];
  if(!file) return;
  const buf = await file.arrayBuffer();
  await loadFromArrayBuffer(buf);
}

async function handleUrlLoad(e){
  e.preventDefault();
  const url = document.getElementById("xlsxUrl").value.trim();
  if(!url) return;
  const resp = await fetch(url, { cache: "no-store" });
  const buf = await resp.arrayBuffer();
  await loadFromArrayBuffer(buf);
}

async function loadFromArrayBuffer(buf){
  const wb = XLSX.read(buf, { type: "array" });
  const firstSheet = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheet];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

  state.rawRows = rows;
  suggestMappings(rows);
  populateMappingSelectors(Object.keys(rows[0] ?? {}));
  applyMapping(); // initial apply with suggested map
  applyFilters();
  refreshAll();
}

/* ---------- Column Mapping ---------- */

const friendlyMap = {
  status: ["status","state","stage","progress"],
  level: ["level","area","dept","department","section"],
  created: ["created","created date","start date","opened","opened date","requested"],
  due: ["due","due date","target","deadline"],
  completed: ["completed","completed date","done","finish","finished date","closed","closed date"],
  title: ["title","task","name","summary","subject","description"]
};

function suggestMappings(rows){
  const columns = Object.keys(rows[0] ?? {});
  const mapOut = {};
  for(const key of Object.keys(friendlyMap)){
    mapOut[key] = bestColumn(columns, friendlyMap[key]);
  }
  state.mapped = mapOut;
}

function bestColumn(columns, candidates){
  const lower = columns.map(c => c.toLowerCase());
  for(const cand of candidates){
    const idx = lower.findIndex(c => c === cand);
    if(idx >= 0) return columns[idx];
  }
  // contains
  for(const cand of candidates){
    const idx = lower.findIndex(c => c.includes(cand));
    if(idx >= 0) return columns[idx];
  }
  return null;
}

function populateMappingSelectors(columns){
  const ids = ["mapStatus","mapLevel","mapCreated","mapDue","mapCompleted","mapTitle"];
  const keys = ["status","level","created","due","completed","title"];

  ids.forEach((id, i) => {
    const sel = document.getElementById(id);
    sel.innerHTML = `<option value="">— Not Mapped —</option>` + columns.map(c => {
      const selected = state.mapped[keys[i]] === c ? "selected" : "";
      return `<option ${selected}>${c}</option>`;
    }).join("");
  });
}

function applyMapping(){
  const ids = ["mapStatus","mapLevel","mapCreated","mapDue","mapCompleted","mapTitle"];
  const keys = ["status","level","created","due","completed","title"];
  ids.forEach((id, i) => state.mapped[keys[i]] = document.getElementById(id).value || null);
  buildFilters();
  applyFilters();
  refreshAll();
}

/* ---------- Categorization ---------- */

function categorize(levelVal = ""){
  const v = String(levelVal || "").toLowerCase();
  if (v.includes("design")) return "Design";
  if (v.includes("nameplate") || v.includes("name plate")) return "Nameplates";
  if (v.includes("ocr") || v.includes("nc") || v.includes("numerical control")) return "OCR & NC";
  return "Other Levels";
}

/* ---------- Filters ---------- */

function buildFilters(){
  const statusSet = new Set();
  const catSet = new Set();

  for(const r of state.rawRows){
    const s = getVal(r, "status");
    if (s) statusSet.add(String(s));
    catSet.add(categorize(getVal(r,"level")));
  }

  const statusSel = document.getElementById("filterStatus");
  statusSel.innerHTML = [...statusSet].sort().map(v => `<option selected>${escapeHtml(v)}</option>`).join("");

  const catSel = document.getElementById("filterCategory");
  catSel.innerHTML = state.categories.map(v => `<option selected>${v}</option>`).join("");

  // date bounds
  const {minDate, maxDate} = createdDateBounds();
  const fromEl = document.getElementById("filterFrom");
  const toEl = document.getElementById("filterTo");
  fromEl.value = minDate ? fmtDateInput(minDate) : "";
  toEl.value = maxDate ? fmtDateInput(maxDate) : "";
}

function applyFilters(){
  const statusChosen = selectedValues(document.getElementById("filterStatus"));
  const catsChosen = selectedValues(document.getElementById("filterCategory"));
  const from = parseDate(document.getElementById("filterFrom").value);
  const to = parseDate(document.getElementById("filterTo").value, true);

  state.filteredRows = state.rawRows.filter(r => {
    const sOk = !statusChosen.length || statusChosen.includes(String(getVal(r,"status")));
    const c = categorize(getVal(r,"level"));
    const cOk = !catsChosen.length || catsChosen.includes(c);
    const cd = toDate(getVal(r,"created"));
    const dOk = (!from || (cd && cd >= from)) && (!to || (cd && cd <= to));
    return sOk && cOk && dOk;
  });
}

function resetFilters(){ buildFilters(); }

/* ---------- Data helpers ---------- */

function getVal(row, key){
  const col = state.mapped[key];
  return col ? row[col] : null;
}

function toDate(v){
  if(!v) return null;
  // SheetJS may provide Date, Serial, or string
  if (v instanceof Date) return v;
  if (typeof v === "number") return XLSX.SSF.parse_date_code(v) ? excelSerialToDate(v) : null;
  const t = Date.parse(v);
  return isNaN(t) ? null : new Date(t);
}

function excelSerialToDate(n){
  // Excel serial days since 1899-12-30
  const utc = new Date(Date.UTC(1899,11,30));
  return new Date(utc.getTime() + n*86400000);
}

function createdDateBounds(){
  const dates = state.rawRows.map(r => toDate(getVal(r,"created"))).filter(Boolean);
  if(!dates.length) return {minDate:null, maxDate:null};
  return {minDate: new Date(Math.min(...dates)), maxDate: new Date(Math.max(...dates))};
}

function fmtDateInput(d){ return d.toISOString().slice(0,10); }
function parseDate(s, endOfDay=false){
  if(!s) return null;
  const d = new Date(s);
  if (endOfDay){ d.setHours(23,59,59,999); }
  return d;
}

function daysBetween(a,b){
  if(!a || !b) return null;
  const ms = b - a;
  return ms/86400000;
}

function selectedValues(sel){
  return Array.from(sel.options).filter(o=>o.selected).map(o=>o.value);
}

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

/* ---------- KPIs ---------- */

function computeKpis(rows){
  const total = rows.length;

  const open = rows.filter(r => !toDate(getVal(r,"completed"))).length;

  // Completed MTD
  const now = new Date();
  const mStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const mEnd = new Date(now.getFullYear(), now.getMonth()+1, 0, 23,59,59,999);
  const completedMTD = rows.filter(r => {
    const cd = toDate(getVal(r,"completed"));
    return cd && cd >= mStart && cd <= mEnd;
  }).length;

  // On-time % last 30 d (completed where completed <= due)
  const last30 = new Date(); last30.setDate(last30.getDate()-30);
  const compLast30 = rows.filter(r => {
    const c = toDate(getVal(r,"completed"));
    return c && c >= last30;
  });
  const onTime = compLast30.length
    ? Math.round(100 * compLast30.filter(r => {
        const c = toDate(getVal(r,"completed"));
        const d = toDate(getVal(r,"due"));
        return c && d && c <= d;
      }).length / compLast30.length)
    : 0;

  // Median cycle (created -> completed days) for completed rows
  const cycles = rows.map(r => {
    const c = toDate(getVal(r,"created"));
    const f = toDate(getVal(r,"completed"));
    const x = daysBetween(c,f);
    return (x!=null && x>=0) ? x : null;
  }).filter(v => v!=null).sort((a,b)=>a-b);
  const med = cycles.length ? (cycles.length%2 ? cycles[(cycles.length-1)/2] : (cycles[cycles.length/2-1]+cycles[cycles.length/2])/2) : 0;

  return { total, open, completedMTD, onTime, medCycle: Math.round(med) };
}

function renderKpis(k){
  document.getElementById("kpiTotal").textContent = k.total.toLocaleString();
  document.getElementById("kpiOpen").textContent = k.open.toLocaleString();
  document.getElementById("kpiCompletedMTD").textContent = k.completedMTD.toLocaleString();
  document.getElementById("kpiOnTime").textContent = `${k.onTime}%`;
  document.getElementById("kpiCycle").textContent = `${k.medCycle}`;
}

/* ---------- Aggregations ---------- */

function groupStatus(rows){
  const m = new Map();
  for(const r of rows){
    const s = String(getVal(r,"status") ?? "Unknown");
    m.set(s, (m.get(s)||0)+1);
  }
  return m;
}

function groupStatusByCategory(rows){
  // categories fixed order
  const byCat = new Map(state.categories.map(c => [c, new Map()]));
  for(const r of rows){
    const cat = categorize(getVal(r,"level"));
    const st = String(getVal(r,"status") ?? "Unknown");
    const m = byCat.get(cat);
    m.set(st, (m.get(st)||0)+1);
  }
  return byCat;
}

function monthlyThroughput(rows){
  // count completed per YYYY-MM
  const m = new Map();
  for(const r of rows){
    const c = toDate(getVal(r,"completed"));
    if(!c) continue;
    const key = `${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`;
    m.set(key, (m.get(key)||0)+1);
  }
  const keys = [...m.keys()].sort();
  return { labels: keys, values: keys.map(k=>m.get(k)) };
}

function wipTrend(rows){
  // open over time at month granularity from min(created) to current month
  const createdDates = rows.map(r=>toDate(getVal(r,"created"))).filter(Boolean);
  if(!createdDates.length) return {labels:[], values:[]};

  const min = new Date(Math.min(...createdDates));
  const today = new Date();
  const cursor = new Date(min.getFullYear(), min.getMonth(), 1);
  const labels = [], values = [];

  while (cursor <= today){
    const periodEnd = new Date(cursor.getFullYear(), cursor.getMonth()+1, 0, 23,59,59,999);
    const openCount = rows.filter(r => {
      const cr = toDate(getVal(r,"created"));
      const cm = toDate(getVal(r,"completed"));
      return cr && cr <= periodEnd && (!cm || cm > periodEnd);
    }).length;
    labels.push(`${cursor.getFullYear()}-${String(cursor.getMonth()+1).padStart(2,"0")}`);
    values.push(openCount);
    cursor.setMonth(cursor.getMonth()+1);
  }
  return {labels, values};
}

/* ---------- Charts ---------- */

function destroyChart(key){
  if(state.charts[key]){ state.charts[key].destroy(); delete state.charts[key]; }
}

function renderCharts(){
  const rows = state.filteredRows;

  // Status Pie
  {
    destroyChart("statusPie");
    const ctx = document.getElementById("statusPie");
    const m = groupStatus(rows);
    state.charts.statusPie = new Chart(ctx, {
      type: "doughnut",
      data: {
        labels: [...m.keys()],
        datasets: [{ data: [...m.values()] }]
      },
      options: { plugins: { legend: { labels: { color: "#cbd5e1" } } } }
    });
  }

  // Category Stacked Bar
  {
    destroyChart("categoryStacked");
    const ctx = document.getElementById("categoryStacked");
    const byCat = groupStatusByCategory(rows);
    const cats = state.categories;
    const allStatuses = new Set();
    for(const m of byCat.values()) for(const s of m.keys()) allStatuses.add(s);
    const statuses = [...allStatuses];

    const datasets = statuses.map((s,i)=>({
      label: s,
      data: cats.map(c => byCat.get(c).get(s) || 0),
      stack: "stack1"
    }));

    state.charts.categoryStacked = new Chart(ctx, {
      type: "bar",
      data: { labels: cats, datasets },
      options: {
        responsive: true,
        scales: {
          x: { ticks:{ color:"#cbd5e1" }, stacked:true, grid:{ color:"#1f2937" } },
          y: { ticks:{ color:"#cbd5e1" }, stacked:true, grid:{ color:"#1f2937" } }
        },
        plugins: { legend: { labels:{ color:"#cbd5e1" } } }
      }
    });
  }

  // Throughput Bar
  {
    destroyChart("throughputBar");
    const ctx = document.getElementById("throughputBar");
    const {labels, values} = monthlyThroughput(rows);
    state.charts.throughputBar = new Chart(ctx, {
      type: "bar",
      data: { labels, datasets: [{ label:"Completed", data: values }] },
      options: {
        scales: {
          x: { ticks:{ color:"#cbd5e1" }, grid:{ color:"#1f2937" } },
          y: { ticks:{ color:"#cbd5e1" }, grid:{ color:"#1f2937" } }
        },
        plugins: { legend:{ labels:{ color:"#cbd5e1" } } }
      }
    });
  }

  // WIP line
  {
    destroyChart("wipLine");
    const ctx = document.getElementById("wipLine");
    const {labels, values} = wipTrend(rows);
    state.charts.wipLine = new Chart(ctx, {
      type: "line",
      data: { labels, datasets: [{ label:"Open Items", data: values, tension: 0.25 }] },
      options: {
        scales: {
          x: { ticks:{ color:"#cbd5e1" }, grid:{ color:"#1f2937" } },
          y: { ticks:{ color:"#cbd5e1" }, grid:{ color:"#1f2937" } }
        },
        plugins: { legend:{ labels:{ color:"#cbd5e1" } } }
      }
    });
  }
}

/* ---------- Table & CSV ---------- */

function renderTable(){
  const rows = state.filteredRows;
  const fields = visibleFields();

  const thead = document.querySelector("#dataTable thead");
  const tbody = document.querySelector("#dataTable tbody");
  thead.innerHTML = `<tr>${fields.map(f=>`<th>${escapeHtml(f)}</th>`).join("")}</tr>`;

  const maxRows = 1000; // safety
  tbody.innerHTML = rows.slice(0,maxRows).map(r => {
    return `<tr>${fields.map(f=>{
      const v = r[f];
      return `<td>${escapeHtml(formatCell(v))}</td>`;
    }).join("")}</tr>`;
  }).join("");
}

function visibleFields(){
  // show mapped columns first, then everything else
  const mappedCols = Object.values(state.mapped).filter(Boolean);
  const all = new Set(Object.keys(state.rawRows[0] ?? {}));
  const rest = [...all].filter(c => !mappedCols.includes(c));
  return [...mappedCols, ...rest];
}

function formatCell(v){
  if (v instanceof Date) return v.toISOString().slice(0,10);
  return v==null ? "" : String(v);
}

function downloadFilteredCsv(){
  const rows = state.filteredRows;
  const fields = visibleFields();
  const header = fields.join(",");
  const body = rows.map(r => fields.map(f => csvEscape(r[f])).join(",")).join("\n");
  const blob = new Blob([header+"\n"+body], {type:"text/csv;charset=utf-8"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "filtered-project-tracking.csv";
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function csvEscape(v){
  if (v==null) return "";
  const s = String(v);
  if (/[",\n]/.test(s)) return `"${s.replace(/"/g,'""')}"`;
  return s;
}

/* ---------- Refresh pipeline ---------- */

function refreshAll(){
  const k = computeKpis(state.filteredRows);
  renderKpis(k);
  renderCharts();
  renderTable();
}
