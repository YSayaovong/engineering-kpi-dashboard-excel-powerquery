/* Simple Engineering Analysis
   - Reads Excel in-browser (SheetJS)
   - Shows counts + one bar chart per required group
   - Overdue only when Project ID present, not completed, and due < today
*/

const state = {
  rows: [],
  mapped: { projectId:null, status:null, level:null, due:null, completed:null },
  charts: {}
};

window.addEventListener("DOMContentLoaded", async () => {
  wireUI();
  // Auto-load /data/Project Tracking.xlsx if available
  try {
    const resp = await fetch("./data/Project%20Tracking.xlsx", { cache: "no-store" });
    if (resp.ok) {
      const buf = await resp.arrayBuffer();
      await loadFromArrayBuffer(buf);
    }
  } catch {}
});

function wireUI(){
  document.getElementById("fileInput").addEventListener("change", async (e)=>{
    const f = e.target.files?.[0]; if(!f) return;
    const buf = await f.arrayBuffer();
    await loadFromArrayBuffer(buf);
  });
  document.getElementById("urlForm").addEventListener("submit", async (e)=>{
    e.preventDefault();
    const url = document.getElementById("xlsxUrl").value.trim();
    if(!url) return;
    const resp = await fetch(url, { cache:"no-store" });
    const buf = await resp.arrayBuffer();
    await loadFromArrayBuffer(buf);
  });
  document.getElementById("applyMapping").addEventListener("click", ()=>{
    const ids = ["mapProjectId","mapStatus","mapLevel","mapDue","mapCompleted"];
    const keys = ["projectId","status","level","due","completed"];
    ids.forEach((id,i)=> state.mapped[keys[i]] = document.getElementById(id).value || null);
    refresh();
  });
}

async function loadFromArrayBuffer(buf){
  const wb = XLSX.read(buf, { type:"array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval:null });
  state.rows = rows;
  suggestMappings(rows);
  populateMappingSelectors(Object.keys(rows[0] ?? {}));
  refresh();
}

function suggestMappings(rows){
  const cols = Object.keys(rows[0] ?? {}).map(c => [c, c.toLowerCase()]);
  const pick = (alts) => {
    for (const a of alts){
      const exact = cols.find(([orig,low]) => low === a);
      if (exact) return exact[0];
    }
    for (const a of alts){
      const part = cols.find(([orig,low]) => low.includes(a));
      if (part) return part[0];
    }
    return null;
  };
  state.mapped.projectId = pick(["project id","projectid","pid","id","project"]);
  state.mapped.status   = pick(["status","state","stage","progress"]);
  state.mapped.level    = pick(["level","area","dept","department","section","discipline"]);
  state.mapped.due      = pick(["due date","due","deadline","target"]);
  state.mapped.completed= pick(["completed date","completed","closed date","closed","finish","finished date","done"]);
}

function populateMappingSelectors(columns){
  const ids = ["mapProjectId","mapStatus","mapLevel","mapDue","mapCompleted"];
  const keys = ["projectId","status","level","due","completed"];
  ids.forEach((id,i)=>{
    const sel = document.getElementById(id);
    sel.innerHTML = `<option value="">— Not Mapped —</option>` + columns.map(c => {
      const selected = state.mapped[keys[i]] === c ? "selected" : "";
      return `<option ${selected}>${c}</option>`;
    }).join("");
  });
}

function toDate(v){
  if(!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === "number"){
    // Excel serial date
    const base = new Date(Date.UTC(1899,11,30));
    return new Date(base.getTime() + v*86400000);
  }
  const t = Date.parse(v);
  return isNaN(t) ? null : new Date(t);
}

function getVal(row, key){
  const col = state.mapped[key];
  return col ? row[col] : null;
}

function whichGroup(levelVal=""){
  const v = String(levelVal||"").toLowerCase();
  const isNameplate = v.includes("nameplate") || v.includes("name plate");
  const isNameplateErr = isNameplate && v.includes("error");
  const isTank = v.includes("tank size") || v.includes("tank");
  const isInternalLayout = v.includes("internal layout") || v.includes("internal");
  const isDesignErr = (v.includes("design") && v.includes("error")) || v.includes("design errors");
  const isDetail = v.includes("detailing") || v.includes("detail");
  const isDetailErr = isDetail && v.includes("error");

  if (isNameplate || isNameplateErr) return "np"; // nameplate group
  if (isTank || isInternalLayout || isDesignErr) return "ts"; // tank/internal/design group
  if (isDetail || isDetailErr) return "dt"; // detailing group
  if (v.includes("layout") || v.includes("design")) return "ts";
  return "dt";
}

function isOverdue(row){
  const pid = String(getVal(row,"projectId") ?? "").trim();
  if (!pid) return false;
  const comp = toDate(getVal(row,"completed"));
  if (comp) return false;
  const due = toDate(getVal(row,"due"));
  if (!due) return false;
  const today = new Date();
  const d0 = new Date(due.getFullYear(), due.getMonth(), due.getDate());
  const t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return d0 < t0;
}

function isCompleted(row){
  return !!toDate(getVal(row,"completed"));
}

function refresh(){
  const groups = { np:{total:0,open:0,completed:0,overdue:0},
                   ts:{total:0,open:0,completed:0,overdue:0},
                   dt:{total:0,open:0,completed:0,overdue:0} };

  for (const r of state.rows){
    const g = whichGroup(getVal(r,"level"));
    groups[g].total += 1;
    if (isCompleted(r)){
      groups[g].completed += 1;
    } else {
      groups[g].open += 1;
      if (isOverdue(r)) groups[g].overdue += 1;
    }
  }

  // Update counts
  setCounts("np", groups.np);
  setCounts("ts", groups.ts);
  setCounts("dt", groups.dt);

  // Charts
  renderBar("np_chart", "np", groups.np);
  renderBar("ts_chart", "ts", groups.ts);
  renderBar("dt_chart", "dt", groups.dt);
}

function setCounts(prefix, obj){
  document.getElementById(`${prefix}_total`).textContent = obj.total;
  document.getElementById(`${prefix}_open`).textContent = obj.open;
  document.getElementById(`${prefix}_completed`).textContent = obj.completed;
  document.getElementById(`${prefix}_overdue`).textContent = obj.overdue;
}

function renderBar(canvasId, key, obj){
  if (state.charts[key]){ state.charts[key].destroy(); delete state.charts[key]; }
  const ctx = document.getElementById(canvasId);
  state.charts[key] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: ["Open","Completed","Overdue"],
      datasets: [{ label: "Count", data: [obj.open, obj.completed, obj.overdue] }]
    },
    options: {
      scales: {
        x: { ticks: { color: "#cbd5e1" }, grid: { color: "#1f2937" } },
        y: { ticks: { color: "#cbd5e1" }, grid: { color: "#1f2937" }, beginAtZero: true }
      },
      plugins: { legend: { labels: { color: "#cbd5e1" } } }
    }
  });
}
