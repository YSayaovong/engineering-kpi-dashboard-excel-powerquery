/* Ultra-simple: read Excel, show a bar (totals by group) and a pie (status distribution).
   Overdue rule:
     - Only if Project ID present (non-blank)
     - Not completed
     - Due date exists and is earlier than today
*/

const state = { rows: [], mapped: { projectId:null, status:null, level:null, due:null, completed:null }, charts:{} };

window.addEventListener("DOMContentLoaded", () => {
  document.getElementById("fileInput").addEventListener("change", onFile);
  document.getElementById("loadDefault").addEventListener("click", loadDefault);
  // try autoload silently
  loadDefault(true);
});

async function onFile(e){
  const f = e.target.files?.[0]; if(!f) return;
  document.getElementById("fileName").textContent = f.name;
  const buf = await f.arrayBuffer();
  await loadFromArrayBuffer(buf);
}

async function loadDefault(silent=false){
  try{
    const resp = await fetch("./data/Project%20Tracking.xlsx", { cache:"no-store" });
    if(!resp.ok){ if(!silent) alert("No default file found at /data/Project Tracking.xlsx"); return; }
    document.getElementById("fileName").textContent = "/data/Project Tracking.xlsx";
    const buf = await resp.arrayBuffer();
    await loadFromArrayBuffer(buf);
  }catch(e){
    if(!silent) alert("Could not load default file.");
  }
}

async function loadFromArrayBuffer(buf){
  const wb = XLSX.read(buf, { type:"array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval:null });
  state.rows = rows;
  suggestMappings(rows);
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
  state.mapped.status    = pick(["status","state","stage","progress"]);
  state.mapped.level     = pick(["level","area","dept","department","section","discipline"]);
  state.mapped.due       = pick(["due date","due","deadline","target"]);
  state.mapped.completed = pick(["completed date","completed","closed date","closed","finish","finished date","done"]);
}

function getVal(row, key){ const col = state.mapped[key]; return col ? row[col] : null; }
function toDate(v){
  if(!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === "number"){
    const base = new Date(Date.UTC(1899,11,30));
    return new Date(base.getTime() + v*86400000);
  }
  const t = Date.parse(v); return isNaN(t) ? null : new Date(t);
}
function isCompleted(r){ return !!toDate(getVal(r,"completed")); }
function isOverdue(r){
  const pid = String(getVal(r,"projectId") ?? "").trim(); if(!pid) return false;
  if (isCompleted(r)) return false;
  const due = toDate(getVal(r,"due")); if(!due) return false;
  const today = new Date();
  const d0 = new Date(due.getFullYear(), due.getMonth(), due.getDate());
  const t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return d0 < t0;
}

function whichGroup(level=""){
  const v = String(level||"").toLowerCase();
  const isNameplate = v.includes("nameplate") || v.includes("name plate");
  const isNameplateErr = isNameplate && v.includes("error");
  const isTank = v.includes("tank size") || v.includes("tank");
  const isInternal = v.includes("internal layout") || v.includes("internal");
  const isDesignErr = (v.includes("design") && v.includes("error")) || v.includes("design errors");
  const isDetail = v.includes("detailing") || v.includes("detail");
  const isDetailErr = isDetail && v.includes("error");
  if (isNameplate || isNameplateErr) return "Nameplates + Nameplate Errors";
  if (isTank || isInternal || isDesignErr) return "Tank Size + Internal Layout + Design Errors";
  if (isDetail || isDetailErr) return "Detailing + Detailing Errors";
  if (v.includes("layout") || v.includes("design")) return "Tank Size + Internal Layout + Design Errors";
  return "Detailing + Detailing Errors";
}

function refresh(){
  if (!state.rows.length) return;

  // Aggregations
  const groups = ["Nameplates + Nameplate Errors",
                  "Tank Size + Internal Layout + Design Errors",
                  "Detailing + Detailing Errors"];
  const totalsByGroup = new Map(groups.map(g => [g,0]));

  let open=0, completed=0, overdue=0;

  for(const r of state.rows){
    const g = whichGroup(getVal(r,"level"));
    totalsByGroup.set(g, (totalsByGroup.get(g)||0) + 1);

    if (isCompleted(r)) completed++;
    else {
      open++;
      if (isOverdue(r)) overdue++;
    }
  }

  renderBar(groups.map(g=>g.split(" + ")[0]), groups.map(g=>totalsByGroup.get(g)||0));
  renderPie(["Open","Completed","Overdue"], [open, completed, overdue]);
}

function destroyChart(key){
  if (state.charts[key]){ state.charts[key].destroy(); delete state.charts[key]; }
}

function renderBar(labels, values){
  destroyChart("bar");
  const ctx = document.getElementById("barChart");
  state.charts.bar = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets: [{ label:"Total", data: values }] },
    options: {
      scales: {
        x: { ticks:{ color:"#cbd5e1" }, grid:{ color:"#1f2937" } },
        y: { ticks:{ color:"#cbd5e1" }, grid:{ color:"#1f2937" }, beginAtZero:true }
      },
      plugins: { legend:{ labels:{ color:"#cbd5e1" } } }
    }
  });
}

function renderPie(labels, values){
  destroyChart("pie");
  const ctx = document.getElementById("pieChart");
  state.charts.pie = new Chart(ctx, {
    type: "pie",
    data: { labels, datasets: [{ data: values }] },
    options: { plugins: { legend:{ labels:{ color:"#cbd5e1" } } } }
  });
}
