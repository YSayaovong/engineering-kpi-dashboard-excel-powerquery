/* Mini dashboard: reads Excel in browser and shows simple counts + bar chart per group */
const state = { rows: [], charts: {}, mapped: { projectId:null, status:null, level:null, due:null, completed:null } };

window.addEventListener("DOMContentLoaded", () => {
  document.getElementById("fileInput").addEventListener("change", onFile);
  document.getElementById("loadDefault").addEventListener("click", loadDefault);
  document.getElementById("downloadCsv").addEventListener("click", downloadCsv);
});

async function onFile(e){
  const f = e.target.files?.[0]; if(!f) return;
  const buf = await f.arrayBuffer();
  await loadFromArrayBuffer(buf);
}

async function loadDefault(){
  try{
    const resp = await fetch("./data/Project%20Tracking.xlsx", { cache:"no-store" });
    if(!resp.ok) return;
    const buf = await resp.arrayBuffer();
    await loadFromArrayBuffer(buf);
  }catch{}
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

function getVal(row, key){
  const col = state.mapped[key]; return col ? row[col] : null;
}
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
  const today = new Date(), d0 = new Date(due.getFullYear(), due.getMonth(), due.getDate());
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
  if (isNameplate || isNameplateErr) return "np";
  if (isTank || isInternal || isDesignErr) return "ts";
  if (isDetail || isDetailErr) return "dt";
  if (v.includes("layout") || v.includes("design")) return "ts";
  return "dt";
}

function refresh(){
  const gr = {
    np:{total:0, open:0, completed:0, overdue:0},
    ts:{total:0, open:0, completed:0, overdue:0},
    dt:{total:0, open:0, completed:0, overdue:0},
  };

  for(const r of state.rows){
    const g = whichGroup(getVal(r,"level"));
    gr[g].total++;
    if (isCompleted(r)) gr[g].completed++; else {
      gr[g].open++;
      if (isOverdue(r)) gr[g].overdue++;
    }
  }

  // global
  const totals = ["np","ts","dt"].reduce((a,k)=>{
    a.total += gr[k].total; a.open += gr[k].open; a.completed += gr[k].completed; a.overdue += gr[k].overdue; return a;
  }, {total:0,open:0,completed:0,overdue:0});
  document.getElementById("k_total").textContent = totals.total;
  document.getElementById("k_open").textContent = totals.open;
  document.getElementById("k_completed").textContent = totals.completed;
  document.getElementById("k_overdue").textContent = totals.overdue;

  // per group counts
  setCounts("np", gr.np);
  setCounts("ts", gr.ts);
  setCounts("dt", gr.dt);

  // charts
  renderBar("np_chart","np",gr.np);
  renderBar("ts_chart","ts",gr.ts);
  renderBar("dt_chart","dt",gr.dt);
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
    data: { labels: ["Open","Completed","Overdue"], datasets: [{ label: "Count", data: [obj.open,obj.completed,obj.overdue] }] },
    options: {
      scales: {
        x: { ticks:{ color:"#cbd5e1" }, grid:{ color:"#1f2937" } },
        y: { ticks:{ color:"#cbd5e1" }, grid:{ color:"#1f2937" }, beginAtZero:true }
      },
      plugins: { legend:{ labels:{ color:"#cbd5e1" } } }
    }
  });
}

function downloadCsv(){
  const rows = [
    ["Group","Total","Open","Completed","Overdue"],
    ["Nameplates + Nameplate Errors",
      document.getElementById("np_total").textContent,
      document.getElementById("np_open").textContent,
      document.getElementById("np_completed").textContent,
      document.getElementById("np_overdue").textContent],
    ["Tank Size + Internal Layout + Design Errors",
      document.getElementById("ts_total").textContent,
      document.getElementById("ts_open").textContent,
      document.getElementById("ts_completed").textContent,
      document.getElementById("ts_overdue").textContent],
    ["Detailing + Detailing Errors",
      document.getElementById("dt_total").textContent,
      document.getElementById("dt_open").textContent,
      document.getElementById("dt_completed").textContent,
      document.getElementById("dt_overdue").textContent]
  ];
  const csv = rows.map(r => r.map(v => /[",\n]/.test(String(v)) ? `"${String(v).replace(/"/g,'""')}"` : String(v)).join(",")).join("\n");
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = "summary.csv"; document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}
