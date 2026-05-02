"""
Nomura FI – Power BI-style interactive HTML dashboard generator.
Usage:  python3 powerbi_dashboard.py <clean_excel_path> [output_dir]
Reads the clean Data sheet produced by preprocessing.py and writes a
self-contained HTML file with embedded Chart.js visuals and slicer filters.
"""

import sys
import json
import math
from pathlib import Path
from datetime import datetime

import pandas as pd


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------

def load(path: str) -> pd.DataFrame:
    xf = pd.ExcelFile(path, engine="openpyxl")
    sheet = "Data" if "Data" in xf.sheet_names else xf.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    for col in ["revenue_usd", "meetings_cnt", "events_attended", "reports_read",
                "Rev_Per_Meeting", "Engagement_Score", "Readership_Rate"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in ["client_name", "analyst_nm", "REGION", "client_tier",
                "MiFID_Status", "Flag", "Research_Access_Status", "Client_Focus_Category"]:
        if col in df.columns:
            df[col] = df[col].fillna("Unknown").astype(str).str.strip()
    return df


def fmt_usd(n: float) -> str:
    if n >= 1_000_000:
        return f"${n/1_000_000:.1f}M"
    if n >= 1_000:
        return f"${n/1_000:.0f}K"
    return f"${n:.0f}"


def records(df: pd.DataFrame) -> list:
    """Return every row as a dict with only the fields the dashboard needs."""
    cols = ["client_name", "analyst_nm", "REGION", "client_tier",
            "revenue_usd", "meetings_cnt", "MiFID_Status", "Flag",
            "Rev_Per_Meeting", "Engagement_Score", "mifid_expiry_dt"]
    out = []
    for _, row in df.iterrows():
        r = {}
        for c in cols:
            v = row.get(c, "")
            if isinstance(v, float) and math.isnan(v):
                v = ""
            elif hasattr(v, "item"):          # numpy scalar
                v = v.item()
            r[c] = v
        out.append(r)
    return out


# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Nomura FI – Client Engagement Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;background:#F3F2F1;color:#252423;font-size:13px}
  /* Header */
  #hdr{background:#1F3864;color:#fff;padding:12px 20px;display:flex;align-items:center;justify-content:space-between}
  #hdr h1{font-size:16px;font-weight:700;letter-spacing:1.5px}
  #hdr span{font-size:11px;color:#a8c6e8}
  /* Layout */
  #app{display:flex;height:calc(100vh - 44px)}
  /* Slicer panel */
  #slicers{width:210px;min-width:210px;background:#fff;border-right:1px solid #e1dfdd;padding:12px;overflow-y:auto}
  #slicers h2{font-size:11px;font-weight:700;color:#605e5c;text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px;padding-bottom:6px;border-bottom:2px solid #0078D4}
  .slicer{margin-bottom:16px}
  .slicer label{display:block;font-size:11px;font-weight:600;color:#605e5c;margin-bottom:5px;text-transform:uppercase;letter-spacing:.6px}
  .slicer select{width:100%;padding:5px 7px;border:1px solid #c8c6c4;border-radius:3px;font-size:12px;background:#faf9f8;color:#252423;cursor:pointer}
  .slicer select:focus{outline:none;border-color:#0078D4}
  #reset-btn{width:100%;padding:7px;background:#0078D4;color:#fff;border:none;border-radius:3px;font-size:12px;font-weight:600;cursor:pointer;margin-top:4px;letter-spacing:.5px}
  #reset-btn:hover{background:#106EBE}
  /* Main canvas */
  #main{flex:1;overflow-y:auto;padding:14px 16px}
  /* KPI row */
  #kpis{display:flex;gap:10px;margin-bottom:14px}
  .kpi{flex:1;background:#fff;border-radius:4px;padding:12px 14px;box-shadow:0 1px 4px rgba(0,0,0,.08);border-top:3px solid #0078D4}
  .kpi.warn{border-top-color:#D83B01}
  .kpi.flag{border-top-color:#FFB900}
  .kpi.green{border-top-color:#107C10}
  .kpi .val{font-size:22px;font-weight:700;color:#252423;line-height:1.2}
  .kpi .lbl{font-size:10px;color:#605e5c;text-transform:uppercase;letter-spacing:.7px;margin-top:3px}
  /* Chart grid */
  .chart-row{display:flex;gap:10px;margin-bottom:14px}
  .card{background:#fff;border-radius:4px;box-shadow:0 1px 4px rgba(0,0,0,.08);padding:12px 14px;flex:1;min-width:0}
  .card h3{font-size:11px;font-weight:700;color:#605e5c;text-transform:uppercase;letter-spacing:.7px;margin-bottom:10px}
  .card.wide{flex:2}
  /* Tables */
  .tbl-wrap{overflow-x:auto;max-height:220px;overflow-y:auto}
  table{width:100%;border-collapse:collapse;font-size:11px}
  th{background:#0078D4;color:#fff;padding:6px 10px;text-align:left;font-weight:600;position:sticky;top:0;white-space:nowrap}
  td{padding:5px 10px;border-bottom:1px solid #f3f2f1;white-space:nowrap}
  tr:nth-child(even) td{background:#f9f8f7}
  tr:hover td{background:#EBF3FB}
  .badge{display:inline-block;padding:1px 6px;border-radius:10px;font-size:10px;font-weight:600}
  .badge.exp{background:#FDE7E9;color:#A4262C}
  .badge.ok{background:#DFF6DD;color:#107C10}
  .badge.unk{background:#F3F2F1;color:#605e5c}
  .badge.rev{background:#FFF4CE;color:#7D5C00}
  #no-data{display:none;text-align:center;padding:40px;color:#605e5c;font-size:13px}
</style>
</head>
<body>
<div id="hdr">
  <h1>NOMURA &nbsp;|&nbsp; Fixed Income Client Engagement</h1>
  <span id="hdr-date">__DATE__</span>
</div>
<div id="app">
  <div id="slicers">
    <h2>Filters</h2>
    <div class="slicer"><label>Region</label><select id="f-region"><option value="">All</option></select></div>
    <div class="slicer"><label>Analyst</label><select id="f-analyst"><option value="">All</option></select></div>
    <div class="slicer"><label>Tier</label><select id="f-tier"><option value="">All</option></select></div>
    <div class="slicer"><label>MiFID Status</label><select id="f-mifid"><option value="">All</option></select></div>
    <div class="slicer"><label>Flag</label><select id="f-flag"><option value="">All</option></select></div>
    <button id="reset-btn">Reset Filters</button>
  </div>
  <div id="main">
    <div id="kpis">
      <div class="kpi"><div class="val" id="k-clients">—</div><div class="lbl">Active Clients</div></div>
      <div class="kpi green"><div class="val" id="k-revenue">—</div><div class="lbl">Total Revenue</div></div>
      <div class="kpi"><div class="val" id="k-meetings">—</div><div class="lbl">Total Meetings</div></div>
      <div class="kpi warn"><div class="val" id="k-mifid">—</div><div class="lbl">MiFID Expired</div></div>
      <div class="kpi flag"><div class="val" id="k-flagged">—</div><div class="lbl">Flagged Clients</div></div>
    </div>
    <div class="chart-row">
      <div class="card wide"><h3>Revenue by Analyst</h3><canvas id="ch-analyst" height="160"></canvas></div>
      <div class="card"><h3>Revenue by Region</h3><canvas id="ch-region" height="160"></canvas></div>
    </div>
    <div class="chart-row">
      <div class="card wide"><h3>Top 15 Clients by Revenue</h3><canvas id="ch-clients" height="170"></canvas></div>
      <div class="card"><h3>Client Tier Distribution</h3><canvas id="ch-tier" height="170"></canvas></div>
    </div>
    <div class="chart-row">
      <div class="card wide"><h3>Flagged Clients — High Meetings / Low Revenue</h3>
        <div class="tbl-wrap"><table id="tbl-flagged">
          <thead><tr><th>Client</th><th>Analyst</th><th>Meetings</th><th>Revenue</th><th>MiFID</th></tr></thead>
          <tbody></tbody>
        </table></div>
      </div>
      <div class="card"><h3>MiFID Expired Clients</h3>
        <div class="tbl-wrap"><table id="tbl-mifid">
          <thead><tr><th>Client</th><th>Analyst</th><th>Expiry</th></tr></thead>
          <tbody></tbody>
        </table></div>
      </div>
    </div>
  </div>
</div>

<script>
const RAW = __DATA__;

const PBI_COLORS = ['#0078D4','#107C10','#D83B01','#FFB900','#8764B8',
  '#00B7C3','#E3008C','#004E8C','#498205','#7A7574','#038387','#C239B3'];

function fmtUSD(n){
  if(n>=1e6) return '$'+(n/1e6).toFixed(1)+'M';
  if(n>=1e3) return '$'+(n/1e3).toFixed(0)+'K';
  return '$'+Math.round(n);
}
function fmtNum(n){ return n.toLocaleString(); }

// Populate slicer dropdowns
function unique(key){ return [...new Set(RAW.map(r=>r[key]).filter(v=>v&&v!=='Unknown'))].sort(); }
function populate(id, key){
  const sel=document.getElementById(id);
  unique(key).forEach(v=>{ const o=document.createElement('option'); o.value=o.textContent=v; sel.appendChild(o); });
}
populate('f-region','REGION'); populate('f-analyst','analyst_nm');
populate('f-tier','client_tier'); populate('f-mifid','MiFID_Status');
populate('f-flag','Flag');

// Filter
function getFilters(){
  return {region:document.getElementById('f-region').value,
          analyst:document.getElementById('f-analyst').value,
          tier:document.getElementById('f-tier').value,
          mifid:document.getElementById('f-mifid').value,
          flag:document.getElementById('f-flag').value};
}
function applyFilters(f){
  return RAW.filter(r=>
    (!f.region  || r.REGION===f.region) &&
    (!f.analyst || r.analyst_nm===f.analyst) &&
    (!f.tier    || r.client_tier===f.tier) &&
    (!f.mifid   || r.MiFID_Status===f.mifid) &&
    (!f.flag    || r.Flag===f.flag)
  );
}

// Aggregations
function groupSum(data, key, val='revenue_usd'){
  const m={};
  data.forEach(r=>{ m[r[key]]=(m[r[key]]||0)+r[val]; });
  return m;
}

// Charts (kept as module-level vars so we can destroy/recreate)
let chAnalyst, chRegion, chClients, chTier;

function makeBar(id, labels, values, color, horizontal=true){
  const ctx=document.getElementById(id).getContext('2d');
  return new Chart(ctx,{
    type:'bar',
    data:{labels, datasets:[{data:values, backgroundColor:color||'#0078D4',
      borderRadius:3, borderSkipped:false}]},
    options:{
      indexAxis: horizontal?'y':'x',
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+fmtUSD(c.raw)}}},
      scales:{
        x:{ticks:{callback:v=>fmtUSD(v),font:{size:10}},grid:{color:'#f3f2f1'}},
        y:{ticks:{font:{size:10}},grid:{display:false}}
      }
    }
  });
}

function makeDoughnut(id, labels, values){
  const ctx=document.getElementById(id).getContext('2d');
  return new Chart(ctx,{
    type:'doughnut',
    data:{labels, datasets:[{data:values, backgroundColor:PBI_COLORS,
      borderWidth:2, borderColor:'#fff', hoverOffset:6}]},
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{position:'right', labels:{font:{size:10}, boxWidth:12}},
        tooltip:{callbacks:{label:c=>' '+c.label+': '+fmtUSD(c.raw)}}
      }
    }
  });
}

function badgeMiFID(s){
  const cls=s==='EXPIRED'?'exp':s==='Valid'?'ok':'unk';
  return `<span class="badge ${cls}">${s}</span>`;
}

function render(){
  const f=getFilters();
  const data=applyFilters(f);

  // KPIs
  const totalRev=data.reduce((a,r)=>a+r.revenue_usd,0);
  const totalMtg=data.reduce((a,r)=>a+r.meetings_cnt,0);
  const expMiFID=data.filter(r=>r.MiFID_Status==='EXPIRED').length;
  const flagged =data.filter(r=>r.Flag==='Review').length;
  document.getElementById('k-clients').textContent=fmtNum(data.length);
  document.getElementById('k-revenue').textContent=fmtUSD(totalRev);
  document.getElementById('k-meetings').textContent=fmtNum(totalMtg);
  document.getElementById('k-mifid').textContent=expMiFID;
  document.getElementById('k-flagged').textContent=flagged;

  // Analyst chart (top 10)
  const analystMap=groupSum(data,'analyst_nm');
  const analystSorted=Object.entries(analystMap).sort((a,b)=>b[1]-a[1]).slice(0,10);
  if(chAnalyst) chAnalyst.destroy();
  chAnalyst=makeBar('ch-analyst',
    analystSorted.map(e=>e[0]), analystSorted.map(e=>e[1]), '#0078D4');

  // Region doughnut
  const regionMap=groupSum(data,'REGION');
  const regionE=Object.entries(regionMap).filter(e=>e[0]!=='Unknown').sort((a,b)=>b[1]-a[1]);
  if(chRegion) chRegion.destroy();
  chRegion=makeDoughnut('ch-region', regionE.map(e=>e[0]), regionE.map(e=>e[1]));

  // Top 15 clients
  const clientsSorted=[...data].sort((a,b)=>b.revenue_usd-a.revenue_usd).slice(0,15);
  if(chClients) chClients.destroy();
  chClients=makeBar('ch-clients',
    clientsSorted.map(r=>r.client_name), clientsSorted.map(r=>r.revenue_usd),
    clientsSorted.map((_,i)=>PBI_COLORS[i%PBI_COLORS.length]));

  // Tier doughnut
  const tierMap=groupSum(data,'client_tier');
  const tierE=Object.entries(tierMap).sort((a,b)=>b[1]-a[1]);
  if(chTier) chTier.destroy();
  chTier=makeDoughnut('ch-tier', tierE.map(e=>e[0]), tierE.map(e=>e[1]));

  // Flagged table
  const flaggedData=data.filter(r=>r.Flag==='Review').sort((a,b)=>b.meetings_cnt-a.meetings_cnt);
  const ftbody=document.querySelector('#tbl-flagged tbody');
  ftbody.innerHTML=flaggedData.length?flaggedData.map(r=>`
    <tr><td>${r.client_name}</td><td>${r.analyst_nm}</td>
    <td>${Math.round(r.meetings_cnt)}</td><td>${fmtUSD(r.revenue_usd)}</td>
    <td>${badgeMiFID(r.MiFID_Status)}</td></tr>`).join('')
    :'<tr><td colspan="5" style="text-align:center;color:#605e5c;padding:20px;">No flagged clients</td></tr>';

  // MiFID expired table
  const mifidData=data.filter(r=>r.MiFID_Status==='EXPIRED').sort((a,b)=>a.client_name.localeCompare(b.client_name));
  const mtbody=document.querySelector('#tbl-mifid tbody');
  mtbody.innerHTML=mifidData.length?mifidData.map(r=>`
    <tr><td>${r.client_name}</td><td>${r.analyst_nm}</td>
    <td><span class="badge exp">${r.mifid_expiry_dt||'—'}</span></td></tr>`).join('')
    :'<tr><td colspan="3" style="text-align:center;color:#605e5c;padding:20px;">None</td></tr>';
}

// Wire up slicers
['f-region','f-analyst','f-tier','f-mifid','f-flag'].forEach(id=>{
  document.getElementById(id).addEventListener('change', render);
});
document.getElementById('reset-btn').addEventListener('click',()=>{
  ['f-region','f-analyst','f-tier','f-mifid','f-flag'].forEach(id=>{
    document.getElementById(id).value='';
  });
  render();
});

render();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 powerbi_dashboard.py <clean_excel_path> [output_dir]", file=sys.stderr)
        sys.exit(1)

    excel_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) >= 3 else str(Path(excel_path).parent)

    df = load(excel_path)

    data = records(df)
    data_json = json.dumps(data, default=str, ensure_ascii=False)

    now = datetime.now().strftime("%d %B %Y, %H:%M")
    html = HTML.replace("__DATA__", data_json).replace("__DATE__", now)

    stem = Path(excel_path).stem.replace("Clean_", "").replace("Formatted_", "")
    out_path = Path(output_dir) / f"Nomura_Dashboard_{stem}.html"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")

    print(str(out_path))


if __name__ == "__main__":
    main()
