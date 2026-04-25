import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, ReferenceLine, Area, AreaChart, ComposedChart,
  Line, LineChart, Scatter, ScatterChart, ZAxis
} from "recharts";

const KHULNA_LAT = 22.8098;
const KHULNA_LON = 89.5644;
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// ─── PHYSICS FORMULAS ────────────────────────────────────────────────────────

function calcWetBulb(T, RH) {
  if (T == null || RH == null) return null;
  const Tw = T * Math.atan(0.151977 * Math.pow(RH + 8.313659, 0.5))
    + Math.atan(T + RH)
    - Math.atan(RH - 1.676331)
    + 0.00391838 * Math.pow(RH, 1.5) * Math.atan(0.023101 * RH)
    - 4.686035;
  return parseFloat(Tw.toFixed(2));
}

function calcWBGT(T, RH) {
  if (T == null || RH == null) return null;
  const es = 6.112 * Math.exp((17.67 * T) / (T + 243.5));
  const e  = (RH / 100) * es;
  return parseFloat((0.567 * T + 0.393 * e + 3.94).toFixed(2));
}

function getWBGTRisk(wbgt) {
  if (wbgt == null)  return { level:"Unknown",  color:"#6b7280", index:0 };
  if (wbgt < 28)     return { level:"Safe",      color:"#4ade80", index:1 };
  if (wbgt < 32)     return { level:"Caution",   color:"#facc15", index:2 };
  if (wbgt < 35)     return { level:"Danger",    color:"#f97316", index:3 };
  return               { level:"Extreme",    color:"#ef4444", index:4 };
}

// ─── STATISTICS ──────────────────────────────────────────────────────────────

// Mann-Kendall test — returns { S, Z, pValue, trend }
function mannKendall(series) {
  const n = series.length;
  if (n < 4) return null;
  let S = 0;
  for (let i = 0; i < n - 1; i++) {
    for (let j = i + 1; j < n; j++) {
      const diff = series[j] - series[i];
      if (diff > 0) S++;
      else if (diff < 0) S--;
    }
  }
  // Variance (no tied groups assumed for annual means)
  const varS = (n * (n - 1) * (2 * n + 5)) / 18;
  let Z = 0;
  if (S > 0)      Z = (S - 1) / Math.sqrt(varS);
  else if (S < 0) Z = (S + 1) / Math.sqrt(varS);

  // Two-tailed p-value from standard normal
  const pValue = 2 * (1 - normalCDF(Math.abs(Z)));

  let trend = "No significant trend";
  if (pValue < 0.01 && Z > 0) trend = "Significant increasing trend (p<0.01)";
  else if (pValue < 0.05 && Z > 0) trend = "Significant increasing trend (p<0.05)";
  else if (pValue < 0.1  && Z > 0) trend = "Marginally increasing trend (p<0.10)";
  else if (pValue < 0.01 && Z < 0) trend = "Significant decreasing trend (p<0.01)";
  else if (pValue < 0.05 && Z < 0) trend = "Significant decreasing trend (p<0.05)";

  return { S, Z: parseFloat(Z.toFixed(3)), pValue: parseFloat(pValue.toFixed(4)), trend };
}

// Sen's Slope estimator — median of all pairwise slopes
function senSlope(series) {
  const slopes = [];
  for (let i = 0; i < series.length - 1; i++) {
    for (let j = i + 1; j < series.length; j++) {
      slopes.push((series[j] - series[i]) / (j - i));
    }
  }
  slopes.sort((a, b) => a - b);
  const mid = Math.floor(slopes.length / 2);
  const median = slopes.length % 2 === 0
    ? (slopes[mid - 1] + slopes[mid]) / 2
    : slopes[mid];
  return parseFloat(median.toFixed(4));
}

// Standard normal CDF (Abramowitz & Stegun approximation)
function normalCDF(z) {
  const t = 1 / (1 + 0.2316419 * Math.abs(z));
  const poly = t * (0.319381530 + t * (-0.356563782 + t * (1.781477937 + t * (-1.821255978 + t * 1.330274429))));
  const pdf  = Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
  const cdf  = 1 - pdf * poly;
  return z >= 0 ? cdf : 1 - cdf;
}

function avg(arr) {
  const v = arr.filter(x => x != null);
  if (!v.length) return null;
  return parseFloat((v.reduce((a,b)=>a+b,0)/v.length).toFixed(2));
}

// ─── DATA GROUPING ────────────────────────────────────────────────────────────

function groupByMonth(rows) {
  const m = {};
  rows.forEach(r => {
    const [yr, mo] = r.date.split("-");
    const k = `${yr}-${mo}`;
    if (!m[k]) m[k] = { label:`${MONTHS[parseInt(mo)-1]} ${yr}`, maxT:[],minT:[],meanT:[],hum:[],sol:[],tw:[],wbgt:[],danger:0,extreme:0,caution:0,total:0 };
    if (r.maxTemp  != null) m[k].maxT.push(r.maxTemp);
    if (r.minTemp  != null) m[k].minT.push(r.minTemp);
    if (r.meanTemp != null) m[k].meanT.push(r.meanTemp);
    if (r.humidity != null) m[k].hum.push(r.humidity);
    if (r.solar    != null) m[k].sol.push(r.solar);
    if (r.wetBulb  != null) m[k].tw.push(r.wetBulb);
    if (r.wbgt     != null) m[k].wbgt.push(r.wbgt);
    m[k].total++;
    if (r.wbgt >= 28 && r.wbgt < 32) m[k].caution++;
    if (r.wbgt >= 32 && r.wbgt < 35) m[k].danger++;
    if (r.wbgt >= 35) m[k].extreme++;
  });
  return Object.entries(m).map(([k,v]) => ({
    key:k, label:v.label,
    maxTemp:avg(v.maxT), minTemp:avg(v.minT), meanTemp:avg(v.meanT),
    humidity:avg(v.hum), solar:avg(v.sol), wetBulb:avg(v.tw), wbgt:avg(v.wbgt),
    cautionDays:v.caution, dangerDays:v.danger, extremeDays:v.extreme, totalDays:v.total,
  }));
}

function groupByYear(rows) {
  const y = {};
  rows.forEach(r => {
    const yr = r.date.split("-")[0];
    if (!y[yr]) y[yr] = { tw:[],wbgt:[],temp:[],caution:0,danger:0,extreme:0,total:0 };
    if (r.wetBulb  != null) y[yr].tw.push(r.wetBulb);
    if (r.wbgt     != null) y[yr].wbgt.push(r.wbgt);
    if (r.meanTemp != null) y[yr].temp.push(r.meanTemp);
    y[yr].total++;
    if (r.wbgt >= 28 && r.wbgt < 32) y[yr].caution++;
    if (r.wbgt >= 32 && r.wbgt < 35) y[yr].danger++;
    if (r.wbgt >= 35) y[yr].extreme++;
  });
  return Object.entries(y).sort().map(([yr,v]) => ({
    year:yr,
    avgTw:   avg(v.tw),
    avgWbgt: avg(v.wbgt),
    avgTemp: avg(v.temp),
    cautionDays: v.caution,
    dangerDays:  v.danger,
    extremeDays: v.extreme,
    totalDays:   v.total,
  }));
}

// ─── UI HELPERS ───────────────────────────────────────────────────────────────

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background:"rgba(4,8,15,0.98)", border:"1px solid #1e3a50", borderRadius:10, padding:"10px 15px", fontSize:12, color:"#e0f4ff" }}>
      <div style={{ fontWeight:700, color:"#7dd3fc", marginBottom:5 }}>{label}</div>
      {payload.map((p,i) => <div key={i} style={{ color:p.color, marginBottom:2 }}>{p.name}: <b>{p.value != null ? p.value : "—"}{p.unit||""}</b></div>)}
    </div>
  );
};

const iStyle = { background:"rgba(255,255,255,0.06)", border:"1px solid #1e4d6b", borderRadius:8, color:"#e0f4ff", padding:"7px 11px", fontSize:12, fontFamily:"Georgia,serif", outline:"none", colorScheme:"dark" };
const box    = { background:"rgba(255,255,255,0.025)", border:"1px solid #102d4a", borderRadius:16, padding:"18px 14px 14px", marginBottom:16 };
const iv     = n => Math.max(0, Math.floor(n/14)-1);

function pBadge(p) {
  if (p < 0.01) return { label:"p < 0.01 ✓✓", color:"#4ade80" };
  if (p < 0.05) return { label:"p < 0.05 ✓",  color:"#a3e635" };
  if (p < 0.10) return { label:"p < 0.10 ~",  color:"#facc15" };
  return              { label:`p = ${p} ✗`,   color:"#f87171" };
}

// ─── MANN-KENDALL RESULT CARD ─────────────────────────────────────────────────
function MKCard({ title, mk, slope, unit, color }) {
  if (!mk) return null;
  const pb = pBadge(mk.pValue);
  return (
    <div style={{ background:"rgba(255,255,255,0.03)", border:`1px solid ${color}44`, borderRadius:14, padding:"16px 18px", flex:"1 1 280px" }}>
      <div style={{ fontSize:10, color, letterSpacing:3, textTransform:"uppercase", marginBottom:8 }}>Mann-Kendall · {title}</div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:12 }}>
        {[
          { label:"Kendall S",    value: mk.S },
          { label:"Z Statistic",  value: mk.Z },
          { label:"p-value",      value: <span style={{ color:pb.color, fontWeight:700 }}>{pb.label}</span> },
          { label:"Sen's Slope",  value: <span style={{ color: slope > 0 ? "#f87171" : "#4ade80", fontWeight:700 }}>{slope > 0 ? "+" : ""}{slope} {unit}/yr</span> },
        ].map((s,i) => (
          <div key={i} style={{ background:"rgba(0,0,0,0.2)", borderRadius:8, padding:"8px 10px" }}>
            <div style={{ fontSize:10, color:"#4a8aaa", marginBottom:3 }}>{s.label}</div>
            <div style={{ fontSize:15, fontWeight:700, color:"#e0f4ff" }}>{s.value}</div>
          </div>
        ))}
      </div>
      <div style={{ fontSize:12, color: mk.Z > 0 ? "#fca5a5" : "#86efac", fontStyle:"italic", borderTop:"1px solid #0f2a3f", paddingTop:8 }}>
        {mk.Z > 0 ? "▲" : "▼"} {mk.trend}
      </div>
    </div>
  );
}

// ─── MAIN COMPONENT ───────────────────────────────────────────────────────────
export default function KhulnaHeatStress() {
  const [rawData,     setRawData]     = useState(null);
  const [loading,     setLoading]     = useState(true);
  const [error,       setError]       = useState(null);
  const [downloading, setDownloading] = useState(false);

  const todayStr    = new Date().toISOString().split("T")[0];
  const currentYear = new Date().getFullYear().toString();

  const [startDate, setStartDate] = useState("2019-01-01");
  const [endDate,   setEndDate]   = useState(todayStr);
  const [tab,       setTab]       = useState("mk");       // mk | heatstroke | wetbulb | trend | temp | humidity | solar
  const [view,      setView]      = useState("monthly");

  const years = ["2019","2020","2021","2022","2023","2024","2025","2026"];

  // ── FETCH ──
  useEffect(() => {
    (async () => {
      try {
        const url = `https://archive-api.open-meteo.com/v1/archive?latitude=${KHULNA_LAT}&longitude=${KHULNA_LON}&start_date=2019-01-01&end_date=${todayStr}&daily=temperature_2m_max,temperature_2m_min,temperature_2m_mean,relative_humidity_2m_mean,shortwave_radiation_sum&timezone=Asia%2FDhaka`;
        const json = await (await fetch(url)).json();
        setRawData(json);
      } catch { setError("Failed to fetch ERA5 data."); }
      finally  { setLoading(false); }
    })();
  }, []);

  // ── PROCESS ──
  const allDailyRows = useMemo(() => {
    if (!rawData) return [];
    return rawData.daily.time.map((d,i) => {
      const T  = rawData.daily.temperature_2m_mean[i];
      const RH = rawData.daily.relative_humidity_2m_mean[i];
      const tw   = calcWetBulb(T, RH);
      const wbgt = calcWBGT(T, RH);
      return { date:d, maxTemp:rawData.daily.temperature_2m_max[i], minTemp:rawData.daily.temperature_2m_min[i], meanTemp:T, humidity:RH, solar:rawData.daily.shortwave_radiation_sum[i], wetBulb:tw, wbgt, wbgtRisk:getWBGTRisk(wbgt).level };
    });
  }, [rawData]);

  const filteredRows = useMemo(() =>
    allDailyRows.filter(r => r.date >= startDate && r.date <= endDate),
    [allDailyRows, startDate, endDate]
  );

  const monthlyData = useMemo(() => groupByMonth(filteredRows), [filteredRows]);
  const yearlyData  = useMemo(() => groupByYear(filteredRows),  [filteredRows]);

  // ── MANN-KENDALL on annual means ──
  const mkStats = useMemo(() => {
    if (yearlyData.length < 4) return null;
    const twSeries   = yearlyData.map(y => y.avgTw).filter(Boolean);
    const wbgtSeries = yearlyData.map(y => y.avgWbgt).filter(Boolean);
    const tempSeries = yearlyData.map(y => y.avgTemp).filter(Boolean);
    const dangerSeries = yearlyData.map(y => y.dangerDays);
    return {
      tw:     { mk: mannKendall(twSeries),     slope: senSlope(twSeries)     },
      wbgt:   { mk: mannKendall(wbgtSeries),   slope: senSlope(wbgtSeries)   },
      temp:   { mk: mannKendall(tempSeries),   slope: senSlope(tempSeries)   },
      danger: { mk: mannKendall(dangerSeries), slope: senSlope(dangerSeries) },
    };
  }, [yearlyData]);

  const chartData = view === "monthly"
    ? monthlyData
    : filteredRows.map(r => ({ ...r, label: r.date.slice(5) }));

  // ── SUMMARY STATS ──
  const allTw    = filteredRows.map(r => r.wetBulb).filter(Boolean);
  const allWbgt  = filteredRows.map(r => r.wbgt).filter(Boolean);
  const allTemps = filteredRows.map(r => r.meanTemp).filter(Boolean);
  const allHum   = filteredRows.map(r => r.humidity).filter(Boolean);
  const allSolar = filteredRows.map(r => r.solar).filter(Boolean);
  const allMax   = filteredRows.map(r => r.maxTemp).filter(Boolean);
  const allMin   = filteredRows.map(r => r.minTemp).filter(Boolean);

  const dangerDays  = filteredRows.filter(r => r.wbgt >= 32 && r.wbgt < 35).length;
  const extremeDays = filteredRows.filter(r => r.wbgt >= 35).length;
  const cautionDays = filteredRows.filter(r => r.wbgt >= 28 && r.wbgt < 32).length;
  const maxTw       = allTw.length ? Math.max(...allTw).toFixed(2) : "—";
  const maxWbgt     = allWbgt.length ? Math.max(...allWbgt).toFixed(2) : "—";

  // ── QUICK DATE SELECTS ──
  const setYear    = y => { setStartDate(`${y}-01-01`); setEndDate(y === currentYear ? todayStr : `${y}-12-31`); };
  const setAllTime = () => { setStartDate("2019-01-01"); setEndDate(todayStr); };
  const isAllTime  = startDate === "2019-01-01" && endDate === todayStr;
  const isYear     = y => startDate === `${y}-01-01` && (endDate === `${y}-12-31` || (y === currentYear && endDate === todayStr));

  // ── EXCEL DOWNLOAD ──
  const downloadExcel = () => {
    setDownloading(true);
    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Daily
      const ws1 = XLSX.utils.aoa_to_sheet([
        ["Date","Max Temp (°C)","Min Temp (°C)","Mean Temp (°C)","Humidity (%)","Solar (MJ/m²)","Wet Bulb Temp (°C)","WBGT (°C)","Heat Stroke Risk"],
        ...filteredRows.map(r=>[r.date,r.maxTemp,r.minTemp,r.meanTemp,r.humidity,r.solar,r.wetBulb,r.wbgt,r.wbgtRisk])
      ]);
      ws1["!cols"] = [12,14,14,15,12,14,18,12,16].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws1, "Daily Data");

      // Sheet 2: Monthly
      const ws2 = XLSX.utils.aoa_to_sheet([
        ["Month","Avg Max T","Avg Min T","Avg Mean T","Avg Humidity","Avg Solar","Avg Wet Bulb","Avg WBGT","Caution Days","Danger Days","Extreme Days"],
        ...monthlyData.map(m=>[m.label,m.maxTemp,m.minTemp,m.meanTemp,m.humidity,m.solar,m.wetBulb,m.wbgt,m.cautionDays,m.dangerDays,m.extremeDays])
      ]);
      ws2["!cols"] = [14,12,12,12,14,12,14,12,13,13,13].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws2, "Monthly Averages");

      // Sheet 3: Yearly
      const ws3 = XLSX.utils.aoa_to_sheet([
        ["Year","Avg Tw (°C)","Avg WBGT (°C)","Avg Temp (°C)","Caution Days","Danger Days","Extreme Days","Total Days"],
        ...yearlyData.map(y=>[y.year,y.avgTw,y.avgWbgt,y.avgTemp,y.cautionDays,y.dangerDays,y.extremeDays,y.totalDays])
      ]);
      ws3["!cols"] = [8,14,14,14,14,14,14,12].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws3, "Yearly Risk Summary");

      // Sheet 4: Mann-Kendall Results
      const mkRows = [["Variable","Kendall S","Z Statistic","p-value","Significant?","Sen's Slope (per year)","Trend Direction"]];
      if (mkStats) {
        const entries = [
          ["Wet Bulb Temp (°C)", mkStats.tw],
          ["WBGT (°C)",          mkStats.wbgt],
          ["Mean Temp (°C)",     mkStats.temp],
          ["Danger Days (count)",mkStats.danger],
        ];
        entries.forEach(([name, {mk, slope}]) => {
          mkRows.push([name, mk.S, mk.Z, mk.pValue, mk.pValue < 0.05 ? "Yes" : "No", slope, mk.trend]);
        });
      }
      const ws4 = XLSX.utils.aoa_to_sheet(mkRows);
      ws4["!cols"] = [22,12,14,10,12,22,40].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws4, "Mann-Kendall Results");

      // Sheet 5: Summary & Citations
      const ws5 = XLSX.utils.aoa_to_sheet([
        ["Metric","Value"],
        ["Study Area","Khulna, Bangladesh"],
        ["Coordinates",`${KHULNA_LAT}°N, ${KHULNA_LON}°E`],
        ["Data Range",`${startDate} to ${endDate}`],
        ["Total Days Analyzed", filteredRows.length],
        ["",""],
        ["Avg Mean Temperature (°C)", avg(allTemps)],
        ["Avg Relative Humidity (%)", avg(allHum)],
        ["Avg Wet Bulb Temperature (°C)", avg(allTw)],
        ["Max Wet Bulb Temperature (°C)", maxTw],
        ["Avg WBGT (°C)", avg(allWbgt)],
        ["Max WBGT (°C)", maxWbgt],
        ["",""],
        ["Caution Days (WBGT 28–32°C)", cautionDays],
        ["Danger Days (WBGT 32–35°C)", dangerDays],
        ["Extreme Days (WBGT ≥35°C)", extremeDays],
        ["",""],
        mkStats ? ["MK Z — Wet Bulb Temp", mkStats.tw.mk.Z] : [],
        mkStats ? ["MK p-value — Wet Bulb Temp", mkStats.tw.mk.pValue] : [],
        mkStats ? ["Sen's Slope — Wet Bulb Temp (°C/yr)", mkStats.tw.slope] : [],
        mkStats ? ["MK Z — WBGT", mkStats.wbgt.mk.Z] : [],
        mkStats ? ["MK p-value — WBGT", mkStats.wbgt.mk.pValue] : [],
        mkStats ? ["Sen's Slope — WBGT (°C/yr)", mkStats.wbgt.slope] : [],
        ["",""],
        ["Data Source","ERA5 Reanalysis, ECMWF via Open-Meteo API"],
        ["Citation 1","Hersbach et al. (2020). The ERA5 global reanalysis. QJRMS, 146(730). doi:10.1002/qj.3803"],
        ["Citation 2","Stull, R. (2011). Wet-Bulb Temperature from RH and Air Temperature. J. Appl. Meteor. Climatol., 50, 2267–2269."],
        ["Citation 3","Raymond et al. (2020). Heat and humidity too severe for human tolerance. Science Advances, 6(19). doi:10.1126/sciadv.aaw1838"],
        ["Citation 4","Bernard & Pourmoghani (1999). Prediction of workplace WBGT. Appl. Occup. Environ. Hyg., 14(2), 126–134."],
        ["Citation 5","Sen, P.K. (1968). Estimates of regression coefficient based on Kendall's tau. JASA, 63(324), 1379–1389."],
        ["Downloaded on", new Date().toLocaleDateString()],
      ]);
      ws5["!cols"] = [{ wch:32 }, { wch:80 }];
      XLSX.utils.book_append_sheet(wb, ws5, "Summary & Citations");

      XLSX.writeFile(wb, `Khulna_HeatStress_MannKendall_${startDate}_to_${endDate}.xlsx`);
    } catch(e) { alert("Download failed: " + e.message); }
    finally { setDownloading(false); }
  };

  // ─── RENDER ───────────────────────────────────────────────────────────────
  const TABS = [
    ["mk",         "📊 Mann-Kendall"],
    ["heatstroke", "🔴 Heat Stroke Risk"],
    ["wetbulb",    "💧 Wet Bulb Temp"],
    ["trend",      "📈 Year Trend"],
    ["temp",       "🌡️ Temperature"],
    ["humidity",   "💧 Humidity"],
    ["solar",      "☀️ Solar"],
  ];

  return (
    <div style={{ minHeight:"100vh", background:"linear-gradient(160deg,#030710 0%,#050c18 60%,#060d1c 100%)", fontFamily:"Georgia,serif", color:"#cce8f8", paddingBottom:56 }}>

      {/* ── HEADER ── */}
      <div style={{ padding:"20px 20px 14px", borderBottom:"1px solid #0c2030", background:"rgba(0,0,0,0.45)", backdropFilter:"blur(12px)", position:"sticky", top:0, zIndex:20 }}>
        <div style={{ maxWidth:1040, margin:"0 auto", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:10 }}>
          <div>
            <div style={{ fontSize:9, letterSpacing:4, color:"#ef4444", textTransform:"uppercase", marginBottom:2 }}>Heat Stress Research · Khulna, Bangladesh · ERA5 + Mann-Kendall</div>
            <h1 style={{ margin:0, fontSize:22, fontWeight:700, color:"#fff" }}>🌡️ Khulna Heat Stroke Risk Index</h1>
            <div style={{ fontSize:11, color:"#64a7c8", marginTop:1 }}>Wet Bulb Temp · WBGT · Mann-Kendall Trend · Sen's Slope · 2019–Present</div>
          </div>
          <button onClick={downloadExcel} disabled={downloading || !filteredRows.length} style={{
            display:"flex", alignItems:"center", gap:7, padding:"9px 16px", borderRadius:10,
            border:"1px solid #ef4444", background:"linear-gradient(135deg,#3b0000,#5c0a0a)",
            color:"#fca5a5", cursor:"pointer", fontSize:12, fontFamily:"inherit", fontWeight:600,
            boxShadow:"0 2px 16px rgba(239,68,68,0.2)", opacity: !filteredRows.length ? 0.5 : 1
          }}>
            📥 {downloading ? "Exporting…" : `Download Excel (${filteredRows.length.toLocaleString()} days)`}
          </button>
        </div>
      </div>

      <div style={{ maxWidth:1040, margin:"0 auto", padding:"0 14px" }}>

        {/* ── DATE FILTER ── */}
        <div style={{ ...box, marginTop:18, padding:"16px 18px" }}>
          <div style={{ fontSize:9, color:"#38bdf8", letterSpacing:3, textTransform:"uppercase", marginBottom:10 }}>📅 Date Range Filter</div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:10, alignItems:"flex-end" }}>
            <div><div style={{ fontSize:9, color:"#4a8aaa", marginBottom:4 }}>START</div><input type="date" value={startDate} min="2019-01-01" max={endDate} onChange={e=>setStartDate(e.target.value)} style={iStyle}/></div>
            <div style={{ color:"#1e4d6b", fontSize:16, paddingBottom:2 }}>→</div>
            <div><div style={{ fontSize:9, color:"#4a8aaa", marginBottom:4 }}>END</div><input type="date" value={endDate} min={startDate} max={todayStr} onChange={e=>setEndDate(e.target.value)} style={iStyle}/></div>
            <div style={{ width:1, height:32, background:"#0f3455", alignSelf:"center" }}/>
            <div>
              <div style={{ fontSize:9, color:"#4a8aaa", marginBottom:4 }}>QUICK SELECT</div>
              <div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>
                <button onClick={setAllTime} style={{ padding:"4px 10px", borderRadius:6, border:`1px solid ${isAllTime?"#38bdf8":"#1e4d6b"}`, background:isAllTime?"#0c3a5e":"transparent", color:isAllTime?"#7dd3fc":"#4a8aaa", cursor:"pointer", fontSize:10, fontFamily:"inherit" }}>All</button>
                {years.map(y=>(
                  <button key={y} onClick={()=>setYear(y)} style={{ padding:"4px 10px", borderRadius:6, border:`1px solid ${isYear(y)?"#38bdf8":"#1e4d6b"}`, background:isYear(y)?"#0c3a5e":"transparent", color:isYear(y)?"#7dd3fc":"#4a8aaa", cursor:"pointer", fontSize:10, fontFamily:"inherit" }}>{y}</button>
                ))}
              </div>
            </div>
          </div>
          {filteredRows.length > 0 && <div style={{ marginTop:7, fontSize:10, color:"#4a8aaa" }}>Showing <b style={{ color:"#7dd3fc" }}>{filteredRows.length.toLocaleString()} days</b> · {startDate} → {endDate}</div>}
        </div>

        {loading && <div style={{ textAlign:"center", padding:"80px 0", color:"#38bdf8" }}><div style={{ fontSize:42 }}>🌀</div><div style={{ marginTop:10, letterSpacing:2 }}>Loading ERA5 archive…</div></div>}
        {error   && <div style={{ textAlign:"center", padding:"60px 0", color:"#f87171" }}><div style={{ fontSize:30 }}>⚠️</div><div style={{ marginTop:8 }}>{error}</div></div>}

        {!loading && !error && filteredRows.length > 0 && (<>

          {/* ── STAT CARDS ── */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(148px,1fr))", gap:10, marginBottom:16 }}>
            {[
              { label:"Avg Wet Bulb",    value: avg(allTw)    != null ? avg(allTw)+"°C"    : "—", icon:"💧", color:"#38bdf8", sub:"Stull 2011" },
              { label:"Max Wet Bulb",    value: maxTw+"°C",                                         icon:"⚠️", color:"#f97316", sub:"Peak recorded" },
              { label:"Avg WBGT",        value: avg(allWbgt)  != null ? avg(allWbgt)+"°C"  : "—", icon:"🌡️", color:"#facc15", sub:"Heat stroke index" },
              { label:"Max WBGT",        value: maxWbgt+"°C",                                        icon:"🔥", color:"#ef4444", sub:"Peak recorded" },
              { label:"🟡 Caution Days", value: cautionDays,                                          icon:"",   color:"#facc15", sub:"WBGT 28–32°C" },
              { label:"🟠 Danger Days",  value: dangerDays,                                           icon:"",   color:"#f97316", sub:"WBGT 32–35°C" },
              { label:"🔴 Extreme Days", value: extremeDays,                                          icon:"",   color:"#ef4444", sub:"WBGT ≥35°C" },
            ].map((s,i) => (
              <div key={i} style={{ background:"rgba(255,255,255,0.04)", border:"1px solid #1a3a50", borderRadius:12, padding:"12px 14px" }}>
                <div style={{ fontSize:18, marginBottom:4 }}>{s.icon}</div>
                <div style={{ fontSize:19, fontWeight:700, color:s.color }}>{s.value}</div>
                <div style={{ fontSize:10, color:"#94a3b8", marginTop:2 }}>{s.label}</div>
                <div style={{ fontSize:9, color:"#4a6a7a", marginTop:1, fontStyle:"italic" }}>{s.sub}</div>
              </div>
            ))}
          </div>

          {/* ── TAB BAR ── */}
          <div style={{ display:"flex", gap:3, background:"rgba(255,255,255,0.03)", borderRadius:12, padding:4, border:"1px solid #0f3455", marginBottom:14, flexWrap:"wrap" }}>
            {TABS.map(([v,l]) => (
              <button key={v} onClick={()=>setTab(v)} style={{ padding:"6px 13px", borderRadius:8, border:"none", background:tab===v?"#0c3a5e":"transparent", color:tab===v?"#7dd3fc":"#4a8aaa", cursor:"pointer", fontSize:11, fontFamily:"inherit", transition:"all 0.15s", fontWeight:tab===v?700:400 }}>{l}</button>
            ))}
          </div>

          {/* ── VIEW TOGGLE (hide for mk tab) ── */}
          {tab !== "mk" && tab !== "trend" && (
            <div style={{ display:"flex", gap:3, background:"rgba(255,255,255,0.03)", borderRadius:10, padding:3, border:"1px solid #0f3455", marginBottom:14, width:"fit-content" }}>
              {[["monthly","Monthly Avg"],["daily","Daily View"]].map(([v,l])=>(
                <button key={v} onClick={()=>setView(v)} style={{ padding:"5px 14px", borderRadius:7, border:"none", background:view===v?"#0c3a5e":"transparent", color:view===v?"#7dd3fc":"#4a8aaa", cursor:"pointer", fontSize:11, fontFamily:"inherit" }}>{l}</button>
              ))}
            </div>
          )}

          {/* ════════════════════════════════════════════════════════════════ */}
          {/* TAB: MANN-KENDALL                                               */}
          {/* ════════════════════════════════════════════════════════════════ */}
          {tab === "mk" && (
            <>
              {/* Explanation */}
              <div style={{ ...box, background:"rgba(139,92,246,0.06)", border:"1px solid #4c1d9588", padding:"16px 20px", marginBottom:16 }}>
                <div style={{ fontSize:9, color:"#a78bfa", letterSpacing:3, textTransform:"uppercase", marginBottom:8 }}>📐 What These Numbers Mean</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))", gap:10, fontSize:11, color:"#94a3b8", lineHeight:1.7 }}>
                  <div><b style={{ color:"#c4b5fd" }}>Mann-Kendall S:</b> Sum of all pairwise trend signs. Positive = upward, Negative = downward.</div>
                  <div><b style={{ color:"#c4b5fd" }}>Z Statistic:</b> Standardized test statistic. |Z| {">"} 1.96 means p {"<"} 0.05 (significant).</div>
                  <div><b style={{ color:"#c4b5fd" }}>p-value:</b> Probability the trend is due to chance. p {"<"} 0.05 = statistically significant ✓</div>
                  <div><b style={{ color:"#c4b5fd" }}>Sen's Slope:</b> Rate of change per year. E.g. +0.05°C/yr = warming 0.05°C every year.</div>
                </div>
              </div>

              {/* MK Cards */}
              {mkStats ? (
                <div style={{ display:"flex", flexWrap:"wrap", gap:12, marginBottom:16 }}>
                  <MKCard title="Wet Bulb Temperature"  mk={mkStats.tw.mk}     slope={mkStats.tw.slope}     unit="°C"   color="#38bdf8"/>
                  <MKCard title="WBGT"                  mk={mkStats.wbgt.mk}   slope={mkStats.wbgt.slope}   unit="°C"   color="#f97316"/>
                  <MKCard title="Mean Air Temperature"  mk={mkStats.temp.mk}   slope={mkStats.temp.slope}   unit="°C"   color="#facc15"/>
                  <MKCard title="Danger Days Count"     mk={mkStats.danger.mk} slope={mkStats.danger.slope} unit=" days" color="#ef4444"/>
                </div>
              ) : (
                <div style={{ textAlign:"center", color:"#4a8aaa", padding:"30px 0" }}>Need at least 4 years of data for Mann-Kendall. Select a wider date range.</div>
              )}

              {/* Copy-paste text for paper */}
              {mkStats && (
                <div style={{ ...box, background:"rgba(56,189,248,0.04)", border:"1px solid #0c4a6e" }}>
                  <div style={{ fontSize:9, color:"#38bdf8", letterSpacing:3, textTransform:"uppercase", marginBottom:10 }}>📝 Ready-to-Paste Results Text for Your Paper</div>
                  <div style={{ fontSize:12, color:"#cbd5e1", lineHeight:1.9, background:"rgba(0,0,0,0.3)", borderRadius:10, padding:"14px 16px" }}>
                    {"The Mann-Kendall trend test revealed a "}
                    <span style={{ color: mkStats.tw.mk.pValue < 0.05 ? "#4ade80" : "#fbbf24" }}>
                      {mkStats.tw.mk.trend.toLowerCase()}
                    </span>
                    {" in annual mean wet bulb temperature over Khulna during the study period (S = "}
                    <b style={{ color:"#7dd3fc" }}>{mkStats.tw.mk.S}</b>
                    {", Z = "}
                    <b style={{ color:"#7dd3fc" }}>{mkStats.tw.mk.Z}</b>
                    {", p = "}
                    <b style={{ color: mkStats.tw.mk.pValue < 0.05 ? "#4ade80" : "#fbbf24" }}>{mkStats.tw.mk.pValue}</b>
                    {"). Sen's slope estimator indicates a rate of change of "}
                    <b style={{ color:"#f97316" }}>{mkStats.tw.slope > 0 ? "+" : ""}{mkStats.tw.slope}°C per year</b>
                    {" for wet bulb temperature. Similarly, WBGT showed a "}
                    <span style={{ color: mkStats.wbgt.mk.pValue < 0.05 ? "#4ade80" : "#fbbf24" }}>
                      {mkStats.wbgt.mk.trend.toLowerCase()}
                    </span>
                    {" (Z = "}
                    <b style={{ color:"#7dd3fc" }}>{mkStats.wbgt.mk.Z}</b>
                    {", p = "}
                    <b style={{ color: mkStats.wbgt.mk.pValue < 0.05 ? "#4ade80" : "#fbbf24" }}>{mkStats.wbgt.mk.pValue}</b>
                    {", Sen's slope = "}
                    <b style={{ color:"#f97316" }}>{mkStats.wbgt.slope > 0 ? "+" : ""}{mkStats.wbgt.slope}°C/yr</b>
                    {"). The annual number of WBGT danger days (≥32°C) showed a "}
                    <span style={{ color: mkStats.danger.mk.pValue < 0.05 ? "#4ade80" : "#fbbf24" }}>
                      {mkStats.danger.mk.trend.toLowerCase()}
                    </span>
                    {" (Z = "}
                    <b style={{ color:"#7dd3fc" }}>{mkStats.danger.mk.Z}</b>
                    {", p = "}
                    <b style={{ color: mkStats.danger.mk.pValue < 0.05 ? "#4ade80" : "#fbbf24" }}>{mkStats.danger.mk.pValue}</b>
                    {") at a rate of "}
                    <b style={{ color:"#f97316" }}>{mkStats.danger.slope > 0 ? "+" : ""}{mkStats.danger.slope} days/yr</b>
                    {"."}
                  </div>
                  <div style={{ fontSize:10, color:"#4a8aaa", marginTop:8, fontStyle:"italic" }}>
                    ↑ Copy this paragraph directly into your Results section. Replace [study period] with your actual dates.
                  </div>
                </div>
              )}

              {/* Annual trend table */}
              {yearlyData.length > 0 && (
                <div style={box}>
                  <div style={{ fontSize:9, color:"#a78bfa", letterSpacing:3, textTransform:"uppercase", marginBottom:12 }}>Annual Summary Table — For Your Paper</div>
                  <div style={{ overflowX:"auto" }}>
                    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
                      <thead>
                        <tr style={{ color:"#64a7c8", borderBottom:"1px solid #1e4d6b" }}>
                          {["Year","Avg Tw (°C)","Avg WBGT (°C)","Avg T (°C)","🟡 Caution","🟠 Danger","🔴 Extreme","Total Days"].map((h,i)=>(
                            <th key={i} style={{ padding:"7px 11px", textAlign:"left", fontSize:10, fontWeight:600 }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {yearlyData.map((y,i)=>(
                          <tr key={i} style={{ borderBottom:"1px solid #0a1e2e", background:i%2===0?"rgba(255,255,255,0.02)":"transparent" }}>
                            <td style={{ padding:"7px 11px", color:"#7dd3fc", fontWeight:700 }}>{y.year}</td>
                            <td style={{ padding:"7px 11px", color:"#38bdf8" }}>{y.avgTw}</td>
                            <td style={{ padding:"7px 11px", color: y.avgWbgt >= 32?"#ef4444":y.avgWbgt >= 28?"#f97316":"#facc15" }}>{y.avgWbgt}</td>
                            <td style={{ padding:"7px 11px", color:"#fb923c" }}>{y.avgTemp}</td>
                            <td style={{ padding:"7px 11px", color:"#facc15" }}>{y.cautionDays}</td>
                            <td style={{ padding:"7px 11px", color:"#f97316" }}>{y.dangerDays}</td>
                            <td style={{ padding:"7px 11px", color:"#ef4444", fontWeight:y.extremeDays>0?700:400 }}>{y.extremeDays}</td>
                            <td style={{ padding:"7px 11px", color:"#64a7c8" }}>{y.totalDays}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </>
          )}

          {/* ════════════════════════════════════════════════════════════════ */}
          {/* TAB: HEAT STROKE RISK                                           */}
          {/* ════════════════════════════════════════════════════════════════ */}
          {tab === "heatstroke" && (
            <div style={box}>
              <div style={{ paddingLeft:6, marginBottom:10 }}>
                <div style={{ fontSize:9, color:"#ef4444", letterSpacing:3, textTransform:"uppercase" }}>Heat Stroke Risk Index (WBGT)</div>
                <div style={{ fontSize:14, fontWeight:700, color:"#e0f4ff", marginTop:2 }}>{view==="monthly" ? "Monthly Average WBGT + Danger Days" : "Daily WBGT Values"}</div>
              </div>
              {/* Risk legend */}
              <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginBottom:12 }}>
                {[["< 28°C","Safe","#4ade80"],["28–32°C","Caution","#facc15"],["32–35°C","Danger","#f97316"],["≥ 35°C","Extreme","#ef4444"]].map(([r,l,c],i)=>(
                  <div key={i} style={{ display:"flex", alignItems:"center", gap:5, background:"rgba(255,255,255,0.03)", border:`1px solid ${c}33`, borderRadius:7, padding:"4px 10px" }}>
                    <div style={{ width:8, height:8, borderRadius:"50%", background:c }}/>
                    <span style={{ color:c, fontWeight:600, fontSize:10 }}>{l}</span>
                    <span style={{ color:"#4a8aaa", fontSize:9 }}>{r}</span>
                  </div>
                ))}
              </div>
              <ResponsiveContainer width="100%" height={290}>
                {view === "monthly" ? (
                  <ComposedChart data={monthlyData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(monthlyData.length)}/>
                    <YAxis yAxisId="l" tick={{ fill:"#4a8aaa", fontSize:9 }} unit="°C" domain={[20,40]}/>
                    <YAxis yAxisId="r" orientation="right" tick={{ fill:"#4a8aaa", fontSize:9 }} unit=" d"/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <Legend wrapperStyle={{ color:"#94a3b8", fontSize:10 }}/>
                    <ReferenceLine yAxisId="l" y={35} stroke="#ef4444" strokeDasharray="4 2" label={{ value:"Extreme ≥35°C", fill:"#ef4444", fontSize:9, position:"insideTopLeft" }}/>
                    <ReferenceLine yAxisId="l" y={32} stroke="#f97316" strokeDasharray="4 2" label={{ value:"Danger ≥32°C", fill:"#f97316", fontSize:9, position:"insideTopLeft" }}/>
                    <ReferenceLine yAxisId="l" y={28} stroke="#facc15" strokeDasharray="4 2" label={{ value:"Caution ≥28°C", fill:"#facc15", fontSize:9, position:"insideTopLeft" }}/>
                    <Area yAxisId="l" type="monotone" dataKey="wbgt" name="Avg WBGT" stroke="#f97316" fill="#f9731622" strokeWidth={2} dot={false} unit="°C"/>
                    <Bar  yAxisId="r" dataKey="dangerDays" name="Danger Days" fill="#ef444444" radius={[3,3,0,0]} unit=" days"/>
                  </ComposedChart>
                ) : (
                  <AreaChart data={chartData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <defs><linearGradient id="gW" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#f97316" stopOpacity={0.35}/><stop offset="95%" stopColor="#f97316" stopOpacity={0}/></linearGradient></defs>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(filteredRows.length)}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit="°C" domain={[18,42]}/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <ReferenceLine y={35} stroke="#ef4444" strokeDasharray="4 2"/>
                    <ReferenceLine y={32} stroke="#f97316" strokeDasharray="4 2"/>
                    <ReferenceLine y={28} stroke="#facc15" strokeDasharray="4 2"/>
                    <Area type="monotone" dataKey="wbgt" name="WBGT" stroke="#f97316" fill="url(#gW)" strokeWidth={1.5} dot={false} unit="°C"/>
                  </AreaChart>
                )}
              </ResponsiveContainer>
            </div>
          )}

          {/* ════════════════════════════════════════════════════════════════ */}
          {/* TAB: WET BULB TEMP                                              */}
          {/* ════════════════════════════════════════════════════════════════ */}
          {tab === "wetbulb" && (
            <div style={box}>
              <div style={{ paddingLeft:6, marginBottom:10 }}>
                <div style={{ fontSize:9, color:"#38bdf8", letterSpacing:3, textTransform:"uppercase" }}>Wet Bulb Temperature (Tw) · Stull (2011)</div>
                <div style={{ fontSize:14, fontWeight:700, color:"#e0f4ff", marginTop:2 }}>{view==="monthly" ? "Monthly Average Tw" : "Daily Tw"} — Survival limit = 35°C</div>
              </div>
              <ResponsiveContainer width="100%" height={270}>
                <AreaChart data={chartData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                  <defs><linearGradient id="gT" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#38bdf8" stopOpacity={0.35}/><stop offset="95%" stopColor="#38bdf8" stopOpacity={0}/></linearGradient></defs>
                  <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                  <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(chartData.length)}/>
                  <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit="°C" domain={[10,38]}/>
                  <Tooltip content={<CustomTooltip/>}/>
                  <ReferenceLine y={35} stroke="#ef4444" strokeWidth={2} strokeDasharray="5 3" label={{ value:"☠ Lethal 35°C — Raymond et al. 2020", fill:"#ef4444", fontSize:9, position:"insideTopLeft" }}/>
                  <ReferenceLine y={32} stroke="#f97316" strokeDasharray="4 2" label={{ value:"Dangerous 32°C", fill:"#f97316", fontSize:9, position:"insideTopRight" }}/>
                  <ReferenceLine y={28} stroke="#facc15" strokeDasharray="4 2"/>
                  <Area type="monotone" dataKey="wetBulb" name="Wet Bulb Temp" stroke="#38bdf8" fill="url(#gT)" strokeWidth={2} dot={false} unit="°C"/>
                </AreaChart>
              </ResponsiveContainer>
            </div>
          )}

          {/* ════════════════════════════════════════════════════════════════ */}
          {/* TAB: YEAR TREND                                                 */}
          {/* ════════════════════════════════════════════════════════════════ */}
          {tab === "trend" && (
            <>
              <div style={box}>
                <div style={{ fontSize:9, color:"#a78bfa", letterSpacing:3, textTransform:"uppercase", marginBottom:10 }}>Risk Days Per Year — Stacked</div>
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart data={yearlyData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="year" tick={{ fill:"#4a8aaa", fontSize:11 }}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit=" d"/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <Legend wrapperStyle={{ color:"#94a3b8", fontSize:10 }}/>
                    <Bar dataKey="cautionDays" name="Caution 28–32°C" stackId="a" fill="#facc15" unit=" days"/>
                    <Bar dataKey="dangerDays"  name="Danger 32–35°C"  stackId="a" fill="#f97316" unit=" days"/>
                    <Bar dataKey="extremeDays" name="Extreme ≥35°C"   stackId="a" fill="#ef4444" radius={[4,4,0,0]} unit=" days"/>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div style={box}>
                <div style={{ fontSize:9, color:"#a78bfa", letterSpacing:3, textTransform:"uppercase", marginBottom:10 }}>Annual Avg Tw & WBGT — Trend Line</div>
                <ResponsiveContainer width="100%" height={210}>
                  <LineChart data={yearlyData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="year" tick={{ fill:"#4a8aaa", fontSize:11 }}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit="°C" domain={["auto","auto"]}/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <Legend wrapperStyle={{ color:"#94a3b8", fontSize:10 }}/>
                    <ReferenceLine y={35} stroke="#ef4444" strokeDasharray="4 2" label={{ value:"Lethal 35°C", fill:"#ef4444", fontSize:9 }}/>
                    <Line type="monotone" dataKey="avgTw"   name="Avg Wet Bulb" stroke="#38bdf8" strokeWidth={2.5} dot={{ fill:"#38bdf8", r:5 }} unit="°C"/>
                    <Line type="monotone" dataKey="avgWbgt" name="Avg WBGT"     stroke="#f97316" strokeWidth={2}   dot={{ fill:"#f97316", r:4 }} unit="°C" strokeDasharray="5 3"/>
                    <Line type="monotone" dataKey="avgTemp" name="Avg Air Temp" stroke="#facc15" strokeWidth={1.5} dot={{ fill:"#facc15", r:3 }} unit="°C" strokeDasharray="3 2"/>
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </>
          )}

          {/* ════════════════════════════════════════════════════════════════ */}
          {/* TAB: TEMPERATURE                                                */}
          {/* ════════════════════════════════════════════════════════════════ */}
          {tab === "temp" && (
            <div style={box}>
              <div style={{ fontSize:9, color:"#fb923c", letterSpacing:3, textTransform:"uppercase", marginBottom:8 }}>Surface Temperature</div>
              <div style={{ fontSize:14, fontWeight:700, color:"#e0f4ff", marginBottom:10 }}>{view==="monthly" ? "Monthly Averages" : "Daily Records"}</div>
              <ResponsiveContainer width="100%" height={260}>
                {view === "monthly" ? (
                  <AreaChart data={monthlyData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <defs>
                      <linearGradient id="gMx" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#f97316" stopOpacity={0.3}/><stop offset="95%" stopColor="#f97316" stopOpacity={0}/></linearGradient>
                      <linearGradient id="gMn" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#38bdf8" stopOpacity={0.2}/><stop offset="95%" stopColor="#38bdf8" stopOpacity={0}/></linearGradient>
                    </defs>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(monthlyData.length)}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit="°C" domain={[10,42]}/>
                    <Tooltip content={<CustomTooltip/>}/><Legend wrapperStyle={{ color:"#94a3b8", fontSize:10 }}/>
                    <Area type="monotone" dataKey="maxTemp"  name="Max Temp"  stroke="#f97316" fill="url(#gMx)" strokeWidth={2} dot={false} unit="°C"/>
                    <Area type="monotone" dataKey="meanTemp" name="Mean Temp" stroke="#facc15" fill="none"       strokeWidth={1.5} dot={false} unit="°C" strokeDasharray="4 2"/>
                    <Area type="monotone" dataKey="minTemp"  name="Min Temp"  stroke="#38bdf8" fill="url(#gMn)" strokeWidth={2} dot={false} unit="°C"/>
                  </AreaChart>
                ) : (
                  <ComposedChart data={chartData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(filteredRows.length)}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit="°C" domain={[10,45]}/>
                    <Tooltip content={<CustomTooltip/>}/><Legend wrapperStyle={{ color:"#94a3b8", fontSize:10 }}/>
                    <Area type="monotone" dataKey="maxTemp" name="Max Temp" stroke="#f97316" fill="#f9731615" strokeWidth={1.5} dot={false} unit="°C"/>
                    <Area type="monotone" dataKey="minTemp" name="Min Temp" stroke="#38bdf8" fill="#38bdf815" strokeWidth={1.5} dot={false} unit="°C"/>
                  </ComposedChart>
                )}
              </ResponsiveContainer>
            </div>
          )}

          {/* ════════════════════════════════════════════════════════════════ */}
          {/* TAB: HUMIDITY                                                   */}
          {/* ════════════════════════════════════════════════════════════════ */}
          {tab === "humidity" && (
            <div style={box}>
              <div style={{ fontSize:9, color:"#38bdf8", letterSpacing:3, textTransform:"uppercase", marginBottom:8 }}>Relative Humidity</div>
              <div style={{ fontSize:14, fontWeight:700, color:"#e0f4ff", marginBottom:10 }}>{view==="monthly" ? "Monthly Averages" : "Daily Records"}</div>
              <ResponsiveContainer width="100%" height={250}>
                {view === "monthly" ? (
                  <BarChart data={monthlyData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(monthlyData.length)}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit="%" domain={[40,100]}/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <ReferenceLine y={80} stroke="#7dd3fc" strokeDasharray="4 2" label={{ value:"Monsoon ~80%", fill:"#7dd3fc", fontSize:9, position:"insideTopLeft" }}/>
                    <Bar dataKey="humidity" name="Avg Humidity" fill="#0ea5e9" radius={[4,4,0,0]} unit="%"/>
                  </BarChart>
                ) : (
                  <AreaChart data={chartData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <defs><linearGradient id="gHm" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.4}/><stop offset="95%" stopColor="#0ea5e9" stopOpacity={0.02}/></linearGradient></defs>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(filteredRows.length)}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit="%" domain={[30,100]}/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <Area type="monotone" dataKey="humidity" name="Humidity" stroke="#0ea5e9" fill="url(#gHm)" strokeWidth={1.5} dot={false} unit="%"/>
                  </AreaChart>
                )}
              </ResponsiveContainer>
            </div>
          )}

          {/* ════════════════════════════════════════════════════════════════ */}
          {/* TAB: SOLAR                                                      */}
          {/* ════════════════════════════════════════════════════════════════ */}
          {tab === "solar" && (
            <div style={box}>
              <div style={{ fontSize:9, color:"#fbbf24", letterSpacing:3, textTransform:"uppercase", marginBottom:8 }}>Solar Radiation (MJ/m²)</div>
              <div style={{ fontSize:14, fontWeight:700, color:"#e0f4ff", marginBottom:10 }}>{view==="monthly" ? "Monthly Avg Daily Radiation" : "Daily Shortwave Radiation"}</div>
              <ResponsiveContainer width="100%" height={240}>
                {view === "monthly" ? (
                  <BarChart data={monthlyData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(monthlyData.length)}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit=" MJ" domain={[0,30]}/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <Bar dataKey="solar" name="Solar Radiation" fill="#f59e0b" radius={[4,4,0,0]} unit=" MJ/m²"/>
                  </BarChart>
                ) : (
                  <AreaChart data={chartData} margin={{ top:5,right:14,left:0,bottom:5 }}>
                    <defs><linearGradient id="gSl" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#f59e0b" stopOpacity={0.4}/><stop offset="95%" stopColor="#f59e0b" stopOpacity={0.02}/></linearGradient></defs>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{ fill:"#4a8aaa", fontSize:9 }} interval={iv(filteredRows.length)}/>
                    <YAxis tick={{ fill:"#4a8aaa", fontSize:9 }} unit=" MJ" domain={[0,35]}/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <Area type="monotone" dataKey="solar" name="Solar Radiation" stroke="#f59e0b" fill="url(#gSl)" strokeWidth={1.5} dot={false} unit=" MJ/m²"/>
                  </AreaChart>
                )}
              </ResponsiveContainer>
            </div>
          )}

          {/* ── CITATIONS ── */}
          <div style={{ ...box, background:"rgba(139,92,246,0.04)", border:"1px solid #3b1f6a" }}>
            <div style={{ fontSize:9, color:"#a78bfa", letterSpacing:3, textTransform:"uppercase", marginBottom:10 }}>📚 Paper Citations</div>
            <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
              {[
                ["ERA5 Data",       "Hersbach et al. (2020). The ERA5 global reanalysis. Q. J. R. Meteorol. Soc., 146(730). doi:10.1002/qj.3803"],
                ["Wet Bulb Formula","Stull, R. (2011). Wet-Bulb Temperature from RH and Air Temperature. J. Appl. Meteor. Climatol., 50, 2267–2269."],
                ["35°C Limit",      "Raymond et al. (2020). Heat and humidity too severe for human tolerance. Science Advances, 6(19). doi:10.1126/sciadv.aaw1838"],
                ["WBGT Formula",    "Bernard & Pourmoghani (1999). Prediction of workplace WBGT. Appl. Occup. Environ. Hyg., 14(2), 126–134."],
                ["Mann-Kendall",    "Mann, H.B. (1945). Nonparametric tests against trend. Econometrica, 13, 245–259. / Kendall, M.G. (1975). Rank Correlation Methods. Griffin, London."],
                ["Sen's Slope",     "Sen, P.K. (1968). Estimates of regression coefficient based on Kendall's tau. JASA, 63(324), 1379–1389."],
              ].map(([ref,cite],i)=>(
                <div key={i} style={{ fontSize:10, color:"#94a3b8", lineHeight:1.7 }}>
                  <span style={{ color:"#a78bfa", fontWeight:600 }}>[{ref}]</span> {cite}
                </div>
              ))}
            </div>
          </div>

          <div style={{ textAlign:"center", fontSize:9, color:"#1e3a4a", marginTop:4 }}>
            ERA5 · ECMWF · Open-Meteo · {KHULNA_LAT}°N {KHULNA_LON}°E · Mann-Kendall + Sen's Slope computed client-side
          </div>

        </>)}
      </div>
    </div>
  );
}