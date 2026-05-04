import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, ReferenceLine, Area, AreaChart, ComposedChart,
  Line, LineChart, Scatter, ScatterChart, Cell,
} from "recharts";

const KHULNA_LAT = 22.8098;
const KHULNA_LON = 89.5644;
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// ─── The earliest ERA5 has reliable data for Bangladesh ───────────────────────
const ERA5_EARLIEST = "2000-01-01";

// ─── PHYSICS ──────────────────────────────────────────────────────────────────
function calcWetBulb(T, RH) {
  if (T == null || RH == null) return null;
  const Tw = T * Math.atan(0.151977 * Math.pow(RH + 8.313659, 0.5))
    + Math.atan(T + RH) - Math.atan(RH - 1.676331)
    + 0.00391838 * Math.pow(RH, 1.5) * Math.atan(0.023101 * RH) - 4.686035;
  return parseFloat(Tw.toFixed(2));
}
function calcWBGT(T, RH) {
  if (T == null || RH == null) return null;
  const es = 6.112 * Math.exp((17.67 * T) / (T + 243.5));
  const e  = (RH / 100) * es;
  return parseFloat((0.567 * T + 0.393 * e + 3.94).toFixed(2));
}
function getWBGTRisk(wbgt) {
  if (wbgt == null) return { level:"Unknown",  color:"#6b7280" };
  if (wbgt < 28)    return { level:"Safe",      color:"#4ade80" };
  if (wbgt < 32)    return { level:"Caution",   color:"#facc15" };
  if (wbgt < 35)    return { level:"Danger",    color:"#f97316" };
  return              { level:"Extreme",    color:"#ef4444" };
}

// ─── STATISTICS ───────────────────────────────────────────────────────────────
function mannKendall(series) {
  const n = series.length;
  if (n < 4) return null;
  let S = 0;
  for (let i = 0; i < n - 1; i++)
    for (let j = i + 1; j < n; j++) {
      const d = series[j] - series[i];
      if (d > 0) S++; else if (d < 0) S--;
    }
  const varS = (n * (n - 1) * (2 * n + 5)) / 18;
  let Z = 0;
  if (S > 0)      Z = (S - 1) / Math.sqrt(varS);
  else if (S < 0) Z = (S + 1) / Math.sqrt(varS);
  const pValue = 2 * (1 - normalCDF(Math.abs(Z)));
  let trend = "No significant trend";
  if (pValue < 0.01 && Z > 0) trend = "Significant increasing trend (p<0.01)";
  else if (pValue < 0.05 && Z > 0) trend = "Significant increasing trend (p<0.05)";
  else if (pValue < 0.1  && Z > 0) trend = "Marginally increasing trend (p<0.10)";
  else if (pValue < 0.01 && Z < 0) trend = "Significant decreasing trend (p<0.01)";
  else if (pValue < 0.05 && Z < 0) trend = "Significant decreasing trend (p<0.05)";
  return { S, Z: parseFloat(Z.toFixed(3)), pValue: parseFloat(pValue.toFixed(4)), trend };
}
function senSlope(series) {
  const slopes = [];
  for (let i = 0; i < series.length - 1; i++)
    for (let j = i + 1; j < series.length; j++)
      slopes.push((series[j] - series[i]) / (j - i));
  slopes.sort((a, b) => a - b);
  const mid = Math.floor(slopes.length / 2);
  return parseFloat((slopes.length % 2 === 0 ? (slopes[mid-1]+slopes[mid])/2 : slopes[mid]).toFixed(4));
}
function normalCDF(z) {
  const t = 1 / (1 + 0.2316419 * Math.abs(z));
  const poly = t*(0.319381530+t*(-0.356563782+t*(1.781477937+t*(-1.821255978+t*1.330274429))));
  const pdf  = Math.exp(-0.5*z*z)/Math.sqrt(2*Math.PI);
  const cdf  = 1 - pdf*poly;
  return z >= 0 ? cdf : 1 - cdf;
}
function avg(arr) {
  const v = arr.filter(x => x != null);
  if (!v.length) return null;
  return parseFloat((v.reduce((a,b)=>a+b,0)/v.length).toFixed(2));
}
function pearsonR(x, y) {
  const n = Math.min(x.length, y.length);
  if (n < 3) return null;
  const mx = x.reduce((a,b)=>a+b,0)/n, my = y.reduce((a,b)=>a+b,0)/n;
  let num=0, dx2=0, dy2=0;
  for (let i=0;i<n;i++) { num+=(x[i]-mx)*(y[i]-my); dx2+=(x[i]-mx)**2; dy2+=(y[i]-my)**2; }
  if (dx2===0||dy2===0) return null;
  return parseFloat((num/Math.sqrt(dx2*dy2)).toFixed(3));
}
function linearRegression(x, y) {
  const n = x.length;
  if (n < 3) return null;
  const mx = x.reduce((a,b)=>a+b,0)/n, my = y.reduce((a,b)=>a+b,0)/n;
  let ssxy=0, ssxx=0;
  for (let i=0;i<n;i++) { ssxy+=(x[i]-mx)*(y[i]-my); ssxx+=(x[i]-mx)**2; }
  const slope=ssxy/ssxx, intercept=my-slope*mx;
  const yPred=x.map(xi=>slope*xi+intercept);
  const ssTot=y.reduce((a,yi)=>a+(yi-my)**2,0);
  const ssRes=y.reduce((a,yi,i)=>a+(yi-yPred[i])**2,0);
  const r2=1-ssRes/ssTot;
  const F=(r2/1)/((1-r2)/(n-2));
  const pValue=1-fCDF(F,1,n-2);
  return { slope:parseFloat(slope.toFixed(4)), intercept:parseFloat(intercept.toFixed(4)), r2:parseFloat(r2.toFixed(4)), pValue:parseFloat(pValue.toFixed(4)), yPred };
}
function fCDF(F,d1,d2){ const x=d2/(d2+d1*F); return incompleteBeta(x,d2/2,d1/2); }
function incompleteBeta(x,a,b){
  if(x<=0)return 0; if(x>=1)return 1;
  const lbeta=lgamma(a)+lgamma(b)-lgamma(a+b);
  const front=Math.exp(Math.log(x)*a+Math.log(1-x)*b-lbeta)/a;
  return front*betaCF(x,a,b);
}
function betaCF(x,a,b){
  const MAXIT=200,EPS=3e-7;
  let qab=a+b,qap=a+1,qam=a-1,c=1,d=1-qab*x/qap;
  if(Math.abs(d)<1e-30)d=1e-30; d=1/d; let h=d;
  for(let m=1;m<=MAXIT;m++){
    let m2=2*m,aa=m*(b-m)*x/((qam+m2)*(a+m2));
    d=1+aa*d; if(Math.abs(d)<1e-30)d=1e-30; c=1+aa/c; if(Math.abs(c)<1e-30)c=1e-30; d=1/d; h*=d*c;
    aa=-(a+m)*(qab+m)*x/((a+m2)*(qap+m2));
    d=1+aa*d; if(Math.abs(d)<1e-30)d=1e-30; c=1+aa/c; if(Math.abs(c)<1e-30)c=1e-30; d=1/d;
    const del=d*c; h*=del; if(Math.abs(del-1)<EPS)break;
  }
  return h;
}
function lgamma(x){
  const c=[76.18009172947146,-86.50532032941677,24.01409824083091,-1.231739572450155,0.1208650973866179e-2,-0.5395239384953e-5];
  let y=x,tmp=x+5.5; tmp-=(x+0.5)*Math.log(tmp); let ser=1.000000000190015;
  for(let i=0;i<6;i++) ser+=c[i]/++y;
  return -tmp+Math.log(2.5066282746310005*ser/x);
}

// ─── SEASON ───────────────────────────────────────────────────────────────────
function getSeason(month) {
  const m = parseInt(month);
  if (m >= 3 && m <= 5)  return "Pre-monsoon";
  if (m >= 6 && m <= 9)  return "Monsoon";
  return "Post-monsoon";
}

// ─── GROUPING ─────────────────────────────────────────────────────────────────
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
  return Object.entries(m).map(([k,v])=>({
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
    if (!y[yr]) y[yr] = { tw:[],wbgt:[],temp:[],hum:[],caution:0,danger:0,extreme:0,total:0 };
    if (r.wetBulb  != null) y[yr].tw.push(r.wetBulb);
    if (r.wbgt     != null) y[yr].wbgt.push(r.wbgt);
    if (r.meanTemp != null) y[yr].temp.push(r.meanTemp);
    if (r.humidity != null) y[yr].hum.push(r.humidity);
    y[yr].total++;
    if (r.wbgt >= 28 && r.wbgt < 32) y[yr].caution++;
    if (r.wbgt >= 32 && r.wbgt < 35) y[yr].danger++;
    if (r.wbgt >= 35) y[yr].extreme++;
  });
  return Object.entries(y).sort().map(([yr,v])=>({
    year:yr, yearNum:parseInt(yr),
    avgTw:avg(v.tw), avgWbgt:avg(v.wbgt), avgTemp:avg(v.temp), avgHum:avg(v.hum),
    cautionDays:v.caution, dangerDays:v.danger, extremeDays:v.extreme, totalDays:v.total,
  }));
}

// ─── UI HELPERS ───────────────────────────────────────────────────────────────
// Note: CT tooltip is defined inside main component to access theme tokens (th)
const iStyle={background:"rgba(255,255,255,0.06)",border:"1px solid #1e4d6b",borderRadius:8,color:"#e0f4ff",padding:"7px 11px",fontSize:12,fontFamily:"Georgia,serif",outline:"none",colorScheme:"dark"};
const box={background:"rgba(255,255,255,0.025)",border:"1px solid #102d4a",borderRadius:16,padding:"18px 14px 14px",marginBottom:16};
const iv=n=>Math.max(0,Math.floor(n/14)-1);
function pBadge(p){
  if(p<0.01)return{label:"p < 0.01 ✓✓",color:"#4ade80"};
  if(p<0.05)return{label:"p < 0.05 ✓",color:"#a3e635"};
  if(p<0.10)return{label:"p < 0.10 ~",color:"#facc15"};
  return{label:`p = ${p} ✗`,color:"#f87171"};
}
function corrColor(r){
  if(r===null)return"#1a1a2e";
  const abs=Math.abs(r);
  if(abs>=0.8)return r>0?"rgba(239,68,68,0.7)":"rgba(56,189,248,0.7)";
  if(abs>=0.6)return r>0?"rgba(249,115,22,0.6)":"rgba(14,165,233,0.6)";
  if(abs>=0.4)return r>0?"rgba(250,204,21,0.4)":"rgba(99,202,183,0.4)";
  if(abs>=0.2)return r>0?"rgba(250,204,21,0.2)":"rgba(99,202,183,0.2)";
  return"rgba(255,255,255,0.04)";
}

function MKCard({title,mk,slope,unit,color}){
  if(!mk)return null;
  const pb=pBadge(mk.pValue);
  return(
    <div style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${color}44`,borderRadius:14,padding:"16px 18px",flex:"1 1 280px"}}>
      <div style={{fontSize:10,color,letterSpacing:3,textTransform:"uppercase",marginBottom:8}}>Mann-Kendall · {title}</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        {[
          {label:"Kendall S",value:mk.S},
          {label:"Z Statistic",value:mk.Z},
          {label:"p-value",value:<span style={{color:pb.color,fontWeight:700}}>{pb.label}</span>},
          {label:"Sen's Slope",value:<span style={{color:slope>0?"#f87171":"#4ade80",fontWeight:700}}>{slope>0?"+":""}{slope} {unit}/yr</span>},
        ].map((s,i)=>(
          <div key={i} style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:"8px 10px"}}>
            <div style={{fontSize:10,color:"#4a8aaa",marginBottom:3}}>{s.label}</div>
            <div style={{fontSize:15,fontWeight:700,color:"#e0f4ff"}}>{s.value}</div>
          </div>
        ))}
      </div>
      <div style={{fontSize:12,color:mk.Z>0?"#fca5a5":"#86efac",fontStyle:"italic",borderTop:"1px solid #0f2a3f",paddingTop:8}}>
        {mk.Z>0?"▲":"▼"} {mk.trend}
      </div>
    </div>
  );
}

// ─── 5-YEAR PRESET RANGES ─────────────────────────────────────────────────────
const PRESET_RANGES = [
  { label:"2000–2025 (25 yrs — recommended)", start:"2000-01-01", endFn: t => t },
  { label:"2005–2025 (20 yrs)", start:"2005-01-01", endFn: t => t },
  { label:"2010–2025 (15 yrs)", start:"2010-01-01", endFn: t => t },
  { label:"2015–2025 (10 yrs)", start:"2015-01-01", endFn: t => t },
  { label:"2019–2025 (6 yrs)", start:"2019-01-01", endFn: t => t },
];

// ─── FIGURE DOWNLOAD UTILITY ─────────────────────────────────────────────────

// Find the best Recharts SVG inside a container.
// Uses multiple strategies so it works regardless of Recharts version / iframe.
function findChartSvg(container) {
  if (!container) return null;

  // Strategy 1: Recharts v2 standard class
  const byClass = container.querySelector('.recharts-surface')
                || container.querySelector('.recharts-wrapper svg');
  if (byClass) {
    const r = byClass.getBoundingClientRect();
    if (r.width > 50 && r.height > 50) return byClass;
    // Class found but zero-size — check SVG attributes
    const w = parseFloat(byClass.getAttribute('width') || '0');
    const h = parseFloat(byClass.getAttribute('height') || '0');
    if (w > 50 && h > 50) return byClass;
  }

  // Strategy 2: pick the SVG with the largest explicit width attribute
  const allSvgs = [...container.querySelectorAll('svg')];
  if (!allSvgs.length) return null;

  let best = null, bestW = 0;
  for (const svg of allSvgs) {
    // Check both rendered size and SVG attribute width
    const rect = svg.getBoundingClientRect();
    const attrW = parseFloat(svg.getAttribute('width') || '0');
    const w = Math.max(rect.width, attrW);
    if (w > bestW) { bestW = w; best = svg; }
  }
  return (bestW > 50) ? best : null;
}

// Retry finding the chart SVG up to maxAttempts times with delay between tries.
// This handles the ResponsiveContainer two-pass render without using RAF
// (which gets throttled in iframes).
async function findChartSvgWithRetry(containerRef, maxAttempts = 25, delayMs = 200) {
  for (let i = 0; i < maxAttempts; i++) {
    const svg = findChartSvg(containerRef.current);
    if (svg) {
      // Verify it has real dimensions
      const rect  = svg.getBoundingClientRect();
      const attrW = parseFloat(svg.getAttribute('width') || '0');
      const w     = Math.max(rect.width, attrW);
      const attrH = parseFloat(svg.getAttribute('height') || '0');
      const h     = Math.max(rect.height, attrH);
      if (w >= 100 && h >= 50) return { svg, w, h };
    }
    await new Promise(r => setTimeout(r, delayMs));
  }
  throw new Error(`Chart not ready after ${maxAttempts} attempts — try the individual Download Figure button instead`);
}

async function renderFigToPng(containerRef, figLabel) {
  if (!containerRef.current) throw new Error('Chart container not found — make sure you are on the correct tab');

  const { svg: svgEl, w: W, h: H } = await findChartSvgWithRetry(containerRef, 25, 200);
  const SCALE = 3; // 3× → ~300 dpi

  const clone = svgEl.cloneNode(true);
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  clone.setAttribute('width',  W);
  clone.setAttribute('height', H);
  const style = document.createElementNS('http://www.w3.org/2000/svg', 'style');
  style.textContent = `text, tspan { font-family: Georgia, serif; } .recharts-text { font-size: 11px; }`;
  clone.insertBefore(style, clone.firstChild);

  const svgStr  = new XMLSerializer().serializeToString(clone);
  const svgBlob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
  const svgUrl  = URL.createObjectURL(svgBlob);

  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width  = Math.round(W * SCALE);
      canvas.height = Math.round(H * SCALE);
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(SCALE, SCALE);
      ctx.drawImage(img, 0, 0, W, H);
      ctx.font = 'bold 9px Georgia, serif';
      ctx.fillStyle = '#777777';
      ctx.fillText(`Khulna Heat Stress Study · ${figLabel}`, 8, H - 6);
      URL.revokeObjectURL(svgUrl);
      canvas.toBlob(b => b ? resolve(b) : reject(new Error('Canvas export failed')), 'image/png', 1.0);
    };
    img.onerror = () => { URL.revokeObjectURL(svgUrl); reject(new Error('SVG image load failed')); };
    img.src = svgUrl;
  });
}


// ─── SINGLE FIGURE DOWNLOAD (Save As dialog) ──────────────────────────────────
async function downloadFig(containerRef, filename, figLabel) {
  const pngBlob = await renderFigToPng(containerRef, figLabel);

  if (typeof window.showSaveFilePicker === 'function') {
    try {
      const fh = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [{ description: 'PNG Image (300 dpi)', accept: { 'image/png': ['.png'] } }],
      });
      const w = await fh.createWritable();
      await w.write(pngBlob);
      await w.close();
      return;
    } catch (err) {
      if (err.name === 'AbortError') return;
    }
  }
  // Fallback
  const url  = URL.createObjectURL(pngBlob);
  const link = document.createElement('a');
  link.href = url; link.download = filename;
  document.body.appendChild(link); link.click(); document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// ─── ALL FIGURES DOWNLOAD (folder picker or sequential) ───────────────────────


// ─── DOWNLOAD BUTTON COMPONENT ────────────────────────────────────────────────
function DownloadBtn({ containerRef, filename, figLabel, figNum }) {
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);

  const handle = async () => {
    setBusy(true);
    try {
      await downloadFig(containerRef, filename, figLabel);
      setDone(true);
      setTimeout(() => setDone(false), 2500);
    } catch(e) {
      alert('Download failed: ' + e.message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between',
                  marginBottom:10, padding:'8px 12px', borderRadius:8,
                  background:'rgba(56,189,248,0.06)', border:'1px solid #0c4a6e' }}>
      <div>
        <span style={{ fontSize:11, color:'#7dd3fc', fontWeight:700 }}>
          {figNum}
        </span>
        <span style={{ fontSize:11, color:'#64a7c8', marginLeft:8 }}>
          {figLabel}
        </span>
        <span style={{ fontSize:10, color:'#4a6a7a', marginLeft:8 }}>
          · PNG 3× resolution (300 dpi) · white background · print-ready
        </span>
      </div>
      <button
        onClick={handle}
        disabled={busy}
        style={{ display:'flex', alignItems:'center', gap:6, padding:'6px 14px',
                 borderRadius:7, border:'1px solid #38bdf8',
                 background: done ? '#064e3b' : '#0c3a5e',
                 color: done ? '#34d399' : '#7dd3fc',
                 cursor: busy ? 'wait' : 'pointer',
                 fontSize:11, fontFamily:'inherit', fontWeight:600,
                 whiteSpace:'nowrap', transition:'all 0.2s' }}>
        {busy ? '⏳ Exporting…' : done ? '✓ Saved!' : '📥 Download Figure'}
      </button>
    </div>
  );
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
export default function KhulnaHeatStress() {
  const todayStr    = new Date().toISOString().split("T")[0];
  const currentYear = new Date().getFullYear();

  // ── fetch range (drives the API call) ──
  const [fetchStart, setFetchStart] = useState("2000-01-01");
  const [fetchEnd,   setFetchEnd]   = useState(todayStr);

  // ── filter range (subset of fetched data for analysis) ──
  const [startDate, setStartDate] = useState("2000-01-01");
  const [endDate,   setEndDate]   = useState(todayStr);

  // ── quick select state ──
  const [presetIdx,    setPresetIdx]    = useState(0);   // dropdown index
  const [customYear,   setCustomYear]   = useState("");  // custom single year field
  const [customYearErr,setCustomYearErr]= useState("");

  const [rawData,     setRawData]     = useState(null);
  const [loading,     setLoading]     = useState(false);
  const [error,       setError]       = useState(null);
  const [downloading, setDownloading] = useState(false);
  const [tab,         setTab]         = useState("mk");
  const [view,        setView]        = useState("monthly");

  // ── Figure refs for paper export ──
  const refFig2 = useRef(null);
  const refFig3 = useRef(null);
  const refFig4 = useRef(null);
  const refFig5 = useRef(null);
  const refFig6 = useRef(null);
  const refFig7 = useRef(null);
  const refFig8 = useRef(null);

  // ── Paper mode — light theme for journal-quality export ──
  const [paperMode,    setPaperMode]    = useState(false);
  const [dlAll,        setDlAll]        = useState({ busy:false, current:0, total:0, filename:'' });

  // Convenience: box style that adapts to paper mode
  const pbox = { ...box,
    background: paperMode ? '#f9fafb' : 'rgba(255,255,255,0.025)',
    border: paperMode ? '1px solid #e5e7eb' : '1px solid #102d4a',
  };

  // Theme tokens — dark (app) vs white (paper/print)
  const th = paperMode ? {
    bg:        '#ffffff',
    grid:      '#e5e7eb',
    gridDash:  '3 3',
    axTick:    '#374151',
    text:      '#111827',
    subtext:   '#6b7280',
    barWbgt:   '#c2410c',
    barTw:     '#0369a1',
    lineWbgt:  '#c2410c',
    lineTw:    '#0369a1',
    lineTemp:  '#ca8a04',
    areaWbgt:  '#c2410c',
    areaFill:  '#fde8d8',
    refExtreme:'#b91c1c',
    refDanger: '#c2410c',
    refCaution:'#a16207',
    refLethal: '#b91c1c',
    caution:   '#fbbf24',
    danger:    '#f97316',
    extreme:   '#ef4444',
    boxBg:     '#f9fafb',
    boxBorder: '#e5e7eb',
    ttBg:      'rgba(255,255,255,0.98)',
    ttBorder:  '#d1d5db',
    ttText:    '#111827',
    ttLabel:   '#1d4ed8',
  } : {
    bg:        'transparent',
    grid:      '#0a1a28',
    gridDash:  '3 3',
    axTick:    '#4a8aaa',
    text:      '#e0f4ff',
    subtext:   '#94a3b8',
    barWbgt:   '#f97316',
    barTw:     '#38bdf8',
    lineWbgt:  '#f97316',
    lineTw:    '#38bdf8',
    lineTemp:  '#facc15',
    areaWbgt:  '#f97316',
    areaFill:  '#f9731622',
    refExtreme:'#ef4444',
    refDanger: '#f97316',
    refCaution:'#facc15',
    refLethal: '#ef4444',
    caution:   '#facc15',
    danger:    '#f97316',
    extreme:   '#ef4444',
    boxBg:     'rgba(255,255,255,0.025)',
    boxBorder: '#102d4a',
    ttBg:      'rgba(4,8,15,0.98)',
    ttBorder:  '#1e3a50',
    ttText:    '#e0f4ff',
    ttLabel:   '#7dd3fc',
  };

  // Theme-aware tooltip — defined after th so it can reference th tokens
  const CT = ({ active, payload, label }) => {
    if (!active||!payload?.length) return null;
    return (
      <div style={{background:th.ttBg,border:`1px solid ${th.ttBorder}`,borderRadius:10,padding:"10px 15px",fontSize:12,color:th.ttText}}>
        <div style={{fontWeight:700,color:th.ttLabel,marginBottom:5}}>{label}</div>
        {payload.map((p,i)=><div key={i} style={{color:p.color,marginBottom:2}}>{p.name}: <b>{p.value!=null?p.value:"—"}{p.unit||""}</b></div>)}
      </div>
    );
  };

  // ─── FIGURE LIST — tab field tells the downloader which tab to activate ──────
  const figList = [
    { ref:refFig2, filename:'Fig2_Annual_Trend_Lines.png',   label:'Annual Tw, WBGT & Air Temp trend lines', num:'Fig. 2', tab:'trend',      view:null      },
    { ref:refFig3, filename:'Fig3_Annual_Danger_Days.png',   label:'Annual stacked heat stress risk days',   num:'Fig. 3', tab:'trend',      view:null      },
    { ref:refFig4, filename:'Fig4_Seasonal_WBGT.png',        label:'Seasonal WBGT comparison bar chart',     num:'Fig. 4', tab:'seasonal',   view:null      },
    { ref:refFig5, filename:'Fig5_Exceedance_Frequency.png', label:'Exceedance frequency bar chart',         num:'Fig. 5', tab:'exceedance', view:null      },
    { ref:refFig6, filename:'Fig6_OLS_Regression.png',       label:'OLS regression actual vs trend line',    num:'Fig. 6', tab:'regression', view:null      },
    { ref:refFig7, filename:'Fig7_Monthly_WBGT_Risk.png',    label:'Monthly WBGT heat stroke risk',          num:'Fig. 7', tab:'heatstroke', view:'monthly' },
    { ref:refFig8, filename:'Fig8_Wet_Bulb_Temperature.png', label:'Wet bulb temperature time series',       num:'Fig. 8', tab:'wetbulb',    view:'monthly' },
  ];

  // ─── DOWNLOAD ALL — navigates tabs automatically so every chart is mounted ───
  const handleDownloadAll = async () => {
    if (!paperMode) {
      alert('Please enable 📄 Paper Mode first — the button turns green in the header.');
      return;
    }

    // ── Step 1: pick folder BEFORE navigating tabs ──────────────────────────
    let dirHandle = null;
    if (typeof window.showDirectoryPicker === 'function') {
      try {
        dirHandle = await window.showDirectoryPicker({ mode:'readwrite', startIn:'downloads' });
      } catch (err) {
        if (err.name === 'AbortError') return; // user cancelled — do nothing
        // API exists but failed for another reason — fall through to sequential
      }
    }

    // ── Step 2: save current tab/view so we can restore them after ──────────
    const savedTab  = tab;
    const savedView = view;

    setDlAll({ busy:true, current:0, total:figList.length, filename:'Preparing…' });

    try {
      let lastTab = null;

      for (let i = 0; i < figList.length; i++) {
        const fig = figList[i];
        setDlAll({ busy:true, current:i+1, total:figList.length, filename:fig.filename });

        // ── Navigate to the right tab + view if not already there ────────────
        if (fig.tab !== lastTab) {
          setTab(fig.tab);
          if (fig.view) setView(fig.view);
          lastTab = fig.tab;
          // Wait for React to mount the tab, then poll until chart is ready
          await new Promise(r => setTimeout(r, 150));
          await findChartSvgWithRetry(fig.ref, 25, 200);
          // One extra wait for legend/axes to finish rendering
          await new Promise(r => setTimeout(r, 100));
        } else {
          // Same tab — chart already mounted, brief pause
          await new Promise(r => setTimeout(r, 100));
        }

        // ── Capture the chart ────────────────────────────────────────────────
        let blob;
        try {
          blob = await renderFigToPng(fig.ref, fig.label);
        } catch (e) {
          console.warn(`Could not capture ${fig.filename}:`, e.message);
          continue; // skip this figure, keep going
        }

        // ── Save the blob ────────────────────────────────────────────────────
        if (dirHandle) {
          // Write directly into the chosen folder
          const fh = await dirHandle.getFileHandle(fig.filename, { create:true });
          const w  = await fh.createWritable();
          await w.write(blob);
          await w.close();
        } else {
          // Fallback: trigger individual browser download
          const url  = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url; link.download = fig.filename;
          document.body.appendChild(link); link.click();
          document.body.removeChild(link);
          URL.revokeObjectURL(url);
          await new Promise(r => setTimeout(r, 700)); // space out browser downloads
        }
      }

      // ── Step 3: restore original tab/view ───────────────────────────────
      setTab(savedTab);
      setView(savedView);
      setDlAll({ busy:false, current:figList.length, total:figList.length, filename:'✓ All figures saved!' });
      setTimeout(() => setDlAll({ busy:false, current:0, total:0, filename:'' }), 3000);

    } catch (e) {
      alert('Download failed: ' + e.message);
      setTab(savedTab);
      setView(savedView);
      setDlAll({ busy:false, current:0, total:0, filename:'' });
    }
  };

  // ─── FETCH — re-runs whenever fetchStart or fetchEnd changes ───────────────
  const doFetch = useCallback(async (start, end) => {
    setLoading(true);
    setError(null);
    setRawData(null);
    try {
      const url = `https://archive-api.open-meteo.com/v1/archive?latitude=${KHULNA_LAT}&longitude=${KHULNA_LON}&start_date=${start}&end_date=${end}&daily=temperature_2m_max,temperature_2m_min,temperature_2m_mean,relative_humidity_2m_mean,shortwave_radiation_sum&timezone=Asia%2FDhaka`;
      const json = await (await fetch(url)).json();
      if (json.error) throw new Error(json.reason || "API error");
      setRawData(json);
    } catch(e) { setError("Failed to fetch ERA5 data: " + e.message); }
    finally    { setLoading(false); }
  }, []);

  // initial load
  useEffect(() => { doFetch(fetchStart, fetchEnd); }, []); // eslint-disable-line

  // ─── APPLY PRESET ──────────────────────────────────────────────────────────
  const applyPreset = (idx) => {
    const p = PRESET_RANGES[idx];
    const end = p.endFn(todayStr);
    setPresetIdx(idx);
    setCustomYear("");
    setCustomYearErr("");
    setFetchStart(p.start);
    setFetchEnd(end);
    setStartDate(p.start);
    setEndDate(end);
    doFetch(p.start, end);
  };

  // ─── APPLY CUSTOM YEAR ─────────────────────────────────────────────────────
  const applyCustomYear = () => {
    const y = parseInt(customYear);
    if (!customYear || isNaN(y) || y < 2000 || y > currentYear) {
      setCustomYearErr(`Enter a year between 2000 and ${currentYear}`);
      return;
    }
    setCustomYearErr("");
    const start = `${y}-01-01`;
    const end   = y === currentYear ? todayStr : `${y}-12-31`;
    // For a single year we still fetch from 2000 so MK has data,
    // but filter the display to just that year
    setFetchStart("2000-01-01");
    setFetchEnd(todayStr);
    setStartDate(start);
    setEndDate(end);
    if (!rawData) doFetch("2000-01-01", todayStr);
  };

  // ─── PROCESS ───────────────────────────────────────────────────────────────
  const allDailyRows = useMemo(() => {
    if (!rawData) return [];
    return rawData.daily.time.map((d,i) => {
      const T  = rawData.daily.temperature_2m_mean[i];
      const RH = rawData.daily.relative_humidity_2m_mean[i];
      const tw   = calcWetBulb(T, RH);
      const wbgt = calcWBGT(T, RH);
      return {
        date:d,
        maxTemp:rawData.daily.temperature_2m_max[i],
        minTemp:rawData.daily.temperature_2m_min[i],
        meanTemp:T, humidity:RH,
        solar:rawData.daily.shortwave_radiation_sum[i],
        wetBulb:tw, wbgt,
        wbgtRisk:getWBGTRisk(wbgt).level,
      };
    });
  }, [rawData]);

  const filteredRows = useMemo(()=>
    allDailyRows.filter(r=>r.date>=startDate&&r.date<=endDate),
    [allDailyRows,startDate,endDate]
  );

  const monthlyData = useMemo(()=>groupByMonth(filteredRows),[filteredRows]);
  const yearlyData  = useMemo(()=>groupByYear(filteredRows), [filteredRows]);

  // ── MANN-KENDALL on annual means ──
  const mkStats = useMemo(()=>{
    if(yearlyData.length<4)return null;
    const tw     = yearlyData.map(y=>y.avgTw).filter(Boolean);
    const wbgt   = yearlyData.map(y=>y.avgWbgt).filter(Boolean);
    const temp   = yearlyData.map(y=>y.avgTemp).filter(Boolean);
    const danger = yearlyData.map(y=>y.dangerDays);
    return{
      tw:    {mk:mannKendall(tw),    slope:senSlope(tw)    },
      wbgt:  {mk:mannKendall(wbgt),  slope:senSlope(wbgt)  },
      temp:  {mk:mannKendall(temp),  slope:senSlope(temp)  },
      danger:{mk:mannKendall(danger),slope:senSlope(danger)},
    };
  },[yearlyData]);

  // ── CORRELATION ──
  const corrMatrix = useMemo(()=>{
    const p=filteredRows.filter(r=>r.meanTemp!=null&&r.humidity!=null&&r.wetBulb!=null&&r.wbgt!=null);
    const T=p.map(r=>r.meanTemp),RH=p.map(r=>r.humidity),Tw=p.map(r=>r.wetBulb),W=p.map(r=>r.wbgt),Mx=p.map(r=>r.maxTemp);
    const vars=["Tw","WBGT","Temp","RH","Max T"],data=[Tw,W,T,RH,Mx];
    return{vars,matrix:vars.map((_,i)=>vars.map((__,j)=>pearsonR(data[i],data[j])))};
  },[filteredRows]);

  // ── SEASONAL ──
  const seasonalStats = useMemo(()=>{
    const s={"Pre-monsoon":[],"Monsoon":[],"Post-monsoon":[]};
    filteredRows.forEach(r=>{
      const mo=r.date.split("-")[1];
      if(r.wbgt!=null)s[getSeason(mo)].push({wbgt:r.wbgt,tw:r.wetBulb,temp:r.meanTemp,hum:r.humidity});
    });
    return Object.entries(s).map(([name,rows])=>({
      name,
      avgWbgt:avg(rows.map(r=>r.wbgt)),avgTw:avg(rows.map(r=>r.tw)),
      avgTemp:avg(rows.map(r=>r.temp)),avgHum:avg(rows.map(r=>r.hum)),
      dangerDays:rows.filter(r=>r.wbgt>=32&&r.wbgt<35).length,
      extremeDays:rows.filter(r=>r.wbgt>=35).length,
      cautionDays:rows.filter(r=>r.wbgt>=28&&r.wbgt<32).length,
      totalDays:rows.length,
    }));
  },[filteredRows]);

  // ── REGRESSION ──
  const regressionStats = useMemo(()=>{
    if(yearlyData.length<3)return null;
    const v=yearlyData.filter(y=>y.avgWbgt!=null);
    const x=v.map(y=>y.yearNum),y=v.map(y=>y.avgWbgt);
    const res=linearRegression(x,y);
    if(!res)return null;
    const trendLine=v.map((yr,i)=>({year:yr.year,predicted:parseFloat(res.yPred[i].toFixed(3)),actual:yr.avgWbgt}));
    const yTw=v.map(yr=>yr.avgTw).filter(Boolean);
    const xTw=v.filter(y=>y.avgTw!=null).map(y=>y.yearNum);
    const resTw=linearRegression(xTw,yTw);
    return{wbgt:res,tw:resTw,trendLine,years:v};
  },[yearlyData]);

  // ── EXCEEDANCE ──
  const exceedance = useMemo(()=>{
    const thresholds=[
      {label:"Tw ≥ 26°C",field:"wetBulb",val:26,color:"#facc15"},
      {label:"Tw ≥ 28°C",field:"wetBulb",val:28,color:"#fb923c"},
      {label:"Tw ≥ 30°C",field:"wetBulb",val:30,color:"#f97316"},
      {label:"Tw ≥ 32°C",field:"wetBulb",val:32,color:"#ef4444"},
      {label:"Tw ≥ 35°C",field:"wetBulb",val:35,color:"#dc2626"},
      {label:"WBGT ≥ 28°C",field:"wbgt",val:28,color:"#fbbf24"},
      {label:"WBGT ≥ 32°C",field:"wbgt",val:32,color:"#f97316"},
      {label:"WBGT ≥ 35°C",field:"wbgt",val:35,color:"#ef4444"},
    ];
    const byYear=yearlyData.map(yd=>{
      const yr=filteredRows.filter(r=>r.date.startsWith(yd.year));
      const obj={year:yd.year};
      thresholds.forEach(t=>{obj[t.label]=yr.filter(r=>r[t.field]!=null&&r[t.field]>=t.val).length;});
      return obj;
    });
    const overall=thresholds.map(t=>({
      label:t.label,color:t.color,
      total:filteredRows.filter(r=>r[t.field]!=null&&r[t.field]>=t.val).length,
      pct:filteredRows.length>0?parseFloat((filteredRows.filter(r=>r[t.field]!=null&&r[t.field]>=t.val).length/filteredRows.length*100).toFixed(1)):0,
    }));
    return{thresholds,byYear,overall};
  },[filteredRows,yearlyData]);

  // ── SUMMARY STATS ──
  const allTw    = filteredRows.map(r=>r.wetBulb).filter(Boolean);
  const allWbgt  = filteredRows.map(r=>r.wbgt).filter(Boolean);
  const dangerDays  = filteredRows.filter(r=>r.wbgt>=32&&r.wbgt<35).length;
  const extremeDays = filteredRows.filter(r=>r.wbgt>=35).length;
  const cautionDays = filteredRows.filter(r=>r.wbgt>=28&&r.wbgt<32).length;
  const maxTw   = allTw.length   ? Math.max(...allTw).toFixed(2)   : "—";
  const maxWbgt = allWbgt.length ? Math.max(...allWbgt).toFixed(2) : "—";

  const chartData = view==="monthly" ? monthlyData : filteredRows.map(r=>({...r,label:r.date.slice(5)}));
  const seasonColors={"Pre-monsoon":"#f97316","Monsoon":"#38bdf8","Post-monsoon":"#a78bfa"};

  // ── EXCEL ──
  const downloadExcel = () => {
    setDownloading(true);
    try {
      const wb=XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ["Date","Max T","Min T","Mean T","RH","Solar","Tw","WBGT","Risk","Season"],
        ...filteredRows.map(r=>[r.date,r.maxTemp,r.minTemp,r.meanTemp,r.humidity,r.solar,r.wetBulb,r.wbgt,r.wbgtRisk,getSeason(r.date.split("-")[1])])
      ]),"Daily Data");
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ["Month","MaxT","MinT","MeanT","RH","Solar","Tw","WBGT","Caution","Danger","Extreme"],
        ...monthlyData.map(m=>[m.label,m.maxTemp,m.minTemp,m.meanTemp,m.humidity,m.solar,m.wetBulb,m.wbgt,m.cautionDays,m.dangerDays,m.extremeDays])
      ]),"Monthly Averages");
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ["Year","Avg Tw","Avg WBGT","Avg T","Avg RH","Caution","Danger","Extreme","Days"],
        ...yearlyData.map(y=>[y.year,y.avgTw,y.avgWbgt,y.avgTemp,y.avgHum,y.cautionDays,y.dangerDays,y.extremeDays,y.totalDays])
      ]),"Yearly Summary");
      const mkRows=[["Variable","S","Z","p-value","Sig?","Sen Slope","Trend"]];
      if(mkStats)[["Tw",mkStats.tw],["WBGT",mkStats.wbgt],["Temp",mkStats.temp],["Danger Days",mkStats.danger]]
        .forEach(([n,{mk,slope}])=>mkRows.push([n,mk.S,mk.Z,mk.pValue,mk.pValue<0.05?"Yes":"No",slope,mk.trend]));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(mkRows),"Mann-Kendall");
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ["",...corrMatrix.vars],...corrMatrix.matrix.map((row,i)=>[corrMatrix.vars[i],...row])
      ]),"Correlation Matrix");
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ["Season","Avg WBGT","Avg Tw","Avg T","Avg RH","Caution","Danger","Extreme","Total"],
        ...seasonalStats.map(s=>[s.name,s.avgWbgt,s.avgTw,s.avgTemp,s.avgHum,s.cautionDays,s.dangerDays,s.extremeDays,s.totalDays])
      ]),"Seasonal Analysis");
      if(regressionStats)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ["Variable","Slope/yr","Intercept","R²","p-value","Sig?"],
        ["WBGT",regressionStats.wbgt.slope,regressionStats.wbgt.intercept,regressionStats.wbgt.r2,regressionStats.wbgt.pValue,regressionStats.wbgt.pValue<0.05?"Yes":"No"],
        ["Tw",regressionStats.tw?.slope,regressionStats.tw?.intercept,regressionStats.tw?.r2,regressionStats.tw?.pValue,regressionStats.tw?.pValue<0.05?"Yes":"No"],
      ]),"OLS Regression");
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ["Threshold","Total Days","% of Period"],...exceedance.overall.map(e=>[e.label,e.total,e.pct+"%"])
      ]),"Exceedance");
      XLSX.writeFile(wb,`Khulna_HeatStress_${startDate}_to_${endDate}.xlsx`);
    } catch(e){alert("Download failed: "+e.message);}
    finally{setDownloading(false);}
  };

  const TABS=[
    ["mk","📊 Mann-Kendall"],["corr","🔗 Correlation"],["seasonal","🌿 Seasonal"],
    ["regression","📐 Regression"],["exceedance","📈 Exceedance"],
    ["heatstroke","🔴 Heat Stroke"],["wetbulb","💧 Wet Bulb"],
    ["trend","📈 Year Trend"],["temp","🌡️ Temperature"],["humidity","💧 Humidity"],["solar","☀️ Solar"],
  ];

  // ─── RENDER ───────────────────────────────────────────────────────────────
  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(160deg,#030710 0%,#050c18 60%,#060d1c 100%)",fontFamily:"Georgia,serif",color:"#cce8f8",paddingBottom:56}}>

      {/* HEADER */}
      <div style={{padding:"20px 20px 14px",borderBottom:"1px solid #0c2030",background:"rgba(0,0,0,0.45)",backdropFilter:"blur(12px)",position:"sticky",top:0,zIndex:20}}>
        <div style={{maxWidth:1100,margin:"0 auto",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
          <div>
            <div style={{fontSize:9,letterSpacing:4,color:"#ef4444",textTransform:"uppercase",marginBottom:2}}>Heat Stress Research · Khulna, Bangladesh · ERA5 · Mann-Kendall + Correlation + Regression</div>
            <h1 style={{margin:0,fontSize:22,fontWeight:700,color:"#fff"}}>🌡️ Khulna Heat Stroke Risk Index</h1>
            <div style={{fontSize:11,color:"#64a7c8",marginTop:1}}>Wet Bulb · WBGT · Mann-Kendall · Sen's Slope · Pearson r · OLS · Seasonal · Exceedance</div>
          </div>
          <div style={{display:'flex', gap:8, flexWrap:'wrap', alignItems:'center'}}>
            <button onClick={()=>setPaperMode(p=>!p)} style={{
              display:"flex", alignItems:"center", gap:7, padding:"9px 16px", borderRadius:10,
              border:`1px solid ${paperMode?'#16a34a':'#38bdf8'}`,
              background: paperMode ? 'linear-gradient(135deg,#052e16,#14532d)' : 'rgba(255,255,255,0.05)',
              color: paperMode ? '#4ade80' : '#7dd3fc',
              cursor:"pointer", fontSize:12, fontFamily:"inherit", fontWeight:600,
            }}>
              {paperMode ? '🟢 Paper Mode ON — click charts then download' : '📄 Paper Mode (white theme for export)'}
            </button>
            <button onClick={downloadExcel} disabled={downloading||!filteredRows.length} style={{
              display:"flex", alignItems:"center", gap:7, padding:"9px 16px", borderRadius:10,
              border:"1px solid #ef4444", background:"linear-gradient(135deg,#3b0000,#5c0a0a)",
              color:"#fca5a5", cursor:"pointer", fontSize:12, fontFamily:"inherit", fontWeight:600,
              opacity:!filteredRows.length?0.5:1
            }}>
              📥 {downloading?"Exporting…":`Download Excel (${filteredRows.length.toLocaleString()} days)`}
            </button>
          </div>
        </div>
      </div>

      <div style={{maxWidth:1100,margin:"0 auto",padding:"0 14px"}}>

        {/* ── DATE RANGE CONTROL ─────────────────────────────────────────── */}
        <div style={{...box,marginTop:18,padding:"18px 20px"}}>
          <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase",marginBottom:14}}>📅 Data Range Selection</div>

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,flexWrap:"wrap"}}>

            {/* LEFT: Preset dropdown */}
            <div>
              <div style={{fontSize:10,color:"#4a8aaa",marginBottom:6,fontWeight:600}}>MULTI-YEAR PRESET RANGE</div>
              <select
                value={presetIdx}
                onChange={e=>{setCustomYear("");applyPreset(parseInt(e.target.value));}}
                style={{...iStyle,width:"100%",cursor:"pointer",appearance:"auto"}}
              >
                {PRESET_RANGES.map((p,i)=>(
                  <option key={i} value={i}>{p.label}</option>
                ))}
              </select>
              <div style={{fontSize:10,color:"#4a6a7a",marginTop:6,fontStyle:"italic"}}>
                Fetches full range from ERA5 for robust Mann-Kendall analysis
              </div>
            </div>

            {/* RIGHT: Custom year + manual dates */}
            <div>
              <div style={{fontSize:10,color:"#4a8aaa",marginBottom:6,fontWeight:600}}>VIEW SINGLE YEAR (filter only)</div>
              <div style={{display:"flex",gap:8,alignItems:"flex-start"}}>
                <input
                  type="number"
                  placeholder={`e.g. 2023`}
                  value={customYear}
                  min="2000"
                  max={currentYear}
                  onChange={e=>{setCustomYear(e.target.value);setCustomYearErr("");}}
                  onKeyDown={e=>e.key==="Enter"&&applyCustomYear()}
                  style={{...iStyle,width:100,textAlign:"center"}}
                />
                <button onClick={applyCustomYear} style={{padding:"7px 14px",borderRadius:8,border:"1px solid #38bdf8",background:"#0c3a5e",color:"#7dd3fc",cursor:"pointer",fontSize:11,fontFamily:"inherit",whiteSpace:"nowrap"}}>
                  View Year
                </button>
                {startDate.slice(0,4)===endDate.slice(0,4) && startDate.slice(0,4)!=="2000" && (
                  <button
                    onClick={()=>{setCustomYear("");applyPreset(presetIdx);}}
                    style={{padding:"7px 14px",borderRadius:8,border:"1px solid #4a8aaa",background:"transparent",color:"#4a8aaa",cursor:"pointer",fontSize:11,fontFamily:"inherit"}}
                  >
                    ← Back to full range
                  </button>
                )}
              </div>
              {customYearErr && <div style={{fontSize:11,color:"#f87171",marginTop:5}}>{customYearErr}</div>}
              <div style={{fontSize:10,color:"#4a6a7a",marginTop:6,fontStyle:"italic"}}>
                Note: single-year view shows descriptive stats only — MK trend needs 4+ years
              </div>
            </div>
          </div>

          {/* Manual date pickers row */}
          <div style={{marginTop:14,paddingTop:14,borderTop:"1px solid #0f3455"}}>
            <div style={{fontSize:10,color:"#4a8aaa",marginBottom:8,fontWeight:600}}>MANUAL DATE RANGE (custom analysis period)</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:10,alignItems:"center"}}>
              <div>
                <div style={{fontSize:9,color:"#4a8aaa",marginBottom:4}}>START</div>
                <input type="date" value={startDate} min={ERA5_EARLIEST} max={endDate}
                  onChange={e=>{setStartDate(e.target.value);setCustomYear("");}}
                  style={iStyle}/>
              </div>
              <div style={{color:"#1e4d6b",fontSize:16}}>→</div>
              <div>
                <div style={{fontSize:9,color:"#4a8aaa",marginBottom:4}}>END</div>
                <input type="date" value={endDate} min={startDate} max={todayStr}
                  onChange={e=>{setEndDate(e.target.value);setCustomYear("");}}
                  style={iStyle}/>
              </div>
              <button
                onClick={()=>{
                  if(startDate<fetchStart||endDate>fetchEnd){
                    doFetch(startDate,endDate);
                    setFetchStart(startDate);setFetchEnd(endDate);
                  }
                }}
                style={{padding:"7px 14px",borderRadius:8,border:"1px solid #38bdf8",background:"#0c3a5e",color:"#7dd3fc",cursor:"pointer",fontSize:11,fontFamily:"inherit",alignSelf:"flex-end"}}
              >
                Apply & Fetch
              </button>
            </div>
          </div>

          {/* Status bar */}
          <div style={{marginTop:12,padding:"8px 12px",background:"rgba(0,0,0,0.2)",borderRadius:8,display:"flex",flexWrap:"wrap",gap:16,alignItems:"center"}}>
            <div style={{fontSize:11,color:"#4a8aaa"}}>
              Fetched: <b style={{color:"#7dd3fc"}}>{fetchStart}</b> → <b style={{color:"#7dd3fc"}}>{fetchEnd}</b>
              <span style={{color:"#4a6a7a",marginLeft:8}}>({allDailyRows.length.toLocaleString()} days loaded)</span>
            </div>
            <div style={{fontSize:11,color:"#4a8aaa"}}>
              Analysing: <b style={{color:"#a78bfa"}}>{startDate}</b> → <b style={{color:"#a78bfa"}}>{endDate}</b>
              <span style={{color:"#4a6a7a",marginLeft:8}}>({filteredRows.length.toLocaleString()} days · {yearlyData.length} years)</span>
            </div>
            {yearlyData.length < 4 && (
              <div style={{fontSize:11,color:"#f87171"}}>⚠ Need ≥4 years for Mann-Kendall trend analysis</div>
            )}
            {yearlyData.length >= 4 && yearlyData.length < 10 && (
              <div style={{fontSize:11,color:"#facc15"}}>⚡ {yearlyData.length} years — trend detectable but limited statistical power</div>
            )}
            {yearlyData.length >= 10 && (
              <div style={{fontSize:11,color:"#4ade80"}}>✓ {yearlyData.length} years — good statistical power for Mann-Kendall</div>
            )}
          </div>
        </div>

        {loading && (
          <div style={{textAlign:"center",padding:"80px 0",color:"#38bdf8"}}>
            <div style={{fontSize:42}}>🌀</div>
            <div style={{marginTop:10,letterSpacing:2}}>Loading ERA5 archive ({fetchStart} → {fetchEnd})…</div>
            <div style={{marginTop:6,fontSize:11,color:"#4a8aaa"}}>Larger ranges take longer — please wait</div>
          </div>
        )}
        {error && (
          <div style={{textAlign:"center",padding:"60px 0",color:"#f87171"}}>
            <div style={{fontSize:30}}>⚠️</div>
            <div style={{marginTop:8}}>{error}</div>
            <button onClick={()=>doFetch(fetchStart,fetchEnd)} style={{marginTop:12,padding:"8px 16px",borderRadius:8,border:"1px solid #ef4444",background:"transparent",color:"#fca5a5",cursor:"pointer",fontFamily:"inherit"}}>Retry</button>
          </div>
        )}

        {!loading && !error && filteredRows.length > 0 && (<>

          {/* STAT CARDS */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(148px,1fr))",gap:10,marginBottom:16}}>
            {[
              {label:"Avg Wet Bulb",  value:avg(allTw)!=null?avg(allTw)+"°C":"—",   icon:"💧",color:"#38bdf8",sub:"Stull 2011"},
              {label:"Max Wet Bulb",  value:maxTw+"°C",                              icon:"⚠️",color:"#f97316",sub:"Peak recorded"},
              {label:"Avg WBGT",      value:avg(allWbgt)!=null?avg(allWbgt)+"°C":"—",icon:"🌡️",color:"#facc15",sub:"Heat stroke index"},
              {label:"Max WBGT",      value:maxWbgt+"°C",                            icon:"🔥",color:"#ef4444",sub:"Peak recorded"},
              {label:"🟡 Caution",    value:cautionDays,                             icon:"", color:"#facc15",sub:"WBGT 28–32°C"},
              {label:"🟠 Danger",     value:dangerDays,                              icon:"", color:"#f97316",sub:"WBGT 32–35°C"},
              {label:"🔴 Extreme",    value:extremeDays,                             icon:"", color:"#ef4444",sub:"WBGT ≥35°C"},
            ].map((s,i)=>(
              <div key={i} style={{background:"rgba(255,255,255,0.04)",border:"1px solid #1a3a50",borderRadius:12,padding:"12px 14px"}}>
                <div style={{fontSize:18,marginBottom:4}}>{s.icon}</div>
                <div style={{fontSize:19,fontWeight:700,color:s.color}}>{s.value}</div>
                <div style={{fontSize:10,color:"#94a3b8",marginTop:2}}>{s.label}</div>
                <div style={{fontSize:9,color:"#4a6a7a",marginTop:1,fontStyle:"italic"}}>{s.sub}</div>
              </div>
            ))}
          </div>

          {/* FIGURES DOWNLOAD PANEL */}
          <div style={{...box, background:'rgba(56,189,248,0.04)', border:'1px solid #0c4a6e', marginBottom:14}}>
            <div style={{display:'flex', alignItems:'center', justifyContent:'space-between', flexWrap:'wrap', gap:10, marginBottom:12}}>
              <div style={{fontSize:9, color:'#38bdf8', letterSpacing:3, textTransform:'uppercase'}}>
                📄 Download Figures for Paper — PNG 300 dpi · White Background · Print-Ready
              </div>

              {/* ── DOWNLOAD ALL BUTTON ── */}
              <button
                onClick={handleDownloadAll}
                disabled={dlAll.busy}
                style={{
                  display:'flex', alignItems:'center', gap:8, padding:'10px 20px',
                  borderRadius:10, cursor: dlAll.busy ? 'wait' : 'pointer',
                  border: dlAll.busy ? '1px solid #4ade80' : '2px solid #38bdf8',
                  background: dlAll.busy
                    ? 'linear-gradient(135deg,#052e16,#14532d)'
                    : 'linear-gradient(135deg,#0c3a5e,#1e4d6b)',
                  color: dlAll.busy ? '#4ade80' : '#7dd3fc',
                  fontSize:13, fontFamily:'inherit', fontWeight:700,
                  boxShadow: dlAll.busy ? '0 0 16px rgba(74,222,128,0.3)' : '0 0 16px rgba(56,189,248,0.2)',
                  transition:'all 0.2s',
                }}>
                {dlAll.busy ? (
                  <>
                    <span style={{fontSize:16}}>⏳</span>
                    <span>
                      Saving {dlAll.current}/{dlAll.total}…
                      <span style={{fontSize:10, fontWeight:400, display:'block', color:'#86efac'}}>
                        {dlAll.filename}
                      </span>
                    </span>
                    {/* progress bar */}
                    <div style={{
                      position:'absolute', bottom:0, left:0,
                      height:3, borderRadius:2, background:'#4ade80',
                      width:`${(dlAll.current/dlAll.total)*100}%`,
                      transition:'width 0.3s',
                    }}/>
                  </>
                ) : (
                  <>📁 Download All Figures to Folder</>
                )}
              </button>
            </div>

            {/* Instruction row */}
            <div style={{fontSize:11, color:'#4a8aaa', marginBottom:10, lineHeight:1.7}}>
              <b style={{color:'#7dd3fc'}}>How to use:</b> Enable 📄 Paper Mode → visit each tab so charts render → click
              <b style={{color:'#38bdf8'}}> Download All Figures to Folder</b> to pick a folder and save all 7 figures at once.
              Or download individually below.
              <span style={{color:'#facc15', marginLeft:6}}>
                ⚠ Fig 1 (Study Area Map) must be created separately in QGIS or Google Maps.
              </span>
            </div>

            {/* ── TAB NAVIGATION GUIDE ── */}
            <div style={{display:'flex', gap:6, flexWrap:'wrap', marginBottom:12}}>
              {[
                {tab:'trend',      figs:'Fig 2 & 3'},
                {tab:'seasonal',   figs:'Fig 4'},
                {tab:'exceedance', figs:'Fig 5'},
                {tab:'regression', figs:'Fig 6'},
                {tab:'heatstroke', figs:'Fig 7'},
                {tab:'wetbulb',    figs:'Fig 8'},
              ].map(({tab:t,figs})=>(
                <button key={t} onClick={()=>setTab(t)} style={{
                  padding:'4px 10px', borderRadius:6, fontSize:10, fontFamily:'inherit',
                  border:'1px solid #1e4d6b', background: tab===t?'#0c3a5e':'transparent',
                  color: tab===t?'#7dd3fc':'#4a8aaa', cursor:'pointer',
                }}>
                  {figs} → {t}
                </button>
              ))}
            </div>

            {/* ── PER-FIGURE DOWNLOAD ROWS ── */}
            <div style={{display:'flex', flexDirection:'column', gap:5}}>
              {figList.map((f,i) => (
                <div key={i} style={{display:'flex', alignItems:'center', justifyContent:'space-between',
                  padding:'7px 10px', borderRadius:7, background:'rgba(0,0,0,0.2)',
                  border:'1px solid #0f3455'}}>
                  <div>
                    <span style={{fontSize:11, color:'#7dd3fc', fontWeight:700}}>{f.num}</span>
                    <span style={{fontSize:10, color:'#64a7c8', marginLeft:6}}>{f.label}</span>
                    <span style={{fontSize:9, color:'#4a6a7a', marginLeft:8}}>{f.filename}</span>
                  </div>
                  <DownloadBtn containerRef={f.ref} filename={f.filename} figLabel={f.label} figNum={f.num}/>
                </div>
              ))}
            </div>

            <div style={{marginTop:10, padding:'8px 12px', borderRadius:7,
              background:'rgba(139,92,246,0.08)', border:'1px solid #4c1d9533', fontSize:11, color:'#94a3b8'}}>
              💡 <b style={{color:'#c4b5fd'}}>For Fig. 1 (Study Area Map):</b> Use QGIS (free) — add OpenStreetMap base layer,
              mark the ERA5 grid cell at 22.81°N 89.56°E, add Khulna city boundary and Sundarbans polygon.
              Export as <code style={{color:'#7dd3fc'}}>Fig1_StudyArea_Khulna.png</code> at 300 dpi.
            </div>
          </div>

          {/* TABS */}
          <div style={{display:"flex",gap:3,background:"rgba(255,255,255,0.03)",borderRadius:12,padding:4,border:"1px solid #0f3455",marginBottom:14,flexWrap:"wrap"}}>
            {TABS.map(([v,l])=>(
              <button key={v} onClick={()=>setTab(v)} style={{padding:"6px 13px",borderRadius:8,border:"none",background:tab===v?"#0c3a5e":"transparent",color:tab===v?"#7dd3fc":"#4a8aaa",cursor:"pointer",fontSize:11,fontFamily:"inherit",transition:"all 0.15s",fontWeight:tab===v?700:400}}>{l}</button>
            ))}
          </div>

          {/* VIEW TOGGLE */}
          {!["mk","trend","corr","seasonal","regression","exceedance"].includes(tab)&&(
            <div style={{display:"flex",gap:3,background:"rgba(255,255,255,0.03)",borderRadius:10,padding:3,border:"1px solid #0f3455",marginBottom:14,width:"fit-content"}}>
              {[["monthly","Monthly Avg"],["daily","Daily View"]].map(([v,l])=>(
                <button key={v} onClick={()=>setView(v)} style={{padding:"5px 14px",borderRadius:7,border:"none",background:view===v?"#0c3a5e":"transparent",color:view===v?"#7dd3fc":"#4a8aaa",cursor:"pointer",fontSize:11,fontFamily:"inherit"}}>{l}</button>
              ))}
            </div>
          )}

          {/* ══ MANN-KENDALL ══ */}
          {tab==="mk"&&(
            <>
              <div style={{...box,background:"rgba(139,92,246,0.06)",border:"1px solid #4c1d9588",padding:"16px 20px",marginBottom:16}}>
                <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:8}}>📐 What These Numbers Mean</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:10,fontSize:11,color:"#94a3b8",lineHeight:1.7}}>
                  <div><b style={{color:"#c4b5fd"}}>Mann-Kendall S:</b> Sum of all pairwise trend signs. Positive = upward.</div>
                  <div><b style={{color:"#c4b5fd"}}>Z Statistic:</b> |Z| {">"}1.96 means p {"<"} 0.05 (significant).</div>
                  <div><b style={{color:"#c4b5fd"}}>p-value:</b> p {"<"} 0.05 = statistically significant ✓</div>
                  <div><b style={{color:"#c4b5fd"}}>Sen's Slope:</b> Rate of change per year.</div>
                </div>
              </div>
              {mkStats?(
                <div style={{display:"flex",flexWrap:"wrap",gap:12,marginBottom:16}}>
                  <MKCard title="Wet Bulb Temperature" mk={mkStats.tw.mk}     slope={mkStats.tw.slope}     unit="°C"    color="#38bdf8"/>
                  <MKCard title="WBGT"                 mk={mkStats.wbgt.mk}   slope={mkStats.wbgt.slope}   unit="°C"    color="#f97316"/>
                  <MKCard title="Mean Air Temperature" mk={mkStats.temp.mk}   slope={mkStats.temp.slope}   unit="°C"    color="#facc15"/>
                  <MKCard title="Danger Days Count"    mk={mkStats.danger.mk} slope={mkStats.danger.slope} unit=" days" color="#ef4444"/>
                </div>
              ):(
                <div style={{textAlign:"center",color:"#f87171",padding:"30px 0",background:"rgba(239,68,68,0.05)",borderRadius:12,border:"1px solid #ef444433"}}>
                  ⚠ Only {yearlyData.length} year(s) in selected range — need ≥ 4 for Mann-Kendall.<br/>
                  <span style={{fontSize:11,color:"#4a8aaa"}}>Use the preset dropdown above to select a multi-year range.</span>
                </div>
              )}
              {mkStats&&(
                <div style={{...box,background:"rgba(56,189,248,0.04)",border:"1px solid #0c4a6e"}}>
                  <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>📝 Ready-to-Paste Results Text</div>
                  <div style={{fontSize:12,color:"#cbd5e1",lineHeight:1.9,background:"rgba(0,0,0,0.3)",borderRadius:10,padding:"14px 16px"}}>
                    {"The Mann-Kendall trend test revealed a "}
                    <span style={{color:mkStats.tw.mk.pValue<0.05?"#4ade80":"#fbbf24"}}>{mkStats.tw.mk.trend.toLowerCase()}</span>
                    {" in annual mean wet bulb temperature over Khulna (S = "}<b style={{color:"#7dd3fc"}}>{mkStats.tw.mk.S}</b>
                    {", Z = "}<b style={{color:"#7dd3fc"}}>{mkStats.tw.mk.Z}</b>
                    {", p = "}<b style={{color:mkStats.tw.mk.pValue<0.05?"#4ade80":"#fbbf24"}}>{mkStats.tw.mk.pValue}</b>
                    {"). Sen's slope = "}<b style={{color:"#f97316"}}>{mkStats.tw.slope>0?"+":""}{mkStats.tw.slope}°C/yr</b>
                    {". WBGT showed a "}<span style={{color:mkStats.wbgt.mk.pValue<0.05?"#4ade80":"#fbbf24"}}>{mkStats.wbgt.mk.trend.toLowerCase()}</span>
                    {" (Z = "}<b style={{color:"#7dd3fc"}}>{mkStats.wbgt.mk.Z}</b>
                    {", p = "}<b style={{color:mkStats.wbgt.mk.pValue<0.05?"#4ade80":"#fbbf24"}}>{mkStats.wbgt.mk.pValue}</b>
                    {", Sen's slope = "}<b style={{color:"#f97316"}}>{mkStats.wbgt.slope>0?"+":""}{mkStats.wbgt.slope}°C/yr</b>
                    {"). Danger days showed a "}<span style={{color:mkStats.danger.mk.pValue<0.05?"#4ade80":"#fbbf24"}}>{mkStats.danger.mk.trend.toLowerCase()}</span>
                    {" (Z = "}<b style={{color:"#7dd3fc"}}>{mkStats.danger.mk.Z}</b>
                    {", p = "}<b style={{color:mkStats.danger.mk.pValue<0.05?"#4ade80":"#fbbf24"}}>{mkStats.danger.mk.pValue}</b>
                    {") at "}<b style={{color:"#f97316"}}>{mkStats.danger.slope>0?"+":""}{mkStats.danger.slope} days/yr</b>{"."}
                  </div>
                </div>
              )}
              {yearlyData.length>0&&(
                <div style={box}>
                  <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:12}}>Annual Summary Table</div>
                  <div style={{overflowX:"auto"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                      <thead><tr style={{color:"#64a7c8",borderBottom:"1px solid #1e4d6b"}}>
                        {["Year","Avg Tw","Avg WBGT","Avg T","Avg RH","🟡 Caution","🟠 Danger","🔴 Extreme","Days"].map((h,i)=>(
                          <th key={i} style={{padding:"7px 11px",textAlign:"left",fontSize:10,fontWeight:600}}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {yearlyData.map((y,i)=>(
                          <tr key={i} style={{borderBottom:"1px solid #0a1e2e",background:i%2===0?"rgba(255,255,255,0.02)":"transparent"}}>
                            <td style={{padding:"7px 11px",color:"#7dd3fc",fontWeight:700}}>{y.year}</td>
                            <td style={{padding:"7px 11px",color:"#38bdf8"}}>{y.avgTw}°C</td>
                            <td style={{padding:"7px 11px",color:y.avgWbgt>=32?"#ef4444":y.avgWbgt>=28?"#f97316":"#facc15"}}>{y.avgWbgt}°C</td>
                            <td style={{padding:"7px 11px",color:"#fb923c"}}>{y.avgTemp}°C</td>
                            <td style={{padding:"7px 11px",color:"#94a3b8"}}>{y.avgHum}%</td>
                            <td style={{padding:"7px 11px",color:"#facc15"}}>{y.cautionDays}</td>
                            <td style={{padding:"7px 11px",color:"#f97316"}}>{y.dangerDays}</td>
                            <td style={{padding:"7px 11px",color:"#ef4444",fontWeight:y.extremeDays>0?700:400}}>{y.extremeDays}</td>
                            <td style={{padding:"7px 11px",color:"#64a7c8"}}>{y.totalDays}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </>
          )}

          {/* ══ CORRELATION ══ */}
          {tab==="corr"&&(
            <>
              <div style={{...box,background:"rgba(56,189,248,0.04)",border:"1px solid #0c4a6e"}}>
                <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>🔗 Pearson Correlation Matrix — Daily Data</div>
                <div style={{fontSize:11,color:"#94a3b8",marginBottom:14,lineHeight:1.7}}>Pearson r: Red = strong positive, Blue = strong negative. Values ≥|0.8| are very strong.</div>
                <div style={{overflowX:"auto"}}>
                  <table style={{borderCollapse:"separate",borderSpacing:4,fontSize:12}}>
                    <thead><tr>
                      <th style={{padding:"8px 16px",fontSize:10,color:"#4a8aaa"}}></th>
                      {corrMatrix.vars.map((v,i)=><th key={i} style={{padding:"8px 16px",fontSize:11,color:"#7dd3fc",fontWeight:700,textAlign:"center"}}>{v}</th>)}
                    </tr></thead>
                    <tbody>
                      {corrMatrix.matrix.map((row,i)=>(
                        <tr key={i}>
                          <td style={{padding:"8px 16px",fontSize:11,color:"#7dd3fc",fontWeight:700,whiteSpace:"nowrap"}}>{corrMatrix.vars[i]}</td>
                          {row.map((r,j)=>(
                            <td key={j} style={{padding:"10px 18px",textAlign:"center",borderRadius:8,background:corrColor(r),fontWeight:i===j?700:r!=null&&Math.abs(r)>=0.8?700:400,color:i===j?"#94a3b8":"#e0f4ff",fontSize:13}}>
                              {i===j?"—":r!=null?r.toFixed(3):"—"}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              <div style={{...box,background:"rgba(56,189,248,0.04)",border:"1px solid #0c4a6e"}}>
                <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>📝 Ready-to-Paste Text</div>
                <div style={{fontSize:12,color:"#cbd5e1",lineHeight:1.9,background:"rgba(0,0,0,0.3)",borderRadius:10,padding:"14px 16px"}}>
                  {(()=>{
                    const tw_wbgt=corrMatrix.matrix[0]?.[1],tw_temp=corrMatrix.matrix[0]?.[2];
                    const tw_rh=corrMatrix.matrix[0]?.[3],wbgt_temp=corrMatrix.matrix[1]?.[2],wbgt_rh=corrMatrix.matrix[1]?.[3];
                    return<>Pearson correlation analysis revealed strong positive relationships among all heat stress variables. Wet bulb temperature (Tw) showed a very strong positive correlation with WBGT (r = <b style={{color:"#7dd3fc"}}>{tw_wbgt}</b>). Tw was also strongly correlated with air temperature (r = <b style={{color:"#7dd3fc"}}>{tw_temp}</b>) and relative humidity (r = <b style={{color:"#7dd3fc"}}>{tw_rh}</b>). WBGT showed correlations of r = <b style={{color:"#7dd3fc"}}>{wbgt_temp}</b> with air temperature and r = <b style={{color:"#7dd3fc"}}>{wbgt_rh}</b> with relative humidity.</>;
                  })()}
                </div>
              </div>
            </>
          )}

          {/* ══ SEASONAL ══ */}
          {tab==="seasonal"&&(
            <>
              <div style={box}>
                <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:12}}>🌿 Seasonal WBGT — Dot Plot</div>
                <DownloadBtn containerRef={refFig4} filename="Fig4_Seasonal_WBGT.png" figLabel="Seasonal WBGT comparison bar chart" figNum="Fig. 4"/>
                <div ref={refFig4} style={{background:th.bg, padding: paperMode?'8px':0, borderRadius:8}}>
                {/* Cleveland dot plot built with ComposedChart */}
                <ResponsiveContainer width="100%" height={300}>
                  <ComposedChart data={seasonalStats} layout="vertical"
                    margin={{top:10,right:60,left:110,bottom:10}}>
                    <CartesianGrid stroke={th.grid} strokeDasharray={th.gridDash} horizontal={false}/>
                    <XAxis type="number" tick={{fill:th.axTick,fontSize:9}} unit="°C"
                           domain={[0,40]} ticks={[0,5,10,15,20,25,30,35,40]}/>
                    <YAxis type="category" dataKey="name" tick={{fill:th.axTick,fontSize:11}} width={105}/>
                    <Tooltip content={<CT/>}/>
                    <Legend wrapperStyle={{color:th.subtext,fontSize:10}}/>
                    <ReferenceLine x={35} stroke={th.refExtreme} strokeDasharray="4 2"
                      label={{value:"Extreme ≥35°C",fill:th.refExtreme,fontSize:9,position:"top"}}/>
                    <ReferenceLine x={32} stroke={th.refDanger}  strokeDasharray="4 2"
                      label={{value:"Danger ≥32°C", fill:th.refDanger, fontSize:9,position:"top"}}/>
                    <ReferenceLine x={28} stroke={th.refCaution} strokeDasharray="4 2"
                      label={{value:"Caution ≥28°C",fill:th.refCaution,fontSize:9,position:"top"}}/>
                    {/* WBGT stem */}
                    <Bar dataKey="avgWbgt" name="Avg WBGT" fill="none"
                         shape={(props)=>{
                           const {x,y,width,height,value} = props;
                           const cy2 = y + height/2;
                           return <line x1={x} y1={cy2} x2={x+width} y2={cy2}
                                        stroke={th.barWbgt} strokeWidth={3}/>;
                         }} unit="°C"/>
                    {/* Tw stem */}
                    <Bar dataKey="avgTw" name="Avg Wet Bulb" fill="none"
                         shape={(props)=>{
                           const {x,y,width,height,value} = props;
                           const cy2 = y + height/2;
                           return <line x1={x} y1={cy2} x2={x+width} y2={cy2}
                                        stroke={th.barTw} strokeWidth={3}/>;
                         }} unit="°C"/>
                    {/* WBGT dot */}
                    <Scatter dataKey="avgWbgt" name="WBGT dot" fill={th.barWbgt}
                             shape={(props)=>{
                               const {cx:sx,cy:sy} = props;
                               return <circle cx={sx} cy={sy} r={8}
                                              fill={th.barWbgt}
                                              stroke={paperMode?'#fff':'#111'} strokeWidth={2}/>;
                             }} unit="°C"/>
                    {/* Tw dot */}
                    <Scatter dataKey="avgTw" name="Tw dot" fill={th.barTw}
                             shape={(props)=>{
                               const {cx:sx,cy:sy} = props;
                               return <circle cx={sx} cy={sy} r={8}
                                              fill={th.barTw}
                                              stroke={paperMode?'#fff':'#111'} strokeWidth={2}/>;
                             }} unit="°C"/>
                  </ComposedChart>
                </ResponsiveContainer>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:12,marginBottom:16}}>
                {seasonalStats.map((s,i)=>(
                  <div key={i} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${seasonColors[s.name]}44`,borderRadius:14,padding:"16px 18px"}}>
                    <div style={{fontSize:11,color:seasonColors[s.name],fontWeight:700,letterSpacing:2,textTransform:"uppercase",marginBottom:10}}>{s.name}</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                      {[
                        {label:"Avg WBGT",value:s.avgWbgt+"°C",color:"#f97316"},
                        {label:"Avg Tw",value:s.avgTw+"°C",color:"#38bdf8"},
                        {label:"Avg Temp",value:s.avgTemp+"°C",color:"#facc15"},
                        {label:"Avg RH",value:s.avgHum+"%",color:"#94a3b8"},
                        {label:"Caution Days",value:s.cautionDays,color:"#facc15"},
                        {label:"Danger Days",value:s.dangerDays,color:"#f97316"},
                        {label:"Extreme Days",value:s.extremeDays,color:"#ef4444"},
                        {label:"Total Days",value:s.totalDays,color:"#64a7c8"},
                      ].map((m,j)=>(
                        <div key={j} style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:"8px 10px"}}>
                          <div style={{fontSize:9,color:"#4a8aaa",marginBottom:2}}>{m.label}</div>
                          <div style={{fontSize:14,fontWeight:700,color:m.color}}>{m.value}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
              <div style={{...box,background:"rgba(56,189,248,0.04)",border:"1px solid #0c4a6e"}}>
                <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>📝 Ready-to-Paste Seasonal Text</div>
                <div style={{fontSize:12,color:"#cbd5e1",lineHeight:1.9,background:"rgba(0,0,0,0.3)",borderRadius:10,padding:"14px 16px"}}>
                  {(()=>{
                    const pre=seasonalStats.find(s=>s.name==="Pre-monsoon");
                    const mon=seasonalStats.find(s=>s.name==="Monsoon");
                    const post=seasonalStats.find(s=>s.name==="Post-monsoon");
                    if(!pre||!mon||!post)return"Loading…";
                    return<>Seasonal analysis revealed marked variation in heat stress across the year. The pre-monsoon season (March–May) recorded the highest average WBGT of <b style={{color:"#f97316"}}>{pre.avgWbgt}°C</b> with <b style={{color:"#ef4444"}}>{pre.dangerDays}</b> danger days and <b style={{color:"#dc2626"}}>{pre.extremeDays}</b> extreme days, representing the peak physiological heat stress period. The monsoon season (June–September) showed an average WBGT of <b style={{color:"#38bdf8"}}>{mon.avgWbgt}°C</b> with avg Tw = <b style={{color:"#38bdf8"}}>{mon.avgTw}°C</b> and avg RH = <b style={{color:"#94a3b8"}}>{mon.avgHum}%</b>. Post-monsoon (October–February) had the lowest heat stress (avg WBGT = <b style={{color:"#a78bfa"}}>{post.avgWbgt}°C</b>), highlighting the pre-monsoon period as the critical intervention window.</>;
                  })()}
                </div>
              </div>
            </>
          )}

          {/* ══ REGRESSION ══ */}
          {tab==="regression"&&(
            <>
              <div style={{...box,background:"rgba(139,92,246,0.06)",border:"1px solid #4c1d9588"}}>
                <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:8}}>📐 Why OLS + Mann-Kendall Together?</div>
                <div style={{fontSize:11,color:"#94a3b8",lineHeight:1.7}}>Mann-Kendall (non-parametric) detects monotonic trend without assuming normality. OLS regression (parametric) gives R² and a visual trend line. Journals expect both.</div>
              </div>
              {regressionStats?(
                <>
                  <div style={{display:"flex",flexWrap:"wrap",gap:12,marginBottom:16}}>
                    {[{label:"WBGT",res:regressionStats.wbgt,color:"#f97316",unit:"°C"},{label:"Wet Bulb Temp",res:regressionStats.tw,color:"#38bdf8",unit:"°C"}]
                      .map((item,i)=>item.res&&(
                        <div key={i} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${item.color}44`,borderRadius:14,padding:"16px 18px",flex:"1 1 280px"}}>
                          <div style={{fontSize:10,color:item.color,letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>OLS Regression · {item.label}</div>
                          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                            {[
                              {label:"Slope",value:<span style={{color:item.res.slope>0?"#f87171":"#4ade80",fontWeight:700}}>{item.res.slope>0?"+":""}{item.res.slope} {item.unit}/yr</span>},
                              {label:"Intercept",value:item.res.intercept},
                              {label:"R²",value:<span style={{color:item.res.r2>0.5?"#4ade80":item.res.r2>0.25?"#facc15":"#f87171",fontWeight:700}}>{item.res.r2}</span>},
                              {label:"p-value",value:<span style={{color:item.res.pValue<0.05?"#4ade80":"#f87171",fontWeight:700}}>{pBadge(item.res.pValue).label}</span>},
                            ].map((s,j)=>(
                              <div key={j} style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:"8px 10px"}}>
                                <div style={{fontSize:10,color:"#4a8aaa",marginBottom:3}}>{s.label}</div>
                                <div style={{fontSize:15,fontWeight:700,color:"#e0f4ff"}}>{s.value}</div>
                              </div>
                            ))}
                          </div>
                          <div style={{marginTop:10,fontSize:11,color:"#94a3b8",fontStyle:"italic"}}>R² = {item.res.r2} → {(item.res.r2*100).toFixed(1)}% of variance explained by time trend.</div>
                        </div>
                      ))}
                  </div>
                  <div style={box}>
                    <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>Annual WBGT — Actual vs OLS Trend Line</div>
                    <DownloadBtn containerRef={refFig6} filename="Fig6_OLS_Regression.png" figLabel="OLS regression actual vs trend line" figNum="Fig. 6"/>
                    <div ref={refFig6} style={{background:th.bg, padding: paperMode?'8px':0, borderRadius:8}}>
                    <ResponsiveContainer width="100%" height={240}>
                      <ComposedChart data={regressionStats.trendLine} margin={{top:5,right:14,left:0,bottom:5}}>
                        <CartesianGrid stroke={th.grid} strokeDasharray={th.gridDash}/>
                        <XAxis dataKey="year" tick={{fill:th.axTick,fontSize:11}}/>
                        <YAxis tick={{fill:th.axTick,fontSize:9}} unit="°C" domain={["auto","auto"]}/>
                        <Tooltip content={<CT/>}/><Legend wrapperStyle={{color:th.subtext,fontSize:10}}/>
                        <Bar dataKey="actual" name="Actual WBGT" fill={paperMode?"#fb923c44":"#f9731644"} unit="°C"/>
                        <Line type="linear" dataKey="predicted" name="OLS Trend" stroke={th.refExtreme} strokeWidth={2.5} dot={false} unit="°C" strokeDasharray="6 3"/>
                      </ComposedChart>
                    </ResponsiveContainer>
                    </div>
                  </div>
                  <div style={{...box,background:"rgba(56,189,248,0.04)",border:"1px solid #0c4a6e"}}>
                    <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>📝 Ready-to-Paste Regression Text</div>
                    <div style={{fontSize:12,color:"#cbd5e1",lineHeight:1.9,background:"rgba(0,0,0,0.3)",borderRadius:10,padding:"14px 16px"}}>
                      OLS regression of annual mean WBGT against time confirmed a statistically {regressionStats.wbgt.pValue<0.05?"significant":"non-significant"} increasing trend
                      (slope = <b style={{color:"#f97316"}}>{regressionStats.wbgt.slope>0?"+":""}{regressionStats.wbgt.slope}°C/yr</b>,
                      R² = <b style={{color:"#7dd3fc"}}>{regressionStats.wbgt.r2}</b>,
                      p = <b style={{color:regressionStats.wbgt.pValue<0.05?"#4ade80":"#f87171"}}>{regressionStats.wbgt.pValue}</b>),
                      accounting for {(regressionStats.wbgt.r2*100).toFixed(1)}% of interannual WBGT variance.
                      Wet bulb temperature showed slope = <b style={{color:"#38bdf8"}}>{regressionStats.tw?.slope>0?"+":""}{regressionStats.tw?.slope}°C/yr</b> (R² = <b style={{color:"#7dd3fc"}}>{regressionStats.tw?.r2}</b>). These results corroborate the non-parametric Mann-Kendall analysis.
                    </div>
                  </div>
                </>
              ):(
                <div style={{textAlign:"center",color:"#4a8aaa",padding:"30px 0"}}>Need ≥ 3 years of data.</div>
              )}
            </>
          )}

          {/* ══ EXCEEDANCE ══ */}
          {tab==="exceedance"&&(
            <>
              <div style={box}>
                <div style={{fontSize:9,color:"#ef4444",letterSpacing:3,textTransform:"uppercase",marginBottom:12}}>📈 Exceedance Frequency — Overall Period</div>
                <div style={{display:"flex",flexDirection:"column",gap:8}}>
                  {exceedance.overall.map((e,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:12}}>
                      <div style={{width:130,fontSize:11,color:"#94a3b8",flexShrink:0}}>{e.label}</div>
                      <div style={{flex:1,height:22,background:"rgba(255,255,255,0.05)",borderRadius:5,overflow:"hidden"}}>
                        <div style={{width:`${Math.min(e.pct,100)}%`,height:"100%",background:e.color,borderRadius:5,opacity:0.85}}/>
                      </div>
                      <div style={{width:80,fontSize:12,color:e.color,fontWeight:700,textAlign:"right"}}>{e.total.toLocaleString()} days</div>
                      <div style={{width:50,fontSize:11,color:"#64a7c8",textAlign:"right"}}>{e.pct}%</div>
                    </div>
                  ))}
                </div>
              </div>
              <div style={box}>
                <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>Exceedance Days Per Year</div>
                <DownloadBtn containerRef={refFig5} filename="Fig5_Exceedance_Frequency.png" figLabel="Exceedance frequency bar chart" figNum="Fig. 5"/>
                <div ref={refFig5} style={{background:th.bg, padding: paperMode?'8px':0, borderRadius:8}}>
                <ResponsiveContainer width="100%" height={230}>
                  <BarChart data={exceedance.byYear} margin={{top:5,right:14,left:0,bottom:5}}>
                    <CartesianGrid stroke={th.grid} strokeDasharray={th.gridDash}/>
                    <XAxis dataKey="year" tick={{fill:th.axTick,fontSize:10}} interval={iv(exceedance.byYear.length)}/>
                    <YAxis tick={{fill:th.axTick,fontSize:9}} unit=" d"/>
                    <Tooltip content={<CT/>}/><Legend wrapperStyle={{color:th.subtext,fontSize:10}}/>
                    <Bar dataKey="Tw ≥ 28°C"   name="Tw ≥ 28°C"   fill="#fb923c" radius={[4,4,0,0]} unit=" days"/>
                    <Bar dataKey="Tw ≥ 32°C"   name="Tw ≥ 32°C"   fill={th.refDanger}  radius={[4,4,0,0]} unit=" days"/>
                    <Bar dataKey="WBGT ≥ 32°C" name="WBGT ≥ 32°C" fill={th.refExtreme} radius={[4,4,0,0]} unit=" days"/>
                  </BarChart>
                </ResponsiveContainer>
                </div>
              </div>
              <div style={{...box,background:"rgba(56,189,248,0.04)",border:"1px solid #0c4a6e"}}>
                <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>📝 Ready-to-Paste Exceedance Text</div>
                <div style={{fontSize:12,color:"#cbd5e1",lineHeight:1.9,background:"rgba(0,0,0,0.3)",borderRadius:10,padding:"14px 16px"}}>
                  {(()=>{
                    const tw28=exceedance.overall.find(e=>e.label==="Tw ≥ 28°C");
                    const tw32=exceedance.overall.find(e=>e.label==="Tw ≥ 32°C");
                    const tw35=exceedance.overall.find(e=>e.label==="Tw ≥ 35°C");
                    const wbgt32=exceedance.overall.find(e=>e.label==="WBGT ≥ 32°C");
                    const wbgt35=exceedance.overall.find(e=>e.label==="WBGT ≥ 35°C");
                    if(!tw28)return"Loading…";
                    return<>Exceedance frequency analysis revealed that wet bulb temperature exceeded 28°C on <b style={{color:"#fb923c"}}>{tw28.total.toLocaleString()} days ({tw28.pct}%)</b> of the study period. The Tw ≥ 32°C threshold was exceeded on <b style={{color:"#ef4444"}}>{tw32?.total.toLocaleString()} days ({tw32?.pct}%)</b>, and the physiologically lethal Tw ≥ 35°C threshold (Raymond et al., 2020) on <b style={{color:"#dc2626"}}>{tw35?.total.toLocaleString()} days ({tw35?.pct}%)</b>. WBGT ≥ 32°C (danger) was exceeded on <b style={{color:"#f97316"}}>{wbgt32?.total.toLocaleString()} days ({wbgt32?.pct}%)</b>, and WBGT ≥ 35°C (extreme) on <b style={{color:"#ef4444"}}>{wbgt35?.total.toLocaleString()} days ({wbgt35?.pct}%)</b>.</>;
                  })()}
                </div>
              </div>
            </>
          )}

          {/* ══ HEAT STROKE ══ */}
          {tab==="heatstroke"&&(
            <div style={box}>
              <div style={{paddingLeft:6,marginBottom:10}}>
                <div style={{fontSize:9,color:"#ef4444",letterSpacing:3,textTransform:"uppercase"}}>Heat Stroke Risk Index (WBGT)</div>
                <div style={{fontSize:14,fontWeight:700,color:"#e0f4ff",marginTop:2}}>{view==="monthly"?"Monthly Average WBGT + Danger Days":"Daily WBGT Values"}</div>
              </div>
              <DownloadBtn containerRef={refFig7} filename="Fig7_Monthly_WBGT_Risk.png" figLabel="Monthly WBGT heat stroke risk" figNum="Fig. 7"/>
              <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:12}}>
                {[["< 28°C","Safe","#4ade80"],["28–32°C","Caution","#facc15"],["32–35°C","Danger","#f97316"],["≥ 35°C","Extreme","#ef4444"]].map(([r,l,c],i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:5,background:"rgba(255,255,255,0.03)",border:`1px solid ${c}33`,borderRadius:7,padding:"4px 10px"}}>
                    <div style={{width:8,height:8,borderRadius:"50%",background:c}}/><span style={{color:c,fontWeight:600,fontSize:10}}>{l}</span><span style={{color:"#4a8aaa",fontSize:9}}>{r}</span>
                  </div>
                ))}
              </div>
              <div ref={refFig7} style={{background:th.bg, padding: paperMode?'8px':0, borderRadius:8}}>
              <ResponsiveContainer width="100%" height={290}>
                {view==="monthly"?(
                  <ComposedChart data={monthlyData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <CartesianGrid stroke={th.grid} strokeDasharray={th.gridDash}/>
                    <XAxis dataKey="label" tick={{fill:th.axTick,fontSize:9}} interval={iv(monthlyData.length)}/>
                    <YAxis yAxisId="l" tick={{fill:th.axTick,fontSize:9}} unit="°C" domain={[20,40]}/>
                    <YAxis yAxisId="r" orientation="right" tick={{fill:th.axTick,fontSize:9}} unit=" d"/>
                    <Tooltip content={<CT/>}/><Legend wrapperStyle={{color:th.subtext,fontSize:10}}/>
                    <ReferenceLine yAxisId="l" y={35} stroke={th.refExtreme} strokeDasharray="4 2" label={{value:"Extreme ≥35°C",fill:th.refExtreme,fontSize:9,position:"insideTopLeft"}}/>
                    <ReferenceLine yAxisId="l" y={32} stroke={th.refDanger}  strokeDasharray="4 2" label={{value:"Danger ≥32°C", fill:th.refDanger, fontSize:9,position:"insideTopLeft"}}/>
                    <ReferenceLine yAxisId="l" y={28} stroke={th.refCaution} strokeDasharray="4 2"/>
                    <Area yAxisId="l" type="monotone" dataKey="wbgt" name="Avg WBGT" stroke={th.lineWbgt} fill={th.areaFill} strokeWidth={2} dot={false} unit="°C"/>
                    <Bar  yAxisId="r" dataKey="dangerDays" name="Danger Days" fill={paperMode?"#c2410c44":"#ef444444"} radius={[3,3,0,0]} unit=" days"/>
                  </ComposedChart>
                ):(
                  <AreaChart data={chartData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <defs><linearGradient id="gW" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%"  stopColor={th.areaWbgt} stopOpacity={0.35}/>
                      <stop offset="95%" stopColor={th.areaWbgt} stopOpacity={0}/>
                    </linearGradient></defs>
                    <CartesianGrid stroke={th.grid} strokeDasharray={th.gridDash}/>
                    <XAxis dataKey="label" tick={{fill:th.axTick,fontSize:9}} interval={iv(filteredRows.length)}/>
                    <YAxis tick={{fill:th.axTick,fontSize:9}} unit="°C" domain={[18,42]}/>
                    <Tooltip content={<CT/>}/>
                    <ReferenceLine y={35} stroke={th.refExtreme} strokeDasharray="4 2"/>
                    <ReferenceLine y={32} stroke={th.refDanger}  strokeDasharray="4 2"/>
                    <ReferenceLine y={28} stroke={th.refCaution} strokeDasharray="4 2"/>
                    <Area type="monotone" dataKey="wbgt" name="WBGT" stroke={th.lineWbgt} fill="url(#gW)" strokeWidth={1.5} dot={false} unit="°C"/>
                  </AreaChart>
                )}
              </ResponsiveContainer>
              </div>
            </div>
          )}

          {/* ══ WET BULB ══ */}
          {tab==="wetbulb"&&(
            <div style={box}>
              <div style={{paddingLeft:6,marginBottom:10}}>
                <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase"}}>Wet Bulb Temperature (Tw) · Stull (2011)</div>
                <div style={{fontSize:14,fontWeight:700,color:"#e0f4ff",marginTop:2}}>{view==="monthly"?"Monthly Average Tw":"Daily Tw"} — Survival limit = 35°C</div>
              </div>
              <DownloadBtn containerRef={refFig8} filename="Fig8_Wet_Bulb_Temperature.png" figLabel="Wet bulb temperature time series" figNum="Fig. 8"/>
              <div ref={refFig8} style={{background:th.bg, padding: paperMode?'8px':0, borderRadius:8}}>
              <ResponsiveContainer width="100%" height={270}>
                <AreaChart data={chartData} margin={{top:5,right:14,left:0,bottom:5}}>
                  <defs><linearGradient id="gT" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%"  stopColor={th.lineTw} stopOpacity={0.35}/>
                    <stop offset="95%" stopColor={th.lineTw} stopOpacity={0}/>
                  </linearGradient></defs>
                  <CartesianGrid stroke={th.grid} strokeDasharray={th.gridDash}/>
                  <XAxis dataKey="label" tick={{fill:th.axTick,fontSize:9}} interval={iv(chartData.length)}/>
                  <YAxis tick={{fill:th.axTick,fontSize:9}} unit="°C" domain={[10,38]}/>
                  <Tooltip content={<CT/>}/>
                  <ReferenceLine y={35} stroke={th.refLethal} strokeWidth={2} strokeDasharray="5 3" label={{value:"☠ Lethal 35°C — Raymond et al. 2020",fill:th.refLethal,fontSize:9,position:"insideTopLeft"}}/>
                  <ReferenceLine y={32} stroke={th.refDanger}  strokeDasharray="4 2"/>
                  <ReferenceLine y={28} stroke={th.refCaution} strokeDasharray="4 2"/>
                  <Area type="monotone" dataKey="wetBulb" name="Wet Bulb Temp" stroke={th.lineTw} fill="url(#gT)" strokeWidth={2} dot={false} unit="°C"/>
                </AreaChart>
              </ResponsiveContainer>
              </div>
            </div>
          )}

          {/* ══ YEAR TREND ══ */}
          {tab==="trend"&&(
            <>
              <div style={box}>
                <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>Risk Days Per Year — Stacked Area</div>
                <DownloadBtn containerRef={refFig3} filename="Fig3_Annual_Danger_Days.png" figLabel="Annual stacked heat stress risk days" figNum="Fig. 3"/>
                <div ref={refFig3} style={{background:th.bg, padding: paperMode?'8px':0, borderRadius:8}}>
                <ResponsiveContainer width="100%" height={260}>
                  <AreaChart data={yearlyData} margin={{top:10,right:14,left:0,bottom:5}}>
                    <defs>
                      <linearGradient id="gCau" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%"  stopColor={th.caution} stopOpacity={0.85}/>
                        <stop offset="95%" stopColor={th.caution} stopOpacity={0.55}/>
                      </linearGradient>
                      <linearGradient id="gDan" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%"  stopColor={th.danger}  stopOpacity={0.90}/>
                        <stop offset="95%" stopColor={th.danger}  stopOpacity={0.60}/>
                      </linearGradient>
                      <linearGradient id="gExt" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%"  stopColor={th.extreme} stopOpacity={0.95}/>
                        <stop offset="95%" stopColor={th.extreme} stopOpacity={0.65}/>
                      </linearGradient>
                    </defs>
                    <CartesianGrid stroke={th.grid} strokeDasharray={th.gridDash}/>
                    <XAxis dataKey="year" tick={{fill:th.axTick,fontSize:11}} interval={iv(yearlyData.length)}/>
                    <YAxis tick={{fill:th.axTick,fontSize:9}} unit=" d"/>
                    <Tooltip content={<CT/>}/><Legend wrapperStyle={{color:th.subtext,fontSize:10}}/>
                    <Area type="monotone" dataKey="cautionDays" name="Caution 28–32°C" stackId="a"
                          stroke={th.caution} fill="url(#gCau)" strokeWidth={1.2} unit=" days"/>
                    <Area type="monotone" dataKey="dangerDays"  name="Danger 32–35°C"  stackId="a"
                          stroke={th.danger}  fill="url(#gDan)" strokeWidth={1.2} unit=" days"/>
                    <Area type="monotone" dataKey="extremeDays" name="Extreme ≥35°C"   stackId="a"
                          stroke={th.extreme} fill="url(#gExt)" strokeWidth={1.5} unit=" days"/>
                  </AreaChart>
                </ResponsiveContainer>
                </div>
              </div>
              <div style={box}>
                <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>Annual Avg Tw & WBGT</div>
                <DownloadBtn containerRef={refFig2} filename="Fig2_Annual_Trend_Lines.png" figLabel="Annual Tw, WBGT & Air Temp trend lines" figNum="Fig. 2"/>
                <div ref={refFig2} style={{background:th.bg, padding: paperMode?'8px':0, borderRadius:8}}>
                <ResponsiveContainer width="100%" height={210}>
                  <LineChart data={yearlyData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <CartesianGrid stroke={th.grid} strokeDasharray={th.gridDash}/>
                    <XAxis dataKey="year" tick={{fill:th.axTick,fontSize:11}} interval={iv(yearlyData.length)}/>
                    <YAxis tick={{fill:th.axTick,fontSize:9}} unit="°C" domain={["auto","auto"]}/>
                    <Tooltip content={<CT/>}/><Legend wrapperStyle={{color:th.subtext,fontSize:10}}/>
                    <ReferenceLine y={35} stroke={th.refLethal} strokeDasharray="4 2" label={{value:"Lethal 35°C",fill:th.refLethal,fontSize:9}}/>
                    <Line type="monotone" dataKey="avgTw"   name="Avg Wet Bulb" stroke={th.lineTw}   strokeWidth={2.5} dot={{fill:th.lineTw,r:3}} unit="°C"/>
                    <Line type="monotone" dataKey="avgWbgt" name="Avg WBGT"     stroke={th.lineWbgt} strokeWidth={2}   dot={{fill:th.lineWbgt,r:3}} unit="°C" strokeDasharray="5 3"/>
                    <Line type="monotone" dataKey="avgTemp" name="Avg Air Temp" stroke={th.lineTemp} strokeWidth={1.5} dot={{fill:th.lineTemp,r:3}} unit="°C" strokeDasharray="3 2"/>
                  </LineChart>
                </ResponsiveContainer>
                </div>
              </div>
            </>
          )}

          {/* ══ TEMPERATURE ══ */}
          {tab==="temp"&&(
            <div style={box}>
              <div style={{fontSize:9,color:"#fb923c",letterSpacing:3,textTransform:"uppercase",marginBottom:8}}>Surface Temperature</div>
              <ResponsiveContainer width="100%" height={260}>
                {view==="monthly"?(
                  <AreaChart data={monthlyData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <defs>
                      <linearGradient id="gMx" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#f97316" stopOpacity={0.3}/><stop offset="95%" stopColor="#f97316" stopOpacity={0}/></linearGradient>
                      <linearGradient id="gMn" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#38bdf8" stopOpacity={0.2}/><stop offset="95%" stopColor="#38bdf8" stopOpacity={0}/></linearGradient>
                    </defs>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{fill:"#4a8aaa",fontSize:9}} interval={iv(monthlyData.length)}/>
                    <YAxis tick={{fill:"#4a8aaa",fontSize:9}} unit="°C" domain={[10,42]}/>
                    <Tooltip content={<CT/>}/><Legend wrapperStyle={{color:"#94a3b8",fontSize:10}}/>
                    <Area type="monotone" dataKey="maxTemp"  name="Max Temp"  stroke="#f97316" fill="url(#gMx)" strokeWidth={2} dot={false} unit="°C"/>
                    <Area type="monotone" dataKey="meanTemp" name="Mean Temp" stroke="#facc15" fill="none"      strokeWidth={1.5} dot={false} unit="°C" strokeDasharray="4 2"/>
                    <Area type="monotone" dataKey="minTemp"  name="Min Temp"  stroke="#38bdf8" fill="url(#gMn)" strokeWidth={2} dot={false} unit="°C"/>
                  </AreaChart>
                ):(
                  <ComposedChart data={chartData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{fill:"#4a8aaa",fontSize:9}} interval={iv(filteredRows.length)}/>
                    <YAxis tick={{fill:"#4a8aaa",fontSize:9}} unit="°C" domain={[10,45]}/>
                    <Tooltip content={<CT/>}/><Legend wrapperStyle={{color:"#94a3b8",fontSize:10}}/>
                    <Area type="monotone" dataKey="maxTemp" name="Max Temp" stroke="#f97316" fill="#f9731615" strokeWidth={1.5} dot={false} unit="°C"/>
                    <Area type="monotone" dataKey="minTemp" name="Min Temp" stroke="#38bdf8" fill="#38bdf815" strokeWidth={1.5} dot={false} unit="°C"/>
                  </ComposedChart>
                )}
              </ResponsiveContainer>
            </div>
          )}

          {/* ══ HUMIDITY ══ */}
          {tab==="humidity"&&(
            <div style={box}>
              <div style={{fontSize:9,color:"#38bdf8",letterSpacing:3,textTransform:"uppercase",marginBottom:8}}>Relative Humidity</div>
              <ResponsiveContainer width="100%" height={250}>
                {view==="monthly"?(
                  <BarChart data={monthlyData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{fill:"#4a8aaa",fontSize:9}} interval={iv(monthlyData.length)}/>
                    <YAxis tick={{fill:"#4a8aaa",fontSize:9}} unit="%" domain={[40,100]}/>
                    <Tooltip content={<CT/>}/>
                    <ReferenceLine y={80} stroke="#7dd3fc" strokeDasharray="4 2" label={{value:"Monsoon ~80%",fill:"#7dd3fc",fontSize:9,position:"insideTopLeft"}}/>
                    <Bar dataKey="humidity" name="Avg Humidity" fill="#0ea5e9" radius={[4,4,0,0]} unit="%"/>
                  </BarChart>
                ):(
                  <AreaChart data={chartData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <defs><linearGradient id="gHm" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.4}/><stop offset="95%" stopColor="#0ea5e9" stopOpacity={0.02}/></linearGradient></defs>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{fill:"#4a8aaa",fontSize:9}} interval={iv(filteredRows.length)}/>
                    <YAxis tick={{fill:"#4a8aaa",fontSize:9}} unit="%" domain={[30,100]}/>
                    <Tooltip content={<CT/>}/>
                    <Area type="monotone" dataKey="humidity" name="Humidity" stroke="#0ea5e9" fill="url(#gHm)" strokeWidth={1.5} dot={false} unit="%"/>
                  </AreaChart>
                )}
              </ResponsiveContainer>
            </div>
          )}

          {/* ══ SOLAR ══ */}
          {tab==="solar"&&(
            <div style={box}>
              <div style={{fontSize:9,color:"#fbbf24",letterSpacing:3,textTransform:"uppercase",marginBottom:8}}>Solar Radiation (MJ/m²)</div>
              <ResponsiveContainer width="100%" height={240}>
                {view==="monthly"?(
                  <BarChart data={monthlyData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{fill:"#4a8aaa",fontSize:9}} interval={iv(monthlyData.length)}/>
                    <YAxis tick={{fill:"#4a8aaa",fontSize:9}} unit=" MJ" domain={[0,30]}/>
                    <Tooltip content={<CT/>}/>
                    <Bar dataKey="solar" name="Solar Radiation" fill="#f59e0b" radius={[4,4,0,0]} unit=" MJ/m²"/>
                  </BarChart>
                ):(
                  <AreaChart data={chartData} margin={{top:5,right:14,left:0,bottom:5}}>
                    <defs><linearGradient id="gSl" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#f59e0b" stopOpacity={0.4}/><stop offset="95%" stopColor="#f59e0b" stopOpacity={0.02}/></linearGradient></defs>
                    <CartesianGrid stroke="#0a1a28" strokeDasharray="3 3"/>
                    <XAxis dataKey="label" tick={{fill:"#4a8aaa",fontSize:9}} interval={iv(filteredRows.length)}/>
                    <YAxis tick={{fill:"#4a8aaa",fontSize:9}} unit=" MJ" domain={[0,35]}/>
                    <Tooltip content={<CT/>}/>
                    <Area type="monotone" dataKey="solar" name="Solar Radiation" stroke="#f59e0b" fill="url(#gSl)" strokeWidth={1.5} dot={false} unit=" MJ/m²"/>
                  </AreaChart>
                )}
              </ResponsiveContainer>
            </div>
          )}

          {/* CITATIONS */}
          <div style={{...box,background:"rgba(139,92,246,0.04)",border:"1px solid #3b1f6a"}}>
            <div style={{fontSize:9,color:"#a78bfa",letterSpacing:3,textTransform:"uppercase",marginBottom:10}}>📚 Paper Citations</div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {[
                ["ERA5","Hersbach et al. (2020). The ERA5 global reanalysis. Q. J. R. Meteorol. Soc., 146(730). doi:10.1002/qj.3803"],
                ["Wet Bulb","Stull, R. (2011). Wet-Bulb Temperature from RH and Air Temperature. J. Appl. Meteor. Climatol., 50, 2267–2269."],
                ["35°C Limit","Raymond et al. (2020). Heat and humidity too severe for human tolerance. Science Advances, 6(19). doi:10.1126/sciadv.aaw1838"],
                ["WBGT","Bernard & Pourmoghani (1999). Prediction of workplace WBGT. Appl. Occup. Environ. Hyg., 14(2), 126–134."],
                ["Mann-Kendall","Mann (1945). Nonparametric tests against trend. Econometrica, 13, 245–259. / Kendall (1975). Rank Correlation Methods."],
                ["Sen's Slope","Sen (1968). Estimates of regression coefficient based on Kendall's tau. JASA, 63(324), 1379–1389."],
                ["Pearson r","Pearson, K. (1895). Notes on regression and inheritance. Proc. R. Soc. London, 58, 240–242."],
              ].map(([ref,cite],i)=>(
                <div key={i} style={{fontSize:10,color:"#94a3b8",lineHeight:1.7}}>
                  <span style={{color:"#a78bfa",fontWeight:600}}>[{ref}]</span> {cite}
                </div>
              ))}
            </div>
          </div>
          <div style={{textAlign:"center",fontSize:9,color:"#1e3a4a",marginTop:4}}>
            ERA5 · ECMWF · Open-Meteo · {KHULNA_LAT}°N {KHULNA_LON}°E · {startDate} → {endDate} · {yearlyData.length} years analysed
          </div>
        </>)}
      </div>
    </div>
  );
}