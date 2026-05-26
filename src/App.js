import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  ReferenceLine, Area, AreaChart, ComposedChart, Line,
  BarChart, Bar, Cell, Scatter,
} from "recharts";

// ── Constants ─────────────────────────────────────────────────────
const LAT = 22.8098, LON = 89.5644;
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const ERA5_START = "2000-01-01";

// ── Physics ───────────────────────────────────────────────────────
function calcTw(T, RH) {
  if (T == null || RH == null) return null;
  return parseFloat((
    T * Math.atan(0.151977 * (RH + 8.313659) ** 0.5)
    + Math.atan(T + RH) - Math.atan(RH - 1.676331)
    + 0.00391838 * RH ** 1.5 * Math.atan(0.023101 * RH) - 4.686035
  ).toFixed(2));
}
function calcWBGT(T, RH) {
  if (T == null || RH == null) return null;
  const es = 6.112 * Math.exp(17.67 * T / (T + 243.5));
  return parseFloat((0.567 * T + 0.393 * (RH / 100) * es + 3.94).toFixed(2));
}

// ── Statistics ────────────────────────────────────────────────────
function avg(arr) {
  const v = arr.filter(x => x != null);
  return v.length ? parseFloat((v.reduce((a,b)=>a+b,0)/v.length).toFixed(2)) : null;
}
function normalCDF(z) {
  const t=1/(1+0.2316419*Math.abs(z));
  const p=t*(0.31938153+t*(-0.356563782+t*(1.781477937+t*(-1.821255978+t*1.330274429))));
  const c=1-Math.exp(-0.5*z*z)/Math.sqrt(2*Math.PI)*p;
  return z>=0?c:1-c;
}
function mannKendall(s) {
  const n=s.length; if(n<4) return null;
  let S=0;
  for(let i=0;i<n-1;i++)
    for(let j=i+1;j<n;j++){const d=s[j]-s[i];if(d>0)S++;else if(d<0)S--;}
  const vS=n*(n-1)*(2*n+5)/18;
  const Z=S>0?(S-1)/Math.sqrt(vS):S<0?(S+1)/Math.sqrt(vS):0;
  const p=2*(1-normalCDF(Math.abs(Z)));
  return{S,Z:+Z.toFixed(3),p:+p.toFixed(4)};
}
function modifiedMannKendall(s) {
  const n=s.length; if(n<4) return null;
  let S=0;
  for(let i=0;i<n-1;i++)
    for(let j=i+1;j<n;j++){const d=s[j]-s[i];if(d>0)S++;else if(d<0)S--;}
  const vS_std=n*(n-1)*(2*n+5)/18;
  const indexed=s.map((v,i)=>({v,i})).sort((a,b)=>a.v-b.v);
  const ranks=new Array(n);
  indexed.forEach((item,rank)=>{ranks[item.i]=rank+1;});
  const mRank=(n+1)/2;
  let denom=0;
  for(let i=0;i<n;i++) denom+=(ranks[i]-mRank)**2;
  let wSum=0;
  for(let lag=1;lag<n-1;lag++){
    let num=0;
    for(let i=0;i<n-lag;i++) num+=(ranks[i]-mRank)*(ranks[i+lag]-mRank);
    wSum+=(n-lag)*(n-lag-1)*(n-lag-2)*(denom>0?num/denom:0);
  }
  const nsRatio=Math.max(0.5,1+(2/(n*(n-1)*(n-2)))*wSum);
  const vS_mod=vS_std*nsRatio;
  const Z=S>0?(S-1)/Math.sqrt(vS_mod):S<0?(S+1)/Math.sqrt(vS_mod):0;
  const p=2*(1-normalCDF(Math.abs(Z)));
  return{S,Z:+Z.toFixed(3),p:+p.toFixed(4),nsRatio:+nsRatio.toFixed(3)};
}
function senSlope(s) {
  const sl=[];
  for(let i=0;i<s.length-1;i++)
    for(let j=i+1;j<s.length;j++) sl.push((s[j]-s[i])/(j-i));
  sl.sort((a,b)=>a-b);
  const m=Math.floor(sl.length/2);
  return+(sl.length%2===0?(sl[m-1]+sl[m])/2:sl[m]).toFixed(4);
}
function pearsonR(x,y) {
  const n=Math.min(x.length,y.length); if(n<3) return null;
  const mx=x.slice(0,n).reduce((a,b)=>a+b,0)/n,my=y.slice(0,n).reduce((a,b)=>a+b,0)/n;
  let num=0,dx2=0,dy2=0;
  for(let i=0;i<n;i++){num+=(x[i]-mx)*(y[i]-my);dx2+=(x[i]-mx)**2;dy2+=(y[i]-my)**2;}
  return dx2&&dy2?+(num/Math.sqrt(dx2*dy2)).toFixed(3):null;
}
function linReg(x,y) {
  const n=x.length; if(n<3) return null;
  const mx=x.reduce((a,b)=>a+b,0)/n,my=y.reduce((a,b)=>a+b,0)/n;
  let ssxy=0,ssxx=0;
  for(let i=0;i<n;i++){ssxy+=(x[i]-mx)*(y[i]-my);ssxx+=(x[i]-mx)**2;}
  const slope=ssxy/ssxx,intercept=my-slope*mx;
  const yp=x.map(xi=>slope*xi+intercept);
  const sse=y.reduce((a,yi,i)=>a+(yi-yp[i])**2,0);
  const sst=y.reduce((a,yi)=>a+(yi-my)**2,0);
  const r2=1-sse/sst;
  const se=Math.sqrt(sse/(n-2));
  const ci=x.map(xi=>1.96*se*Math.sqrt(1/n+(xi-mx)**2/ssxx));
  return{slope:+slope.toFixed(4),intercept:+intercept.toFixed(4),r2:+r2.toFixed(4),yp,ci};
}
function getSeason(mo){
  const m=parseInt(mo);
  if(m>=3&&m<=5)return"Pre-monsoon";
  if(m>=6&&m<=9)return"Monsoon";
  return"Post-monsoon";
}

// ── Data aggregation ──────────────────────────────────────────────
function byYear(rows){
  const y={};
  rows.forEach(r=>{
    const yr=r.date.slice(0,4);
    if(!y[yr])y[yr]={tw:[],wbgt:[],t:[],rh:[],c:0,d:0,e:0};
    if(r.tw!=null)y[yr].tw.push(r.tw);
    if(r.wbgt!=null)y[yr].wbgt.push(r.wbgt);
    if(r.t!=null)y[yr].t.push(r.t);
    if(r.rh!=null)y[yr].rh.push(r.rh);
    if(r.wbgt>=28&&r.wbgt<32)y[yr].c++;
    if(r.wbgt>=32&&r.wbgt<35)y[yr].d++;
    if(r.wbgt>=35)y[yr].e++;
  });
  return Object.entries(y).sort().map(([yr,v])=>({
    year:yr,yn:+yr,tw:avg(v.tw),wbgt:avg(v.wbgt),t:avg(v.t),rh:avg(v.rh),
    c:v.c,d:v.d,e:v.e,
  }));
}
function byMonth(rows){
  const m={};
  rows.forEach(r=>{
    const[yr,mo]=r.date.slice(0,7).split("-");
    const k=`${yr}-${mo}`;
    if(!m[k])m[k]={label:`${MONTHS[+mo-1]} '${yr.slice(2)}`,yr:+yr,mo:+mo,tw:[],wbgt:[]};
    if(r.tw!=null)m[k].tw.push(r.tw);
    if(r.wbgt!=null)m[k].wbgt.push(r.wbgt);
  });
  return Object.entries(m).sort().map(([,v])=>({...v,tw:avg(v.tw),wbgt:avg(v.wbgt)}));
}

// ── Figure download — 4× PNG with legend injection ───────────────
// Recharts renders Legend as an HTML <div> outside the SVG, so a plain
// SVG capture misses it. This function also parses the legend DOM items
// and injects them as native SVG elements into the clone, so every
// downloaded figure includes a complete, properly styled legend.
async function downloadFig(ref,filename){
  const c=ref.current; if(!c) throw new Error("Container not ready");
  let s=null,tries=30;
  while(tries--){
    const svgs=[...c.querySelectorAll('svg')].sort((a,b)=>{
      const wa=Math.max(a.getBoundingClientRect().width,+a.getAttribute('width')||0);
      const wb=Math.max(b.getBoundingClientRect().width,+b.getAttribute('width')||0);
      return wb-wa;
    });
    s=svgs[0]||null;
    if(s){
      const w=Math.max(s.getBoundingClientRect().width,+s.getAttribute('width')||0);
      const h=Math.max(s.getBoundingClientRect().height,+s.getAttribute('height')||0);
      if(w>=100&&h>=60){s._w=w;s._h=h;break;}
    }
    await new Promise(r=>setTimeout(r,200)); s=null;
  }
  if(!s) throw new Error("Chart not ready — try again in a moment");
  const{_w:w,_h:h}=s,SC=4;

  // ── Clone the chart SVG ───────────────────────────────────────
  const clone=s.cloneNode(true);
  clone.setAttribute('xmlns','http://www.w3.org/2000/svg');
  clone.setAttribute('width',w); clone.setAttribute('height',h);
  const st=document.createElementNS('http://www.w3.org/2000/svg','style');
  st.textContent=
    'text,tspan{font-family:"Times New Roman",Georgia,serif;}'+
    '.recharts-cartesian-axis-tick-value{font-size:13px;}'+
    '.recharts-label{font-size:12px;}';
  clone.insertBefore(st,clone.firstChild);

  // ── Inject legend as SVG elements ────────────────────────────
  // Recharts' legend is an HTML div (class recharts-legend-wrapper)
  // positioned absolutely inside recharts-wrapper, outside the SVG.
  // We read its items and draw them as SVG text + icons in the clone.
  const legendWrapper=c.querySelector('.recharts-legend-wrapper');
  if(legendWrapper){
    // Is the legend at the top or bottom of the chart?
    const lTop=parseFloat(legendWrapper.style.top||'0');
    const isTop=lTop<h*0.5;

    // Parse each visible legend item
    const items=[];
    legendWrapper.querySelectorAll('.recharts-legend-item').forEach(li=>{
      const span=li.querySelector('.recharts-legend-item-text');
      const text=span?.textContent?.trim()||'';
      if(!text||text.startsWith('_')) return; // skip hidden series

      // Detect icon type: area/bar → rect fill, line/scatter → stroke
      const iconSvg=li.querySelector('svg');
      const rectEl=iconSvg?.querySelector('rect[fill]');
      const lineEl=iconSvg?.querySelector('path[stroke],line[stroke]');
      const isArea=!!rectEl&&!lineEl;
      const color=rectEl?.getAttribute('fill')
        ||lineEl?.getAttribute('stroke')
        ||span?.style?.color
        ||'#666';
      items.push({text,color,isArea});
    });

    if(items.length>0){
      const NS='http://www.w3.org/2000/svg';
      const g=document.createElementNS(NS,'g');
      g.setAttribute('class','injected-legend');

      // Centre the legend row horizontally, position at top or bottom
      // Measure approximate total width first
      const CHAR_W=7.5, ICON_W=24, GAP=14;
      const totalW=items.reduce((acc,{text})=>acc+ICON_W+text.length*CHAR_W+GAP,0)-GAP;
      let lx=Math.max(16,(w-totalW)/2);
      const ly=isTop?14:h-10;

      items.forEach(({text,color,isArea})=>{
        if(isArea){
          // Filled rectangle icon (area / stacked-area series)
          const rect=document.createElementNS(NS,'rect');
          rect.setAttribute('x',String(lx));
          rect.setAttribute('y',String(ly-7));
          rect.setAttribute('width','18');
          rect.setAttribute('height','10');
          rect.setAttribute('fill',color);
          rect.setAttribute('opacity','0.85');
          g.appendChild(rect);
        } else {
          // Line + dot icon
          const ln=document.createElementNS(NS,'line');
          ln.setAttribute('x1',String(lx));    ln.setAttribute('y1',String(ly));
          ln.setAttribute('x2',String(lx+18)); ln.setAttribute('y2',String(ly));
          ln.setAttribute('stroke',color);     ln.setAttribute('stroke-width','2');
          g.appendChild(ln);
          const dot=document.createElementNS(NS,'circle');
          dot.setAttribute('cx',String(lx+9)); dot.setAttribute('cy',String(ly));
          dot.setAttribute('r','3.5');          dot.setAttribute('fill',color);
          g.appendChild(dot);
        }
        // Label text
        const txt=document.createElementNS(NS,'text');
        txt.setAttribute('x',String(lx+ICON_W));
        txt.setAttribute('y',String(ly+4));
        txt.setAttribute('fill','#374151');
        txt.setAttribute('font-size','12');
        txt.setAttribute('font-family','"Times New Roman",serif');
        txt.textContent=text;
        g.appendChild(txt);

        lx+=ICON_W+text.length*CHAR_W+GAP;
      });
      clone.appendChild(g);
    }
  }

  // ── Render clone to 4× PNG canvas ────────────────────────────
  const blob=new Blob([new XMLSerializer().serializeToString(clone)],
    {type:'image/svg+xml;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  await new Promise((res,rej)=>{
    const img=new Image();
    img.onload=()=>{
      const cv=document.createElement('canvas');
      cv.width=Math.round(w*SC); cv.height=Math.round(h*SC);
      const ctx=cv.getContext('2d');
      ctx.fillStyle='#ffffff'; ctx.fillRect(0,0,cv.width,cv.height);
      ctx.scale(SC,SC); ctx.drawImage(img,0,0,w,h);
      URL.revokeObjectURL(url);
      cv.toBlob(b=>{
        if(!b) return rej(new Error('Export failed'));
        const a=document.createElement('a');
        a.href=URL.createObjectURL(b); a.download=filename;
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        setTimeout(()=>URL.revokeObjectURL(a.href),1000); res();
      },'image/png',1.0);
    };
    img.onerror=()=>{URL.revokeObjectURL(url);rej(new Error('SVG render failed'));};
    img.src=url;
  });
}

// ── WBGT cell colour ──────────────────────────────────────────────
function wbgtColor(v){
  if(v==null)return'#dbeafe';
  if(v<22)return'#dbeafe';if(v<24)return'#bfdbfe';
  if(v<26)return'#a7f3d0';if(v<28)return'#6ee7b7';
  if(v<30)return'#fde68a';if(v<32)return'#fbbf24';
  if(v<34)return'#fb923c';if(v<35)return'#f97316';
  return'#dc2626';
}

// ── Design tokens ─────────────────────────────────────────────────
const C={
  bg:'#080d18',surface:'#0f172a',border:'#1e293b',hover:'#162032',
  accent:'#60a5fa',text:'#f1f5f9',sub:'#94a3b8',muted:'#475569',
  caution:'#fbbf24',danger:'#f97316',extreme:'#ef4444',
  tw:'#38bdf8',wbgt:'#fb923c',temp:'#fcd34d',grid:'#1e293b',
};

// ── Custom reference-line label — staggered to prevent overlap ────
// Used for Fig 4 where 3 vertical reference lines would otherwise clash.
// viewBox.x = pixel x of the reference line; place text to its LEFT at
// a different yOff for each line so they never overlap.
function RefLineLabel({viewBox,value,fill,yOff}){
  if(!viewBox) return null;
  return(
    <text x={viewBox.x-5} y={(viewBox.y||0)+yOff}
      fill={fill} fontSize={9} fontStyle="italic" textAnchor="end">
      {value}
    </text>
  );
}

// ── Small UI components ───────────────────────────────────────────
function ChartTip({active,payload,label,pm}){
  if(!active||!payload?.length) return null;
  const vis=payload.filter(p=>p.name&&!p.name.startsWith('_'));
  return(
    <div style={{background:pm?'rgba(255,255,255,0.98)':'rgba(8,13,24,0.97)',
      border:`1px solid ${pm?'#d1d5db':'#334155'}`,borderRadius:10,
      padding:'10px 16px',fontSize:12,color:pm?'#111':C.text}}>
      <div style={{color:C.accent,fontWeight:700,marginBottom:6,fontSize:11}}>{label}</div>
      {vis.map((p,i)=>(
        <div key={i} style={{display:'flex',justifyContent:'space-between',gap:20,marginBottom:2}}>
          <span style={{color:pm?'#6b7280':C.sub}}>{p.name}</span>
          <span style={{fontWeight:700,color:p.color}}>
            {typeof p.value==='number'?p.value.toFixed(2):'—'}{p.unit||''}
          </span>
        </div>
      ))}
    </div>
  );
}
function DlBtn({containerRef,filename}){
  const[st,setSt]=useState('idle');
  const go=async()=>{
    setSt('busy');
    try{await downloadFig(containerRef,filename);setSt('done');setTimeout(()=>setSt('idle'),2500);}
    catch(e){alert(e.message);setSt('idle');}
  };
  const col=st==='done'?'#34d399':C.accent;
  return(
    <button onClick={go} disabled={st==='busy'} style={{padding:'5px 14px',borderRadius:8,
      fontSize:11,fontFamily:'inherit',fontWeight:600,cursor:st==='busy'?'wait':'pointer',
      border:`1px solid ${col}`,background:`${col}18`,color:col}}>
      {st==='busy'?'⏳ Saving…':st==='done'?'✓ Saved':'⬇ PNG (300 dpi)'}
    </button>
  );
}
function StatCard({label,value,unit,icon,color,sub}){
  return(
    <div style={{background:C.surface,border:`1px solid ${C.border}`,
      borderLeft:`3px solid ${color||C.accent}`,borderRadius:12,
      padding:'14px 18px',flex:'1 1 130px',minWidth:130}}>
      <div style={{fontSize:20,marginBottom:4}}>{icon}</div>
      <div style={{fontSize:22,fontWeight:700,color:color||C.text}}>
        {value}<span style={{fontSize:12,fontWeight:400,marginLeft:3,color:C.sub}}>{unit}</span>
      </div>
      <div style={{fontSize:10,color:C.sub,marginTop:4,fontWeight:600,
        letterSpacing:0.5,textTransform:'uppercase'}}>{label}</div>
      {sub&&<div style={{fontSize:9,color:C.muted,marginTop:2}}>{sub}</div>}
    </div>
  );
}
function FigSection({title,subtitle,figNum,figRef,filename,children,pm}){
  return(
    <div style={{background:pm?'#ffffff':C.surface,
      border:`1px solid ${pm?'#e5e7eb':C.border}`,borderRadius:16,
      padding:'20px',marginBottom:20}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',
        flexWrap:'wrap',gap:10,marginBottom:14}}>
        <div>
          {figNum&&<div style={{fontSize:9,color:C.accent,letterSpacing:3,
            textTransform:'uppercase',fontWeight:700,marginBottom:3}}>{figNum}</div>}
          <div style={{fontSize:14,fontWeight:700,color:pm?'#111':C.text}}>{title}</div>
          {subtitle&&<div style={{fontSize:11,color:C.sub,marginTop:2}}>{subtitle}</div>}
        </div>
        {figRef&&<DlBtn containerRef={figRef} filename={filename}/>}
      </div>
      {children}
    </div>
  );
}

// ── Fig 7: Canvas heatmap ─────────────────────────────────────────
function HeatmapCanvas({monthly}){
  const canvasRef=useRef(null);
  const years=useMemo(()=>[...new Set(monthly.map(m=>m.yr))].sort(),[monthly]);
  useEffect(()=>{
    const cv=canvasRef.current; if(!cv||!years.length) return;
    const SC=2,CW=50+years.length*31,CH=32+12*29+60;
    cv.width=CW*SC;cv.height=CH*SC;
    cv.style.width=CW+'px';cv.style.height=CH+'px';
    const ctx=cv.getContext('2d');ctx.scale(SC,SC);
    ctx.fillStyle='#ffffff';ctx.fillRect(0,0,CW,CH);
    const ax='#374151';
    // Year labels
    ctx.font='600 11px Inter,sans-serif';ctx.fillStyle=ax;ctx.textAlign='center';
    years.forEach((yr,xi)=>{
      if(yr%5===0)ctx.fillText(String(yr),50+xi*31+14,24);
    });
    // Month rows
    MONTHS.forEach((mn,mi)=>{
      const y=32+mi*29;
      ctx.font='700 11px Inter,sans-serif';ctx.fillStyle=ax;ctx.textAlign='right';
      ctx.fillText(mn,44,y+18);
      years.forEach((yr,xi)=>{
        const row=monthly.find(m=>m.yr===yr&&m.mo===mi+1);
        const x=50+xi*31;
        ctx.fillStyle=wbgtColor(row?.wbgt??null);
        ctx.beginPath();ctx.roundRect(x,y,28,26,3);ctx.fill();
      });
    });
    // Legend
    const legY=32+12*29+8;
    const stops=[[22,'<22'],[24,'22–24'],[26,'24–26'],[28,'26–28'],[30,'28–30'],
                 [32,'30–32'],[34,'32–34'],[35,'34–35'],[36,'≥35']];
    ctx.font='600 11px Inter,sans-serif';ctx.fillStyle=ax;ctx.textAlign='left';
    ctx.fillText('WBGT (°C):',50,legY+13);
    let lx=50+90;
    stops.forEach(([v,l])=>{
      ctx.fillStyle=wbgtColor(v);ctx.beginPath();ctx.roundRect(lx,legY,26,13,3);ctx.fill();
      ctx.fillStyle=ax;ctx.font='500 10px Inter,sans-serif';ctx.textAlign='left';
      ctx.fillText(l,lx+28,legY+10);lx+=28+ctx.measureText(l).width+8;
    });
  },[monthly,years]);
  const download=()=>{
    canvasRef.current?.toBlob(b=>{
      if(!b)return;
      const a=document.createElement('a');
      a.href=URL.createObjectURL(b);a.download='Fig7_Monthly_WBGT_Risk.png';
      document.body.appendChild(a);a.click();document.body.removeChild(a);
      setTimeout(()=>URL.revokeObjectURL(a.href),1000);
    },'image/png',1.0);
  };
  return(
    <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:16,padding:'20px',marginBottom:20}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:14}}>
        <div>
          <div style={{fontSize:9,color:C.accent,letterSpacing:3,textTransform:'uppercase',fontWeight:700,marginBottom:3}}>Fig. 7</div>
          <div style={{fontSize:14,fontWeight:700,color:C.text}}>Monthly WBGT Calendar Heatmap</div>
          <div style={{fontSize:11,color:C.sub,marginTop:2}}>
            Indoor/shaded WBGT · colour-coded by risk category · rows = months, columns = years
          </div>
        </div>
        <button onClick={download} style={{padding:'5px 14px',borderRadius:8,fontSize:11,
          fontFamily:'inherit',fontWeight:600,cursor:'pointer',
          border:`1px solid ${C.accent}`,background:`${C.accent}18`,color:C.accent}}>
          ⬇ PNG (300 dpi)
        </button>
      </div>
      <div style={{overflowX:'auto'}}>
        <canvas ref={canvasRef} style={{display:'block',maxWidth:'100%',borderRadius:8}}/>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
export default function App(){
  const today=new Date().toISOString().slice(0,10);
  const[tab,setTab]=useState('overview');
  const[startDate,setStartDate]=useState('2000-01-01');
  const[endDate,setEndDate]=useState('2025-12-31');
  const[rawRows,setRawRows]=useState(null);
  const[loading,setLoading]=useState(false);
  const[err,setErr]=useState(null);
  const[pm,setPm]=useState(false);

  const r2=useRef(null),r3=useRef(null),r4=useRef(null),
        r5=useRef(null),r6=useRef(null),r8=useRef(null);

  // ── Fetch ────────────────────────────────────────────────────────
  const fetchData=useCallback(async(s,e)=>{
    setLoading(true);setErr(null);setRawRows(null);
    try{
      const url=
        'https://archive-api.open-meteo.com/v1/archive'+
        '?latitude='+LAT+'&longitude='+LON+
        '&start_date='+s+'&end_date='+e+
        '&daily=temperature_2m_max,temperature_2m_min,temperature_2m_mean,'+
        'relative_humidity_2m_mean'+
        '&timezone=Asia%2FDhaka';
      const res=await fetch(url);
      if(!res.ok) throw new Error('HTTP '+res.status+' — '+res.statusText);
      const j=await res.json();
      if(j.error) throw new Error(j.reason||'API returned error');
      setRawRows(j.daily.time.map((d,i)=>{
        const T=j.daily.temperature_2m_mean[i],RH=j.daily.relative_humidity_2m_mean[i];
        return{date:d,tmax:j.daily.temperature_2m_max[i],tmin:j.daily.temperature_2m_min[i],
          t:T,rh:RH,tw:calcTw(T,RH),wbgt:calcWBGT(T,RH)};
      }));
    }catch(e){setErr(e.message);}
    finally{setLoading(false);}
  },[]);

  useEffect(()=>{fetchData(startDate,endDate);},[]);

  const data=useMemo(()=>
    (rawRows||[]).filter(r=>r.date>=startDate&&r.date<=endDate),
    [rawRows,startDate,endDate]);

  const Tip=useCallback(({active,payload,label})=>
    <ChartTip active={active} payload={payload} label={label} pm={pm}/>,[pm]);

  const yearly=useMemo(()=>byYear(data),[data]);
  const monthly=useMemo(()=>byMonth(data),[data]);

  const mk=useMemo(()=>{
    if(yearly.length<4) return null;
    const tw=yearly.map(y=>y.tw).filter(Boolean);
    const wbgt=yearly.map(y=>y.wbgt).filter(Boolean);
    const t=yearly.map(y=>y.t).filter(Boolean);
    const d=yearly.map(y=>y.d);
    const run=arr=>({std:mannKendall(arr),mod:modifiedMannKendall(arr),slope:senSlope(arr)});
    return{tw:run(tw),wbgt:run(wbgt),t:run(t),danger:run(d)};
  },[yearly]);

  const reg=useMemo(()=>{
    const vW=yearly.filter(y=>y.wbgt!=null);
    const vTw=yearly.filter(y=>y.tw!=null);
    const rW=linReg(vW.map(y=>y.yn),vW.map(y=>y.wbgt));
    const rTw=linReg(vTw.map(y=>y.yn),vTw.map(y=>y.tw));
    if(!rW) return null;
    const chart=vW.map((yr,i)=>({
      year:yr.year,wbgt:yr.wbgt,
      pred:+rW.yp[i].toFixed(3),
      ciHi:+(rW.yp[i]+rW.ci[i]).toFixed(3),
      ciLo:+(rW.yp[i]-rW.ci[i]).toFixed(3),
    }));
    const all=chart.flatMap(d=>[d.wbgt,d.ciHi,d.ciLo]).filter(v=>v!=null&&!isNaN(v));
    return{wbgt:rW,tw:rTw,chart,
      domain:[Math.floor((Math.min(...all)-0.2)*10)/10,
              Math.ceil( (Math.max(...all)+0.2)*10)/10]};
  },[yearly]);

  const seasonal=useMemo(()=>{
    const s={'Pre-monsoon':[],'Monsoon':[],'Post-monsoon':[]};
    data.forEach(r=>{if(r.wbgt!=null)s[getSeason(r.date.slice(5,7))].push(r);});
    return Object.entries(s).map(([name,rows])=>({
      name,
      avgWbgt:avg(rows.map(r=>r.wbgt)),avgTw:avg(rows.map(r=>r.tw)),
      avgT:avg(rows.map(r=>r.t)),avgRH:avg(rows.map(r=>r.rh)),
      c:rows.filter(r=>r.wbgt>=28&&r.wbgt<32).length,
      d:rows.filter(r=>r.wbgt>=32&&r.wbgt<35).length,
      e:rows.filter(r=>r.wbgt>=35).length,
    }));
  },[data]);

  const exceed=useMemo(()=>{
    const TH=[
      {label:'Tw ≥ 26°C',  field:'tw',  val:26,color:'#fbbf24'},
      {label:'Tw ≥ 28°C',  field:'tw',  val:28,color:'#fb923c'},
      {label:'Tw ≥ 30°C',  field:'tw',  val:30,color:'#f97316'},
      {label:'WBGT ≥ 28°C',field:'wbgt',val:28,color:'#a78bfa'},
      {label:'WBGT ≥ 32°C',field:'wbgt',val:32,color:'#ef4444'},
      {label:'WBGT ≥ 35°C',field:'wbgt',val:35,color:'#dc2626'},
    ];
    return TH.map(t=>({
      label:t.label,color:t.color,
      n:data.filter(r=>r[t.field]!=null&&r[t.field]>=t.val).length,
      pct:data.length?+(data.filter(r=>r[t.field]!=null&&r[t.field]>=t.val).length/data.length*100).toFixed(1):0,
    }));
  },[data]);

  const corr=useMemo(()=>{
    const p=data.filter(r=>r.t!=null&&r.rh!=null&&r.tw!=null&&r.wbgt!=null&&r.tmax!=null);
    const vars=['Tw','WBGT','T','RH','Tmax'];
    const d=[p.map(r=>r.tw),p.map(r=>r.wbgt),p.map(r=>r.t),p.map(r=>r.rh),p.map(r=>r.tmax)];
    return{vars,mat:vars.map((_,i)=>vars.map((__,j)=>pearsonR(d[i],d[j]))),n:p.length};
  },[data]);

  const totals=useMemo(()=>{
    const wbgts=data.map(r=>r.wbgt).filter(Boolean);
    const tws=data.map(r=>r.tw).filter(Boolean);
    return{n:data.length,meanTw:avg(tws),meanWbgt:avg(wbgts),
      maxTw:tws.length?+Math.max(...tws).toFixed(2):null,
      maxWbgt:wbgts.length?+Math.max(...wbgts).toFixed(2):null,
      c:data.filter(r=>r.wbgt>=28&&r.wbgt<32).length,
      d:data.filter(r=>r.wbgt>=32&&r.wbgt<35).length,
      e:data.filter(r=>r.wbgt>=35).length};
  },[data]);

  const dlExcel=()=>{
    try{
      const wb=XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ['Date','Tmax_C','Tmin_C','Tmean_C','RH_pct','Tw_C','WBGT_indoor_C'],
        ...data.map(r=>[r.date,r.tmax,r.tmin,r.t,r.rh,r.tw,r.wbgt])
      ]),'Daily');
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
        ['Year','Avg_Tw','Avg_WBGT','Avg_T','Avg_RH','Caution','Danger','Extreme'],
        ...yearly.map(y=>[y.year,y.tw,y.wbgt,y.t,y.rh,y.c,y.d,y.e])
      ]),'Yearly');
      XLSX.writeFile(wb,`Khulna_HeatStress_${startDate}_${endDate}.xlsx`);
    }catch(e){alert(e.message);}
  };

  const TABS=[
    {id:'overview',   icon:'📊',label:'Overview'         },
    {id:'mk',         icon:'📉',label:'MK Analysis'      },
    {id:'trends',     icon:'📈',label:'Fig 2–3: Trends'  },
    {id:'seasonal',   icon:'🌿',label:'Fig 4: Seasonal'  },
    {id:'exceedance', icon:'⚠️', label:'Fig 5: Exceedance'},
    {id:'regression', icon:'📐',label:'Fig 6: Regression'},
    {id:'heatmap',    icon:'🔥',label:'Fig 7: Heatmap'   },
    {id:'wetbulb',    icon:'💧',label:'Fig 8: Wet Bulb'  },
    {id:'correlation',icon:'🔗',label:'Correlations'     },
  ];

  const sigLabel=p=>p<0.01?'✓✓ p<0.01':p<0.05?'✓ p<0.05':p<0.1?'~ p<0.10':'✗ n.s.';
  const sigColor=p=>p<0.01?'#34d399':p<0.05?'#86efac':p<0.1?'#fbbf24':'#f87171';

  // ── Chart theme — font sizes scale with paper mode ────────────
  const tc={
    grid:pm?'#d1d5db':C.grid,ax:pm?'#1f2937':C.sub,
    ref35:pm?'#b91c1c':'#ef4444',ref32:pm?'#c2410c':'#f97316',ref28:pm?'#854d0e':'#eab308',
  };
  // Consistent professional font sizes: slightly larger in paper mode
  const FS={tick:pm?12:11, axLabel:pm?13:12, legend:pm?12:11, refLabel:pm?10:9};
  const atick={fill:tc.ax,fontSize:FS.tick};
  const legStyle={color:tc.ax,fontSize:FS.legend};

  // Y-axis label helper — rotated, inside left
  const yLab=(val,extra={})=>({
    value:val,angle:-90,position:'insideLeft',fill:tc.ax,fontSize:FS.axLabel,offset:14,...extra
  });
  // X-axis label helper — bottom
  const xLab=(val)=>({value:val,position:'insideBottom',offset:-10,fill:tc.ax,fontSize:FS.axLabel});

  const chartBg={background:pm?'#ffffff':'transparent',padding:pm?'20px 12px 8px':0};

  // ── Loading screen ──────────────────────────────────────────────
  if(loading) return(
    <div style={{minHeight:'100vh',background:C.bg,display:'flex',alignItems:'center',
      justifyContent:'center',flexDirection:'column',gap:16}}>
      <div style={{fontSize:48}}>🌀</div>
      <div style={{color:C.sub,letterSpacing:2,fontSize:13}}>Fetching ERA5 data…</div>
      <div style={{color:C.muted,fontSize:11}}>{startDate} → {endDate}</div>
    </div>
  );

  // ── Error screen ────────────────────────────────────────────────
  if(err&&!data.length) return(
    <div style={{minHeight:'100vh',background:C.bg,display:'flex',alignItems:'center',
      justifyContent:'center',flexDirection:'column',gap:20,padding:40}}>
      <div style={{fontSize:48}}>⚠️</div>
      <div style={{fontSize:16,fontWeight:700,color:'#ef4444'}}>Fetch failed</div>
      <div style={{background:'rgba(239,68,68,0.1)',border:'1px solid #ef4444',borderRadius:12,
        padding:'16px 24px',maxWidth:540,fontSize:13,color:'#fca5a5',lineHeight:1.7}}>
        <code>{err}</code>
      </div>
      <div style={{fontSize:12,color:C.sub,maxWidth:540,textAlign:'center',lineHeight:1.8}}>
        <strong style={{color:C.text}}>Fix:</strong> run this app locally with{' '}
        <code style={{color:'#34d399'}}>npm start</code>.
        The Open-Meteo API is accessible from a local dev server.
        Sandboxed browser environments (like Claude.ai artifacts) block external fetch calls.
      </div>
      <div style={{display:'flex',gap:12,flexWrap:'wrap',justifyContent:'center'}}>
        <button onClick={()=>fetchData(startDate,endDate)} style={{padding:'10px 24px',borderRadius:10,
          fontSize:13,fontFamily:'inherit',fontWeight:600,cursor:'pointer',
          border:'1px solid '+C.accent,background:'rgba(96,165,250,0.12)',color:C.accent}}>
          ↻ Retry
        </button>
        <a href={
          'https://archive-api.open-meteo.com/v1/archive?latitude='+LAT+'&longitude='+LON+
          '&start_date='+startDate+'&end_date='+endDate+
          '&daily=temperature_2m_max,temperature_2m_min,temperature_2m_mean,relative_humidity_2m_mean'+
          '&timezone=Asia%2FDhaka'
        } target="_blank" rel="noreferrer" style={{padding:'10px 24px',borderRadius:10,
          fontSize:13,fontFamily:'inherit',fontWeight:600,cursor:'pointer',textDecoration:'none',
          border:'1px solid '+C.muted,background:'transparent',color:C.muted}}>
          🔗 Test API URL
        </a>
      </div>
    </div>
  );

  return(
    <div style={{minHeight:'100vh',background:pm?'#f3f4f6':C.bg,
      fontFamily:"'Inter',system-ui,sans-serif",color:pm?'#111827':C.text,paddingBottom:60}}>

      {/* ── HEADER ────────────────────────────────────────────── */}
      <div style={{background:pm?'rgba(243,244,246,0.97)':'rgba(8,13,24,0.97)',
        backdropFilter:'blur(12px)',borderBottom:`1px solid ${pm?'#e5e7eb':C.border}`,
        position:'sticky',top:0,zIndex:50}}>
        <div style={{maxWidth:1200,margin:'0 auto',padding:'0 20px'}}>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',
            flexWrap:'wrap',gap:10,padding:'12px 0 8px'}}>
            <div>
              <div style={{fontSize:8,color:C.accent,letterSpacing:4,textTransform:'uppercase',marginBottom:2}}>
                ERA5 · Khulna 22.81°N 89.56°E · WBGT indoor (Bernard 1999)
              </div>
              <h1 style={{margin:0,fontSize:16,fontWeight:800,color:pm?'#111':C.text}}>
                🌡 Khulna Heat Stress Analysis
              </h1>
            </div>
            <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center'}}>
              <button onClick={()=>setPm(p=>!p)} style={{padding:'6px 14px',borderRadius:8,
                fontSize:11,fontFamily:'inherit',fontWeight:600,cursor:'pointer',
                border:`1px solid ${pm?'#22c55e':C.border}`,
                background:pm?'rgba(34,197,94,0.12)':'transparent',
                color:pm?'#22c55e':C.sub}}>
                {pm?'🟢 Paper Mode ON':'📄 Paper Mode'}
              </button>
              <button onClick={dlExcel} disabled={!data.length} style={{padding:'6px 14px',
                borderRadius:8,fontSize:11,fontFamily:'inherit',fontWeight:600,cursor:'pointer',
                border:`1px solid ${C.border}`,background:'transparent',
                color:C.sub,opacity:!data.length?0.4:1}}>
                📥 Excel
              </button>
              <button onClick={()=>fetchData(startDate,endDate)} style={{padding:'6px 14px',
                borderRadius:8,fontSize:11,fontFamily:'inherit',fontWeight:600,cursor:'pointer',
                border:`1px solid ${C.accent}`,background:'rgba(96,165,250,0.1)',color:C.accent}}>
                ↻ Refresh
              </button>
            </div>
          </div>
          <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap',
            paddingBottom:8,borderTop:`1px solid ${pm?'#e5e7eb':C.border}`,paddingTop:8}}>
            <span style={{fontSize:11,color:C.sub}}>Period:</span>
            {[['Start',startDate,v=>setStartDate(v),ERA5_START,endDate],
              ['End',endDate,v=>setEndDate(v),startDate,today]].map(([l,val,set,mn,mx])=>(
              <input key={l} type="date" value={val} min={mn} max={mx}
                onChange={e=>set(e.target.value)}
                style={{background:pm?'#fff':C.surface,border:`1px solid ${pm?'#d1d5db':C.border}`,
                  borderRadius:6,color:pm?'#111':C.text,padding:'4px 10px',fontSize:11,
                  colorScheme:pm?'light':'dark',outline:'none'}}/>
            ))}
            <button onClick={()=>fetchData(startDate,endDate)} style={{padding:'4px 12px',
              borderRadius:6,fontSize:11,fontFamily:'inherit',fontWeight:600,cursor:'pointer',
              border:`1px solid ${C.accent}`,background:'rgba(96,165,250,0.12)',color:C.accent}}>
              Apply
            </button>
            <span style={{fontSize:11,color:C.muted}}>
              {data.length.toLocaleString()} days · {yearly.length} years
            </span>
          </div>
          <div style={{display:'flex',overflowX:'auto',scrollbarWidth:'none',
            margin:'0 -20px',padding:'0 20px'}}>
            {TABS.map(t=>(
              <button key={t.id} onClick={()=>setTab(t.id)} style={{display:'flex',
                alignItems:'center',gap:4,padding:'8px 11px',border:'none',background:'none',
                cursor:'pointer',fontSize:11,fontFamily:'inherit',
                fontWeight:tab===t.id?700:400,whiteSpace:'nowrap',
                color:tab===t.id?C.accent:C.muted,
                borderBottom:`2px solid ${tab===t.id?C.accent:'transparent'}`,
                transition:'all 0.15s'}}>
                <span style={{fontSize:12}}>{t.icon}</span>{t.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {err&&data.length>0&&(
        <div style={{background:'rgba(239,68,68,0.08)',borderBottom:'1px solid rgba(239,68,68,0.3)',
          padding:'8px 20px',fontSize:11,color:'#f87171'}}>
          ⚠ Refresh failed: {err}
        </div>
      )}

      {data.length>0&&(
        <div style={{maxWidth:1200,margin:'0 auto',padding:'20px 20px 0'}}>

          {/* ══ OVERVIEW ════════════════════════════════════════ */}
          {tab==='overview'&&(
            <>
              <div style={{display:'flex',flexWrap:'wrap',gap:10,marginBottom:18}}>
                <StatCard icon="💧" label="Mean Wet Bulb" unit="°C" value={totals.meanTw??'—'} color={C.tw} sub="Stull (2011)"/>
                <StatCard icon="🌡" label="Mean WBGT"     unit="°C" value={totals.meanWbgt??'—'} color={C.wbgt} sub="Bernard (1999)"/>
                <StatCard icon="⬆"  label="Max WBGT"     unit="°C" value={totals.maxWbgt??'—'} color={C.extreme} sub="Record day"/>
                <StatCard icon="⬆"  label="Max Tw"       unit="°C" value={totals.maxTw??'—'} color={C.tw} sub="Record day"/>
                <StatCard icon="🟡" label="Caution days"  unit=""   value={totals.c.toLocaleString()} color={C.caution} sub="28–32°C"/>
                <StatCard icon="🟠" label="Danger days"   unit=""   value={totals.d.toLocaleString()} color={C.danger}  sub="32–35°C"/>
                <StatCard icon="🔴" label="Extreme days"  unit=""   value={totals.e.toLocaleString()} color={C.extreme} sub="≥35°C"/>
              </div>
              {mk&&(
                <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:16,padding:'18px 20px'}}>
                  <div style={{fontSize:9,color:C.accent,letterSpacing:3,textTransform:'uppercase',fontWeight:700,marginBottom:14}}>
                    Mann-Kendall — Standard + Modified (Hamed & Rao 1998)
                  </div>
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(250px,1fr))',gap:10}}>
                    {[
                      {n:'Wet Bulb Temp (Tw)',col:C.tw,   u:'°C',...mk.tw    },
                      {n:'WBGT (indoor)',     col:C.wbgt, u:'°C',...mk.wbgt  },
                      {n:'Mean Air Temp',     col:C.temp, u:'°C',...mk.t     },
                      {n:'Danger Days ≥32°C', col:C.danger,u:' d',...mk.danger},
                    ].map((item,i)=>(
                      <div key={i} style={{background:C.hover,border:`1px solid ${C.border}`,
                        borderLeft:`3px solid ${item.col}`,borderRadius:10,padding:'12px'}}>
                        <div style={{fontSize:11,color:item.col,fontWeight:700,marginBottom:6}}>{item.n}</div>
                        <div style={{fontSize:10,color:C.sub,marginBottom:2}}>
                          Std: Z={item.std?.Z} · <span style={{color:sigColor(item.std?.p),fontWeight:700}}>{sigLabel(item.std?.p)}</span>
                        </div>
                        <div style={{fontSize:10,color:'#a78bfa',marginBottom:6}}>
                          Mod: Z={item.mod?.Z} · <span style={{color:sigColor(item.mod?.p),fontWeight:700}}>{sigLabel(item.mod?.p)}</span>
                          <span style={{color:C.muted,marginLeft:6}}>n/n*={item.mod?.nsRatio}</span>
                        </div>
                        <div style={{fontSize:13,fontWeight:700,color:item.slope>0?C.danger:C.tw}}>
                          {item.slope>0?'+':''}{item.slope} {item.u}/yr (Sen)
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </>
          )}

          {/* ══ MK ANALYSIS ═════════════════════════════════════ */}
          {tab==='mk'&&mk&&(
            <>
              <div style={{background:'rgba(167,139,250,0.08)',border:'1px solid rgba(167,139,250,0.25)',
                borderRadius:12,padding:'14px 18px',marginBottom:18,fontSize:11,color:C.sub,lineHeight:1.7}}>
                <span style={{color:'#a78bfa',fontWeight:700}}>Modified MK (Hamed & Rao 1998): </span>
                Adjusts Var(S) for serial autocorrelation via effective sample size ratio n/n*.
                n/n* &gt; 1 = positive autocorrelation; standard test was too liberal.
                Both tests significant = most defensible for publication.
              </div>
              <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:16,padding:'18px 20px'}}>
                <div style={{fontSize:9,color:C.accent,letterSpacing:3,textTransform:'uppercase',fontWeight:700,marginBottom:14}}>
                  Table 4 — Standard vs Modified MK
                </div>
                <div style={{overflowX:'auto'}}>
                  <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                    <thead>
                      <tr style={{borderBottom:`2px solid ${C.border}`}}>
                        {['Variable','S','Std Z','Std p','Mod Z','Mod p','n/n*',"Sen's slope",'Verdict'].map((h,i)=>(
                          <th key={i} style={{padding:'8px 10px',textAlign:'left',fontSize:10,color:C.muted,fontWeight:700}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {[
                        ['Tw (°C)',          mk.tw,    C.tw,   `+${mk.tw.slope}°C/yr`    ],
                        ['WBGT indoor (°C)', mk.wbgt,  C.wbgt, `+${mk.wbgt.slope}°C/yr`  ],
                        ['Air Temp (°C)',    mk.t,     C.temp, `+${mk.t.slope}°C/yr`      ],
                        ['Danger days',      mk.danger,C.danger,`+${mk.danger.slope} d/yr`],
                      ].map(([name,m,col,slope],i)=>{
                        const both=m.std?.p<0.05&&m.mod?.p<0.05;
                        return(
                          <tr key={i} style={{borderBottom:`1px solid ${C.grid}`,
                            background:i%2?'rgba(255,255,255,0.015)':'transparent'}}>
                            <td style={{padding:'7px 10px',fontWeight:600,color:col}}>{name}</td>
                            <td style={{padding:'7px 10px',color:C.sub}}>{m.std?.S}</td>
                            <td style={{padding:'7px 10px',color:C.sub}}>{m.std?.Z}</td>
                            <td style={{padding:'7px 10px'}}><span style={{color:sigColor(m.std?.p),fontWeight:700}}>{sigLabel(m.std?.p)}</span></td>
                            <td style={{padding:'7px 10px',color:C.sub}}>{m.mod?.Z}</td>
                            <td style={{padding:'7px 10px'}}><span style={{color:sigColor(m.mod?.p),fontWeight:700}}>{sigLabel(m.mod?.p)}</span></td>
                            <td style={{padding:'7px 10px',color:'#a78bfa'}}>{m.mod?.nsRatio}</td>
                            <td style={{padding:'7px 10px',color:C.danger,fontWeight:700}}>{slope}</td>
                            <td style={{padding:'7px 10px',fontSize:10,fontStyle:'italic',
                              color:both?'#34d399':m.mod?.p<0.1?'#fbbf24':'#f87171'}}>
                              {both?'Robust (both)':m.mod?.p<0.1?'Marginal':'Weakened'}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}

          {/* ══ TRENDS — Fig 2 + Fig 3 ════════════════════════════ */}
          {tab==='trends'&&(
            <>
              {/* FIX: domain [18,36] so 35°C line is visible; label moved to
                  insideBottomLeft so it sits BELOW the line and never clips at top */}
              <FigSection pm={pm} figNum="Fig. 2"
                title="Annual Mean Tw, WBGT (indoor) & Air Temperature"
                subtitle="Sen's slope trend lines overlaid (dashed) · Bernard & Pourmoghani (1999)"
                figRef={r2} filename="Fig2_Annual_Trend_Lines.png">
                <div ref={r2} style={chartBg}>
                  <ResponsiveContainer width="100%" height={340}>
                    <ComposedChart data={yearly}
                      margin={{top:8,right:24,left:8,bottom:36}}>
                      <CartesianGrid stroke={tc.grid} strokeDasharray="3 3"/>
                      <XAxis dataKey="year" tick={atick} interval={2}
                        label={xLab('Year')}/>
                      <YAxis tick={atick} unit="°C" domain={[18,36]}
                        label={yLab('Temperature (°C)')}/>
                      <Tooltip content={<Tip/>}/>
                      <Legend wrapperStyle={legStyle} iconSize={12}
                        verticalAlign="top" height={32}/>
                      {/* Label at insideBottomLeft = below the 35°C line, avoids clipping */}
                      <ReferenceLine y={35} stroke={tc.ref35} strokeWidth={1.5}
                        strokeDasharray="6 3"
                        label={{value:'35°C — survivability limit',fill:tc.ref35,
                          fontSize:FS.refLabel,position:'insideBottomLeft'}}/>
                      {mk&&(()=>{
                        const b0=yearly[0]?.yn||2000,w0=yearly[0]?.wbgt||28,tw0=yearly[0]?.tw||22;
                        return(<>
                          <Line type="linear" name="_wbgt_trend"
                            dataKey={y=>+(mk.wbgt.slope*(y.yn-b0)+w0).toFixed(3)}
                            stroke={C.wbgt} strokeWidth={1.5} strokeDasharray="8 4"
                            dot={false} legendType="none" unit="°C"/>
                          <Line type="linear" name="_tw_trend"
                            dataKey={y=>+(mk.tw.slope*(y.yn-b0)+tw0).toFixed(3)}
                            stroke={C.tw} strokeWidth={1.5} strokeDasharray="8 4"
                            dot={false} legendType="none" unit="°C"/>
                        </>);
                      })()}
                      <Line type="monotone" dataKey="wbgt" name="WBGT (indoor)"
                        stroke={C.wbgt} strokeWidth={2.5}
                        dot={{fill:C.wbgt,r:4,strokeWidth:0}} unit="°C"/>
                      <Line type="monotone" dataKey="tw" name="Avg Tw"
                        stroke={C.tw} strokeWidth={2.5}
                        dot={{fill:C.tw,r:4,strokeWidth:0}} unit="°C"/>
                      <Line type="monotone" dataKey="t" name="Air Temp"
                        stroke={C.temp} strokeWidth={1.8} strokeDasharray="5 3"
                        dot={{fill:C.temp,r:3,strokeWidth:0}} unit="°C"/>
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
                {mk&&(
                  <div style={{display:'flex',flexWrap:'wrap',gap:12,marginTop:10,
                    padding:'8px 12px',background:'rgba(0,0,0,0.15)',borderRadius:8,fontSize:11}}>
                    <span style={{color:C.wbgt}}>WBGT: +{mk.wbgt.slope}°C/yr · p(std)={mk.wbgt.std?.p} · p(mod)={mk.wbgt.mod?.p}</span>
                    <span style={{color:C.tw}}>Tw: +{mk.tw.slope}°C/yr · p(mod)={mk.tw.mod?.p}</span>
                    <span style={{color:C.muted}}>Air T: p(mod)={mk.t.mod?.p} (n.s.)</span>
                  </div>
                )}
              </FigSection>

              <FigSection pm={pm} figNum="Fig. 3"
                title="Annual WBGT Heat Stress Risk Days"
                subtitle="Stacked area · indoor WBGT thresholds · Bernard & Pourmoghani (1999)"
                figRef={r3} filename="Fig3_Annual_Danger_Days.png">
                <div ref={r3} style={chartBg}>
                  <ResponsiveContainer width="100%" height={320}>
                    <AreaChart data={yearly}
                      margin={{top:8,right:24,left:8,bottom:36}}>
                      <defs>
                        {[['aC',C.caution],['aD',C.danger],['aE',C.extreme]].map(([id,col])=>(
                          <linearGradient key={id} id={id} x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%"  stopColor={col} stopOpacity={0.9}/>
                            <stop offset="95%" stopColor={col} stopOpacity={0.65}/>
                          </linearGradient>
                        ))}
                      </defs>
                      <CartesianGrid stroke={tc.grid} strokeDasharray="3 3"/>
                      <XAxis dataKey="year" tick={atick} interval={2} label={xLab('Year')}/>
                      <YAxis tick={atick} label={yLab('Days per year')}/>
                      <Tooltip content={<Tip/>}/>
                      <Legend wrapperStyle={legStyle} iconSize={12}
                        verticalAlign="top" height={32}/>
                      <Area type="monotone" dataKey="c" name="Caution 28–32°C"
                        stackId="a" stroke={C.caution} fill="url(#aC)"
                        strokeWidth={1.5} unit=" days"/>
                      <Area type="monotone" dataKey="d" name="Danger 32–35°C"
                        stackId="a" stroke={C.danger} fill="url(#aD)"
                        strokeWidth={1.5} unit=" days"/>
                      <Area type="monotone" dataKey="e" name="Extreme ≥35°C"
                        stackId="a" stroke={C.extreme} fill="url(#aE)"
                        strokeWidth={2} unit=" days"/>
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              </FigSection>
            </>
          )}

          {/* ══ SEASONAL — Fig 4 ══════════════════════════════════ */}
          {/* FIX: reference line labels now use RefLineLabel with staggered
              yOff (14 / 30 / 46 px from top) — they no longer fight at the
              same position and overlap each other.                          */}
          {tab==='seasonal'&&(
            <FigSection pm={pm} figNum="Fig. 4"
              title="Seasonal Heat Stress — Dot Plot"
              subtitle="26-year seasonal means · indoor WBGT and Tw · Bernard & Pourmoghani (1999)"
              figRef={r4} filename="Fig4_Seasonal_WBGT.png">
              <div ref={r4} style={chartBg}>
                <ResponsiveContainer width="100%" height={300}>
                  <ComposedChart data={seasonal} layout="vertical"
                    margin={{top:16,right:30,left:130,bottom:36}}>
                    <CartesianGrid stroke={tc.grid} strokeDasharray="3 3" horizontal={false}/>
                    <XAxis type="number" tick={atick} unit="°C"
                      domain={[15,38]} ticks={[15,20,25,30,35]}
                      label={xLab('Temperature (°C)')}/>
                    <YAxis type="category" dataKey="name"
                      tick={{...atick,fontSize:FS.tick}} width={125}/>
                    <Tooltip content={<Tip/>}/>
                    <Legend wrapperStyle={legStyle} verticalAlign="top"
                      iconSize={12} height={30}/>
                    {/* Staggered labels: each at a different y offset so they
                        never overlap. textAnchor="end" + x = line_x - 5 so
                        text sits neatly to the left of each reference line.  */}
                    <ReferenceLine x={35} stroke={tc.ref35} strokeWidth={1.5}
                      strokeDasharray="5 3"
                      label={<RefLineLabel value="Extreme 35°C"
                        fill={tc.ref35} yOff={14}/>}/>
                    <ReferenceLine x={32} stroke={tc.ref32} strokeWidth={1.5}
                      strokeDasharray="5 3"
                      label={<RefLineLabel value="Danger 32°C"
                        fill={tc.ref32} yOff={30}/>}/>
                    <ReferenceLine x={28} stroke={tc.ref28} strokeWidth={1.5}
                      strokeDasharray="5 3"
                      label={<RefLineLabel value="Caution 28°C"
                        fill={tc.ref28} yOff={46}/>}/>
                    <Scatter dataKey="avgWbgt" name="WBGT (indoor)" fill={C.wbgt}
                      shape={p=><circle cx={p.x} cy={p.y} r={11}
                        fill={C.wbgt} stroke={pm?'#fff':'#080d18'} strokeWidth={2}/>}
                      unit="°C"/>
                    <Scatter dataKey="avgTw" name="Avg Tw" fill={C.tw}
                      shape={p=><circle cx={p.x} cy={p.y} r={11}
                        fill={C.tw} stroke={pm?'#fff':'#080d18'} strokeWidth={2}/>}
                      unit="°C"/>
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
              {/* Seasonal summary — for cross-checking with Table 6 */}
              <div style={{overflowX:'auto',marginTop:14}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                  <thead><tr style={{borderBottom:`1px solid ${C.border}`}}>
                    {['Season','WBGT (°C)','Tw (°C)','T (°C)','RH (%)','Caution','Danger','Extreme'].map((h,i)=>(
                      <th key={i} style={{padding:'6px 10px',textAlign:'left',
                        color:C.muted,fontWeight:600,fontSize:10}}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>{seasonal.map((s,i)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${C.grid}`}}>
                      <td style={{padding:'6px 10px',fontWeight:600,
                        color:[C.danger,C.tw,C.sub][i]}}>{s.name}</td>
                      <td style={{padding:'6px 10px'}}>{s.avgWbgt}</td>
                      <td style={{padding:'6px 10px'}}>{s.avgTw}</td>
                      <td style={{padding:'6px 10px'}}>{s.avgT}</td>
                      <td style={{padding:'6px 10px'}}>{s.avgRH}</td>
                      <td style={{padding:'6px 10px',color:C.caution}}>{s.c}</td>
                      <td style={{padding:'6px 10px',color:C.danger}}>{s.d}</td>
                      <td style={{padding:'6px 10px',color:C.extreme}}>{s.e}</td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
            </FigSection>
          )}

          {/* ══ EXCEEDANCE — Fig 5 ════════════════════════════════ */}
          {/* FIX: right margin 220px so "6,084 (64.1%)" labels never clip */}
          {tab==='exceedance'&&(
            <FigSection pm={pm} figNum="Fig. 5"
              title="Threshold Exceedance Frequency"
              subtitle="Indoor WBGT and Tw · matches Table 7 in paper"
              figRef={r5} filename="Fig5_Exceedance_Frequency.png">
              <div ref={r5} style={chartBg}>
                <ResponsiveContainer width="100%" height={300}>
                  <BarChart data={exceed} layout="vertical"
                    margin={{top:12,right:220,left:115,bottom:8}}>
                    <CartesianGrid stroke={tc.grid} strokeDasharray="3 3" horizontal={false}/>
                    <XAxis type="number" tick={atick}
                      label={{value:'Number of days',position:'insideBottom',
                        offset:-4,fill:tc.ax,fontSize:FS.axLabel}}/>
                    <YAxis type="category" dataKey="label"
                      tick={{...atick,fontSize:FS.tick}} width={110}/>
                    <Tooltip content={<Tip/>}/>
                    <Bar dataKey="n" name="Days" radius={[0,5,5,0]} unit=" days"
                      label={({x,y,width,height,value})=>
                        value?(
                          <text x={x+width+10} y={y+height/2}
                            fill={tc.ax} fontSize={FS.tick} fontWeight={600}
                            textAnchor="start" dominantBaseline="middle">
                            {value.toLocaleString()} ({data.length?+(value/data.length*100).toFixed(1):0}%)
                          </text>
                        ):null
                      }>
                      {exceed.map((e,i)=><Cell key={i} fill={e.color}/>)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </FigSection>
          )}

          {/* ══ OLS REGRESSION — Fig 6 ════════════════════════════ */}
          {/* FIX: domain zoomed to data range; CI = two thin dashed lines */}
          {tab==='regression'&&reg&&(
            <FigSection pm={pm} figNum="Fig. 6"
              title="OLS Regression of Annual Mean WBGT against Year"
              subtitle="Scatter = observed annual means · dashed red = OLS trend · thin dashed = 95% CI bounds"
              figRef={r6} filename="Fig6_OLS_Regression.png">
              <div ref={r6} style={chartBg}>
                <ResponsiveContainer width="100%" height={340}>
                  <ComposedChart data={reg.chart}
                    margin={{top:8,right:24,left:20,bottom:36}}>
                    <CartesianGrid stroke={tc.grid} strokeDasharray="3 3"/>
                    <XAxis dataKey="year" tick={atick} interval={2}
                      label={xLab('Year')}/>
                    <YAxis tick={atick} unit="°C" domain={reg.domain} width={80}
                      label={yLab('Annual mean WBGT (°C)')}/>
                    <Tooltip content={<Tip/>}/>
                    <Legend wrapperStyle={legStyle} iconSize={12}
                      verticalAlign="top" height={32}/>
                    {/* 95% CI upper and lower bounds — thin dashed, semi-transparent */}
                    <Line type="linear" dataKey="ciHi" name="_ci_hi"
                      stroke={C.wbgt} strokeWidth={1.2} dot={false}
                      strokeDasharray="4 3" opacity={0.55} legendType="none" unit="°C"/>
                    <Line type="linear" dataKey="ciLo" name="_ci_lo"
                      stroke={C.wbgt} strokeWidth={1.2} dot={false}
                      strokeDasharray="4 3" opacity={0.55} legendType="none" unit="°C"/>
                    {/* OLS trend line */}
                    <Line type="linear" dataKey="pred" name="OLS trend"
                      stroke={tc.ref35} strokeWidth={2.5} dot={false}
                      strokeDasharray="10 5" unit="°C"/>
                    {/* Observed data points */}
                    <Scatter dataKey="wbgt" name="Observed annual WBGT" fill={C.wbgt}
                      shape={p=><circle cx={p.x} cy={p.y} r={5.5}
                        fill={C.wbgt} stroke={pm?'#fff':'#080d18'}
                        strokeWidth={1.5} fillOpacity={0.9}/>}
                      unit="°C"/>
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
              <div style={{display:'flex',flexWrap:'wrap',gap:12,marginTop:12}}>
                {[{l:'WBGT indoor',r:reg.wbgt,c:C.wbgt},{l:'Tw',r:reg.tw,c:C.tw}]
                  .filter(x=>x.r).map((item,i)=>(
                  <div key={i} style={{background:'rgba(0,0,0,0.15)',borderRadius:8,
                    padding:'8px 14px',fontSize:11,borderLeft:`3px solid ${item.c}`}}>
                    <span style={{color:item.c,fontWeight:700}}>{item.l}: </span>
                    slope = <span style={{color:C.danger,fontWeight:700}}>+{item.r.slope}°C/yr</span>
                    &nbsp;· R² = <span style={{fontWeight:700}}>{item.r.r2}</span>
                  </div>
                ))}
              </div>
            </FigSection>
          )}

          {/* ══ HEATMAP — Fig 7 ═══════════════════════════════════ */}
          {tab==='heatmap'&&<HeatmapCanvas monthly={monthly}/>}

          {/* ══ WET BULB — Fig 8 ══════════════════════════════════ */}
          {/* FIX: 35°C label at insideBottomLeft so it sits below the line
              and is always inside the chart; interval=23 → one tick per 2 years */}
          {tab==='wetbulb'&&(
            <FigSection pm={pm} figNum="Fig. 8"
              title="Monthly Mean Wet Bulb Temperature (Tw)"
              subtitle="Stull (2011) formula · red dashed = 35°C physiological survivability limit"
              figRef={r8} filename="Fig8_Wet_Bulb_Temperature.png">
              <div ref={r8} style={chartBg}>
                <ResponsiveContainer width="100%" height={320}>
                  <AreaChart data={monthly}
                    margin={{top:16,right:24,left:8,bottom:36}}>
                    <defs>
                      <linearGradient id="gTw8" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%"  stopColor={C.tw} stopOpacity={0.45}/>
                        <stop offset="95%" stopColor={C.tw} stopOpacity={0.02}/>
                      </linearGradient>
                    </defs>
                    <CartesianGrid stroke={tc.grid} strokeDasharray="3 3"/>
                    {/* interval=23 → label every 24 months ≈ once per 2 years,
                        year-aligned and never crowded                          */}
                    <XAxis dataKey="label" tick={atick} interval={23}
                      label={xLab('Month / Year')}/>
                    <YAxis tick={atick} unit="°C" domain={[10,36]}
                      label={yLab('Wet Bulb Temperature (°C)')}/>
                    <Tooltip content={<Tip/>}/>
                    <ReferenceLine y={35} stroke={tc.ref35} strokeWidth={2}
                      strokeDasharray="7 4"
                      label={{value:'35°C — survivability limit',fill:tc.ref35,
                        fontSize:FS.refLabel,fontWeight:600,
                        position:'insideBottomLeft'}}/>
                    <Area type="monotone" dataKey="tw" name="Monthly mean Tw"
                      stroke={C.tw} fill="url(#gTw8)"
                      strokeWidth={2.5} dot={false} unit="°C"/>
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </FigSection>
          )}

          {/* ══ CORRELATIONS ══════════════════════════════════════ */}
          {tab==='correlation'&&(
            <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:16,padding:'20px'}}>
              <div style={{fontSize:9,color:C.accent,letterSpacing:3,textTransform:'uppercase',fontWeight:700,marginBottom:4}}>
                Pearson Correlation Matrix — Table 5 in paper
              </div>
              <div style={{fontSize:11,color:C.sub,marginBottom:14}}>
                Daily data · n={corr.n.toLocaleString()}
              </div>
              <div style={{overflowX:'auto'}}>
                <table style={{borderCollapse:'separate',borderSpacing:5,fontSize:12}}>
                  <thead><tr>
                    <th style={{padding:'8px 14px',color:C.muted,fontSize:11}}/>
                    {corr.vars.map((v,i)=>(
                      <th key={i} style={{padding:'8px 14px',textAlign:'center',
                        color:C.accent,fontWeight:700,fontSize:12}}>{v}</th>
                    ))}
                  </tr></thead>
                  <tbody>{corr.mat.map((row,i)=>(
                    <tr key={i}>
                      <td style={{padding:'8px 14px',color:C.accent,fontWeight:700,
                        whiteSpace:'nowrap'}}>{corr.vars[i]}</td>
                      {row.map((r,j)=>{
                        const bg=i===j?C.hover:r==null?C.hover:
                          r>0?`rgba(249,115,22,${(Math.abs(r)*0.75).toFixed(2)})`:
                          `rgba(56,189,248,${(Math.abs(r)*0.75).toFixed(2)})`;
                        return(
                          <td key={j} style={{padding:'10px 18px',textAlign:'center',
                            borderRadius:10,background:bg,
                            fontWeight:Math.abs(r||0)>=0.8||i===j?700:400,
                            color:C.text,minWidth:70}}>
                            {i===j?'—':r!=null?r.toFixed(3):'—'}
                          </td>
                        );
                      })}
                    </tr>
                  ))}</tbody>
                </table>
              </div>
              <div style={{marginTop:10,fontSize:11,color:C.muted}}>
                🟠 Positive · 🔵 Negative · |r| ≥ 0.80 very strong
              </div>
            </div>
          )}

        </div>
      )}
    </div>
  );
}