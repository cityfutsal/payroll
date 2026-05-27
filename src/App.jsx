import React, { useState, useEffect, useRef, useMemo, Component } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  CartesianGrid, Legend, AreaChart, Area, ComposedChart, Line,
  ReferenceLine, Cell
} from "recharts";

/* ─── DESIGN TOKENS ───────────────────────────────────────────────────────── */
const T = {
  navy:"#08152A", navy2:"#0F2240", navy3:"#1A3460",
  gold:"#C9A84C", goldf:"#F5DFA0",
  white:"#FFFFFF", off:"#F7F8FC", border:"#E2E8F2",
  muted:"#7B8DAB", text:"#1A2A42",
  green:"#0F7A52", greenl:"#E8F7F2",
  red:"#C0392B",   redl:"#FDECEA",
  amber:"#B86A00", amberl:"#FEF3E2",
  blue2:"#2563EB", blue2l:"#EFF6FF",
  dallas:"#1A3460", colony:"#0F7A52",
  shadow:"0 2px 16px rgba(8,21,42,0.08)",
  shadow2:"0 8px 40px rgba(8,21,42,0.16)",
};
const LOCS = ["Dallas","The Colony"];
const LOC_CLR = {"Dallas":T.dallas,"The Colony":T.colony};

// Per-location weekly Square sales goal (Venue Team target).
const GOAL_MIN = 3600;
const GOAL_MAX = 7000;

const FORM_TYPES = {
  preopen: {id:"preopen", label:"Pre-Open Checklist", icon:"☀️", color:"#B86A00"},
  closing: {id:"closing", label:"Closing Checklist",  icon:"🌙", color:"#1A3460"},
  hourly:  {id:"hourly",  label:"Hourly Reset",       icon:"🔄", color:"#2563EB"},
  other:   {id:"other",   label:"Form",               icon:"📋", color:"#7B8DAB"},
};
function normalizeVenue(s=""){
  const v=String(s).toLowerCase();
  if(v.includes("colony")) return "The Colony";
  if(v.includes("dallas")||v.includes("downtown")) return "Dallas";
  return s||"Unknown";
}
function detectFormType(headerRow1, headerRow2){
  const all=[...(headerRow1||[]),...(headerRow2||[])].map(s=>String(s||"").toLowerCase()).join(" | ");
  if(all.includes("pre-open")||all.includes("preopen")) return "preopen";
  if(all.includes("closing")) return "closing";
  if(all.includes("hourly")||all.includes("reset")) return "hourly";
  return "other";
}
const ROLE_CLR = (r="")=>{
  const s=r.toLowerCase();
  if(s.includes("lead")||s.includes("captain")) return T.gold;
  if(s.includes("senior")) return T.colony;
  if(s.includes("janitor")||s.includes("cleaning")) return T.amber;
  return T.navy3;
};
const DAYS = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

/* ─── UTILS ───────────────────────────────────────────────────────────────── */
const fmtUSD   = n=>"$"+(+n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtHHMM  = h=>{const m=Math.round((+h||0)*60);return `${Math.floor(m/60)}:${String(m%60).padStart(2,"0")}`;};
const fmtHHMMsigned = h=>{const neg=h<0;const m=Math.round(Math.abs(h)*60);return `${neg?"-":"+"} ${Math.floor(m/60)}:${String(m%60).padStart(2,"0")}`;};

function parseHHMM(v){
  if(!v||v==="NaN") return 0;
  const s=String(v).trim();
  const m=s.match(/^-?(\d+):(\d{2})$/);
  if(m) return (s.startsWith("-")?-1:1)*(+m[1]+ +m[2]/60);
  return parseFloat(s)||0;
}
function localDate(d){
  const s=String(d||"").slice(0,10);
  return s.match(/^\d{4}-\d{2}-\d{2}$/)?new Date(s+"T12:00:00"):new Date(d);
}
function toDateStr(v){
  if(!v) return "";
  const pad=n=>String(n).padStart(2,"0");
  if(typeof v==="number"){const d=new Date(Math.round((v-25569)*864e5));return d.toISOString().slice(0,10);}
  if(v instanceof Date){return `${v.getFullYear()}-${pad(v.getMonth()+1)}-${pad(v.getDate())}`;}
  const s=String(v).trim();
  if(/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  let m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);                       // M/D/YYYY
  if(m) return `${m[3]}-${pad(m[1])}-${pad(m[2])}`;
  m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\b/);                          // M/D/YY
  if(m) return `20${m[3]}-${pad(m[1])}-${pad(m[2])}`;
  const d=new Date(s);
  if(!isNaN(d)) return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  return s.slice(0,10);
}
function weekKey(d){const dt=localDate(d),j=new Date(dt.getFullYear(),0,1),w=Math.ceil(((dt-j)/864e5+j.getDay()+1)/7);return `${dt.getFullYear()}-W${String(w).padStart(2,"0")}`;}
function monthKey(d){const dt=localDate(d);return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,"0")}`;}
function yearKey(d){return String(localDate(d).getFullYear());}
function periodLabel(s,e){const f=d=>localDate(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});return `${f(s)} – ${f(e)}`;}
function dayOfWeek(d){return localDate(d).getDay();}
function shortDate(d){return localDate(d).toLocaleDateString("en-US",{month:"short",day:"numeric"});}

/* ─── EXCEL PARSER ────────────────────────────────────────────────────────── */
async function parseExcel(file){
  return new Promise((res,rej)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        const hi=rows.findIndex(row=>row.some(c=>String(c).toLowerCase().includes("first")));
        if(hi<0){rej("Header row not found");return;}
        const hdrs=rows[hi].map(h=>String(h).toLowerCase().trim());
        const data=rows.slice(hi+1).filter(row=>row.some(c=>c!==""));
        const get=(o,...keys)=>{for(const k of keys){const v=o[k];if(v!==undefined&&v!==""&&v!==null&&String(v)!=="NaN")return v;}return "";};
        const rowObjs=data.map(row=>{const o={};hdrs.forEach((h,i)=>{o[h]=row[i];});return o;});
        const empMap={};

        for(const o of rowObjs){
          const rawFirst=String(get(o,"first name","first_name","firstname")||"").trim();
          const rawLast =String(get(o,"last name","last_name","lastname")||"").trim();
          if(!rawFirst||!rawLast||rawFirst.toLowerCase()==="undefined") continue;
          const first=rawFirst.charAt(0).toUpperCase()+rawFirst.slice(1);
          const last=rawLast.split(/\s+/).map(w=>w.charAt(0).toUpperCase()+w.slice(1)).join(" ");
          const key=`${first}|${last}`;

          if(!empMap[key]){
            empMap[key]={
              location: String(get(o,"location")||"").trim(),
              firstName:first, lastName:last, role:"",
              rate:0, hours:0, scheduledHours:0, difference:0,
              overtime:0, pay:0, workedDays:0, notes:"",
              startDate:toDateStr(get(o,"start date","startdate","start")),
              endDate:toDateStr(get(o,"end date","enddate","end")),
              dailyBreakdown:[],
            };
          }
          const emp=empMap[key];

          // Role
          const role=String(get(o,"title","role","department","dept","position","job title")||"").trim();
          if(role&&role.length>1&&!emp.role) emp.role=role;

          // Rate
          const rate=parseFloat(get(o,"hourly rate (usd)","hourly rate","rate")||0);
          if(rate>0) emp.rate=rate;

          // Summary row fields (only populated on first/summary row per employee)
          const rawSched=get(o,"total scheduled hours","scheduled hours","scheduled");
          if(rawSched){const h=parseHHMM(rawSched);if(h>0) emp.scheduledHours=h;}

          const rawWorked=get(o,"total work hours","hours","total hours");
          if(rawWorked){const h=parseHHMM(rawWorked);if(h>0) emp.hours=h;}

          const rawDiff=get(o,"total difference (scheduled vs actual)","difference","diff");
          if(rawDiff){emp.difference=parseHHMM(rawDiff);}

          const rawOT=get(o,"overtime x1.5","overtime");
          if(rawOT){emp.overtime=parseHHMM(rawOT);}

          const wd=parseInt(get(o,"worked days","days")||0);
          if(wd>0) emp.workedDays=wd;

          // Daily pay — sum ALL rows
          const dp=parseFloat(get(o,"total pay","pay","total pay ($)")||0);
          if(dp>0) emp.pay+=dp;

          // Daily breakdown — one entry per row that has a date
          const dateVal=get(o,"date");
          if(dateVal){
            const dateStr=toDateStr(dateVal);
            const dailyPay=parseFloat(get(o,"total pay","pay")||0)||0;
            if(dateStr&&dailyPay>0){
              emp.dailyBreakdown.push({date:dateStr,pay:dailyPay,rate:rate||emp.rate});
            }
          }
        }

        // SECOND PASS — summary-style exports have a "Dates" / "Days Worked"
        // cell listing multiple dates as text (e.g. "5/18/2026, 5/19/2026,...").
        // Parse those into dailyBreakdown so sales attribution works for that
        // format too. Tagged fromSummary so downstream code knows pay is an
        // estimate (total pay / day count), not a per-row value.
        for(const o of rowObjs){
          const rawFirst=String(get(o,"first name","first_name","firstname")||"").trim();
          const rawLast =String(get(o,"last name","last_name","lastname")||"").trim();
          if(!rawFirst||!rawLast||rawFirst.toLowerCase()==="undefined") continue;
          const first=rawFirst.charAt(0).toUpperCase()+rawFirst.slice(1);
          const last=rawLast.split(/\s+/).map(w=>w.charAt(0).toUpperCase()+w.slice(1)).join(" ");
          const emp=empMap[`${first}|${last}`];
          if(!emp) continue;
          // Look at every cell except the single-date one we already processed.
          for(const [colHdr,v] of Object.entries(o)){
            if(!v||colHdr==="date") continue;
            const matches=String(v).match(/\b\d{1,2}\/\d{1,2}\/\d{2,4}\b|\b\d{4}-\d{2}-\d{2}\b/g);
            if(!matches||matches.length<2) continue;
            const parsed=matches.map(t=>toDateStr(t)).filter(d=>/^\d{4}-\d{2}-\d{2}$/.test(d));
            const unique=[...new Set(parsed)];
            const existing=new Set((emp.dailyBreakdown||[]).map(d=>d.date));
            const fresh=unique.filter(d=>!existing.has(d));
            if(!fresh.length) continue;
            const totalDays=(emp.dailyBreakdown.length+fresh.length)||emp.workedDays||fresh.length;
            const perDay=emp.pay>0&&totalDays>0?emp.pay/totalDays:0;
            fresh.forEach(d=>emp.dailyBreakdown.push({date:d,pay:perDay,rate:emp.rate,fromSummary:true}));
          }
        }

        const result=Object.values(empMap)
          .filter(e=>e.firstName&&e.lastName)
          .map(e=>{
            if(!e.role||e.role.length<2){
              e.role=e.rate>=22?"Venue Lead":e.rate>=18?"Senior Host":"Venue Host";
            }
            // Calculate hours from pay/rate if still 0
            if(e.hours===0&&e.pay>0&&e.rate>0) e.hours=e.pay/e.rate;
            // Calculate difference if missing
            if(e.difference===0&&e.scheduledHours>0&&e.hours>0) e.difference=e.hours-e.scheduledHours;
            return e;
          });
        res(result);
      }catch(err){rej(String(err));}
    };
    reader.readAsArrayBuffer(file);
  });
}

async function parseSquare(file){
  // Square exports vary in shape between report types:
  //   - "Sales summary - Daily"  → first line is the title, second line has date headers
  //   - "Sales summary - Weekly" → single total
  //   - "Item Sales" etc.        → may put title later, may have a blank/metadata line first
  // Strategy: search the WHOLE file for (a) a row whose first cell looks like dates and
  // (b) the first row whose first cell starts with "gross"; tolerate BOM, whitespace,
  // any case, and any line position.
  return new Promise((res,rej)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        let raw=String(e.target.result||"");
        // Strip UTF-8 BOM if present
        if(raw.charCodeAt(0)===0xFEFF) raw=raw.slice(1);
        const lines=raw.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
        if(!lines.length){rej("Empty file");return;}

        function parseCsvLine(line){
          const out=[];let cur='',inQ=false;
          for(let i=0;i<line.length;i++){
            const ch=line[i];
            if(ch==='"'){inQ=!inQ;}
            else if(ch===','&&!inQ){out.push(cur.trim());cur='';}
            else{cur+=ch;}
          }
          out.push(cur.trim());
          return out;
        }
        function parseMoney(s){
          s=String(s||"").replace(/"/g,'').replace(/\$/g,'').replace(/,/g,'').trim();
          if(!s) return 0;
          if(s.startsWith('(')&&s.endsWith(')')) return -parseFloat(s.slice(1,-1))||0;
          return parseFloat(s)||0;
        }
        function isDateLike(s){return /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(String(s||"").replace(/"/g,'').trim());}
        function toIso(s){
          const p=String(s||"").replace(/"/g,'').trim().split('/');
          if(p.length!==3) return null;
          const yyyy=p[2].length===2?`20${p[2]}`:p[2];
          return `${yyyy}-${p[0].padStart(2,'0')}-${p[1].padStart(2,'0')}`;
        }

        // Find the date-header row anywhere in the file (any row where 2+ cells are M/D/Y).
        let dates=null;
        for(let i=0;i<Math.min(lines.length,15);i++){
          const cols=parseCsvLine(lines[i]);
          const dateCells=cols.filter(isDateLike);
          if(dateCells.length>=2){ dates=cols.map(c=>isDateLike(c)?c.replace(/"/g,'').trim():""); break; }
        }

        // Find the first row whose first cell starts with "gross" (case-insensitive).
        let grossCols=null;
        for(let i=0;i<lines.length;i++){
          const cols=parseCsvLine(lines[i]);
          const m=(cols[0]||"").toLowerCase().replace(/"/g,'').trim();
          if(m.startsWith('gross')){ grossCols=cols; break; }
        }

        const results=[];
        if(grossCols){
          if(dates){
            dates.forEach((dateStr,di)=>{
              if(!dateStr) return;
              const val=parseMoney(grossCols[di+1]||"0");
              if(val<=0) return;
              const iso=toIso(dateStr);
              if(iso) results.push({date:iso,gross:Math.round(val*100)/100});
            });
          } else {
            const val=parseMoney(grossCols[1]||"0");
            if(val>0) results.push({date:'summary',gross:Math.round(val*100)/100,isWeeklyTotal:true});
          }
        }

        if(!results.length){
          const preview=lines.slice(0,3).map(l=>l.slice(0,80)).join(" | ");
          rej(`No Gross sales row found. First lines: ${preview}`);
          return;
        }
        res(results);
      }catch(err){rej("Parse error: "+String(err));}
    };
    reader.onerror=()=>rej("File read error");
    reader.readAsText(file,'utf-8');
  });
}

/* ─── CONNECTEAM FORM PARSER ──────────────────────────────────────────────── */
async function parseConnecteamForm(file){
  return new Promise((res,rej)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:false});
        if(rows.length<3){rej("Form has no submissions");return;}
        const sectionRow=rows[0]||[];
        const headerRow=rows[1]||[];
        const formType=detectFormType(sectionRow,headerRow);
        const findIdx=(matcher)=>headerRow.findIndex(h=>matcher(String(h||"").toLowerCase().trim()));
        const fullNameIdx=findIdx(h=>h==="full name");
        const dateIdx=findIdx(h=>h==="submission date");
        const timeIdx=findIdx(h=>h==="submission time");
        const venueIdx=findIdx(h=>h.includes("venue")&&!h.endsWith("- location stamp")&&!h.endsWith("- timestamp"));
        if(fullNameIdx<0||dateIdx<0){rej("Required columns (Full name, Submission Date) not found — is this a Connecteam Entries export?");return;}
        const isMetaIdx=i=>{
          const h=String(headerRow[i]||"").toLowerCase().trim();
          return h.endsWith("- location stamp")||h.endsWith("- timestamp")||
                 h==="#"||h==="status"||h==="audit"||
                 h==="status - last status change"||h==="audit - last status change";
        };
        const questionIdxs=headerRow
          .map((_,i)=>i)
          .filter(i=>i!==fullNameIdx&&i!==dateIdx&&i!==timeIdx&&i!==venueIdx&&!isMetaIdx(i));
        const submissions=[];
        for(let r=2;r<rows.length;r++){
          const row=rows[r];
          if(!row||!row.some(c=>c!==""&&c!=null)) continue;
          const fullName=String(row[fullNameIdx]||"").trim();
          if(!fullName) continue;
          const dateStr=toDateStr(row[dateIdx]);
          const submissionTime=String(row[timeIdx]||"").trim();
          const rawVenue=String(row[venueIdx]>=0?row[venueIdx]:"").trim();
          const venue=normalizeVenue(rawVenue);
          const answers={};
          let complete=0,positive=0,negative=0,naCount=0,images=0;
          const flags=[];
          questionIdxs.forEach(i=>{
            const q=String(headerRow[i]||"").trim();
            const v=row[i];
            if(v==null||v==="") return;
            const sv=String(v).trim();
            answers[q]=sv;
            const lc=sv.toLowerCase();
            if(lc==="complete"){complete++;positive++;}
            else if(lc==="yes") positive++;
            else if(lc==="no"){negative++;flags.push(q);}
            else if(lc==="n/a"||lc==="na") naCount++;
            else if(lc==="image") images++;
          });
          submissions.push({
            id:`${dateStr}|${fullName}|${submissionTime}`.replace(/\s+/g,"_"),
            formType, submittedBy:fullName, venue, rawVenue,
            submissionDate:dateStr, submissionTime,
            answers,
            summary:{questions:questionIdxs.length,complete,positive,negative,naCount,images,flags},
          });
        }
        if(!submissions.length){rej("No valid submission rows found");return;}
        res({formType,submissions});
      }catch(err){rej("Parse error: "+String(err));}
    };
    reader.onerror=()=>rej("File read error");
    reader.readAsArrayBuffer(file);
  });
}

// Aggregate an array of square rows (possibly from multiple files/locations)
// into per-date objects with a byLocation breakdown and combined gross.
function aggregateSquareRows(rows){
  const m={};
  (rows||[]).forEach(r=>{
    const d=r.date||"";
    if(!d) return;
    const loc=r.location||"Unknown";
    if(d==="summary") return; // skip weekly summary rows here
    if(!m[d]) m[d]={date:d,gross:0,byLocation:{}};
    m[d].byLocation[loc]=(m[d].byLocation[loc]||0)+(+r.gross||0);
    m[d].gross=Object.values(m[d].byLocation).reduce((s,v)=>s+v,0);
    m[d].gross=Math.round(m[d].gross*100)/100;
  });
  return Object.values(m).sort((a,b)=>a.date.localeCompare(b.date));
}

// Migration: pre-v3.3.1 uploads stored dates as broken strings ("Mon May 18")
// that re-parsed to year 2001. Shift the year forward to match uploadedAt so
// day-of-week labels and Square sales attribution work without re-uploading.
function fixWeekYear(week){
  if(!week) return null;
  const uploadYear=week.uploadedAt?new Date(week.uploadedAt).getFullYear():new Date().getFullYear();
  const sample=week.startDate||week.employees?.[0]?.dailyBreakdown?.[0]?.date||"";
  const m=String(sample).match(/^(\d{4})-/);
  if(!m) return null;
  const dataYear=+m[1];
  if(Math.abs(uploadYear-dataYear)<2) return null;
  const shift=s=>{
    if(typeof s!=="string") return s;
    const m2=s.match(/^\d{4}(-\d{2}-\d{2}.*)$/);
    return m2?`${uploadYear}${m2[1]}`:s;
  };
  return{
    ...week,
    startDate:shift(week.startDate),
    endDate:shift(week.endDate),
    period:(week.startDate&&week.endDate)?periodLabel(shift(week.startDate),shift(week.endDate)):week.period,
    employees:(week.employees||[]).map(e=>({
      ...e,
      startDate:shift(e.startDate),
      endDate:shift(e.endDate),
      dailyBreakdown:(e.dailyBreakdown||[]).map(d=>({...d,date:shift(d.date)})),
    })),
  };
}

// Storage backend is window.storage from main.jsx — Supabase when env vars
// are set, localStorage fallback otherwise. Keys are JSON strings under
// "cf:week:..." and "cf:att:..." prefixes.
const S=()=>window.storage;
async function storeLoad(){
  try{
    const weeks={},attachments={};
    const {keys=[]} = await S().list("cf:");
    const pairs = await Promise.all(keys.map(async k=>{
      const r=await S().get(k);
      try{ return [k, r?JSON.parse(r.value):null]; }catch{ return [k,null]; }
    }));
    pairs.forEach(([k,val])=>{
      if(val==null) return;
      if(k.startsWith("cf:week:")) weeks[k.replace("cf:week:","")] = val;
      else if(k.startsWith("cf:att:"))  attachments[k.replace("cf:att:","")] = val;
    });
    // Migrate broken-year stored weeks (pre-v3.3.1 toDateStr bug).
    await Promise.all(Object.keys(weeks).map(async k=>{
      const fixed=fixWeekYear(weeks[k]);
      if(fixed){
        weeks[k]=fixed;
        try{ await S().set(`cf:week:${k}`,JSON.stringify(fixed)); }catch{}
      }
    }));
    return {weeks,attachments};
  }catch{return {weeks:{},attachments:{}};}
}
const storeWeek=(k,v)=>S().set(`cf:week:${k}`,JSON.stringify(v));
const storeAtt =(k,v)=>S().set(`cf:att:${k}`,JSON.stringify(v));
const delWeek  =k=>S().delete(`cf:week:${k}`);
const delAtt   =k=>S().delete(`cf:att:${k}`);

/* ─── MICRO COMPONENTS ────────────────────────────────────────────────────── */
function Pill({children,color=T.navy3}){
  return <span style={{background:color+"18",color,border:`1px solid ${color}30`,borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",whiteSpace:"nowrap"}}>{children}</span>;
}
function Delta({val,hours=false}){
  if(val===undefined||val===null) return null;
  const up=val>0,dn=val<0;
  const disp=hours?fmtHHMMsigned(val):(Math.abs(val).toFixed(1)+"%");
  return <span style={{color:up?T.green:dn?T.red:T.muted,fontWeight:700,fontSize:12,fontFamily:"'DM Mono',monospace"}}>{up?"↑":dn?"↓":"—"} {disp}</span>;
}
function KPI({label,value,sub,delta,accent=T.gold,mono=false}){
  return(
    <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:"18px 22px",flex:1,minWidth:140,borderLeft:`4px solid ${accent}`,boxShadow:T.shadow}}>
      <div style={{fontSize:10,fontWeight:800,color:T.muted,letterSpacing:"0.12em",textTransform:"uppercase",marginBottom:8}}>{label}</div>
      <div style={{fontSize:22,fontWeight:900,color:T.navy,fontFamily:mono?"'DM Mono',monospace":"inherit"}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:T.muted,marginTop:3}}>{sub}</div>}
      {delta!==undefined&&<div style={{marginTop:5}}><Delta val={delta}/><span style={{color:T.muted,fontSize:10,marginLeft:4}}>vs prior</span></div>}
    </div>
  );
}
function Toast({msg,type="ok",onClose}){
  useEffect(()=>{const t=setTimeout(onClose,3400);return()=>clearTimeout(t);},[]);
  return <div style={{position:"fixed",top:20,right:20,zIndex:9999,background:type==="err"?T.red:type==="warn"?T.amber:T.green,color:T.white,padding:"12px 20px",borderRadius:10,fontWeight:700,fontSize:13,boxShadow:T.shadow2,display:"flex",alignItems:"center",gap:12,minWidth:260,animation:"toastIn .25s ease"}}><span>{type==="err"?"✖":type==="warn"?"⚠":"✓"}</span><span style={{flex:1}}>{msg}</span><span onClick={onClose} style={{cursor:"pointer",opacity:.7}}>×</span></div>;
}
function Modal({title,children,onClose,wide}){
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(8,21,42,0.65)",zIndex:9998,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{background:T.white,borderRadius:18,padding:32,width:"100%",maxWidth:wide?900:520,maxHeight:"92vh",overflow:"auto",boxShadow:T.shadow2}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
          <div style={{fontWeight:900,fontSize:18,color:T.navy}}>{title}</div>
          <button onClick={onClose} style={{background:"none",border:`1px solid ${T.border}`,width:32,height:32,borderRadius:8,cursor:"pointer",fontSize:18,color:T.muted,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}
function UploadZone({onFiles,label,sub,accept=".xlsx,.xls,.pdf,.csv",icon="📂",compact}){
  const [drag,setDrag]=useState(false);const ref=useRef();
  const go=files=>{if(files.length)onFiles(Array.from(files));};
  return(
    <div onClick={()=>ref.current.click()} onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);go(e.dataTransfer.files);}}
      style={{border:`2px dashed ${drag?T.gold:T.border}`,borderRadius:12,padding:compact?"16px 20px":"28px 24px",textAlign:"center",cursor:"pointer",background:drag?T.gold+"08":T.off,transition:"all .2s",userSelect:"none"}}>
      <input ref={ref} type="file" accept={accept} multiple style={{display:"none"}} onChange={e=>go(e.target.files)}/>
      <div style={{fontSize:compact?20:26,marginBottom:6}}>{icon}</div>
      <div style={{fontWeight:800,color:T.navy,fontSize:compact?12:14}}>{label}</div>
      {sub&&<div style={{fontSize:11,color:T.muted,marginTop:4}}>{sub}</div>}
    </div>
  );
}

/* ─── COLLAPSIBLE SECTION ─────────────────────────────────────────────────── */
function Section({title,icon,subtitle,right,defaultOpen=false,children}){
  const [open,setOpen]=useState(defaultOpen);
  return(
    <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,boxShadow:T.shadow,overflow:"hidden"}}>
      <button onClick={()=>setOpen(o=>!o)} style={{width:"100%",display:"flex",alignItems:"center",gap:12,padding:"14px 20px",background:"transparent",border:"none",cursor:"pointer",fontFamily:"inherit",textAlign:"left"}}>
        {icon&&<span style={{fontSize:18,lineHeight:1}}>{icon}</span>}
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontWeight:800,color:T.navy,fontSize:14}}>{title}</div>
          {subtitle&&<div style={{color:T.muted,fontSize:11,marginTop:2}}>{subtitle}</div>}
        </div>
        {right}
        <span style={{color:T.muted,fontSize:18,display:"inline-block",transform:`rotate(${open?90:0}deg)`,transition:"transform .18s",lineHeight:1}}>›</span>
      </button>
      {open&&<div style={{padding:"4px 20px 22px",borderTop:`1px solid ${T.border}`}}>{children}</div>}
    </div>
  );
}

/* ─── SCHEDULED VS ACTUAL CHART ───────────────────────────────────────────── */
function ScheduleVsActual({employees}){
  const data=useMemo(()=>
    (employees||[])
      .map(e=>({e,sched:Number(e.scheduledHours)||0,worked:Number(e.hours)||0}))
      .filter(({sched,worked})=>sched>0||worked>0)
      .sort((a,b)=>b.worked-a.worked)
      .map(({e,sched,worked})=>({
        name:`${e.firstName} ${e.lastName?.charAt(0)||""}.`,
        fullName:`${e.firstName} ${e.lastName||""}`,
        role:e.role,
        location:e.location,
        scheduled:+sched.toFixed(2),
        worked:+worked.toFixed(2),
        diff:+(worked-sched).toFixed(2),
        over:worked>sched?+(worked-sched).toFixed(2):0,
        under:worked<sched?+(sched-worked).toFixed(2):0,
      }))
  ,[employees]);

  if(!data.length) return <div style={{textAlign:"center",color:T.muted,padding:40}}>Upload Connecteam exports to see schedule vs worked hours.</div>;

  const CustomTooltip=({active,payload,label})=>{
    if(!active||!payload?.length) return null;
    const d=data.find(r=>r.name===label)||{};
    return(
      <div style={{background:T.navy,border:`1px solid ${T.navy3}`,borderRadius:10,padding:"12px 16px",boxShadow:T.shadow2,minWidth:200}}>
        <div style={{color:T.gold,fontWeight:800,fontSize:13,marginBottom:8}}>{d.fullName}</div>
        <div style={{color:T.muted,fontSize:11,marginBottom:8}}>{d.role} · {d.location}</div>
        {payload.map(p=>(
          <div key={p.name} style={{display:"flex",justifyContent:"space-between",gap:16,marginBottom:3,fontSize:12,color:T.white}}>
            <span style={{color:p.color+"CC"}}>{p.name}</span>
            <span style={{fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmtHHMM(p.value)}</span>
          </div>
        ))}
        <div style={{borderTop:`1px solid ${T.navy3}`,marginTop:8,paddingTop:8,fontSize:12,color:d.diff>0?T.green:d.diff<0?T.red:T.muted,fontWeight:700}}>
          {d.diff>0?"↑ Over":"d.diff<0"?"↓ Under":"On schedule"}: {fmtHHMM(Math.abs(d.diff))}
        </div>
      </div>
    );
  };

  const totalOver=data.reduce((s,r)=>s+r.over,0);
  const totalUnder=data.reduce((s,r)=>s+r.under,0);
  const overCount=data.filter(r=>r.over>0).length;
  const underCount=data.filter(r=>r.under>0).length;

  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      {/* Summary pills */}
      <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
        <div style={{background:T.greenl,border:`1px solid ${T.green}30`,borderRadius:10,padding:"10px 16px",fontSize:13}}>
          <span style={{fontWeight:800,color:T.green}}>{overCount} staff over schedule</span>
          <span style={{color:T.muted,marginLeft:8}}>+{fmtHHMM(totalOver)} total</span>
        </div>
        <div style={{background:T.redl,border:`1px solid ${T.red}30`,borderRadius:10,padding:"10px 16px",fontSize:13}}>
          <span style={{fontWeight:800,color:T.red}}>{underCount} staff under schedule</span>
          <span style={{color:T.muted,marginLeft:8}}>-{fmtHHMM(totalUnder)} total</span>
        </div>
      </div>

      {/* Bar chart */}
      <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:"20px 20px 8px",boxShadow:T.shadow}}>
        <div style={{fontWeight:800,color:T.navy,fontSize:14,marginBottom:16}}>Scheduled vs Worked Hours</div>
        <ResponsiveContainer width="100%" height={Math.max(200,data.length*38)}>
          <BarChart data={data} layout="vertical" barCategoryGap="25%" barGap={2}>
            <CartesianGrid strokeDasharray="3 3" stroke={T.border} horizontal={false}/>
            <XAxis type="number" tickFormatter={v=>`${v}h`} tick={{fontSize:11,fill:T.muted}}/>
            <YAxis type="category" dataKey="name" width={90} tick={{fontSize:12,fill:T.text,fontWeight:600}}/>
            <Tooltip content={<CustomTooltip/>}/>
            <Legend wrapperStyle={{fontSize:12}}/>
            <Bar dataKey="scheduled" name="Scheduled" fill={T.border} radius={[0,4,4,0]}>
              {data.map((e,i)=><Cell key={i} fill={T.navy3+"60"}/>)}
            </Bar>
            <Bar dataKey="worked" name="Worked" radius={[0,4,4,0]}>
              {data.map((e,i)=><Cell key={i} fill={e.diff>0?T.green:e.diff<0?T.red:T.colony}/>)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      {/* Detail table */}
      <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow}}>
        <div style={{background:T.navy,padding:"12px 18px"}}>
          <span style={{color:T.gold,fontWeight:800,fontSize:13}}>Schedule Variance by Employee</span>
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead>
            <tr style={{background:T.off}}>
              {["Employee","Role","Location","Scheduled","Worked","Variance","Overtime"].map(h=>(
                <th key={h} style={{padding:"10px 14px",textAlign:["Scheduled","Worked","Variance","Overtime"].includes(h)?"right":"left",
                  fontSize:10,fontWeight:800,color:T.muted,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${T.border}`}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {[...data].sort((a,b)=>Math.abs(b.diff)-Math.abs(a.diff)).map((e,i)=>{
              const emp=employees.find(x=>x.firstName+" "+x.lastName===e.fullName)||{};
              return(
                <tr key={i} style={{background:i%2===0?T.white:"#F9FAFD",borderBottom:`1px solid ${T.border}`}}>
                  <td style={{padding:"10px 14px",fontWeight:700,color:T.navy}}>{e.fullName}</td>
                  <td style={{padding:"10px 14px"}}><Pill color={ROLE_CLR(e.role)}>{e.role}</Pill></td>
                  <td style={{padding:"10px 14px",color:T.muted,fontSize:12}}>{e.location}</td>
                  <td style={{padding:"10px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{fmtHHMM(e.scheduled)}</td>
                  <td style={{padding:"10px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmtHHMM(e.worked)}</td>
                  <td style={{padding:"10px 14px",textAlign:"right"}}>
                    <span style={{fontFamily:"'DM Mono',monospace",fontWeight:700,color:e.diff>0?T.green:e.diff<0?T.red:T.muted}}>
                      {e.diff>0?"+":""}{fmtHHMM(Math.abs(e.diff))} {e.diff>0?"↑":e.diff<0?"↓":""}
                    </span>
                  </td>
                  <td style={{padding:"10px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:emp.overtime>0?T.amber:T.muted}}>
                    {emp.overtime>0?fmtHHMM(emp.overtime):"—"}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ─── STAFFING ACTIVITY CHART (daily coverage + Square overlay) ───────────── */
function StaffingActivityChart({week,squareData}){
  // Build daily staffing coverage from daily breakdown
  const dailyStaffing=useMemo(()=>{
    if(!week?.employees) return [];
    const byDate={};
    week.employees.forEach(emp=>{
      (emp.dailyBreakdown||[]).forEach(d=>{
        const iso=toDateStr(d.date);
        if(!iso) return;
        if(!byDate[iso]) byDate[iso]={date:iso,headcount:0,totalPay:0,byRole:{}};
        byDate[iso].headcount+=1;
        byDate[iso].totalPay+=Number(d.pay)||0;
        const role=emp.role||"Venue Host";
        byDate[iso].byRole[role]=(byDate[iso].byRole[role]||0)+1;
      });
    });
    return Object.values(byDate).sort((a,b)=>String(a.date).localeCompare(String(b.date))).map(d=>({
      ...d,
      day:DAYS[dayOfWeek(d.date)]+" "+shortDate(d.date),
      dateStr:d.date,
      "Venue Lead":d.byRole["Venue Lead"]||0,
      "Senior Host":d.byRole["Senior Host"]||0,
      "Venue Host":d.byRole["Venue Host"]||0,
    }));
  },[week]);

  // Merge Square data
  const squareByDate=useMemo(()=>{
    const m={};
    (squareData||[]).forEach(s=>{
      // s = {date, gross, byLocation:{loc:amount}}
      m[s.date]={gross:s.gross,byLocation:s.byLocation||{}};
    });
    return m;
  },[squareData]);

  const chartData=useMemo(()=>
    dailyStaffing.map(d=>{
      const s=squareByDate[d.dateStr]||{gross:0,byLocation:{}};
      const out={...d,sales_total:s.gross};
      LOCS.forEach(l=>{out[`sales_${l.replace(/\s/g,'_')}`]=s.byLocation[l]||0;});
      return out;
    })
  ,[dailyStaffing,squareByDate]);

  if(!chartData.length) return <div style={{textAlign:"center",color:T.muted,padding:40,fontSize:13}}>Upload Connecteam exports to see daily staffing activity.</div>;

  const hasSales=chartData.some(d=>d.sales>0);

  const CustomTooltip=({active,payload,label})=>{
    if(!active||!payload?.length) return null;
    return(
      <div style={{background:T.navy,border:`1px solid ${T.navy3}`,borderRadius:10,padding:"12px 16px",boxShadow:T.shadow2}}>
        <div style={{color:T.gold,fontWeight:800,fontSize:12,marginBottom:8}}>{label}</div>
        {payload.map(p=>{
          const isSales=p.name?.startsWith("sales_")||p.name==="Square Sales"||p.dataKey==="sales_total";
          return (
            <div key={p.name} style={{display:"flex",justifyContent:"space-between",gap:16,marginBottom:3,fontSize:12,color:T.white}}>
              <span style={{color:p.color+"CC"}}>{p.name.replace(/^sales_/,'').replace(/_/g,' ')}</span>
              <span style={{fontFamily:"'DM Mono',monospace",fontWeight:700}}>{isSales?fmtUSD(p.value):`${p.value} staff`}</span>
            </div>
          );
        })}
      </div>
    );
  };

  return(
    <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:"20px 20px 8px",boxShadow:T.shadow}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div style={{fontWeight:800,color:T.navy,fontSize:14}}>
          Daily Staffing Coverage{hasSales?" + Square Sales":""}
        </div>
        {!hasSales&&<span style={{fontSize:12,color:T.muted}}>Upload a Square export to overlay sales data</span>}
      </div>
      <ResponsiveContainer width="100%" height={280}>
        <ComposedChart data={chartData} barCategoryGap="25%" barGap={1}>
          <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
          <XAxis dataKey="day" tick={{fontSize:11,fill:T.muted}}/>
          <YAxis yAxisId="staff" orientation="left" tick={{fontSize:11,fill:T.muted}} label={{value:"Staff",angle:-90,position:"insideLeft",style:{fontSize:10,fill:T.muted}}}/>
          {hasSales&&<YAxis yAxisId="sales" orientation="right" tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} tick={{fontSize:11,fill:T.muted}}/>}
          <Tooltip content={<CustomTooltip/>}/>
          <Legend wrapperStyle={{fontSize:12}}/>
          <Bar yAxisId="staff" dataKey="Venue Lead"   name="Venue Lead"   stackId="staff" fill={T.gold}    radius={[0,0,0,0]}/>
          <Bar yAxisId="staff" dataKey="Senior Host"  name="Senior Host"  stackId="staff" fill={T.colony}  radius={[0,0,0,0]}/>
          <Bar yAxisId="staff" dataKey="Venue Host"   name="Venue Host"   stackId="staff" fill={T.navy3}   radius={[4,4,0,0]}/>
          {hasSales&&LOCS.map(l=>{
            const key=`sales_${l.replace(/\s/g,'_')}`;
            return <Line key={key} yAxisId="sales" type="monotone" dataKey={key} name={l} stroke={LOC_CLR[l]||T.gold} strokeWidth={2} dot={{r:3}}/>;
          })}
          {hasSales&&<Line yAxisId="sales" type="monotone" dataKey="sales_total" name="Total Sales" stroke={T.gold} strokeWidth={2.5} dot={{fill:T.gold,r:5,stroke:T.white,strokeWidth:2}}/>}
        </ComposedChart>
      </ResponsiveContainer>
    </div>
  );
}

/* ─── WEEKLY GANTT (employee × day grid) ─────────────────────────────────── */
function WeeklyGantt({week}){
  const [locFilter,setLocFilter]=useState("All");
  if(!week?.employees) return <div style={{textAlign:"center",color:T.muted,padding:40}}>No data.</div>;

  // Get all dates in the week, normalized to ISO, ordered Mon → Sun.
  // Normalization covers legacy stored data that wasn't ISO (e.g. "Fri May 18").
  const allDates=useMemo(()=>{
    const d=new Set();
    week.employees.forEach(e=>(e.dailyBreakdown||[]).forEach(b=>{
      const iso=toDateStr(b.date);
      if(iso) d.add(iso);
    }));
    const dowOrder=s=>{const x=dayOfWeek(s);return x===0?6:x-1;}; // Mon=0..Sun=6
    return [...d].sort((a,b)=>dowOrder(a)-dowOrder(b));
  },[week]);

  const filtered=week.employees.filter(e=>locFilter==="All"||e.location===locFilter||normalizeVenue(e.location)===locFilter);

  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      {/* Filter */}
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        {["All",...LOCS].map(l=>(
          <button key={l} onClick={()=>setLocFilter(l)} style={{padding:"7px 16px",borderRadius:8,border:`1px solid ${locFilter===l?(LOC_CLR[l]||T.gold):T.border}`,background:locFilter===l?(LOC_CLR[l]||T.gold)+"12":T.white,color:locFilter===l?(LOC_CLR[l]||T.gold):T.navy,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
            {l}
          </button>
        ))}
      </div>

      <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow}}>
        <div style={{background:T.navy,padding:"12px 18px"}}>
          <span style={{color:T.gold,fontWeight:800,fontSize:13}}>Staff Coverage Grid — {week.period}</span>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead>
              <tr style={{background:T.off}}>
                <th style={{padding:"10px 14px",textAlign:"left",fontWeight:800,color:T.muted,fontSize:10,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${T.border}`,minWidth:160,position:"sticky",left:0,background:T.off}}>Employee</th>
                <th style={{padding:"10px 14px",textAlign:"left",fontWeight:800,color:T.muted,fontSize:10,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${T.border}`,minWidth:100}}>Role</th>
                {allDates.map(d=>(
                  <th key={d} style={{padding:"10px 12px",textAlign:"center",fontWeight:700,color:T.muted,fontSize:10,borderBottom:`1px solid ${T.border}`,minWidth:64,whiteSpace:"nowrap"}}>
                    <div>{DAYS[dayOfWeek(d)]}</div>
                    <div style={{fontWeight:900,color:T.navy}}>{shortDate(d).split(" ")[1]}</div>
                  </th>
                ))}
                <th style={{padding:"10px 14px",textAlign:"right",fontWeight:800,color:T.muted,fontSize:10,borderBottom:`1px solid ${T.border}`}}>Total</th>
              </tr>
            </thead>
            <tbody>
              {filtered.sort((a,b)=>a.lastName.localeCompare(b.lastName)).map((emp,i)=>{
                const payByDate={};
                (emp.dailyBreakdown||[]).forEach(d=>{const iso=toDateStr(d.date);if(iso)payByDate[iso]=d.pay;});
                const clr=LOC_CLR[normalizeVenue(emp.location)]||LOC_CLR[emp.location]||T.navy;
                return(
                  <tr key={i} style={{background:i%2===0?T.white:"#F9FAFD",borderBottom:`1px solid ${T.border}`}}>
                    <td style={{padding:"10px 14px",position:"sticky",left:0,background:i%2===0?T.white:"#F9FAFD",zIndex:1}}>
                      <div style={{fontWeight:700,color:T.navy,fontSize:12}}>{emp.firstName} {emp.lastName}</div>
                      <div style={{fontSize:10,color:clr,fontWeight:600}}>{emp.location}</div>
                    </td>
                    <td style={{padding:"10px 14px"}}><Pill color={ROLE_CLR(emp.role)}>{(emp.role||"").replace("Facility Coordinator","Coord").replace("Senior Host","Sr. Host")}</Pill></td>
                    {allDates.map(d=>{
                      const pay=payByDate[d];
                      const hrs=pay&&emp.rate?pay/emp.rate:0;
                      return(
                        <td key={d} style={{padding:"8px 6px",textAlign:"center"}}>
                          {pay?(
                            <div style={{background:clr+"18",border:`1px solid ${clr}30`,borderRadius:8,padding:"4px 6px",minWidth:52}}>
                              <div style={{fontWeight:800,color:clr,fontSize:11,fontFamily:"'DM Mono',monospace"}}>{fmtHHMM(hrs)}</div>
                              <div style={{fontSize:9,color:T.muted}}>{fmtUSD(pay).replace("$","$")}</div>
                            </div>
                          ):(
                            <div style={{width:52,height:36,display:"flex",alignItems:"center",justifyContent:"center"}}>
                              <div style={{width:8,height:2,background:T.border,borderRadius:2}}/>
                            </div>
                          )}
                        </td>
                      );
                    })}
                    <td style={{padding:"10px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:800,color:T.navy}}>{fmtUSD(emp.pay)}</td>
                  </tr>
                );
              })}
              {/* Headcount row */}
              <tr style={{background:T.navy,borderTop:`2px solid ${T.gold}`}}>
                <td style={{padding:"10px 14px",color:T.gold,fontWeight:800,fontSize:11,position:"sticky",left:0,background:T.navy}} colSpan={2}>DAILY HEADCOUNT</td>
                {allDates.map(d=>{
                  const count=filtered.filter(e=>(e.dailyBreakdown||[]).some(b=>toDateStr(b.date)===d)).length;
                  return(
                    <td key={d} style={{padding:"10px 6px",textAlign:"center"}}>
                      <div style={{background:T.gold+"20",borderRadius:8,padding:"4px 6px",color:T.gold,fontWeight:900,fontSize:14,fontFamily:"'DM Mono',monospace"}}>{count||"—"}</div>
                    </td>
                  );
                })}
                <td style={{padding:"10px 14px",textAlign:"right",color:T.gold,fontWeight:900,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(filtered.reduce((s,e)=>s+e.pay,0))}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ─── SQUARE SALES PANEL ──────────────────────────────────────────────────── */
function SquarePanel({weekKey,squareData,onUpload}){
  const [locFilter,setLocFilter]=useState("All");
  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:12,padding:20}}>
          <div style={{fontWeight:800,color:T.navy,fontSize:13,marginBottom:12}}>🟦 Square Sales Export</div>
          <UploadZone onFiles={f=>onUpload(f,"square")} label="Upload Square sales CSV" sub="Daily Totals export from Square Dashboard" accept=".csv,.xlsx,.xls" icon="🟦" compact/>
        </div>
        <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:12,padding:20}}>
          <div style={{fontWeight:800,color:T.navy,fontSize:13,marginBottom:12}}>📋 Form Submissions</div>
          <UploadZone onFiles={f=>onUpload(f,"form")} label="Attach form submissions" sub="End of shift forms, compliance docs" accept=".pdf,.xlsx,.xls,.csv" icon="📋" compact/>
        </div>
      </div>
      {squareData.length>0&&(
        <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:20,boxShadow:T.shadow}}>
          <div style={{fontWeight:800,color:T.navy,fontSize:14,marginBottom:16}}>Square Sales — {weekKey}</div>
          <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:10}}>
            <div style={{fontSize:12,color:T.muted,marginRight:8}}>Show</div>
            <select value={locFilter} onChange={e=>setLocFilter(e.target.value)} style={{padding:"6px 10px",borderRadius:8,border:`1px solid ${T.border}`,background:T.white,fontSize:12}}>
              <option value="All">All locations (combined)</option>
              {LOCS.map(l=><option key={l} value={l}>{l}</option>)}
            </select>
          </div>

          {/** build display data based on filter **/}
          {(()=>{
            const display=squareData.map(d=>{
              const day=DAYS[dayOfWeek(d.date)]+" "+shortDate(d.date);
              const val=locFilter==="All"?d.gross:(d.byLocation&&d.byLocation[locFilter]||0);
              return {date:d.date,day,gross:val,raw:d};
            });
            return (
              <ResponsiveContainer width="100%" height={180}>
                <AreaChart data={display}>
                  <defs>
                    <linearGradient id="sqGrad" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor={T.gold} stopOpacity={0.3}/>
                      <stop offset="95%" stopColor={T.gold} stopOpacity={0}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
                  <XAxis dataKey="day" tick={{fontSize:11,fill:T.muted}}/>
                  <YAxis tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} tick={{fontSize:11,fill:T.muted}}/>
                  <Tooltip formatter={v=>[fmtUSD(v),"Sales"]} contentStyle={{borderRadius:8,fontSize:12}}/>
                  <Area type="monotone" dataKey="gross" name="Gross Sales" stroke={T.gold} strokeWidth={2.5} fill="url(#sqGrad)" dot={{fill:T.gold,r:4,stroke:T.white,strokeWidth:2}}/>
                </AreaChart>
              </ResponsiveContainer>
            );
          })()}
          <div style={{marginTop:12,display:"flex",gap:20}}>
            <div><span style={{fontSize:11,color:T.muted}}>Total:</span> <span style={{fontWeight:800,color:T.navy,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(squareData.reduce((s,d)=>s+(locFilter==="All"?d.gross:(d.byLocation&&d.byLocation[locFilter]||0)),0))}</span></div>
            <div><span style={{fontSize:11,color:T.muted}}>Best day:</span> <span style={{fontWeight:800,color:T.green,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(Math.max(...squareData.map(d=>(locFilter==="All"?d.gross:(d.byLocation&&d.byLocation[locFilter]||0)))))}</span></div>
            <div><span style={{fontSize:11,color:T.muted}}>Days:</span> <span style={{fontWeight:800,color:T.navy}}>{squareData.length}</span></div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ─── COMPARISON ENGINE ───────────────────────────────────────────────────── */
function buildGroups(weeks,mode,customRange){
  const sortedKeys=Object.keys(weeks).sort();
  let groups=[];
  if(mode==="week"){
    groups=sortedKeys.slice(-16).map(k=>{
      const w=weeks[k];
      const byLoc={};LOCS.forEach(l=>{byLoc[l]=0;});
      w.employees.forEach(e=>{if(LOC_CLR[e.location])byLoc[e.location]+=e.pay;});
      return{name:w.period?.split("–")[0]?.trim()||k,key:k,total:w.employees.reduce((s,e)=>s+e.pay,0),employees:w.employees,...byLoc};
    });
  }else if(mode==="month"){
    const m={};
    sortedKeys.forEach(k=>{
      const w=weeks[k];const mk=monthKey(w.startDate||k);
      if(!m[mk]){m[mk]={name:new Date(mk+"-01").toLocaleDateString("en-US",{month:"short",year:"numeric"}),key:mk,total:0,employees:[]};LOCS.forEach(l=>{m[mk][l]=0;});}
      w.employees.forEach(e=>{m[mk].total+=e.pay;m[mk].employees.push(e);if(LOC_CLR[e.location])m[mk][e.location]+=e.pay;});
    });
    groups=Object.values(m).sort((a,b)=>a.key.localeCompare(b.key));
  }else if(mode==="year"){
    const y={};
    sortedKeys.forEach(k=>{
      const w=weeks[k];const yk=yearKey(w.startDate||k);
      if(!y[yk]){y[yk]={name:yk,key:yk,total:0,employees:[]};LOCS.forEach(l=>{y[yk][l]=0;});}
      w.employees.forEach(e=>{y[yk].total+=e.pay;y[yk].employees.push(e);if(LOC_CLR[e.location])y[yk][e.location]+=e.pay;});
    });
    groups=Object.values(y).sort((a,b)=>a.key.localeCompare(b.key));
  }else if(mode==="custom"&&customRange.from&&customRange.to){
    const from=localDate(customRange.from),to=localDate(customRange.to);
    const filtered=sortedKeys.filter(k=>{const w=weeks[k];const d=localDate(w.startDate||"2020-01-01");return d>=from&&d<=to;});
    const agg={name:`${customRange.from} – ${customRange.to}`,key:"custom",total:0,employees:[]};
    LOCS.forEach(l=>{agg[l]=0;});
    filtered.forEach(k=>{const w=weeks[k];w.employees.forEach(e=>{agg.total+=e.pay;agg.employees.push(e);if(LOC_CLR[e.location])agg[e.location]+=e.pay;});});
    groups=[agg];
  }
  return groups;
}

const CTooltip=({active,payload,label})=>{
  if(!active||!payload?.length) return null;
  return <div style={{background:T.navy,border:`1px solid ${T.navy3}`,borderRadius:10,padding:"12px 16px",boxShadow:T.shadow2}}><div style={{color:T.gold,fontWeight:800,fontSize:12,marginBottom:8}}>{label}</div>{payload.map(p=><div key={p.name} style={{display:"flex",justifyContent:"space-between",gap:16,marginBottom:3,fontSize:12,color:T.white}}><span style={{color:p.color+"CC"}}>{p.name}</span><span style={{fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmtUSD(p.value)}</span></div>)}</div>;
};

/* ─── SIDEBAR ─────────────────────────────────────────────────────────────── */
// Premium SVG icons — inline, no dependency, scales cleanly
const ICONS = {
  upload: (c)=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>,
  report: (c)=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>,
  schedule: (c)=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="14" x2="8" y2="14"/><line x1="12" y1="14" x2="12" y2="14"/><line x1="16" y1="14" x2="16" y2="14"/><line x1="8" y1="18" x2="8" y2="18"/><line x1="12" y1="18" x2="12" y2="18"/></svg>,
  compare: (c)=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/><line x1="2" y1="20" x2="22" y2="20"/></svg>,
  history: (c)=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>,
  forms: (c)=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="5" y="3" width="14" height="18" rx="2"/><line x1="9" y1="3" x2="9" y2="5"/><line x1="15" y1="3" x2="15" y2="5"/><line x1="8" y1="11" x2="16" y2="11"/><line x1="8" y1="15" x2="13" y2="15"/></svg>,
  export: (c)=><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
  import: (c)=><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>,
};

const NAV=[
  {id:"upload",   icon:"upload",   label:"Upload"},
  {id:"report",   icon:"report",   label:"Report"},
  {id:"schedule", icon:"schedule", label:"Schedule"},
  {id:"compare",  icon:"compare",  label:"Compare"},
  {id:"history",  icon:"history",  label:"History"},
  {id:"forms",    icon:"forms",    label:"Forms"},
];

function Sidebar({tab,setTab,weekCount,onExport,onImport}){
  const importRef=useRef();
  return(
    <div style={{width:220,minHeight:"100vh",background:T.navy,display:"flex",flexDirection:"column",flexShrink:0,position:"sticky",top:0,height:"100vh",overflowY:"auto"}}>
      <div style={{padding:"16px 16px 14px",borderBottom:`1px solid ${T.navy3}`}}>
        <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAXgAAADrCAYAAABjGI3/AAAAAXNSR0IArs4c6QAAAERlWElmTU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAABeKADAAQAAAABAAAA6wAAAABxFr7PAABAAElEQVR4Ae19CWBdRb333Jt935M2TZt0L23pvkOx4MLyrIoKfk8Qy476UMHt8d7nZ3i+pyIKCgJSBAHR56OKImBZhQIt3Ve6N2lo0rTZt5ub5OYu3+8358zNyW2W2zY8aPqf9mTmzPKfmd+55zcz/1mOUmIEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBAFBQBAQBAQBQUAQEAQEAUFAEIgaAVfUMYd5xL///e8JgUAgxlQzJycn1NDQ0Asfp5/TzTS8N2mNvWvXLv/NN9/cbe4ffvjhuLlz58b5fD4dN1L+k08+2bVq1aqAHd/19NNPJyYlJbmceY0cOdKI0/axY8e0bfzNfa9IuGG4CXOWtaqqynfllVcGNm/eHIfwOGc6k69Ju3z58g6En1BPZxq6DZYdHR06blFRkY7C+lJWeXl5CHlS1smbW+9LUDkJMWpk4YlpLShQ2RODlAkzQeE4cBzbotQbx3zqjVK/CVb3/T1BNeL3kBJv/QbS8cwoIx11am0PqdJru3A3GBYuddvTiSojqdfvCA/DyuZYuV+VXumz83SrFaXxasZUl0qx45v87Ai9rNYImQxk2foyjJuO32fs5qD6+tdZ7hPNsmWx6tzP4vd/UKm2HJcux4mxlKqkZ1VECJ7vaIfXvia/Wtnzu3eERDqBzz2JYc/R1u9EKWMjn0pcGVl2vd7zqdLSYDh+f45S4Jg+NUanNXG0bMhtfRnPrjSaZ2dSntF27Bld+iEo/MGDB/MyMjKK/H5/STAYTAXJ82UM4j5UWFioYmJiXPB3ud3uIIhZ0Y8G8bQ7Pj4+fK8d+MN4SMfbPbi20UGzaNGiOSDNyV1dXZrER48e7aIJwbS3t4euvvrqV0HwNYy7bt26rOzs7EtTUlJcbBBGjBjhgkw3w5wmLy+PeYXscrNMWjbvjT/zQBoXZDArXQ/KQL1Uc3PzGjirvF7vxBkzZsxFvfULxPKbOiJ5TG5ubuDVV1998WMf+1gD0/Zn0FBkpKWlXZSQkJCM9CEbO0WcUF9FWZ2dnYeR/m1cfRNSf8Kv+GqqmjntQlUwJl25QLwuPKtQ0K1c7qDyB0KqEPKCygVAKNmt+Ai6gy4VGxdUhR1BhCn4WiYYA3/IcCN94UV+5X9ls3pD7deBJLvM3E+o4vQMFY/GBM9aBdx44F1B5YOQ5roOdcVtL6pV9w7cSJU+NhZEuUDFOZ5bMNalEuJjVCzyD9buRH7W76P0sRI1ccZclZIeB5BQdvyuCkHYATfqhDrQuOHWBuXhM3Xr35jlxb9B/eite8b14T4W8eLAoQH0M5oWNCm17GW0Zv5wotLnktWMkmLV4R2jAsEUFVgYVAlxMfqBsYwulNcVAr5+5I26lyBlaDrQDYV0GcKCXFZhGD/jaLlaqd4JB/XneODZ8WrUxLmKxaY8v9+t3LHWE4qH1ZUSUsXnBPEsQ6rT41J16ZsQs7w/cba/SxUuWaxGjC9UY84JKD/kxhCHuFj9vJsv9aqfjHxd/evNLYPIGRbBZzPBu8rKyiYmJydfhV7yZ2JjY8eB/OL53sDmFaIbBpblB1t7M5xu299ECpMVw9EgKJD2vyG9IXh3VlbWHfn5+Zey8QgLhoPxQYiduC7ArSZ4lGtJQUHBI4mJiXxxdL52Glr63WL+JGlzT9t5T7kmjP4sE8tN4mZa9LADqPvFiFOVnp4+HQ3dI4xjx9NxEIaoLrx7flVXV3c57v9OmX2Z119/PRaN0ufRGNyNNCnIn4VjQ6MbPJJ9d3e3Fzj8K/xJ8CdnpsyZqEaWPKiyC3KREAyA+hnUWVk8JvjjeeA/GmYtnEUgDn44XEE+GPIUwkJoCNwWSXJElTPyFsS3CH7B5/LViJG/VElpI1VcHEBDCBuQUAAgukIqPrFMpWS8Dt+BCX7i1BWqYOS3VAAAKuTtYqOCVy4GBErSjkv+JmRYv4/Rk65XxWO/hjjxaLD0g0ISFB75asNC2E5aFp9aZbOqTBq0IoQcRE+Y9AOFLLd7q1pW8g80ZBbB/+qZHJVfdLHKylmB+p8LMfEgeZQT+cegHG40TKB3Sy78aXQeAFSXQfvgD8rmRmtDPzZm3vbfwjUwwV+BXnb+2K+r/JE34HkEwvIAMdx4LswZDbeuNvIjXoHuOyD3V7j6N4sXJ6rMgh+o7OxFyp/hRwcgQOjxHEHwwN8d16AqEi6CACH4/lE880N27tw5GaR2F3qbnwKxagLTHIF3EWSvSZe1JLnROMnSee/0Z1xzj156ED3V3Tox/jz44IMZyGcGZMeT8PjO0RjiRQ/6GNQlldoTf0C8izMzM5MYTrmR5WA8Z34mndM24SwTL6ccukHwNY2NjWVMU19fv724uLgLDUo6689wGqbjiATlUcBqEbz6JXiMHmZjhPJ/ETfL5E0ZJl/I8mPE8ALy/R1FM+ykTMmkmaoA+h5wM8gZgnXbZ4uwyhuW57zVOZnsGGC7afHW5+tWns4KuCwz7pwZ6EmPUQlJIDr2YOGdgD+MzwFbfU25ery0zYrcz99lK6CayfuoyspKVp12I0MB7HX7QcAdre2qq4M9UqVuKk1Wo0suVmkZGSpk6uQsJ90waFt6mYhbAN0rWBcYWet0QfTgu7sPqDce79SRfgRyzxt7vRox5rsqMzMHjRDi4SIZsoFgObU4W6Yzb+JPb5O/joIbDv4YL+A7FFGQE28/Pn+qyh15JRrKpLAcxqJMtms0GMjofCif19GUDHoPaD5yTYHKzJmiEpOTMBqw0uE3rOvG+rW37VctdccHlDGMAi2WGUYViqYq69evTwcR/QA9zU+RzNijpTE9V94bUqSbF4nWXM57pnPGZU+X8UDwDZWVlVTRaDNz5swJ6JUXMozxjSzaNEi3/4Ybbmimu7S01A3VzHzGc+Zl0jrzM26nPCOfael2GhNGG43Knj/+8Y+1DH/77bePoHd9KC4uLlw+pudlDMq0EPMCZKATzI4dO/LRey9Fo1liAg2xUwZxaW1t3VlTU3PnvHnzvCbOSdnZuVB3gGHZi9ZEBOxYP2JIUqKKghcJgv7hi/e4GC8cDncAHVmqLjo8x1XlLqv3zgKlpZ8H+W6EIxJMN+KZPDj4amvdCN8eYHSkiD/nnzcSDcQ41cG0CONjoLhu5MeRgKe1QlVut4hwwtRJKil1kuqA/odlAlYAHhfLiMvUy9xTjnY749AP9+G4cOt4yA/tl/J4AmiY1uhSLiuNVaPyrlYjx3xfpWflqA6Ea2wggwVlWU3+lEe3LjvCu+182EjpOLTpD7u7Cz8qj0+1NWBSYwBzxRUxaFy+DN16gS4bZWl5kKOfHdLCqctk8qVtDzwGkKxUyeQ5aJwxwgO16bKjMvyt6N8LUno8b6mffad9QBnDKPCsJHj0MD9BVQJJjqRuLkNItI1x+tFtCNPpb3rk9DONBNQz+zHZGO4pIL85IPgExqExhGzc6NVupTfvp06dmoOe9DlOefR35sn7SGPC6W/ckTbDTB3QM99oJnXRqHSiDLpHyXAThzYbQRqU6VxYefrG8QeqmUSod27CdamJH5kvRjN1mMT9j9mzZx90JI3eeemlCSo5fY4mExdUKzTEkhdHQ+zt0dZuKzgsXMeLCNe9RLz4HAV0dZapA+vqdHzq3zOyzsOQhTr6GE08TpVHe3tAdXmoOx/YjB43USWn5qG8+JHZ5eTcAMsSC7UPe9M/uaNJC8kfNVslJ6eh8UIhHcbUzdTL1I2jF7opj25zr8NxH8aC+TIObJ/Po2qOWuqg80eMVdmFt6us7FRNzExPwwaReWq5SGfk0dYXw2y3s2x084rDaKfDW6M8deGOjSU44u+cy0ajQVtu5YPfFmUzT2d+zEfnAblw6hFFyFa7RYjrdVtQsAT6zjiNeQzVYbZsygp0YzjaYWHQK9HwvSF0Z5XhShaQ7Q0gq1j2LE0P2hCzAYNEZUiKfsZtbGc807M28SgLBL9t5cqV6Bpp40J+S0weRgbzYBlAfqG2trb1dlw1fvz4KYjPGVH87mPCefOelzM/k8bI5D3j9GUYh4Y2eush5NtLT+pB7wY9bSTvSW/KzHQoU35JSck0uh3GhXmFZdCrfwNyaXT+TMeLsuDXCVXQo9OmTXveke7knNMvLYIedRxqh2YQPT6SwakaqrV5kfwIibdpm3rDnng87wsFUBtMwSQjgUIge5WwiQl7gZ7WenVg+8AExnLFJ8xTScnQxYH0mJayWGYSGecRux1Ek5SyAKogKw7zGDLDfHHFYnGU23VMHX+nQoseP/2foHcvssYgwFIzKCyWkWXt69IJ7T8Mdxo+Dxpi5vXsUN+5ZqCJeJcaO+MKqFHGWo01MR7AsEz6QpwYp56ojzRzb4pTsQkz0TAjMEIuG9l2j1cdOrijj5TD1us03pIzExOQzMjU1NT5JFZDRoboncTG2vHeJigdl35MYy7e05jeuLEpDwS6wQrVKpck9N5nG/nGJgGSwBG3Dfr3d0186N7nI34s4zllmnQmHstBQ39TF94bUjbpWR7jZjgN1DMt7733XjhP+tXW1m4BybcZuaaeJl+UKQZlm8O4xmzbtq0Yk8H/D5hiWNyDGfPkRS+Q+5sVFRU/h9tmAnqfpJkyYwZUJ7ma3HWBnCRjk1JYpDMs7AkH/KmKoK17dsDPj16dp3N7OFbxlHNA8Pla/UBP3RumA2kId1fHYbXj5aP06ddoFcSIJZjUI7HygfCHY11sWLzQ23R2bdbpr7giHvri6RhF2GVj+ey49vPFw+vJyvjRx/izXDS8N5dRubDBYMe3s2O3uv/+NsURSnb+ZSohka2vrpauF+WG1Vhs1ByyjDuct10+nSncVCsFqVrC422q4yiw/+f8098WqIzcFSouwa38WCFK2WG5WuCJfxCF8Fvj2xODwz5Xz83FyGCajoti9RgmZjl9VWrPW+U9/sPfddYRPFeLQEWTyUdriIxuQ4BOoqSfuUyv2dwzDQ3vTRpDiFB1dKBHHh4KQudcDHIca+LrhPYfEjz09Yc2btyoSYP6d+jBLzBlo3yavu6dfqYhMOWkTWPSG7fpWaP3fmj16tVVOpL9Z82aNRXw32/imPoYm/6ox3xE16/P7t27U7H08jsg/cUmH5K6KQPFYiRTjgnc71944YX1djanZmXkLVZJiciXpIlLv/T4Q1JiQ8JfsovEhEvXHTY7fFxZqDt+jItLEzX9IYqE1NXpgUqB6jHLpKTMAjFjaaR9r8mZ+SE+FmIon3ezWr26y0Tv0571iVw0EjOsRoJ5GgMZLKfXU6eOl+3SvsUXjMLQaKy17JDlQgT28EnK9rPX9QmSqZ2y4DbEiqrqMNbJAga2jQVt6sc9rRythZQeoaRN1yMGyufIgXkyHe8pgmlom0vLZRQ7DqPr+MyD2BN4eHg7Auglc36if5NedCkmkyfr56SJHTKdxuThtE3ZdCJn5Ah33NipKikFKkTI5LPlyIIX5zTwCDHBuk39/n5PRKphfYune3YZ9DQXYTULVt1aPyySFg3vbVJshM0XWP/UdSD+mHiGOPu6t3usCuRevmXLliMmLTb6zMA68FRzT9uQM23q36kDpz/WtSeD4EdDVXIceZlfP9fhm/LRj1cGGocUI4flpxt69Xak54+Ya9Bh6bi0dR0QznqF0Ki8df/99/ciKpZhxYoVmyCHJK4NZRjZ9OBKIGzISsZmpQ6UaTnIfQUbKU6i0pi4TIfwVkyq/gyjpoFfep1ygD/sEScnz9TkyJ43n41+foCBk2+xYHC/rxH+mOGGH1dOc5kdiYlrYIgWcaCtaw8XOakDc70NNbvU22+V6RiACGoG5AM2cGoDmIZk5MM6+OamsCrNTnOilTziHKhFRpAntWF6bVAAYKU8bQfU96625mey0ktUfHKsinPXqC6wUgiTuyG9F8OaHHXHcNYeyxcDGSolDUsodQ0gzaoIJnybMUnSgQldED4qz4YhiNaIhIafNCYxA8rX4scEq6WOGzvxXJWahtEWcGRjaEhWY4S0oUAb8mpF44EIpuA6krlBHCREKXX9mI57wbCqWDXVNqim6t26qn39+e5daaqg8AsgYegdmb9+lrAdQNnvZe/k4ax1KXuHOe7S0TjHx2MJHEhdFw70Zn4nXdCWNjUTA5OZI+HwdZ5VBM/eMXrvmrxIhiQjGpsISVJ+qCl+h7AqeMPSP2xG0Zud6ABphXghjf7VwQ2ndpOEdRhUEnu+9rWvhXsKmHxcbCYqEUfnS9FGjQG1SFidg6WIIRD+49XV1dwo5Gd8xOXrSvlB+FF3zs1En4dqZArLbupBom1qaipD+DPIj+TN+PxBs6Isry4zwgLoVb+C+xMMetzrAMNXkFf4rWIk5sG8QPBjxo0bV3LOOef4qZqBXp7lZBm1LIdNLJ9Zu3bt4zrgdP6MWpKPIf2UXosoUBarWljW0hnwqKOHfwNCbwGoIHss9YjDBCnXc5Os2R1ndFI97SBIj8sfSdjVFbvCG5a+WZquElNnaLQQrZchHA1Nbari0Lu9/Pu6yc6ZDxUIlyMhz14wWkVurnsLyawWsTtQg2V794Ecu1V7M9bVY/05SYomLs7FIZDq7pyP1SGXK1dagvY3f9hk1VSuB1G+okcdXR0+NEzQ+7POFBHw6VUtXm+rqi+3JoYzCuZgbiBO+Zi9zZcsI3vyPvQxKsteQ0/3NchAY4IAcnAMysOJYW0AIHoH0CuibhgecYMSy8kJ0MZjlerVJ49Z8fr4O2HW+Vhvv1TnZfT2VFnpcqAMbFQj8XKKMU2106/H7Vbp2QtRFoxYOPWF4uLx62fAZ97l9anWGr2IoCfJ8HdZP6ThX09dQ2yTzwa5TTeESMKiMaSEd6nuueee+6+vf/3r1ooKHXp6fzipy0aFRG2MyY/3tjonrCK45ppr2uF9v4nbh+3as2fPN9FzzjdyWB+6SbSoX8rf/va3lbfddlv/L1ofQo0XGqdNwMGDBiTN4GTCeA8VTRImqT+O/Kfi0g0M62bi0uY9RjFb0Uj98Nprr7XWXRshp2JPO3e6Sk4bpSdGIV+TAHtpLvx8Y9B9bDi6X10x4wcQ3WtEctJZTV40SaVlTuzp2doSOCpAe6DaWw+rss1HBpHrxi7YxXrCkaMa9lS1gQy2r12dARD6OtsTEzTXccK2/0nbu55NUyWjl6Fc6KIjvdPEo1Dd/kZ19YJ7nN4DuDFCweQvl02SSIkljV6CCJsqn7K9v1Xfu/Jv2n8o/9x2T5LKHXOtSk1PtkdZmoM1qTNfF55dF1b6JKXk9NTTxsw0ACHnsCqicDc9nIYln3OseuGBWe+bBRjr6vPVqta6QxGphv2t+fUN+4qygliiNwkEVUBid/Z8GcbeLyY797z11lsc6g+ZgXqmEL3eyRRo8jV5k5Qx2XkUm66MimDQfO++++5krN+/CfXINgTvtEHw2ZMmTeIyglMymzZtqoQap8wQNoWQsIkP86GNhvJW7CO4mvcMo23cjI861WBJ5J3z588v5/1pm4y8BSoRRwawQSYp6YYZP12qIziR2d7ElRGnR+4sZEr6fKgvkkB0nHxBPnZe5BW6u7Dy5Q8P6b0K/dbp+tJMDHOma/UQiQnYaKPLjPv21iZ15NiBftNHBsSHFmL37nL0oKF6sDok4Sgc2CWn9LFkJByjt+NrP8rCRPUsjZ9WQ6GBZCMZi4sjn2CoTcXHHuydaIjups6Yp3ILPqHJm/UwA0SOQoh3W9u7GIW9aHE5MCNuxJxxTUOEV7Tf0kxOH4fVM4V6hBaDEQV3DWthkMHG2dO6XW1bPaTvdr9l+RAFnFUEjwnW+VApxBlCMrZ5HlBtbDLrwo3f6drUvyPPLPauaQxxmp4uRrs7cSBZ1NumcVhZCch9tJFDmaYeJFvIOwCSPqXeO2XdfvvtHejBbzXyjWyTD+uBnvs4zCkk0W0uxrcbrg6Q+0NYy7+aaYbAuPHizsT2/ghRfPlx+bC5xuM9PR2/JdkF1cVcakg0oJpUIJ+GvXAeZ9DWyJUvtqcOOfHPpKkToA8faY02kM4QFFORtNpbD6mKV6tPTNiPT14hG514a9LQQXbszXOE0N46eJmM6ElTJ0PfP0JvFmNZOJ9BouVFd3vrHnWkosJEHzK79GkcSzDmevSwM/QghL1r5s8eeQgszsblaPmTeM57dINtiFnP3TKeXRJXbEQL5yhhej4mx+MxykFkI58Trd1Q17F+3dhn0nOQnyPh8Hai6T5rjIs7MUmChoxMzUliICquRR8KojBitc08QfBuTkI6yZKBJEeoaNb2SjDIDXrOczGJmkIyjZTHexA8J2yx/uzUDXrga4HRdQYr4mUu5mFGIJE5IE4IRxG8tn///vsQZl7LyGgnd89eZ3zctEjthNYzs1fs9fpVtzegnj04G2pXvMkwMXEYoiOwV3/PvjE62g7o39945rB6sNSaKznvU6kqKWMm9PMgCPSMNcFDFomI+XS0dqjm2sF1uPmjFmIyNFWrOwgBjx4gwbBktLu8W0A0A59hwzpoU+rGpqslugxW19QqD2WR3Dva/aqxIfrf7IjRC1USZmOZ3mnYU+YyIU9rlbronyapj12Oh0wAgAVtN0CIBaY4skJjSjfxPXasUb16f+2gxDk69xz0Ci7TSym5Gok4UDwnjKm/b20+rLa982d1waU3aqBYHJK01qMjrqF1HnrWn0nPOV8lkt8p11FBLoP1Qp/X2jbw7tr+5J7h/mcNwWOLfQp04XoNtyFGkpYx6LW24jhbayLKeJ6mzW39IPdFRozJz+QPIu3GAV7Rv6AQhN7zYp7MaGQZ2WwsQLw8lXLwVR4mUT82JmC3jR07tgPqGLwxfGd6cDJJTB1MQ8PGAA3kwUOHDv3gk5/8ZJOJd9r25MkToTop1oQW1mdTKsrEFSKJ2HFaMuU7WMIIFY2OYBUW/UKVCObiRCB5IZCEFSYglBDixMA/6PJhKeNXIcia//jIxePRUx6v88FZYL0MCcOHNdRlhwfT4bqwO3Qelj2CyNFztKYoUVTkTxIl2fi6wstne+XR180d8Tmo+7loxEKavMwKGpIk1Uad3jrV2Xywr6R9+Lkwj2Ed9UAdPNUyREqTIRwk8izMHajgo1oPx3v2rll4oIiImPRNtCrCkyqh/VfB/PXqoltWqWnTNqv+j/F1q6zCq1VqZp7GgOSuDUTzcXERQE3F79W9tx9V538c2aPhYl5hQzcLChPE4WN9mcW3JamMnDm64WCRacxvFudMqbaWZtVUPvjkuJVyWP09awgeyw8ngGzHkAhJRk5DsgLZHoL+/ajT/3TdIMh8kPE0EiTzMMbkD113zYEDB/Yb/8Fs9Mx5PvyCyHiGgLEax4tJUg7ZT8uApA9BxfIe8tKTqCy7sw6mLsafmWHk0ATVzI8WL15sEeZplcCRODlvnkrNSNEvvQND/c5zaz9X+xSMmhzmBPv91vc9kIMcHDLZ+01IbFe5uZzQtkzJ1Fk4WiBD3xhyYH6URz1/c8M29dAdA+vfOQpISJxl7ZLFGkDqyCmLMrissLWlXR3ZFz0+Y+ZNAqkWYMKWpbdOa9R1gkCORPy+3Wrzs9Gp4z7xpWTIOkeXRQ/ByK4sGAzfBx7/kJWH87FdhZZnH3/t6HpvAYk6KeN8yDhPecZdhdiH+0ih1PceLlIZGZ/VvXEcxgpC51JOrPlFbGpc2hqrVVXNUzot/XjpOQ/t0N76DxvJkF5y0+NnXB+Zhr0ESWOskYHtqX8rxAmrqTwtZeofa6xlqSbNWWL3ZrphXGlMDH4EBB/ukZKcDNHSDYJfAwI9/RUfDgyxnHABlkjms1Ex+TlJEb3tddddd129I8mAzo9+9KOT0WCEl0byPdXvKt4b1gV12Ll9+/aoJ2z7y+z6669vQ2PxEkndXCy3uYwf86bB5LQH5F6KA8esF7U/wSfvH6Mys5dpMtNkZBjGFuTmRBr+873nIWC86Db36DArfaGc1MdqnSxscltr00a197X3bEkurO5YqnXTevKRhAe5lE2bPeiWppfgAYEDmEs/P1MlpozTsfRJhpBDgieB6iWPbTvU/sN7B5DQOygt86MqJSURywrZPNhlgk3csVJR1Ve/NKh6xEi8ZPk4LDVF2SCID7KXsevJPDR2kM/NUebiKhte+oAz2J1oIPlJAx78VlS8SI2fd2svcT03bjVrwS04FmEslmDyx8SeAn9IaK5w8dsDx488pL73aXvSGQ/QFC1cQhaKF4xLK+Utt/PvzBkXqoz0TF0+4085lMGln21NL6rnVmLTw9ln8Ms7OwxW/c3nCpBIw986ep9DotqIlA1ynw9ChgbT+rWSGGkMKYNEqU6xf706aMA/WD0zG73qZKanTJI6L+PGiGAjlnjizT99A3XVWuJiJBlSN3Uw/rB5FMGfoHd/il+GcvifvvOrpXkqI3u2Vm0EsXCaqomw4QvM4oFwBjVMx586Ljr5PNox3/K4fXTupVf1LLHj6ZImjk6GP5h2VrWVg59hUjRxvkrHaMOZnm72Ptm4+Do345jhKDsR2NyVmY0zalBmvx9SrP/WzwX17sBxB40YVURrsvIzMRJJ0eXQREvciJ+5KAhu/kb1ZcLhTbwiLzYyjBeP3cVJyYsRC0OKCPPDJ4tVZtbndDyqhPQzYEWQjng0NlSro+89gwB4wPipZzQ/IUTT5aHFJLj44Za+TGLaefpDIb1+H4jIxhrHdqu2QXbX9iVzmPidFQT/6KOPpmH54MxIcjLEi55vM85l2T6Uz/QK7L7E8sgFJk/aJGZD7lix4wfBn9TED1bPLMEEqyZ0U1ZTB8gLQd4643+6dktLy1ZMADca+aYeLD+NaWRwBPAm6Ox/fskllzSebp4npJ84Zyp00Fj/jpc+yK83kQccF/Wtujyw9RCepGNf9NcX09KPxMF4IIouHNDe1Ngz9zFjzlisvy62wm2yMVYM0nR2VKiDZeUnlK+3hwtLELEKB2vyyE3GcLTAeYAufFGKW+WjNXcszUaZpoX11lSJsAetdcyQGeg+rpoqoh8NeLvr0VA2oGzsJbM3DVkoDCzdePCe2JFgNaasBC5dFx2JES3DspCsjQw3TveaNk3HNFG0XTLpc9h8NE6TOZct6mRIR3JnQ+pteV4deLksnMaPCvJ8Gqdh1kiiQhg1UO8fab707RSVkoqlnwiILAEbxwDq7K3pf3dtpLxhdk/ohr3BWewlINuxJCknYZl7rvsGSVUOJRAg+GyohPSRv8wz8gIh10HVvT/aPDFhm4TRwDzGpywn4fIedWg5cuTIkDVSOGemCmXc6yy3ydPghgagGkcR/AT4vj8TWBmZ87h1Fj0xDFc4pofhi6wv/sHbT/UHSTtM6PDThA4m5CQhC6snC5kYhoM4D7b3Vx2wzoKh37iZi6AKSdM7IFlhGs1p+EO7rfld9LwHXsp61a1pmFw9V8fXPUktgILwo8O/ttYW5Bn988kdOVHFJo7UdTFqIwonqZJgve271Bt/rmVRozI71r4HNdNqkCVWyLCxtIE0DSLl8qKhfJIjGZNetjeDrBs7nI0X43Z1Vqjdu3uT738+MUrljbgKB6lxFyyHm/zhIj3lIqOuzuOqteFx1XPiKuc67J3HzIdxcelRGmz9fHWhGdhjpi8Zo+KSSnR07esoMOvT5durXl4X3TxFj9Rh4zorJlnR852NK9k8tfA7jB8A3dAhbx2SHZcmA9j4utEUqFPyTY/d5Gn0/lB/7MUEa4MjyYBOqGfGoArjjT6fkQ3RUjbIuAxqkiFrpKjqufzyy/dA9vmm186ya760iQArgP5nw4YNQ7XePbL+2BGaM9cicL60qKR52Q3RsUcbQtec32XVUfBH9061KDKEZVhe6m9JRgxvbdyp3ltr1qK7VG4hzn8HEXVCu6VVCRSGy+ijG/RhZPTs30ydV4LJ4BJkgKQsF6LaXKYTdXUcUkfKenqr/UuyQvJHz8fHt6HkhiB9ZotDJnvZniZ8Q9Y+4ngwWQxfWepV8897CF+YKlCuOCyXRINmjMYHowxdaHiSuK1RE4iepMvTz2xjYWndEE+OKDzNXD7KGveYMRMuU+m507SHPpbZiEc9uAO3BnMgr/7xkLr1vnQsHcUOumasbopL13p9JuLT42WYm/mwTJFmxMgF2F6dxjZDG4jXhiXmh1raULbVvc9csmOcFdZZQfDUv0O1wfNk0BHAe27rrUmMXJ+Otdsbhvppg5DnoMeNFxQ/UeRjjCFlLCnchEld/AKjMzgeYCYaDH18gGkkmNLIhpppEzcpRSctqlgsdCpjmjLTpmGe6L37jh8//vRQN4w6A/5Z8c10nDqIden2i0sy0YzJNxjujvZWqDxeQOnYs2a/HA1AjM2CsDgxqc8IIjGg4G4XZrq5ZBL9/5qaN9XvfmdNun3p28kY4lvnz/DcdPIUGw7afG5dOKWq7vjgPe9R42bhHBuLaDQJWiJ0udij9/t2Yc19z6odBPRvSt3Y9bkExwpgaIaOseFeqpnQlkHtEFTezuhX45iMjq7diG+S3qnyRl2ISd8CqGyIEeRj+aGuMyLq82XgF0CDyA/Qx7jjQbqfwvcai6xnweeAyzRiVD11tPSen7jjVzkqM++LGNHgPB7EZRL9GPHHELHLnak+d+s3sOoIdcKD9XXFoNG5wApnGsa1E9Lis+hrHXx23vmYB8AHVLARjR9R5zws8+LFBqG1ZcjfbUg+Y8ywJ3iQKHlWqzYMQRmSJ1Gi994BPfLJvyyDPGKzqcpEY94mf6hTgierL8cu3AWYR3CxB28MiZYXG6mhWP9u5NKGSigdKqY5psy0TWPCcIxAqqESslc/0GeITcmMcSDMIt3rJqmRJfjSk/CpPmhvParW/+M/cIRwK87VwiZ2qHCSQfD++BBedqggAkxkG8xrxidawPmxPn7TKjYKpACsAplQrPXvmuiQB33ZQ+aEPEnF11mj6ssG13Vn5CyEisFSfbAHbFiNqgk/RgZeHDNs8oRjQPOvCRlQbcy2SBdlYMOmF5DQDQCgjoN64+T1ytZa9XfUTaU71KKLUnR5lQdYQmwChitwapOAPQP6Bu37kT1J6oJPLlHp6UU6TDe0cLENJD7e9gZVfXSfDjN/Cicsw7r0+RqD8KiLgawHRHN1TvYIhPuxdp2YISiEA/ti0CHi89UjAyu6xkA/C6QNmlaIsmAuvTUBI6+Z1qYy3BMePdrho8ZNd1eb8ja+P+pD5HAmmGFP8EuXLh0J1cZE8zBIUk6yBVFVHT58uNyED4V91113pYEcZ5q8jG1kc1IXuuuoN1XddNNNPLBsIeVwJZCz/JSJBqMTp0gO3ss0BYjChoppIvIsYV5GNcNk5h4N4y4cetYchahTi5JbOAdrm6FWs0mBUkjCfHHZo2337FOl13IOg/RwKkbTgSosWowNTulaLiWRnDmhp/OBu7Fut1r3XM2AGVzx1VQ0IPPCjQ+JTxvb9ra1qqZj0fckM0pK0BvF5DLqS7KjGJaNunhORXiO71cV/6iy8jiFv1TX8IrGfOMnY5BxgS6L7n0DExaGdeQZL53ePWpn+dGwKH5AfOSoL6JBSMGPxRoNUcNjMNHkjXQJWLJMUTSsm74olw4a2oiH/9rQ7tY/ANsD1uWfLkYvvxguflcSMZged2wkaHd6D6rDe4/AddaaYU/w2Pl5Lsg2i0+YBGls9uJpqCq5+uqr2/TNEP3BZp+x6G0XmzxIijTMnxcI+dDBgwejnvi5+OKLC1CHyZHFo3wSPtQllSDbisjw07nHZ/jmIE+Mn3twM/JYH4x6NmB0ZIFoAobSzs1fpL/xyedkyIEblNiba8eHJeor30R2FrCnk29+0XnQCWMoZwvhs2Kvm4ZEUX9srdqyxYRa/pF/z10yFscTTEIvFMMcMrAxcJKwvO3Qv+85aHwHtUeMwm7Y5ERN7oYFWVOqRaiG8ECleO+9Q6mO679IMxZg6WdmjtW4AhfWjqNINoQ+2N62zWqV42iMKTPmq/xRy7R6xahnekm3MWFdoAGyDCsHw1utlrHd2oMBMPodivi5JWbMwSR8thVmy+K8iVYf4b6t+az6wLYFVO+/w57g8a3QhSDb8Ac+WH1DtCQqkG05jrWdAlsjY6vNe6Nk35k4vGU8yOGvyoUllvWLFi0K9/Kwa3YOVDR6vbrRlzMq86ONUcOWk1mvXlxcPBVzCNnM18hwukHwO3BgWSv9hsqg934+Gw+qhJinMXSj/H4so4y+R2oSR2uzR5yRg7NlQCKaTJgQbqM6aW9vU4fL1kYrzo7HZwUhjkbhU9h5mpYJPT98bX7QoYbgeZBZd1e7evilKZiwo546BEIhMUEphAfJXbRBbOCJTf4YPhWXAUIGUEa/YOfK287OHeqn3zMKEDugX8utz0xno0MCDJeLzwDgd3bwcDWqe6Izdz+TrzJTsP4dZfd3u3AqJ9Lpdhu2WZKP+26E0cSBIentx+auhDQ3VCkXYRITh50hf/7czYUvIqBeXPrZ8zu4FSqTksnXYLloNpZC8oPCVgXCvx/IYKMZfgQIJj7awNbxGA5DoudITefHe7h59rzTJCfibPsUPAusZeXvhCMcPkw2Puz9tDRVqHtXlyg3d2nZJsbOsNu2DRR8tjSVu5rVXf9KFd6wMMOa4LkWHeqZBZFPyhAWfwP45NwFILJJIOLwRzwi45OkTW+c8RiOe/2VJThd2OjzLOzf058G5M9GxUXduNNQDhsJ6N97XgpnhH7c2DC1CJ1p/np1A2HKz3uWCyqfd+Ds/eNn4CmaBx54IBW4zTX5mLpTHBsoNChc4rnnFMUPnmz2kjFQz4zXxMvYfD15DAkJlHpuv69elW0tG1yQjkFSNyxCNw0lBlXhnGJMZJboddmaHOBLQtFIwqZ2uGTa5ZiQxG8IrUsAE3nUkcTh6F4u/eNogis1QoEJ1tkuYCV9loqdDS12HFrqo9e/f+8naVD3zNUreEypTfH5UY76miZVVdmzxBNZDGBi1IQpN2En6UTlI4EZEkQ5yW96bsPmPt2o0Q9xiAD13W5UMDbuPI0WcdEhBIfxUO8gvv5UdWSHFYC/0z49B99b/aTGRYsFALo3jbh6XsLERHpbnJbNyWNiTZUUcdUG+fA7Ovh9W2oz3LORNYbflk1ImGOfMoqhMctlB+tniPviKZeqUd4ZOp3en8B8NbGj/oyOONT9c2KZp5WyrImp6xHySI8wuM5gM6wJfsWKFXno+U43z8f0onlv3OipXkAi470xJszc92UzDomPaUHma0wcqC0SQY5znJOhJow2ljO2o0GIesMLGymMBhYZ3buzbMwb8rrRm46+R+csTD9unJs/Fg1KiSF2kycbKBoQ/C4syYx+DXY/+fTrnZnPnlnP0jcyLnt0VH/w/fRC/15XFs2IhQ+VFwuuG0j7nkzQpYrGzkXPG/p3GkYjSdDYNlfi5BWcr+/58lvHbqEMEEUsaJMQdccS9zyJkobqYBqSmtfbgd2a0U/i54wei9UnY3R6UxyWh79P9og9OFysYmu5Dh/sT+nDWVi6+WVgOV4lo1BGdcF0muciBGhihB8Jlz8ua+8Bd+hZqif6W8YK6/KWqboNldqrtDRWFZWswLkz+GA5fDi6odHkDSeJmgSewIVlNj46AvzYVecn9dhgGuwZR5O2jmSlD2L1kDFL/jkPx0JMCouiSI2XLZvLOzMzP6G6Eulr18XGkYm4m9d6cEYiX2iFnchJ8FhpPM90e1gTPM5inwqyzTMPib9ZY4ybtnFHhpn7vmyTDgTbiU1SG00cqGqKQI6TeG8I0YSRKNGDr4T+/bDxG8xevnx5DkYE55p4pqy0SfqYQ6itqqrqvYrBRD5FG/MWsyFbf4qPedCYurAOqPL6k1niedLFyMEHtnHEsiYEk1g/OryTJImm+nfU4GvASep821kBXnTzoqHtVqPG8UPeeAcYbLgDQXRqcoEIHuqle5HwZ+XZVmiiRBwaTYooHInMbgCtxghR+XZ1dVSpY7uj178XjJ+DQ89SesoAOYb0+Pv1+Xaqx34anbonJWsizp8ZhXXsKBsLo0GETUO5NPQjVLCNl+7Z6xYV3gCDDZUuhh2B34zFIAaN1/bwXEDxeeegMVmue8MupKVYfeEPe8eeplr0pFtUMBl7Fqgo1wCz542RENantnu6Mdmdo9zxmTqdzsoUCLIISIjbim1TVDwNOBWYWysN4vN56MIyPqLzKAX93ODWWhiKMPHowD3j8TuzQVQ8JoHvEiMNC8OnPmwNPrA9BeQYZ8jY2CQr/cztmjvdBgxDpLzvK5x+jAOyqzx69Gh4uSAONZuJfNMZ7kzH3jDzRfzt119/fXQvKPKeMmXKOWik9IYpUxaTN+VBH74XX4RqZNhQGeRHFVNY/+6sB1bPBN6PfQPhsvPTbqkZ863hM3zN+8p3jjzk7ehSzTWDqbhMKqYw5M4s6E8Ti5MkY1TOiOnahy84SZ15MEav3i08jHpBkwfj0c9wANyMr8NsoqT6guoUcqSndZd64MfRrzbKyFwM1QMmaqDaoQxDQJTfBeV4RztHfyZzFqZ/U4jNUglYaO5C7529cPaqTUqI02XmPd3WHzpsfwSwHmYHMTHSRidEGEja5+k5Hz+l4PNoTEZYahQMbdjYMSr1477OTlVz7FcqOakMun4sacX5MAScxM+1+An4Rqy3md+P/ZwqLPmMDmN9mafBXI8AHD349GzUDRvBdBzkw3imIWI5dfvEAhjDcLqNH2xdJ+OPAJcL38Vtorpz2JhhTfAgqXSSoHP3p5Os+BQNWZonStKOjGPCaBviZxybsN/FIVvhVTggx8Xw55sels00vKiTh3pjLcUwPBrD9e/Yrc9xrTbO8rHRgLzNQ9mbpooJ+c1n/Zw4GDfW2zeiQYtWB2yKHb2dklwElcI43WsmjBopvpS4+N3nxvpatfVNLo/szxD7yItxLcyzs7n5IU4lZCbic3e5evJQT+bZ4vjSW48vnMQSx+Qsh21IINrA1sTCe8SxevtWZG8HNklVv4iAsG7DStPP3+U3JWMCxzpXhbJIWnq0ADLkbtBOb5dqrIp+OWxqNr5RCrWObr1QYMqzYQjXUftFlIfRWFc3Ry8oBzGxvbSb+xDamqugenpL+//7Y6Mxsfpp3TDhO/FabWUaSc6Z1FWvV6/86SEszay3BPXz9+ntE9FYfkbPDRBf1t1gzklU8/LRMy0T+ndgoifhYTvJneL7rBd/FnZdGMfUi30AHuHQ0Y7PKVYNq3XzrPGwNSC/Gk5qRurDDXkZ0jL3TiD68mM4/W3+1tHRI2cvhr9ExfXq0JfPM3LpZ+QwDcrShe39J6Mvd2GOQK9/d8o0v3PIC3G5IvMZKnPhhReORp4T2XgwT+ZFNy8ajBgO4EjimqHK7wQ5+aOmYI00Pu1mc7QhG0JM0mjBxpUNq+tOSGd5MDYT4o3VF93wyzP+FrknJcUof1u86vB6NUnwRdeEgmiMaQy5gP70cxKG5ggdyAjWxR4947Bx0FjhvuH4BlVdvhoRojPTZxSp+KQSvXGH9dfkjqS0Saq+rkpVXh6duue672KyFt+G1aMAu1wsqpPgBisV600EwwYebC8CmOxsOP579ex9VlkmT/0UjpWYoiO7MVowWPEn04Xe+/GjjwxK7syDaqRwGbXDytn4hawFDmpFaQY6AfjKFzFCoMbeinrCX/1sEUfbKBDtsGHlcJln19b6rvLv6e+3FU51Jjl6Pb4zqeDRlBVnlG/CCpMaJyEbwjU2CcxclGlIzch3pjV+xga5B9GjDZ8I+YUvfCEfvV99XruJY/KhjbJUY0nlQL1Pk0zbTz31VBrIdoazfMbNcoFsW1DHIe1N44iFWcgzgwVgHrxYdkPwaFR4xAL0B++TycieixU0sRap8uXFT5S/UureORHXVLNWHTrUV/6GjkjuHJna97mwu3GfGYsDguL19vlgkJMXHeq9Q+v0KY9axcK8mIovvJ2aREX3CcbEZbgdR/MG/tCmIqSlqUxVV/1cld5YdULy/jwKi6dCDw0dNISwd8qyaIM8WMZO6Lxrt0Wn7hk/ayxUXewRo8FB71fLiiS4/gpi+7NxIwCUwYvkzvo117+ujlU9qudBfvaHXByrcDWeWYJuiIgHJ59ZBzhVQ/16VX3wZVviwFaIDxmGeVh/LJtydDm4VhNm7hJ89xYT0X78Hpifzsi2HLdahhkBOW0j28TVjTtEtzVsUu/nb5tl/182fBGGrUFv+QCWJP4ahHUziHEEJwxJVsaQLJ33xt/pRzcvxjWGbhIfyL2hsrJyt/HneTHoweeYeyPH2BhR7MQEa9RrbEtKSibhHJ0ik96QrCFdyDuIHbFHTX5DYaOBWoB5C31uj8mXculGfgqYchnZ+2PmzoXqJIVH7lpDb764+sUk9rja8e3VpgY2qD0Po3dJ7IfLubcasCPn1zEhmBbAOSdoNHy+GBWfEaPi/DhYrNOjNrz4jCocM0cVFs/DBqI4q7dqv/U6S5sQWQZDIiY/3fAgrkVJKBESkJ+8be3QMe9QdTWPq6ceo3omepOVP0+XU68GtBsa9tyBvV7K2dK0PuoPfIwYcy4mIdM0MSvowblqhOoWGyFN+iwZZdPo3zfc+pZ52nF1veAOoZXwd9ajR74BH+b+hfraJYd0uqyiT6q0rFlQe1nr3g0uOL4aR0Z4Vc2R36h/+3KDjjvYnyCA1hol5meXSZcPN2ykOrmzCiYdH0d342x7bVhgxmccOMP1sR8M/fVlh+no8HP+hPAuYz2/XzW8j79t5PhBmGFN8NCN+zZu3PhrkFYD9cq4MgkyiNI8Yf1zhhfvjVuTtx2PVqQJUuUDknVBPVK2ZcuWsLqCa98ReTVVJ7DRDtDib06LdqFBeAa9X0MJOmygP0gXh7K+QmKFO2SrmkLMmzLReL1x7bXXYuZt6AxGJXUYaTxL+TD4yBqYAXnTwnEIXRUVFVRJvU+mMA4HaR3AxOQzgI89cU7CcXEi9EUYnrc0dKvqsn3wZeF49YNlDYCHWiYb59H4QeaBYCxWw+DgKxxUEoQiPyYUp/LyktXzTxxS+cUPqfkfuQyHY+XplSGst8vNL6Tz4xMgRv3b6J0Pe9fspdLWMfh7wuXDQeeNNeWqofZV9doLb6s3Hj+5Z8Pz2rs6/6I3CaFyWrbugUM8CbT22Bp6R2VCwXYtS5+PgxRcQ64nWokbCC3IT13BmFGCPlLZbkx0S8dF6DDwghyQu78ZHx7fixb+TXXDR8w8AHo/wXTg+gIIEt+NxYoVIsXlnK0tQaycqVYV+6JXUXV4dmCi9c8YqaEHhdpTFroaKCN6U20BDIGPwAd16WzBhq8/Y808fh9819gawfB5ES/tprf+jehHY3ni4bI+us5MYyXT5W2sbVMHd7+Pv22rBP/bf4c1wRPMBQsWHH/44Yd/jeMDXoD6IZPvOYjWfrInwG3/Onr8ET+ElSNhDzIdCRwk68byyCYQNpYaWAZEvAVr3EsZh2kYj/nRkJGx+uQ9K2Z0f7EEcj922d7JMjCF3XDoxJQHdU91dJKij4UvOa1CD/5llpt5sB6sD+8xWvDt2rXLesmiFxl9zLQ2n2qqfRgEl6hiE0DOHXgebpwCiUPEfFh4VHs0qP760nFboPNZ8bU1z9S2oUoNZWHnZk6civOh4eDmJAwNcACA8qM3TzVOekpI/c/929WhHRXq3EUpODUxBL11AN8u9SFPv0pOBNEnBFUsykATwJeVEmOxIQdEYT9XFcR9LFpyXUYcobBrfY36+bcHnkzUwvr442l5Rh2rfI2A61Mk+eFrTq9TIRXErtqd28r6SNW3l9ezVlUdLlNpqdgYhl2qHVi5Qhl+tmAwlO00Ji+nn3ajAG1NARB7m/rvVXVq9yqnegzLyDzPq+oj/9DyYkHExMKN4xS8bUFVjr0KpTdEv8Kr4uCbyp1QrmUwbzZIHShzDFYVdaPBqKmy3p+a2tehfnpXNcHfzwkPu0gcNdDE22sS9JewcB+LZ0S38TegMlks8nCjYao/3IkPfx9j8uFknC/JcKqX1GV4I8DfLUmdxEs3e/vs7+n+GWyrVc3KgtolJRVHCSSoFH+MCqbg7AX06Pn+x2LNczCJPXxMGNZ4VVywAz05P3qJIPTYblwBEEIA2/QD0PlTNsnDeeE23KjQLUYQ+NAhIAT/oXskUqAoESCZG1I3BM+kdFvr3HkefHx8Amajk1QwHqoadwJ67YkqBWer8MMf3TjH1w3lbiCA9avoBrqhBuCHQXxxftXZwMaDkzQkfL9qSsIwrpr50Z+jtt49YHiIEQQ+bAgMexXNhw1wKc+QIUCypSHR0m168jHQr7uhUuA9/b087qFo3LiMGHzkorHdn1bV1Oo/Ul3bsG/b+iYomNEcYIVNUmK86saBN37s7onFZGx6egA9+gAmY3moFsT4KMuQurHhJUYQ+PAiIAT/4X02UrLoEDAEb/Tw2FhTF8osKXH/4Yknxk8pKZkeHxMzGdMJo1wxsVld3b749i5fR7e/+3hbQ8O+rbv27PvRz35ZXltfF1AjMtNwyBcO3MdkeQf074HWTujDu3Vv30zk9ZB8dKWTWILAB4iAqGg+QPAl6yFFgL9lrYP/y1/+Ujz13HMvLcjPvzDQ7Z8WCAY4J94cGxPj4Zwcjg3lUc7ZnGaMiYs9WFV9fM2br73x2jfu/I8OlV2E83TxRahYfCuvKeRRwcYOrGDpVM1Q26g6qmZ4WeqbIS2+CBMEhh4BIfihx1QkfjAIaJ38pu3bFxTm5d2EA98+Fhcff6ypuXFX7fH6Pa3tbUfbmpvbsdLJhdVUianp6YVZ2dnnFOTlTcNnV0twANa6P7344l/+5fs/LlfpOAe9y9uO7562qUA7zlEBwatGQ+xU1RiCq/G6lgAAF+RJREFUF1XNB/OsJdcoERCCjxIoifahRSD8G96yY8/C0YUF38ex8fN8Pv+zrd6OZ+66+ydHfrtqFY4WTlLZ2UkqGef5V7VhZytW0yxfen7cLdetmDBm3JjPji4s/FSHz7/393/666+/feed27HsEb13HPXbjslYz3Eq4Z2Tq0LwH9qfgxRMEBAEhgsCRi2j1m3ZMuF4be1zLa2tDYcrK+9cs2bNaFQSK2rwdaGcnFEqP7/gL88+e/Xmbdu+rUZPGK8KCsapguKxqnjy2Cu//OXz3tm89edt7e111cdr3/jZLx+ci9YgXaWmYitsEb+klYoL34fVCyy5yJKjBTGCwIceAS4pEyMInMkIuO+5557EhQsX3pGWnPx5j8f70Kvr1v7y8htvbFFpabGqrd43pbjY/cgDD4yZM2vWbfl5eZ+cUjJ6bcPRo61H6hoS8cGP1N27D3YeKC/fcsGShSo/K/PyglEjfXlJSS+/sW9fjGrpgGoGn6+zjLPnbvzOZOyk7IKAICAIfLgRWLNx4xIPet/Hj9e8+vbWrYWqqChJZYzJAsHn3v3QQ/nv7tnz1dra2mdw1EIjjpform9sfGnbrl2/+M+7f3m5Gjt5Mg7dP1el5I348b0PjjtaXf1ic0vL8bfXr1+EWrO3zp47zz1JwsUNVLLyDCCIEQQEAUHgfwMBF74P+0scP+HdvmvvZ5EhR6Wp+Fwbzx3K3r5r93UtLW1HmpqbD+Foh98dOXLkURD4+samppZ/vL3uDhzuNVoVjsdVOAbxk55/7Y0r2jyejoNl5Q/M5eFnFqmT4LG6RhN8WOePezGCwIcaAdElfqgfjxRuMAR++9vfZuC00Itb2jwHNuzd9RriU3WCZY3NPEDIm5OVcXlSUsLo6mM1T76zdWvpvffe+4sX/v7SA8fqGp5p93XuwKlx+IpPvUd1YsdqdnbcE3996Y2WtrYDGWmpl31s+fKRkMFePGVx9QxX0ohqBiCIOTMQEII/M56TlLIfBEaPHz/GHRMzNujrWnvzlVeao5hJwh24eNRgAjY5qcT42MSJY8bE79xf0XXVP1/55m+e/9t//efdd7+NQ8M6caiVW8UH8LUkpVbd/+N6b3v7OhzTPHrGubMmQYZefgnbSe7SiwcgYj78CIg+8X1+RuhhJi5btuwSaAzSzJeljI2sdW8Q9zy4UZMGNuTwOGPj7tVbNP4sMuKFw+AfvmccyjDGDjO3lB12Ox0mjZ1ey45M64zfn5vpHWHhMho/u37mluWOLBPLb/J38RTLyuPH98yePr3Po1yzUlKKk5OS4o80Nx+whbLyTM9yxLa0tDyXmZ5+Lr6V+6V8rJx56L67a9p+9IODiXFxa/c9+EqlaoZuPS2Nn2vzq2xodxqVy9vZeQDr6GOmjhuXDxlcIkl5zrq43ly3bv7USZPOIUaO54lo2jjj6jqaACc+rKfBmDYN8Yg0zjR2mJGvscKJn8TJ+Onfj/lYOsvmwFMn7yePE/J2lM35TI2MXvk5y+yoVzgOw1kPkzeOum7HUdt/x5HebIjFvE8ICMG/T8AasUVFRQl4yT6Ll7AANk/epSGfa9uOxxs6+cIbrufxwPqlc8Y1L50dX58sy3DKNrKQTr/45oViGl546U0cnRdlIC0JSr94dMPPxW/H2kbHNy+lkQ8bUXX5w/FMXpQBY86uN+ULxws7rLx1viwb0+DS98wf5TB5x/i7u59FWJ8En56VlYHv4CpPF84/7zFc7UL9uXvz7t2rgn5/YmFR0YVJCUmzcnNysnNzsjuCgeDH7/7r3U+uX7r01aamTp8md7thaenya1kFBblkW6pnnARHd8gVGzvPFwhcDvBYVJZVk7VdbD5gU3fWiWmsdBbefBY4yNKvn4shRMQJG2Ju40I/83x0VsTajqjzZbjzmfM5k/R1Qjwnx+9IJ2PcyAaAzxPlCbcujrx1GpbRrhMsHVf7mz8Mo2H9IYfyKY/l1c8RFgZaOMwTvzXGQ/lqcf0DTiF4AvI+GSH49wlYI3bt2rXtxcXFP8YPPI4/eL5sOG/dBA9qR75oTNCXX6QgE8fYkfny3hj2/vChD1dkuRiHYZChX8pIGSa98Td50d+82LbbRO1lI52W6+x9shy8p0walMnV2tHR75nivo6OVj96qVkp8Vm2cNM6aZXKtVdeeRyHjf3ylltu+QMa21xfMFiQHB9/aU5Ozs15+flz8K2AC9GLrEDPneTGliaYFudOBRHhSJs6bJDSRpOUw+12+WL/6g6F3kILTLLSjbGzHjav0f8EXCmH8o1BnTXO5p62wYFEzedCbB2Y6KjEmH4MJ+sa/PmBmEgT+Wwj82f8yDhOGUa28TNlMeki7438SH9gZPDqxtfNwh+rN3LFFgQEAUEgjMBbW7fOxIdU/NXHjj0e9uxxuN58882pmzZt+rTt5cIHWmLXrVs3Ch8vOYxlk8G3394wz4TB1vtCsNrmCXxspfvll1+eBb9wr9aOJ5YgcMYgID34M+ZRSUH7QqCxsrLcN2pUOT7HuPS+xx7L+/p119UhniblW2+9L270mDH/jrNn5lccObKgobH1zeSEmC58N3dpelpagcfrra2ufq8W8WOwdj5eVVXFPvnMM0lp6enLsOxyf3Nc3CGEWcrxvjIXP0FAEBAEBIH3FQFXeXn5f+A7st379++/wc6JBB/D3vq+ffu+iR5+eZfPVw/VxV4Q926oPWqw4enY/kOHvrV8+XJsZJpGnRk3MsVSRjcOstm7f38p7i3FMhxiBAFBQBAQBD4ABN7etGkyvoVbiWvXzp07x9lF0OS8efPm3P3l5Uura2vvQHgZVC/eqqqqH+3ed/iS119/PdcmdzYILsyXjMeO110NjY2HX1+3bsIHUBXJUhAQBAQBQcCJAHvqmLC7rR29eBD0fz/33HMg7h7DLzpt2LAh5/B77z1+/PjxnS+++OJITK5ylQ0bAa3OefbZZwuOHTv2P2gAfPv3H/zq008/Lec09UAoLkFAEBAEPjgENuzdmwPyfrirq8uDCdQn33777cm9S1OS+MILL8zCsQZL4N9rGdPbGzdOa2ho+AM2OHlwlMGDb+3caVbk9BYhd4KAICAICAIfDAIg79HohT8AVQy+7dH8Dsj6K++8804JSsPeuHvZsmWxds9cq2+gkik+fPjwzfUNDRtwzHBLVXX1L3bt2jUaccUIAsMCAZlEGhaPUSphEICqpggbn/4Z1404biARvfJqbLfZ3tXRcajR423EAv1gWmJ8NnaqjovFZiWsqClEY9CK3asrW5qa/jxr1qyjRpbYgsCZjoAQ/Jn+BKX8JyCAde7Z2Mg0LT8//yJsLvsYNteUYB9QTExsrI87MgN+fzw24HCLZRXcL+E4g79jVc2+efPmtZwgTDwEgTMYASH4M/jhSdH7RwATr+6rrroqJwHr3V3d3UXY9FmYmJicQqIPBv1enIVSjd77e5hUPT5x4sQGSJL17v3DKSFnKAJC8Gfog5NiR48AlkPGoseegF59LM5D4dZ/PyZhu26++WaeMyNGEBAEBAFBQBAQBAQBQUAQEAQEAUFAEBAEBIFhjsApqQCfDoWi2WTkjnIzEsvgLIc+xiDC75QeAzZLJWNyNge7ZTNOSYAkEgQEAUHg/UQA674/iTNafg2i+jXslTir5REcVbsSG3/uxbksXwCJ8oyWkzZYg/41yHkYR+7+FLtE005GwPZduy5HWR45evToD+677z5+CzVseI/THu/EJOlvdu/efRUCnOQdjmccazZuW9Lc2vpwc3PjQ6jjSnyc+zetrZ7HGpubHz5wqPzW13CypIkbrf3KmjXnVFdX/ztOo/wTJmn/gVU4q/Eh719hiebn77nnnqjxwtxAUX19/a94dMLdd9/N776etMFZO5fgaN6HgNW1J51YEgx7BOQ0yWH/iAeuICYdZ2HS8YtYThgAIXux2sSNlSZurDCJxzry5TilcS6I6P9deOGF/LJRVAbHAoxF+tuxzjwfMrpxHv6bSPh8VIkRCQebHwTBL8WkaDbyLYPXUybtBRdc8Bksf/wKiLob5Pob+DvPajfRwnZinCpJTU7+Ig4jC2KitT0WH9/ASew84D6xePSof0pLSfrEW2+99Y2lS5eWhxMN4CgrK/scTqe8HevsJ6CRwQKd7nosvUxJSUpakl5cfDHOlj9/9uzZpSh38wBidBDS5wCjL2EZZw2ew8/g6fxoyWDJdTjSzUb+V+Fiw/LbqBJJpLMGAX0Ox1lTW6noCQiAzONwpYEs30Fv+yb0Sm+E+yZs/rkL2/6zQKbXZWdnn3dCwgE8EP/K1NTUsSAwD2RnQsYKngczQJJeQRg5HMAIYiXIK7ewsPC7WPEyhhFwPjtuC78DQsxrbWv7HRqB7b0S9nHjwhe3IScVRTmAHa43NjQ13VBfV8c6fhOjlVqQ9cfxOb+oer/4xNyinNzc/4K8hZD1InrfN6JRvAaYfRn3P4K8RNT1y2PHjv1ONPVFI0FM0vExjRRgPWBD1UfVjFciHGlomHsdv2ACxT67ERCCP7ufPz8dxx67Qg+3Cj31l1999dWX1q9fvxpqkEdAguux4zMHvfDF0cL0/PPPZ+H7s/8HaoMgVA8/BQlXIf1Fd9zxg+nRykAv2AfCfALE+Up6evpUfInpX5A2ZsKECV/JysqaAXLf2lBX90A0o4o47G9ivuhp19r1e/nnP//5Kzh18s9oxP4bPeiE+MTERYOVjfp+fAHqVnz/dTKI/BmU7wfbtm1bjQZn8xNPPPEmZD0Cwv831Jsjn+u/9a1vTRtMph3ObzRiUHFqBo2DXr+PfE9ZxqnlLKnOBARERXMmPKX3sYyOT7EFnevCn3rqqe7Ro0fryUOob3q+7zdIWdB7vTQjI2MqCO9d6LyfQMNxTkFBwY3J6cnUl+8YJHk4eOHChQ3oyf8QKqKZGBFcA11zPXrb16HB6Ko+evRHCxYsqApHHsCB+mkdPVQYveoHue6ZM2eOYNJAMHjiN+4iZKbm5Y1MS037OBqKTowufjV16tQKEwWbqoK4GjHS+GtmdvZN6amp56MhugDhO02c/mziD4J3oTy6nP3F68/fEDvsUx0B9Cda/IcBAkLww+Ahnk4VMLTXHAGd+Vz0uL+PHj2/7YlvSsdOBknNxARiHXqrr0WTB0guHuS+Ag1CPNI9vmTJkkaQ9O/gd3VeVsbnN27ceD+IuTIaWYyzY8eOjZgfeHDEiBF3lpSU3AYeLMBk4q+OBgIvIjgqQgsE8e1ndG5BoNMqKyt/CqJn/UKoXx5GGpegN97haWt7erAyubu6CuPiYnPQwNRXVFTs7Sv++eef34by7Y3Pzj4f4WP7iuP0I7HjXhM76nlKBM/6UCafm1O2uAUBIiAEf5b/DkAQ+gPQ0JlPgzpmDEf6JA3olpNAfodB7r/A9W40MF188cXzoVJZjLR1e/fuXfsYPqGHCdf34Lc3Ly9vJhqMT0HOA9HIYhyqavBN1Ueg8liOCdd56DnvAoH+4uKlS6OejIwPdvvx1XCVlJRUFBcffz22srK3rNgIQQXFD4D81O/zDT4BHBeHdiIQgv4/FmTcr2oTcnnOPE3Uox7E7VeeljTAH7QRUc9tDCBGgoYpAkLww/TBRlstrppB709BnbIdZPdHkjtJHyTdCvXKLqhaDlx22WVdUchzQVd+HUYCqUgbv2jRorvRU+5mgwFCLYHMGBD8l6DLfgLE7YlCno6CVTPHcAzwVhI8GpwtWO1yONq0jIeuOs4Yi1UYUeyt93juTUtIiE8LBm5PTU2bjB79Jnwg5HHIHHTFS2wodBTk3YCGMAd6d85J/DWyHKhbHuq/EPGI377I8Mh7YK9HIXZPPjI4qns+L0aEDK2LjyqRRDprEBCCP2sedf8VJaFjCeFB9JCfYCwQvQv33eiRe3EblSoEJzhOhirmn0BuAeipj4C7RpC4YILwq8cqkUT0xGdOmjRpGWQO3mNGJGMgT/fYqf+GX1TlMWljseQTZdCTyIfebXg6mNUZMyknp624OH4lVD8f7w4GFyDuSyZ+fzZGMfUg7RcwH3At5hS+887mzUcWz5u31cTnF6HmzJlzB9Q+44HfYUy4vmHC+rOpKiLH08ZmqVOaJGVaygfGoqLpD+iz2F8I/ix++Kw6GYYEiCtwOsflQgXDM9jz0Uisw/VtjAq4OsSlGQz5gIBuwnLE6zFReg2WEK5etWpV1ISGHrjppZ40ibkDId2zBREGli+fxwZLYTXNc+iJPzpi5MhvjMjN/SHuy7Ai5xDD+jMYdQSgdroXo5GZaMjmTR437hF8QWoNRijHkCYVqpv5GKEsRCPQgeWmPwQGVf3JMv5QiblRN84HZGL+4j9hOtHYaoMGUWF08dD8+fP3m/j92Jrg0XguQHl+gThhjDD6qsVu23vuvffeQSeR+5Et3mc4AkLwZ/gDPN3ik/jQW+fVfaqyqJoAWX0G5OYHuT0+ffr0DZGyMGHqYhwQ5Eduuu22qSD4XZFx+rsHiaJ98LOMJ6PX1uK8XR0cQWDRZjDcoIDMPfia073Qy89KTEo6D6ql7z366KO3X3/99W39lYH+2PW7Z9q0ad9Ai/H1+NjYj0NtNA7eHFXEYJ19Cnrue7kjGL39P7NBYJqBDFROIWDWDbVOAiei0R4Gwe5aRQacXJjQfRbpByR41M2HxiCIeYFxqM+XmZ5qMY4MfN3dFZB7P2QIwQ/0IIZxmBD8MH640VQNqoTnQFRHoaMeVGfcnzwQVTfS3wOyceOIAqpfTlCjgPR2oId/S1xcYmZHd3fUOnjmibmA/0HDsR8kOuiyw8gyelpa1iHtLa0ez2Fn2HnnnXcEuv1vxSUkzGkGM1e2tg46WUnSRk97/XXXXXcMrP57qIwmgUyz0QPHxtr2SoxadmAlUjlX0zjzGsBdAdy+ApLHBls9z0rcQnQTS/gfGCCtDkIvfzVGI9xNG15JAzfa7aCrtaXFgyMVhNwHA3EYhwvBD+OHG03VsL78XZDBQZDgSfeOjXz0QNuQ/s9Y4RLChGyfK1zQa+5ET/8F9IBjOmtqTiovEPFONkCnUsY1a1QFOsF/xNJGvymvbYfuuGPr9q9+NW8/et0uzMJGRYQgeD+uMpyJUwWV1ltQjcShB009jXf5PEsFFJFPv7fYUNaMnvqq/iKgvoOW6cCBA7uRvk/1EsoWXLly5SmPzPorl/gLAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAgIAoKAICAICAKCgCAgCAgCgoAgIAicOQj8f14QTZ01EwOjAAAAAElFTkSuQmCC" alt="City Futsal Payroll" style={{width:"100%",maxWidth:160,height:"auto",objectFit:"contain",display:"block"}}/>
      </div>
      <nav style={{padding:"14px 12px",flex:1}}>
        {NAV.map(n=>{
          const active=tab===n.id;
          return <button key={n.id} onClick={()=>setTab(n.id)} style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:10,border:"none",cursor:"pointer",marginBottom:3,background:active?T.gold:"transparent",color:active?T.navy:T.white+"99",fontWeight:active?800:500,fontSize:13,transition:"all .15s",textAlign:"left",fontFamily:"inherit"}}>
            <span style={{display:"flex",alignItems:"center"}}>{ICONS[n.icon]?.(active?T.navy:T.white+"99")}</span>{n.label}
            {n.id==="history"&&weekCount>0&&<span style={{marginLeft:"auto",background:active?T.navy:T.navy3,color:T.gold,borderRadius:20,padding:"1px 8px",fontSize:10,fontWeight:800}}>{weekCount}</span>}
          </button>;
        })}
      </nav>
      <div style={{padding:"14px 12px",borderTop:`1px solid ${T.navy3}`}}>
        <div style={{fontSize:10,fontWeight:800,color:T.muted,letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:10}}>Share with Wesley</div>
        <button onClick={onExport} style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"9px 14px",borderRadius:9,border:`1px solid ${T.gold}40`,cursor:"pointer",marginBottom:6,background:T.gold+"18",color:T.gold,fontWeight:700,fontSize:12,textAlign:"left",fontFamily:"inherit"}}>
          {ICONS.export(T.gold)}Export Data
        </button>
        <input ref={importRef} type="file" accept=".json" style={{display:"none"}} onChange={e=>{if(e.target.files[0]){onImport(e.target.files[0]);e.target.value="";}}}/>
        <button onClick={()=>importRef.current.click()} style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"9px 14px",borderRadius:9,border:`1px solid ${T.navy3}`,cursor:"pointer",background:"transparent",color:T.white+"70",fontWeight:600,fontSize:12,textAlign:"left",fontFamily:"inherit"}}>
          {ICONS.import(T.white+"70")}Import Data
        </button>
        <div style={{fontSize:10,color:T.muted,marginTop:8,lineHeight:1.6}}>Export → send .json to Wesley → he imports to see live data.</div>
      </div>
      <div style={{padding:"10px 16px",borderTop:`1px solid ${T.navy3}`,textAlign:"center"}}>
        <div style={{color:T.white+"25",fontSize:9,letterSpacing:"0.06em",textTransform:"uppercase"}}>© 2026 City Futsal Holdings</div>
      </div>
    </div>
  );
}

/* ─── ERROR BOUNDARY ──────────────────────────────────────────────────────── */
// Catches render errors so a single bad memo or chart can't blank the page.
export class ErrorBoundary extends Component {
  constructor(props){ super(props); this.state={err:null,info:null}; }
  static getDerivedStateFromError(err){ return {err}; }
  componentDidCatch(err,info){
    this.setState({info});
    try{ console.error("App crashed:",err,info?.componentStack); }catch{}
  }
  render(){
    if(this.state.err){
      const msg=this.state.err?.message||String(this.state.err);
      const stack=this.state.err?.stack||"";
      return(
        <div style={{minHeight:"100vh",background:T.off,padding:32,fontFamily:"'DM Sans',sans-serif"}}>
          <div style={{maxWidth:760,margin:"0 auto"}}>
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:900,fontSize:24,color:T.navy,marginBottom:8}}>Something broke</div>
            <div style={{color:T.muted,fontSize:13,marginBottom:18}}>The app caught a crash and stopped it from blanking the page. Your data is still safe in Supabase — refresh to recover.</div>
            <div style={{display:"flex",gap:10,marginBottom:20}}>
              <button onClick={()=>this.setState({err:null,info:null})} style={{padding:"10px 22px",borderRadius:9,border:"none",background:T.gold,color:T.navy,cursor:"pointer",fontWeight:800,fontSize:13,fontFamily:"inherit"}}>Try again</button>
              <button onClick={()=>location.reload()} style={{padding:"10px 22px",borderRadius:9,border:`1px solid ${T.border}`,background:T.white,color:T.navy,cursor:"pointer",fontWeight:700,fontSize:13,fontFamily:"inherit"}}>Reload page</button>
            </div>
            <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:10,padding:16,fontSize:12,color:T.red,fontFamily:"'DM Mono',monospace",overflow:"auto",marginBottom:12}}>
              <strong>{msg}</strong>
            </div>
            {stack&&<pre style={{background:T.navy,color:T.gold+"CC",padding:14,borderRadius:10,fontSize:10,overflow:"auto",fontFamily:"'DM Mono',monospace",maxHeight:400}}>{stack}</pre>}
            {this.state.info?.componentStack&&<pre style={{background:T.white,border:`1px solid ${T.border}`,color:T.muted,padding:14,borderRadius:10,fontSize:10,overflow:"auto",fontFamily:"'DM Mono',monospace",maxHeight:300,marginTop:8}}>{this.state.info.componentStack}</pre>}
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

/* ─── PASSWORD GATE ───────────────────────────────────────────────────────── */
// Bump AUTH_KEY's "vN" suffix to force everyone to re-enter after a password change.
const AUTH_KEY="cf:auth:v1";
function PasswordGate({onAuth}){
  const required=import.meta.env.VITE_APP_PASSWORD;
  const [input,setInput]=useState("");
  const [err,setErr]=useState(false);
  const submit=e=>{
    e.preventDefault();
    if(input===required){
      try{localStorage.setItem(AUTH_KEY,"ok");}catch{}
      onAuth();
    }else{
      setErr(true);
      setInput("");
    }
  };
  return(
    <div style={{position:"fixed",inset:0,background:`linear-gradient(135deg, ${T.navy} 0%, ${T.navy2} 100%)`,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",padding:20}}>
      <form onSubmit={submit} style={{background:T.white,padding:36,borderRadius:18,boxShadow:T.shadow2,minWidth:340,maxWidth:380}}>
        <div style={{textAlign:"center",fontFamily:"'Syne',sans-serif",fontWeight:900,fontSize:22,color:T.navy,letterSpacing:"-0.01em"}}>CF Payroll</div>
        <div style={{textAlign:"center",color:T.muted,fontSize:12,marginTop:6,marginBottom:24}}>Enter the team password</div>
        <input type="password" autoFocus value={input} onChange={e=>{setInput(e.target.value);setErr(false);}} placeholder="Password" style={{width:"100%",padding:"12px 14px",borderRadius:9,border:`1px solid ${err?T.red:T.border}`,fontSize:14,fontFamily:"inherit",outline:"none",marginBottom:err?6:18,color:T.navy}}/>
        {err&&<div style={{color:T.red,fontSize:12,marginBottom:14,textAlign:"center",fontWeight:600}}>Incorrect password</div>}
        <button type="submit" style={{width:"100%",padding:"12px",borderRadius:9,border:"none",background:T.gold,color:T.navy,fontWeight:800,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>Enter</button>
      </form>
    </div>
  );
}

/* ═══ MAIN APP ════════════════════════════════════════════════════════════════ */
export default function App(){
  const [tab,setTab]=useState("upload");
  const [weeks,setWeeks]=useState({});
  const [attachments,setAttachments]=useState({});
  const [squareByWeek,setSquareByWeek]=useState({});
  const [loading,setLoading]=useState(true);
  const [uploading,setUploading]=useState(false);
  const [preview,setPreview]=useState(null);
  const [selectedWeek,setSelectedWeek]=useState(null);
  const [compareWeek,setCompareWeek]=useState(null);
  const [toast,setToast]=useState(null);
  const [deleteConfirm,setDeleteConfirm]=useState(null);
  const [compareMode,setCompareMode]=useState("week");
  const [customRange,setCustomRange]=useState({from:"",to:""});
  const [schedLocFilter,setSchedLocFilter]=useState("All");
  const [reportLocFilter,setReportLocFilter]=useState("all");
  const [formDetail,setFormDetail]=useState(null);
  const [formFilter,setFormFilter]=useState({type:"all",venue:"all",employee:"",weekScope:"all"});
  const [authed,setAuthed]=useState(()=>{
    const req=import.meta.env.VITE_APP_PASSWORD;
    if(!req) return true; // no password configured = open
    try{ return localStorage.getItem(AUTH_KEY)==="ok"; }catch{ return false; }
  });

  useEffect(()=>{
    storeLoad().then(({weeks:w,attachments:a})=>{
      setWeeks(w);setAttachments(a);
      const keys=Object.keys(w).sort();
      if(keys.length){setSelectedWeek(keys[keys.length-1]);setCompareWeek(keys[keys.length-2]||null);}
      // Load square data from attachments — distribute rows to their actual
      // week (not the attachment's stored weekKey) so legacy uploads that
      // landed under the wrong week self-heal.
      const sqRaw={};
      Object.entries(a).forEach(([k,v])=>{
        if(v.type!=="square"||!v.parsedData) return;
        v.parsedData.forEach(row=>{
          const wk=row.date&&/^\d{4}-\d{2}-\d{2}$/.test(row.date)?weekKey(row.date):v.weekKey;
          sqRaw[wk]=(sqRaw[wk]||[]).concat(row);
        });
      });
      const sq={};
      Object.entries(sqRaw).forEach(([wk,rows])=>{sq[wk]=aggregateSquareRows(rows);});
      setSquareByWeek(sq);
      setLoading(false);
    });
  },[]);

  const toast2=(msg,type="ok")=>setToast({msg,type});

  const handleFiles=async files=>{
    setUploading(true);
    try{
      const all=[];
      for(const f of files){const rows=await parseExcel(f);all.push(...rows);}
      if(!all.length){toast2("No rows found in files","err");setUploading(false);return;}
      const starts=all.map(r=>r.startDate).filter(Boolean).sort();
      const ends=all.map(r=>r.endDate).filter(Boolean).sort().reverse();
      const sd=starts[0]||new Date().toISOString().slice(0,10);
      const ed=ends[0]||sd;
      const wk=weekKey(sd);
      setPreview({weekKey:wk,period:periodLabel(sd,ed),startDate:sd,endDate:ed,
        employees:all,totalPay:all.reduce((s,e)=>s+e.pay,0),uploadedAt:new Date().toISOString()});
    }catch(err){toast2("Parse error: "+err,"err");}
    setUploading(false);
  };

  const confirmSave=async()=>{
    if(!preview) return;
    await storeWeek(preview.weekKey,preview);
    const updated={...weeks,[preview.weekKey]:preview};
    setWeeks(updated);
    setSelectedWeek(preview.weekKey);
    const keys=Object.keys(updated).sort();
    setCompareWeek(keys[keys.indexOf(preview.weekKey)-1]||null);
    setPreview(null);
    toast2(`${preview.period} saved`);
    setTab("report");
  };

  const handleSquareUpload=async(files,fallbackWeekK)=>{
    for(const f of files){
      try{
        const data=await parseSquare(f);
        if(!data.length){toast2("No sales data found in "+f.name,"warn");continue;}

        // Detect location from filename (match any LOCS value case-insensitive)
        const nameLower=(f.name||"").toLowerCase();
        let detectedLoc="Unknown";
        for(const L of LOCS){if(nameLower.includes(L.toLowerCase())){detectedLoc=L;break;}}
        const dailyRows=data.filter(d=>!d.isWeeklyTotal).map(d=>({...d,location:detectedLoc}));
        const summaryRows=data.filter(d=>d.isWeeklyTotal);

        // Derive the attachment's primary week from the file's first date so
        // the file appears under that week's Documents panel (for delete UI).
        // The aggregated daily data below is distributed by each row's own
        // week, so multi-week files still route correctly.
        const firstDate=dailyRows[0]?.date;
        const attWeekK=firstDate&&/^\d{4}-\d{2}-\d{2}$/.test(firstDate)?weekKey(firstDate):fallbackWeekK;

        const attKey=`${attWeekK}:square:${Date.now()}:${f.name}`;
        const val={name:f.name,type:"square",weekKey:attWeekK,size:f.size,
          uploadedAt:new Date().toISOString(),parsedData:dailyRows,summaryData:summaryRows};
        await storeAtt(attKey,val);
        setAttachments(a=>({...a,[attKey]:val}));

        if(dailyRows.length>0){
          // Group rows by their actual week (a file may span >1 week).
          const byWeek={};
          dailyRows.forEach(row=>{
            const wk=row.date&&/^\d{4}-\d{2}-\d{2}$/.test(row.date)?weekKey(row.date):attWeekK;
            (byWeek[wk]=byWeek[wk]||[]).push(row);
          });
          setSquareByWeek(sq=>{
            const result={...sq};
            Object.entries(byWeek).forEach(([wk,rows])=>{
              const curRaw=(result[wk]||[]).reduce((acc,d)=>{
                if(d.byLocation){
                  Object.entries(d.byLocation).forEach(([loc,g])=>acc.push({date:d.date,gross:g,location:loc}));
                } else { acc.push({date:d.date,gross:d.gross,location:d.location||"Unknown"}); }
                return acc;
              },[]);
              result[wk]=aggregateSquareRows([...curRaw,...rows]);
            });
            return result;
          });
          const total=dailyRows.reduce((s,d)=>s+d.gross,0);
          const wkCount=Object.keys(byWeek).length;
          toast2(`✓ ${f.name}: ${dailyRows.length} days · ${fmtUSD(total)}${wkCount>1?` across ${wkCount} weeks`:""}`);
        } else if(summaryRows.length>0){
          const total=summaryRows.reduce((s,d)=>s+d.gross,0);
          toast2(`✓ ${f.name}: Weekly total ${fmtUSD(total)} recorded`);
        }
      }catch(e){toast2("Could not parse "+f.name+". Make sure it's a Square Sales Summary CSV. Error: "+String(e),"err");}
    }
  };

  const handleFormUpload=async(files,weekK)=>{
    for(const f of files){
      const isExcel=/\.(xlsx|xls)$/i.test(f.name||"");
      let parsedData=null;
      let parseErr=null;
      if(isExcel){
        try{ parsedData=await parseConnecteamForm(f); }
        catch(e){ parseErr=String(e); }
      }
      const attKey=`${weekK}:form:${Date.now()}:${f.name}`;
      const val={name:f.name,type:"form",weekKey:weekK,size:f.size,
        uploadedAt:new Date().toISOString(),
        ...(parsedData?{parsedData}:{})};
      await storeAtt(attKey,val);
      setAttachments(a=>({...a,[attKey]:val}));
      if(parsedData){
        const t=FORM_TYPES[parsedData.formType]||FORM_TYPES.other;
        toast2(`✓ ${f.name}: ${parsedData.submissions.length} ${t.label} submission${parsedData.submissions.length===1?"":"s"}`);
      }else if(isExcel&&parseErr){
        toast2(`${f.name} attached (couldn't parse: ${parseErr})`,"warn");
      }else{
        toast2(`${f.name} attached`);
      }
    }
  };

  const handleDelWeek=async k=>{
    await delWeek(k);const u={...weeks};delete u[k];setWeeks(u);
    const keys=Object.keys(u).sort();setSelectedWeek(keys[keys.length-1]||null);
    setDeleteConfirm(null);toast2("Week deleted","err");
  };

  const handleDelAtt=async k=>{
    const att=attachments[k];
    const weekK=att?.weekKey;
    const wasSquare=att?.type==="square";
    await delAtt(k);
    const updated={...attachments};delete updated[k];
    setAttachments(updated);
    if(wasSquare&&weekK){
      const remaining=Object.values(updated)
        .filter(v=>v.type==="square"&&v.weekKey===weekK)
        .flatMap(v=>v.parsedData||[]);
      setSquareByWeek(sq=>{
        const u={...sq};
        if(remaining.length) u[weekK]=aggregateSquareRows(remaining);
        else delete u[weekK];
        return u;
      });
    }
    toast2(`${att?.name||"Upload"} deleted`,"err");
  };

  const handleExport=()=>{
    const payload={weeks,attachments,squareByWeek,exportedAt:new Date().toISOString(),exportedBy:"CFH Payroll Dashboard"};
    const blob=new Blob([JSON.stringify(payload,null,2)],{type:"application/json"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");a.href=url;
    a.download=`CFH_Payroll_Export_${new Date().toISOString().slice(0,10)}.json`;
    a.click();URL.revokeObjectURL(url);
    toast2("Exported — send the .json to Wesley");
  };

  const handleImport=file=>{
    const reader=new FileReader();
    reader.onload=async e=>{
      try{
        const p=JSON.parse(e.target.result);
        if(!p.weeks) throw new Error("Invalid export file");
        for(const [k,v] of Object.entries(p.weeks)) await storeWeek(k,v);
        if(p.attachments) for(const [k,v] of Object.entries(p.attachments)) await storeAtt(k,v);
        const merged={...weeks,...p.weeks};
        setWeeks(merged);
        setAttachments(a=>({...a,...(p.attachments||{})}));
        if(p.squareByWeek) setSquareByWeek(sq=>({...sq,...p.squareByWeek}));
        const keys=Object.keys(merged).sort();
        if(keys.length){setSelectedWeek(keys[keys.length-1]);setCompareWeek(keys[keys.length-2]||null);}
        toast2(`Imported ${Object.keys(p.weeks).length} week(s)`);
        setTab("report");
      }catch(err){toast2("Import failed: "+err.message,"err");}
    };
    reader.readAsText(file);
  };

  const sortedKeys=Object.keys(weeks).sort().reverse();
  const currentWeek=selectedWeek?weeks[selectedWeek]:null;
  const cmpWeekData=compareWeek?weeks[compareWeek]:null;

  // Report-tab location filter applied to all employee-derived stats below.
  const matchLoc=(e,loc)=>loc==="all"||e.location===loc||normalizeVenue(e.location)===loc;
  const filteredEmps=currentWeek?currentWeek.employees.filter(e=>matchLoc(e,reportLocFilter)):[];
  const filteredCmpEmps=cmpWeekData?cmpWeekData.employees.filter(e=>matchLoc(e,reportLocFilter)):[];
  const filteredWeek=currentWeek?(reportLocFilter==="all"?currentWeek:{...currentWeek,employees:filteredEmps}):null;

  const totalPay=filteredEmps.reduce((s,e)=>s+e.pay,0);
  const totalHrs=filteredEmps.reduce((s,e)=>s+e.hours,0);
  const totalSched=filteredEmps.reduce((s,e)=>s+e.scheduledHours,0);
  const totalStaff=new Set(filteredEmps.map(e=>`${e.firstName} ${e.lastName}`)).size;
  const priorPay=filteredCmpEmps.length?filteredCmpEmps.reduce((s,e)=>s+e.pay,0):null;
  const payDelta=priorPay?((totalPay-priorPay)/Math.max(priorPay,1))*100:undefined;
  const groups=buildGroups(weeks,compareMode,customRange);
  const currentSquare=selectedWeek?squareByWeek[selectedWeek]||[]:[];
  const weekAttachments=Object.entries(attachments).filter(([k])=>k.startsWith((selectedWeek||"")+":")).map(([k,v])=>({k,v}));

  // ── Form submissions (deduped across attachments)
  const allFormSubmissions=useMemo(()=>{
    const seen=new Set();
    const list=[];
    Object.entries(attachments).forEach(([attKey,v])=>{
      if(v.type!=="form"||!v.parsedData?.submissions) return;
      v.parsedData.submissions.forEach(s=>{
        const key=`${s.formType}|${s.id}`;
        if(seen.has(key)) return;
        seen.add(key);
        list.push({...s, weekKey:v.weekKey, fileName:v.name, attKey});
      });
    });
    return list.sort((a,b)=>(b.submissionDate||"").localeCompare(a.submissionDate||"")||(b.submissionTime||"").localeCompare(a.submissionTime||""));
  },[attachments]);

  const filteredSubs=useMemo(()=>{
    return allFormSubmissions.filter(s=>{
      if(formFilter.type!=="all"&&s.formType!==formFilter.type) return false;
      if(formFilter.venue!=="all"&&s.venue!==formFilter.venue) return false;
      if(formFilter.employee&&!s.submittedBy.toLowerCase().includes(formFilter.employee.toLowerCase())) return false;
      if(formFilter.weekScope==="current"&&selectedWeek&&s.weekKey!==selectedWeek) return false;
      return true;
    });
  },[allFormSubmissions,formFilter,selectedWeek]);

  const subsByEmployee=useMemo(()=>{
    const m={};
    filteredSubs.forEach(s=>{(m[s.submittedBy]=m[s.submittedBy]||[]).push(s);});
    return m;
  },[filteredSubs]);

  // ── Sales × Staff attribution: per-day × per-location rows for current week.
  // Normalizes legacy stored data: dailyBreakdown.date may be "Tue May 26" or
  // "5/26/2026" from older parseExcel; employee.location may be full venue name.
  // Honors reportLocFilter so the table scopes to a single location when chosen.
  const salesAttribution=useMemo(()=>{
    if(!currentWeek||!currentSquare?.length) return [];
    const rows=[];
    currentSquare.forEach(day=>{
      const dayIso=toDateStr(day.date);
      const locsForDay=day.byLocation?Object.keys(day.byLocation):[];
      const list=locsForDay.length?locsForDay:(day.gross>0?["Unknown"]:[]);
      list.forEach(loc=>{
        const sales=day.byLocation?(day.byLocation[loc]||0):day.gross;
        if(sales<=0) return;
        const locNorm=normalizeVenue(loc);
        if(reportLocFilter!=="all"&&locNorm!==reportLocFilter&&loc!==reportLocFilter&&loc!=="Unknown") return;
        const staff=currentWeek.employees
          .filter(e=>(loc==="Unknown"||normalizeVenue(e.location)===locNorm||e.location===loc)&&matchLoc(e,reportLocFilter))
          .map(e=>{
            const d=(e.dailyBreakdown||[]).find(x=>toDateStr(x.date)===dayIso);
            if(!d) return null;
            const hours=d.rate>0?d.pay/d.rate:0;
            return (hours>0||d.fromSummary)?{name:`${e.firstName} ${e.lastName.charAt(0)}.`,fullName:`${e.firstName} ${e.lastName}`,hours,pay:d.pay,rate:d.rate}:null;
          })
          .filter(Boolean)
          .sort((a,b)=>b.hours-a.hours);
        const staffHours=staff.reduce((s,e)=>s+e.hours,0);
        rows.push({date:dayIso,location:loc,sales,staff,staffHours});
      });
    });
    return rows.sort((a,b)=>a.date.localeCompare(b.date)||a.location.localeCompare(b.location));
  },[currentWeek,currentSquare,reportLocFilter]);

  // ── Per-employee sales allocation: each day's location sales split equally
  // among the employees who worked that location that day.
  // Honors reportLocFilter so allocation scopes to one venue when chosen.
  const salesPerEmployee=useMemo(()=>{
    if(!currentWeek||!currentSquare?.length) return [];
    const map={};
    currentWeek.employees.filter(e=>matchLoc(e,reportLocFilter)).forEach(e=>{
      map[`${e.firstName}|${e.lastName}`]={
        firstName:e.firstName, lastName:e.lastName,
        location:e.location, role:e.role,
        pay:e.pay, hours:e.hours,
        salesShare:0, daysWithSales:0,
        breakdown:[],
      };
    });
    currentSquare.forEach(day=>{
      const dayIso=toDateStr(day.date);
      const dailyLocs=day.byLocation?Object.entries(day.byLocation):[];
      const list=dailyLocs.length?dailyLocs:(day.gross>0?[["Unknown",day.gross]]:[]);
      list.forEach(([loc,sales])=>{
        if(sales<=0) return;
        const locNorm=normalizeVenue(loc);
        if(reportLocFilter!=="all"&&locNorm!==reportLocFilter&&loc!==reportLocFilter&&loc!=="Unknown") return;
        const workers=currentWeek.employees.filter(e=>
          (loc==="Unknown"||normalizeVenue(e.location)===locNorm||e.location===loc)&&
          matchLoc(e,reportLocFilter)&&
          (e.dailyBreakdown||[]).some(d=>toDateStr(d.date)===dayIso&&(d.pay>0||d.fromSummary))
        );
        if(!workers.length) return;
        const share=sales/workers.length;
        workers.forEach(e=>{
          const k=`${e.firstName}|${e.lastName}`;
          if(!map[k]) return;
          map[k].salesShare+=share;
          map[k].daysWithSales+=1;
          map[k].breakdown.push({date:dayIso,location:loc,daySales:sales,workers:workers.length,share});
        });
      });
    });
    return Object.values(map)
      .filter(e=>e.salesShare>0)
      .sort((a,b)=>b.salesShare-a.salesShare);
  },[currentWeek,currentSquare,reportLocFilter]);

  // ── Per-venue Goal Performance: actual share vs expected (shift-proportional)
  // share of weekly sales. Per the team rule, an employee who worked 3 of 30
  // shifts should produce at least 10% of the weekly sales total at that venue.
  const goalPerformance=useMemo(()=>{
    if(!currentWeek||!currentSquare?.length) return [];
    const locsToReport=reportLocFilter==="all"?LOCS:[reportLocFilter];
    return locsToReport.map(loc=>{
      // Total Square sales at this location for the week
      const weekTotal=currentSquare.reduce((s,day)=>{
        const v=day.byLocation?(day.byLocation[loc]||0):(loc==="Unknown"?day.gross:0);
        return s+v;
      },0);
      const locEmps=currentWeek.employees.filter(e=>matchLoc(e,loc));
      // Pre-index workers per day for this venue
      const workersByDay={};
      currentSquare.forEach(day=>{
        const dayIso=toDateStr(day.date);
        workersByDay[dayIso]=locEmps.filter(em=>(em.dailyBreakdown||[]).some(d=>toDateStr(d.date)===dayIso&&(d.pay>0||d.fromSummary)));
      });
      const empData=locEmps.map(e=>{
        const shifts=(e.dailyBreakdown||[]).filter(d=>d.pay>0||d.fromSummary).length;
        let actual=0;
        currentSquare.forEach(day=>{
          const dayIso=toDateStr(day.date);
          const sales=day.byLocation?(day.byLocation[loc]||0):(loc==="Unknown"?day.gross:0);
          if(sales<=0) return;
          const dayWorkers=workersByDay[dayIso]?.length||0;
          const worked=(e.dailyBreakdown||[]).some(d=>toDateStr(d.date)===dayIso&&(d.pay>0||d.fromSummary));
          if(worked&&dayWorkers>0) actual+=sales/dayWorkers;
        });
        return {firstName:e.firstName,lastName:e.lastName,role:e.role,shifts,actual};
      }).filter(e=>e.shifts>0);
      const totalShifts=empData.reduce((s,e)=>s+e.shifts,0);
      const employees=empData.map(e=>{
        const expected=totalShifts>0?(e.shifts/totalShifts)*weekTotal:0;
        return {...e,expected,status:e.actual>=expected?"above":"below",deltaPct:expected>0?((e.actual-expected)/expected)*100:0};
      }).sort((a,b)=>b.actual-a.actual);
      const teamStatus=weekTotal>=GOAL_MAX?"exceeded":weekTotal>=GOAL_MIN?"met":"below";
      return {location:loc,weekTotal,totalShifts,employees,teamStatus};
    }).filter(r=>r.weekTotal>0||r.employees.length>0);
  },[currentWeek,currentSquare,reportLocFilter]);

  if(!authed) return <PasswordGate onAuth={()=>setAuthed(true)}/>;
  if(loading) return <div style={{height:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.navy}}><div style={{textAlign:"center"}}><div style={{width:48,height:48,background:T.gold,borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,color:T.navy,fontSize:20,margin:"0 auto 14px"}}>CF</div><div style={{color:T.white,fontWeight:700}}>Loading…</div></div></div>;

  return(
    <div style={{display:"flex",minHeight:"100vh",background:T.off,fontFamily:"'DM Sans','Segoe UI',sans-serif",color:T.text}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&family=DM+Mono:wght@400;500;700&family=Syne:wght@700;800;900&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        @keyframes toastIn{from{transform:translateX(30px);opacity:0}to{transform:none;opacity:1}}
        @keyframes fadeUp{from{transform:translateY(10px);opacity:0}to{transform:none;opacity:1}}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px}
        input:focus,select:focus{outline:2px solid ${T.gold}80;outline-offset:1px}
        button:not([disabled]):hover{filter:brightness(1.08)}
      `}</style>

      {toast&&<Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)}/>}

      {deleteConfirm&&<Modal title="Delete this week?" onClose={()=>setDeleteConfirm(null)}>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:40,marginBottom:12}}>🗑️</div>
          <div style={{color:T.muted,fontSize:13,marginBottom:24}}>{weeks[deleteConfirm]?.period}</div>
          <div style={{display:"flex",gap:12,justifyContent:"center"}}>
            <button onClick={()=>setDeleteConfirm(null)} style={{padding:"10px 24px",borderRadius:9,border:`1px solid ${T.border}`,background:T.white,cursor:"pointer",fontWeight:700,fontSize:13,fontFamily:"inherit",color:T.navy}}>Cancel</button>
            <button onClick={()=>handleDelWeek(deleteConfirm)} style={{padding:"10px 24px",borderRadius:9,border:"none",background:T.red,color:T.white,cursor:"pointer",fontWeight:700,fontSize:13,fontFamily:"inherit"}}>Delete</button>
          </div>
        </div>
      </Modal>}

      <Sidebar tab={tab} setTab={setTab} weekCount={sortedKeys.length} onExport={handleExport} onImport={handleImport}/>

      <div style={{flex:1,padding:"28px 32px",maxWidth:1140,animation:"fadeUp .3s ease"}}>

        {/* ══ UPLOAD ══ */}
        {tab==="upload"&&(
          <div style={{display:"flex",flexDirection:"column",gap:24}}>
            <div>
              <div style={{fontFamily:"'Syne',sans-serif",fontWeight:900,fontSize:26,color:T.navy,letterSpacing:"-0.02em"}}>Upload Payroll</div>
              <div style={{color:T.muted,fontSize:13,marginTop:4}}>Drop Connecteam exports — all locations at once.</div>
            </div>
            {!preview?(
              <div style={{maxWidth:520}}>
                <UploadZone onFiles={handleFiles} label="Drop Connecteam Excel exports" sub="Dallas · The Colony — all at once" icon="📂"/>
                {uploading&&<div style={{textAlign:"center",color:T.muted,padding:20,fontSize:13}}>Parsing files…</div>}
              </div>
            ):(
              <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:16,overflow:"hidden",boxShadow:T.shadow}}>
                <div style={{background:T.navy,padding:"16px 22px",display:"flex",alignItems:"center",gap:16}}>
                  <div style={{flex:1}}>
                    <div style={{color:T.gold,fontWeight:900,fontSize:15,fontFamily:"'Syne',sans-serif"}}>{preview.period}</div>
                    <div style={{color:T.white+"70",fontSize:12,marginTop:2}}>{preview.employees.length} employees · {fmtUSD(preview.totalPay)}</div>
                  </div>
                  <button onClick={()=>setPreview(null)} style={{padding:"8px 18px",borderRadius:8,border:`1px solid ${T.navy3}`,background:"transparent",color:T.white+"80",cursor:"pointer",fontWeight:600,fontSize:13,fontFamily:"inherit"}}>Discard</button>
                  <button onClick={confirmSave} style={{padding:"8px 20px",borderRadius:8,border:"none",background:T.gold,color:T.navy,cursor:"pointer",fontWeight:800,fontSize:13,fontFamily:"inherit"}}>✓ Save Week</button>
                </div>
                <div style={{padding:20}}>
                  {LOCS.filter(l=>preview.employees.some(e=>e.location===l)).map(l=>{
                    const emps=preview.employees.filter(e=>e.location===l);
                    const tot=emps.reduce((s,e)=>s+e.pay,0);
                    const clr=LOC_CLR[l]||T.navy;
                    return <div key={l} style={{marginBottom:12,border:`1px solid ${T.border}`,borderRadius:10,overflow:"hidden"}}>
                      <div style={{background:clr,padding:"10px 16px",display:"flex",justifyContent:"space-between"}}>
                        <span style={{color:T.white,fontWeight:800,fontSize:13}}>{l}</span>
                        <span style={{color:T.white+"CC",fontSize:12}}>{emps.length} staff · <strong style={{color:T.white}}>{fmtUSD(tot)}</strong></span>
                      </div>
                      {emps.map((e,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:12,padding:"8px 16px",background:i%2===0?T.white:T.off,borderTop:`1px solid ${T.border}`,fontSize:13}}>
                        <span style={{flex:1,fontWeight:600,color:T.navy}}>{e.firstName} {e.lastName}</span>
                        <Pill color={ROLE_CLR(e.role)}>{e.role}</Pill>
                        <span style={{color:T.muted,fontFamily:"'DM Mono',monospace"}}>{fmtHHMM(e.hours)}</span>
                        <span style={{fontWeight:800,color:T.navy,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(e.pay)}</span>
                      </div>)}
                    </div>;
                  })}
                </div>
              </div>
            )}
            {!preview&&sortedKeys.length>0&&(
              <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow,maxWidth:520}}>
                <div style={{background:T.off,padding:"11px 18px",fontWeight:800,color:T.navy,fontSize:11,borderBottom:`1px solid ${T.border}`,letterSpacing:"0.06em",textTransform:"uppercase"}}>Recent Uploads</div>
                {sortedKeys.slice(0,6).map(k=>{const w=weeks[k];return(
                  <div key={k} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 18px",borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}
                    onClick={()=>{setSelectedWeek(k);setTab("report");}}>
                    <div style={{width:8,height:8,borderRadius:"50%",background:T.gold,flexShrink:0}}/>
                    <div style={{flex:1}}><div style={{fontWeight:700,color:T.navy,fontSize:13}}>{w.period}</div><div style={{color:T.muted,fontSize:11,marginTop:1}}>{w.employees.length} employees</div></div>
                    <div style={{fontFamily:"'DM Mono',monospace",fontWeight:800,color:T.navy,fontSize:13}}>{fmtUSD(w.totalPay)}</div>
                  </div>
                );})}
              </div>
            )}
          </div>
        )}

        {/* ══ REPORT ══ */}
        {tab==="report"&&(
          <div style={{display:"flex",flexDirection:"column",gap:20}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:12}}>
              <div>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:900,fontSize:26,color:T.navy,letterSpacing:"-0.02em"}}>Weekly Report</div>
                <div style={{color:T.muted,fontSize:13,marginTop:3}}>Full payroll with schedule variance and daily breakdown.</div>
              </div>
              {sortedKeys.length>0&&(
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  <select value={compareWeek||""} onChange={e=>setCompareWeek(e.target.value||null)} style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${T.border}`,background:T.white,fontSize:12,color:T.muted,cursor:"pointer",fontFamily:"inherit"}}>
                    <option value="">Compare with…</option>
                    {sortedKeys.filter(k=>k!==selectedWeek).map(k=><option key={k} value={k}>{weeks[k].period}</option>)}
                  </select>
                  <select value={selectedWeek||""} onChange={e=>setSelectedWeek(e.target.value)} style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${T.border}`,background:T.white,fontSize:13,fontWeight:700,color:T.navy,cursor:"pointer",fontFamily:"inherit"}}>
                    {sortedKeys.map(k=><option key={k} value={k}>{weeks[k].period}</option>)}
                  </select>
                </div>
              )}
            </div>

            {!currentWeek?(
              <div style={{textAlign:"center",padding:80,color:T.muted}}>
                <div style={{fontSize:48,marginBottom:16}}>📭</div>
                <div style={{fontWeight:800,fontSize:18,color:T.navy}}>No data yet</div>
                <button onClick={()=>setTab("upload")} style={{marginTop:20,padding:"10px 24px",borderRadius:9,border:"none",background:T.gold,color:T.navy,cursor:"pointer",fontWeight:800,fontSize:14,fontFamily:"inherit"}}>Upload →</button>
              </div>
            ):(
              <>
                {/* Location filter — scopes the entire weekly report */}
                <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                  <div style={{fontSize:11,color:T.muted,fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",marginRight:4}}>View</div>
                  {[{id:"all",label:"All Locations (Combined)"},...LOCS.map(l=>({id:l,label:l}))].map(opt=>{
                    const active=reportLocFilter===opt.id;
                    const accent=opt.id==="all"?T.gold:(LOC_CLR[opt.id]||T.navy3);
                    return(
                      <button key={opt.id} onClick={()=>setReportLocFilter(opt.id)} style={{padding:"7px 16px",borderRadius:8,border:`1px solid ${active?accent:T.border}`,background:active?accent+"14":T.white,color:active?accent:T.navy,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
                        {opt.label}
                      </button>
                    );
                  })}
                </div>

                {/* KPIs */}
                <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
                  <KPI label={reportLocFilter==="all"?"Total Payroll":`${reportLocFilter} Payroll`} value={fmtUSD(totalPay)} sub={currentWeek.period} accent={T.gold} delta={payDelta} mono/>
                  <KPI label="Hours Worked"     value={fmtHHMM(totalHrs)}     sub={totalSched>0?`Sched: ${fmtHHMM(totalSched)}`:""}  accent={T.navy3} mono/>
                  {totalSched>0&&<KPI label="Schedule Var."  value={fmtHHMMsigned(totalHrs-totalSched)} sub={totalHrs>totalSched?"over schedule":"under schedule"} accent={totalHrs>totalSched?T.green:T.red} mono/>}
                  <KPI label="Active Staff"     value={totalStaff}            sub={reportLocFilter==="all"?"this week":`at ${reportLocFilter}`} accent={T.colony}/>
                  {currentSquare.length>0&&(()=>{const sq=reportLocFilter==="all"?currentSquare.reduce((s,d)=>s+d.gross,0):currentSquare.reduce((s,d)=>s+(d.byLocation?(d.byLocation[reportLocFilter]||0):0),0);return <KPI label="Square Sales" value={fmtUSD(sq)} sub={`${currentSquare.length} days`} accent={T.gold} mono/>;})()}
                </div>

                {/* Sales Summary callout — total + per-location vs goal */}
                {currentSquare.length>0&&(()=>{
                  const locsShown=reportLocFilter==="all"?LOCS:[reportLocFilter];
                  const perLoc=locsShown.map(loc=>({
                    location:loc,
                    total:currentSquare.reduce((s,d)=>s+(d.byLocation?(d.byLocation[loc]||0):0),0),
                  }));
                  const grand=perLoc.reduce((s,p)=>s+p.total,0);
                  const fmtStatus=t=>t>=GOAL_MAX?{label:"Exceeded",clr:T.green,bg:T.greenl}:t>=GOAL_MIN?{label:"On Goal",clr:T.green,bg:T.greenl}:{label:"Below Goal",clr:T.red,bg:T.redl};
                  return(
                    <div style={{background:`linear-gradient(135deg, ${T.navy} 0%, ${T.navy2} 100%)`,borderRadius:14,padding:"22px 26px",color:T.white,boxShadow:T.shadow2}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:14}}>
                        <div>
                          <div style={{fontSize:10,fontWeight:800,color:T.gold,letterSpacing:"0.14em",textTransform:"uppercase",marginBottom:4}}>Square Sales · This Week</div>
                          <div style={{fontFamily:"'DM Mono',monospace",fontSize:32,fontWeight:900,color:T.white,lineHeight:1.1}}>{fmtUSD(grand)}</div>
                          <div style={{fontSize:11,color:T.white+"88",marginTop:4}}>{currentSquare.length} day{currentSquare.length===1?"":"s"} · {currentWeek.period}</div>
                        </div>
                        <div style={{display:"flex",flexWrap:"wrap",gap:10}}>
                          {perLoc.map(p=>{
                            const s=fmtStatus(p.total);
                            return(
                              <div key={p.location} style={{background:T.white+"0E",border:`1px solid ${T.white}20`,borderRadius:10,padding:"10px 14px",minWidth:180}}>
                                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                                  <div style={{width:10,height:10,borderRadius:"50%",background:LOC_CLR[p.location]||T.gold}}/>
                                  <div style={{fontSize:12,fontWeight:700,color:T.white}}>{p.location}</div>
                                </div>
                                <div style={{fontFamily:"'DM Mono',monospace",fontSize:18,fontWeight:800,color:T.goldf}}>{fmtUSD(p.total)}</div>
                                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:6}}>
                                  <div style={{fontSize:9,color:T.white+"66"}}>Goal: ${(GOAL_MIN/1000).toFixed(1)}k–${(GOAL_MAX/1000).toFixed(0)}k</div>
                                  <span style={{background:s.clr+"30",color:s.clr==="#0F7A52"?"#A7E8C8":s.clr==="#C0392B"?"#F8B6AE":T.gold,border:`1px solid ${s.clr}55`,borderRadius:20,padding:"2px 8px",fontSize:9,fontWeight:800,letterSpacing:"0.05em",textTransform:"uppercase"}}>{s.label}</span>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  );
                })()}

                {/* Combined — all locations in one table (only when "all" selected) */}
                {reportLocFilter==="all"&&(
                <div style={{border:`1px solid ${T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow}}>
                  <div style={{background:`linear-gradient(90deg, ${T.navy} 0%, ${T.navy3} 100%)`,padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div style={{color:T.white,fontWeight:900,fontSize:14}}>🏢 Combined — All Locations</div>
                    <div style={{display:"flex",alignItems:"center",gap:16}}>
                      <div style={{color:T.white+"99",fontSize:11}}>{currentWeek.employees.length} staff · {fmtHHMM(totalHrs)}</div>
                      {payDelta!==undefined&&<Delta val={payDelta}/>}
                      <div style={{color:T.goldf,fontWeight:900,fontSize:18,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(totalPay)}</div>
                    </div>
                  </div>
                  <div style={{background:T.white,overflowX:"auto"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr style={{background:T.off}}>
                        {["Employee","Location","Role","Days","Sched","Worked","Var","Rate","Pay","Δ","Notes"].map(h=>(
                          <th key={h} style={{padding:"9px 12px",textAlign:["Pay","Rate","Sched","Worked","Var","Δ"].includes(h)?"right":"left",fontSize:9,fontWeight:800,color:T.muted,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {[...filteredEmps]
                          .sort((a,b)=>(a.location||"").localeCompare(b.location||"")||a.lastName.localeCompare(b.lastName))
                          .map((e,i)=>{
                            const ce=cmpWeekData?.employees.find(c=>c.firstName===e.firstName&&c.lastName===e.lastName);
                            const d=ce?((e.pay-ce.pay)/Math.max(ce.pay,1))*100:null;
                            const varHrs=e.hours-e.scheduledHours;
                            const locNorm=normalizeVenue(e.location);
                            const clr=LOC_CLR[locNorm]||LOC_CLR[e.location]||T.muted;
                            return <tr key={i} style={{background:i%2===0?T.white:"#F9FAFD",borderBottom:`1px solid ${T.border}`}}>
                              <td style={{padding:"9px 12px",fontWeight:700,color:T.navy}}>{e.firstName} {e.lastName}</td>
                              <td style={{padding:"9px 12px"}}><Pill color={clr}>{locNorm||e.location||"—"}</Pill></td>
                              <td style={{padding:"9px 12px"}}><Pill color={ROLE_CLR(e.role)}>{(e.role||"").replace("Senior Host","Sr. Host")}</Pill></td>
                              <td style={{padding:"9px 12px",textAlign:"center",color:T.muted}}>{e.workedDays||"—"}</td>
                              <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{e.scheduledHours>0?fmtHHMM(e.scheduledHours):"—"}</td>
                              <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace"}}>{fmtHHMM(e.hours)}</td>
                              <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontSize:11,fontWeight:700,color:varHrs>0?T.green:varHrs<0?T.red:T.muted}}>
                                {e.scheduledHours>0?(varHrs>0?"+":"")+fmtHHMM(Math.abs(varHrs)):"—"}
                              </td>
                              <td style={{padding:"9px 12px",textAlign:"right",color:T.muted}}>${e.rate}/hr</td>
                              <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:800}}>{fmtUSD(e.pay)}</td>
                              <td style={{padding:"9px 12px",textAlign:"right"}}>{d!==null?<Delta val={d}/>:"—"}</td>
                              <td style={{padding:"9px 12px",color:T.muted,fontSize:11,maxWidth:140}}>{e.notes||"—"}</td>
                            </tr>;
                          })}
                        <tr style={{background:T.navy+"08",borderTop:`2px solid ${T.navy}`}}>
                          <td colSpan={5} style={{padding:"10px 12px",fontWeight:800,color:T.navy,fontSize:11,textAlign:"right",letterSpacing:"0.06em",textTransform:"uppercase"}}>Week Total</td>
                          <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:800,color:T.navy}}>{fmtHHMM(totalHrs)}</td>
                          <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:700,color:totalSched>0?(totalHrs>totalSched?T.green:T.red):T.muted,fontSize:11}}>
                            {totalSched>0?((totalHrs-totalSched)>0?"+":"")+fmtHHMM(Math.abs(totalHrs-totalSched)):"—"}
                          </td>
                          <td/>
                          <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:900,color:T.gold,fontSize:14}}>{fmtUSD(totalPay)}</td>
                          <td/><td/>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
                )}

                {/* Location breakdown — when filter is "all", show all locations;
                    when a specific location is picked, show only that one's table. */}
                {LOCS
                  .filter(l=>reportLocFilter==="all"||l===reportLocFilter)
                  .filter(l=>currentWeek.employees.some(e=>e.location===l||normalizeVenue(e.location)===l))
                  .map(l=>{
                  const emps=currentWeek.employees.filter(e=>e.location===l||normalizeVenue(e.location)===l);
                  const cmpEmps=cmpWeekData?.employees.filter(e=>e.location===l||normalizeVenue(e.location)===l)||[];
                  const total=emps.reduce((s,e)=>s+e.pay,0);
                  const cmpTotal=cmpEmps.reduce((s,e)=>s+e.pay,0);
                  const delta=cmpTotal?((total-cmpTotal)/Math.max(cmpTotal,1))*100:null;
                  const clr=LOC_CLR[l]||T.navy;
                  return(
                    <div key={l} style={{border:`1px solid ${T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow}}>
                      <div style={{background:clr,padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <div style={{color:T.white,fontWeight:900,fontSize:14}}>🏟 {l}</div>
                        <div style={{display:"flex",alignItems:"center",gap:16}}>
                          {delta!==null&&<Delta val={delta}/>}
                          <div style={{color:T.goldf,fontWeight:900,fontSize:18,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(total)}</div>
                        </div>
                      </div>
                      <div style={{background:T.white,overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead><tr style={{background:T.off}}>
                            {["Employee","Role","Days","Sched","Worked","Var","Rate","Pay","Δ","Notes"].map(h=>(
                              <th key={h} style={{padding:"9px 12px",textAlign:["Pay","Rate","Sched","Worked","Var","Δ"].includes(h)?"right":"left",fontSize:9,fontWeight:800,color:T.muted,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                            ))}
                          </tr></thead>
                          <tbody>
                            {[...emps].sort((a,b)=>a.lastName.localeCompare(b.lastName)).map((e,i)=>{
                              const ce=cmpEmps.find(c=>c.firstName===e.firstName&&c.lastName===e.lastName);
                              const d=ce?((e.pay-ce.pay)/Math.max(ce.pay,1))*100:null;
                              const varHrs=e.hours-e.scheduledHours;
                              return <tr key={i} style={{background:i%2===0?T.white:"#F9FAFD",borderBottom:`1px solid ${T.border}`}}>
                                <td style={{padding:"9px 12px",fontWeight:700,color:T.navy}}>{e.firstName} {e.lastName}</td>
                                <td style={{padding:"9px 12px"}}><Pill color={ROLE_CLR(e.role)}>{(e.role||"").replace("Senior Host","Sr. Host")}</Pill></td>
                                <td style={{padding:"9px 12px",textAlign:"center",color:T.muted}}>{e.workedDays||"—"}</td>
                                <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{e.scheduledHours>0?fmtHHMM(e.scheduledHours):"—"}</td>
                                <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace"}}>{fmtHHMM(e.hours)}</td>
                                <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontSize:11,fontWeight:700,color:varHrs>0?T.green:varHrs<0?T.red:T.muted}}>
                                  {e.scheduledHours>0?(varHrs>0?"+":"")+fmtHHMM(Math.abs(varHrs)):"—"}
                                </td>
                                <td style={{padding:"9px 12px",textAlign:"right",color:T.muted}}>${e.rate}/hr</td>
                                <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:800}}>{fmtUSD(e.pay)}</td>
                                <td style={{padding:"9px 12px",textAlign:"right"}}>{d!==null?<Delta val={d}/>:"—"}</td>
                                <td style={{padding:"9px 12px",color:T.muted,fontSize:11,maxWidth:140}}>{e.notes||"—"}</td>
                              </tr>;
                            })}
                            <tr style={{background:clr+"10",borderTop:`2px solid ${clr}`}}>
                              <td colSpan={7} style={{padding:"9px 12px",fontWeight:800,color:clr,fontSize:11,textAlign:"right"}}>SUBTOTAL</td>
                              <td style={{padding:"9px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:900,color:clr,fontSize:14}}>{fmtUSD(total)}</td>
                              <td/><td/>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>
                  );
                })}

                {/* GOAL PERFORMANCE — per-location, expected vs actual share */}
                {goalPerformance.length>0&&goalPerformance.map(gp=>{
                  const teamPill=gp.teamStatus==="exceeded"?{label:"Goal Exceeded",bg:T.green,fg:T.white}:gp.teamStatus==="met"?{label:"On Goal",bg:T.green,fg:T.white}:{label:"Below Goal",bg:T.red,fg:T.white};
                  const pct=gp.weekTotal/GOAL_MAX*100;
                  const progressClr=gp.teamStatus==="below"?T.red:T.green;
                  const above=gp.employees.filter(e=>e.status==="above").length;
                  const below=gp.employees.filter(e=>e.status==="below").length;
                  return(
                    <div key={gp.location} style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,boxShadow:T.shadow,overflow:"hidden"}}>
                      <div style={{background:LOC_CLR[gp.location]||T.navy3,padding:"14px 20px",color:T.white,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
                        <div>
                          <div style={{fontSize:10,fontWeight:800,letterSpacing:"0.1em",textTransform:"uppercase",color:T.white+"BB"}}>🎯 Goal Performance</div>
                          <div style={{fontWeight:900,fontSize:16,marginTop:2}}>{gp.location}</div>
                        </div>
                        <div style={{display:"flex",alignItems:"center",gap:14}}>
                          <div style={{textAlign:"right"}}>
                            <div style={{fontSize:10,color:T.white+"99"}}>This Week</div>
                            <div style={{fontFamily:"'DM Mono',monospace",fontWeight:900,fontSize:18}}>{fmtUSD(gp.weekTotal)}</div>
                          </div>
                          <span style={{background:teamPill.bg,color:teamPill.fg,padding:"4px 10px",borderRadius:20,fontSize:10,fontWeight:800,letterSpacing:"0.05em",textTransform:"uppercase"}}>{teamPill.label}</span>
                        </div>
                      </div>
                      <div style={{padding:"16px 20px",borderBottom:`1px solid ${T.border}`}}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:T.muted,marginBottom:6}}>
                          <span>$0</span><span>Min ${(GOAL_MIN/1000).toFixed(1)}k</span><span>Max ${(GOAL_MAX/1000).toFixed(0)}k</span>
                        </div>
                        <div style={{position:"relative",height:10,background:T.off,borderRadius:5,overflow:"hidden"}}>
                          <div style={{position:"absolute",left:`${GOAL_MIN/GOAL_MAX*100}%`,top:0,bottom:0,width:1,background:T.navy3,zIndex:2}}/>
                          <div style={{height:"100%",width:`${Math.min(100,pct)}%`,background:progressClr,transition:"width .3s"}}/>
                        </div>
                        <div style={{marginTop:10,fontSize:11,color:T.muted}}>
                          {gp.employees.length} employee{gp.employees.length===1?"":"s"} worked · {gp.totalShifts} shift{gp.totalShifts===1?"":"s"} · <span style={{color:T.green,fontWeight:700}}>{above} at/above</span> · <span style={{color:T.red,fontWeight:700}}>{below} below</span>
                        </div>
                      </div>
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead><tr style={{background:T.off}}>
                            {["Employee","Shifts","% of Team","Expected","Actual","vs Expected","Status"].map(h=>(
                              <th key={h} style={{padding:"9px 12px",textAlign:["Shifts","% of Team","Expected","Actual","vs Expected"].includes(h)?"right":"left",fontSize:9,fontWeight:800,color:T.muted,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                            ))}
                          </tr></thead>
                          <tbody>
                            {gp.employees.map((e,i)=>{
                              const pct=gp.totalShifts>0?(e.shifts/gp.totalShifts)*100:0;
                              const clr=e.status==="above"?T.green:T.red;
                              return(
                                <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.white:"#F9FAFD"}}>
                                  <td style={{padding:"10px 12px",fontWeight:700,color:T.navy}}>{e.firstName} {e.lastName}</td>
                                  <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{e.shifts}</td>
                                  <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{pct.toFixed(1)}%</td>
                                  <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{fmtUSD(e.expected)}</td>
                                  <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:800,color:T.gold}}>{fmtUSD(e.actual)}</td>
                                  <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:700,color:clr}}>{e.deltaPct>0?"+":""}{e.deltaPct.toFixed(0)}%</td>
                                  <td style={{padding:"10px 12px"}}><span style={{background:clr+"18",color:clr,border:`1px solid ${clr}40`,padding:"3px 10px",borderRadius:20,fontSize:10,fontWeight:800,letterSpacing:"0.05em",textTransform:"uppercase"}}>{e.status==="above"?"✓ Above":"✗ Below"}</span></td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                      <div style={{padding:"10px 20px",background:T.off,fontSize:10,color:T.muted,fontStyle:"italic"}}>
                        Expected = (employee shifts ÷ team shifts) × week sales. Actual = sum of each day's location sales ÷ employees on shift that day.
                      </div>
                    </div>
                  );
                })}

                {/* Collapsible secondary analytics */}
                <Section title="📊 Staffing × Sales — Daily Activity" subtitle="Headcount and revenue overlaid by day" defaultOpen={false}>
                  <StaffingActivityChart week={filteredWeek} squareData={currentSquare}/>
                </Section>

                <Section title={`⏱ Scheduled vs Worked Hours${reportLocFilter!=="all"?` — ${reportLocFilter}`:""}`} subtitle="Per-employee schedule variance" defaultOpen={false}>
                  <ScheduleVsActual employees={filteredEmps}/>
                </Section>

                {salesAttribution.length>0&&(
                  <Section title="💰 Sales × Staff — by Day" subtitle="Who was on the floor when sales came in" defaultOpen={false} right={<span style={{fontSize:10,color:T.muted,marginRight:6}}>{salesAttribution.length} rows</span>}>
                    <div style={{overflowX:"auto"}}>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                        <thead><tr style={{background:T.off}}>
                          {["Date","Location","Sales","Hours","$/hr","Staff working"].map(h=>(
                            <th key={h} style={{padding:"9px 12px",textAlign:["Sales","Hours","$/hr"].includes(h)?"right":"left",fontSize:9,fontWeight:800,color:T.muted,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                          ))}
                        </tr></thead>
                        <tbody>
                          {salesAttribution.map((row,i)=>(
                            <tr key={row.date+row.location+i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.white:"#F9FAFD"}}>
                              <td style={{padding:"10px 12px",fontWeight:700,color:T.navy,fontFamily:"'DM Mono',monospace",whiteSpace:"nowrap"}}>{DAYS[dayOfWeek(row.date)]} {shortDate(row.date)}</td>
                              <td style={{padding:"10px 12px"}}><Pill color={LOC_CLR[row.location]||T.muted}>{row.location}</Pill></td>
                              <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:800,color:T.gold}}>{fmtUSD(row.sales)}</td>
                              <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{row.staffHours>0?row.staffHours.toFixed(1)+"h":"—"}</td>
                              <td style={{padding:"10px 12px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:700,color:row.staffHours>0?T.green:T.muted}}>{row.staffHours>0?fmtUSD(row.sales/row.staffHours):"—"}</td>
                              <td style={{padding:"10px 12px",color:T.navy}}>
                                {row.staff.length===0?<span style={{color:T.muted,fontSize:11}}>— no staff matched —</span>:(
                                  <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                                    {row.staff.map((s,j)=>(
                                      <span key={j} title={s.fullName} style={{fontSize:11,padding:"3px 8px",background:T.off,border:`1px solid ${T.border}`,borderRadius:14,whiteSpace:"nowrap"}}>
                                        {s.name} <span style={{color:T.muted,fontFamily:"'DM Mono',monospace",fontWeight:700}}>({s.hours.toFixed(1)}h)</span>
                                      </span>
                                    ))}
                                  </div>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </Section>
                )}

                {/* Square + Forms upload */}
                <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:22,boxShadow:T.shadow}}>
                  <div style={{fontWeight:800,color:T.navy,fontSize:15,marginBottom:16}}>📎 Sales & Documents — {currentWeek.period}</div>
                  <SquarePanel weekKey={selectedWeek} squareData={currentSquare}
                    onUpload={(files,type)=>type==="square"?handleSquareUpload(files,selectedWeek):handleFormUpload(files,selectedWeek)}/>
                  {weekAttachments.filter(({v})=>v.type==="square").length>0&&(
                    <div style={{marginTop:16}}>
                      <div style={{fontWeight:700,color:T.navy,fontSize:13,marginBottom:10}}>🟦 Uploaded Square Reports</div>
                      {weekAttachments.filter(({v})=>v.type==="square").map(({k,v})=>{
                        const rows=v.parsedData||[];
                        const total=rows.reduce((s,r)=>s+(+r.gross||0),0);
                        const locs=[...new Set(rows.map(r=>r.location).filter(Boolean))];
                        return(
                          <div key={k} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:T.off,borderRadius:9,marginBottom:6,border:`1px solid ${T.border}`}}>
                            <span style={{fontSize:18}}>🟦</span>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontWeight:700,color:T.navy,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{v.name}</div>
                              <div style={{fontSize:10,color:T.muted,marginTop:2}}>
                                {rows.length>0?<>{rows.length} day{rows.length===1?"":"s"} · <strong style={{color:T.gold,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(total)}</strong>{locs.length?<> · {locs.join(", ")}</>:null}</>:<>uploaded {new Date(v.uploadedAt).toLocaleDateString()}</>}
                              </div>
                            </div>
                            <button onClick={()=>handleDelAtt(k)} title="Delete this upload" style={{background:"none",border:`1px solid ${T.border}`,color:T.red,cursor:"pointer",fontSize:13,padding:"4px 10px",borderRadius:6,fontFamily:"inherit",fontWeight:700}}>Delete</button>
                          </div>
                        );
                      })}
                    </div>
                  )}
                  {weekAttachments.filter(({v})=>v.type==="form").length>0&&(
                    <div style={{marginTop:16}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                        <div style={{fontWeight:700,color:T.navy,fontSize:13}}>📋 Form Submissions</div>
                        <button onClick={()=>setTab("forms")} style={{background:"none",border:`1px solid ${T.border}`,padding:"4px 10px",borderRadius:6,fontSize:11,color:T.navy,cursor:"pointer",fontFamily:"inherit"}}>Open Forms tab →</button>
                      </div>
                      {weekAttachments.filter(({v})=>v.type==="form").map(({k,v})=>{
                        const subs=v.parsedData?.submissions||[];
                        const t=v.parsedData?(FORM_TYPES[v.parsedData.formType]||FORM_TYPES.other):FORM_TYPES.other;
                        return(
                          <div key={k} style={{padding:"10px 14px",background:T.off,borderRadius:9,marginBottom:6,border:`1px solid ${T.border}`}}>
                            <div style={{display:"flex",alignItems:"center",gap:10}}>
                              <span style={{fontSize:18}}>{t.icon}</span>
                              <div style={{flex:1,minWidth:0}}>
                                <div style={{fontWeight:700,color:T.navy,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{v.name}</div>
                                <div style={{fontSize:10,color:T.muted,marginTop:2}}>
                                  {subs.length>0?<>{t.label} · <strong style={{color:T.navy}}>{subs.length}</strong> submission{subs.length===1?"":"s"} from {new Set(subs.map(s=>s.submittedBy)).size} employee{new Set(subs.map(s=>s.submittedBy)).size===1?"":"s"}</>:<>Attached {new Date(v.uploadedAt).toLocaleDateString()} (not parsed)</>}
                                </div>
                              </div>
                              <button onClick={()=>handleDelAtt(k)} title="Delete this upload" style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:16}}>×</button>
                            </div>
                            {subs.length>0&&(
                              <div style={{marginTop:8,paddingTop:8,borderTop:`1px solid ${T.border}`,display:"flex",flexWrap:"wrap",gap:6}}>
                                {subs.slice(0,8).map((s,i)=>(
                                  <button key={i} onClick={()=>setFormDetail(s)} style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:"3px 10px",fontSize:11,cursor:"pointer",color:T.navy,fontFamily:"inherit",display:"flex",gap:6,alignItems:"center"}}>
                                    <span style={{fontWeight:700}}>{s.submittedBy}</span>
                                    <span style={{color:T.muted}}>· {s.submissionDate}</span>
                                    {s.summary.negative>0&&<span style={{color:T.red,fontWeight:800}}>⚠{s.summary.negative}</span>}
                                  </button>
                                ))}
                                {subs.length>8&&<span style={{fontSize:11,color:T.muted,alignSelf:"center"}}>+{subs.length-8} more</span>}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
        )}

        {/* ══ SCHEDULE ══ */}
        {tab==="schedule"&&(
          <div style={{display:"flex",flexDirection:"column",gap:20}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:12}}>
              <div>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:900,fontSize:26,color:T.navy,letterSpacing:"-0.02em"}}>Schedule View</div>
                <div style={{color:T.muted,fontSize:13,marginTop:3}}>Daily coverage grid, schedule vs actual, and sales overlay.</div>
              </div>
              {sortedKeys.length>0&&(
                <select value={selectedWeek||""} onChange={e=>setSelectedWeek(e.target.value)} style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${T.border}`,background:T.white,fontSize:13,fontWeight:700,color:T.navy,cursor:"pointer",fontFamily:"inherit"}}>
                  {sortedKeys.map(k=><option key={k} value={k}>{weeks[k].period}</option>)}
                </select>
              )}
            </div>

            {!currentWeek?(
              <div style={{textAlign:"center",padding:80,color:T.muted}}>
                <div style={{fontSize:48,marginBottom:16}}>📅</div>
                <div style={{fontWeight:800,fontSize:18,color:T.navy}}>No data yet</div>
                <button onClick={()=>setTab("upload")} style={{marginTop:20,padding:"10px 24px",borderRadius:9,border:"none",background:T.gold,color:T.navy,cursor:"pointer",fontWeight:800,fontSize:14,fontFamily:"inherit"}}>Upload →</button>
              </div>
            ):(
              <>
                <StaffingActivityChart week={currentWeek} squareData={currentSquare}/>
                <WeeklyGantt week={currentWeek}/>
                <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:22,boxShadow:T.shadow}}>
                  <div style={{fontWeight:800,color:T.navy,fontSize:15,marginBottom:16}}>⏱ Schedule vs Actual — All Staff</div>
                  <ScheduleVsActual employees={currentWeek.employees}/>
                </div>
              </>
            )}
          </div>
        )}

        {/* ══ COMPARE ══ */}
        {tab==="compare"&&(
          <div style={{display:"flex",flexDirection:"column",gap:20}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:12}}>
              <div>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:900,fontSize:26,color:T.navy,letterSpacing:"-0.02em"}}>Comparison</div>
                <div style={{color:T.muted,fontSize:13,marginTop:3}}>Payroll trends across weeks, months, years, or custom range.</div>
              </div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {[["week","Weekly"],["month","Monthly"],["year","Yearly"],["custom","Custom"]].map(([id,label])=>(
                  <button key={id} onClick={()=>setCompareMode(id)} style={{padding:"7px 16px",borderRadius:8,border:`1px solid ${compareMode===id?T.gold:T.border}`,background:compareMode===id?T.gold:T.white,color:compareMode===id?T.navy:T.navy,cursor:"pointer",fontWeight:700,fontSize:12,fontFamily:"inherit"}}>
                    {label}
                  </button>
                ))}
              </div>
            </div>

            {compareMode==="custom"&&(
              <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:12,padding:18,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontWeight:700,color:T.navy,fontSize:13}}>From:</span>
                <input type="date" value={customRange.from} onChange={e=>setCustomRange(r=>({...r,from:e.target.value}))} style={{padding:"7px 12px",borderRadius:8,border:`1px solid ${T.border}`,fontSize:13,fontFamily:"inherit"}}/>
                <span style={{fontWeight:700,color:T.navy,fontSize:13}}>To:</span>
                <input type="date" value={customRange.to} onChange={e=>setCustomRange(r=>({...r,to:e.target.value}))} style={{padding:"7px 12px",borderRadius:8,border:`1px solid ${T.border}`,fontSize:13,fontFamily:"inherit"}}/>
              </div>
            )}

            {groups.length===0?(
              <div style={{textAlign:"center",color:T.muted,padding:60,fontSize:13}}>No data — upload weekly files first.</div>
            ):(
              <>
                <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:"20px 20px 8px",boxShadow:T.shadow}}>
                  <div style={{fontWeight:800,color:T.navy,fontSize:14,marginBottom:16}}>Payroll by Location</div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={groups} barCategoryGap="30%">
                      <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
                      <XAxis dataKey="name" tick={{fontSize:11,fill:T.muted}}/>
                      <YAxis tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} tick={{fontSize:11,fill:T.muted}}/>
                      <Tooltip content={<CTooltip/>}/><Legend wrapperStyle={{fontSize:12}}/>
                      {LOCS.map(l=><Bar key={l} dataKey={l} fill={LOC_CLR[l]||T.navy} stackId="a" radius={l===LOCS[LOCS.length-1]?[4,4,0,0]:[0,0,0,0]}/>)}
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:"20px 20px 8px",boxShadow:T.shadow}}>
                  <div style={{fontWeight:800,color:T.navy,fontSize:14,marginBottom:16}}>Total Payroll Trend</div>
                  <ResponsiveContainer width="100%" height={200}>
                    <AreaChart data={groups}>
                      <defs><linearGradient id="tGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={T.gold} stopOpacity={0.25}/><stop offset="95%" stopColor={T.gold} stopOpacity={0}/></linearGradient></defs>
                      <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
                      <XAxis dataKey="name" tick={{fontSize:11,fill:T.muted}}/><YAxis tickFormatter={v=>`$${(v/1000).toFixed(0)}k`} tick={{fontSize:11,fill:T.muted}}/>
                      <Tooltip content={<CTooltip/>}/>
                      <Area type="monotone" dataKey="total" stroke={T.gold} strokeWidth={2.5} fill="url(#tGrad)" dot={{fill:T.gold,r:4,stroke:T.white,strokeWidth:2}}/>
                    </AreaChart>
                  </ResponsiveContainer>
                </div>

                <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow}}>
                  <div style={{background:T.navy,padding:"12px 20px"}}><span style={{color:T.gold,fontWeight:800,fontSize:13}}>Summary Table</span></div>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr style={{background:T.off}}>
                      {["Period","Dallas","The Colony","Total"].map(h=>(
                        <th key={h} style={{padding:"10px 14px",textAlign:h==="Period"?"left":"right",fontSize:10,fontWeight:800,color:T.muted,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                      ))}
                    </tr></thead>
                    <tbody>
                      {groups.map((g,i)=>(
                        <tr key={i} style={{background:i%2===0?T.white:"#F9FAFD",borderBottom:`1px solid ${T.border}`}}>
                          <td style={{padding:"10px 14px",fontWeight:700,color:T.navy}}>{g.name}</td>
                          {LOCS.map(l=><td key={l} style={{padding:"10px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:g[l]>0?T.text:T.muted}}>{g[l]>0?fmtUSD(g[l]):"—"}</td>)}
                          <td style={{padding:"10px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:900,color:T.navy}}>{fmtUSD(g.total)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        )}

        {/* ══ HISTORY ══ */}
        {tab==="history"&&(
          <div style={{display:"flex",flexDirection:"column",gap:20}}>
            <div>
              <div style={{fontFamily:"'Syne',sans-serif",fontWeight:900,fontSize:26,color:T.navy,letterSpacing:"-0.02em"}}>History</div>
              <div style={{color:T.muted,fontSize:13,marginTop:3}}>All stored weeks — click any row to view.</div>
            </div>
            {!sortedKeys.length?(
              <div style={{textAlign:"center",padding:80,color:T.muted}}><div style={{fontSize:48,marginBottom:16}}>📂</div><div style={{fontWeight:800,color:T.navy}}>No weeks stored</div></div>
            ):(
              <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{background:T.navy}}>
                    {["Pay Period","Staff","Hours","Sched","Var","Dallas","Colony","Total","Attached",""].map(h=>(
                      <th key={h} style={{padding:"11px 14px",textAlign:["Dallas","Colony","Total","Hours","Sched","Var"].includes(h)?"right":"left",color:T.gold,fontSize:9,fontWeight:800,letterSpacing:"0.1em",textTransform:"uppercase"}}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {sortedKeys.map((k,i)=>{
                      const w=weeks[k];
                      const byL={};LOCS.forEach(l=>{byL[l]=w.employees.filter(e=>e.location===l).reduce((s,e)=>s+e.pay,0);});
                      const total=w.employees.reduce((s,e)=>s+e.pay,0);
                      const hrs=w.employees.reduce((s,e)=>s+e.hours,0);
                      const sched=w.employees.reduce((s,e)=>s+e.scheduledHours,0);
                      const varHrs=sched>0?hrs-sched:null;
                      const staff=new Set(w.employees.map(e=>`${e.firstName} ${e.lastName}`)).size;
                      const attCount=Object.keys(attachments).filter(ak=>ak.startsWith(k+":")).length;
                      return(
                        <tr key={k} style={{background:i%2===0?T.white:"#F9FAFD",borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}
                          onClick={()=>{setSelectedWeek(k);setTab("report");}}>
                          <td style={{padding:"11px 14px",fontWeight:800,color:T.navy}}>{w.period}</td>
                          <td style={{padding:"11px 14px",color:T.muted}}>{staff}</td>
                          <td style={{padding:"11px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{fmtHHMM(hrs)}</td>
                          <td style={{padding:"11px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:T.muted}}>{sched>0?fmtHHMM(sched):"—"}</td>
                          <td style={{padding:"11px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontSize:11,fontWeight:700,color:varHrs===null?T.muted:varHrs>0?T.green:T.red}}>
                            {varHrs!==null?(varHrs>0?"+":"")+fmtHHMM(Math.abs(varHrs)):"—"}
                          </td>
                          {LOCS.map(l=><td key={l} style={{padding:"11px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",color:byL[l]>0?T.text:T.muted}}>{byL[l]>0?fmtUSD(byL[l]):"—"}</td>)}
                          <td style={{padding:"11px 14px",textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:900,color:T.navy}}>{fmtUSD(total)}</td>
                          <td style={{padding:"11px 14px",textAlign:"center"}}>{attCount>0?<Pill color={T.colony}>{attCount}</Pill>:<span style={{color:T.muted}}>—</span>}</td>
                          <td style={{padding:"11px 14px",textAlign:"center"}} onClick={e=>e.stopPropagation()}>
                            <button onClick={()=>setDeleteConfirm(k)} style={{background:"none",border:`1px solid ${T.border}`,color:T.muted,cursor:"pointer",fontSize:12,padding:"4px 10px",borderRadius:6,fontFamily:"inherit"}}>🗑</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* ══ FORMS ══ */}
        {tab==="forms"&&(
          <div style={{display:"flex",flexDirection:"column",gap:20}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:12}}>
              <div>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:900,fontSize:26,color:T.navy,letterSpacing:"-0.02em"}}>Form Submissions</div>
                <div style={{color:T.muted,fontSize:13,marginTop:3}}>Connecteam checklists — Pre-Open, Closing, Hourly Reset.</div>
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <div style={{color:T.muted,fontSize:11}}>Total: <strong style={{color:T.navy}}>{allFormSubmissions.length}</strong></div>
              </div>
            </div>

            <div style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,padding:20,boxShadow:T.shadow,maxWidth:560}}>
              <UploadZone
                onFiles={f=>handleFormUpload(f, selectedWeek||weekKey(new Date().toISOString().slice(0,10)))}
                label="Drop Connecteam Excel exports"
                sub="Pre-Open · Closing · Hourly Reset — multiple files OK"
                accept=".xlsx,.xls"
                icon="📋"
              />
              {!selectedWeek&&<div style={{marginTop:10,fontSize:11,color:T.amber}}>⚠ Tip: select a week in Report first so uploads link to it.</div>}
            </div>

            {allFormSubmissions.length>0&&(
              <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
                <select value={formFilter.type} onChange={e=>setFormFilter(f=>({...f,type:e.target.value}))} style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${T.border}`,background:T.white,fontSize:12,fontFamily:"inherit",color:T.navy}}>
                  <option value="all">All form types</option>
                  {Object.values(FORM_TYPES).filter(t=>t.id!=="other").map(t=><option key={t.id} value={t.id}>{t.icon} {t.label}</option>)}
                </select>
                <select value={formFilter.venue} onChange={e=>setFormFilter(f=>({...f,venue:e.target.value}))} style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${T.border}`,background:T.white,fontSize:12,fontFamily:"inherit",color:T.navy}}>
                  <option value="all">All venues</option>
                  {LOCS.map(l=><option key={l} value={l}>{l}</option>)}
                </select>
                <select value={formFilter.weekScope} onChange={e=>setFormFilter(f=>({...f,weekScope:e.target.value}))} style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${T.border}`,background:T.white,fontSize:12,fontFamily:"inherit",color:T.navy}}>
                  <option value="all">All weeks</option>
                  <option value="current">Selected week only</option>
                </select>
                <input value={formFilter.employee} onChange={e=>setFormFilter(f=>({...f,employee:e.target.value}))} placeholder="Filter by employee name…" style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${T.border}`,background:T.white,fontSize:12,fontFamily:"inherit",minWidth:220,color:T.navy}}/>
                <div style={{marginLeft:"auto",color:T.muted,fontSize:12}}>{filteredSubs.length} of {allFormSubmissions.length}</div>
              </div>
            )}

            {allFormSubmissions.length===0?(
              <div style={{textAlign:"center",padding:80,color:T.muted}}>
                <div style={{fontSize:48,marginBottom:16}}>📋</div>
                <div style={{fontWeight:800,fontSize:18,color:T.navy}}>No form submissions yet</div>
                <div style={{fontSize:13,marginTop:6}}>Upload a Connecteam Excel export above to get started.</div>
              </div>
            ):filteredSubs.length===0?(
              <div style={{textAlign:"center",padding:60,color:T.muted,background:T.white,border:`1px solid ${T.border}`,borderRadius:12}}>No submissions match the current filters.</div>
            ):(
              Object.entries(subsByEmployee).sort(([a],[b])=>a.localeCompare(b)).map(([emp,subs])=>(
                <div key={emp} style={{background:T.white,border:`1px solid ${T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow}}>
                  <div style={{background:T.navy,padding:"12px 18px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div style={{color:T.white,fontWeight:800,fontSize:14}}>👤 {emp}</div>
                    <div style={{color:T.goldf,fontSize:11,fontFamily:"'DM Mono',monospace"}}>{subs.length} submission{subs.length===1?"":"s"}</div>
                  </div>
                  <div>
                    {subs.map((s,i)=>{
                      const t=FORM_TYPES[s.formType]||FORM_TYPES.other;
                      return(
                        <div key={s.id+i} onClick={()=>setFormDetail(s)} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 16px",borderTop:i?`1px solid ${T.border}`:"none",cursor:"pointer",background:i%2===0?T.white:T.off}}>
                          <div style={{width:32,textAlign:"center",fontSize:18}}>{t.icon}</div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:700,color:T.navy,fontSize:13}}>{t.label}</div>
                            <div style={{color:T.muted,fontSize:11,marginTop:2}}>{s.submissionDate||"—"} {s.submissionTime||""} · {s.venue}</div>
                          </div>
                          <div style={{display:"flex",gap:10,fontSize:11,fontFamily:"'DM Mono',monospace",alignItems:"center"}}>
                            <span style={{color:T.green}}>✓ {s.summary.complete}</span>
                            {s.summary.negative>0&&<span style={{color:T.red,fontWeight:700}}>✗ {s.summary.negative}</span>}
                            {s.summary.naCount>0&&<span style={{color:T.muted}}>N/A {s.summary.naCount}</span>}
                            {s.summary.images>0&&<span style={{color:T.blue2}}>🖼 {s.summary.images}</span>}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))
            )}
          </div>
        )}
      </div>

      {formDetail&&<Modal wide title={`${FORM_TYPES[formDetail.formType]?.icon||"📋"} ${FORM_TYPES[formDetail.formType]?.label||"Form"} — ${formDetail.submittedBy}`} onClose={()=>setFormDetail(null)}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:18}}>
          <Pill color={T.navy3}>{formDetail.submissionDate||"—"}</Pill>
          {formDetail.submissionTime&&<Pill color={T.gold}>{formDetail.submissionTime}</Pill>}
          <Pill color={LOC_CLR[formDetail.venue]||T.muted}>{formDetail.venue}</Pill>
          <Pill color={T.green}>{formDetail.summary.complete} complete</Pill>
          {formDetail.summary.negative>0&&<Pill color={T.red}>{formDetail.summary.negative} issues</Pill>}
          {formDetail.summary.naCount>0&&<Pill color={T.muted}>{formDetail.summary.naCount} N/A</Pill>}
        </div>
        {formDetail.summary.flags.length>0&&(
          <div style={{background:T.redl,border:`1px solid ${T.red}40`,borderRadius:9,padding:"12px 14px",marginBottom:14}}>
            <div style={{fontWeight:800,color:T.red,fontSize:12,marginBottom:6,letterSpacing:"0.05em",textTransform:"uppercase"}}>⚠ Items marked No</div>
            {formDetail.summary.flags.map((f,i)=><div key={i} style={{fontSize:12,color:T.navy,marginBottom:2}}>• {f}</div>)}
          </div>
        )}
        <div style={{background:T.off,border:`1px solid ${T.border}`,borderRadius:10,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:T.navy}}>
              <th style={{padding:"9px 12px",textAlign:"left",color:T.gold,fontSize:9,fontWeight:800,letterSpacing:"0.1em",textTransform:"uppercase"}}>Question</th>
              <th style={{padding:"9px 12px",textAlign:"left",color:T.gold,fontSize:9,fontWeight:800,letterSpacing:"0.1em",textTransform:"uppercase",width:180}}>Answer</th>
            </tr></thead>
            <tbody>
              {Object.entries(formDetail.answers).map(([q,v],i)=>{
                const lc=String(v).toLowerCase();
                const clr=lc==="complete"||lc==="yes"?T.green:lc==="no"?T.red:lc==="image"?T.blue2:lc==="n/a"?T.muted:T.navy;
                const mono=["complete","yes","no","n/a","image"].includes(lc);
                return(
                  <tr key={i} style={{borderTop:i?`1px solid ${T.border}`:"none",background:i%2===0?T.white:T.off}}>
                    <td style={{padding:"8px 12px",color:T.navy}}>{q}</td>
                    <td style={{padding:"8px 12px",fontWeight:700,color:clr,fontFamily:mono?"'DM Mono',monospace":"inherit"}}>{v}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div style={{display:"flex",justifyContent:"flex-end",marginTop:18}}>
          <button onClick={()=>setFormDetail(null)} style={{padding:"9px 22px",borderRadius:9,border:"none",background:T.navy,color:T.white,cursor:"pointer",fontWeight:700,fontSize:13,fontFamily:"inherit"}}>Close</button>
        </div>
      </Modal>}
    </div>
  );
}
