// ============================================================
// WARD NURSING ROSTER SYSTEM v3.1 — Supabase + Roster Locking
// ============================================================
import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  loadStaff, saveStaff,
  loadRosters, saveRoster, updateRosterLock,
  loadNightPlan, saveNightPlan,
  loadADOAdjustments, saveADOAdjustment, deleteADOAdjustment,
} from "./supabase.js";

const CLASSIFICATIONS = {
  NUM:  { label:"Nurse Unit Manager",        color:"#9b59b6", inCharge:true  },
  ANUM: { label:"Assoc. Nurse Unit Manager", color:"#e74c3c", inCharge:true  },
  CNS:  { label:"Clinical Nurse Specialist", color:"#e67e22", inCharge:true  },
  RN:   { label:"Registered Nurse",          color:"#3498db", inCharge:false },
  GNP:  { label:"Graduate Nurse Program",    color:"#27ae60", inCharge:false },
  EN:   { label:"Enrolled Nurse",            color:"#1abc9c", inCharge:false },
};

const SHIFT_DEF = {
  D:  { label:"Day",            hours:8,  color:"#f39c12", bg:"#2a1f00", text:"#f39c12" },
  E:  { label:"Evening",        hours:8,  color:"#e67e22", bg:"#2a1200", text:"#e67e22" },
  N:  { label:"Night",          hours:10, color:"#7986cb", bg:"#0d1233", text:"#7986cb" },
  AL: { label:"Annual Leave",   hours:8,  color:"#9c27b0", bg:"#1a0a22", text:"#ce93d8" },
  SL: { label:"Sick Leave",     hours:8,  color:"#e74c3c", bg:"#220a0a", text:"#ef9a9a" },
  PDL:{ label:"Prof. Dev.",     hours:8,  color:"#2196f3", bg:"#0a1422", text:"#90caf9" },
  UL: { label:"Union Leave",    hours:8,  color:"#78909c", bg:"#111618", text:"#b0bec5" },
  ADO:{ label:"ADO",            hours:8,  color:"#43a047", bg:"#0a1a0a", text:"#a5d6a7" },
  LSL:{ label:"Long Svc Leave", hours:8,  color:"#00acc1", bg:"#001a1e", text:"#80deea" },
  PL: { label:"Parental Leave", hours:8,  color:"#ec407a", bg:"#220010", text:"#f48fb1" },
};

const CARD_CODES = ["LSL","PL"];
const DAY_NAMES   = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const FWA_CONDITIONS = {
  NO_NIGHTS:     "No night shifts",
  NO_WEEKENDS:   "No weekend shifts",
  DAYS_ONLY:     "Day shifts only",
  EVENINGS_ONLY: "Evening shifts only",
  SPECIFIC_DAYS: "Specific days of week only",
  SPECIFIC_SHIFTS:"Specific shift types only",
  MAX_HOURS_WEEK:"Max hours per week",
  REDUCED_NIGHTS:"Reduced night shift frequency",
  CUSTOM:        "Custom arrangement",
};

// Date/formatting utilities

// Parse a date string as LOCAL time (not UTC) to avoid timezone day-shift bugs.
// new Date("2026-05-25") parses as UTC midnight which in AEST = Sun 24 May locally.
// This function always returns the correct local date.
function parseLocalDate(dateStr) {
  if (!dateStr) return null;
  const [y, m, d] = dateStr.split("-").map(Number);
  return new Date(y, m - 1, d); // local midnight — no UTC shift
}

const addDays  = (d,n) => { const x=new Date(d); x.setDate(x.getDate()+n); return x; };
const isoDate  = d => {
  // Format as YYYY-MM-DD in LOCAL time (not UTC)
  const dt = new Date(d);
  const y  = dt.getFullYear();
  const m  = String(dt.getMonth()+1).padStart(2,"0");
  const dd = String(dt.getDate()).padStart(2,"0");
  return `${y}-${m}-${dd}`;
};
const fmtDate  = d => parseLocalDate(typeof d==="string"?d:isoDate(d))?.toLocaleDateString("en-AU",{day:"2-digit",month:"2-digit",year:"numeric"}) || "";
const fmtShort = d => { const dt=parseLocalDate(typeof d==="string"?d:isoDate(d))||new Date(d); return `${dt.getDate()} ${MONTH_NAMES[dt.getMonth()]}`; };
const dayIdx   = d => { const w=(parseLocalDate(typeof d==="string"?d:isoDate(d))||new Date(d)).getDay(); return w===0?6:w-1; };
const isWknd   = d => dayIdx(d)>=5;
// getMon: always returns the Monday of the week containing date d, in local time.
const getMon = d => {
  const x = parseLocalDate(typeof d==="string"?d:isoDate(d)) || new Date(d);
  x.setHours(0,0,0,0);
  const dow  = x.getDay(); // 0=Sun,1=Mon,...,6=Sat
  const diff = dow===0 ? -6 : 1-dow; // Mon→0, Sun→-6, Tue→-1, etc.
  x.setDate(x.getDate()+diff);
  return x;
};

// buildDays: always produces exactly weeks*7 days starting from startMon (local Monday).
function buildDays(startMon, weeks) {
  const total = weeks * 7; // 2 weeks = 14, 4 weeks = 28
  return Array.from({length: total}, (_, i) => {
    const date = addDays(startMon, i);
    const di   = dayIdx(date); // 0=Mon … 6=Sun (local)
    return {
      date,
      iso:  isoDate(date), // local YYYY-MM-DD
      di,
      wknd: di >= 5,       // Sat=5, Sun=6
      wk:   Math.floor(i / 7),
    };
  });
}

const isInCharge  = s => !!(CLASSIFICATIONS[s.cls]?.inCharge || s.inCharge);
const fullName    = s => s ? `${s.firstName||""} ${s.lastName||""}`.trim() || s.name || "" : "";

function fwaAllows(s,iso,shift) {
  for (const c of (s.fwaConditions||[])) {
    if (c.type==="NO_NIGHTS"       && shift==="N")                              return false;
    if (c.type==="NO_WEEKENDS"     && isWknd(iso))                              return false;
    if (c.type==="DAYS_ONLY"       && shift!=="D")                              return false;
    if (c.type==="EVENINGS_ONLY"   && shift!=="E")                              return false;
    if (c.type==="SPECIFIC_DAYS"   && c.days   && !c.days.includes(dayIdx(iso))) return false;
    if (c.type==="SPECIFIC_SHIFTS" && c.shifts && !c.shifts.includes(shift))    return false;
  }
  return true;
}

function buildLeaveMap(staff) {
  const lm={};
  staff.forEach(s=>{
    lm[s.id]={...(s.gridLeave||{})};
    (s.leaveCard||[]).forEach(le=>{
      let cur=new Date(le.from); const end=new Date(le.to);
      while(cur<=end){ lm[s.id][isoDate(cur)]=le.code; cur=addDays(cur,1); }
    });
  });
  return lm;
}

function nightAdjustedHrs(s) {
  // Hours a staff member can contribute to nights per fortnight block
  // Nights are 10h shifts — round down contracted hours to nearest 10h multiple
  return Math.floor(s.hrs / 10) * 10;
}

function autoComputeNightGroups(staff) {
  const eligible = staff.filter(s =>
    !s.permNights && s.cls!=="NUM" &&
    !s.fwaConditions?.some(c=>c.type==="NO_NIGHTS"||
      (c.type==="SPECIFIC_SHIFTS" && c.shifts && !c.shifts.includes("N")))
  );
  const permNights     = staff.filter(s=>s.permNights);
  // Perm nights also contribute night-adjusted hours
  const permHoursPerFn = permNights.reduce((a,s)=>a+nightAdjustedHrs(s),0);

  // Base requirement: 5 staff × 10h × 14 nights = 700h per fortnight block
  // With 25% buffer: 700 × 1.25 = 875h target
  const BASE_REQUIRED  = 700;
  const BUFFER         = 1.25;
  const TARGET_HOURS   = BASE_REQUIRED * BUFFER; // 875h
  const rotNeeded      = Math.max(0, TARGET_HOURS - permHoursPerFn);

  // Each eligible staff contributes their night-adjusted hours per block
  const avgNightHrs    = eligible.length
    ? eligible.reduce((a,s)=>a+nightAdjustedHrs(s),0) / eligible.length
    : 70;

  // How many staff per block needed to hit the buffered target
  const staffPerBlock  = Math.max(1, Math.ceil(rotNeeded / avgNightHrs));

  // Number of groups = total eligible / staff per block
  // Fewer groups = each group is on nights more frequently (aim for 6-8wk gap)
  const numGroups      = Math.max(2, Math.ceil(eligible.length / staffPerBlock));

  // Sort: ensure each group has in-charge coverage, then by classification priority, then hrs desc
  const clsPri = {ANUM:0,CNS:1,RN:2,GNP:3,EN:4};
  const sorted  = [...eligible].sort((a,b)=>{
    const cp=(clsPri[a.cls]||5)-(clsPri[b.cls]||5); if(cp!==0)return cp;
    return b.hrs-a.hrs;
  });

  const groups = Array.from({length:numGroups},(_,i)=>({
    id:i+1, members:[], totalHours:0, nightHours:0
  }));

  // Round-robin assignment
  sorted.forEach((s,i)=>{
    const g=groups[i%numGroups];
    g.members.push(s.id);
    g.totalHours  += s.hrs;
    g.nightHours  += nightAdjustedHrs(s);
  });

  return { groups, numGroups, staffPerBlock, permHoursPerFn,
           targetHours:TARGET_HOURS, baseRequired:BASE_REQUIRED, bufferPct:25 };
}

function autoAssignNightPlan(staff, year, firstMonday) {
  const {groups}=autoComputeNightGroups(staff);
  const plan={}, groupAssignments={};

  // Use the provided first Monday, or fall back to first Monday of the year
  const startMon = firstMonday ? getMon(parseLocalDate(firstMonday)) : getMon(new Date(year,0,1));

  // Build 26 fortnights from that start date
  const fns=Array.from({length:26},(_,i)=>({
    idx:i,
    start: addDays(startMon, i*14),
    end:   addDays(startMon, i*14+13),
    key:   isoDate(addDays(startMon, i*14)),
  }));

  // Assign groups to fortnights with max gap between reuse (6-8 week / 3-4 fn gap)
  const lastUsed={};
  fns.forEach((fn,i)=>{
    let best=null,bestGap=-1;
    groups.forEach(g=>{ const gap=i-(lastUsed[g.id]??-99); if(gap>bestGap){bestGap=gap;best=g;} });
    groupAssignments[fn.key]=best.id; lastUsed[best.id]=i;
  });

  // Build per-staff plan from group assignments
  fns.forEach(fn=>{
    const group=groups.find(g=>g.id===groupAssignments[fn.key]); if(!group)return;
    group.members.forEach(sid=>{
      for(let d=0;d<14;d++){
        const iso=isoDate(addDays(fn.start,d));
        if(!plan[sid])plan[sid]={};
        plan[sid][iso]=true;
      }
    });
  });

  return {plan, groupAssignments, groups, fns, firstMonday: isoDate(startMon), year};
}

const SAMPLE_STAFF = [
  {id:"s1", firstName:"Alexandra",cls:"NUM",  hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2020-01-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s2", firstName:"James",    lastName:"Hartley",  cls:"ANUM",hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2019-03-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s3", firstName:"Maria",    lastName:"Santos",   cls:"ANUM",hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"Prefers day shifts",   resign:null,commencementDate:"2018-06-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s4", firstName:"David",    lastName:"Okonkwo",  cls:"ANUM",hrs:64,permNights:false,inCharge:true, fwaConditions:[{type:"SPECIFIC_DAYS",days:[0,1,2,3,4],note:"Mon-Fri only (FWA)"}],prefs:"",resign:null,commencementDate:"2021-02-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s5", firstName:"Priya",    lastName:"Patel",    cls:"CNS", hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2017-09-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s6", firstName:"Tom",      lastName:"Nguyen",   cls:"CNS", hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2016-11-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s7", firstName:"Sarah",    lastName:"Kim",      cls:"RN",  hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2020-07-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s8", firstName:"Luke",     lastName:"Andersen", cls:"RN",  hrs:64,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2022-01-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s9", firstName:"Fatima",   lastName:"Al-Rawi",  cls:"RN",  hrs:48,permNights:false,inCharge:false,fwaConditions:[{type:"NO_NIGHTS",note:"FWA approved"},{type:"SPECIFIC_DAYS",days:[0,1,2,3,4],note:"Weekdays only"}],prefs:"",resign:null,commencementDate:"2021-08-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s10",firstName:"Grace",    lastName:"Torres",   cls:"RN",  hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2019-05-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s11",firstName:"Ben",      lastName:"Murphy",   cls:"RN",  hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"Happy to work weekends",resign:null,commencementDate:"2020-03-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s12",firstName:"Amara",    lastName:"Diallo",   cls:"RN",  hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2018-10-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s13",firstName:"Chen",     lastName:"Wei",      cls:"RN",  hrs:64,permNights:false,inCharge:false,fwaConditions:[{type:"MAX_HOURS_WEEK",value:32,note:"Max 32h/wk (FWA)"}],prefs:"",resign:null,commencementDate:"2023-01-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s14",firstName:"Nina",     lastName:"Rodriguez",cls:"RN",  hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2021-04-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s15",firstName:"Oscar",    lastName:"Pietersen",cls:"RN",  hrs:48,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2022-06-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s16",firstName:"Yuki",     lastName:"Tanaka",   cls:"GNP", hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2025-01-15",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s17",firstName:"Chloe",    lastName:"Martin",   cls:"GNP", hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2025-01-15",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s18",firstName:"Raj",      lastName:"Sharma",   cls:"GNP", hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2025-03-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s19",firstName:"Mei",      lastName:"Lin",      cls:"EN",  hrs:64,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,commencementDate:"2022-09-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s20",firstName:"Patrick",  lastName:"Flynn",    cls:"EN",  hrs:80,permNights:true, inCharge:false,fwaConditions:[],prefs:"Permanent nights",     resign:null,commencementDate:"2018-01-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s21",firstName:"Zara",     lastName:"Ahmed",    cls:"EN",  hrs:80,permNights:true, inCharge:false,fwaConditions:[],prefs:"Permanent nights",     resign:null,commencementDate:"2019-07-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s22",firstName:"Sofia",    lastName:"Russo",    cls:"RN",  hrs:80,permNights:true, inCharge:true, fwaConditions:[],prefs:"Permanent nights",     resign:null,commencementDate:"2017-03-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s23",firstName:"James",    lastName:"Brennan",  cls:"RN",  hrs:80,permNights:false,inCharge:false,fwaConditions:[{type:"EVENINGS_ONLY",note:"Evening only (FWA)"}],prefs:"",resign:null,commencementDate:"2023-06-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s24",firstName:"Lily",     lastName:"Thompson", cls:"RN",  hrs:64,permNights:false,inCharge:true, fwaConditions:[{type:"SPECIFIC_SHIFTS",shifts:["D","E"],note:"D/E only, no nights (FWA)"}],prefs:"",resign:null,commencementDate:"2024-02-01",leaveCard:[],gridLeave:{},requests:{}},
];

// ─── ROSTER GENERATOR v3 ─────────────────────────────────────
function generateRoster({staff,startDate,weeks,nightPlanData,previousRoster,recentWkndCounts,bumpHistory={}}) {
  const startMon=getMon(parseLocalDate(startDate));
  const days=buildDays(startMon,weeks);
  const roster={};
  days.forEach(d=>{ roster[d.iso]={D:[],E:[],N:[]}; });

  // Only roster active staff — commenced before roster end, not yet resigned/archived
  const rosterEnd = isoDate(addDays(startMon, weeks*7-1));
  const activeStaff = staff.filter(s=>{
    if(s.archived)return false;
    if(s.commencementDate && s.commencementDate > rosterEnd)return false;
    if(s.resign && s.resign < isoDate(startMon))return false;
    return true;
  });
  // Use activeStaff throughout generator instead of staff
  const leaveMap=buildLeaveMap(activeStaff);

  // Period targets — contracted hours is already per fortnight (2 weeks)
  const multiplier = weeks / 2;
  const periodTarget = {};
  const maxShifts = {};
  activeStaff.forEach(s => {
    const raw = s.hrs * multiplier;
    periodTarget[s.id] = s.permNights ? Math.floor(raw / 10) * 10 : raw;
  });

  // Hours worked — pre-load leave hours FIRST so maxShifts accounts for them
  const hw = {};
  activeStaff.forEach(s => { hw[s.id] = 0; });
  activeStaff.forEach(s => {
    days.forEach(d => {
      const lc = leaveMap[s.id]?.[d.iso];
      if (!lc || !SHIFT_DEF[lc]) return;
      hw[s.id] += (s.permNights || nightPlanData?.plan?.[s.id]?.[d.iso]) ? 10 : 8;
    });
  });

  // maxShifts: hard ceiling — canWork() refuses once reached.
  activeStaff.forEach(s => {
    const remainingHrs = Math.max(0, periodTarget[s.id] - hw[s.id]);
    maxShifts[s.id] = s.permNights ? Math.floor(remainingHrs/10) : Math.floor(remainingHrs/8);
  });

  // Shift count tracker
  const shiftCount = {};
  activeStaff.forEach(s => { shiftCount[s.id] = 0; });

  // ADO tracking
  const adoAccrued={},adoTaken={},adoMap={};
  activeStaff.forEach(s=>{ adoAccrued[s.id]=0; adoTaken[s.id]=0; adoMap[s.id]={}; });

  // Weekend counts (seeded from history)
  const wkndCnt={};
  activeStaff.forEach(s=>{ wkndCnt[s.id]=(recentWkndCounts?.[s.id]||0); });

  const prevTail=previousRoster?.tail||{};

  function onShift(sId,iso,sh){ return roster[iso]?.[sh]?.includes(sId)||prevTail[iso]?.[sh]?.includes(sId)||false; }
  function working(sId,iso){ return onShift(sId,iso,"D")||onShift(sId,iso,"E")||onShift(sId,iso,"N"); }
  function assignedToday(sId,iso){ const d=roster[iso]; return !!(d&&(d.D.includes(sId)||d.E.includes(sId)||d.N.includes(sId))); }

  function consecNights(sId,iso){ let c=0; for(let i=1;i<=5;i++){if(onShift(sId,isoDate(addDays(parseLocalDate(iso),-i)),"N"))c++;else break;} return c; }
  function consecShifts(sId,iso){ let c=0; for(let i=1;i<=6;i++){if(working(sId,isoDate(addDays(parseLocalDate(iso),-i))))c++;else break;} return c; }
  function nightWithin47h(sId,iso){ for(let i=1;i<=2;i++){if(onShift(sId,isoDate(addDays(parseLocalDate(iso),-i)),"N"))return true;} return false; }

  function weekHrs(sId,iso,adding){
    const mon=getMon(parseLocalDate(iso)); let tot=adding;
    for(let i=0;i<7;i++){
      const k=isoDate(addDays(mon,i)); if(k===iso)continue;
      if(leaveMap[sId][k]||adoMap[sId][k]){tot+=8;continue;}
      if(roster[k]?.D.includes(sId)||roster[k]?.E.includes(sId))tot+=8;
      if(roster[k]?.N.includes(sId))tot+=10;
    }
    return tot;
  }

  function canWork(s, iso, shift) {
    if (!s || s.archived) return false;
    // Must have commenced before this date
    if (s.commencementDate && parseLocalDate(iso) < parseLocalDate(s.commencementDate)) return false;
    // Must not have resigned — resignation date is last day of work, so exclude days AFTER it
    if (s.resign && parseLocalDate(iso) > parseLocalDate(s.resign)) return false;
    if (leaveMap[s.id][iso]) return false;
    if (adoMap[s.id][iso])   return false;
    if (assignedToday(s.id, iso)) return false;
    if (s.permNights && shift !== "N") return false;
    if (!s.permNights && shift === "N" && !nightPlanData?.plan?.[s.id]?.[iso]) return false;
    if (!fwaAllows(s, iso, shift)) return false;
    // DUAL HARD CEILING
    const shiftHrs = shift === "N" ? 10 : 8;
    if (hw[s.id] + shiftHrs > periodTarget[s.id]) return false;
    if (shiftCount[s.id] >= maxShifts[s.id]) return false;
    const maxW = s.fwaConditions?.find(c => c.type === "MAX_HOURS_WEEK")?.value;
    if (maxW && weekHrs(s.id, iso, shiftHrs) > maxW) return false;
    // No night shifts in first 3 months for ALL staff (not just GNP)
    if (s.commencementDate && shift === "N") {
      const cutoff = addDays(parseLocalDate(s.commencementDate), 91);
      if (parseLocalDate(iso) < cutoff) return false;
    }
    if (shift === "N" && consecNights(s.id, iso) >= 4) return false;
    if (shift !== "N" && consecShifts(s.id, iso) >= 5) return false;
    if ((shift === "D" || shift === "E") && nightWithin47h(s.id, iso)) return false;
    // E→N block: can't do Night if worked Evening yesterday (< 47h gap)
    if (shift === "E" && onShift(s.id, isoDate(addDays(parseLocalDate(iso), -1)), "N")) return false;
    // E→D short changeover: allowed once per week, unless requested
    if (shift === "D") {
      const prevIso = isoDate(addDays(parseLocalDate(iso), -1));
      if (onShift(s.id, prevIso, "E")) {
        const alreadyHadEtoD = countEtoDThisWeek(s.id, isoDate(addDays(parseLocalDate(iso), -1))) >= 1;
        const requested = !!(s.requests?.[`${iso}_D`]);
        if (alreadyHadEtoD && !requested) return false;
      }
    }
    return true;
  }

  // Count E→D short changeovers in the same calendar week as iso
  // E→D = worked Evening yesterday, working Day today (9.5h gap)
  // Allowed once per week per staff member (unless requested)
  function countEtoDThisWeek(sId, iso) {
    const mon = getMon(parseLocalDate(iso));
    let count = 0;
    for (let i = 1; i < 7; i++) {
      const dayIso  = isoDate(addDays(mon, i));
      const prevIso = isoDate(addDays(mon, i - 1));
      if (dayIso > iso) break;
      const workedDtoday = roster[dayIso]?.D.includes(sId) || prevTail[dayIso]?.D?.includes(sId);
      const workedEprev  = roster[prevIso]?.E.includes(sId) || prevTail[prevIso]?.E?.includes(sId);
      if (workedDtoday && workedEprev) count++;
    }
    return count;
  }

  function tryInsertADO(s,upToIso,nightCtx){
    if(s.hrs<80)return;
    const thresh=nightCtx?10:8;
    if(adoAccrued[s.id]<thresh)return;
    adoAccrued[s.id]-=thresh; adoTaken[s.id]++;
    let best=null,bestScore=999;
    days.forEach(d=>{
      if(d.iso>upToIso)return;
      if(leaveMap[s.id][d.iso]||adoMap[s.id][d.iso]||assignedToday(s.id,d.iso))return;
      if(nightCtx&&!s.permNights&&!nightPlanData?.plan?.[s.id]?.[d.iso])return;
      if(!nightCtx&&d.wknd)return;
      const prevK=isoDate(addDays(d.date,-1)),nextK=isoDate(addDays(d.date,1));
      const score=(!working(s.id,prevK)&&!leaveMap[s.id][prevK]?0:1)+(!assignedToday(s.id,nextK)&&!leaveMap[s.id][nextK]?0:1);
      if(score<bestScore){bestScore=score;best=d.iso;}
    });
    if(best){adoMap[s.id][best]=true;hw[s.id]+=thresh;}
  }

  const totalGNP=activeStaff.filter(s=>s.cls==="GNP").length;
  const maxGNPShift=Math.max(1,Math.floor(totalGNP/2));
  let numHasWorked=false;

  // ── UPFRONT SURPLUS RESOLUTION ────────────────────────────
  // Determine who is in the night plan for this fortnight
  const inNightPlanThisFn = new Set(
    activeStaff
      .filter(s => !s.permNights && days.some(d => nightPlanData?.plan?.[s.id]?.[d.iso]))
      .map(s => s.id)
  );

  // Calculate actual night-hours need for this fortnight
  // 5 staff × 10h × 14 nights = 700h, minus perm nights contribution,
  // minus leave taken by night-plan staff during this period
  const permNightStaff = activeStaff.filter(s => s.permNights);
  const permNightHrs   = permNightStaff.reduce((a,s) => a + nightAdjustedHrs(s), 0);

  // Leave hours already charged to night-plan rotating staff
  const rotatingNightLeaveHrs = activeStaff
    .filter(s => inNightPlanThisFn.has(s.id))
    .reduce((total, s) => {
      const leaveCount = days.filter(d => leaveMap[s.id]?.[d.iso]).length;
      return total + leaveCount * 10;
    }, 0);

  // Effective rotating-hours available from night-plan staff
  const rotatingNightStaff = activeStaff.filter(s => inNightPlanThisFn.has(s.id));
  const rotatingNightHrsAvailable = rotatingNightStaff.reduce((a,s) => {
    const adjHrs = nightAdjustedHrs(s);
    const leaveHrs = days.filter(d => leaveMap[s.id]?.[d.iso]).length * 10;
    return a + Math.max(0, adjHrs - leaveHrs);
  }, 0);

  const nightHrsNeeded = Math.max(0, 700 - permNightHrs); // rotating need

  // Surplus = how many rotating hours over what's needed
  const surplusHrs = rotatingNightHrsAvailable - nightHrsNeeded;

  // Identify surplus staff to bump — using bump history for fairness
  // bumpHistory: { staffId -> bumpCount } passed in from previous rosters
  const bumpCounts = bumpHistory;

  // Never bump: ANUM, GNP, permanent nights
  // Bumpable: RN, CNS, EN (with EN last-resort protection)
  const bumpable = rotatingNightStaff.filter(s =>
    !s.permNights &&
    s.cls !== "ANUM" &&
    s.cls !== "GNP"
  );

  // EN protection: count ENs in night plan for this fortnight
  const enInNightPlan = bumpable.filter(s => s.cls === "EN");
  const nonEnBumpable = bumpable.filter(s => s.cls !== "EN");

  // Sort bumpable by: fewest historical bumps first, then alpha
  const sortBumpable = (arr) => [...arr].sort((a,b) => {
    const ba = bumpCounts[a.id] || 0;
    const bb = bumpCounts[b.id] || 0;
    if (ba !== bb) return ba - bb;
    return fullName(a).localeCompare(fullName(b));
  });

  const bumpedFromNights = new Set(); // staff IDs bumped for entire fortnight
  const bumpedDetails   = [];         // { staffId, name, reason, bumpCount }

  if (surplusHrs > 0) {
    // Work through bumpable candidates until surplus is resolved
    // Try non-EN first, then EN if still surplus remains
    const candidates = [...sortBumpable(nonEnBumpable), ...sortBumpable(enInNightPlan)];

    let remainingSurplus = surplusHrs;
    for (const s of candidates) {
      if (remainingSurplus <= 0) break;

      // EN protection: don't bump if this is the last EN in night plan
      if (s.cls === "EN") {
        const enRemaining = enInNightPlan.filter(e => !bumpedFromNights.has(e.id));
        if (enRemaining.length <= 1) break; // protect last EN
      }

      // Coverage check: don't bump if it would leave nights with no in-charge
      // (we check this after, it's a soft guide — we'll enforce via warnings)

      bumpedFromNights.add(s.id);
      const adjHrs = nightAdjustedHrs(s);
      const leaveHrs = days.filter(d => leaveMap[s.id]?.[d.iso]).length * 10;
      remainingSurplus -= Math.max(0, adjHrs - leaveHrs);

      bumpedDetails.push({
        staffId:   s.id,
        name:      fullName(s),
        reason:    `Surplus night coverage — rostered to Day/Evening for this entire fortnight`,
        bumpCount: (bumpCounts[s.id] || 0) + 1,
      });
    }
  }

  // Track night hours ceiling and assigned
  const nightHoursCeiling = {};
  activeStaff.forEach(s => {
    nightHoursCeiling[s.id] = s.permNights
      ? periodTarget[s.id]
      : Math.floor((s.hrs * (weeks/2)) / 10) * 10;
  });
  const nightHrsAssigned = {};
  activeStaff.forEach(s => { nightHrsAssigned[s.id] = 0; });

  // Helper: is this person eligible for nights on this day?
  function eligibleForNights(s, iso) {
    if (bumpedFromNights.has(s.id)) return false; // bumped for whole fortnight
    if (!canWork(s, iso, "N")) return false;
    return true;
  }

  // Helper: in-charge coverage check for a night shift roster
  function nightInChargeCoverage(shiftStaffIds) {
    const onShiftStaff = shiftStaffIds.map(id => activeStaff.find(x=>x.id===id)).filter(Boolean);
    const anums   = onShiftStaff.filter(s => s.cls === "ANUM");
    const icCapable = onShiftStaff.filter(s => isInCharge(s) && s.cls !== "ANUM");
    if (anums.length >= 1) {
      return icCapable.length >= 1;
    } else {
      return icCapable.length >= 2;
    }
  }

  // ── Week-balanced night targets ───────────────────────────
  // For rotating staff: split their total nights roughly half week1/half week2
  // (give or take 1 shift), to avoid front-loading one week and leaving gaps.
  const week1Days = days.filter(d => d.wk === 0);
  const week2Days = days.filter(d => d.wk === 1);
  const nightWeekTarget = {}; // staffId -> { w1: targetShifts, w2: targetShifts }
  activeStaff.forEach(s => {
    if (s.permNights || bumpedFromNights.has(s.id)) return;
    const totalShifts = Math.floor(nightHoursCeiling[s.id] / 10);
    const w1 = Math.ceil(totalShifts / 2);
    const w2 = totalShifts - w1;
    nightWeekTarget[s.id] = { w1, w2, assignedW1: 0, assignedW2: 0 };
  });

  // Helper: does this person have a 2-3 day rest block coming up if we
  // assign them tonight? We avoid more than 4 consecutive nights and
  // try to ensure when they DO stop, they get 2-3 days off before
  // the next block (handled by canWork's consecNights check + this scoring).
  function nightAssignScore(s, iso, wk) {
    if (s.permNights) return 0; // perm nights always score equally
    const target = nightWeekTarget[s.id];
    if (!target) return 0;
    const assignedThisWeek = wk === 0 ? target.assignedW1 : target.assignedW2;
    const targetThisWeek    = wk === 0 ? target.w1 : target.w2;
    // Strongly prefer staff who are behind their target for this week
    const deficit = targetThisWeek - assignedThisWeek;
    return deficit; // higher deficit = higher priority
  }

  // ── PHASE 1: Night shifts ─────────────────────────────────
  const nightInChargeMissing = new Set();

  days.forEach(({iso, wk}) => {
    const eligible = activeStaff
      .filter(s => eligibleForNights(s, iso))
      .sort((a,b) => {
        if (a.permNights && !b.permNights) return -1;
        if (!a.permNights && b.permNights) return  1;
        // Week-balance priority: staff behind their week target go first
        const scoreA = nightAssignScore(a, iso, wk);
        const scoreB = nightAssignScore(b, iso, wk);
        if (scoreA !== scoreB) return scoreB - scoreA;
        return (periodTarget[b.id]-hw[b.id]) - (periodTarget[a.id]-hw[a.id]);
      });

    let enCnt=0, gnpCnt=0, anumCnt=0;

    for (const s of eligible) {
      if (roster[iso].N.length >= 5) break;
      if (nightHrsAssigned[s.id] + 10 > nightHoursCeiling[s.id]) continue;

      if (s.cls === "EN" && enCnt >= 1) continue;
      if (s.cls === "GNP" && gnpCnt >= 1) continue;
      if (s.cls === "ANUM") {
        if (anumCnt >= 1) {
          // Two-ANUM exception: both must be 80h FT, AND every other night
          // shift this week must already have exactly 1 ANUM
          const other = roster[iso].N.map(id=>activeStaff.find(x=>x.id===id)).find(x=>x?.cls==="ANUM");
          if (!other || other.hrs < 80 || s.hrs < 80) continue;
          // Check all other nights this week have exactly 1 ANUM
          const weekStart = addDays(parseLocalDate(iso), -dayIdx(iso)); // Mon of this week
          const otherNightsThisWeek = Array.from({length:7},(_,i)=>isoDate(addDays(weekStart,i)))
            .filter(k => k !== iso && roster[k]);
          const allOthersHaveOneAnum = otherNightsThisWeek.every(k => {
            const anumCount = roster[k].N.filter(id=>activeStaff.find(x=>x.id===id)?.cls==="ANUM").length;
            return anumCount === 1;
          });
          if (!allOthersHaveOneAnum) continue;
        }
        anumCnt++;
      }
      if (s.cls === "EN")  enCnt++;
      if (s.cls === "GNP") gnpCnt++;

      roster[iso].N.push(s.id);
      hw[s.id] += 10;
      shiftCount[s.id]++;
      nightHrsAssigned[s.id] += 10;
      if (nightWeekTarget[s.id]) {
        if (wk === 0) nightWeekTarget[s.id].assignedW1++;
        else nightWeekTarget[s.id].assignedW2++;
      }
    }

    // In-charge coverage enforcement
    if (!nightInChargeCoverage(roster[iso].N) && roster[iso].N.length > 0) {
      const anumOnShift = roster[iso].N.some(id=>activeStaff.find(x=>x.id===id)?.cls==="ANUM");
      const icNeeded = anumOnShift ? 1 : 2;
      const icHave = roster[iso].N.filter(id=>{
        const x=activeStaff.find(y=>y.id===id);
        return x && isInCharge(x) && x.cls!=="ANUM";
      }).length;
      const icShortfall = icNeeded - icHave;

      // Try to add in-charge capable staff up to the shortfall
      for (let i=0; i<icShortfall; i++) {
        if (roster[iso].N.length >= 5) break;
        const ic = eligible.find(s =>
          !roster[iso].N.includes(s.id) &&
          isInCharge(s) &&
          s.cls !== "ANUM" &&
          eligibleForNights(s, iso) &&
          nightHrsAssigned[s.id] + 10 <= nightHoursCeiling[s.id]
        );
        if (ic) {
          roster[iso].N.push(ic.id);
          hw[ic.id] += 10;
          shiftCount[ic.id]++;
          nightHrsAssigned[ic.id] += 10;
          if (nightWeekTarget[ic.id]) {
            if (wk === 0) nightWeekTarget[ic.id].assignedW1++;
            else nightWeekTarget[ic.id].assignedW2++;
          }
        } else {
          nightInChargeMissing.add(iso);
        }
      }
      // Final check
      if (!nightInChargeCoverage(roster[iso].N)) nightInChargeMissing.add(iso);
    }
  });

  // ── ADO accrual at each week boundary ─────────────────────
  days.filter(d=>d.di===6).forEach(sun => {
    const monIso = isoDate(addDays(sun.date,-6));
    activeStaff.filter(s=>s.hrs>=80).forEach(s => {
      let worked = false;
      for (let i=0;i<7;i++) {
        const k=isoDate(addDays(parseLocalDate(monIso),i));
        if (roster[k]?.D.includes(s.id)||roster[k]?.E.includes(s.id)||
            roster[k]?.N.includes(s.id)||leaveMap[s.id]?.[k]) { worked=true; break; }
      }
      if (!worked) return;
      adoAccrued[s.id] += 2;
      const nightCtx = s.permNights ||
        Array.from({length:7},(_,i)=>isoDate(addDays(parseLocalDate(monIso),i)))
          .some(k=>nightPlanData?.plan?.[s.id]?.[k]);
      tryInsertADO(s, sun.iso, nightCtx);
    });
  });

  // ── PHASE 2: Day & Evening — Three-Pass Approach ────────────
  //
  // PASS 1: Priority fill — weekends + Mon/Fri to BASE
  // PASS 2: Pre-allocate mid-week using block preference scoring
  // PASS 3: Refinement — fix isolated shifts by adding adjacent days

  // ── Dynamic ceiling ──────────────────────────────────────
  const deEligibleStaff = activeStaff.filter(s =>
    !s.permNights && (!inNightPlanThisFn.has(s.id) || bumpedFromNights.has(s.id))
  );
  const totalDEshifts = deEligibleStaff.reduce((sum,s) => {
    const leaveHrs = days.filter(d => leaveMap[s.id]?.[d.iso]).length * 8;
    return sum + Math.max(0, Math.floor((periodTarget[s.id] - leaveHrs) / 8));
  }, 0);
  const weekdayCount = days.filter(d=>!d.wknd).length;
  const weekendCount = days.filter(d=>d.wknd).length;
  const totalSlots   = weekdayCount * 20 + weekendCount * 18; // 2 shifts × staffing target per day
  const surplusDE    = Math.max(0, totalDEshifts - totalSlots);
  const surplusPerSlot = weekdayCount > 0 ? surplusDE / (weekdayCount * 2) : 0;
  const WEEKDAY_CEILING = Math.min(14, Math.round(10 + surplusPerSlot));

  // ── Helpers ───────────────────────────────────────────────
  function canWorkForFill(s, iso, shift) {
    if (s.permNights) return false;
    if (inNightPlanThisFn.has(s.id) && !bumpedFromNights.has(s.id)) return false;
    return canWork(s, iso, shift);
  }

  function classOk(s, iso, shift) {
    if (s.cls==="NUM") {
      const di=dayIdx(iso);
      if (shift!=="D"||di<1||di>3||numHasWorked) return false;
    }
    if (s.cls==="ANUM") {
      if (roster[iso][shift].some(id=>activeStaff.find(x=>x.id===id)?.cls==="ANUM")) return false;
      if (roster[iso][shift].some(id=>activeStaff.find(x=>x.id===id)?.cls==="NUM"))  return false;
    }
    if (s.cls==="EN") {
      if (roster[iso][shift].some(id=>activeStaff.find(x=>x.id===id)?.cls==="EN")) return false;
    }
    if (s.cls==="GNP") {
      const gc=roster[iso][shift].filter(id=>activeStaff.find(x=>x.id===id)?.cls==="GNP").length;
      if (gc>=maxGNPShift) return false;
    }
    return true;
  }

  function assignSlot(s, iso, shift, cap, wknd=false) {
    if (!canWorkForFill(s,iso,shift)) return false;
    if (!classOk(s,iso,shift)) return false;
    if (roster[iso][shift].length >= cap) return false;
    if (hw[s.id]+8 > periodTarget[s.id]) return false;
    roster[iso][shift].push(s.id);
    hw[s.id]+=8; shiftCount[s.id]++;
    if (wknd) wkndCnt[s.id]=(wkndCnt[s.id]||0)+1;
    if (s.cls==="NUM") numHasWorked=true;
    return true;
  }

  function ensureIC(iso, shift, wknd=false) {
    const hasIC=roster[iso][shift].some(id=>{const x=activeStaff.find(y=>y.id===id);return x&&isInCharge(x);});
    if (hasIC||roster[iso][shift].length===0) return;
    const ic=activeStaff.find(s=>
      !roster[iso][shift].includes(s.id)&&isInCharge(s)&&
      canWorkForFill(s,iso,shift)&&classOk(s,iso,shift)&&
      hw[s.id]+8<=periodTarget[s.id]
    );
    if (ic) { roster[iso][shift].push(ic.id); hw[ic.id]+=8; shiftCount[ic.id]++; if(wknd)wkndCnt[ic.id]=(wkndCnt[ic.id]||0)+1; }
  }

  function bestCandidates(iso, shift, wknd, cap) {
    return activeStaff
      .filter(s=>canWorkForFill(s,iso,shift)&&classOk(s,iso,shift)&&roster[iso][shift].length<cap&&hw[s.id]+8<=periodTarget[s.id])
      .sort((a,b)=>{
        const aR=!!(a.requests?.[`${iso}_${shift}`]), bR=!!(b.requests?.[`${iso}_${shift}`]);
        if(aR&&!bR)return -1; if(!aR&&bR)return 1;
        if(wknd){const dw=(wkndCnt[a.id]||0)-(wkndCnt[b.id]||0);if(dw)return dw;}
        return (periodTarget[b.id]-hw[b.id])-(periodTarget[a.id]-hw[a.id]);
      });
  }

  function blockScore(sId, iso) {
    const p1=isoDate(addDays(parseLocalDate(iso),-1));
    const n1=isoDate(addDays(parseLocalDate(iso),1));
    const p2=isoDate(addDays(parseLocalDate(iso),-2));
    const n2=isoDate(addDays(parseLocalDate(iso),2));
    const worked = k => roster[k]?.D.includes(sId)||roster[k]?.E.includes(sId)||roster[k]?.N.includes(sId)||prevTail[k]?.D?.includes(sId)||prevTail[k]?.E?.includes(sId);
    let score=0;
    if(worked(p1))score+=3; if(worked(n1))score+=3;
    if(worked(p2))score+=1; if(worked(n2))score+=1;
    return score;
  }

  const priorityDays = [...days.filter(d=>d.wknd), ...days.filter(d=>!d.wknd&&(d.di===0||d.di===4))];
  const midWeekDays  = days.filter(d=>!d.wknd&&d.di!==0&&d.di!==4);

  // ── PASS 1: Priority fill (weekends + Mon/Fri) ────────────
  priorityDays.forEach(({iso, di, wknd}) => {
    const BASE = wknd ? 9 : 10;
    let round=0, stuck=0;
    while ((roster[iso].D.length<BASE||roster[iso].E.length<BASE) && stuck<40) {
      const sh=round%2===0?"D":"E"; round++;
      if (roster[iso][sh].length>=BASE){stuck++;continue;}
      const prev=roster[iso][sh].length;
      const cand=bestCandidates(iso,sh,wknd,BASE)[0];
      if(cand) assignSlot(cand,iso,sh,BASE,wknd);
      if(roster[iso][sh].length===prev) stuck++; else stuck=0;
    }
    ensureIC(iso,"D",wknd); ensureIC(iso,"E",wknd);
  });

  // ── PASS 2: Pre-allocate mid-week with block preference ───
  // Calculate each person's remaining shifts needed
  const midWeekNeed={};
  deEligibleStaff.forEach(s=>{
    midWeekNeed[s.id]=Math.floor(Math.max(0,periodTarget[s.id]-hw[s.id])/8);
  });

  // Sort staff: most constrained first (need / available days ratio)
  const midWeekElig = deEligibleStaff
    .filter(s=>(midWeekNeed[s.id]||0)>0)
    .sort((a,b)=>{
      const avA=midWeekDays.filter(d=>canWorkForFill(a,d.iso,"D")||canWorkForFill(a,d.iso,"E")).length||1;
      const avB=midWeekDays.filter(d=>canWorkForFill(b,d.iso,"D")||canWorkForFill(b,d.iso,"E")).length||1;
      return (midWeekNeed[b.id]/avB)-(midWeekNeed[a.id]/avA);
    });

  midWeekElig.forEach(s=>{
    let needed=midWeekNeed[s.id]||0;
    if(!needed)return;
    // Score all available mid-week day+shift combos
    const opts=[];
    midWeekDays.forEach(({iso})=>{
      ["D","E"].forEach(sh=>{
        if(!canWorkForFill(s,iso,sh))return;
        if(!classOk(s,iso,sh))return;
        if(roster[iso][sh].length>=WEEKDAY_CEILING)return;
        if(hw[s.id]+8>periodTarget[s.id])return;
        // Prefer E slightly to balance with Pass 1 which fills D first on priority days
        const shPref=sh==="E"?0.5:0;
        opts.push({iso,sh,score:blockScore(s.id,iso)+shPref});
      });
    });
    opts.sort((a,b)=>b.score-a.score);
    let assigned=0;
    for(const {iso,sh} of opts){
      if(assigned>=needed)break;
      if(assignSlot(s,iso,sh,WEEKDAY_CEILING,false)) assigned++;
    }
  });

  // Top up mid-week shifts to BASE if under, then ensure IC
  midWeekDays.forEach(({iso})=>{
    ["D","E"].forEach(sh=>{
      let stuck=0;
      while(roster[iso][sh].length<10&&stuck<20){
        const cand=bestCandidates(iso,sh,false,WEEKDAY_CEILING)[0];
        if(!cand){stuck++;break;}
        assignSlot(cand,iso,sh,WEEKDAY_CEILING,false);
        stuck=0;
      }
      ensureIC(iso,sh,false);
    });
  });

  // ── PASS 3: Refinement — fix isolated shifts ──────────────
  for(let p=0;p<3;p++){
    activeStaff.forEach(s=>{
      if(s.permNights||(inNightPlanThisFn.has(s.id)&&!bumpedFromNights.has(s.id)))return;
      days.forEach(({iso,wknd})=>{
        if(!roster[iso].D.includes(s.id)&&!roster[iso].E.includes(s.id))return;
        const p1=isoDate(addDays(parseLocalDate(iso),-1));
        const n1=isoDate(addDays(parseLocalDate(iso),1));
        const worked=k=>roster[k]?.D.includes(s.id)||roster[k]?.E.includes(s.id)||roster[k]?.N.includes(s.id)||leaveMap[s.id]?.[k];
        if(worked(p1)||worked(n1))return; // not isolated
        // Try to add adjacent day to fix isolation
        const sh=roster[iso].D.includes(s.id)?"D":"E";
        for(const fixIso of[p1,n1]){
          if(!roster[fixIso])continue;
          if(hw[s.id]+8>periodTarget[s.id])break;
          const fixWknd=isWknd(fixIso);
          const cap=fixWknd?9:WEEKDAY_CEILING;
          for(const trySh of[sh,sh==="D"?"E":"D"]){
            if(assignSlot(s,fixIso,trySh,cap,fixWknd)) break;
          }
          if(worked(fixIso))break;
        }
      });
    });
  }

  // Merge ADO into leaveMap for display
  activeStaff.forEach(s=>{ Object.keys(adoMap[s.id]).forEach(iso=>{ leaveMap[s.id][iso]="ADO"; }); });

  // Tail for next roster (last 7 days)
  const tail={};
  days.slice(-7).forEach(d=>{ tail[d.iso]={D:[...roster[d.iso].D],E:[...roster[d.iso].E],N:[...roster[d.iso].N]}; });

  // Hours summary
  const hoursSummary={};
  activeStaff.forEach(s=>{
    hoursSummary[s.id]={
      target:periodTarget[s.id],
      worked:hw[s.id],
      adoCount:adoTaken[s.id],
      variance:hw[s.id]-periodTarget[s.id],
      shifts:shiftCount[s.id],
      maxShifts:maxShifts[s.id],
    };
  });

  // Warnings
  const warnings=[];
  days.forEach(({iso,wknd})=>{
    const d=roster[iso];
    ["D","E"].forEach(sh=>{
      const exp=wknd?9:10;
      if(d[sh].length<exp)warnings.push({iso,sh,type:"staffing",msg:`Understaffed: ${d[sh].length}/${exp}`});
      const hasIC=d[sh].some(id=>{const x=activeStaff.find(y=>y.id===id);return x&&isInCharge(x);});
      if(!hasIC&&d[sh].length>0)warnings.push({iso,sh,type:"incharge",msg:"No In-Charge nurse"});
    });
    if(d.N.length<5)warnings.push({iso,sh:"N",type:"staffing",msg:`Night understaffed: ${d.N.length}/5`});
    // Night in-charge coverage check
    if(d.N.length>0){
      if(nightInChargeMissing.has(iso)){
        warnings.push({iso,sh:"N",type:"nightincharge",msg:"Insufficient In-Charge coverage on nights — ANUM present needs 1 additional, no ANUM needs 2"});
      }
    }
  });

  // Night surplus — bumped staff
  bumpedDetails.forEach(b=>{
    warnings.push({
      iso:"—",sh:"N→D/E",type:"nightSurplus",
      msg:`${b.name} moved from Nights to Day/Evening for entire fortnight (surplus night coverage). Historical bumps: ${b.bumpCount}`,
      staffId:b.staffId,
      bumpCount:b.bumpCount,
    });
  });

  // Isolated shift detection
  activeStaff.forEach(s=>{
    days.forEach(({iso})=>{
      const d=roster[iso];
      const workingToday=d.D.includes(s.id)||d.E.includes(s.id)||d.N.includes(s.id);
      if(!workingToday)return;
      const todayShift=d.D.includes(s.id)?"D":d.E.includes(s.id)?"E":"N";
      const prevIso=isoDate(addDays(parseLocalDate(iso),-1));
      const nextIso=isoDate(addDays(parseLocalDate(iso), 1));
      function isOff(checkIso){
        const dr=roster[checkIso];
        if(!dr)return true;
        const onShift=dr.D.includes(s.id)||dr.E.includes(s.id)||dr.N.includes(s.id);
        const onLeave=leaveMap[s.id]?.[checkIso];
        return !onShift||!!onLeave;
      }
      if(isOff(prevIso)&&isOff(nextIso)){
        const wasRequested=Object.keys(s.requests||{}).some(k=>k.startsWith(iso+"_"));
        warnings.push({
          iso,sh:todayShift,type:"isolated",
          msg:`${fullName(s)}: isolated ${todayShift} shift${wasRequested?" (requested by staff)":""}`,
          staffId:s.id, requested:wasRequested,
        });
      }
    });
  });

  // Hours variance warnings
  activeStaff.forEach(s=>{
    const v=hoursSummary[s.id];
    if(Math.abs(v.variance)>8)warnings.push({iso:"—",sh:"Hrs",type:"hours",msg:`${fullName(s)}: ${v.worked}h worked vs ${v.target}h target (${v.variance>0?"+":""}${v.variance}h)`});
  });

  const adoTotal=Object.values(adoMap).reduce((a,m)=>a+Object.keys(m).length,0);

  // Build updated bump history for this roster
  const updatedBumpHistory = { ...(previousRoster?.bumpHistory || {}) };
  bumpedDetails.forEach(b => {
    updatedBumpHistory[b.staffId] = (updatedBumpHistory[b.staffId] || 0) + 1;
  });

  return {
    roster,leaveMap,hoursWorked:hw,hoursSummary,warnings,
    days:days.map(d=>d.iso),startDate:isoDate(startMon),
    tail,wkndCountEnd:{...wkndCnt},
    adoInserted:Object.fromEntries(activeStaff.map(s=>[s.id,Object.keys(adoMap[s.id])])),
    adoTotal,
    bumpHistory: updatedBumpHistory,
    bumpedThisPeriod: bumpedDetails,
    nightInChargeMissing: [...nightInChargeMissing],
  };
}

// ─── PRE-GENERATION VALIDATOR ────────────────────────────────
// Runs before generateRoster() and returns issues grouped by severity.
// BLOCKING issues prevent generation. WARNINGS are shown but allow proceeding.

function validateRosterConfig({ staff, startDate, weeks, nightPlanData }) {
  const errors   = []; // blocking — must fix before generating
  const warnings = []; // non-blocking — should review

  const startMon = getMon(parseLocalDate(startDate));

  // ── Start date checks ──
  const dow = parseLocalDate(startDate).getDay();
  if (dow !== 1) {
    errors.push({
      code: "START_NOT_MONDAY",
      msg:  `Start date is a ${["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][dow]}. Rosters must start on a Monday.`,
      fix:  `Change start date to ${isoDate(startMon)}`,
    });
  }

  // ── Staff checks ──
  const activeStaff = staff.filter(s => {
    if (!s || s.archived) return false;
    if (s.resign) {
      const rosterEnd = addDays(startMon, weeks * 7 - 1);
      if (parseLocalDate(s.resign) <= parseLocalDate(startDate)) return false;
    }
    return true;
  });

  if (activeStaff.length === 0) {
    errors.push({ code:"NO_STAFF", msg:"No active staff members found.", fix:"Add staff in the Staff tab." });
  }

  // Need at least one in-charge per shift type
  const hasInCharge = activeStaff.some(s => isInCharge(s) && !s.permNights);
  if (!hasInCharge) {
    errors.push({ code:"NO_INCHARGE", msg:"No In-Charge capable staff available for Day/Evening shifts.", fix:"Ensure at least one ANUM, CNS, or RN with In-Charge skill is active." });
  }

  const hasNightInCharge = activeStaff.some(s => isInCharge(s));
  if (!hasNightInCharge) {
    errors.push({ code:"NO_NIGHT_INCHARGE", msg:"No In-Charge capable staff available for Night shifts.", fix:"Ensure at least one ANUM, CNS, or RN with In-Charge skill can work nights." });
  }

  // Check each classification requirement
  const clsCounts = {};
  Object.keys(CLASSIFICATIONS).forEach(c => { clsCounts[c] = 0; });
  activeStaff.forEach(s => { clsCounts[s.cls] = (clsCounts[s.cls]||0)+1; });

  if ((clsCounts.ANUM||0) === 0) {
    warnings.push({ code:"NO_ANUM", msg:"No ANUM on roster. Every shift requires an In-Charge nurse — ensure enough CNS/RN In-Charge staff to cover.", fix:"Add an ANUM or ensure RNs have In-Charge skill set." });
  }
  if ((clsCounts.EN||0) === 0) {
    warnings.push({ code:"NO_EN", msg:"No Enrolled Nurses on roster. EN slots on each shift will be empty.", fix:"Add Enrolled Nurse staff if required." });
  }

  // ── Night plan checks ──
  if (!nightPlanData) {
    warnings.push({ code:"NO_NIGHT_PLAN", msg:"No night shift plan has been generated. Night shifts will be understaffed.", fix:"Go to Night Planner tab and click Auto-Generate." });
  } else {
    // Check there are staff assigned to nights in this roster period
    const rosterDays = buildDays(startMon, weeks);
    const nightStaffInPeriod = new Set();
    rosterDays.forEach(d => {
      Object.keys(nightPlanData.plan||{}).forEach(sId => {
        if (nightPlanData.plan[sId]?.[d.iso]) nightStaffInPeriod.add(sId);
      });
    });
    const permNightStaff = activeStaff.filter(s => s.permNights);
    const totalNightAvailable = nightStaffInPeriod.size + permNightStaff.length;

    if (totalNightAvailable < 5) {
      errors.push({
        code: "INSUFFICIENT_NIGHT_STAFF",
        msg:  `Only ${totalNightAvailable} staff available for nights in this period (need 5 per shift minimum).`,
        fix:  "Check night plan covers this fortnight, or add permanent night staff.",
      });
    } else if (totalNightAvailable < 8) {
      warnings.push({
        code: "LOW_NIGHT_STAFF",
        msg:  `Only ${totalNightAvailable} staff available for nights. Some night shifts may be understaffed if leave is taken.`,
        fix:  "Consider reviewing night shift group assignments for this period.",
      });
    }
  }

  // ── Hours viability checks ──
  // Can the ward produce enough staff-hours to fill all required shifts?
  const rosterDays    = buildDays(startMon, weeks);
  const weekendDays   = rosterDays.filter(d => d.wknd);
  const weekdayDays   = rosterDays.filter(d => !d.wknd);
  const requiredHours =
    (weekdayDays.length * 2 * 10 * 8) +  // D+E, 10 staff, 8hrs each
    (weekendDays.length * 2 *  9 * 8) +  // D+E, 9 staff, 8hrs each
    (rosterDays.length  * 1 *  5 * 10);  // N, 5 staff, 10hrs each

  const totalAvailableHours = activeStaff.reduce((sum, s) => {
    // Subtract leave already entered
    const leaveMap = buildLeaveMap([s]);
    let leaveHrs = 0;
    rosterDays.forEach(d => {
      if (leaveMap[s.id]?.[d.iso]) leaveHrs += 8;
    });
    return sum + Math.max(0, s.hrs * (weeks/2) - leaveHrs);
  }, 0);

  const coveragePct = Math.round((totalAvailableHours / requiredHours) * 100);
  if (coveragePct < 70) {
    errors.push({
      code: "INSUFFICIENT_HOURS",
      msg:  `Total available staff hours (${totalAvailableHours}h) covers only ${coveragePct}% of required shift hours (${requiredHours}h). Roster will have significant understaffing.`,
      fix:  "Check leave entries, add staff, or reduce required staffing levels.",
    });
  } else if (coveragePct < 90) {
    warnings.push({
      code: "LOW_COVERAGE",
      msg:  `Staff hours cover ${coveragePct}% of required shifts. Some shifts will likely be understaffed.`,
      fix:  "Review leave in this period or consider calling in agency staff.",
    });
  }

  // ── Individual staff checks ──
  activeStaff.forEach(s => {
    // FWA: no nights + in night group = conflict
    const noNights = s.fwaConditions?.some(c => c.type==="NO_NIGHTS" ||
      (c.type==="SPECIFIC_SHIFTS" && c.shifts && !c.shifts.includes("N")));
    if (!s.permNights && !noNights && !nightPlanData) {
      // Fine — just no plan yet
    }

    // Part-time staff with very low hours — can they complete a useful shift?
    if (s.hrs < 16) {
      warnings.push({
        code: "VERY_LOW_HOURS",
        msg:  `${fullName(s)} is contracted to only ${s.hrs}h/fortnight — fewer than 2 shifts. They may be difficult to roster.`,
        fix:  "Verify contracted hours are correct.",
      });
    }

    // Staff with no commencement date — night restriction can't enforce
    if (s.cls === "GNP" && !s.commencementDate) {
      warnings.push({
        code: "GNP_NO_START",
        msg:  `${fullName(s)} is a GNP with no Commencement Date set. Night shift restrictions (first 3 months) cannot be enforced.`,
        fix:  "Set Commencement Date in the Staff tab.",
      });
    }

    // Resignation during roster period
    if (s.resign) {
      const resignDate = parseLocalDate(s.resign);
      const rosterEnd  = addDays(startMon, weeks*7-1);
      if (resignDate > parseLocalDate(startDate) && resignDate <= rosterEnd) {
        warnings.push({
          code: "RESIGNING_MID_ROSTER",
          msg:  `${fullName(s)} resigns on ${fmtDate(s.resign)}, which falls within this roster period.`,
          fix:  "Check roster around their last day and adjust manually after generating.",
        });
      }
    }
  });

  // ── Weekend coverage check ──
  const wkndEligible = activeStaff.filter(s =>
    !s.permNights &&
    !s.fwaConditions?.some(c => c.type==="NO_WEEKENDS") &&
    !s.fwaConditions?.some(c => c.type==="SPECIFIC_DAYS" && c.days && !c.days.includes(5) && !c.days.includes(6))
  );
  if (wkndEligible.length < 9) {
    errors.push({
      code: "INSUFFICIENT_WEEKEND_STAFF",
      msg:  `Only ${wkndEligible.length} staff are eligible for weekend shifts (need 9 minimum per shift).`,
      fix:  "Review FWA conditions — some staff may be unnecessarily excluded from weekends.",
    });
  } else if (wkndEligible.length < 12) {
    warnings.push({
      code: "LOW_WEEKEND_STAFF",
      msg:  `Only ${wkndEligible.length} staff eligible for weekends. Weekend shifts may be tight if leave is taken.`,
      fix:  "Consider weekend staffing levels when approving leave during this period.",
    });
  }

  return { errors, warnings, canGenerate: errors.length === 0 };
}


export default function App() {
  const [tab,setTab]                     = useState("roster");
  const [staff,setStaffState]            = useState([]);
  const [rosters,setRostersState]        = useState({});
  const [nightPlanData,setNightPlanState]= useState(null);
  const [activeKey,setActiveKey]         = useState(null);
  const [notif,setNotif]                 = useState(null);
  const [loading,setLoading]             = useState(true);
  const [genCfg,setGenCfg]               = useState({startDate:isoDate(getMon(new Date())),weeks:2});

  // ── Initial data load from Supabase ──
  useEffect(()=>{
    async function init(){
      setLoading(true);
      try {
        const [s,r,n,ado]=await Promise.all([
          loadStaff(SAMPLE_STAFF),
          loadRosters({}),
          loadNightPlan(null),
          loadADOAdjustments({}),
        ]);
        setStaffState(s); setRostersState(r); setNightPlanState(n); setAdoAdjustments(ado);
        const keys=Object.keys(r).sort();
        if(keys.length>0)setActiveKey(keys[keys.length-1]);
      }catch(e){ console.error("Init:",e); }
      setLoading(false);
    }
    init();
  },[]);

  // ── Setters that persist to Supabase ──
  const setStaff = useCallback((updater)=>{
    setStaffState(prev=>{
      const next=typeof updater==="function"?updater(prev):updater;
      saveStaff(next);
      return next;
    });
  },[]);

  const setRosters = useCallback((updater)=>{
    setRostersState(prev=>{
      const next=typeof updater==="function"?updater(prev):updater;
      Object.entries(next).forEach(([key,val])=>{
        if(JSON.stringify(val)!==JSON.stringify(prev[key])) saveRoster(key,val);
      });
      return next;
    });
  },[]);

  const setNightPlanData = useCallback((val)=>{
    setNightPlanState(val); if(val)saveNightPlan(val);
  },[]);

  function toast(msg,type="ok"){setNotif({msg,type});setTimeout(()=>setNotif(null),4000);}

  // ── Lock / Unlock ──
  async function handleLockRoster(key){
    const r=rosters[key]; if(!r)return;
    if(r.locked){
      if(!confirm("Unlock this roster? It will become editable again."))return;
      const updated={...r,locked:false,lockedAt:null};
      setRostersState(p=>({...p,[key]:updated}));
      await updateRosterLock(key,false);
      toast("Roster unlocked");
    }else{
      if(!confirm("Lock this roster? All editing will be blocked until you unlock it."))return;
      const updated={...r,locked:true,lockedAt:new Date().toISOString()};
      setRostersState(p=>({...p,[key]:updated}));
      await updateRosterLock(key,true);
      toast("Roster locked ✅");
    }
  }

  // ── Suggest next start date after last locked roster ──
  function suggestNextStartDate(){
    const sortedKeys=Object.keys(rosters).sort();
    const lastLocked=[...sortedKeys].reverse().find(k=>rosters[k]?.locked);
    if(!lastLocked){toast("No locked roster found — set start date manually","err");return;}
    const r=rosters[lastLocked];
    const lastDay=r.days?.[r.days.length-1];
    if(!lastDay)return;
    const nextMon=getMon(addDays(parseLocalDate(lastDay),1));
    setGenCfg(c=>({...c,startDate:isoDate(nextMon)}));
    toast(`Start date set to ${fmtDate(nextMon)} — Monday after the last locked roster`);
  }

  // ── Generate ──
  function handleGenerate(){
    const validation=validateRosterConfig({staff,startDate:genCfg.startDate,weeks:genCfg.weeks,nightPlanData});
    if(!validation.canGenerate){toast(`Cannot generate — ${validation.errors.length} blocking issue(s) must be fixed first`,"err");return;}
    try{
      const key=`${isoDate(getMon(new Date(genCfg.startDate)))}_w${genCfg.weeks}`;
      if(rosters[key]?.locked){toast("That roster is locked. Unlock it first to regenerate.","err");return;}
      const sortedKeys=Object.keys(rosters).sort();
      const prevRoster=sortedKeys.length?rosters[sortedKeys[sortedKeys.length-1]]:null;
      const recentWknd={};
      sortedKeys.slice(-3).forEach(k=>{
        const r=rosters[k]; if(!r?.wkndCountEnd)return;
        Object.entries(r.wkndCountEnd).forEach(([id,cnt])=>{recentWknd[id]=(recentWknd[id]||0)+cnt;});
      });
      const result=generateRoster({
        staff,startDate:genCfg.startDate,weeks:genCfg.weeks,
        nightPlanData,previousRoster:prevRoster,recentWkndCounts:recentWknd,
        bumpHistory: prevRoster?.bumpHistory || {},
      });
      result.generatedAt=new Date().toISOString();
      setRosters(r=>({...r,[key]:result}));
      setActiveKey(key); setTab("roster");
      const wm=validation.warnings.length>0?`, ${validation.warnings.length} pre-gen warning(s)`:"";
      toast(`Roster generated — ${result.warnings.length} issue(s), ${result.adoTotal} ADO(s) inserted${wm}`);
    }catch(e){toast("Error: "+e.message,"err");}
  }

  // ── Export ──
  function handleExport(){
    const r=activeKey&&rosters[activeKey]; if(!r)return toast("No roster","err");
    const wb=XLSX.utils.book_new();
    const days=r.days||[];
    const hdrs=["Staff","Cls","Contract","Target","Worked","Variance","ADOs",...days.map(d=>{const dt=parseLocalDate(d);return `${DAY_NAMES[dayIdx(d)]} ${dt.getDate()}/${dt.getMonth()+1}`;})];
    const rows=staff.filter(s=>!s.archived).map(s=>{
      const hs=r.hoursSummary?.[s.id]||{target:s.hrs,worked:0,variance:0,adoCount:0};
      const row=[fullName(s),s.cls,s.hrs,hs.target,hs.worked,hs.variance,hs.adoCount];
      days.forEach(dk=>{
        const lc=r.leaveMap?.[s.id]?.[dk];
        if(lc)row.push(lc);
        else if(r.roster[dk]?.D.includes(s.id))row.push("D");
        else if(r.roster[dk]?.E.includes(s.id))row.push("E");
        else if(r.roster[dk]?.N.includes(s.id))row.push("N");
        else row.push("");
      });
      return row;
    });
    const ws=XLSX.utils.aoa_to_sheet([hdrs,...rows]);
    ws["!cols"]=hdrs.map((_,i)=>({wch:i<7?14:5}));
    XLSX.utils.book_append_sheet(wb,ws,"Roster");
    if((r.warnings||[]).length>0)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Date","Shift","Type","Warning"],...r.warnings.map(w=>[w.iso,w.sh,w.type,w.msg])]),"Warnings");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Staff","Cls","Contract","Target","Worked","Variance","ADOs"],...staff.filter(s=>!s.archived).map(s=>{const hs=r.hoursSummary?.[s.id]||{};return[fullName(s),s.cls,s.hrs,hs.target||0,hs.worked||0,hs.variance||0,hs.adoCount||0];})]),"Hours Summary");
    XLSX.writeFile(wb,`WardRoster_${activeKey}.xlsx`);
    toast("Exported to Excel");
  }

  const activeRoster=activeKey?rosters[activeKey]:null;
  const [adoAdjustments,setAdoAdjustments] = useState({});

  // Persist ADO adjustments — loaded from Supabase in init()
  // No separate useEffect needed; saves happen per-operation in ADOTab

  const rosterKeys=Object.keys(rosters).sort().reverse();

  // Auto-archive staff whose resignation date has passed
  useEffect(()=>{
    const today = isoDate(new Date());
    const needsArchive = staff.some(s=>s.resign&&!s.archived&&s.resign<today);
    if(!needsArchive)return;
    setStaff(prev=>prev.map(s=>{
      if(s.resign&&!s.archived&&s.resign<today){
        return {...s,archived:true};
      }
      return s;
    }));
  },[staff]);
  async function handleDeleteRoster(key) {
    const r = rosters[key]; if (!r) return;
    if (r.locked) {
      if (!confirm(`⚠️ WARNING: "${key}" is LOCKED.\n\nDeleting a locked roster is permanent and cannot be undone.\n\nAre you absolutely sure?`)) return;
      if (!confirm(`FINAL CONFIRMATION\n\nDelete locked roster "${key}"?\n\nThis permanently removes all shift data for this period and cannot be reversed.\n\nClick OK only if you are certain.`)) return;
    } else {
      if (!confirm(`Delete roster "${key}"?\n\nThis is permanent and cannot be undone.\n\nAre you sure?`)) return;
    }
    setRostersState(prev => { const next={...prev}; delete next[key]; return next; });
    try {
      const { supabase } = await import("./supabase.js");
      if (supabase) await supabase.from("rosters").delete().eq("id", key);
    } catch(e) { console.error("Delete roster:", e); }
    try {
      const local=JSON.parse(localStorage.getItem("wr3_rosters")||"{}");
      delete local[key]; localStorage.setItem("wr3_rosters", JSON.stringify(local));
    } catch {}
    if (activeKey===key) setActiveKey(null);
    toast(`Roster deleted: ${key}`);
  }

  if(loading)return(
    <div style={{minHeight:"100vh",background:"#06101a",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
      <div style={{fontSize:48}}>⚕</div>
      <div style={{fontSize:18,fontWeight:700,color:"#4fc3f7",letterSpacing:2}}>WARDROS TER</div>
      <div style={{fontSize:13,color:"#4a7fa0"}}>Loading ward data…</div>
    </div>
  );

  return(
    <div style={C.app}>
      {notif&&<div style={{...C.notif,background:notif.type==="err"?"#7f0000":"#1a5e20"}}>{notif.msg}</div>}
      <header style={C.header}>
        <div style={C.brand}>
          <span style={{fontSize:26}}>⚕</span>
          <div><div style={C.brandTitle}>WardRoster</div><div style={C.brandSub}>Clinical Scheduling System v3.1</div></div>
        </div>
        <nav style={C.nav}>
          {[["roster","📋 Roster"],["generate","⚡ Generate"],["staff","👥 Staff"],["leave","📅 Leave & Requests"],["nightplan","🌙 Night Planner"],["ado","🗓 ADO Ledger"],["history","🕐 History"]].map(([v,l])=>(
            <button key={v} style={{...C.navBtn,...(tab===v?C.navActive:{})}} onClick={()=>setTab(v)}>{l}</button>
          ))}
        </nav>
        <button style={C.exportBtn} onClick={handleExport}>⬇ Export XLSX</button>
      </header>
      <div style={C.main}>
        {tab==="roster"   &&<RosterTab   roster={activeRoster} staff={staff} rosterKeys={rosterKeys} activeKey={activeKey} setActiveKey={setActiveKey} rosters={rosters} setRosters={setRosters} onLock={handleLockRoster}/>}
        {tab==="generate" &&<GenerateTab cfg={genCfg} setCfg={setGenCfg} onGenerate={handleGenerate} staff={staff} nightPlanData={nightPlanData} rosters={rosters} onSuggestDate={suggestNextStartDate}/>}
        {tab==="staff"    &&<StaffTab    staff={staff} setStaff={setStaff} toast={toast}/>}
        {tab==="leave"    &&<LeaveTab    staff={staff} setStaff={setStaff} toast={toast}/>}
        {tab==="nightplan"&&<NightPlanTab staff={staff} nightPlanData={nightPlanData} setNightPlanData={setNightPlanData} toast={toast} rosters={rosters}/>}
        {tab==="ado"      &&<ADOTab      staff={staff} rosters={rosters} adoAdjustments={adoAdjustments} setAdoAdjustments={setAdoAdjustments} toast={toast}/>}
        {tab==="history"  &&<HistoryTab  rosters={rosters} staff={staff} activeKey={activeKey} setActiveKey={setActiveKey} setTab={setTab} onDelete={handleDeleteRoster}/>}
      </div>
    </div>
  );
}

// ─── ROSTER TAB ──────────────────────────────────────────────
function RosterTab({roster,staff,rosterKeys,activeKey,setActiveKey,rosters,setRosters,onLock}){
  const [filter,setFilter]=useState("");
  const [subTab,setSubTab]=useState("grid");
  const [showAll,setShowAll]=useState(false);
  const CLS_ORDER = {NUM:0,ANUM:1,CNS:2,RN:3,GNP:4,EN:5};
  const visible = staff
    .filter(s=>!s.archived && (s.firstName||s.name||"").toLowerCase().includes(filter.toLowerCase()))
    .sort((a,b)=>{
      const co=(CLS_ORDER[a.cls]??9)-(CLS_ORDER[b.cls]??9);
      if(co!==0)return co;
      const an=`${a.lastName||""} ${a.firstName||a.name||""}`.trim();
      const bn=`${b.lastName||""} ${b.firstName||b.name||""}`.trim();
      return an.localeCompare(bn);
    });
  const isLocked = roster?.locked || false;

  function getCell(sId,iso){
    if(!roster)return null;
    const lc=roster.leaveMap?.[sId]?.[iso]; if(lc)return lc;
    const d=roster.roster[iso];
    if(d?.D.includes(sId))return "D"; if(d?.E.includes(sId))return "E"; if(d?.N.includes(sId))return "N";
    return null;
  }

  function cycleCell(sId,iso){
    if(!roster||!activeKey)return;
    if(isLocked)return; // blocked when locked
    const cur=getCell(sId,iso);
    if(cur&&!["D","E","N",null].includes(cur))return;
    const cycle=[null,"D","E","N"];
    const next=cycle[(cycle.indexOf(cur)+1)%cycle.length];
    const newR=JSON.parse(JSON.stringify(roster));
    newR.roster[iso].D=newR.roster[iso].D.filter(id=>id!==sId);
    newR.roster[iso].E=newR.roster[iso].E.filter(id=>id!==sId);
    newR.roster[iso].N=newR.roster[iso].N.filter(id=>id!==sId);
    if(next)newR.roster[iso][next].push(sId);
    setRosters(r=>({...r,[activeKey]:newR}));
  }

  if(!roster)return(
    <div style={C.empty}>
      <div style={{fontSize:52,marginBottom:12}}>📋</div>
      <div style={C.emptyH}>No Roster Generated</div>
      <div style={C.emptySub}>Use ⚡ Generate to build a new roster, or select a saved one below.</div>
      {rosterKeys.length>0&&<div style={{display:"flex",gap:10,flexWrap:"wrap",justifyContent:"center",marginTop:16}}>
        {rosterKeys.map(k=><button key={k} style={C.pill} onClick={()=>setActiveKey(k)}>📋 {k}</button>)}
      </div>}
    </div>
  );

  const days=roster.days||[];
  const wBT={staffing:0,incharge:0,hours:0,isolated:0,nightSurplus:0,nightincharge:0};
  (roster.warnings||[]).forEach(w=>{ wBT[w.type]=(wBT[w.type]||0)+1; });

  return(
    <div>
      <div style={C.toolbar}>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <select style={C.sel} value={activeKey||""} onChange={e=>setActiveKey(e.target.value)}>
            {rosterKeys.map(k=>{
              const r=rosters[k];
              return <option key={k} value={k}>{k}{r?.locked?" 🔒":""}</option>;
            })}
          </select>
          {/* Lock status badge */}
          {isLocked ? (
            <span style={{background:"#1a0a00",border:"1px solid #f39c12",color:"#f39c12",borderRadius:8,padding:"4px 12px",fontSize:11,fontWeight:700,display:"flex",alignItems:"center",gap:6}}>
              🔒 Locked — {roster.lockedAt ? fmtDate(roster.lockedAt) : ""}
            </span>
          ) : (
            roster && <span style={{background:"#0a1a0a",border:"1px solid #27ae60",color:"#66bb6a",borderRadius:8,padding:"4px 12px",fontSize:11,fontWeight:700}}>
              ✏️ Draft — editable
            </span>
          )}
          {/* Warning badges */}
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {wBT.staffing>0&&<span style={{...C.warnBadge,background:"#4a2000"}}>👥 {wBT.staffing} staffing</span>}
            {wBT.incharge>0&&<span style={{...C.warnBadge,background:"#3a0030"}}>⭐ {wBT.incharge} in-charge</span>}
            {wBT.hours>0&&<span style={{...C.warnBadge,background:"#003030"}}>⏱ {wBT.hours} hours</span>}
            {wBT.isolated>0&&<span style={{...C.warnBadge,background:"#3a2a00"}}>⚠ {wBT.isolated} isolated</span>}
            {wBT.nightSurplus>0&&<span style={{...C.warnBadge,background:"#002040"}}>🔀 {wBT.nightSurplus} night surplus</span>}
            {wBT.nightincharge>0&&<span style={{...C.warnBadge,background:"#3a0020"}}>🌙 {wBT.nightincharge} night in-charge</span>}
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <input style={C.searchBox} placeholder="Search staff…" value={filter} onChange={e=>setFilter(e.target.value)}/>
          {/* Lock / Unlock button */}
          {activeKey && (
            <button
              style={{
                ...C.btnSec,
                borderColor: isLocked ? "#f39c12" : "#27ae60",
                color:       isLocked ? "#f39c12" : "#66bb6a",
                fontWeight:  700,
              }}
              onClick={()=>onLock(activeKey)}
            >
              {isLocked ? "🔓 Unlock" : "🔒 Lock Roster"}
            </button>
          )}
        </div>
      </div>

      {/* Lock banner */}
      {isLocked && (
        <div style={{background:"#1a0f00",border:"1px solid #f39c1244",borderRadius:8,padding:"10px 16px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:20}}>🔒</span>
          <div>
            <div style={{fontSize:12,fontWeight:700,color:"#f39c12"}}>This roster is locked</div>
            <div style={{fontSize:11,color:"#8a6020",marginTop:2}}>Manual edits are disabled. Click <b>Unlock</b> to make changes.</div>
          </div>
        </div>
      )}

      <div style={{display:"flex",gap:4,marginBottom:12,borderBottom:"1px solid #1a3050"}}>
        {[["grid","📋 Roster Grid"],["hours","⏱ Hours Summary"],["warnings",`⚠ Warnings (${(roster.warnings||[]).length})`]].map(([v,l])=>(
          <button key={v} style={{...C.navBtn,...(subTab===v?{...C.navActive,borderBottom:"2px solid #4fc3f7"}:{})}} onClick={()=>setSubTab(v)}>{l}</button>
        ))}
      </div>

      {subTab==="grid"&&(
        <>
          <div style={{overflowX:"auto",border:"1px solid #1a3050",borderRadius:8}}>
            <table style={{borderCollapse:"collapse",tableLayout:"fixed",width:"100%"}}>
              <thead>
                <tr>
                  <th style={{...C.th,...C.fix0,minWidth:175,textAlign:"left",paddingLeft:10}}>Staff</th>
                  <th style={{...C.th,...C.fix1,minWidth:52}}>Cls</th>
                  <th style={{...C.th,...C.fix2,minWidth:64,borderRight:"2px solid #1a4070"}}>Hrs</th>
                  {days.map(iso=>{
                    const dt=parseLocalDate(iso),di=dayIdx(iso),wknd=di>=5;
                    return(
                      <th key={iso} style={{...C.th,minWidth:36,borderLeft:wknd?"2px solid #2a1030":"1px solid #1a3050",background:wknd?"#130a1a":"#080f1a",color:wknd?"#ce93d8":"#4a7fa0"}}>
                        <div style={{fontSize:8}}>{DAY_NAMES[di]}</div>
                        <div style={{fontSize:10}}>{dt.getDate()}/{dt.getMonth()+1}</div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {visible.map(nurse=>{
                  const cls=CLASSIFICATIONS[nurse.cls];
                  const hs=roster.hoursSummary?.[nurse.id]||{target:nurse.hrs,worked:0,variance:0,adoCount:0};
                  const varOk=Math.abs(hs.variance)<=8, varWarn=Math.abs(hs.variance)<=16;
                  const varCol=varOk?"#66bb6a":varWarn?"#ffa726":"#ef5350";
                  return(
                    <tr key={nurse.id}>
                      <td style={{...C.td,...C.fix0,padding:"2px 10px"}}>
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          <span style={{width:7,height:7,borderRadius:"50%",background:cls?.color,flexShrink:0}}/>
                          <span style={{fontSize:11,fontWeight:600,color:"#c8d8e8",whiteSpace:"nowrap"}}>{nurse.name}</span>
                          {nurse.permNights&&<span style={C.bdg("#1a237e","#7986cb")}>PN</span>}
                          {nurse.fwaConditions?.length>0&&<span style={C.bdg("#1b4000","#81c784")}>FWA</span>}
                        </div>
                      </td>
                      <td style={{...C.td,...C.fix1,textAlign:"center"}}>
                        <span style={{fontSize:9,fontWeight:700,color:cls?.color,background:cls?.color+"22",borderRadius:8,padding:"1px 5px"}}>{nurse.cls}</span>
                      </td>
                      <td style={{...C.td,...C.fix2,textAlign:"center",borderRight:"2px solid #1a4070",padding:"2px 4px"}}>
                        <div style={{fontSize:10,fontWeight:700,color:varCol}}>{hs.worked}</div>
                        <div style={{fontSize:8,color:"#2a5070"}}>/{hs.target}h</div>
                        {hs.adoCount>0&&<div style={{fontSize:7,color:"#a5d6a7"}}>ADO×{hs.adoCount}</div>}
                      </td>
                      {days.map(iso=>{
                        const code=getCell(nurse.id,iso);
                        const def=SHIFT_DEF[code]||null;
                        const wknd=dayIdx(iso)>=5;
                        const autoADO=code==="ADO"&&roster.adoInserted?.[nurse.id]?.includes(iso);
                        const isolated=["D","E","N"].includes(code)&&
                          (roster.warnings||[]).some(w=>
                            w.type==="isolated"&&w.iso===iso&&w.staffId===nurse.id
                          );
                        const isolatedRequested=isolated&&
                          (roster.warnings||[]).some(w=>
                            w.type==="isolated"&&w.iso===iso&&w.staffId===nurse.id&&w.requested
                          );
                        // Night in-charge missing — highlight entire N column for that date
                        const nightICMissing=code==="N"&&
                          (roster.nightInChargeMissing||[]).includes(iso);
                        return(
                          <td key={iso}
                            style={{
                              ...C.td,
                              background:def?def.bg:wknd?"#0d0814":"transparent",
                              textAlign:"center",padding:1,
                              cursor:CARD_CODES.includes(code)||isLocked?"default":"pointer",
                              borderLeft:wknd?"2px solid #2a1030":"1px solid #111a28",
                              outline:nightICMissing?"2px solid #ef5350":isolated?`2px solid ${isolatedRequested?"#64b5f6":"#ffa726"}`:"none",
                              outlineOffset:"-2px",
                              position:"relative",
                            }}
                            onClick={()=>!CARD_CODES.includes(code)&&!isLocked&&cycleCell(nurse.id,iso)}
                            title={
                              nightICMissing
                                ? `🌙 Insufficient In-Charge coverage on this night shift`
                                : isolated
                                  ? `⚠ Isolated shift — ${fullName(nurse)} works alone this day${isolatedRequested?" (requested)":""}\n${code}: ${def?.label||""}`
                                  : code?(def?.label+(autoADO?" ★ auto-inserted":"")):"Off — click to add"
                            }>
                            <span style={{fontSize:9,fontWeight:700,color:def?def.text:"#1a3050",display:"block",lineHeight:"22px"}}>
                              {code||""}{autoADO&&<span style={{fontSize:6,verticalAlign:"super"}}>★</span>}
                              {isolated&&<span style={{fontSize:6,verticalAlign:"super",color:isolatedRequested?"#64b5f6":"#ffa726"}}>!</span>}
                            </span>
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:10}}>
            {Object.entries(SHIFT_DEF).map(([k,v])=>(
              <span key={k} style={{fontSize:9,padding:"2px 8px",borderRadius:10,background:v.bg,color:v.text,border:`1px solid ${v.color}44`}}>
                <b>{k}</b> {v.label}
              </span>
            ))}
            <span style={{fontSize:9,padding:"2px 8px",borderRadius:10,background:"#0a1a0a",color:"#a5d6a7",border:"1px solid #43a04744"}}><b>ADO★</b> Auto-inserted</span>
            <span style={{fontSize:9,padding:"2px 8px",borderRadius:10,background:"#1a1200",color:"#ffa726",border:"1px solid #ffa72644"}}><b>⚠!</b> Isolated shift</span>
            <span style={{fontSize:9,padding:"2px 8px",borderRadius:10,background:"#0a1422",color:"#64b5f6",border:"1px solid #64b5f644"}}><b>!</b> Isolated (requested)</span>
            <span style={{fontSize:9,padding:"2px 8px",borderRadius:10,background:"#220a0a",color:"#ef5350",border:"1px solid #ef535044"}}><b>🌙</b> Night in-charge missing</span>
          </div>
          <h3 style={{...C.sectionH,marginTop:20}}>Daily Staffing Counts</h3>
          <div style={{overflowX:"auto",border:"1px solid #1a3050",borderRadius:8,marginTop:6}}>
            <table style={{borderCollapse:"collapse",tableLayout:"fixed",width:"100%"}}>
              <thead>
                <tr>
                  <th style={{...C.th,minWidth:70,textAlign:"left",paddingLeft:10}}>Shift</th>
                  {days.map(iso=>{const dt=parseLocalDate(iso),di=dayIdx(iso),wknd=di>=5;
                    return<th key={iso} style={{...C.th,minWidth:36,fontSize:9,background:wknd?"#130a1a":"#080f1a",color:wknd?"#ce93d8":"#4a7fa0"}}>{DAY_NAMES[di]}<br/>{dt.getDate()}</th>;})}
                </tr>
              </thead>
              <tbody>
                {["D","E","N"].map(sh=>{
                  const def=SHIFT_DEF[sh];
                  return(
                    <tr key={sh}>
                      <td style={{...C.td,paddingLeft:10,fontWeight:700,color:def.color,fontSize:11}}>{def.label}</td>
                      {days.map(iso=>{
                        const cnt=roster.roster[iso]?.[sh]?.length||0;
                        const exp=sh==="N"?5:dayIdx(iso)>=5?9:10;
                        return<td key={iso} style={{...C.td,textAlign:"center",fontWeight:700,fontSize:11,color:cnt>=exp?"#66bb6a":"#ef5350"}}>{cnt}</td>;
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}

      {subTab==="hours"&&(
        <div>
          <h3 style={C.sectionH}>Hours Summary — {activeKey}</h3>
          <div style={{fontSize:11,color:"#4a7fa0",marginBottom:12}}>🟢 Within 8h of target &nbsp; 🟡 8–16h variance &nbsp; 🔴 Over 16h variance</div>
          <div style={{border:"1px solid #1a3050",borderRadius:8,overflow:"hidden"}}>
            <table style={{borderCollapse:"collapse",width:"100%"}}>
              <thead>
                <tr style={{background:"#060e18"}}>
                  {["Staff","Class","Contract","Target","Worked","Variance","Shifts","ADOs"].map(h=>(
                    <th key={h} style={{...C.th,textAlign:"left",padding:"8px 12px",fontSize:11}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {staff.filter(s=>!s.archived).map((s,i)=>{
                  const hs=roster.hoursSummary?.[s.id]||{target:s.hrs,worked:0,variance:0,adoCount:0,shifts:0,maxShifts:0};
                  const cls=CLASSIFICATIONS[s.cls];
                  const ok=Math.abs(hs.variance)<=8,warn=Math.abs(hs.variance)<=16;
                  const col=ok?"#66bb6a":warn?"#ffa726":"#ef5350";
                  return(
                    <tr key={s.id} style={{background:i%2===0?"#0a1828":"#07101e",borderTop:"1px solid #0d1e30"}}>
                      <td style={{...C.td,padding:"8px 12px",fontWeight:600,color:"#c8d8e8"}}>{fullName(s)}</td>
                      <td style={{...C.td,padding:"8px 12px"}}><span style={{fontSize:10,color:cls?.color,background:cls?.color+"22",borderRadius:6,padding:"1px 7px",fontWeight:700}}>{s.cls}</span></td>
                      <td style={{...C.td,padding:"8px 12px",color:"#7fb3d3"}}>{s.hrs}h/fn</td>
                      <td style={{...C.td,padding:"8px 12px",color:"#a8dadc",fontWeight:700}}>{hs.target}h</td>
                      <td style={{...C.td,padding:"8px 12px",color:col,fontWeight:700}}>{hs.worked}h</td>
                      <td style={{...C.td,padding:"8px 12px"}}>
                        <span style={{color:col,fontWeight:700}}>{hs.variance>0?"+":""}{hs.variance}h</span>
                        {!ok&&<span style={{fontSize:9,marginLeft:4}}>{warn?"⚠":"✗"}</span>}
                      </td>
                      <td style={{...C.td,padding:"8px 12px",color:"#7fb3d3"}}>{hs.shifts}/{hs.maxShifts}</td>
                      <td style={{...C.td,padding:"8px 12px",color:"#a5d6a7"}}>{hs.adoCount>0?`${hs.adoCount} ADO${hs.adoCount>1?"s":""}`:""}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {subTab==="warnings"&&(
        <div>
          <h3 style={C.sectionH}>Scheduling Warnings — {(roster.warnings||[]).length} total</h3>
          {(roster.warnings||[]).length===0&&<div style={{color:"#2a5070",padding:16}}>✅ No warnings — roster looks good.</div>}
          {["staffing","incharge","nightincharge","nightSurplus","hours","isolated"].map(type=>{
            const items=(roster.warnings||[]).filter(w=>w.type===type);
            if(!items.length)return null;
            const meta={
              staffing:     {title:"Staffing Levels",              icon:"👥",color:"#ffa726"},
              incharge:     {title:"In-Charge Coverage (D/E)",     icon:"⭐",color:"#ce93d8"},
              nightincharge:{title:"Night In-Charge Coverage",     icon:"🌙",color:"#ef9a9a"},
              nightSurplus: {title:"Night Surplus — Moved to D/E", icon:"🔀",color:"#64b5f6"},
              hours:        {title:"Hours Variance",               icon:"⏱",color:"#80deea"},
              isolated:     {title:"Isolated Shifts",              icon:"⚠",color:"#ffa726"},
            };
            const m=meta[type];
            return(
              <div key={type} style={{...C.card,marginBottom:14}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                  <span style={{fontSize:16}}>{m.icon}</span>
                  <span style={{fontWeight:700,color:m.color,fontSize:13}}>{m.title}</span>
                  <span style={{fontSize:10,color:"#4a7fa0"}}>{items.length} issue{items.length!==1?"s":""}</span>
                </div>
                {(showAll?items:items.slice(0,12)).map((w,i)=>(
                  <div key={i} style={{display:"flex",gap:10,background:"#060d18",borderRadius:6,padding:"6px 12px",alignItems:"center",marginBottom:4}}>
                    <span style={{fontSize:10,color:"#4a7fa0",minWidth:90,flexShrink:0}}>{w.iso}</span>
                    <span style={{fontSize:10,color:m.color,fontWeight:700,minWidth:30}}>{w.sh}</span>
                    <span style={{fontSize:11,color:"#c8d8e8"}}>{w.msg}</span>
                  </div>
                ))}
                {!showAll&&items.length>12&&<button style={{...C.btnSec,fontSize:10,marginTop:6}} onClick={()=>setShowAll(true)}>Show all {items.length} →</button>}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── GENERATE TAB ────────────────────────────────────────────
function GenerateTab({cfg,setCfg,onGenerate,staff,nightPlanData,rosters,onSuggestDate}){
  const {numGroups,staffPerBlock,permHoursPerFn}=autoComputeNightGroups(staff);
  const fwaCount   = staff.filter(s=>s.fwaConditions?.length>0).length;
  const permNights = staff.filter(s=>s.permNights);
  const sortedKeys = Object.keys(rosters).sort();
  const prevRoster = sortedKeys.length?rosters[sortedKeys[sortedKeys.length-1]]:null;
  const hasTail    = !!(prevRoster?.tail&&Object.keys(prevRoster.tail).length>0);
  const prevEnd    = prevRoster?.days?.[prevRoster.days.length-1];
  const recentCount= sortedKeys.slice(-3).length;

  // Run validation live as config changes
  const validation = validateRosterConfig({ staff, startDate:cfg.startDate, weeks:cfg.weeks, nightPlanData });
  const { errors, warnings, canGenerate } = validation;

  return(
    <div>
      <h2 style={C.pageH}>⚡ Generate Roster</h2>

      {/* Config cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(255px,1fr))",gap:16,marginBottom:20}}>
        <div style={C.card}>
          <div style={C.cardH}>Roster Period</div>
          <label style={C.lbl}>Start Date (must be a Monday)</label>
          <input type="date" style={C.inp} value={cfg.startDate} onChange={e=>setCfg(c=>({...c,startDate:e.target.value}))}/>
          <button style={{...C.btnSec,fontSize:10,marginBottom:10}} onClick={onSuggestDate}>
            📅 Use next date after locked roster
          </button>
          <div style={{display:"flex",gap:8}}>
            {[[2,"1 Fortnight"],[4,"4 Weeks"]].map(([w,l])=>(
              <button key={w} style={{...C.selBtn,...(cfg.weeks===w?C.selBtnOn:{})}} onClick={()=>setCfg(c=>({...c,weeks:w}))}>{l}</button>
            ))}
          </div>
        </div>

        <div style={C.card}>
          <div style={C.cardH}>Ward Overview</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {[["Total Staff",staff.length],["Perm Nights",permNights.length],["FWA Staff",fwaCount],["Night Groups",numGroups]].map(([l,v])=>(
              <div key={l} style={{background:"#060d18",borderRadius:7,padding:10,textAlign:"center"}}>
                <div style={{fontSize:22,fontWeight:700,color:"#64b5f6"}}>{v}</div>
                <div style={{fontSize:9,color:"#4a7fa0"}}>{l}</div>
              </div>
            ))}
          </div>
        </div>

        <div style={C.card}>
          <div style={C.cardH}>Continuity & History</div>
          <div style={{fontSize:11,lineHeight:2.1,color:"#7fb3d3"}}>
            <div>Previous roster: <b style={{color:prevRoster?"#a8dadc":"#4a7fa0"}}>{prevRoster?fmtDate(prevEnd):"None saved"}</b></div>
            <div>Tail carry-forward: <b style={{color:hasTail?"#66bb6a":"#ffa726"}}>{hasTail?"✅ Active":"⚠ None (first roster)"}</b></div>
            <div>Weekend history: <b style={{color:"#a8dadc"}}>{recentCount} roster{recentCount!==1?"s":""} loaded</b></div>
            <div>Night plan: <b style={{color:nightPlanData?"#66bb6a":"#ef5350"}}>{nightPlanData?"✅ Loaded":"⚠ Not generated"}</b></div>
            <div>ADO accrual: <b style={{color:"#a5d6a7"}}>Auto-insert (FT staff)</b></div>
          </div>
        </div>
      </div>

      {/* ── VALIDATION PANEL ── */}
      <div style={{marginBottom:24}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}>
          <h3 style={{...C.sectionH,marginBottom:0}}>
            {canGenerate ? "✅ Pre-Generation Check" : "❌ Pre-Generation Check"}
          </h3>
          <span style={{
            fontSize:11,fontWeight:700,padding:"3px 12px",borderRadius:10,
            background: canGenerate ? "#1b5e20" : "#7f0000",
            color: canGenerate ? "#a5d6a7" : "#ef9a9a",
          }}>
            {canGenerate
              ? errors.length===0&&warnings.length===0 ? "All clear" : `${warnings.length} warning(s)`
              : `${errors.length} blocking issue(s)`}
          </span>
        </div>

        {/* Blocking errors */}
        {errors.length > 0 && (
          <div style={{marginBottom:12}}>
            <div style={{fontSize:11,fontWeight:700,color:"#ef5350",marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>
              🚫 Must Fix Before Generating
            </div>
            {errors.map((e,i) => (
              <div key={i} style={{
                background:"#1a0505",border:"1px solid #7f000088",
                borderRadius:8,padding:"12px 16px",marginBottom:8,
              }}>
                <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
                  <span style={{fontSize:16,flexShrink:0}}>🚫</span>
                  <div style={{flex:1}}>
                    <div style={{fontSize:12,fontWeight:600,color:"#ef9a9a",marginBottom:4}}>{e.msg}</div>
                    <div style={{fontSize:11,color:"#7f4040"}}>
                      <span style={{color:"#ef535066"}}>How to fix: </span>
                      <span style={{color:"#cd8080"}}>{e.fix}</span>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Non-blocking warnings */}
        {warnings.length > 0 && (
          <div style={{marginBottom:12}}>
            <div style={{fontSize:11,fontWeight:700,color:"#ffa726",marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>
              ⚠ Warnings — Generation Allowed
            </div>
            {warnings.map((w,i) => (
              <div key={i} style={{
                background:"#1a1000",border:"1px solid #ffa72644",
                borderRadius:8,padding:"12px 16px",marginBottom:8,
              }}>
                <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
                  <span style={{fontSize:16,flexShrink:0}}>⚠</span>
                  <div style={{flex:1}}>
                    <div style={{fontSize:12,fontWeight:600,color:"#ffd54f",marginBottom:4}}>{w.msg}</div>
                    <div style={{fontSize:11}}>
                      <span style={{color:"#ffa72666"}}>Suggestion: </span>
                      <span style={{color:"#ffcc80"}}>{w.fix}</span>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* All clear */}
        {errors.length === 0 && warnings.length === 0 && (
          <div style={{background:"#0a1a0a",border:"1px solid #2e7d3244",borderRadius:8,padding:"14px 18px",display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:22}}>✅</span>
            <div>
              <div style={{fontSize:13,fontWeight:700,color:"#81c784"}}>All checks passed</div>
              <div style={{fontSize:11,color:"#4a7fa0",marginTop:2}}>Configuration looks good — ready to generate.</div>
            </div>
          </div>
        )}
      </div>

      {/* Generate button */}
      <button
        style={{
          ...C.btnPrimary,
          fontSize:15,padding:"13px 44px",
          opacity: canGenerate ? 1 : 0.45,
          cursor:  canGenerate ? "pointer" : "not-allowed",
          background: canGenerate
            ? "linear-gradient(135deg,#0d47a1,#1565c0)"
            : "#1a1a1a",
        }}
        onClick={onGenerate}
        title={canGenerate ? "Generate roster" : "Fix blocking issues above before generating"}
      >
        {canGenerate ? "⚡ Generate Roster" : "🚫 Fix Issues First"}
      </button>

      {!canGenerate && (
        <div style={{fontSize:11,color:"#4a5070",marginTop:10}}>
          Resolve the {errors.length} blocking issue{errors.length!==1?"s":""} above to enable generation.
        </div>
      )}

      {/* Rules summary — collapsed by default */}
      <details style={{marginTop:24}}>
        <summary style={{fontSize:11,color:"#4a7fa0",cursor:"pointer",userSelect:"none",marginBottom:8}}>
          📋 View active scheduling rules
        </summary>
        <div style={{...C.card,marginTop:8}}>
          <div style={{display:"flex",flexDirection:"column",gap:4}}>
            {[
              "✅ Contracted hours targeted per staff (hard ceiling)",
              "✅ Night hrs rounded down to nearest 10h shift multiple",
              "✅ Leave hours deducted from target before scheduling",
              "✅ ADO auto-inserted at threshold (FT staff, resets each period)",
              "✅ Forward rotation preferred (D→E→N)",
              "✅ Max 5 consecutive D/E shifts, max 4 consecutive nights",
              "✅ 47h break enforced after night shifts",
              "✅ Max 1 EN per shift, max 1 ANUM per shift",
              "✅ ≤50% GNP per D/E shift, max 1 GNP per night",
              "✅ In-Charge nurse guaranteed on every shift",
              "✅ NUM: 1×Day shift per fortnight (Tue–Thu), no ANUM clash",
              "✅ All FWA conditions enforced",
              "✅ GNP: no night shifts in first 3 months",
              "✅ Weekend fairness across last 3 rosters",
              "✅ Priority fill: Weekends & Mon/Fri before mid-week",
              "✅ No overstaffing on weekends or nights",
              "✅ Cross-fortnight tail continuity (consecutive shifts, 47h break)",
            ].map(r=><div key={r} style={{fontSize:10,color:"#7fb3d3"}}>{r}</div>)}
          </div>
        </div>
      </details>
    </div>
  );
}

// ─── STAFF TAB ───────────────────────────────────────────────
function StaffTab({staff,setStaff,toast}){
  const [editing,setEditing]=useState(null);
  const [form,setForm]=useState(null);
  const [filterCls,setFilterCls]=useState("ALL");
  const importRef=useRef();
  const blank={
    id:`s${Date.now()}`,firstName:"",lastName:"",cls:"RN",hrs:"",
    permNights:false,inCharge:false,fwaConditions:[],prefs:"",
    resign:null,commencementDate:null,
    leaveCard:[],gridLeave:{},requests:{},archived:false,
  };

  function startEdit(s){setForm({...s,fwaConditions:[...(s.fwaConditions||[])]});setEditing(s.id);}
  function startNew(){setForm({...blank,id:`s${Date.now()}`});setEditing("new");}

  function saveForm(){
    const errors=[];
    if(!form.firstName?.trim())errors.push("First Name is required");
    if(!form.lastName?.trim())errors.push("Last Name is required");
    if(!form.hrs||isNaN(Number(form.hrs))||Number(form.hrs)<=0)errors.push("Contracted Hours is required");
    if(!form.commencementDate)errors.push("Commencement Date is required");
    if(errors.length>0){toast(errors[0],"err");return;}
    const saved={...form,hrs:Number(form.hrs)};
    if(editing==="new")setStaff(p=>[...p,saved]); else setStaff(p=>p.map(s=>s.id===form.id?saved:s));
    setEditing(null);toast("Saved");
  }
  function remove(id){if(!confirm("Remove this staff member?"))return;setStaff(p=>p.filter(s=>s.id!==id));toast("Removed");}

  function handleImport(e){
    const file=e.target.files[0]; if(!file)return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        const imported=rows.map((r,i)=>({...blank,id:r.ID||`imp${i}`,name:r.Name||"",cls:r.Classification||"RN",hrs:Number(r.ContractedHours)||80,
          permNights:r.PermanentNights==="Yes",inCharge:r.InChargeSkill==="Yes",prefs:r.Preferences||"",
          resign:r.ResignationDate||null,commencementDate:r.CommencementDate||r.GNP_StartDate||null,
          fwaConditions:r.FWA==="Yes"?[{type:"CUSTOM",note:r.FWA_Notes||"See agreement"}]:[],
        }));
        setStaff(p=>{const ids=new Set(p.map(s=>s.id));return[...p,...imported.filter(s=>!ids.has(s.id))];});
        toast(`Imported ${imported.length} staff`);
      }catch(err){toast("Import failed: "+err.message,"err");}
    };
    reader.readAsBinaryString(file); e.target.value="";
  }

  function exportTemplate(){
    const wb=XLSX.utils.book_new();
    const ws=XLSX.utils.aoa_to_sheet([
      ["ID","Name","Classification","ContractedHours","PermanentNights","InChargeSkill","FWA","FWA_Notes","Preferences","ResignationDate","GNP_StartDate"],
      ["s1","Jane Smith","RN","80","No","Yes","No","","Prefers mornings","",""],
      ["s2","John Doe","EN","64","No","No","Yes","No nights","","",""],
    ]);
    XLSX.utils.book_append_sheet(wb,ws,"Staff");
    XLSX.writeFile(wb,"StaffImportTemplate.xlsx");
    toast("Template downloaded");
  }

  if(editing){
    const f=form,sf=fn=>setForm(p=>({...p,...fn(p)}));
    return(
      <div style={{maxWidth:740}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:20}}>
          <button style={C.backBtn} onClick={()=>setEditing(null)}>← Back</button>
          <h2 style={C.pageH}>{editing==="new"?"Add Staff Member":"Edit: "+f.name}</h2>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
          <div>
            <label style={C.lbl}>First Name <span style={{color:"#ef5350"}}>*</span></label>
            <input type="text" style={{...C.inp,borderColor:!f.firstName?.trim()&&editing?"#ef535066":undefined}}
              value={f.firstName||""} onChange={e=>sf(p=>({...p,firstName:e.target.value}))}/>
          </div>
          <div>
            <label style={C.lbl}>Last Name <span style={{color:"#ef5350"}}>*</span></label>
            <input type="text" style={C.inp} value={f.lastName||""} onChange={e=>sf(p=>({...p,lastName:e.target.value}))}/>
          </div>
          <div>
            <label style={C.lbl}>Contracted Hours / Fortnight <span style={{color:"#ef5350"}}>*</span></label>
            <input type="number" style={C.inp} placeholder="e.g. 80" value={f.hrs||""} onChange={e=>sf(p=>({...p,hrs:e.target.value}))}/>
          </div>
          <div>
            <label style={C.lbl}>Commencement Date <span style={{color:"#ef5350"}}>*</span></label>
            <input type="date" style={C.inp} value={f.commencementDate||""} onChange={e=>sf(p=>({...p,commencementDate:e.target.value||null}))}/>
            <div style={{fontSize:9,color:"#4a7fa0",marginTop:-8}}>No night shifts for first 3 months from this date</div>
          </div>
          <div>
            <label style={C.lbl}>Resignation Date</label>
            <div style={{display:"flex",gap:6}}>
              <input type="date" style={{...C.inp,flex:1}} value={f.resign||""} onChange={e=>sf(p=>({...p,resign:e.target.value||null}))}/>
              {f.resign&&(
                <button type="button" style={{...C.btnSec,fontSize:10,padding:"0 10px"}} onClick={()=>sf(p=>({...p,resign:null}))}>✕ Clear</button>
              )}
            </div>
            <div style={{fontSize:9,color:f.resign?"#ef9a9a":"#4a7fa0",marginTop:2}}>
              {f.resign ? `Last working day: ${fmtDate(f.resign)} — staff auto-archived after this date` : "Leave blank unless this staff member has resigned"}
            </div>
          </div>
          <div>
            <label style={C.lbl}>Classification</label>
            <select style={C.inp} value={f.cls} onChange={e=>sf(p=>({...p,cls:e.target.value}))}>
              {Object.entries(CLASSIFICATIONS).map(([k,v])=><option key={k} value={k}>{k} — {v.label}</option>)}
            </select>
          </div>
          <div style={{gridColumn:"1/-1"}}>
            <label style={C.lbl}>Known Preferences</label>
            <input type="text" style={C.inp} value={f.prefs||""} onChange={e=>sf(p=>({...p,prefs:e.target.value}))}/>
          </div>
        </div>
        <div style={{display:"flex",gap:20,marginBottom:16,flexWrap:"wrap"}}>
          {[["permNights","Permanent Night Shift"],["inCharge","In-Charge Skill (RN)"]].map(([k,l])=>(
            <label key={k} style={{display:"flex",alignItems:"center",gap:8,fontSize:12,color:"#7fb3d3",cursor:"pointer"}}>
              <input type="checkbox" checked={!!f[k]} onChange={e=>sf(p=>({...p,[k]:e.target.checked}))}/>{l}
            </label>
          ))}
        </div>
        <div style={C.card}>
          <div style={C.cardH}>Flexible Work Arrangement Conditions</div>
          <div style={{fontSize:11,color:"#4a7fa0",marginBottom:10}}>Each condition is a specific approved restriction on rostering.</div>
          {(f.fwaConditions||[]).map((cond,i)=>(
            <div key={i} style={{background:"#060d18",borderRadius:7,padding:10,marginBottom:8,display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-start"}}>
              <div style={{minWidth:200,flex:1}}>
                <label style={C.lbl}>Condition</label>
                <select style={C.inp} value={cond.type} onChange={e=>{const nc=[...f.fwaConditions];nc[i]={...nc[i],type:e.target.value};sf(p=>({...p,fwaConditions:nc}));}}>
                  {Object.entries(FWA_CONDITIONS).map(([k,v])=><option key={k} value={k}>{v}</option>)}
                </select>
              </div>
              {cond.type==="SPECIFIC_DAYS"&&(
                <div style={{flex:2,minWidth:220}}>
                  <label style={C.lbl}>Allowed Days</label>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    {DAY_NAMES.map((dn,di)=>(
                      <label key={di} style={{display:"flex",alignItems:"center",gap:3,fontSize:11,color:"#7fb3d3",cursor:"pointer"}}>
                        <input type="checkbox" checked={(cond.days||[]).includes(di)} onChange={e=>{
                          const nc=[...f.fwaConditions],days=new Set(nc[i].days||[]);
                          e.target.checked?days.add(di):days.delete(di);nc[i]={...nc[i],days:[...days].sort()};sf(p=>({...p,fwaConditions:nc}));
                        }}/>{dn}
                      </label>
                    ))}
                  </div>
                </div>
              )}
              {cond.type==="SPECIFIC_SHIFTS"&&(
                <div style={{flex:1,minWidth:160}}>
                  <label style={C.lbl}>Allowed Shifts</label>
                  <div style={{display:"flex",gap:10}}>
                    {["D","E","N"].map(sh=>(
                      <label key={sh} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#7fb3d3",cursor:"pointer"}}>
                        <input type="checkbox" checked={(cond.shifts||[]).includes(sh)} onChange={e=>{
                          const nc=[...f.fwaConditions],shifts=new Set(nc[i].shifts||[]);
                          e.target.checked?shifts.add(sh):shifts.delete(sh);nc[i]={...nc[i],shifts:[...shifts]};sf(p=>({...p,fwaConditions:nc}));
                        }}/>{sh}
                      </label>
                    ))}
                  </div>
                </div>
              )}
              {cond.type==="MAX_HOURS_WEEK"&&(
                <div style={{flex:1}}>
                  <label style={C.lbl}>Max Hours/Week</label>
                  <input type="number" style={C.inp} value={cond.value||40} onChange={e=>{const nc=[...f.fwaConditions];nc[i]={...nc[i],value:Number(e.target.value)};sf(p=>({...p,fwaConditions:nc}));}}/>
                </div>
              )}
              <div style={{flex:2,minWidth:200}}>
                <label style={C.lbl}>Agreement Note</label>
                <input type="text" style={C.inp} value={cond.note||""} placeholder="e.g. Approved by NUM…" onChange={e=>{const nc=[...f.fwaConditions];nc[i]={...nc[i],note:e.target.value};sf(p=>({...p,fwaConditions:nc}));}}/>
              </div>
              <button style={{...C.iconBtn,color:"#ef5350",marginTop:20}} onClick={()=>sf(p=>({...p,fwaConditions:p.fwaConditions.filter((_,j)=>j!==i)}))}>✕</button>
            </div>
          ))}
          <button style={C.btnSec} onClick={()=>sf(p=>({...p,fwaConditions:[...(p.fwaConditions||[]),{type:"NO_NIGHTS",note:""}]}))}>+ Add Condition</button>
        </div>
        <div style={{display:"flex",gap:12,marginTop:20}}>
          <button style={C.btnPrimary} onClick={saveForm}>💾 Save</button>
          <button style={C.btnSec} onClick={()=>setEditing(null)}>Cancel</button>
        </div>
      </div>
    );
  }

  const CLS_ORDER={NUM:0,ANUM:1,CNS:2,RN:3,GNP:4,EN:5};
  const filtered=staff
    .filter(s=>(filterCls==="ALL"||s.cls===filterCls)&&
      fullName(s).toLowerCase().includes(""))
    .sort((a,b)=>{
      // Active before archived
      if(a.archived&&!b.archived)return 1;
      if(!a.archived&&b.archived)return -1;
      const co=(CLS_ORDER[a.cls]??9)-(CLS_ORDER[b.cls]??9);
      if(co!==0)return co;
      return fullName(a).localeCompare(fullName(b));
    });
  return(
    <div>
      <div style={C.toolbar}>
        <h2 style={C.pageH}>👥 Staff ({staff.length})</h2>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <select style={C.sel} value={filterCls} onChange={e=>setFilterCls(e.target.value)}>
            <option value="ALL">All</option>{Object.keys(CLASSIFICATIONS).map(k=><option key={k} value={k}>{k}</option>)}
          </select>
          <button style={C.btnSec} onClick={exportTemplate}>⬇ Template</button>
          <button style={C.btnSec} onClick={()=>importRef.current.click()}>⬆ Import XLSX</button>
          <input ref={importRef} type="file" accept=".xlsx" style={{display:"none"}} onChange={handleImport}/>
          <button style={C.btnPrimary} onClick={startNew}>+ Add Staff</button>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(290px,1fr))",gap:12}}>
        {filtered.map(nurse=>{
          const cls=CLASSIFICATIONS[nurse.cls];
          return(
            <div key={nurse.id} style={{...C.card,opacity:nurse.archived?0.5:1}}>
              <div style={{display:"flex",alignItems:"flex-start",gap:10,marginBottom:8}}>
                <div style={{width:38,height:38,borderRadius:"50%",background:cls?.color+"33",color:cls?.color,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:14,flexShrink:0}}>
                  {`${(nurse.firstName||"")[0]||""}${(nurse.lastName||"")[0]||""}`.toUpperCase()||"?"}
                </div>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,fontWeight:700,color:"#c8d8e8"}}>{fullName(nurse)}</div>
                  <div style={{display:"flex",gap:5,flexWrap:"wrap",marginTop:3}}>
                    <span style={{fontSize:9,color:cls?.color,background:cls?.color+"22",borderRadius:8,padding:"1px 6px",fontWeight:700}}>{nurse.cls}</span>
                    <span style={{fontSize:9,color:"#64b5f6",background:"#1a3050",borderRadius:8,padding:"1px 6px"}}>{nurse.hrs}h/fn</span>
                    {nurse.permNights&&<span style={C.bdg("#1a237e","#7986cb")}>Perm Nights</span>}
                    {nurse.inCharge&&<span style={C.bdg("#1a3000","#aed581")}>In-Charge</span>}
                    {nurse.fwaConditions?.length>0&&<span style={C.bdg("#1b4000","#81c784")}>FWA ×{nurse.fwaConditions.length}</span>}
                    {nurse.archived&&<span style={C.bdg("#2a0000","#ef9a9a")}>Archived</span>}
                  </div>
                </div>
                <div style={{display:"flex",gap:4}}>
                  <button style={C.iconBtn} onClick={()=>startEdit(nurse)}>✏️</button>
                  <button style={C.iconBtn} onClick={()=>remove(nurse.id)}>🗑️</button>
                </div>
              </div>
              {nurse.fwaConditions?.length>0&&(
                <div style={{background:"#050f00",border:"1px solid #2e7d3244",borderRadius:6,padding:"6px 10px",marginBottom:6}}>
                  {nurse.fwaConditions.map((c,i)=><div key={i} style={{fontSize:10,color:"#81c784"}}>📋 {FWA_CONDITIONS[c.type]}{c.note?`: ${c.note}`:""}</div>)}
                </div>
              )}
              {nurse.commencementDate&&<div style={{fontSize:10,color:"#4a7fa0"}}>🗓 Started: {fmtDate(nurse.commencementDate)}</div>}
              {nurse.prefs&&<div style={{fontSize:10,color:"#4a7fa0",fontStyle:"italic"}}>💭 {nurse.prefs}</div>}
              {nurse.resign&&<div style={{fontSize:10,color:"#ef5350",marginTop:3}}>🚪 Resigning {fmtDate(nurse.resign)}</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── LEAVE TAB ───────────────────────────────────────────────
const PAINT_OPTIONS=[
  {code:"AL", label:"Annual Leave", color:"#9c27b0",isLeave:true},
  {code:"SL", label:"Sick Leave",   color:"#e74c3c",isLeave:true},
  {code:"PDL",label:"Prof. Dev.",   color:"#2196f3",isLeave:true},
  {code:"UL", label:"Union Leave",  color:"#78909c",isLeave:true},
  {code:"ADO",label:"ADO",          color:"#43a047",isLeave:true},
  {code:"RQ", label:"Shift Request",color:"#ffa726",isLeave:false},
];

function LeaveTab({staff,setStaff,toast}){
  const [mode,setMode]=useState("grid");
  const [weekStart,setWeekStart]=useState(()=>getMon(new Date()));
  const [paintCode,setPaintCode]=useState("AL");
  const [reqShift,setReqShift]=useState("D");
  const [filterName,setFilterName]=useState("");
  const [filterCls,setFilterCls]=useState("ALL");
  const [cardForm,setCardForm]=useState({staffId:"",code:"LSL",from:"",to:"",note:""});

  const weekDays=Array.from({length:7},(_,i)=>{const date=addDays(weekStart,i);return{date,iso:isoDate(date),di:dayIdx(date),wknd:isWknd(date)};});
  const weekLabel=()=>`${fmtShort(weekDays[0].date)} – ${fmtShort(weekDays[6].date)} ${weekDays[6].date.getFullYear()}`;
  const visible=staff.filter(s=>!s.archived&&fullName(s).toLowerCase().includes(filterName.toLowerCase())&&(filterCls==="ALL"||s.cls===filterCls));

  function getLongLeave(sId,iso){return(staff.find(x=>x.id===sId)?.leaveCard||[]).find(l=>iso>=l.from&&iso<=l.to)||null;}
  function getLeaveCode(sId,iso){return staff.find(x=>x.id===sId)?.gridLeave?.[iso]||null;}
  function getRequest(sId,iso){const s=staff.find(x=>x.id===sId);const k=Object.keys(s?.requests||{}).find(k=>k.startsWith(iso+"_"));return k?k.split("_")[1]:null;}

  function handleCell(sId,iso){
    setStaff(prev=>prev.map(s=>{
      if(s.id!==sId)return s;
      if((s.leaveCard||[]).find(l=>iso>=l.from&&iso<=l.to))return s;
      if(paintCode==="RQ"){
        const newReqs={...(s.requests||{})};
        const existing=Object.keys(newReqs).find(k=>k.startsWith(iso+"_"));
        if(existing){const sh=existing.split("_")[1];delete newReqs[existing];if(sh!==reqShift)newReqs[`${iso}_${reqShift}`]=true;}
        else{newReqs[`${iso}_${reqShift}`]=true;}
        return{...s,requests:newReqs};
      }else{
        const newGL={...(s.gridLeave||{})},newReqs={...(s.requests||{})};
        Object.keys(newReqs).filter(k=>k.startsWith(iso+"_")).forEach(k=>delete newReqs[k]);
        if(newGL[iso]===paintCode)delete newGL[iso]; else newGL[iso]=paintCode;
        return{...s,gridLeave:newGL,requests:newReqs};
      }
    }));
  }

  function clearRow(sId){
    const isos=new Set(weekDays.map(d=>d.iso));
    setStaff(prev=>prev.map(s=>{
      if(s.id!==sId)return s;
      const gl={...(s.gridLeave||{})},reqs={...(s.requests||{})};
      isos.forEach(iso=>{delete gl[iso];Object.keys(reqs).filter(k=>k.startsWith(iso+"_")).forEach(k=>delete reqs[k]);});
      return{...s,gridLeave:gl,requests:reqs};
    }));
  }

  function clearWeek(){
    if(!confirm(`Clear all leave & requests for visible staff: ${weekLabel()}?`))return;
    const isos=new Set(weekDays.map(d=>d.iso));
    setStaff(prev=>prev.map(s=>{
      if(!visible.find(x=>x.id===s.id))return s;
      const gl={...(s.gridLeave||{})},reqs={...(s.requests||{})};
      isos.forEach(iso=>{delete gl[iso];Object.keys(reqs).filter(k=>k.startsWith(iso+"_")).forEach(k=>delete reqs[k]);});
      return{...s,gridLeave:gl,requests:reqs};
    }));
    toast("Week cleared");
  }

  function addCardLeave(){
    if(!cardForm.staffId||!cardForm.from||!cardForm.to)return toast("Fill all fields","err");
    if(cardForm.from>cardForm.to)return toast("From must be before To","err");
    setStaff(prev=>prev.map(s=>s.id!==cardForm.staffId?s:{...s,leaveCard:[...(s.leaveCard||[]),{...cardForm,id:Date.now()}]}));
    toast("Long leave added");setCardForm(f=>({...f,from:"",to:"",note:""}));
  }

  function removeCard(sId,lid){setStaff(prev=>prev.map(s=>s.id!==sId?s:{...s,leaveCard:(s.leaveCard||[]).filter(l=>l.id!==lid)}));toast("Removed");}

  return(
    <div>
      <div style={C.toolbar}>
        <h2 style={C.pageH}>📅 Leave & Requests</h2>
        <div style={{display:"flex",gap:8}}>
          <button style={{...C.selBtn,...(mode==="grid"?C.selBtnOn:{})}} onClick={()=>setMode("grid")}>📊 Weekly Grid</button>
          <button style={{...C.selBtn,...(mode==="card"?C.selBtnOn:{})}} onClick={()=>setMode("card")}>📋 Long Leave (LSL/PL)</button>
        </div>
      </div>

      {mode==="grid"&&(
        <div>
          <div style={{display:"flex",alignItems:"center",width:"fit-content",background:"#0a1828",borderRadius:9,border:"1px solid #1a3050",overflow:"hidden",marginBottom:14}}>
            <button style={C.weekNavBtn} onClick={()=>setWeekStart(d=>addDays(d,-7))}>‹ Prev</button>
            <div style={{padding:"10px 24px",borderLeft:"1px solid #1a3050",borderRight:"1px solid #1a3050",textAlign:"center",minWidth:230}}>
              <div style={{fontSize:15,fontWeight:700,color:"#4fc3f7"}}>{weekLabel()}</div>
              <div style={{fontSize:10,color:"#2a6080",marginTop:2}}>{weekDays.filter(d=>!d.wknd).length} weekdays · {weekDays.filter(d=>d.wknd).length} weekend days</div>
            </div>
            <button style={C.weekNavBtn} onClick={()=>setWeekStart(d=>addDays(d,7))}>Next ›</button>
            <button style={{...C.weekNavBtn,color:"#ffa726",borderLeft:"1px solid #1a3050"}} onClick={()=>setWeekStart(getMon(new Date()))}>Today</button>
          </div>

          <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap",marginBottom:12,padding:"12px 14px",background:"#0a1525",border:"1px solid #1a3050",borderRadius:8}}>
            <div>
              <label style={C.lbl}>Filter</label>
              <div style={{display:"flex",gap:6}}>
                <input style={{...C.inp,marginBottom:0,width:140}} placeholder="Name…" value={filterName} onChange={e=>setFilterName(e.target.value)}/>
                <select style={{...C.inp,marginBottom:0,width:100}} value={filterCls} onChange={e=>setFilterCls(e.target.value)}>
                  <option value="ALL">All</option>{Object.keys(CLASSIFICATIONS).map(k=><option key={k} value={k}>{k}</option>)}
                </select>
              </div>
            </div>
            <div style={{borderLeft:"1px solid #1a3050",paddingLeft:14}}>
              <label style={C.lbl}>Paint code — click cells to apply</label>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {PAINT_OPTIONS.map(o=>(
                  <button key={o.code} onClick={()=>setPaintCode(o.code)} style={{fontSize:11,fontFamily:"inherit",fontWeight:700,padding:"6px 13px",borderRadius:7,cursor:"pointer",
                    background:paintCode===o.code?o.color+"33":"#060e18",color:o.color,
                    border:`1px solid ${paintCode===o.code?o.color:o.color+"44"}`,
                    boxShadow:paintCode===o.code?`0 0 8px ${o.color}55`:"none"}}>
                    {o.code} {o.label}
                  </button>
                ))}
              </div>
            </div>
            {paintCode==="RQ"&&(
              <div>
                <label style={C.lbl}>Request shift</label>
                <div style={{display:"flex",gap:6}}>
                  {["D","E","N"].map(sh=>(
                    <button key={sh} onClick={()=>setReqShift(sh)} style={{fontSize:12,fontFamily:"inherit",fontWeight:700,padding:"6px 14px",borderRadius:7,cursor:"pointer",
                      background:reqShift===sh?SHIFT_DEF[sh].color+"33":"#060e18",color:SHIFT_DEF[sh].color,
                      border:`1px solid ${reqShift===sh?SHIFT_DEF[sh].color:SHIFT_DEF[sh].color+"44"}`}}>{sh} {SHIFT_DEF[sh].label}</button>
                  ))}
                </div>
              </div>
            )}
            <button style={{...C.btnSec,fontSize:10,marginLeft:"auto"}} onClick={clearWeek}>🗑 Clear Week</button>
          </div>

          <div style={{fontSize:10,color:"#2a6080",marginBottom:8,display:"flex",gap:20,flexWrap:"wrap"}}>
            <span>🔒 <b style={{color:"#ef9a9a"}}>Leave codes</b> — mandatory, staff will not be rostered on these days</span>
            <span>📌 <b style={{color:"#ffa726"}}>Shift Requests</b> — best-effort, honoured where rules allow</span>
          </div>

          <div style={{border:"1px solid #1a3050",borderRadius:9,overflow:"hidden"}}>
            <div style={{display:"grid",gridTemplateColumns:"220px 55px repeat(7,1fr) 32px",background:"#060e18",borderBottom:"2px solid #1a3050"}}>
              <div style={{padding:"10px 14px",fontSize:10,fontWeight:700,color:"#2a6080",textTransform:"uppercase",letterSpacing:1}}>Staff</div>
              <div style={{padding:"10px 4px",fontSize:10,fontWeight:700,color:"#2a6080",textAlign:"center"}}>Hrs</div>
              {weekDays.map(d=>(
                <div key={d.iso} style={{padding:"8px 4px",textAlign:"center",borderLeft:d.wknd?"2px solid #3a1050":"1px solid #1a3050",background:d.wknd?"#130820":"transparent"}}>
                  <div style={{fontSize:14,fontWeight:700,color:d.wknd?"#ce93d8":"#4fc3f7"}}>{DAY_NAMES[d.di]}</div>
                  <div style={{fontSize:11,color:d.wknd?"#9c27b0":"#2a6080",marginTop:1}}>{fmtShort(d.date)}</div>
                </div>
              ))}
              <div style={{padding:"10px 4px",fontSize:10,color:"#2a6080",textAlign:"center",borderLeft:"1px solid #1a3050"}}>✕</div>
            </div>
            {visible.length===0&&<div style={{padding:"32px",textAlign:"center",color:"#2a5070",fontSize:12}}>No staff match filter.</div>}
            {visible.map((nurse,ri)=>{
              const cls=CLASSIFICATIONS[nurse.cls];
              const leaveDays=weekDays.filter(d=>nurse.gridLeave?.[d.iso]||(nurse.leaveCard||[]).some(l=>d.iso>=l.from&&d.iso<=l.to)).length;
              const reqDays=weekDays.filter(d=>Object.keys(nurse.requests||{}).some(k=>k.startsWith(d.iso+"_"))).length;
              return(
                <div key={nurse.id} style={{display:"grid",gridTemplateColumns:"220px 55px repeat(7,1fr) 32px",background:ri%2===0?"#0a1525":"#07101e",borderTop:"1px solid #0d1e30"}}>
                  <div style={{padding:"8px 14px",display:"flex",alignItems:"center",gap:7}}>
                    <span style={{width:8,height:8,borderRadius:"50%",background:cls?.color,flexShrink:0}}/>
                    <div>
                      <div style={{fontSize:12,fontWeight:600,color:"#c8d8e8",whiteSpace:"nowrap",overflow:"hidden",maxWidth:145}}>{nurse.name}</div>
                      <div style={{display:"flex",gap:4,marginTop:2}}>
                        <span style={{fontSize:9,color:cls?.color,background:cls?.color+"22",borderRadius:5,padding:"0 4px",fontWeight:700}}>{nurse.cls}</span>
                        {leaveDays>0&&<span style={{fontSize:9,color:"#ce93d8"}}>🏖{leaveDays}d</span>}
                        {reqDays>0&&<span style={{fontSize:9,color:"#ffa726"}}>📌{reqDays}d</span>}
                      </div>
                    </div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#4fc3f7",borderLeft:"1px solid #0d1e30"}}>{nurse.hrs}h</div>
                  {weekDays.map(d=>{
                    const ll=getLongLeave(nurse.id,d.iso);
                    const lc=ll?null:getLeaveCode(nurse.id,d.iso);
                    const rq=ll||lc?null:getRequest(nurse.id,d.iso);
                    const def=ll?SHIFT_DEF[ll.code]:lc?SHIFT_DEF[lc]:null;
                    return(
                      <div key={d.iso} onClick={()=>!ll&&handleCell(nurse.id,d.iso)}
                        style={{borderLeft:d.wknd?"2px solid #3a1050":"1px solid #0d1e30",
                          background:ll?def.bg:lc?def.bg:rq?"#1a1000":d.wknd?"#0e0618":"transparent",
                          cursor:ll?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",minHeight:54}}>
                        {ll&&<div style={{textAlign:"center"}}><div style={{fontSize:13,fontWeight:700,color:def.text}}>{ll.code}</div><div style={{fontSize:8,color:def.text+"88"}}>Long leave</div></div>}
                        {lc&&<div style={{textAlign:"center"}}><div style={{fontSize:14,fontWeight:700,color:def.text}}>{lc}</div><div style={{fontSize:8,color:def.text+"88"}}>{SHIFT_DEF[lc]?.label}</div></div>}
                        {rq&&<div style={{textAlign:"center"}}><div style={{fontSize:14,fontWeight:700,color:"#ffa726"}}>{rq}</div><div style={{fontSize:8,color:"#ffa72688"}}>Request 📌</div></div>}
                        {!ll&&!lc&&!rq&&<div style={{width:24,height:24,borderRadius:"50%",border:"1px dashed #1a3050",display:"flex",alignItems:"center",justifyContent:"center",opacity:.3,fontSize:12,color:"#4fc3f7"}}>+</div>}
                      </div>
                    );
                  })}
                  <div style={{display:"flex",alignItems:"center",justifyContent:"center",borderLeft:"1px solid #0d1e30"}}>
                    <button style={{background:"none",border:"none",color:"#2a5070",cursor:"pointer",fontSize:14}} onClick={()=>clearRow(nurse.id)}>✕</button>
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:12}}>
            {PAINT_OPTIONS.map(o=>(
              <span key={o.code} style={{fontSize:10,padding:"3px 10px",borderRadius:10,background:SHIFT_DEF[o.code]?.bg||"#1a1000",color:SHIFT_DEF[o.code]?.text||"#ffa726",border:`1px solid ${o.color}44`}}>
                <b>{o.code}</b> {o.label} {o.isLeave?"🔒":"📌"}
              </span>
            ))}
            {CARD_CODES.map(k=>(
              <span key={k} style={{fontSize:10,padding:"3px 10px",borderRadius:10,background:SHIFT_DEF[k].bg,color:SHIFT_DEF[k].text,border:`1px solid ${SHIFT_DEF[k].color}44`}}>
                <b>{k}</b> {SHIFT_DEF[k].label} 🔒
              </span>
            ))}
          </div>
        </div>
      )}

      {mode==="card"&&(
        <div>
          <div style={C.card}>
            <div style={C.cardH}>Add Long-Term Leave (LSL / Parental Leave)</div>
            <div style={{fontSize:11,color:"#4a7fa0",marginBottom:14,lineHeight:1.7}}>
              For <b style={{color:"#80deea"}}>Long Service Leave</b> and <b style={{color:"#f48fb1"}}>Parental Leave</b> — full weeks or months. Always mandatory and blocks rostering.
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
              <div style={{gridColumn:"1/-1"}}>
                <label style={C.lbl}>Staff Member</label>
                <select style={C.inp} value={cardForm.staffId} onChange={e=>setCardForm(f=>({...f,staffId:e.target.value}))}>
                  <option value="">— Select —</option>
                  {staff.filter(s=>!s.archived).sort((a,b)=>fullName(a).localeCompare(fullName(b))).map(s=><option key={s.id} value={s.id}>{fullName(s)} ({s.cls})</option>)}
                </select>
              </div>
              <div><label style={C.lbl}>Leave Type</label>
                <select style={C.inp} value={cardForm.code} onChange={e=>setCardForm(f=>({...f,code:e.target.value}))}>
                  {CARD_CODES.map(k=><option key={k} value={k}>{k} — {SHIFT_DEF[k]?.label}</option>)}
                </select>
              </div>
              <div><label style={C.lbl}>From</label><input type="date" style={C.inp} value={cardForm.from} onChange={e=>setCardForm(f=>({...f,from:e.target.value}))}/></div>
              <div><label style={C.lbl}>To</label><input type="date" style={C.inp} value={cardForm.to} onChange={e=>setCardForm(f=>({...f,to:e.target.value}))}/></div>
              <div style={{gridColumn:"1/-1"}}><label style={C.lbl}>Note (optional)</label>
                <input type="text" style={C.inp} placeholder="e.g. Maternity leave approved 12 months" value={cardForm.note} onChange={e=>setCardForm(f=>({...f,note:e.target.value}))}/>
              </div>
            </div>
            <button style={C.btnPrimary} onClick={addCardLeave}>+ Add Leave</button>
          </div>
          <h3 style={{...C.sectionH,marginTop:24}}>All Long-Term Leave Records</h3>
          {staff.filter(s=>(s.leaveCard||[]).length>0).sort((a,b)=>a.name.localeCompare(b.name)).map(s=>{
            const cls=CLASSIFICATIONS[s.cls];
            return(
              <div key={s.id} style={{...C.card,marginBottom:10}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                  <span style={{width:9,height:9,borderRadius:"50%",background:cls?.color}}/>
                  <span style={{fontWeight:700,color:"#c8d8e8",fontSize:13}}>{fullName(s)}</span>
                  <span style={{fontSize:9,color:cls?.color,background:cls?.color+"22",borderRadius:6,padding:"1px 6px",fontWeight:700}}>{s.cls}</span>
                </div>
                {(s.leaveCard||[]).sort((a,b)=>a.from>b.from?1:-1).map(le=>{
                  const def=SHIFT_DEF[le.code];
                  const daysCount=Math.round((new Date(le.to)-new Date(le.from))/86400000)+1;
                  return(
                    <div key={le.id} style={{display:"flex",alignItems:"center",gap:10,background:"#060d18",border:`1px solid ${def?.color}33`,borderRadius:7,padding:"10px 14px",marginBottom:6,flexWrap:"wrap"}}>
                      <span style={{fontSize:11,fontWeight:700,color:def?.text,background:def?.bg,borderRadius:6,padding:"3px 10px"}}>{le.code}</span>
                      <span style={{fontSize:12,color:"#a8dadc",fontWeight:600}}>{fmtDate(le.from)}</span>
                      <span style={{color:"#2a6080"}}>→</span>
                      <span style={{fontSize:12,color:"#a8dadc",fontWeight:600}}>{fmtDate(le.to)}</span>
                      <span style={{fontSize:10,color:"#4a7fa0"}}>{daysCount}d ({Math.round(daysCount/7)}w)</span>
                      {le.note&&<span style={{fontSize:10,color:"#4a7fa0",fontStyle:"italic"}}>"{le.note}"</span>}
                      <button style={{marginLeft:"auto",background:"none",border:"none",color:"#ef5350",cursor:"pointer",fontSize:14}} onClick={()=>removeCard(s.id,le.id)}>✕</button>
                    </div>
                  );
                })}
              </div>
            );
          })}
          {staff.every(s=>!(s.leaveCard||[]).length)&&<div style={{color:"#2a5070",fontSize:12,padding:16,textAlign:"center"}}>No long-term leave entries yet.</div>}
        </div>
      )}
    </div>
  );
}

// ─── NIGHT PLAN TAB ──────────────────────────────────────────
function NightPlanTab({staff,nightPlanData,setNightPlanData,toast,rosters}){
  const [year,setYear]             = useState(new Date().getFullYear());
  const [view,setView]             = useState("planner");
  const [monthOffset,setMonthOffset] = useState(new Date().getMonth());
  const [firstMonday,setFirstMonday] = useState(
    // Default: first Monday of current year, or stored from previous plan
    ()=> nightPlanData?.firstMonday || isoDate(getMon(new Date(new Date().getFullYear(),0,1)))
  );
  const {groups,numGroups,staffPerBlock,permHoursPerFn,targetHours,baseRequired,bufferPct}=autoComputeNightGroups(staff);
  const permStaff=staff.filter(s=>s.permNights);
  const gc=["#e53935","#fb8c00","#fdd835","#43a047","#1e88e5","#8e24aa","#00acc1","#f06292","#00897b","#6d4c41"];

  // Validate firstMonday is actually a Monday
  const firstMondayValid = firstMonday && parseLocalDate(firstMonday).getDay() === 1;

  function autoGen(){
    if(!firstMondayValid){
      toast("First Monday must be a Monday — please check the date","err"); return;
    }
    const result=autoAssignNightPlan(staff, year, firstMonday);
    setNightPlanData(result);
    toast(`Night plan generated for ${year}: ${result.groups.length} groups, starting ${fmtDate(firstMonday)}`);
  }

  // Keep firstMonday in sync with stored plan
  useEffect(()=>{
    if(nightPlanData?.firstMonday) setFirstMonday(nightPlanData.firstMonday);
  },[nightPlanData]);

  // Build fortnight list for selected year
  const yearFns = (nightPlanData?.fns||[]).filter(fn=>new Date(fn.start).getFullYear()===year);

  return(
    <div>
      <div style={C.toolbar}>
        <h2 style={C.pageH}>🌙 Night Shift Planner</h2>
        <div style={{display:"flex",gap:8}}>
          <button style={{...C.selBtn,...(view==="planner"?C.selBtnOn:{})}} onClick={()=>setView("planner")}>⚙️ Planner</button>
          <button style={{...C.selBtn,...(view==="timeline"?C.selBtnOn:{})}} onClick={()=>setView("timeline")}>📅 Annual Timeline</button>
        </div>
      </div>

      {/* ── PLANNER VIEW ── */}
      {view==="planner"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:14,marginBottom:24}}>
            <div style={C.card}>
              <div style={C.cardH}>Coverage Calculation</div>
              <div style={{fontSize:11,color:"#7fb3d3",lineHeight:2.1}}>
                <div>Required: <b style={{color:"#a8dadc"}}>5 staff × 10h = 50h/shift</b></div>
                <div>14 nights/block: <b style={{color:"#a8dadc"}}>{baseRequired}h base target</b></div>
                <div>+{bufferPct}% buffer: <b style={{color:"#66bb6a"}}>{targetHours}h target</b></div>
                <div>Perm nights (adj.): <b style={{color:"#a8dadc"}}>{permHoursPerFn}h/fn</b></div>
                <div>Rotating needed: <b style={{color:"#a8dadc"}}>{Math.max(0,targetHours-permHoursPerFn)}h/block</b></div>
                <div>Staff per block: <b style={{color:"#66bb6a"}}>{staffPerBlock}</b></div>
                <div>Groups: <b style={{color:"#66bb6a"}}>{numGroups}</b></div>
                <div style={{fontSize:9,color:"#2a6080",marginTop:4}}>
                  Hours rounded down to nearest 10h (night shift length)
                </div>
              </div>
            </div>
            <div style={C.card}>
              <div style={C.cardH}>Permanent Nights</div>
              {permStaff.length===0&&<div style={{fontSize:11,color:"#2a5070"}}>None configured</div>}
              {permStaff.map(s=>{const cls=CLASSIFICATIONS[s.cls];return(
                <div key={s.id} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                  <span style={{width:7,height:7,borderRadius:"50%",background:cls?.color}}/>
                  <span style={{fontSize:11,color:"#c8d8e8"}}>{fullName(s)}</span>
                  <span style={{fontSize:9,color:cls?.color}}>{s.cls}</span>
                  <span style={{fontSize:9,color:"#64b5f6",marginLeft:"auto"}}>{s.hrs}h</span>
                </div>
              );})}
            </div>
            <div style={C.card}>
              <div style={C.cardH}>Auto-Generate Plan</div>
              <div style={{fontSize:11,color:"#4a7fa0",marginBottom:10,lineHeight:1.7}}>
                Groups staff by classification and hours, ensures In-Charge coverage, rotates with 6–8 week gaps.
              </div>

              <label style={C.lbl}>First Monday of your roster year</label>
              <input type="date" style={{...C.inp,borderColor:firstMondayValid?"#1a3050":"#ef5350"}}
                value={firstMonday||""}
                onChange={e=>setFirstMonday(e.target.value)}
              />
              {firstMonday&&!firstMondayValid&&(
                <div style={{fontSize:10,color:"#ef5350",marginBottom:8,marginTop:-6}}>
                  ⚠ This date is not a Monday — please select a Monday
                </div>
              )}
              {firstMondayValid&&(
                <div style={{fontSize:10,color:"#66bb6a",marginBottom:8,marginTop:-6}}>
                  ✅ Fortnights will start from {fmtDate(firstMonday)}
                </div>
              )}

              <label style={C.lbl}>Year</label>
              <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:14}}>
                <button style={C.btnSec} onClick={()=>setYear(y=>y-1)}>‹</button>
                <strong style={{color:"#a8dadc",fontSize:18,minWidth:48,textAlign:"center"}}>{year}</strong>
                <button style={C.btnSec} onClick={()=>setYear(y=>y+1)}>›</button>
              </div>
              <button style={{...C.btnPrimary,opacity:firstMondayValid?1:0.4}} onClick={autoGen}>
                🔄 Auto-Generate {year}
              </button>
              {nightPlanData&&(
                <div style={{fontSize:10,color:"#66bb6a",marginTop:8}}>
                  ✅ Plan active — {nightPlanData.groups?.length} groups
                  {nightPlanData.firstMonday&&` · from ${fmtDate(nightPlanData.firstMonday)}`}
                </div>
              )}
            </div>
          </div>

          {groups.length>0&&(
            <>
              <h3 style={C.sectionH}>Night Shift Groups</h3>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(210px,1fr))",gap:12,marginBottom:24}}>
                {groups.map(g=>{
                  const color=gc[g.id-1]||"#64b5f6";
                  const members=staff.filter(s=>g.members.includes(s.id));
                  const hasIC=members.some(s=>isInCharge(s));
                  return(
                    <div key={g.id} style={{...C.card,borderColor:color+"66"}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                        <div style={{width:11,height:11,borderRadius:"50%",background:color}}/>
                        <span style={{fontWeight:700,color,fontSize:13}}>Group {g.id}</span>
                        <span style={{fontSize:9,color:"#4a7fa0"}}>{members.length} staff · {g.totalHours}h</span>
                      </div>
                      <div style={{fontSize:9,color:hasIC?"#66bb6a":"#ef5350",marginBottom:8}}>{hasIC?"✅ In-Charge covered":"⚠ No In-Charge!"}</div>
                      {members.map(s=>{const cls=CLASSIFICATIONS[s.cls];return(
                        <div key={s.id} style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}>
                          <span style={{fontSize:9,color:cls?.color,background:cls?.color+"22",borderRadius:5,padding:"0 4px",fontWeight:700}}>{s.cls}</span>
                          <span style={{fontSize:10,color:"#a8dadc"}}>{fullName(s)}</span>
                          <span style={{fontSize:9,color:"#4a7fa0",marginLeft:"auto"}}>{s.hrs}h</span>
                        </div>
                      );})}
                    </div>
                  );
                })}
              </div>
            </>
          )}

          {nightPlanData&&(
            <>
              <h3 style={C.sectionH}>Fortnight Assignment — {year}</h3>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(158px,1fr))",gap:8}}>
                {yearFns.map(fn=>{
                  const gid=nightPlanData.groupAssignments?.[fn.key];
                  const color=gid?gc[gid-1]||"#64b5f6":"#1a3050";
                  const group=nightPlanData.groups?.find(g=>g.id===gid);
                  return(
                    <div key={fn.key} style={{background:"#0a1525",border:`1px solid ${color}55`,borderRadius:7,padding:"8px 10px"}}>
                      <div style={{fontSize:9,color:"#4a7fa0"}}>Fn {fn.idx+1}</div>
                      <div style={{fontSize:10,color:"#7fb3d3",marginBottom:5}}>{fmtShort(fn.start)} – {fmtShort(fn.end)}</div>
                      {gid?(
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          <div style={{width:8,height:8,borderRadius:"50%",background:color}}/>
                          <span style={{fontSize:10,fontWeight:700,color}}>Group {gid}</span>
                          <span style={{fontSize:9,color:"#4a7fa0"}}>{group?.members?.length||0} staff</span>
                        </div>
                      ):<span style={{fontSize:9,color:"#2a5070"}}>Unassigned</span>}
                    </div>
                  );
                })}
              </div>
            </>
          )}

          {/* Bump History Ledger */}
          {(()=>{
            const allBumps = Object.values(rosters||{}).reduce((acc,r)=>{
              Object.entries(r.bumpHistory||{}).forEach(([k,v])=>{acc[k]=(acc[k]||0)+v;});
              return acc;
            },{});
            if(!Object.keys(allBumps).length) return null;
            return(
              <>
                <h3 style={{...C.sectionH,marginTop:24}}>Night Surplus — Bump History</h3>
                <div style={{fontSize:11,color:"#4a7fa0",marginBottom:10}}>
                  Staff moved from nights to Day/Evening due to surplus night coverage.
                  Staff with fewer bumps are prioritised for bumping next time to ensure fairness.
                  🟢 = 1 bump &nbsp; 🟡 = 2 bumps &nbsp; 🔴 = 3+ bumps
                </div>
                <div style={{border:"1px solid #1a3050",borderRadius:8,overflow:"hidden",marginBottom:16}}>
                  <table style={{borderCollapse:"collapse",width:"100%"}}>
                    <thead>
                      <tr style={{background:"#060e18"}}>
                        {["Staff","Class","Total Bumps","Protection"].map(h=>(
                          <th key={h} style={{...C.th,textAlign:"left",padding:"8px 12px",fontSize:11}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries(allBumps)
                        .sort(([,a],[,b])=>b-a)
                        .map(([sId,count],i)=>{
                          const s=staff.find(x=>x.id===sId); if(!s)return null;
                          const cls=CLASSIFICATIONS[s.cls];
                          const col=count>=3?"#ef5350":count>=2?"#ffa726":"#66bb6a";
                          const neverBump=s.cls==="ANUM"||s.cls==="GNP"||s.permNights;
                          return(
                            <tr key={sId} style={{background:i%2===0?"#0a1828":"#07101e",borderTop:"1px solid #0d1e30"}}>
                              <td style={{...C.td,padding:"8px 12px",fontWeight:600,color:"#c8d8e8"}}>{fullName(s)}</td>
                              <td style={{...C.td,padding:"8px 12px"}}>
                                <span style={{fontSize:10,color:cls?.color,background:cls?.color+"22",borderRadius:6,padding:"1px 7px",fontWeight:700}}>{s.cls}</span>
                              </td>
                              <td style={{...C.td,padding:"8px 12px"}}>
                                <span style={{fontWeight:700,fontSize:13,color:col}}>{count}</span>
                              </td>
                              <td style={{...C.td,padding:"8px 12px",fontSize:10,color:neverBump?"#ef9a9a":"#4a7fa0"}}>
                                {neverBump?"🛡 Never bumped (protected)":"Eligible for bumping"}
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>
              </>
            );
          })()}
        </div>
      )}

      {/* ── ANNUAL TIMELINE VIEW ── */}
      {view==="timeline"&&(()=>{
        const viewMonths=[0,1]; // show 2 months at a time for performance

        return(
        <div>
          <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:16,flexWrap:"wrap"}}>
            <div style={{display:"flex",gap:8,alignItems:"center",background:"#0a1828",borderRadius:8,border:"1px solid #1a3050",padding:"6px 12px"}}>
              <button style={C.weekNavBtn} onClick={()=>setYear(y=>y-1)}>‹ Year</button>
              <strong style={{color:"#a8dadc",fontSize:16,minWidth:44,textAlign:"center"}}>{year}</strong>
              <button style={C.weekNavBtn} onClick={()=>setYear(y=>y+1)}>Year ›</button>
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center",background:"#0a1828",borderRadius:8,border:"1px solid #1a3050",padding:"6px 12px"}}>
              <button style={C.weekNavBtn} onClick={()=>setMonthOffset(m=>Math.max(0,m-2))}>‹ Prev 2mo</button>
              <strong style={{color:"#7fb3d3",fontSize:13,minWidth:120,textAlign:"center"}}>
                {MONTH_NAMES[monthOffset%12]} – {MONTH_NAMES[(monthOffset+1)%12]}
              </strong>
              <button style={C.weekNavBtn} onClick={()=>setMonthOffset(m=>Math.min(10,m+2))}>Next 2mo ›</button>
            </div>
            {!nightPlanData&&<span style={{fontSize:11,color:"#ef5350"}}>⚠ No night plan generated — go to Planner tab first</span>}
          </div>

          <div style={{fontSize:10,color:"#2a6080",marginBottom:10}}>
            Showing 2 months at a time for performance. Use Prev/Next to scroll through the year.
          </div>

          {/* Group colour legend */}
          {nightPlanData&&(
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
              {(nightPlanData.groups||[]).map(g=>{
                const color=gc[g.id-1]||"#64b5f6";
                const members=staff.filter(s=>g.members.includes(s.id));
                return(
                  <div key={g.id} style={{display:"flex",alignItems:"center",gap:6,background:"#0a1828",border:`1px solid ${color}55`,borderRadius:7,padding:"5px 10px"}}>
                    <div style={{width:10,height:10,borderRadius:"50%",background:color,flexShrink:0}}/>
                    <span style={{fontSize:11,fontWeight:700,color}}> Group {g.id}</span>
                    <span style={{fontSize:10,color:"#4a7fa0"}}>({members.length} staff)</span>
                  </div>
                );
              })}
              <div style={{display:"flex",alignItems:"center",gap:6,background:"#0a1828",border:"1px solid #1a3050",borderRadius:7,padding:"5px 10px"}}>
                <div style={{width:10,height:10,borderRadius:"50%",background:"#3949ab",flexShrink:0}}/>
                <span style={{fontSize:11,fontWeight:700,color:"#7986cb"}}>Permanent Nights</span>
              </div>
            </div>
          )}

          {/* Timeline grid — 2 months at a time, staff down the side */}
          {nightPlanData&&(()=>{
            // Build a day-level map only for the visible 2-month window
            const dayMap={};
            const windowStart=new Date(year,monthOffset,1);
            const windowEnd=new Date(year,monthOffset+2,0); // last day of 2nd month
            (nightPlanData.fns||[]).forEach(fn=>{
              const gid=nightPlanData.groupAssignments?.[fn.key];
              if(!gid)return;
              const fnStart=parseLocalDate(typeof fn.start==="string"?fn.start.split("T")[0]:isoDate(fn.start));
              for(let d=0;d<14;d++){
                const dt=addDays(fnStart,d);
                if(dt>=windowStart&&dt<=windowEnd){
                  dayMap[isoDate(dt)]=gid;
                }
              }
            });

            const nightStaff=[
              ...permStaff,
              ...(nightPlanData.groups||[]).flatMap(g=>
                staff.filter(s=>g.members.includes(s.id)&&!s.permNights)
              ),
            ].filter((s,i,arr)=>arr.findIndex(x=>x.id===s.id)===i);

            // Build only the 2 visible months
            const months=viewMonths.map(offset=>{
              const m=monthOffset+offset;
              const days=[];
              let d=new Date(year,m,1);
              const targetMonth=((m%12)+12)%12;
              while(d.getMonth()===targetMonth){ days.push(isoDate(d)); d=addDays(d,1); }
              return {month:targetMonth,label:MONTH_NAMES[targetMonth],days};
            });

            if(nightStaff.length===0){
              return <div style={{color:"#2a5070",fontSize:12,padding:16}}>No staff assigned to night groups yet.</div>;
            }

            return(
              <div style={{overflowX:"auto"}}>
                <table style={{borderCollapse:"collapse",tableLayout:"fixed",fontSize:10}}>
                  <thead>
                    <tr>
                      <th style={{...C.th,minWidth:160,textAlign:"left",paddingLeft:10,position:"sticky",left:0,background:"#06101a",zIndex:4}}>Staff</th>
                      {months.map(({label,days,month})=>(
                        <th key={month} colSpan={days.length} style={{
                          ...C.th,background:"#060e18",color:"#4fc3f7",
                          borderLeft:"2px solid #1a3050",fontSize:11,padding:"6px 4px",
                          minWidth:days.length*14,
                        }}>{label}</th>
                      ))}
                    </tr>
                    <tr>
                      <th style={{...C.th,position:"sticky",left:0,background:"#06101a",zIndex:4}}/>
                      {months.flatMap(({days})=>days.map(iso=>{
                        const dt=parseLocalDate(iso); const di=dayIdx(iso); const wknd=di>=5;
                        return(
                          <th key={iso} style={{
                            ...C.th,minWidth:14,maxWidth:14,padding:"2px 0",fontSize:8,
                            background:wknd?"#130a1a":"#080f1a",
                            color:wknd?"#3a1a50":"#2a5070",
                            borderLeft:dt.getDate()===1?"2px solid #1a3050":"none",
                          }}>{dt.getDate()}</th>
                        );
                      }))}
                    </tr>
                  </thead>
                  <tbody>
                    {nightStaff.map((s,ri)=>{
                      const cls=CLASSIFICATIONS[s.cls];
                      const staffGroup=nightPlanData.groups?.find(g=>g.members.includes(s.id));
                      const staffGroupColor=staffGroup?gc[staffGroup.id-1]||"#64b5f6":null;

                      return(
                        <tr key={s.id} style={{background:ri%2===0?"#0a1828":"#07101e"}}>
                          <td style={{
                            ...C.td,padding:"3px 10px",position:"sticky",left:0,
                            background:ri%2===0?"#0a1828":"#07101e",zIndex:2,
                            borderRight:"1px solid #1a3050",minWidth:160,
                          }}>
                            <div style={{display:"flex",alignItems:"center",gap:6}}>
                              {staffGroupColor&&<div style={{width:7,height:7,borderRadius:"50%",background:staffGroupColor,flexShrink:0}}/>}
                              {s.permNights&&<div style={{width:7,height:7,borderRadius:"50%",background:"#3949ab",flexShrink:0}}/>}
                              <span style={{fontSize:11,color:"#c8d8e8",whiteSpace:"nowrap"}}>{fullName(s)}</span>
                              <span style={{fontSize:9,color:cls?.color,marginLeft:"auto"}}>{s.cls}</span>
                            </div>
                          </td>
                          {months.flatMap(({days})=>days.map(iso=>{
                            const dt=parseLocalDate(iso);
                            const isNightBlock = s.permNights || !!(nightPlanData.plan?.[s.id]?.[iso]);
                            const gid=isNightBlock&&!s.permNights?dayMap[iso]:null;
                            const color=s.permNights?"#3949ab":(gid?gc[gid-1]||"#64b5f6":null);
                            const wknd=dayIdx(iso)>=5;
                            return(
                              <td key={iso} style={{
                                ...C.td,
                                padding:0,
                                minWidth:14,maxWidth:14,
                                background:isNightBlock
                                  ?(color?color+"99":"#1a237e99")
                                  :wknd?"#0d0814":"transparent",
                                borderLeft:dt.getDate()===1?"2px solid #1a3050":"none",
                                borderBottom:"1px solid #0a1525",
                              }} title={isNightBlock?`${fullName(s)}: Night${s.permNights?" (Perm)":` - Group ${gid}`} — ${iso}`:""}/>
                            );
                          }))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <div style={{fontSize:10,color:"#2a5070",marginTop:10}}>
                  Each column = 1 day. Coloured blocks = night shifts assigned.
                </div>
              </div>
            );
          })()}

          {!nightPlanData&&(
            <div style={C.empty}>
              <div style={{fontSize:36,marginBottom:8}}>🌙</div>
              <div style={C.emptyH}>No Night Plan Generated</div>
              <div style={C.emptySub}>Go to the Planner tab, select a year, and click Auto-Generate.</div>
              <button style={{...C.btnPrimary,marginTop:16}} onClick={()=>setView("planner")}>Go to Planner →</button>
            </div>
          )}
        </div>
        );
      })()}
    </div>
  );
}

// ─── ADO LEDGER TAB ──────────────────────────────────────────
// Shows each FT staff member's ADO balance across all saved rosters.
// Balance = ADOs auto-inserted by generator + manual adjustments by NUM.
// NUM can manually add/subtract ADOs to account for changes after roster was locked.

function ADOTab({ staff, rosters, adoAdjustments, setAdoAdjustments, toast }) {
  const [editingId, setEditingId]   = useState(null); // staffId being edited
  const [adjNote, setAdjNote]       = useState("");
  const [adjValue, setAdjValue]     = useState(0);    // +/- number of ADOs
  const [filterCls, setFilterCls]   = useState("ALL");
  const [showHistory, setShowHistory] = useState(null); // staffId whose history is open

  // Only FT staff (80h+) accrue ADOs
  const eligibleStaff = staff.filter(s => s.hrs >= 80 && !s.archived);
  const filtered = eligibleStaff.filter(s => filterCls === "ALL" || s.cls === filterCls);

  // Sort rosters chronologically
  const sortedRosterKeys = Object.keys(rosters).sort();

  // Compute ADO summary for each staff member across ALL rosters
  function computeADO(staffId) {
    let autoTotal = 0;
    const rosterBreakdown = [];

    sortedRosterKeys.forEach(key => {
      const r = rosters[key];
      if (!r) return;
      const inserted = r.adoInserted?.[staffId]?.length || 0;
      if (inserted > 0) {
        rosterBreakdown.push({
          key,
          startDate: r.startDate || key.split("_w")[0],
          weeks: r.weeks || 2,
          count: inserted,
          dates: r.adoInserted?.[staffId] || [],
          locked: r.locked || false,
        });
        autoTotal += inserted;
      }
    });

    // Manual adjustments: array of { date, value, note }
    const adjustments = adoAdjustments[staffId] || [];
    const manualTotal = adjustments.reduce((sum, a) => sum + a.value, 0);

    return {
      autoTotal,
      manualTotal,
      balance: autoTotal + manualTotal,
      rosterBreakdown,
      adjustments,
    };
  }

  function addAdjustment() {
    if (!editingId) return;
    if (adjValue === 0) return toast("Enter a non-zero adjustment value", "err");
    const entry = {
      id:    Date.now(),
      date:  isoDate(new Date()),
      value: adjValue,
      note:  adjNote.trim() || (adjValue > 0 ? "Manual addition" : "Manual deduction"),
    };
    setAdoAdjustments(prev => {
      const next = { ...prev, [editingId]: [...(prev[editingId] || []), entry] };
      saveADOAdjustment(editingId, entry); // persist to Supabase
      return next;
    });
    setEditingId(null); setAdjNote(""); setAdjValue(0);
    toast(`ADO adjustment saved for ${staff.find(s=>s.id===editingId)?.name}`);
  }

  function removeAdjustment(staffId, entryId) {
    setAdoAdjustments(prev => {
      const next = { ...prev, [staffId]: (prev[staffId]||[]).filter(a=>a.id!==entryId) };
      deleteADOAdjustment(entryId); // remove from Supabase
      return next;
    });
    toast("Adjustment removed");
  }

  // Balance colour
  function balColor(bal) {
    if (bal <= 0) return "#ef5350";
    if (bal === 1) return "#ffa726";
    return "#66bb6a";
  }

  // Export ADO ledger to Excel
  function exportLedger() {
    const wb = XLSX.utils.book_new();
    const rows = [["Staff","Classification","Contract","ADOs (auto)","ADOs (manual adj.)","Balance","Notes"]];
    eligibleStaff.forEach(s => {
      const ado = computeADO(s.id);
      rows.push([fullName(s), s.cls, `${s.hrs}h/fn`, ado.autoTotal, ado.manualTotal, ado.balance,
        ado.adjustments.map(a=>`${a.date}: ${a.value>0?"+":""}${a.value} (${a.note})`).join("; ")
      ]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [20,8,8,10,14,8,40].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws, "ADO Ledger");
    XLSX.writeFile(wb, `ADO_Ledger_${isoDate(new Date())}.xlsx`);
    toast("ADO ledger exported");
  }

  return (
    <div>
      <div style={C.toolbar}>
        <h2 style={C.pageH}>🗓 ADO Ledger</h2>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <select style={C.sel} value={filterCls} onChange={e=>setFilterCls(e.target.value)}>
            <option value="ALL">All Classifications</option>
            {Object.keys(CLASSIFICATIONS).map(k=><option key={k} value={k}>{k}</option>)}
          </select>
          <button style={C.btnSec} onClick={exportLedger}>⬇ Export Ledger</button>
        </div>
      </div>

      {/* Info banner */}
      <div style={{background:"#0a1525",border:"1px solid #1a3050",borderRadius:8,padding:"10px 16px",marginBottom:16,fontSize:11,color:"#4a7fa0",lineHeight:1.8}}>
        <b style={{color:"#7fb3d3"}}>ADO Accrual Rules:</b> Full-time staff (80h+/fn) accrue 2h/week.
        Day/evening staff earn 1 ADO per 4-week cycle (8h). Night staff earn 1 ADO per 5-week cycle (10h).
        ADOs are auto-inserted by the roster generator. Use <b style={{color:"#a5d6a7"}}>+ Adjustment</b> to record
        ADOs taken outside of a generated roster, or to correct the balance manually.
      </div>

      {eligibleStaff.length === 0 && (
        <div style={C.empty}>
          <div style={{fontSize:36,marginBottom:8}}>🗓</div>
          <div style={C.emptyH}>No Full-Time Staff Found</div>
          <div style={C.emptySub}>ADOs only apply to staff contracted at 80h/fortnight or more.</div>
        </div>
      )}

      {/* Ledger table */}
      {filtered.length > 0 && (
        <div style={{border:"1px solid #1a3050",borderRadius:8,overflow:"hidden",marginBottom:24}}>
          <table style={{borderCollapse:"collapse",width:"100%"}}>
            <thead>
              <tr style={{background:"#060e18"}}>
                {["Staff","Class","Contract","ADOs Rostered","Adjustments","Balance",""].map(h=>(
                  <th key={h} style={{...C.th,textAlign:"left",padding:"10px 14px",fontSize:11}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((s,i) => {
                const ado    = computeADO(s.id);
                const cls    = CLASSIFICATIONS[s.cls];
                const bal    = ado.balance;
                const isOpen = showHistory === s.id;

                return (
                  <>
                    <tr key={s.id} style={{background:i%2===0?"#0a1828":"#07101e",borderTop:"1px solid #0d1e30"}}>
                      <td style={{...C.td,padding:"10px 14px",fontWeight:600,color:"#c8d8e8"}}>
                        <div style={{display:"flex",alignItems:"center",gap:7}}>
                          <span style={{width:8,height:8,borderRadius:"50%",background:cls?.color,flexShrink:0}}/>
                          {fullName(s)}
                          {s.permNights && <span style={C.bdg("#1a237e","#7986cb")}>Nights</span>}
                        </div>
                      </td>
                      <td style={{...C.td,padding:"10px 14px"}}>
                        <span style={{fontSize:10,color:cls?.color,background:cls?.color+"22",borderRadius:6,padding:"1px 7px",fontWeight:700}}>{s.cls}</span>
                      </td>
                      <td style={{...C.td,padding:"10px 14px",color:"#7fb3d3",fontSize:11}}>{s.hrs}h/fn</td>
                      <td style={{...C.td,padding:"10px 14px",color:"#a5d6a7",fontWeight:700}}>
                        {ado.autoTotal} ADO{ado.autoTotal!==1?"s":""}
                        {ado.rosterBreakdown.length>0&&(
                          <button style={{...C.btnSec,fontSize:9,padding:"2px 7px",marginLeft:8}} onClick={()=>setShowHistory(isOpen?null:s.id)}>
                            {isOpen?"▲ Hide":"▼ Detail"}
                          </button>
                        )}
                      </td>
                      <td style={{...C.td,padding:"10px 14px"}}>
                        {ado.manualTotal !== 0 && (
                          <span style={{color:ado.manualTotal>0?"#a5d6a7":"#ef9a9a",fontWeight:700,fontSize:12}}>
                            {ado.manualTotal>0?"+":""}{ado.manualTotal}
                          </span>
                        )}
                        {ado.adjustments.length===0&&<span style={{color:"#2a5070",fontSize:11}}>—</span>}
                      </td>
                      <td style={{...C.td,padding:"10px 14px"}}>
                        <span style={{
                          fontSize:15,fontWeight:700,color:balColor(bal),
                          background:balColor(bal)+"22",borderRadius:8,
                          padding:"3px 12px",display:"inline-block"
                        }}>
                          {bal} ADO{bal!==1?"s":""}
                        </span>
                      </td>
                      <td style={{...C.td,padding:"8px 14px"}}>
                        <button style={{...C.btnSec,fontSize:10,padding:"4px 10px",color:"#a5d6a7",borderColor:"#2e7d32"}}
                          onClick={()=>{setEditingId(s.id);setAdjValue(0);setAdjNote("");}}>
                          + Adjustment
                        </button>
                      </td>
                    </tr>

                    {/* Roster breakdown detail */}
                    {isOpen && (
                      <tr key={s.id+"_detail"} style={{background:"#060d18"}}>
                        <td colSpan={7} style={{padding:"0 14px 12px 32px"}}>
                          <div style={{fontSize:11,color:"#4a7fa0",marginBottom:6,marginTop:8}}>ADOs auto-inserted by roster generator:</div>
                          {ado.rosterBreakdown.map(rb=>(
                            <div key={rb.key} style={{display:"flex",gap:12,alignItems:"center",marginBottom:4,flexWrap:"wrap"}}>
                              <span style={{color:"#64b5f6",fontWeight:600,minWidth:90}}>{fmtDate(rb.startDate)}</span>
                              <span style={{color:"#7fb3d3"}}>{rb.weeks===2?"Fortnight":"4 weeks"}</span>
                              <span style={{color:"#a5d6a7",fontWeight:700}}>{rb.count} ADO{rb.count!==1?"s":""}</span>
                              <span style={{color:"#2a5070",fontSize:10}}>{rb.dates.map(d=>fmtDate(d)).join(", ")}</span>
                              {rb.locked&&<span style={C.bdg("#4a2000","#f39c12")}>🔒 Locked</span>}
                            </div>
                          ))}
                          {/* Manual adjustments */}
                          {ado.adjustments.length>0&&(
                            <>
                              <div style={{fontSize:11,color:"#4a7fa0",marginTop:10,marginBottom:6}}>Manual adjustments:</div>
                              {ado.adjustments.map(adj=>(
                                <div key={adj.id} style={{display:"flex",gap:12,alignItems:"center",marginBottom:4,background:"#0a1525",borderRadius:6,padding:"5px 10px",flexWrap:"wrap"}}>
                                  <span style={{color:"#64b5f6",minWidth:90}}>{fmtDate(adj.date)}</span>
                                  <span style={{color:adj.value>0?"#a5d6a7":"#ef9a9a",fontWeight:700}}>
                                    {adj.value>0?"+":""}{adj.value} ADO{Math.abs(adj.value)!==1?"s":""}
                                  </span>
                                  <span style={{color:"#7fb3d3",flex:1}}>{adj.note}</span>
                                  <button style={{background:"none",border:"none",color:"#ef5350",cursor:"pointer",fontSize:13}}
                                    onClick={()=>removeAdjustment(s.id,adj.id)}>✕</button>
                                </div>
                              ))}
                            </>
                          )}
                        </td>
                      </tr>
                    )}
                  </>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* Summary stats */}
      {eligibleStaff.length > 0 && (
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:12,marginBottom:24}}>
          {[
            ["Total ADOs Rostered", eligibleStaff.reduce((s,x)=>s+computeADO(x.id).autoTotal,0), "#a5d6a7"],
            ["Manual Adjustments",  eligibleStaff.reduce((s,x)=>s+computeADO(x.id).adjustments.length,0), "#90caf9"],
            ["Staff with Balance>0",eligibleStaff.filter(x=>computeADO(x.id).balance>0).length, "#66bb6a"],
            ["Staff with Balance=0",eligibleStaff.filter(x=>computeADO(x.id).balance===0).length, "#4a7fa0"],
          ].map(([l,v,col])=>(
            <div key={l} style={{background:"#0a1828",border:"1px solid #1a3050",borderRadius:8,padding:"12px 16px",textAlign:"center"}}>
              <div style={{fontSize:24,fontWeight:700,color:col}}>{v}</div>
              <div style={{fontSize:10,color:"#4a7fa0",marginTop:4}}>{l}</div>
            </div>
          ))}
        </div>
      )}

      {/* Adjustment modal */}
      {editingId && (()=>{
        const s = staff.find(x=>x.id===editingId);
        const ado = computeADO(editingId);
        return(
          <div style={{position:"fixed",inset:0,background:"#000000aa",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}}
            onClick={()=>setEditingId(null)}>
            <div style={{background:"#0a1828",border:"1px solid #1a3050",borderRadius:12,padding:28,minWidth:340,maxWidth:420}}
              onClick={e=>e.stopPropagation()}>
              <h3 style={{color:"#4fc3f7",fontSize:15,fontWeight:700,marginBottom:4}}>ADO Adjustment</h3>
              <div style={{color:"#7fb3d3",fontSize:12,marginBottom:16}}>{s?.name} — current balance: <b style={{color:balColor(ado.balance)}}>{ado.balance} ADO{ado.balance!==1?"s":""}</b></div>

              <label style={C.lbl}>Adjustment (+/−)</label>
              <div style={{display:"flex",gap:8,marginBottom:10}}>
                {[-3,-2,-1,1,2,3].map(v=>(
                  <button key={v} onClick={()=>setAdjValue(v)} style={{
                    flex:1,padding:"8px 0",borderRadius:7,fontFamily:"inherit",fontWeight:700,fontSize:13,cursor:"pointer",
                    background:adjValue===v?(v>0?"#1b5e20":"#7f0000"):"#0f1e35",
                    color:adjValue===v?"#fff":(v>0?"#a5d6a7":"#ef9a9a"),
                    border:`1px solid ${adjValue===v?(v>0?"#27ae60":"#ef5350"):"#1a3050"}`,
                  }}>{v>0?"+":""}{v}</button>
                ))}
              </div>
              <div style={{display:"flex",gap:8,marginBottom:10}}>
                <input type="number" style={{...C.inp,marginBottom:0,flex:1}} placeholder="Or enter custom value…"
                  value={adjValue||""} onChange={e=>setAdjValue(Number(e.target.value))}/>
              </div>

              <label style={C.lbl}>Reason / Note</label>
              <input type="text" style={C.inp} placeholder="e.g. ADO taken 12 Jun not recorded in roster"
                value={adjNote} onChange={e=>setAdjNote(e.target.value)}/>

              <div style={{background:"#060d18",borderRadius:6,padding:"8px 12px",marginBottom:16,fontSize:11,color:"#4a7fa0"}}>
                New balance will be: <b style={{color:balColor(ado.balance+adjValue)}}>{ado.balance+adjValue} ADO{(ado.balance+adjValue)!==1?"s":""}</b>
              </div>

              <div style={{display:"flex",gap:10}}>
                <button style={{...C.btnPrimary,flex:1}} onClick={addAdjustment}>Save Adjustment</button>
                <button style={C.btnSec} onClick={()=>setEditingId(null)}>Cancel</button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}
// ─── HISTORY TAB ─────────────────────────────────────────────
function HistoryTab({ rosters, staff, activeKey, setActiveKey, setTab, onDelete }) {
  const [compareA, setCompareA] = useState(null);
  const [compareB, setCompareB] = useState(null);
  const [mode, setMode]         = useState("list");

  const sortedKeys = Object.keys(rosters).sort().reverse();

  function openRoster(key) { setActiveKey(key); setTab("roster"); }

  function rosterStats(key) {
    const r = rosters[key]; if (!r) return null;
    const days = r.days || [], warnings = r.warnings || [];
    const hs = r.hoursSummary || {};
    const onTarget  = Object.values(hs).filter(h=>Math.abs(h.variance||0)<=8).length;
    const offTarget = Object.values(hs).filter(h=>Math.abs(h.variance||0)>8).length;
    const totalShifts = days.reduce((acc,d)=>acc+(r.roster[d]?.D?.length||0)+(r.roster[d]?.E?.length||0)+(r.roster[d]?.N?.length||0),0);
    return { days:days.length, warnings, onTarget, offTarget, totalShifts,
             adoTotal:r.adoTotal||0, locked:r.locked, lockedAt:r.lockedAt,
             generatedAt:r.generatedAt, weeks:r.weeks||2, startDate:r.startDate };
  }

  function ComparePanel({ keyA, keyB }) {
    const rA=rosters[keyA], rB=rosters[keyB];
    if (!rA||!rB) return null;
    const sA=rosterStats(keyA), sB=rosterStats(keyB);
    const rows=[
      ["Period",           `${fmtDate(sA.startDate)} (${sA.weeks}w)`,`${fmtDate(sB.startDate)} (${sB.weeks}w)`],
      ["Total days",       sA.days,          sB.days],
      ["Total warnings",   sA.warnings.length,sB.warnings.length],
      ["Staffing issues",  sA.warnings.filter(w=>w.type==="staffing").length, sB.warnings.filter(w=>w.type==="staffing").length],
      ["In-charge issues", sA.warnings.filter(w=>w.type==="incharge").length, sB.warnings.filter(w=>w.type==="incharge").length],
      ["Hours issues",     sA.warnings.filter(w=>w.type==="hours").length,    sB.warnings.filter(w=>w.type==="hours").length],
      ["Staff on target",  sA.onTarget,      sB.onTarget],
      ["Staff off target", sA.offTarget,     sB.offTarget],
      ["ADOs inserted",    sA.adoTotal,      sB.adoTotal],
      ["Status",           sA.locked?"🔒 Locked":"✏️ Draft", sB.locked?"🔒 Locked":"✏️ Draft"],
    ];
    const staffToShow = staff.filter(s=>!s.archived).slice(0,30);
    return (
      <div>
        <h3 style={{...C.sectionH,marginBottom:16}}>Roster Comparison</h3>
        <div style={{border:"1px solid #1a3050",borderRadius:8,overflow:"hidden",marginBottom:24}}>
          <table style={{borderCollapse:"collapse",width:"100%"}}>
            <thead>
              <tr style={{background:"#060e18"}}>
                <th style={{...C.th,textAlign:"left",padding:"10px 14px",minWidth:160}}>Metric</th>
                <th style={{...C.th,padding:"10px 14px",color:"#64b5f6",minWidth:180}}>{keyA}</th>
                <th style={{...C.th,padding:"10px 14px",color:"#ffa726",minWidth:180}}>{keyB}</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(([label,vA,vB],i)=>{
                const diff=String(vA)!==String(vB);
                return (
                  <tr key={label} style={{background:i%2===0?"#0a1828":"#07101e",borderTop:"1px solid #0d1e30"}}>
                    <td style={{...C.td,padding:"9px 14px",color:"#7fb3d3",fontWeight:600}}>{label}</td>
                    <td style={{...C.td,padding:"9px 14px",textAlign:"center",color:diff?"#64b5f6":"#a8dadc",fontWeight:diff?700:400}}>{vA}</td>
                    <td style={{...C.td,padding:"9px 14px",textAlign:"center",color:diff?"#ffa726":"#a8dadc",fontWeight:diff?700:400}}>{vB}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <h3 style={{...C.sectionH,marginBottom:12}}>Hours Comparison — per staff member</h3>
        <div style={{border:"1px solid #1a3050",borderRadius:8,overflow:"hidden"}}>
          <table style={{borderCollapse:"collapse",width:"100%"}}>
            <thead>
              <tr style={{background:"#060e18"}}>
                <th style={{...C.th,textAlign:"left",padding:"9px 14px"}}>Staff</th>
                <th style={{...C.th,padding:"9px 14px",color:"#4a7fa0"}}>Contract</th>
                <th style={{...C.th,padding:"9px 14px",color:"#64b5f6"}}>{keyA}</th>
                <th style={{...C.th,padding:"9px 14px",color:"#ffa726"}}>{keyB}</th>
                <th style={{...C.th,padding:"9px 14px"}}>Diff</th>
              </tr>
            </thead>
            <tbody>
              {staffToShow.map((s,i)=>{
                const wA=rA.hoursSummary?.[s.id]?.worked??"—";
                const wB=rB.hoursSummary?.[s.id]?.worked??"—";
                const diff=(typeof wA==="number"&&typeof wB==="number")?wB-wA:null;
                const cls=CLASSIFICATIONS[s.cls];
                return (
                  <tr key={s.id} style={{background:i%2===0?"#0a1828":"#07101e",borderTop:"1px solid #0d1e30"}}>
                    <td style={{...C.td,padding:"7px 14px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:6}}>
                        <span style={{width:7,height:7,borderRadius:"50%",background:cls?.color,flexShrink:0}}/>
                        <span style={{fontSize:11,color:"#c8d8e8"}}>{fullName(s)}</span>
                      </div>
                    </td>
                    <td style={{...C.td,padding:"7px 14px",textAlign:"center",color:"#4a7fa0",fontSize:11}}>{s.hrs}h</td>
                    <td style={{...C.td,padding:"7px 14px",textAlign:"center",color:"#64b5f6",fontWeight:600}}>{wA}{typeof wA==="number"?"h":""}</td>
                    <td style={{...C.td,padding:"7px 14px",textAlign:"center",color:"#ffa726",fontWeight:600}}>{wB}{typeof wB==="number"?"h":""}</td>
                    <td style={{...C.td,padding:"7px 14px",textAlign:"center"}}>
                      {diff!==null&&diff!==0&&<span style={{color:diff>0?"#ffa726":"#64b5f6",fontWeight:700,fontSize:11}}>{diff>0?"+":""}{diff}h</span>}
                      {diff===0&&<span style={{color:"#2a5070",fontSize:10}}>same</span>}
                    </td>
                  </tr>
                );
              })}
              {staff.filter(s=>!s.archived).length>30&&(
                <tr><td colSpan={5} style={{...C.td,padding:"8px 14px",color:"#2a5070",fontSize:11,textAlign:"center"}}>Showing first 30 staff</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={C.toolbar}>
        <h2 style={C.pageH}>🕐 Roster History</h2>
        <div style={{display:"flex",gap:8}}>
          <button style={{...C.selBtn,...(mode==="list"?C.selBtnOn:{})}} onClick={()=>setMode("list")}>📋 All Rosters</button>
          <button style={{...C.selBtn,...(mode==="compare"?C.selBtnOn:{})}} onClick={()=>setMode("compare")}>⚖️ Compare Two</button>
        </div>
      </div>

      {sortedKeys.length===0&&(
        <div style={C.empty}>
          <div style={{fontSize:48,marginBottom:12}}>🕐</div>
          <div style={C.emptyH}>No Roster History</div>
          <div style={C.emptySub}>Generated rosters will appear here. Go to ⚡ Generate to build your first roster.</div>
        </div>
      )}

      {/* LIST MODE */}
      {mode==="list"&&sortedKeys.length>0&&(
        <div>
          <div style={{fontSize:11,color:"#4a7fa0",marginBottom:16}}>
            {sortedKeys.length} roster{sortedKeys.length!==1?"s":""} saved.
            Click <b style={{color:"#64b5f6"}}>Open →</b> to view in the Roster tab.
            <b style={{color:"#ef5350"}}> Deleting is permanent and cannot be undone.</b>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {sortedKeys.map(key=>{
              const s=rosterStats(key); if(!s)return null;
              const isActive=key===activeKey;
              const warnCount=s.warnings.length;
              const warnCol=warnCount===0?"#66bb6a":warnCount<5?"#ffa726":"#ef5350";
              return (
                <div key={key} style={{
                  background:isActive?"#0d2035":"#0a1828",
                  border:`1px solid ${isActive?"#1565c0":"#112a42"}`,
                  borderRadius:10,padding:"16px 18px",
                  display:"flex",alignItems:"center",gap:16,flexWrap:"wrap",
                }}>
                  <div style={{flexShrink:0,textAlign:"center",minWidth:48}}>
                    <div style={{fontSize:22}}>{s.locked?"🔒":"✏️"}</div>
                    <div style={{fontSize:9,color:s.locked?"#f39c12":"#66bb6a",marginTop:2,fontWeight:700}}>
                      {s.locked?"Locked":"Draft"}
                    </div>
                  </div>
                  <div style={{flex:1,minWidth:200}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap"}}>
                      <span style={{fontSize:14,fontWeight:700,color:isActive?"#64b5f6":"#c8d8e8"}}>{key}</span>
                      {isActive&&<span style={{fontSize:9,background:"#1565c0",color:"#90caf9",borderRadius:6,padding:"1px 7px",fontWeight:700}}>CURRENT</span>}
                    </div>
                    <div style={{display:"flex",gap:14,flexWrap:"wrap",fontSize:11,color:"#4a7fa0"}}>
                      <span>📅 {s.days} days ({s.weeks===2?"1 fortnight":"4 weeks"})</span>
                      <span style={{color:warnCol}}>⚠ {warnCount} warning{warnCount!==1?"s":""}</span>
                      <span style={{color:"#a5d6a7"}}>🗓 {s.adoTotal} ADO{s.adoTotal!==1?"s":""}</span>
                      {s.offTarget>0&&<span style={{color:"#ffa726"}}>⏱ {s.offTarget} off-target</span>}
                    </div>
                    <div style={{display:"flex",gap:14,flexWrap:"wrap",fontSize:10,color:"#2a5070",marginTop:3}}>
                      {s.generatedAt&&<span>Generated: {fmtDate(s.generatedAt)}</span>}
                      {s.locked&&s.lockedAt&&<span>Locked: {fmtDate(s.lockedAt)}</span>}
                    </div>
                  </div>
                  <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
                    {s.warnings.filter(w=>w.type==="staffing").length>0&&<span style={{...C.warnBadge,background:"#4a2000",fontSize:10}}>👥 {s.warnings.filter(w=>w.type==="staffing").length}</span>}
                    {s.warnings.filter(w=>w.type==="incharge").length>0&&<span style={{...C.warnBadge,background:"#3a0030",fontSize:10}}>⭐ {s.warnings.filter(w=>w.type==="incharge").length}</span>}
                    {s.warnings.filter(w=>w.type==="hours").length>0&&<span style={{...C.warnBadge,background:"#003030",fontSize:10}}>⏱ {s.warnings.filter(w=>w.type==="hours").length}</span>}
                  </div>
                  <div style={{display:"flex",gap:8,flexShrink:0}}>
                    <button style={{...C.btnPrimary,fontSize:11,padding:"6px 14px"}} onClick={()=>openRoster(key)}>Open →</button>
                    <button style={{...C.btnSec,fontSize:11,padding:"6px 14px",color:"#ef5350",borderColor:"#7f000066"}} onClick={()=>onDelete(key)}>🗑 Delete</button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* COMPARE MODE */}
      {mode==="compare"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:24}}>
            {[["A","#64b5f6",compareA,setCompareA],["B","#ffa726",compareB,setCompareB]].map(([lbl,col,val,setter])=>(
              <div key={lbl} style={C.card}>
                <div style={{...C.cardH,color:col}}>Roster {lbl}</div>
                <label style={C.lbl}>Select roster</label>
                <select style={C.inp} value={val||""} onChange={e=>setter(e.target.value||null)}>
                  <option value="">— Select —</option>
                  {sortedKeys.map(k=><option key={k} value={k}>{k}{rosters[k]?.locked?" 🔒":""}</option>)}
                </select>
                {val&&rosterStats(val)&&(
                  <div style={{fontSize:11,color:"#4a7fa0",marginTop:4}}>
                    {rosterStats(val).days} days · {rosterStats(val).warnings.length} warnings · {rosterStats(val).locked?"🔒 Locked":"✏️ Draft"}
                  </div>
                )}
              </div>
            ))}
          </div>
          {compareA&&compareB&&compareA!==compareB
            ?<ComparePanel keyA={compareA} keyB={compareB}/>
            :compareA&&compareB&&compareA===compareB
              ?<div style={{color:"#ef5350",fontSize:12,padding:12}}>Select two different rosters to compare.</div>
              :<div style={{color:"#2a5070",fontSize:12,padding:12}}>Select a roster for both A and B above.</div>
          }
        </div>
      )}
    </div>
  );
}

// ─── STYLES ──────────────────────────────────────────────────
const C = {
  app:       {minHeight:"100vh",background:"#06101a",color:"#c8d8e8",fontFamily:"'DM Mono','Courier New',monospace",fontSize:13},
  header:    {display:"flex",alignItems:"center",gap:16,padding:"10px 22px",background:"#04090f",borderBottom:"1px solid #112030",flexWrap:"wrap"},
  brand:     {display:"flex",alignItems:"center",gap:10,marginRight:6},
  brandTitle:{fontSize:16,fontWeight:700,color:"#4fc3f7",letterSpacing:2,textTransform:"uppercase"},
  brandSub:  {fontSize:9,color:"#1e4060",letterSpacing:1},
  nav:       {display:"flex",gap:3,flex:1,flexWrap:"wrap"},
  navBtn:    {background:"none",border:"none",color:"#3a6080",fontFamily:"inherit",fontSize:11,padding:"7px 13px",cursor:"pointer",borderRadius:5,fontWeight:600},
  navActive: {background:"#0d2035",color:"#4fc3f7"},
  exportBtn: {background:"#0d1e30",color:"#4fc3f7",border:"1px solid #1a4060",borderRadius:6,padding:"7px 14px",fontFamily:"inherit",fontSize:11,cursor:"pointer",fontWeight:600},
  main:      {padding:"20px 22px",overflowX:"auto"},
  notif:     {position:"fixed",top:14,right:14,padding:"10px 18px",borderRadius:7,color:"#fff",fontWeight:700,fontSize:12,zIndex:9999,boxShadow:"0 4px 24px #000a"},
  pageH:     {fontSize:18,fontWeight:700,color:"#4fc3f7",marginBottom:18,letterSpacing:.5},
  sectionH:  {fontSize:13,fontWeight:700,color:"#4fc3f7",marginBottom:10},
  card:      {background:"#0a1828",border:"1px solid #112a42",borderRadius:9,padding:16,marginBottom:0},
  cardH:     {fontSize:11,fontWeight:700,color:"#4fc3f7",marginBottom:12,letterSpacing:.5,textTransform:"uppercase"},
  toolbar:   {display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:10},
  sel:       {background:"#0a1828",color:"#7fb3d3",border:"1px solid #1a3050",borderRadius:6,padding:"6px 10px",fontFamily:"inherit",fontSize:11},
  searchBox: {background:"#0a1828",color:"#7fb3d3",border:"1px solid #1a3050",borderRadius:6,padding:"6px 12px",fontFamily:"inherit",fontSize:11,width:180},
  inp:       {background:"#040c16",color:"#7fb3d3",border:"1px solid #1a3050",borderRadius:6,padding:"7px 10px",fontFamily:"inherit",fontSize:11,width:"100%",marginBottom:10,boxSizing:"border-box"},
  lbl:       {display:"block",fontSize:10,color:"#2a6080",marginBottom:3,letterSpacing:.3},
  btnPrimary:{background:"linear-gradient(135deg,#0d47a1,#1565c0)",color:"#90caf9",border:"none",borderRadius:7,padding:"8px 18px",fontFamily:"inherit",fontWeight:700,fontSize:11,cursor:"pointer"},
  btnSec:    {background:"#0a1828",color:"#4fc3f7",border:"1px solid #1a3050",borderRadius:7,padding:"7px 13px",fontFamily:"inherit",fontSize:11,cursor:"pointer"},
  selBtn:    {background:"#0a1828",color:"#3a6080",border:"1px solid #112030",borderRadius:7,padding:"7px 13px",fontFamily:"inherit",fontSize:11,cursor:"pointer"},
  selBtnOn:  {background:"#0d2535",color:"#4fc3f7",border:"1px solid #1565c0"},
  iconBtn:   {background:"none",border:"none",cursor:"pointer",fontSize:13,padding:"3px"},
  backBtn:   {background:"#0a1828",border:"1px solid #1a3050",color:"#4fc3f7",borderRadius:6,padding:"5px 12px",fontFamily:"inherit",fontSize:11,cursor:"pointer"},
  weekNavBtn:{background:"none",border:"none",color:"#4fc3f7",fontFamily:"inherit",fontSize:12,fontWeight:700,padding:"12px 18px",cursor:"pointer"},
  pill:      {background:"#0a1828",border:"1px solid #1a3050",color:"#4fc3f7",borderRadius:8,padding:"7px 14px",cursor:"pointer",fontFamily:"inherit",fontSize:11},
  warnBadge: {color:"#ffa726",borderRadius:10,padding:"3px 10px",fontSize:10,fontWeight:700},
  th:        {background:"#060e18",color:"#2a6080",padding:"5px 3px",textAlign:"center",fontSize:9,fontWeight:700,borderBottom:"1px solid #112030",whiteSpace:"nowrap",position:"sticky",top:0,zIndex:2},
  td:        {padding:"1px 2px",borderBottom:"1px solid #0a1525",verticalAlign:"middle"},
  fix0:      {position:"sticky",left:0,background:"#06101a",zIndex:3,borderRight:"1px solid #112030"},
  fix1:      {position:"sticky",left:175,background:"#06101a",zIndex:3},
  fix2:      {position:"sticky",left:227,background:"#06101a",zIndex:3},
  bdg:       (bg,fg)=>({fontSize:9,background:bg,color:fg,borderRadius:8,padding:"1px 5px"}),
  empty:     {textAlign:"center",padding:"70px 20px"},
  emptyH:    {fontSize:20,fontWeight:700,color:"#4fc3f7",marginBottom:8},
  emptySub:  {color:"#2a5070",maxWidth:380,margin:"0 auto",lineHeight:1.7,fontSize:12},
};
