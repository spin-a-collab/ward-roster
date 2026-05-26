// ============================================================
// WARD NURSING ROSTER SYSTEM v3 — Completion Pass
// ============================================================
import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

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

const SK = { STAFF:"wr3_staff", ROSTERS:"wr3_rosters", NIGHTPLAN:"wr3_nightPlan" };
const load = (k,fb) => { try { const r=localStorage.getItem(k); return r?JSON.parse(r):fb; } catch { return fb; } };
const save = (k,v)  => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} };

const addDays  = (d,n) => { const x=new Date(d); x.setDate(x.getDate()+n); return x; };
const isoDate  = d => new Date(d).toISOString().split("T")[0];
const fmtDate  = d => new Date(d).toLocaleDateString("en-AU",{day:"2-digit",month:"2-digit",year:"numeric"});
const fmtShort = d => { const dt=new Date(d); return `${dt.getDate()} ${MONTH_NAMES[dt.getMonth()]}`; };
const dayIdx   = d => { const w=new Date(d).getDay(); return w===0?6:w-1; };
const isWknd   = d => dayIdx(d)>=5;
const getMon   = d => { const x=new Date(d),di=x.getDay(); x.setDate(x.getDate()+(di===0?-6:1-di)); x.setHours(0,0,0,0); return x; };

function buildDays(startMon,weeks) {
  return Array.from({length:weeks*7},(_,i)=>{
    const date=addDays(startMon,i);
    return { date, iso:isoDate(date), di:dayIdx(date), wknd:isWknd(date), wk:Math.floor(i/7) };
  });
}

const isInCharge = s => !!(CLASSIFICATIONS[s.cls]?.inCharge || s.inCharge);

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

function autoComputeNightGroups(staff) {
  const eligible = staff.filter(s =>
    !s.permNights && s.cls!=="NUM" &&
    !s.fwaConditions?.some(c=>c.type==="NO_NIGHTS"||
      (c.type==="SPECIFIC_SHIFTS" && c.shifts && !c.shifts.includes("N")))
  );
  const permNights     = staff.filter(s=>s.permNights);
  const permHoursPerFn = permNights.reduce((a,s)=>a+s.hrs,0);
  const rotNeeded      = Math.max(0, 700-permHoursPerFn);
  const avgHrs         = eligible.length ? eligible.reduce((a,s)=>a+s.hrs,0)/eligible.length : 80;
  const staffPerBlock  = Math.max(1,Math.ceil(rotNeeded/avgHrs));
  const numGroups      = Math.max(2,Math.ceil(eligible.length/staffPerBlock));
  const clsPri         = {ANUM:0,CNS:1,RN:2,GNP:3,EN:4};
  const sorted         = [...eligible].sort((a,b)=>{
    const cp=(clsPri[a.cls]||5)-(clsPri[b.cls]||5); if(cp!==0)return cp;
    return b.hrs-a.hrs;
  });
  const groups = Array.from({length:numGroups},(_,i)=>({id:i+1,members:[],totalHours:0}));
  sorted.forEach((s,i)=>{ const g=groups[i%numGroups]; g.members.push(s.id); g.totalHours+=s.hrs; });
  return { groups, numGroups, staffPerBlock, permHoursPerFn };
}

function autoAssignNightPlan(staff,year) {
  const {groups}=autoComputeNightGroups(staff);
  const plan={}, groupAssignments={};
  const startOfYear=getMon(new Date(year,0,1));
  const fns=Array.from({length:26},(_,i)=>({
    idx:i, start:addDays(startOfYear,i*14), end:addDays(startOfYear,i*14+13),
    key:isoDate(addDays(startOfYear,i*14))
  }));
  const lastUsed={};
  fns.forEach((fn,i)=>{
    let best=null,bestGap=-1;
    groups.forEach(g=>{ const gap=i-(lastUsed[g.id]??-99); if(gap>bestGap){bestGap=gap;best=g;} });
    groupAssignments[fn.key]=best.id; lastUsed[best.id]=i;
  });
  fns.forEach(fn=>{
    const group=groups.find(g=>g.id===groupAssignments[fn.key]); if(!group)return;
    group.members.forEach(sid=>{
      for(let d=0;d<14;d++){ const iso=isoDate(addDays(fn.start,d)); if(!plan[sid])plan[sid]={}; plan[sid][iso]=true; }
    });
  });
  return {plan,groupAssignments,groups,fns};
}

const SAMPLE_STAFF = [
  {id:"s1", name:"Alexandra Chen",  cls:"NUM",  hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s2", name:"James Hartley",   cls:"ANUM", hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s3", name:"Maria Santos",    cls:"ANUM", hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"Prefers day shifts",   resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s4", name:"David Okonkwo",   cls:"ANUM", hrs:64,permNights:false,inCharge:true, fwaConditions:[{type:"SPECIFIC_DAYS",days:[0,1,2,3,4],note:"Mon-Fri only (FWA)"}],prefs:"",resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s5", name:"Priya Patel",     cls:"CNS",  hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s6", name:"Tom Nguyen",      cls:"CNS",  hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s7", name:"Sarah Kim",       cls:"RN",   hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s8", name:"Luke Andersen",   cls:"RN",   hrs:64,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s9", name:"Fatima Al-Rawi",  cls:"RN",   hrs:48,permNights:false,inCharge:false,fwaConditions:[{type:"NO_NIGHTS",note:"FWA approved"},{type:"SPECIFIC_DAYS",days:[0,1,2,3,4],note:"Weekdays only"}],prefs:"",resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s10",name:"Grace Torres",    cls:"RN",   hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s11",name:"Ben Murphy",      cls:"RN",   hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"Happy to work weekends",resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s12",name:"Amara Diallo",    cls:"RN",   hrs:80,permNights:false,inCharge:true, fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s13",name:"Chen Wei",        cls:"RN",   hrs:64,permNights:false,inCharge:false,fwaConditions:[{type:"MAX_HOURS_WEEK",value:32,note:"Max 32h/wk (FWA)"}],prefs:"",resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s14",name:"Nina Rodriguez",  cls:"RN",   hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s15",name:"Oscar Pietersen", cls:"RN",   hrs:48,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s16",name:"Yuki Tanaka",     cls:"GNP",  hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,gnpStart:"2025-01-15",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s17",name:"Chloe Martin",    cls:"GNP",  hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,gnpStart:"2025-01-15",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s18",name:"Raj Sharma",      cls:"GNP",  hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,gnpStart:"2025-03-01",leaveCard:[],gridLeave:{},requests:{}},
  {id:"s19",name:"Mei Lin",         cls:"EN",   hrs:64,permNights:false,inCharge:false,fwaConditions:[],prefs:"",                     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s20",name:"Patrick Flynn",   cls:"EN",   hrs:80,permNights:true, inCharge:false,fwaConditions:[],prefs:"Permanent nights",     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s21",name:"Zara Ahmed",      cls:"EN",   hrs:80,permNights:true, inCharge:false,fwaConditions:[],prefs:"Permanent nights",     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s22",name:"Sofia Russo",     cls:"RN",   hrs:80,permNights:true, inCharge:true, fwaConditions:[],prefs:"Permanent nights",     resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s23",name:"James Brennan",   cls:"RN",   hrs:80,permNights:false,inCharge:false,fwaConditions:[{type:"EVENINGS_ONLY",note:"Evening only (FWA)"}],prefs:"",resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
  {id:"s24",name:"Lily Thompson",   cls:"RN",   hrs:64,permNights:false,inCharge:true, fwaConditions:[{type:"SPECIFIC_SHIFTS",shifts:["D","E"],note:"D/E only, no nights (FWA)"}],prefs:"",resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}},
];

// ─── ROSTER GENERATOR v3 ─────────────────────────────────────
function generateRoster({staff,startDate,weeks,nightPlanData,previousRoster,recentWkndCounts}) {
  const startMon=getMon(new Date(startDate));
  const days=buildDays(startMon,weeks);
  const roster={};
  days.forEach(d=>{ roster[d.iso]={D:[],E:[],N:[]}; });

  const leaveMap=buildLeaveMap(staff);

  // Period targets — contracted hours is already per fortnight (2 weeks)
  // For a 4-week roster, double it. weeks=2 → multiplier=1. weeks=4 → multiplier=2.
  const multiplier = weeks / 2;
  const periodTarget={};
  const maxShifts={}; // hard ceiling on number of shifts to prevent runaway assignment
  staff.forEach(s=>{
    const raw = s.hrs * multiplier;
    periodTarget[s.id] = s.permNights ? Math.floor(raw/10)*10 : raw;
    // Max shifts = target hours / shift length, rounded up + 1 buffer
    maxShifts[s.id] = s.permNights
      ? Math.floor(periodTarget[s.id]/10)
      : Math.round(periodTarget[s.id]/8);
  });

  // Shift count tracker — incremented every time a shift is assigned
  const shiftCount={};
  staff.forEach(s=>{ shiftCount[s.id]=0; });

  // Hours worked — pre-load leave hours
  const hw={};
  staff.forEach(s=>{ hw[s.id]=0; });
  staff.forEach(s=>{
    days.forEach(d=>{
      const lc=leaveMap[s.id][d.iso]; if(!lc||!SHIFT_DEF[lc])return;
      hw[s.id]+=(s.permNights||nightPlanData?.plan?.[s.id]?.[d.iso])?10:8;
    });
  });

  // ADO tracking
  const adoAccrued={},adoTaken={},adoMap={};
  staff.forEach(s=>{ adoAccrued[s.id]=0; adoTaken[s.id]=0; adoMap[s.id]={}; });

  // Weekend counts (seeded from history)
  const wkndCnt={};
  staff.forEach(s=>{ wkndCnt[s.id]=(recentWkndCounts?.[s.id]||0); });

  const prevTail=previousRoster?.tail||{};

  function onShift(sId,iso,sh){ return roster[iso]?.[sh]?.includes(sId)||prevTail[iso]?.[sh]?.includes(sId)||false; }
  function working(sId,iso){ return onShift(sId,iso,"D")||onShift(sId,iso,"E")||onShift(sId,iso,"N"); }
  function assignedToday(sId,iso){ const d=roster[iso]; return !!(d&&(d.D.includes(sId)||d.E.includes(sId)||d.N.includes(sId))); }

  function consecNights(sId,iso){ let c=0; for(let i=1;i<=5;i++){if(onShift(sId,isoDate(addDays(new Date(iso),-i)),"N"))c++;else break;} return c; }
  function consecShifts(sId,iso){ let c=0; for(let i=1;i<=6;i++){if(working(sId,isoDate(addDays(new Date(iso),-i))))c++;else break;} return c; }
  function nightWithin47h(sId,iso){ for(let i=1;i<=2;i++){if(onShift(sId,isoDate(addDays(new Date(iso),-i)),"N"))return true;} return false; }

  function weekHrs(sId,iso,adding){
    const mon=getMon(new Date(iso)); let tot=adding;
    for(let i=0;i<7;i++){
      const k=isoDate(addDays(mon,i)); if(k===iso)continue;
      if(leaveMap[sId][k]||adoMap[sId][k]){tot+=8;continue;}
      if(roster[k]?.D.includes(sId)||roster[k]?.E.includes(sId))tot+=8;
      if(roster[k]?.N.includes(sId))tot+=10;
    }
    return tot;
  }

  function canWork(s,iso,shift){
    if(!s||s.resigned)return false;
    if(s.resign&&new Date(iso)>=new Date(s.resign))return false;
    if(leaveMap[s.id][iso])return false;
    if(adoMap[s.id][iso])return false;
    if(assignedToday(s.id,iso))return false;
    if(s.permNights&&shift!=="N")return false;
    if(!s.permNights&&shift==="N"&&!nightPlanData?.plan?.[s.id]?.[iso])return false;
    if(!fwaAllows(s,iso,shift))return false;
    // Hard shift count ceiling — never exceed max shifts for this period
    if(shiftCount[s.id]>=maxShifts[s.id])return false;
    const maxW=s.fwaConditions?.find(c=>c.type==="MAX_HOURS_WEEK")?.value;
    if(maxW&&weekHrs(s.id,iso,shift==="N"?10:8)>maxW)return false;
    if(s.cls==="GNP"&&s.gnpStart&&shift==="N"&&new Date(iso)<addDays(new Date(s.gnpStart),91))return false;
    if(shift==="N"&&consecNights(s.id,iso)>=4)return false;
    if(shift!=="N"&&consecShifts(s.id,iso)>=5)return false;
    if((shift==="D"||shift==="E")&&nightWithin47h(s.id,iso))return false;
    if(shift==="E"&&onShift(s.id,isoDate(addDays(new Date(iso),-1)),"N"))return false;
    return true;
  }

  function rotScore(sId,iso,shift){
    if(shift==="D"&&onShift(sId,isoDate(addDays(new Date(iso),-1)),"E"))return 2;
    if(shift==="E"&&onShift(sId,isoDate(addDays(new Date(iso),-1)),"D"))return 0;
    if(shift==="N"&&onShift(sId,isoDate(addDays(new Date(iso),-1)),"E"))return 0;
    return 1;
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

  const totalGNP=staff.filter(s=>s.cls==="GNP").length;
  const maxGNPShift=Math.max(1,Math.floor(totalGNP/2));
  let numHasWorked=false;

  // Phase 1: Nights
  days.forEach(({iso})=>{
    const eligible=staff.filter(s=>canWork(s,iso,"N")).sort((a,b)=>{
      if(a.permNights&&!b.permNights)return -1; if(!a.permNights&&b.permNights)return 1;
      return (periodTarget[b.id]-hw[b.id])-(periodTarget[a.id]-hw[a.id]);
    });
    let enCnt=0,gnpCnt=0,anumCnt=0,icOk=false;
    for(const s of eligible){
      if(roster[iso].N.length>=5)break;
      if(s.cls==="EN"&&enCnt>=1)continue;
      if(s.cls==="GNP"&&gnpCnt>=1)continue;
      if(s.cls==="ANUM"){
        if(anumCnt>=1){
          const other=roster[iso].N.map(id=>staff.find(x=>x.id===id)).find(x=>x?.cls==="ANUM");
          if(!other||other.hrs<80||s.hrs<80)continue;
        }
        anumCnt++;
      }
      if(s.cls==="EN")enCnt++; if(s.cls==="GNP")gnpCnt++;
      roster[iso].N.push(s.id); hw[s.id]+=10; shiftCount[s.id]++;
      if(isInCharge(s))icOk=true;
    }
    if(!icOk&&roster[iso].N.length>0){
      const ic=eligible.find(s=>!roster[iso].N.includes(s.id)&&isInCharge(s));
      if(ic){roster[iso].N.push(ic.id);hw[ic.id]+=10;shiftCount[ic.id]++;}
    }
  });

  // ADO accrual at each week boundary
  days.filter(d=>d.di===6).forEach(sun=>{
    const monIso=isoDate(addDays(sun.date,-6));
    staff.filter(s=>s.hrs>=80).forEach(s=>{
      let worked=false;
      for(let i=0;i<7;i++){
        const k=isoDate(addDays(new Date(monIso),i));
        if(roster[k]?.D.includes(s.id)||roster[k]?.E.includes(s.id)||roster[k]?.N.includes(s.id)||leaveMap[s.id][k]){worked=true;break;}
      }
      if(!worked)return;
      adoAccrued[s.id]+=2;
      const nightCtx=s.permNights||Array.from({length:7},(_,i)=>isoDate(addDays(new Date(monIso),i))).some(k=>nightPlanData?.plan?.[s.id]?.[k]);
      tryInsertADO(s,sun.iso,nightCtx);
    });
  });

  // Phase 2: Day & Evening
  days.forEach(({iso,di,wknd})=>{
    const BASE=wknd?9:10;
    // Hard overstaffing cap for weekdays: never more than BASE+4 on any single shift
    const OVERSTAFF_CAP=BASE+4;

    for(const shift of ["D","E"]){
      const isNUMSlot=shift==="D"&&di>=1&&di<=3&&!numHasWorked;
      const eligible=staff.filter(s=>{
        if(!canWork(s,iso,shift))return false;
        if(s.permNights)return false;
        if(nightPlanData?.plan?.[s.id]?.[iso])return false;
        // Hard hours cap — never roster someone who has already met their target
        // UNLESS we are below base staffing level (must fill the shift first)
        const rem=periodTarget[s.id]-hw[s.id];
        if(rem<=0&&roster[iso][shift].length>=BASE)return false;
        return true;
      }).sort((a,b)=>{
        const aReq=!!(a.requests?.[`${iso}_${shift}`]),bReq=!!(b.requests?.[`${iso}_${shift}`]);
        if(aReq&&!bReq)return -1; if(!aReq&&bReq)return 1;
        if(wknd){const dw=(wkndCnt[a.id]||0)-(wkndCnt[b.id]||0);if(dw!==0)return dw;}
        const rA=rotScore(a.id,iso,shift),rB=rotScore(b.id,iso,shift);
        if(rA!==rB)return rA-rB;
        // Sort by most hours remaining first (fills those with bigger deficits first)
        return (periodTarget[b.id]-hw[b.id])-(periodTarget[a.id]-hw[a.id]);
      });

      let anumCnt=0,enCnt=0,gnpCnt=0,numOn=false,icOk=false;
      for(const s of eligible){
        const rem=periodTarget[s.id]-hw[s.id];
        // Weekend: hard cap at BASE, no overstaffing ever
        if(wknd&&roster[iso][shift].length>=BASE)break;
        // Weekday: stop at OVERSTAFF_CAP regardless
        if(roster[iso][shift].length>=OVERSTAFF_CAP)break;
        // Weekday: once at BASE, only add staff who still need hours
        if(!wknd&&roster[iso][shift].length>=BASE&&rem<=0)continue;

        if(s.cls==="NUM"){if(!isNUMSlot||numHasWorked)continue;numOn=true;}
        if(s.cls==="ANUM"){
          if(anumCnt>=1)continue;
          if(numOn||roster[iso][shift].some(id=>staff.find(x=>x.id===id)?.cls==="NUM"))continue;
          anumCnt++;
        }
        if(s.cls==="EN"&&enCnt>=1)continue;
        if(s.cls==="GNP"){if(gnpCnt>=maxGNPShift)continue;gnpCnt++;}
        if(s.cls==="EN")enCnt++;

        roster[iso][shift].push(s.id); hw[s.id]+=8; shiftCount[s.id]++;
        if(wknd)wkndCnt[s.id]=(wkndCnt[s.id]||0)+1;
        if(isInCharge(s))icOk=true;
        if(s.cls==="NUM")numHasWorked=true;
      }

      // Guarantee in-charge — but only add if they still have hours remaining
      if(!icOk&&roster[iso][shift].length>0){
        const ic=eligible.find(s=>
          !roster[iso][shift].includes(s.id)&&
          isInCharge(s)&&
          !s.permNights&&
          !nightPlanData?.plan?.[s.id]?.[iso]&&
          (periodTarget[s.id]-hw[s.id])>0&&
          roster[iso][shift].length<OVERSTAFF_CAP
        );
        if(ic){roster[iso][shift].push(ic.id);hw[ic.id]+=8;shiftCount[ic.id]++;if(wknd)wkndCnt[ic.id]=(wkndCnt[ic.id]||0)+1;}
      }
    }
  });

  // Merge ADO into leaveMap for display
  staff.forEach(s=>{ Object.keys(adoMap[s.id]).forEach(iso=>{ leaveMap[s.id][iso]="ADO"; }); });

  // Tail for next roster (last 7 days)
  const tail={};
  days.slice(-7).forEach(d=>{ tail[d.iso]={D:[...roster[d.iso].D],E:[...roster[d.iso].E],N:[...roster[d.iso].N]}; });

  // Hours summary
  const hoursSummary={};
  staff.forEach(s=>{
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
      const hasIC=d[sh].some(id=>{const x=staff.find(y=>y.id===id);return x&&isInCharge(x);});
      if(!hasIC&&d[sh].length>0)warnings.push({iso,sh,type:"incharge",msg:"No In-Charge nurse"});
    });
    if(d.N.length<5)warnings.push({iso,sh:"N",type:"staffing",msg:`Night understaffed: ${d.N.length}/5`});
    const hasNIC=d.N.some(id=>{const x=staff.find(y=>y.id===id);return x&&isInCharge(x);});
    if(!hasNIC&&d.N.length>0)warnings.push({iso,sh:"N",type:"incharge",msg:"No In-Charge (nights)"});
  });
  staff.forEach(s=>{
    const v=hoursSummary[s.id];
    if(Math.abs(v.variance)>8)warnings.push({iso:"—",sh:"Hrs",type:"hours",msg:`${s.name}: ${v.worked}h worked vs ${v.target}h target (${v.variance>0?"+":""}${v.variance}h)`});
  });

  const adoTotal=Object.values(adoMap).reduce((a,m)=>a+Object.keys(m).length,0);
  return {
    roster,leaveMap,hoursWorked:hw,hoursSummary,warnings,
    days:days.map(d=>d.iso),startDate:isoDate(startMon),
    tail,wkndCountEnd:{...wkndCnt},
    adoInserted:Object.fromEntries(staff.map(s=>[s.id,Object.keys(adoMap[s.id])])),
    adoTotal,
  };
}

// ─── APP ─────────────────────────────────────────────────────
export default function App() {
  const [tab,setTab]                    = useState("roster");
  const [staff,setStaff]                = useState(()=>load(SK.STAFF,SAMPLE_STAFF));
  const [rosters,setRosters]            = useState(()=>load(SK.ROSTERS,{}));
  const [nightPlanData,setNightPlanData]= useState(()=>load(SK.NIGHTPLAN,null));
  const [activeKey,setActiveKey]        = useState(null);
  const [notif,setNotif]                = useState(null);
  const [genCfg,setGenCfg]              = useState({startDate:isoDate(getMon(new Date())),weeks:2});

  useEffect(()=>save(SK.STAFF,staff),[staff]);
  useEffect(()=>save(SK.ROSTERS,rosters),[rosters]);
  useEffect(()=>save(SK.NIGHTPLAN,nightPlanData),[nightPlanData]);

  function toast(msg,type="ok"){setNotif({msg,type});setTimeout(()=>setNotif(null),3500);}

  function handleGenerate(){
    try{
      const sortedKeys=Object.keys(rosters).sort();
      const prevRoster=sortedKeys.length?rosters[sortedKeys[sortedKeys.length-1]]:null;
      const recentWknd={};
      sortedKeys.slice(-3).forEach(k=>{
        const r=rosters[k]; if(!r?.wkndCountEnd)return;
        Object.entries(r.wkndCountEnd).forEach(([id,cnt])=>{recentWknd[id]=(recentWknd[id]||0)+cnt;});
      });
      const result=generateRoster({staff,startDate:genCfg.startDate,weeks:genCfg.weeks,nightPlanData,previousRoster:prevRoster,recentWkndCounts:recentWknd});
      const key=`${result.startDate}_w${genCfg.weeks}`;
      setRosters(r=>({...r,[key]:result}));
      setActiveKey(key);
      setTab("roster");
      toast(`Roster generated — ${result.warnings.length} warning(s), ${result.adoTotal} ADO(s) inserted`);
    }catch(e){toast("Error: "+e.message,"err");}
  }

  function handleExport(){
    const r=activeKey&&rosters[activeKey]; if(!r)return toast("No roster","err");
    const wb=XLSX.utils.book_new();
    const hdrs=["Staff","Cls","Contract","Target","Worked","Variance","ADOs",...r.days.map(d=>{const dt=new Date(d);return `${DAY_NAMES[dayIdx(d)]} ${dt.getDate()}/${dt.getMonth()+1}`;})];
    const rows=staff.filter(s=>!s.resigned).map(s=>{
      const hs=r.hoursSummary?.[s.id]||{target:s.hrs,worked:0,variance:0,adoCount:0};
      const row=[s.name,s.cls,s.hrs,hs.target,hs.worked,hs.variance,hs.adoCount];
      r.days.forEach(dk=>{
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
    if(r.warnings.length>0){
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Date","Shift","Type","Warning"],...r.warnings.map(w=>[w.iso,w.sh,w.type,w.msg])]),"Warnings");
    }
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
      ["Staff","Cls","Contract","Target","Worked","Variance","ADOs"],
      ...staff.filter(s=>!s.resigned).map(s=>{const hs=r.hoursSummary?.[s.id]||{target:s.hrs,worked:0,variance:0,adoCount:0};return[s.name,s.cls,s.hrs,hs.target,hs.worked,hs.variance,hs.adoCount];})
    ]),"Hours Summary");
    XLSX.writeFile(wb,`WardRoster_${activeKey}.xlsx`);
    toast("Exported — Roster + Warnings + Hours Summary");
  }

  const activeRoster=activeKey?rosters[activeKey]:null;
  const rosterKeys=Object.keys(rosters).sort().reverse();

  return (
    <div style={C.app}>
      {notif&&<div style={{...C.notif,background:notif.type==="err"?"#7f0000":"#1a5e20"}}>{notif.msg}</div>}
      <header style={C.header}>
        <div style={C.brand}>
          <span style={{fontSize:26}}>⚕</span>
          <div><div style={C.brandTitle}>WardRoster</div><div style={C.brandSub}>Clinical Scheduling System v3</div></div>
        </div>
        <nav style={C.nav}>
          {[["roster","📋 Roster"],["generate","⚡ Generate"],["staff","👥 Staff"],["leave","📅 Leave & Requests"],["nightplan","🌙 Night Planner"]].map(([v,l])=>(
            <button key={v} style={{...C.navBtn,...(tab===v?C.navActive:{})}} onClick={()=>setTab(v)}>{l}</button>
          ))}
        </nav>
        <button style={C.exportBtn} onClick={handleExport}>⬇ Export XLSX</button>
      </header>
      <div style={C.main}>
        {tab==="roster"   &&<RosterTab   roster={activeRoster} staff={staff} rosterKeys={rosterKeys} activeKey={activeKey} setActiveKey={setActiveKey} rosters={rosters} setRosters={setRosters}/>}
        {tab==="generate" &&<GenerateTab cfg={genCfg} setCfg={setGenCfg} onGenerate={handleGenerate} staff={staff} nightPlanData={nightPlanData} rosters={rosters}/>}
        {tab==="staff"    &&<StaffTab    staff={staff} setStaff={setStaff} toast={toast}/>}
        {tab==="leave"    &&<LeaveTab    staff={staff} setStaff={setStaff} toast={toast}/>}
        {tab==="nightplan"&&<NightPlanTab staff={staff} nightPlanData={nightPlanData} setNightPlanData={setNightPlanData} toast={toast}/>}
      </div>
    </div>
  );
}

// ─── ROSTER TAB ──────────────────────────────────────────────
function RosterTab({roster,staff,rosterKeys,activeKey,setActiveKey,rosters,setRosters}){
  const [filter,setFilter]=useState("");
  const [subTab,setSubTab]=useState("grid");
  const [showAll,setShowAll]=useState(false);
  const visible=staff.filter(s=>!s.resigned&&s.name.toLowerCase().includes(filter.toLowerCase()));

  function getCell(sId,iso){
    if(!roster)return null;
    const lc=roster.leaveMap?.[sId]?.[iso]; if(lc)return lc;
    const d=roster.roster[iso];
    if(d?.D.includes(sId))return "D"; if(d?.E.includes(sId))return "E"; if(d?.N.includes(sId))return "N";
    return null;
  }

  function cycleCell(sId,iso){
    if(!roster||!activeKey)return;
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
  const wBT={staffing:0,incharge:0,hours:0};
  (roster.warnings||[]).forEach(w=>{ wBT[w.type]=(wBT[w.type]||0)+1; });

  return(
    <div>
      <div style={C.toolbar}>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <select style={C.sel} value={activeKey||""} onChange={e=>setActiveKey(e.target.value)}>
            {rosterKeys.map(k=><option key={k} value={k}>{k}</option>)}
          </select>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {wBT.staffing>0&&<span style={{...C.warnBadge,background:"#4a2000"}}>👥 {wBT.staffing} staffing</span>}
            {wBT.incharge>0&&<span style={{...C.warnBadge,background:"#3a0030"}}>⭐ {wBT.incharge} in-charge</span>}
            {wBT.hours>0&&<span style={{...C.warnBadge,background:"#003030"}}>⏱ {wBT.hours} hours</span>}
          </div>
        </div>
        <input style={C.searchBox} placeholder="Search staff…" value={filter} onChange={e=>setFilter(e.target.value)}/>
      </div>

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
                    const dt=new Date(iso),di=dayIdx(iso),wknd=di>=5;
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
                        return(
                          <td key={iso}
                            style={{...C.td,background:def?def.bg:wknd?"#0d0814":"transparent",textAlign:"center",padding:1,
                              cursor:CARD_CODES.includes(code)?"default":"pointer",
                              borderLeft:wknd?"2px solid #2a1030":"1px solid #111a28"}}
                            onClick={()=>!CARD_CODES.includes(code)&&cycleCell(nurse.id,iso)}
                            title={code?(def?.label+(autoADO?" ★ auto-inserted":"")):("Off — click to add")}>
                            <span style={{fontSize:9,fontWeight:700,color:def?def.text:"#1a3050",display:"block",lineHeight:"22px"}}>
                              {code||""}{autoADO&&<span style={{fontSize:6,verticalAlign:"super"}}>★</span>}
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
          </div>
          <h3 style={{...C.sectionH,marginTop:20}}>Daily Staffing Counts</h3>
          <div style={{overflowX:"auto",border:"1px solid #1a3050",borderRadius:8,marginTop:6}}>
            <table style={{borderCollapse:"collapse",tableLayout:"fixed",width:"100%"}}>
              <thead>
                <tr>
                  <th style={{...C.th,minWidth:70,textAlign:"left",paddingLeft:10}}>Shift</th>
                  {days.map(iso=>{const dt=new Date(iso),di=dayIdx(iso),wknd=di>=5;
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
                {staff.filter(s=>!s.resigned).map((s,i)=>{
                  const hs=roster.hoursSummary?.[s.id]||{target:s.hrs,worked:0,variance:0,adoCount:0,shifts:0,maxShifts:0};
                  const cls=CLASSIFICATIONS[s.cls];
                  const ok=Math.abs(hs.variance)<=8,warn=Math.abs(hs.variance)<=16;
                  const col=ok?"#66bb6a":warn?"#ffa726":"#ef5350";
                  return(
                    <tr key={s.id} style={{background:i%2===0?"#0a1828":"#07101e",borderTop:"1px solid #0d1e30"}}>
                      <td style={{...C.td,padding:"8px 12px",fontWeight:600,color:"#c8d8e8"}}>{s.name}</td>
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
          {["staffing","incharge","hours"].map(type=>{
            const items=(roster.warnings||[]).filter(w=>w.type===type);
            if(!items.length)return null;
            const meta={staffing:{title:"Staffing Levels",icon:"👥",color:"#ffa726"},incharge:{title:"In-Charge Coverage",icon:"⭐",color:"#ce93d8"},hours:{title:"Hours Variance",icon:"⏱",color:"#80deea"}};
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
function GenerateTab({cfg,setCfg,onGenerate,staff,nightPlanData,rosters}){
  const {numGroups,staffPerBlock,permHoursPerFn}=autoComputeNightGroups(staff);
  const fwaCount=staff.filter(s=>s.fwaConditions?.length>0).length;
  const permNights=staff.filter(s=>s.permNights);
  const sortedKeys=Object.keys(rosters).sort();
  const prevRoster=sortedKeys.length?rosters[sortedKeys[sortedKeys.length-1]]:null;
  const hasTail=!!(prevRoster?.tail&&Object.keys(prevRoster.tail).length>0);
  const prevEnd=prevRoster?.days?.[prevRoster.days.length-1];
  const recentCount=sortedKeys.slice(-3).length;

  return(
    <div>
      <h2 style={C.pageH}>⚡ Generate Roster</h2>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(265px,1fr))",gap:16,marginBottom:24}}>
        <div style={C.card}>
          <div style={C.cardH}>Roster Period</div>
          <label style={C.lbl}>Start Date (Monday)</label>
          <input type="date" style={C.inp} value={cfg.startDate} onChange={e=>setCfg(c=>({...c,startDate:e.target.value}))}/>
          <label style={C.lbl}>Period Length</label>
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
        <div style={C.card}>
          <div style={C.cardH}>Rules Active</div>
          <div style={{display:"flex",flexDirection:"column",gap:4}}>
            {["✅ Contracted hours targeted per staff","✅ Night hrs rounded to 10h multiples","✅ Leave deducted from hours target","✅ ADO auto-inserted at threshold","✅ Forward rotation (D→E→N preferred)","✅ Max 5 consec D/E, 4 consec night shifts","✅ 47h break after nights enforced","✅ Max 1 EN per shift","✅ Max 1 ANUM per shift (FT exception)","✅ ≤50% GNP per shift, max 1 on nights","✅ In-Charge guaranteed every shift","✅ NUM: 1×Day/fn Tue–Thu, no ANUM clash","✅ All FWA conditions enforced","✅ GNP: no nights first 3 months","✅ Weekend fairness (3-roster history)","✅ No overstaffing weekends or nights","✅ Cross-fortnight tail continuity"].map(r=>(
              <div key={r} style={{fontSize:10,color:"#7fb3d3"}}>{r}</div>
            ))}
          </div>
        </div>
      </div>
      <button style={{...C.btnPrimary,fontSize:15,padding:"13px 44px"}} onClick={onGenerate}>⚡ Generate Roster</button>
    </div>
  );
}

// ─── STAFF TAB ───────────────────────────────────────────────
function StaffTab({staff,setStaff,toast}){
  const [editing,setEditing]=useState(null);
  const [form,setForm]=useState(null);
  const [filterCls,setFilterCls]=useState("ALL");
  const importRef=useRef();
  const blank={id:`s${Date.now()}`,name:"",cls:"RN",hrs:80,permNights:false,inCharge:false,fwaConditions:[],prefs:"",resign:null,gnpStart:null,leaveCard:[],gridLeave:{},requests:{}};

  function startEdit(s){setForm({...s,fwaConditions:[...(s.fwaConditions||[])]});setEditing(s.id);}
  function startNew(){setForm({...blank,id:`s${Date.now()}`});setEditing("new");}
  function saveForm(){
    if(!form.name.trim())return toast("Name required","err");
    if(editing==="new")setStaff(p=>[...p,form]); else setStaff(p=>p.map(s=>s.id===form.id?form:s));
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
          resign:r.ResignationDate||null,gnpStart:r.GNP_StartDate||null,
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
          {[["Full Name","name","text"],["Contracted Hours / Fortnight","hrs","number"],["GNP Start Date","gnpStart","date"],["Resignation Date","resign","date"]].map(([l,k,t])=>(
            <div key={k}><label style={C.lbl}>{l}</label>
              <input type={t} style={C.inp} value={f[k]||""} onChange={e=>sf(p=>({...p,[k]:t==="number"?Number(e.target.value)||0:(e.target.value||null)}))}/>
            </div>
          ))}
          <div><label style={C.lbl}>Classification</label>
            <select style={C.inp} value={f.cls} onChange={e=>sf(p=>({...p,cls:e.target.value}))}>
              {Object.entries(CLASSIFICATIONS).map(([k,v])=><option key={k} value={k}>{k} — {v.label}</option>)}
            </select>
          </div>
          <div><label style={C.lbl}>Known Preferences</label>
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

  const filtered=staff.filter(s=>filterCls==="ALL"||s.cls===filterCls);
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
            <div key={nurse.id} style={{...C.card,opacity:nurse.resigned?0.5:1}}>
              <div style={{display:"flex",alignItems:"flex-start",gap:10,marginBottom:8}}>
                <div style={{width:38,height:38,borderRadius:"50%",background:cls?.color+"33",color:cls?.color,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:14,flexShrink:0}}>
                  {nurse.name.split(" ").map(w=>w[0]).join("").slice(0,2)}
                </div>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,fontWeight:700,color:"#c8d8e8"}}>{nurse.name}</div>
                  <div style={{display:"flex",gap:5,flexWrap:"wrap",marginTop:3}}>
                    <span style={{fontSize:9,color:cls?.color,background:cls?.color+"22",borderRadius:8,padding:"1px 6px",fontWeight:700}}>{nurse.cls}</span>
                    <span style={{fontSize:9,color:"#64b5f6",background:"#1a3050",borderRadius:8,padding:"1px 6px"}}>{nurse.hrs}h/fn</span>
                    {nurse.permNights&&<span style={C.bdg("#1a237e","#7986cb")}>Perm Nights</span>}
                    {nurse.inCharge&&<span style={C.bdg("#1a3000","#aed581")}>In-Charge</span>}
                    {nurse.fwaConditions?.length>0&&<span style={C.bdg("#1b4000","#81c784")}>FWA ×{nurse.fwaConditions.length}</span>}
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
  const visible=staff.filter(s=>!s.resigned&&s.name.toLowerCase().includes(filterName.toLowerCase())&&(filterCls==="ALL"||s.cls===filterCls));

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
                  {staff.filter(s=>!s.resigned).sort((a,b)=>a.name.localeCompare(b.name)).map(s=><option key={s.id} value={s.id}>{s.name} ({s.cls})</option>)}
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
                  <span style={{fontWeight:700,color:"#c8d8e8",fontSize:13}}>{s.name}</span>
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
function NightPlanTab({staff,nightPlanData,setNightPlanData,toast}){
  const [year,setYear]=useState(new Date().getFullYear());
  const {groups,numGroups,staffPerBlock,permHoursPerFn}=autoComputeNightGroups(staff);
  const permStaff=staff.filter(s=>s.permNights);
  const gc=["#e53935","#fb8c00","#fdd835","#43a047","#1e88e5","#8e24aa","#00acc1","#f06292","#00897b","#6d4c41"];

  function autoGen(){
    const result=autoAssignNightPlan(staff,year);
    setNightPlanData(result);
    toast(`Night plan generated for ${year}: ${result.groups.length} groups`);
  }

  return(
    <div>
      <h2 style={C.pageH}>🌙 Night Shift Planner</h2>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:14,marginBottom:24}}>
        <div style={C.card}>
          <div style={C.cardH}>Coverage Calculation</div>
          <div style={{fontSize:11,color:"#7fb3d3",lineHeight:2.1}}>
            <div>Required: <b style={{color:"#a8dadc"}}>5 staff × 10h = 50h/shift</b></div>
            <div>14 nights/block: <b style={{color:"#a8dadc"}}>700h total</b></div>
            <div>Perm nights: <b style={{color:"#a8dadc"}}>{permHoursPerFn}h/fn</b></div>
            <div>Rotating needed: <b style={{color:"#a8dadc"}}>{Math.max(0,700-permHoursPerFn)}h/block</b></div>
            <div>Staff per block: <b style={{color:"#66bb6a"}}>{staffPerBlock}</b></div>
            <div>Groups: <b style={{color:"#66bb6a"}}>{numGroups}</b></div>
          </div>
        </div>
        <div style={C.card}>
          <div style={C.cardH}>Permanent Nights</div>
          {permStaff.length===0&&<div style={{fontSize:11,color:"#2a5070"}}>None configured</div>}
          {permStaff.map(s=>{const cls=CLASSIFICATIONS[s.cls];return(
            <div key={s.id} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
              <span style={{width:7,height:7,borderRadius:"50%",background:cls?.color}}/>
              <span style={{fontSize:11,color:"#c8d8e8"}}>{s.name}</span>
              <span style={{fontSize:9,color:cls?.color}}>{s.cls}</span>
              <span style={{fontSize:9,color:"#64b5f6",marginLeft:"auto"}}>{s.hrs}h</span>
            </div>
          );})}
        </div>
        <div style={C.card}>
          <div style={C.cardH}>Auto-Generate Plan</div>
          <div style={{fontSize:11,color:"#4a7fa0",marginBottom:12,lineHeight:1.7}}>Groups staff by classification and hours, ensures In-Charge coverage, rotates with 6–8 week gaps.</div>
          <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:14}}>
            <button style={C.btnSec} onClick={()=>setYear(y=>y-1)}>‹</button>
            <strong style={{color:"#a8dadc",fontSize:18,minWidth:48,textAlign:"center"}}>{year}</strong>
            <button style={C.btnSec} onClick={()=>setYear(y=>y+1)}>›</button>
          </div>
          <button style={C.btnPrimary} onClick={autoGen}>🔄 Auto-Generate {year}</button>
          {nightPlanData&&<div style={{fontSize:10,color:"#66bb6a",marginTop:8}}>✅ Plan active — {nightPlanData.groups?.length} groups</div>}
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
                      <span style={{fontSize:10,color:"#a8dadc"}}>{s.name}</span>
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
            {(nightPlanData.fns||[]).filter(fn=>new Date(fn.start).getFullYear()===year).map(fn=>{
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
