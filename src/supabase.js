// ─── SUPABASE CLIENT ─────────────────────────────────────────
// Replace the placeholder values below with your actual
// Supabase project URL and anon key from:
// supabase.com → your project → Settings → API

import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL  = import.meta.env.VITE_SUPABASE_URL  || "";
const SUPABASE_ANON = import.meta.env.VITE_SUPABASE_ANON || "";

export const supabase = (SUPABASE_URL && SUPABASE_ANON)
  ? createClient(SUPABASE_URL, SUPABASE_ANON)
  : null;

// ─── DATA ACCESS LAYER ───────────────────────────────────────
// All database operations go through these functions.
// Falls back to localStorage if Supabase is not configured
// (so the app still works during initial setup).

const LS = {
  get: (k, fb) => { try { const r = localStorage.getItem(k); return r ? JSON.parse(r) : fb; } catch { return fb; } },
  set: (k, v)  => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
};

// ── Staff ────────────────────────────────────────────────────
export async function loadStaff(fallback = []) {
  if (!supabase) return LS.get("wr3_staff", fallback);
  const { data, error } = await supabase.from("staff").select("*").order("name");
  if (error) { console.error("loadStaff:", error); return LS.get("wr3_staff", fallback); }
  // Parse JSON columns
  return (data || []).map(deserialiseStaff);
}

export async function saveStaff(staffArray) {
  LS.set("wr3_staff", staffArray); // always keep local copy as backup
  if (!supabase) return;
  const rows = staffArray.map(serialiseStaff);
  const { error } = await supabase.from("staff").upsert(rows, { onConflict: "id" });
  if (error) console.error("saveStaff:", error);
}

export async function deleteStaffMember(id) {
  if (!supabase) return;
  const { error } = await supabase.from("staff").delete().eq("id", id);
  if (error) console.error("deleteStaffMember:", error);
}

// ── Rosters ──────────────────────────────────────────────────
export async function loadRosters(fallback = {}) {
  if (!supabase) return LS.get("wr3_rosters", fallback);
  const { data, error } = await supabase.from("rosters").select("*").order("start_date");
  if (error) { console.error("loadRosters:", error); return LS.get("wr3_rosters", fallback); }
  const result = {};
  (data || []).forEach(row => {
    const key = `${row.start_date}_w${row.weeks}`;
    result[key] = deserialiseRoster(row);
  });
  return result;
}

export async function saveRoster(key, rosterData) {
  const current = LS.get("wr3_rosters", {});
  LS.set("wr3_rosters", { ...current, [key]: rosterData });
  if (!supabase) return;
  const row = serialiseRoster(key, rosterData);
  const { error } = await supabase.from("rosters").upsert(row, { onConflict: "id" });
  if (error) console.error("saveRoster:", error);
}

export async function updateRosterLock(key, locked) {
  const current = LS.get("wr3_rosters", {});
  if (current[key]) {
    current[key].locked = locked;
    LS.set("wr3_rosters", current);
  }
  if (!supabase) return;
  const id = keyToId(key);
  const { error } = await supabase.from("rosters").update({ locked, locked_at: locked ? new Date().toISOString() : null }).eq("id", id);
  if (error) console.error("updateRosterLock:", error);
}

// ── Night Plan ───────────────────────────────────────────────
export async function loadNightPlan(fallback = null) {
  if (!supabase) return LS.get("wr3_nightPlan", fallback);
  const { data, error } = await supabase.from("night_plan").select("*").order("created_at", { ascending: false }).limit(1);
  if (error) { console.error("loadNightPlan:", error); return LS.get("wr3_nightPlan", fallback); }
  if (!data || data.length === 0) return LS.get("wr3_nightPlan", fallback);
  return {
    plan:             data[0].plan,
    groupAssignments: data[0].group_assignments,
    groups:           data[0].groups,
    fns:              data[0].fns,
  };
}

export async function saveNightPlan(planData) {
  LS.set("wr3_nightPlan", planData);
  if (!supabase) return;
  const row = {
    id:                1, // single row, always upserted
    plan:              planData.plan,
    group_assignments: planData.groupAssignments,
    groups:            planData.groups,
    fns:               planData.fns,
    created_at:        new Date().toISOString(),
  };
  const { error } = await supabase.from("night_plan").upsert(row, { onConflict: "id" });
  if (error) console.error("saveNightPlan:", error);
}

// ─── SERIALISATION ───────────────────────────────────────────
function serialiseStaff(s) {
  return {
    id:              s.id,
    name:            s.name,
    cls:             s.cls,
    hrs:             s.hrs,
    perm_nights:     s.permNights,
    in_charge:       s.inCharge,
    fwa_conditions:  s.fwaConditions  || [],
    prefs:           s.prefs          || "",
    resign:          s.resign         || null,
    gnp_start:       s.gnpStart       || null,
    leave_card:      s.leaveCard      || [],
    grid_leave:      s.gridLeave      || {},
    requests:        s.requests       || {},
  };
}

function deserialiseStaff(row) {
  return {
    id:            row.id,
    name:          row.name,
    cls:           row.cls,
    hrs:           row.hrs,
    permNights:    row.perm_nights,
    inCharge:      row.in_charge,
    fwaConditions: row.fwa_conditions || [],
    prefs:         row.prefs          || "",
    resign:        row.resign         || null,
    gnpStart:      row.gnp_start      || null,
    leaveCard:     row.leave_card     || [],
    gridLeave:     row.grid_leave     || {},
    requests:      row.requests       || {},
  };
}

function keyToId(key) {
  // "2025-06-02_w2" → stable string id
  return key;
}

function serialiseRoster(key, r) {
  const [startDate, wPart] = key.split("_w");
  return {
    id:               key,
    start_date:       startDate,
    weeks:            parseInt(wPart),
    roster:           r.roster,
    leave_map:        r.leaveMap,
    hours_worked:     r.hoursWorked,
    hours_summary:    r.hoursSummary,
    warnings:         r.warnings,
    days:             r.days,
    tail:             r.tail,
    wknd_count_end:   r.wkndCountEnd,
    ado_inserted:     r.adoInserted,
    ado_total:        r.adoTotal,
    locked:           r.locked        || false,
    locked_at:        r.lockedAt      || null,
    generated_at:     r.generatedAt   || new Date().toISOString(),
  };
}

function deserialiseRoster(row) {
  return {
    roster:       row.roster,
    leaveMap:     row.leave_map,
    hoursWorked:  row.hours_worked,
    hoursSummary: row.hours_summary,
    warnings:     row.warnings      || [],
    days:         row.days          || [],
    tail:         row.tail          || {},
    wkndCountEnd: row.wknd_count_end|| {},
    adoInserted:  row.ado_inserted  || {},
    adoTotal:     row.ado_total     || 0,
    locked:       row.locked        || false,
    lockedAt:     row.locked_at     || null,
    generatedAt:  row.generated_at  || null,
    startDate:    row.start_date,
    weeks:        row.weeks,
  };
}
