/* =============================================================================
 * app.js  — orchestration
 * Holds the parsed model + global filters, computes aggregates, renders the
 * five dashboards, the executive report, and the data-quality panel.
 * Depends on: excel-loader.js, chart-generator.js, report-generator.js
 * ===========================================================================*/

var QBR = (window.QBR = window.QBR || {});
const APP = {
  model: null,
  files: [],    // accumulated workbook files this session ({name, blob}); dedup by name
  filters: { quarters: [], month: "ALL", org: "ALL", school: "ALL", secDefault: "ALL", usageCat: "ALL", sy: "ALL", status: "ALL", authMethod: "ALL" },
  activeTab: "dash-overview",
  expand: {},   // per-table "show all" flags (Option A)
  thiDrill: { open: false, band: "ALL" },   // Tenant Health Index per-tenant drill-down
  fd: { dataset: "risky", search: "", sort: null, page: 1, pageSize: 50, applyFilters: true }, // Full Data Explorer (Option B)
};

/* ---------- utilities ---------------------------------------------------- */
const $ = (id) => document.getElementById(id);
const sum = (arr, f) => arr.reduce((a, x) => a + (f(x) || 0), 0);
const avg = (arr) => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
const byDesc = (f) => (a, b) => (f(b) || 0) - (f(a) || 0);
function esc(s) { return String(s == null ? "" : s).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c])); }
function escAttr(s) { return esc(s).replace(/"/g, "&quot;"); }
function fmt(n, d = 0) { return n == null ? "—" : Number(n).toLocaleString(undefined, { maximumFractionDigits: d }); }
function tb(gb) { return gb == null ? "—" : (gb / 1024).toFixed(1) + " TB"; }

/* ---------- quarter multi-select helpers -------------------------------- */
// filters.quarters is an array of selected quarters; [] (or all four) == "All".
const ALL_Q = ["Q1", "Q2", "Q3", "Q4"];
function selQuarters() { const s = APP.filters.quarters; return (!s || !s.length) ? ALL_Q.slice() : ALL_Q.filter(q => s.includes(q)); }
function qActive(q) { const s = APP.filters.quarters; return !s || !s.length || s.includes(q); }
function isAllQ() { const n = (APP.filters.quarters || []).length; return n === 0 || n === 4; }
function quarterLabel() { return isAllQ() ? "All" : selQuarters().join(", "); }

/* ---------- filtering ---------------------------------------------------- */
// Apply global filters to a fact array. `hasMonth` = risky sheet (monthly).
function applyFilters(rows, opts = {}) {
  const f = APP.filters;
  return rows.filter(r => {
    // A row whose quarter cell was blank/unparseable (normQuarter -> null) is NOT
    // attributable to any quarter. It used to short-circuit past this test and so
    // matched EVERY quarter selection, double-counting that school in storage,
    // usage and Canva totals. It is now in scope only while the filter is "All".
    if (!r.quarter) { if (!isAllQ()) return false; }
    else if (!qActive(r.quarter)) return false;
    if (opts.hasMonth && f.month !== "ALL" && r.month && r.month !== f.month) return false;
    if (f.org !== "ALL" && (r.org || (APP.model.orgByKey && APP.model.orgByKey[r.key]) || "—") !== f.org) return false;
    if (f.school !== "ALL" && r.schoolRaw !== f.school) return false;
    return true;
  });
}

// resolve a school's security default from the security snapshot (by key)
function secDefaultFor(key) {
  const rec = APP.model.security.find(s => s.key === key);
  return rec ? rec.securityDefault : null;
}

// Collapse multiple quarterly rows for the same school into one representative
// row (the row with the max value of metricFn). Used when a specific quarter
// is not selected, so rankings/KPIs don't double-count a school across quarters.
function dedupeBySchool(rows, metricFn) {
  // NB: no single-quarter short-circuit. Selecting one quarter does not guarantee
  // one row per school (a sheet can legitimately repeat a school, and unattributed
  // rows reach here too), so the collapse must run in every scope. When rows really
  // are unique per school this is a no-op.
  const best = new Map();
  rows.forEach(r => {
    const cur = best.get(r.key);
    if (!cur || (metricFn(r) || 0) > (metricFn(cur) || 0)) best.set(r.key, r);
  });
  return [...best.values()];
}

/* ---------- aggregates for the executive report -------------------------- */
function computeAggregates() {
  const m = APP.model, f = APP.filters;
  const risky = applyFilters(m.risky, { hasMonth: true });
  const sec = m.security.filter(s => f.school === "ALL" || s.schoolRaw === f.school);
  const storage = applyFilters(m.storage);
  const usage = applyFilters(m.usage);
  const canva = applyFilters(m.canva);

  // risky per school (summed over selected period)
  const perSchool = {};
  risky.forEach(r => { if (r.risky != null) perSchool[r.schoolRaw] = (perSchool[r.schoolRaw] || 0) + r.risky; });
  const riskySchools = Object.entries(perSchool).map(([schoolRaw, risky]) => ({ schoolRaw, risky })).sort(byDesc(r => r.risky));

  // Domain health counted per SCHOOL (representative = latest month with a
  // status), NOT per monthly row — otherwise a 12-month sheet inflates the
  // totals (e.g. "458 of 589 records" instead of "71 of 110 schools").
  const healthPerSchool = new Map();
  risky.forEach(r => { if (!r.health) return; const e = healthPerSchool.get(r.key); if (!e || r.monthIdx >= e.mi) healthPerSchool.set(r.key, { h: r.health, mi: r.monthIdx }); });
  const health = {}; healthPerSchool.forEach(e => { health[e.h] = (health[e.h] || 0) + 1; });
  const secDefault = {}; sec.forEach(s => { secDefault[s.securityDefault] = (secDefault[s.securityDefault] || 0) + 1; });

  const usageVals = dedupeBySchool(usage.filter(u => u.usagePct != null), u => u.usagePct);
  const usageSorted = usageVals.slice().sort(byDesc(u => u.usagePct));
  const storVals = dedupeBySchool(storage.filter(s => s.usedGB != null), s => s.usedGB);
  const storSorted = storVals.slice().sort(byDesc(s => s.usedGB));
  const canvaVals = dedupeBySchool(canva.filter(c => c.users != null), c => c.users);

  return {
    totalSchools: m.master.size,
    health, secDefault, riskySchools,
    topRisk: riskySchools[0] || null,
    mfaEnabled: sec.filter(s => s.mfa === "YES").length,
    ssprEnabled: sec.filter(s => s.sspr === "YES").length,
    ssprGap: sec.filter(s => s.sspr !== "YES").length,
    usageAvg: avg(usageVals.map(u => u.usagePct)),
    topUsage: usageSorted[0] || null,
    lowUsage: usageSorted[usageSorted.length - 1] || null,
    storageTotalGB: storVals.length ? sum(storVals, s => s.usedGB) : null,
    storageAvgPct: avg(storVals.filter(s => s.pct != null).map(s => s.pct)),
    topStorage: storSorted[0] || null,
    canvaUsers: sum(canvaVals, c => c.users),
    canvaActiveSchools: canvaVals.filter(c => c.users > 0).length,
    topCanva: canvaVals.slice().sort(byDesc(c => c.users))[0] || null,
    ...umAndPmAgg(m, f),
  };
}

// User Management + Postmaster summary for the executive report
function umAndPmAgg(m, f) {
  const out = {};
  if (m.usermgmt.length) {
    const syList = [...new Set(m.usermgmt.map(u => u.sy))];
    const sy = (f.sy && f.sy !== "ALL" && syList.includes(f.sy)) ? f.sy : syList[0];
    const rows = m.usermgmt.filter(u => u.sy === sy);
    const upd = rows.filter(u => u.status === "Updated").length;
    out.umSY = sy; out.umTotal = rows.length; out.umUpdated = upd;
    out.umCompletion = rows.length ? upd / rows.length * 100 : 0;
  }
  if (m.postmaster.length) {
    // Count DISTINCT schools by their reputation across the loaded period. "Issues
    // detected" (the tracker's live problem state) and BAD (rated files) are the
    // actionable deliverability flags; both feed the executive narrative.
    const ACTIONABLE = { BAD: 1, "Issues detected": 1, LOW: 1, "Verify to see health": 1, MEDIUM: 1, HIGH: 1 };
    const bad = new Set(), high = new Set(), issues = new Set(), reported = new Set(), flagged = new Map();
    m.postmaster.forEach(p => {
      const r = p.reputation;
      if (r === "BAD") { bad.add(p.key); flagged.set(p.key, p.schoolRaw); }
      if (r === "HIGH") high.add(p.key);
      if (r === "Issues detected") { issues.add(p.key); flagged.set(p.key, p.schoolRaw); }
      if (ACTIONABLE[r]) reported.add(p.key);
    });
    out.pmBad = bad.size;
    out.pmHigh = high.size;
    out.pmIssues = issues.size;
    out.pmReported = reported.size;
    out.pmBadSchools = [...flagged.values()];   // schools flagged BAD or Issues detected
  }
  return out;
}

/* ---------- resizable table columns -------------------------------------- */
// Adds drag handles to every data table's header so users can widen/narrow
// columns. Attaches once per table (thead persists across tbody re-renders).
function makeTablesResizable() {
  document.querySelectorAll(".dash-panel table").forEach(tbl => {
    if (tbl.dataset.resizable) return;
    tbl.dataset.resizable = "1";
    tbl.classList.add("resizable");
    tbl.querySelectorAll("thead th").forEach(th => {
      const grip = document.createElement("span");
      grip.className = "col-resizer";
      th.appendChild(grip);
      let startX = 0, startW = 0;
      grip.addEventListener("mousedown", e => {
        e.preventDefault(); e.stopPropagation();
        startX = e.clientX; startW = th.offsetWidth; grip.classList.add("dragging");
        const move = ev => { const w = Math.max(48, startW + (ev.clientX - startX)); th.style.width = w + "px"; };
        const up = () => { grip.classList.remove("dragging"); document.removeEventListener("mousemove", move); document.removeEventListener("mouseup", up); };
        document.addEventListener("mousemove", move); document.addEventListener("mouseup", up);
      });
    });
  });
}

/* ---------- data-label spec helper (in-chart value labels) --------------- */
// Builds the dataLabels option consumed by the qbrLabels plugin. `mode` is one
// of donut|hbar|vbar|line. Merges into an existing opts object when given.
function DL(mode, labels, opts) { return Object.assign({}, opts, { dataLabels: { mode, labels } }); }
// value · % strings for a donut/pie (null for zero slices so they're skipped)
function pctLabels(values) { const t = values.reduce((a, b) => a + (b || 0), 0) || 1; return values.map(v => v ? `${fmt(v)} · ${Math.round(v / t * 100)}%` : null); }

/* ---------- KPI card helper --------------------------------------------- */
// v1.9.0 Phase 3 — quarter-over-quarter delta chips on page KPIs.
// renderAll() runs a silent "capture pass" of the time-scoped pages for the
// previous quarter (charts suppressed); kpi() records each numeric value under
// "<page>|<label>", and the real pass attaches a chip. Only labels listed in
// KPI_DELTA get a chip — snapshot metrics (security, GDAP, readiness, domain
// registration) and name-valued KPIs never do. The page key comes from
// withKpiPage() wrappers (bottom of this file), so no renderer changed.
const KPI_DELTA = {
  "dash-risky|Total Risky Users": { good: "down", pct: true },
  "dash-risky|Healthy Domains": { good: "up" },
  "dash-risky|Not Managed": { good: "down" },
  "dash-health|Healthy": { good: "up" },
  "dash-health|Possible Service Issues": { good: "down" },
  "dash-health|Incomplete Setup": { good: "down" },
  "dash-health|No Services Selected": { good: "down" },
  "dash-health|Not Managed": { good: "down" },
  "dash-health|Not Connected": { good: "down" },
  "dash-canva|Total Canva Users": { good: "up", pct: true },
  "dash-canva|Active Schools": { good: "up" },
  "dash-canva|Certificates Active": { good: "up" },
  "dash-usage|Average Usage %": { good: "up", d: 1, unit: " pts" },
  "dash-usage|Total Active (O365)": { good: "up", pct: true },
  "dash-storage|Total Storage": { good: "neutral", pct: true },
  "dash-storage|Average Usage %": { good: "neutral", d: 1, unit: " pts" },
  "dash-postmaster|Issues Detected": { good: "down" },
  "dash-postmaster|Verify to See Health": { good: "down" },
};
function kpiNum(v) {
  const m = String(v == null ? "" : v).replace(/<[^>]*>/g, "").replace(/,/g, "").match(/-?\d+(\.\d+)?/);
  return m ? +m[0] : NaN;
}
// Delta chip shared by kpi() and kpiHero(). goodDir: "up" | "down" | "neutral".
function deltaChip(cur, prev, goodDir, prevQ, o) {
  o = o || {};
  if (prevQ == null || !isFinite(cur) || !isFinite(prev)) return "";
  const diff = cur - prev;
  if (Math.abs(diff) < 1e-9) return `<span class="kpi-delta flat">No change<span class="kpi-delta-q"> vs ${prevQ}</span></span>`;
  const up = diff > 0;
  const cls = goodDir === "neutral" ? "flat" : ((goodDir === "up") === up ? "up" : "down");
  const arrow = up ? "▲" : "▼";
  const txt = o.pct
    ? (prev === 0 ? "new" : Math.abs(diff / prev * 100).toFixed(0) + "%")
    : fmt(Math.abs(diff), o.d || 0) + (o.unit || "");
  const sr = cls === "up" ? "improved" : cls === "down" ? "worsened" : "changed";
  return `<span class="kpi-delta ${cls}" title="${up ? "Up" : "Down"} ${txt} vs ${prevQ} (${sr})">${arrow} ${txt}<span class="kpi-delta-q"> vs ${prevQ}</span></span>`;
}
function kpi(label, value, tone) {
  const K = APP._kpi, page = APP._kpiPage;
  let chip = "";
  if (K && page) {
    const key = page + "|" + label, spec = KPI_DELTA[key];
    if (spec) {
      const n = kpiNum(value);
      if (K.mode === "capture") { if (isFinite(n)) K.prev[key] = n; }
      else if (K.mode === "apply") chip = deltaChip(n, K.prev[key], spec.good, K.prevQ, spec);
    }
  }
  // Name-valued KPIs (e.g. "Highest-Risk School") get a smaller type size so a
  // long school name doesn't balloon the card.
  const txt = !isFinite(kpiNum(value)) && String(value).replace(/<[^>]*>/g, "").trim().length > 3 ? " kpi-text" : "";
  return `<div class="col"><div class="kpi kpi-${tone || "blue"}${txt}">
    <div class="kpi-val">${value}</div><div class="kpi-lbl">${label}</div>${chip}</div></div>`;
}

/* ======================= DASHBOARD 0: EXECUTIVE OVERVIEW =======================
 * A single executive landing page: pulls only the headline summary from every
 * module. Reuses computeAggregates() (which already honours the global filters)
 * so nothing is re-sourced, then adds quarter-over-quarter deltas, a readiness
 * gauge, priority-flag counts, and quick-jump actions to the deep-dive tabs.
 * ============================================================================*/

// Which quarters actually appear in the fact data, in calendar order.
function quartersPresent() {
  const m = APP.model, set = new Set();
  [].concat(m.risky, m.canva, m.storage, m.usage).forEach(r => { if (r && r.quarter) set.add(r.quarter); });
  return ["Q1", "Q2", "Q3", "Q4"].filter(q => set.has(q));
}

// Aggregates for a specific quarter (org/school filters preserved, month reset)
// — used for the QoQ delta chips. Restores the live filters afterward.
function aggForQuarter(q) {
  const pq = APP.filters.quarters, pm = APP.filters.month;
  APP.filters.quarters = [q]; APP.filters.month = "ALL";
  const a = computeAggregates();
  APP.filters.quarters = pq; APP.filters.month = pm;
  return a;
}

// Hero KPI card with an optional QoQ delta chip. goodDir = which direction is
// favourable ("up" or "down"); colours the arrow green/red accordingly.
function kpiHero(label, value, tone, delta) {
  let chip = "";
  if (delta && delta.prevQ != null && isFinite(delta.cur) && isFinite(delta.prev) && Math.abs(delta.cur - delta.prev) > 1e-9)
    chip = deltaChip(delta.cur, delta.prev, delta.goodDir, delta.prevQ, delta);
  return `<div class="col"><div class="kpi kpi-hero kpi-${tone || "blue"}">
    <div class="kpi-val">${value}</div><div class="kpi-lbl">${label}</div>${chip}</div></div>`;
}

function goToTab(tabId) {
  const btn = document.querySelector('.tab-btn[data-tab="' + tabId + '"]');
  if (btn) btn.click();
}

/* ---------- Tenant Health Index (composite 0–100 per school) --------------
 * Blends six dimensions, each scored 0–100 or NULL when the school has no data
 * for it. Missing dimensions are EXCLUDED (weights renormalized), never zeroed —
 * a data gap must not read as a health failure. A school needs ≥2 dimensions to
 * be scored; sparser tenants are counted but not indexed. Honors the global
 * filters (quarter / org / school) via applyFilters, like the rest of Overview. */
// Scoring math, weights, and bands (min/label) now live in excel-loader.js as
// the pure, node-testable QBR.scoreTenantHealth / QBR.THI_WEIGHTS / QBR.THI_BANDS
// / QBR.thiScorers. Only the VIEW concern — band colors — stays here.
const THI_BAND_COLOR = { "Excellent": QBR.COLORS.green, "Healthy": QBR.COLORS.teal,
  "Attention Needed": QBR.COLORS.orange, "Critical": QBR.COLORS.red };
const thiBandColor = label => THI_BAND_COLOR[label] || QBR.COLORS.gray;
// Thin view-side wrapper: band descriptor (min/label from the engine) + its color.
function thiBand(v) { const b = QBR.thiBand(v); return { min: b.min, label: b.label, color: thiBandColor(b.label) }; }

// Latest month with a SCORABLE Postmaster reputation per school → { key: { rep, mi } }.
// Rows without a monthIdx rank lowest; ties go to the later row. Pure (exposed for tests).
function latestScorableReputation(rows) {
  const out = {};
  (rows || []).forEach(p => {
    if (!p || QBR.thiScorers.postmaster(p.reputation) == null) return;
    const mi = p.monthIdx == null ? -1 : p.monthIdx, e = out[p.key];
    if (!e || mi >= e.mi) out[p.key] = { rep: p.reputation, mi };
  });
  return out;
}
QBR._latestScorableReputation = latestScorableReputation;

function computeHealthIndex() {
  const m = APP.model, f = APP.filters;
  const risky = applyFilters(m.risky, { hasMonth: false });
  const usage = applyFilters(m.usage), storage = applyFilters(m.storage), pm = applyFilters(m.postmaster);
  const secRows = m.security.filter(s => (f.org === "ALL" || orgOf(s.key) === f.org) && (f.school === "ALL" || s.schoolRaw === f.school));
  // per-key dimension inputs (data-shaping stays here; scoring is delegated)
  const riskySum = {}, riskyHas = {}, domLatest = {};
  risky.forEach(r => {
    if (r.risky != null) { riskySum[r.key] = (riskySum[r.key] || 0) + r.risky; riskyHas[r.key] = true; }
    if (r.health) { const e = domLatest[r.key]; if (!e || r.monthIdx >= e.mi) domLatest[r.key] = { s: r.health, mi: r.monthIdx }; }
  });
  const usageByKey = {}; dedupeBySchool(usage.filter(u => u.usagePct != null), u => u.usagePct).forEach(u => usageByKey[u.key] = u.usagePct);
  const storByKey = {}; dedupeBySchool(storage.filter(s => s.pct != null), s => s.usedGB).forEach(s => storByKey[s.key] = s.pct);
  const secByKey = {}; secRows.forEach(s => { if (!secByKey[s.key]) secByKey[s.key] = s; });
  // pick the latest *scorable* reputation per school (skip unknowns), using the engine's own scorer.
  // 2026-10-03 fix: the map used to store the bare string and compare against an undefined
  // `.mi`, so the FIRST scorable month won instead of the latest. See latestScorableReputation().
  const pmLatest = latestScorableReputation(pm);
  const keys = new Set(); [risky, usage, storage, pm].forEach(a => a.forEach(r => keys.add(r.key))); secRows.forEach(s => keys.add(s.key));
  const perSchool = [];
  keys.forEach(key => {
    if (!key) return;
    const res = QBR.scoreTenantHealth({
      security: secByKey[key],
      risky: { sum: riskySum[key] || 0, has: !!riskyHas[key] },
      adoption: usageByKey[key],
      domain: domLatest[key] && domLatest[key].s,
      postmaster: pmLatest[key] && pmLatest[key].rep,
      storage: storByKey[key],
    });
    if (!res) return;                                  // too sparse to score
    perSchool.push({ key, name: (m.master.get(key) || {}).name || key, thi: res.thi, scores: res.scores });
  });
  const scored = perSchool.length;
  const dist = { "Excellent": 0, "Healthy": 0, "Attention Needed": 0, "Critical": 0 };
  perSchool.forEach(p => dist[QBR.thiBand(p.thi).label]++);
  return { index: scored ? Math.round(perSchool.reduce((a, p) => a + p.thi, 0) / scored) : null,
    scored, inScope: keys.size, dist, perSchool: perSchool.sort((a, b) => a.thi - b.thi) };
}

/* ---- Tenant Health Index drill-down (per-tenant report) -----------------
 * A ranked, weakest-first table of every scored tenant with its composite
 * index, band, and six sub-scores. The tenant's weakest present dimension is
 * highlighted so the "why" is visible at a glance. Band-filter chips (and a
 * click on a distribution segment) narrow the list; the whole panel is
 * collapsed by default so Overview stays uncluttered. Honors the same global
 * filters as the rest of Overview (it reads the already-filtered thi result). */
const THI_DIMS = [
  ["security", "Security"], ["risky", "Risky"], ["adoption", "Adoption"],
  ["domain", "Domain"], ["postmaster", "Email Rep."], ["storage", "Storage"],
];
function thiScoreCell(v, weak) {
  if (v == null) return '<td class="text-center text-muted" title="No data — excluded from the index (not scored as zero)">—</td>';
  const c = v >= 80 ? QBR.COLORS.green : v >= 60 ? QBR.COLORS.orange : QBR.COLORS.red;
  const style = `color:${c};background:${c}22` + (weak ? `;outline:2px solid ${c};outline-offset:-2px;font-weight:700` : "");
  return `<td class="text-center" style="${style}"${weak ? ' title="Weakest scored dimension for this tenant"' : ""}>${v}</td>`;
}
/* ---- Tenant Health detail pane (v1.9.0 Phase 3) ---------------------------
 * Azure-style right blade opened from a row of the Per-Tenant Health table:
 * the tenant's index + band, each dimension's sub-score with its weight, the
 * weakest dimension, and one-click jumps to the page behind each dimension.
 * Re-rendered after every renderAll() so it follows the filters; closes if the
 * tenant drops out of scope. */
const THI_DIM_PAGE = { security: "dash-sec", risky: "dash-risky", adoption: "dash-usage",
  domain: "dash-health", postmaster: "dash-postmaster", storage: "dash-storage" };
function openThiPane(key, opener) {
  APP.thiPane = { key, opener: opener || null };
  refreshThiPane(true);
}
function closeThiPane() {
  const pane = $("detail-pane"), scrim = $("dp-scrim");
  const op = APP.thiPane && APP.thiPane.opener;
  APP.thiPane = null;
  if (pane) pane.hidden = true;
  if (scrim) scrim.hidden = true;
  document.body.classList.remove("dp-open");
  if (op && document.contains(op)) op.focus();
  else if (op && op.dataset && op.dataset.thiKey) {
    const again = document.querySelector('.thi-row[data-thi-key="' + CSS.escape(op.dataset.thiKey) + '"]');
    if (again) again.focus();
  }
}
function refreshThiPane(focus) {
  const pane = $("detail-pane"), body = $("dp-body"), title = $("dp-title");
  if (!pane || !body || !APP.thiPane || !APP.model) { if (pane && !APP.thiPane) pane.hidden = true; return; }
  const thi = computeHealthIndex();
  const p = thi.perSchool.find(x => x.key === APP.thiPane.key);
  if (!p) { closeThiPane(); return; }
  const b = thiBand(p.thi), W = QBR.THI_WEIGHTS || {};
  let weakDim = null, weakVal = Infinity;
  THI_DIMS.forEach(([d]) => { const v = p.scores[d]; if (v != null && v < weakVal) { weakVal = v; weakDim = d; } });
  const tone = v => v == null ? "none" : v >= 80 ? "ok" : v >= 60 ? "warn" : "bad";
  title.textContent = p.name;
  const org = orgOf(p.key);
  body.innerHTML =
    `<div class="dp-score"><div class="dp-score-num" style="color:${b.color}">${p.thi}<span>/100</span></div>` +
      `<div><span class="badge" style="background:${b.color};color:#fff">${esc(b.label)}</span>` +
      `<div class="dp-meta">${org && org !== "—" ? esc(org) + " · " : ""}Tenant Health Index</div></div></div>` +
    (weakDim ? `<div class="dp-weak">Weakest dimension: <b>${esc((THI_DIMS.find(d => d[0] === weakDim) || [])[1])}</b> (${weakVal}/100)</div>` : "") +
    `<h3 class="dp-h">Dimension scores</h3><ul class="dp-dims">` +
    THI_DIMS.map(([d, name]) => {
      const v = p.scores[d], t = tone(v), w = W[d];
      return `<li class="dp-dim dp-${t}${d === weakDim ? " dp-weakest" : ""}">` +
        `<div class="dp-dim-top"><span class="dp-dim-name">${esc(name)}${w != null ? `<span class="dp-w">weight ${w}%</span>` : ""}</span>` +
        `<span class="dp-dim-val">${v == null ? "No data" : v}</span></div>` +
        `<div class="dp-bar" aria-hidden="true"><span style="width:${v == null ? 0 : v}%"></span></div>` +
        (v == null ? `<div class="dp-note">Excluded from the index — no data, not scored as zero.</div>` : "") +
        `<button type="button" class="dp-link" data-dp-tab="${THI_DIM_PAGE[d]}">Open ${esc(name === "Email Rep." ? "Google Postmaster" : name === "Risky" ? "Risky sign-ins" : name === "Adoption" ? "Microsoft 365 usage" : name === "Domain" ? "Tenant status" : name === "Security" ? "Security defaults" : name)}</button>` +
        `</li>`;
    }).join("") + `</ul>` +
    `<div class="dp-actions"><button type="button" class="btn btn-sm btn-primary" id="dp-s360">Open School 360</button> ` +
      `<button type="button" class="btn btn-sm btn-outline-secondary" id="dp-filter">Show this school across the dashboard</button></div>` +
    `<p class="dp-foot text-muted">Scores follow the current Quarter / Organization filters. Weights renormalize over the dimensions that have data.</p>`;
  body.querySelectorAll("[data-dp-tab]").forEach(x => x.addEventListener("click", () => { const t = x.dataset.dpTab; closeThiPane(); goToTab(t); }));
  const s3 = $("dp-s360");
  if (s3) s3.addEventListener("click", () => { const k = p.key; closeThiPane(); openSchool360(k); });
  const fb = $("dp-filter");
  if (fb) fb.addEventListener("click", () => {
    const inp = $("f-school"); if (!inp) return;
    inp.value = p.name; inp.dispatchEvent(new Event("change", { bubbles: true }));
    closeThiPane();
  });
  pane.hidden = false;
  const scrim = $("dp-scrim"); if (scrim) scrim.hidden = false;
  document.body.classList.add("dp-open");
  if (focus) { const c = $("dp-close"); if (c) c.focus(); }
}

function setThiBand(label) {
  APP.thiDrill.band = (APP.thiDrill.band === label) ? "ALL" : label;
  APP.thiDrill.open = true;                            // reveal detail when a chip/segment drives it
  if (APP.model) renderOverview();
}
function renderThiDrill(thi) {
  const host = $("thi-drill"), body = $("thi-drill-body"), bf = $("thi-band-filter"), tog = $("thi-drill-toggle");
  if (!host || !body) return;
  const open = !!APP.thiDrill.open;
  host.hidden = !open;
  if (tog) { tog.textContent = open ? "Hide detail ▲" : "Show detail ▾"; tog.setAttribute("aria-expanded", open ? "true" : "false"); }

  // Band-filter chips (always shown so the mix is visible even when collapsed).
  if (bf) {
    const chip = (band, label, count) => {
      const active = APP.thiDrill.band === band;
      const dis = band !== "ALL" && !count ? " disabled" : "";
      const col = band === "ALL" ? QBR.COLORS.blue : thiBandColor(band);
      const cls = active ? "btn btn-sm thi-band-btn" : "btn btn-sm btn-outline-secondary thi-band-btn";
      const style = active ? `background:${col};border-color:${col};color:#fff` : "";
      return `<button type="button" class="${cls}" data-band="${escAttr(band)}" style="${style}"${dis}>${esc(label)} <span class="opacity-75">${count}</span></button>`;
    };
    let html = chip("ALL", "All", thi.scored);
    QBR.THI_BANDS.forEach(b => { html += chip(b.label, b.label, thi.dist[b.label] || 0); });
    bf.innerHTML = html;
  }
  if (!open) return;                                   // don't build the table while collapsed

  if (thi.index == null) {
    body.innerHTML = '<tr><td colspan="10" class="text-muted">Not enough data to score tenant health for the current filter.</td></tr>';
    return;
  }
  const want = APP.thiDrill.band;
  const rows = thi.perSchool.map((p, i) => ({ p, rank: i + 1 }))
    .filter(x => want === "ALL" || QBR.thiBand(x.p.thi).label === want);
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="10" class="text-muted">No tenants in the "${esc(want)}" band for the current filter.</td></tr>`;
    return;
  }
  body.innerHTML = rows.map(({ p, rank }) => {
    const b = thiBand(p.thi);
    let weakDim = null, weakVal = Infinity;
    THI_DIMS.forEach(([d]) => { const v = p.scores[d]; if (v != null && v < weakVal) { weakVal = v; weakDim = d; } });
    const cells = THI_DIMS.map(([d]) => thiScoreCell(p.scores[d], d === weakDim)).join("");
    return `<tr class="thi-row" data-thi-key="${escAttr(p.key)}" tabindex="0" title="Open ${escAttr(p.name)} details"><td class="text-muted">${rank}</td><td>${esc(p.name)}</td>`
      + `<td class="text-center fw-bold" style="color:${b.color}">${p.thi}</td>`
      + `<td><span class="badge" style="background:${b.color};color:#fff">${esc(b.label)}</span></td>`
      + cells + "</tr>";
  }).join("");
}

function renderOverview() {
  const m = APP.model, f = APP.filters;
  const agg = computeAggregates();

  // ---- QoQ delta context ------------------------------------------------
  // Deltas are only meaningful when a single quarter is in view — comparing an
  // all-quarters total against one prior quarter is ambiguous, so on "All
  // Quarters" we show headline values with no delta chip.
  // Deltas anchor to the LATEST selected quarter vs the quarter before it.
  const qs = quartersPresent();
  const curQ = isAllQ() ? undefined : selQuarters().slice(-1)[0];
  const prevQ = (curQ && qs.includes(curQ)) ? qs[qs.indexOf(curQ) - 1] : undefined;
  const prevAgg = prevQ ? aggForQuarter(prevQ) : null;
  const totalRisky = agg.riskySchools.reduce((a, r) => a + (r.risky || 0), 0);

  const D = (cur, prev, goodDir, opts) => prevQ == null ? null
    : Object.assign({ cur, prev, prevQ, goodDir }, opts || {});

  // ---- Hero strip -------------------------------------------------------
  const healthy = agg.health["Healthy"] || 0;
  const stTB = agg.storageTotalGB == null ? null : agg.storageTotalGB / 1024;
  const prevStTB = prevAgg && prevAgg.storageTotalGB != null ? prevAgg.storageTotalGB / 1024 : null;
  $("kpi-overview").innerHTML =
    kpiHero("Managed Tenants", fmt(agg.totalSchools), "blue") +
    kpiHero("Healthy Domains", fmt(healthy), "green") +
    kpiHero("Total Risky Users", fmt(totalRisky), "red",
      D(totalRisky, prevAgg ? prevAgg.riskySchools.reduce((a, r) => a + (r.risky || 0), 0) : NaN, "down", { pct: true })) +
    kpiHero("Avg O365 Usage", agg.usageAvg == null ? "—" : agg.usageAvg.toFixed(1) + "%", "teal",
      D(agg.usageAvg, prevAgg ? prevAgg.usageAvg : NaN, "up", { d: 1 })) +
    kpiHero("Canva Users", fmt(agg.canvaUsers), "purple",
      D(agg.canvaUsers, prevAgg ? prevAgg.canvaUsers : NaN, "up", { pct: true })) +
    kpiHero("Total Storage", stTB == null ? "—" : stTB.toFixed(1) + " TB", "orange",
      D(stTB, prevStTB == null ? NaN : prevStTB, "up", { d: 1 }));

  // ---- Tenant Health Index (composite) ----------------------------------
  const thi = computeHealthIndex();
  const band = thi.index == null ? null : thiBand(thi.index);
  if ($("thi-score")) { $("thi-score").textContent = thi.index == null ? "—" : thi.index; $("thi-score").style.color = band ? band.color : ""; }
  if ($("thi-band")) $("thi-band").innerHTML = band ? `<span class="badge" style="background:${band.color};color:#fff">${band.label}</span>` : "";
  QBR.chart.stacked100("ch-ov-thi", QBR.THI_BANDS.map(b => ({ label: b.label, value: thi.dist[b.label] || 0, color: thiBandColor(b.label) })));
  if ($("thi-note")) $("thi-note").innerHTML = thi.index == null
    ? "Not enough data to score tenant health for the current filter."
    : `Composite of Security, Risky Sign-ins, O365 Adoption, Domain Health, Email Reputation &amp; Storage — scored for <b>${thi.scored}</b> of ${thi.inScope} tenants in scope. Missing dimensions are excluded, not zeroed. Bands: ≥95 Excellent · 80–94 Healthy · 60–79 Attention · &lt;60 Critical.`;
  renderThiDrill(thi);

  // ---- Period label -----------------------------------------------------
  const nice = v => (v == null || v === "ALL") ? "All" : v;
  $("ov-period").textContent = "Quarter: " + quarterLabel() +
    (f.org !== "ALL" ? " · " + f.org : "") + (f.school !== "ALL" ? " · " + f.school : "");
  $("ov-scope").textContent = `Portfolio summary across ${agg.totalSchools} managed tenants` +
    (prevQ ? ` · deltas vs ${prevQ}.` : ".");

  // ---- Risky monthly trend (same series as the Risky tab) ---------------
  const rowsAllMonths = applyFilters(m.risky, { hasMonth: false });
  const monthly = QBR.util.MONTHS.map(mo => ({ mo, v: sum(rowsAllMonths.filter(r => r.month === mo && r.risky != null), r => r.risky) }))
    .filter(x => rowsAllMonths.some(r => r.month === x.mo));
  QBR.chart.line("ch-ov-risky-trend", monthly.map(x => x.mo.slice(0, 3)),
    [{ label: "Risky Users", data: monthly.map(x => x.v), color: QBR.COLORS.red, fill: true }],
    { dataLabels: { mode: "line", labels: monthly.map(x => x.v ? fmt(x.v) : null) } });

  // ---- Security posture + Domain health: ranked status lists --------------
  // v1.9.0 Phase 3: the 100% bars became ranked, labelled lists (no legend to
  // decode, cards align row for row, each row drills into its page). The same
  // stacked100 charts are still drawn on hidden canvases so Export Images keeps
  // offering them as slides.
  const sd = agg.secDefault;
  const sdOrder = ["ENABLED", "DISABLED", "CONDITIONAL ACCESS", "NOT MANAGED"];
  const sdVals = sdOrder.map(k => sd[k] || 0);
  const sdSegLbl = ["Enabled", "Not Enabled", "Cond. Access", "Not Managed"];
  const sdSegCol = [QBR.COLORS.green, QBR.COLORS.red, QBR.COLORS.teal, QBR.COLORS.orange];
  QBR.chart.stacked100("ch-ov-secposture", sdSegLbl.map((l, i) => ({ label: l, value: sdVals[i], color: sdSegCol[i] })));
  renderStatusList("ov-secposture", sdOrder.map((k, i) => ({ label: SD_LABEL[k] || sdSegLbl[i], value: sdVals[i], color: sdSegCol[i],
    tone: STATUS_TONE_SD[k], go: { tab: "dash-sec", select: "f-secdef", value: k } })), "tenants");

  QBR.chart.stacked100("ch-ov-health", HEALTH_ORDER.map(h => ({ label: h, value: agg.health[h] || 0, color: HEALTH_COLOR[h] })));
  renderStatusList("ov-health", HEALTH_ORDER.map(h => ({ label: h, value: agg.health[h] || 0, color: HEALTH_COLOR[h],
    tone: STATUS_TONE_HEALTH[h], go: { tab: "dash-health", select: "health-status", value: h } })), "tenants");
  APP._agg = agg;                                       // reused by updateNavBadges()

  // ---- Readiness gauge (half doughnut) ---------------------------------
  const upd = agg.umUpdated || 0, umTot = agg.umTotal || 0, pend = Math.max(0, umTot - upd);
  const pct = umTot ? (upd / umTot * 100) : null;
  QBR.chart.doughnut("ch-ov-readiness", ["Updated", "Pending"], [upd, pend],
    [QBR.COLORS.green, "rgba(150,150,150,.25)"],
    { rotation: -90, circumference: 180, cutout: "72%",
      plugins: { legend: { display: false },
        qbrCenter: { value: pct == null ? "—" : pct.toFixed(0) + "%", label: "Updated" } } });

  // ---- Top 5 storage + Top 5 canva -------------------------------------
  const stor = dedupeBySchool(applyFilters(m.storage).filter(s => s.usedGB != null), s => s.usedGB)
    .sort(byDesc(s => s.usedGB)).slice(0, 5);
  QBR.chart.hbar("ch-ov-storage", stor.map(s => s.schoolRaw), stor.map(s => +(s.usedGB / 1024).toFixed(2)), QBR.COLORS.orange,
    { dataLabels: { mode: "hbar", labels: stor.map(s => (s.usedGB / 1024).toFixed(1) + " TB") } });
  const canva = dedupeBySchool(applyFilters(m.canva).filter(c => c.users != null), c => c.users)
    .sort(byDesc(c => c.users)).slice(0, 5);
  QBR.chart.hbar("ch-ov-canva", canva.map(c => c.schoolRaw), canva.map(c => c.users), QBR.COLORS.purple,
    { dataLabels: { mode: "hbar", labels: canva.map(c => fmt(c.users)) } });

  // ---- Top 5 risky table ------------------------------------------------
  const top = agg.riskySchools.slice(0, 5);
  $("tbl-ov-risky").innerHTML = top.map(x =>
    `<tr><td>${esc(x.schoolRaw)}</td><td class="text-end">${fmt(x.risky)}</td><td class="text-end">${sharePct(x.risky, totalRisky) || "—"}</td><td>${QBR.riskBadge(x.risky)}</td></tr>`
  ).join("") || '<tr><td colspan="4" class="text-muted">No data</td></tr>';
  setNote("ov-risky-note", top.length && totalRisky ? `These 5 schools account for <b>${sharePct(top.reduce((a, x) => a + (x.risky || 0), 0), totalRisky)}</b> of all ${fmt(totalRisky)} risky users.` : "");

  // ---- Priority flags ---------------------------------------------------
  const noAccess = agg.health["Not Managed"] || 0;
  const disabled = sd["DISABLED"] || 0;
  const badRep = (agg.pmBadSchools || []).length;   // BAD + "Issues detected" (the tracker's live problem state)
  const dregRows = (m.domainreg || [])
    .filter(d => f.org === "ALL" || orgOf(d.key) === f.org)
    .filter(d => f.school === "ALL" || d.schoolRaw === f.school);
  const expiring = dregRows.filter(d => d.group === "critical").length;
  const alert = (icon, tone, count, label, tab) =>
    `<button class="ov-alert ov-alert-${tone}" data-goto="${tab}">
       <span class="ov-alert-ico" aria-hidden="true">${icon}</span>
       <span class="ov-alert-num">${fmt(count)}</span>
       <span class="ov-alert-lbl">${label}</span></button>`;
  $("ov-alerts").innerHTML =
    alert("", "red", noAccess, "Not-Managed Tenants", "dash-health") +
    alert("", "orange", disabled, "Security Defaults Not Enabled", "dash-sec") +
    alert("", "red", expiring, "Domains Needing Action", "dash-domainreg") +
    alert("", "orange", badRep, "Email Reputation Issues", "dash-postmaster") +
    (() => { const cap = storageCapacityList(); const crit = cap.some(x => x.band.tone === "bad");
      return alert("", crit ? "red" : "orange", cap.length, `Tenants ≥${storWarn()}% Storage Capacity`, "dash-storage"); })();

  // ---- Quick actions ----------------------------------------------------
  const qa = [
    ["", "Security defaults", "dash-sec"],
    ["", "Risky sign-ins", "dash-risky"],
    ["", "Microsoft 365 usage", "dash-usage"],
    ["", "Executive report", "dash-report"],
  ].map(([i, l, t]) => `<button class="ov-qbtn" data-goto="${t}">${l}</button>`).join("");
  $("ov-quick").innerHTML = qa +
    `<button class="ov-qbtn ov-qbtn-primary" id="ov-export-deck">Export Deck</button>`;

  // wire jump + export buttons (re-created each render, so bind here)
  $("dash-overview").querySelectorAll("[data-goto]").forEach(b =>
    b.addEventListener("click", () => goToTab(b.getAttribute("data-goto"))));
  const ed = $("ov-export-deck");
  if (ed) ed.addEventListener("click", () => { const d = $("btn-deck"); if (d) d.click(); });
}

/* ---------- ranked status list (Overview posture / health cards) ----------
 * One row per status, sorted by count: dot · label · count · % · bar. Rows are
 * buttons that open the owning page with its status filter preset. Statuses
 * with no tenants collapse into one muted "None:" line. */
const STATUS_TONE_SD = { "ENABLED": "ok", "CONDITIONAL ACCESS": "ok", "DISABLED": "bad", "NOT MANAGED": "warn" };
const STATUS_TONE_HEALTH = { "Healthy": "ok", "Possible Service Issues": "warn", "Incomplete Setup": "warn",
  "No Services Selected": "neutral", "Not Managed": "bad", "Not Connected": "bad", "Not Applicable": "neutral",
  "End Contract": "neutral", "No Status": "neutral" };
function renderStatusList(hostId, rows, noun) {
  const host = $(hostId); if (!host) return;
  const total = rows.reduce((a, r) => a + (r.value || 0), 0);
  const live = rows.filter(r => r.value > 0).sort((a, b) => b.value - a.value);
  const none = rows.filter(r => !r.value).map(r => r.label);
  if (!total) { host.innerHTML = '<div class="stl-empty text-muted">No data for the current filter.</div>'; return; }
  const max = live[0].value;
  host.innerHTML = `<div class="stl-total"><b>${fmt(total)}</b> ${noun}</div>` +
    `<ul class="stl-list">` + live.map(r => {
      const pct = r.value / total * 100;
      const g = r.go || {};
      return `<li><button type="button" class="stl-row stl-${r.tone || "neutral"}" data-stl-tab="${escAttr(g.tab || "")}" data-stl-select="${escAttr(g.select || "")}" data-stl-value="${escAttr(g.value || "")}" title="Open ${escAttr(r.label)} on its page">` +
        `<span class="stl-dot" style="background:${r.color}" aria-hidden="true"></span>` +
        `<span class="stl-lbl">${esc(r.label)}</span>` +
        `<span class="stl-num">${fmt(r.value)}</span><span class="stl-pct">${pct < 1 ? "<1" : Math.round(pct)}%</span>` +
        `<span class="stl-bar" aria-hidden="true"><span style="width:${Math.max(2, r.value / max * 100).toFixed(1)}%;background:${r.color}"></span></span>` +
        `</button></li>`;
    }).join("") + `</ul>` +
    (none.length ? `<div class="stl-none">None: ${none.map(esc).join(", ")}</div>` : "");
  host.querySelectorAll(".stl-row").forEach(b => b.addEventListener("click", () => {
    const tab = b.dataset.stlTab, sel = b.dataset.stlSelect, val = b.dataset.stlValue;
    if (tab) goToTab(tab);
    const el = sel && $(sel);
    if (el && Array.from(el.options).some(o => o.value === val)) { el.value = val; el.dispatchEvent(new Event("change", { bubbles: true })); }
  }));
}

/* ---------- sidebar attention counts ----------------------------------------
 * Small count pills on the navigation items whose page has something that
 * needs action under the current filters. Counts, not "unread" markers, so they
 * stay until the underlying data changes. */
function updateNavBadges() {
  const agg = APP._agg, m = APP.model, f = APP.filters;
  if (!agg || !m) return;
  const dreg = (m.domainreg || []).filter(d => f.org === "ALL" || orgOf(d.key) === f.org)
    .filter(d => f.school === "ALL" || d.schoolRaw === f.school);
  const crit = QBR.THRESH && QBR.THRESH.CRITICAL != null ? QBR.THRESH.CRITICAL : 100;
  const counts = {
    "dash-risky": [agg.riskySchools.filter(r => r.risky > crit).length, "bad", "schools above " + crit + " risky users"],
    "dash-sec": [(agg.secDefault || {})["DISABLED"] || 0, "warn", "tenants with Security Defaults not enabled"],
    "dash-health": [(agg.health || {})["Not Managed"] || 0, "bad", "tenants not managed"],
    "dash-domainreg": [dreg.filter(d => d.group === "critical").length, "bad", "domains needing action"],
    "dash-postmaster": [(agg.pmBadSchools || []).length, "warn", "domains with email reputation issues"],
    "dash-inventory": (() => { try {
      const inv = m.inventory;
      if (!inv || !inv.tickets || !inv.tickets.length) return [0, "warn", ""];
      const n = inv.tickets.filter(t => !t.completed && !/completed|resolved|closed/i.test(String(t.status || ""))).length;
      return [n, n ? "warn" : "", n === 1 ? "open support ticket" : "open support tickets"];
    } catch (e) { return [0, "warn", ""]; } })(),
    "dash-storage": (() => { const c = storageCapacityList(); return [c.length, c.some(x => x.band.tone === "bad") ? "bad" : "warn", (c.length === 1 ? "tenant" : "tenants") + " at or above " + storWarn() + "% of storage capacity"]; })(),
  };
  document.querySelectorAll("#app-sidebar .sb-item[data-tab]").forEach(btn => {
    const old = btn.querySelector(".sb-count"); if (old) old.remove();
    const c = counts[btn.dataset.tab];
    if (!c || !c[0]) return;
    const span = document.createElement("span");
    span.className = "sb-count sb-count-" + c[1];
    span.title = fmt(c[0]) + " " + c[2];
    span.innerHTML = `<span class="vh">, </span>${fmt(c[0])}<span class="vh"> ${esc(c[2])}</span>`;
    btn.appendChild(span);
  });
}

// Risk-band pill without the emoji dot (report-generator.js keeps its own copy
// for the narrative; the CSS draws a shape-and-colour dot instead).
QBR.riskBadge = function (v) { const b = QBR.riskBand(v); return `<span class="risk-badge risk-${b.key}">${b.label}</span>`; };

/* ---------- share-of-total helpers (v1.9.0 · 2026-10-01) -------------------
 * "603 · 38%" labels and a concentration caption ("Top 3 = 56% · Top 10 = 78%
 * of 1,598 risky users across 109 schools") for ranked reports. */
function sharePct(v, total) {
  if (!total || v == null) return "";
  const p = v / total * 100;
  return p > 0 && p < 1 ? "<1%" : Math.round(p) + "%";
}
function shareLabel(text, v, total) { const sp = sharePct(v, total); return sp ? text + " · " + sp : text; }
function concentrationNote(valuesDesc, total, unit, noun) {
  const n = valuesDesc.length;
  if (!total || !n) return "";
  const top = k => valuesDesc.slice(0, k).reduce((a, b) => a + (b || 0), 0);
  const parts = [];
  if (n > 3) parts.push(`Top 3 = <b>${sharePct(top(3), total)}</b>`);
  if (n > 10) parts.push(`Top 10 = <b>${sharePct(top(10), total)}</b>`);
  const lead = valuesDesc[0] ? `The largest ${noun} alone holds <b>${sharePct(valuesDesc[0], total)}</b>. ` : "";
  return lead + (parts.length ? parts.join(" · ") + " of " : "Share of ") + `${unit} across ${fmt(n)} ${noun}${n === 1 ? "" : "s"}.`;
}
function setNote(id, html) { const el = $(id); if (el) el.innerHTML = html || ""; }

/* ---------- storage capacity flags (v1.9.0 · 2026-10-01) -------------------
 * Each tenant's LATEST storage snapshot in the current filters (same snapshot
 * the Per-Tenant detail shows) is flagged when Used / Pooled capacity reaches
 * the warning (80%) or critical (90%) level. Thresholds sit on QBR.THRESH so
 * they can be changed in one place alongside the risk thresholds. */
if (QBR.THRESH) {
  if (QBR.THRESH.STORAGE_WARN == null) QBR.THRESH.STORAGE_WARN = 80;
  if (QBR.THRESH.STORAGE_CRIT == null) QBR.THRESH.STORAGE_CRIT = 90;
}
const storWarn = () => (QBR.THRESH && QBR.THRESH.STORAGE_WARN) || 80;
const storCrit = () => (QBR.THRESH && QBR.THRESH.STORAGE_CRIT) || 90;
const CAP_KPI_LABEL = () => `Near Capacity (≥${storWarn()}%)`;
KPI_DELTA["dash-storage|" + CAP_KPI_LABEL()] = { good: "down" };
function storagePct(r) {
  if (!r) return null;
  if (r.pct != null) return r.pct;
  return (r.totalGB > 0 && r.usedGB != null) ? r.usedGB / r.totalGB * 100 : null;
}
function capacityBand(p) {
  if (p == null) return null;
  if (p >= storCrit()) return { label: "At capacity", tone: "bad" };
  if (p >= storWarn()) return { label: "Near capacity", tone: "warn" };
  return null;
}
function capBadge(p) { const b = capacityBand(p); return b ? ` <span class="cap-badge cap-${b.tone}">${b.label}</span>` : ""; }
function storageCapacityList() {
  if (!APP.model) return [];
  const QIDX = { Q1: 0, Q2: 1, Q3: 2, Q4: 3 };
  const latest = new Map();
  applyFilters(APP.model.storage).forEach(r => {
    const pc = storagePct(r);
    if (pc == null || !(r.totalGB > 0)) return;
    const e = latest.get(r.key), qi = QIDX[r.quarter] == null ? -1 : QIDX[r.quarter];
    if (!e || qi > (QIDX[e.quarter] == null ? -1 : QIDX[e.quarter])) latest.set(r.key, r);
  });
  return [...latest.values()].map(r => ({ r, pct: storagePct(r), band: capacityBand(storagePct(r)) }))
    .filter(x => x.band).sort((a, b) => b.pct - a.pct);
}

/* ======================= DASHBOARD 1: RISKY USERS ======================= */
function renderRisky() {
  const m = APP.model;
  const rows = applyFilters(m.risky, { hasMonth: true });
  const perSchool = {};
  rows.forEach(r => { if (r.risky != null) perSchool[r.schoolRaw] = (perSchool[r.schoolRaw] || 0) + r.risky; });
  const schools = Object.entries(perSchool).map(([s, v]) => ({ s, v }));
  const schoolsDesc = schools.slice().sort(byDesc(x => x.v));
  const top = schoolsDesc.slice(0, 10);
  // Lowest-risk = smallest NON-ZERO totals. Schools with 0 risky users would
  // otherwise fill the bottom 10 with blank (zero-length) bars; we surface the
  // count of those fully-clean schools in a caption instead.
  const zeroCount = schools.filter(x => x.v === 0).length;
  const low = schools.slice().filter(x => x.v > 0).sort((a, b) => a.v - b.v).slice(0, 10);
  const totalRisky = sum(rows.filter(r => r.risky != null), r => r.risky);
  // Domain health counted per SCHOOL (latest month with a status), not per
  // monthly row — otherwise the 12-month sheet inflates the KPI.
  const healthPerSchool = new Map();
  rows.forEach(r => { if (!r.health) return; const e = healthPerSchool.get(r.key); if (!e || r.monthIdx >= e.mi) healthPerSchool.set(r.key, { h: r.health, mi: r.monthIdx }); });
  const health = {}; healthPerSchool.forEach(e => { health[e.h] = (health[e.h] || 0) + 1; });

  $("kpi-risky").innerHTML =
    kpi("Total Risky Users", fmt(totalRisky), "red") +
    kpi("Healthy Domains", fmt(health["Healthy"] || 0), "green") +
    kpi("Not Managed", fmt(health["Not Managed"] || 0), "orange") +
    kpi("Highest-Risk School", top[0] ? esc(top[0].s) : "—", "blue");

  // monthly trend — always spans all 12 months (ignores the Month sub-filter,
  // which only scopes KPIs/tables). Still respects Quarter / Org / School.
  const rowsAllMonths = applyFilters(m.risky, { hasMonth: false });

  // ---- Security Defaults DISABLED vs ENABLED — monthly total risky users ----
  // Every risky tenant is folded into exactly TWO buckets so the two lines add
  // up to the Monthly Trend total:
  //   ENABLED (protected)   = Security Defaults ENABLED  OR  Conditional Access
  //   DISABLED (unprotected)= everything else (Disabled, Not Managed, Unknown,
  //                           or no matching SECURITY_DATA record)
  const secByKey = {}; m.security.forEach(s => { secByKey[s.key] = s.securityDefault; });
  const isProtected = key => { const sd = secByKey[key]; return sd === "ENABLED" || sd === "CONDITIONAL ACCESS"; };
  const disSeries = [], enSeries = [];
  QBR.util.MONTHS.forEach(mo => {
    const monthRows = rowsAllMonths.filter(r => r.month === mo && r.risky != null);
    enSeries.push(sum(monthRows.filter(r => isProtected(r.key)), r => r.risky));
    disSeries.push(sum(monthRows.filter(r => !isProtected(r.key)), r => r.risky));
  });
  const monthLabels = QBR.util.MONTHS.map(mo => mo[0] + mo.slice(1).toLowerCase());
  const disLbl = disSeries.map(v => v ? fmt(v) : null), enLbl = enSeries.map(v => v ? fmt(v) : null);
  // Card-local posture filter: Both / Protected only / Unprotected only.
  // Colours: DISABLED/unprotected = red, ENABLED/protected = green.
  const secView = ($("risky-secdef-view") && $("risky-secdef-view").value) || "BOTH";
  const secSeries = [];
  if (secView !== "ENABLED") secSeries.push({ label: "Not Enabled + Not Managed", data: disSeries, color: QBR.COLORS.red, lbl: disLbl });
  if (secView !== "DISABLED") secSeries.push({ label: "Enabled + Conditional Access", data: enSeries, color: QBR.COLORS.green, lbl: enLbl });
  QBR.chart.line("ch-risky-secdef", monthLabels,
    secSeries.map(s => ({ label: s.label, data: s.data, color: s.color })),
    { dataLabels: { mode: "line", labels: secSeries.map(s => s.lbl), colored: true } });
  const disTot = sum(disSeries.map(v => ({ v })), x => x.v), enTot = sum(enSeries.map(v => ({ v })), x => x.v);
  if (m.security.length && (disTot || enTot)) {
    const pk = disSeries.indexOf(Math.max(...disSeries));
    $("risky-secdef-note").innerHTML =
      `Tenants with Security Defaults <b>Not Enabled</b> (or Not Managed) generated <b>${fmt(disTot)}</b> risky-user detections over the period vs <b>${fmt(enTot)}</b> for tenants with them <b>Enabled</b> (or Conditional Access)` +
      (disTot > enTot && enTot >= 0 ? ` — a <b>${enTot ? Math.round((disTot - enTot) / enTot * 100) + "% higher" : "far higher"}</b> concentration of identity risk` : "") +
      (disSeries[pk] ? `, peaking in <b>${monthLabels[pk]}</b> (${fmt(disSeries[pk])})` : "") +
      `. The two lines now reconcile to the Monthly Trend total.`;
  } else {
    $("risky-secdef-note").textContent = "Needs both SECURITY_DATA (Security Defaults) and RISKY_USERS_AND_DOMAIN loaded to compare postures.";
  }
  const monthly = QBR.util.MONTHS.map(mo => {
    const v = sum(rowsAllMonths.filter(r => r.month === mo && r.risky != null), r => r.risky);
    return { mo, v };
  }).filter(x => rowsAllMonths.some(r => r.month === x.mo));
  QBR.chart.line("ch-risky-monthly", monthly.map(x => x.mo.slice(0, 3)),
    [{ label: "Risky Users", data: monthly.map(x => x.v), color: QBR.COLORS.red, fill: true }],
    DL("line", monthly.map(x => x.v ? fmt(x.v) : null)));

  // quarterly comparison
  const quarters = ["Q1", "Q2", "Q3", "Q4"];
  const qData = quarters.map(q => sum(m.risky.filter(r => r.quarter === q && r.risky != null &&
    (APP.filters.school === "ALL" || r.schoolRaw === APP.filters.school)), r => r.risky));
  QBR.chart.bar("ch-risky-quarter", quarters, [{ label: "Risky Users", data: qData, backgroundColor: quarters.map(q => QBR.QUARTER_COLORS[q]) }],
    DL("vbar", qData.map(v => v ? fmt(v) : null), { plugins: { legend: { display: false } } }));

  const leftAlign = { scales: { y: { ticks: { crossAlign: "far" } } } };
  const riskyTot = schools.reduce((a, x) => a + (x.v || 0), 0);
  QBR.chart.hbar("ch-risky-top", top.map((x, i) => `${i + 1}. ${x.s}`), top.map(x => x.v), QBR.COLORS.red, DL("hbar", top.map(x => shareLabel(fmt(x.v), x.v, riskyTot)), leftAlign));
  setNote("risky-top-note", concentrationNote(schoolsDesc.filter(x => x.v > 0).map(x => x.v), riskyTot, fmt(riskyTot) + " risky users", "school"));
  QBR.chart.hbar("ch-risky-low", low.map((x, i) => `${i + 1}. ${x.s}`), low.map(x => x.v), QBR.COLORS.green, DL("hbar", low.map(x => fmt(x.v)), leftAlign));
  $("risky-low-note").textContent = zeroCount
    ? `Showing the ${low.length} lowest schools with at least 1 risky user · ${zeroCount} school${zeroCount === 1 ? "" : "s"} had 0 risky sign-ins (fully clean).`
    : (low.length ? "Lowest risky-user counts among schools with activity." : "No risky-user data for the current filter.");

  // by organization
  const org = {}; rows.forEach(r => { if (r.risky != null) { const o = r.org || orgOf(r.key); org[o] = (org[o] || 0) + r.risky; } });
  const orgE = Object.entries(org).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const orgTot = Object.values(org).reduce((a, b) => a + b, 0);
  QBR.chart.bar("ch-risky-org", orgE.map(x => x[0]), [{ label: "Risky Users", data: orgE.map(x => x[1]), backgroundColor: QBR.paletteColors(orgE.length) }],
    DL("vbar", orgE.map(x => sharePct(x[1], orgTot) || fmt(x[1])), { plugins: { legend: { display: false } } }));
  setNote("risky-org-note", orgE.length ? `Bar labels show each organization's share of ${fmt(orgTot)} risky users. ` +
    `${esc(orgE[0][0])} leads with ${fmt(orgE[0][1])} (${sharePct(orgE[0][1], orgTot)}).` : "");

  // drill-down table w/ risk bands + auto-suggested action (expandable)
  const flagged = expandSlice("risky", schoolsDesc, 10).map(x => {
    const band = QBR.riskBand(x.v);
    return `<tr><td>${esc(x.s)}</td><td class="text-end">${fmt(x.v)}</td><td class="text-end">${sharePct(x.v, riskyTot) || "—"}</td>
      <td>${QBR.riskBadge(x.v)}</td><td class="small">${esc(band.action)}</td></tr>`;
  }).join("");
  $("tbl-risky").innerHTML = flagged || '<tr><td colspan="5" class="text-muted">No data</td></tr>';
  setExpandBtn("risky", 10);
}

/* ======================= DOMAIN HEALTH ======================= */
const HEALTH_ORDER = ["Healthy", "Possible Service Issues", "Incomplete Setup", "No Services Selected", "Not Managed", "Not Connected", "Not Applicable", "End Contract", "No Status"];
const HEALTH_COLOR = {
  "Healthy": QBR.COLORS.green, "Possible Service Issues": QBR.COLORS.orange,
  "Incomplete Setup": QBR.COLORS.yellow, "No Services Selected": QBR.COLORS.gray,
  "Not Managed": QBR.COLORS.red, "Not Connected": "#A4262C",
  "Not Applicable": "#8A8886", "End Contract": QBR.COLORS.purple, "No Status": "#C8C6C4",
};
const HEALTH_TONE = { "Healthy": "green", "Possible Service Issues": "orange", "Incomplete Setup": "teal",
  "No Services Selected": "purple", "Not Managed": "red", "Not Connected": "red",
  "Not Applicable": "blue", "End Contract": "purple", "No Status": "blue" };

function renderDomainHealth() {
  const m = APP.model, f = APP.filters;
  const rows = applyFilters(m.risky, { hasMonth: true }); // respects quarter/month/org/school
  // one representative record per school = most recent month with a non-null status
  const perSchool = new Map();
  rows.forEach(r => {
    let e = perSchool.get(r.key);
    if (!e) { e = { schoolRaw: r.schoolRaw, org: r.org, status: null, statusMonth: null, mi: -1, risky: 0, remarks: null }; perSchool.set(r.key, e); }
    if (r.schoolRaw.length > e.schoolRaw.length) e.schoolRaw = r.schoolRaw;
    if (r.org && !e.org) e.org = r.org;
    if (r.risky != null) e.risky += r.risky;
    if (r.errorCause && !e.remarks) e.remarks = r.errorCause;
    if (r.health && r.monthIdx >= e.mi) { e.status = r.health; e.statusMonth = r.month; e.mi = r.monthIdx; if (r.errorCause) e.remarks = r.errorCause; }
  });
  // Scope: Domain Status lists only tenants that actually have a RISKY_USERS_AND_DOMAIN
  // record (the domain-health source). Schools that reach the Master School Dimension
  // solely through other sheets (e.g. Canva-only tenants) are intentionally NOT injected
  // here — they have no domain to report and would otherwise show as "Unassigned / No
  // Status". The Overview health donut already counts from RISKY only, so this keeps the
  // two views consistent. A tenant that IS in RISKY but has a blank status this period
  // still appears, as "No Status".
  perSchool.forEach(e => { if (!e.status) e.status = "No Status"; });   // in-risky but blank status
  const list = [...perSchool.values()].sort((a, b) => a.schoolRaw.localeCompare(b.schoolRaw));

  // Counts across ALL categories present (HEALTH_ORDER first, any strays appended).
  const cats = HEALTH_ORDER.slice();
  list.forEach(e => { if (!cats.includes(e.status)) cats.push(e.status); });
  const counts = {}; cats.forEach(c => counts[c] = 0);
  list.forEach(e => counts[e.status]++);
  const nonZero = cats.filter(c => counts[c] > 0);

  // KPI cards — one per non-empty status (so Sum reconciles to the total).
  $("kpi-health").innerHTML = nonZero.map(h => kpi(h, fmt(counts[h]), HEALTH_TONE[h] || "blue")).join("");

  // Reconciliation caption (Audit Control 1 & 2).
  const total = list.length;
  $("health-recon").innerHTML = `<b>${fmt(total)}</b> domain-tracked tenants in scope = ` +
    nonZero.map(h => `${fmt(counts[h])} ${esc(h)}`).join(" + ") +
    `. Scope = tenants with a RISKY_USERS_AND_DOMAIN record; Canva-only tenants are excluded. Every tenant is counted in exactly one status.`;

  // distribution
  const hSorted = nonZero.map(h => ({ h, c: counts[h] })).sort((a, b) => b.c - a.c);
  QBR.chart.hbar("ch-health-dist", hSorted.map(x => x.h), hSorted.map(x => x.c),
    hSorted.map(x => HEALTH_COLOR[x.h] || QBR.COLORS.gray), DL("hbar", hSorted.map(x => fmt(x.c))));

  // health by organization (stacked)
  const orgs = [...new Set(list.map(e => e.org || orgOf(e.key)))].sort();
  const datasets = nonZero.map(h => ({
    label: h, backgroundColor: HEALTH_COLOR[h] || QBR.COLORS.gray,
    data: orgs.map(o => list.filter(e => (e.org || orgOf(e.key)) === o && e.status === h).length),
  }));
  QBR.chart.bar("ch-health-org", orgs, datasets,
    { scales: { x: { stacked: true }, y: { stacked: true } }, dataLabels: { mode: "stack" } });

  // table (all schools) — stored for live search filtering
  APP._healthList = list;
  renderHealthTable("");
}

function renderHealthTable(q) {
  const sf = ($("health-status") && $("health-status").value) || "ALL";
  const list = (APP._healthList || []).filter(e =>
    (!q || e.schoolRaw.toLowerCase().includes(q.toLowerCase())) &&
    (sf === "ALL" || e.status === sf));
  const badge = (s) => {
    if (!s) return '<span class="text-muted">—</span>';
    const cls = s === "Healthy" ? "bg-success"
      : (s === "Not Managed" || s === "Not Connected") ? "bg-danger"
      : s === "Possible Service Issues" ? "bg-warning text-dark" : "bg-secondary";
    return `<span class="badge ${cls}">${esc(s)}</span>`;
  };
  // Organization label: blank org → "Unassigned"; else the org name.
  const orgLabel = (e) => e.org ? esc(e.org) : '<span class="text-muted">Unassigned</span>';
  $("tbl-health").innerHTML = list.map(e =>
    `<tr><td>${esc(e.schoolRaw)}</td><td>${orgLabel(e)}</td><td>${badge(e.status)}</td>
     <td class="small">${e.remarks ? esc(e.remarks) : '<span class="text-muted">—</span>'}</td>
     <td class="text-end">${fmt(e.risky)}</td><td class="small text-muted">${e.statusMonth ? esc(e.statusMonth[0] + e.statusMonth.slice(1).toLowerCase()) : "—"}</td></tr>`
  ).join("") || '<tr><td colspan="6" class="text-muted">No schools match.</td></tr>';
}

/* ============ DASHBOARD 2: SECURITY DEFAULTS vs RISKY USERS ============ */
// Display labels for the security-default states. The MODEL keeps the canonical
// uppercase values (ENABLED/DISABLED/…) for matching & filtering; these are only
// what the user sees. Neutral wording: "Not Enabled" rather than "Disabled".
const SD_LABEL = { "ENABLED": "Enabled", "DISABLED": "Not Enabled",
  "CONDITIONAL ACCESS": "Conditional Access", "NOT MANAGED": "Not Managed", "UNKNOWN": "Unknown" };
// Does a school's canonical method set satisfy the Auth-method filter? Handles
// specific methods ("includes X") and the posture presets. Empty method sets
// (no method data / "NO ACCESS") only satisfy "ALL".
function matchAuthMethod(methods, want) {
  if (want === "ALL") return true;
  const set = methods || [];
  if (want === "__PHISH") return set.some(x => QBR.AUTH_PHISH_RESISTANT[x]);
  if (want === "__WEAKONLY") return set.length > 0 && !set.some(x => QBR.AUTH_STRONG[x]);
  if (want === "__OTPONLY") return set.length === 1 && set[0] === "Email OTP";
  if (want === "__NOMFA") return set.length > 0 && set.indexOf("MFA") < 0;
  return set.indexOf(want) >= 0;   // a specific method
}

function renderSecurity() {
  const m = APP.model;
  const sec = m.security.filter(s => APP.filters.school === "ALL" || s.schoolRaw === APP.filters.school)
    .filter(s => APP.filters.secDefault === "ALL" || s.securityDefault === APP.filters.secDefault);
  const dist = {}; sec.forEach(s => dist[s.securityDefault] = (dist[s.securityDefault] || 0) + 1);

  // risky per school key over selected quarter
  const risky = applyFilters(m.risky, { hasMonth: true });
  const riskyByKey = {};
  risky.forEach(r => { if (r.risky != null) riskyByKey[r.key] = (riskyByKey[r.key] || 0) + r.risky; });
  const joined = sec.map(s => ({ ...s, riskyTotal: riskyByKey[s.key] || 0 }));

  const enabledRisk = avg(joined.filter(j => j.securityDefault === "ENABLED").map(j => j.riskyTotal));
  const disabledRisk = avg(joined.filter(j => j.securityDefault === "DISABLED").map(j => j.riskyTotal));

  $("kpi-sec").innerHTML =
    kpi("Enabled", fmt(dist["ENABLED"] || 0), "green") +
    kpi("Not Enabled", fmt(dist["DISABLED"] || 0), "red") +
    kpi("Conditional Access", fmt(dist["CONDITIONAL ACCESS"] || 0), "blue") +
    kpi("Not Managed", fmt(dist["NOT MANAGED"] || 0), "orange");

  // Monthly ENABLED vs DISABLED risky users: join each monthly risky row to its
  // school's security-default status (snapshot) and sum per month by posture.
  const secByKey = {}; m.security.forEach(s => { secByKey[s.key] = s.securityDefault; });
  const riskyAllMonths = applyFilters(m.risky, { hasMonth: false });
  const enSeries = [], disSeries = [];
  QBR.util.MONTHS.forEach(mo => {
    const monthRows = riskyAllMonths.filter(r => r.month === mo && r.risky != null);
    enSeries.push(sum(monthRows.filter(r => secByKey[r.key] === "ENABLED"), r => r.risky));
    disSeries.push(sum(monthRows.filter(r => secByKey[r.key] === "DISABLED"), r => r.risky));
  });
  QBR.chart.line("ch-sec-endis", QBR.util.MONTHS.map(x => x.slice(0, 3)), [
    { label: "Enabled", data: enSeries, color: QBR.COLORS.green },
    { label: "Not Enabled", data: disSeries, color: QBR.COLORS.red },
  ]);

  // Average risky users by security posture (bar).
  const cmpData = [
    avg(joined.filter(j => j.securityDefault === "ENABLED").map(j => j.riskyTotal)) || 0,
    avg(joined.filter(j => j.securityDefault === "DISABLED").map(j => j.riskyTotal)) || 0,
    avg(joined.filter(j => j.securityDefault === "CONDITIONAL ACCESS").map(j => j.riskyTotal)) || 0,
    avg(joined.filter(j => j.securityDefault === "NOT MANAGED").map(j => j.riskyTotal)) || 0,
  ];
  QBR.chart.bar("ch-sec-compare", ["Enabled", "Not Enabled", "Conditional Access", "Not Managed"],
    [{ label: "Avg Risky Users", data: cmpData, backgroundColor: [QBR.COLORS.green, QBR.COLORS.red, QBR.COLORS.blue, QBR.COLORS.orange] }],
    DL("vbar", cmpData.map(v => v ? v.toFixed(1) : null), { plugins: { legend: { display: false } } }));

  QBR.chart.pie("ch-sec-dist", Object.keys(dist).map(k => SD_LABEL[k] || k), Object.values(dist),
    [QBR.COLORS.green, QBR.COLORS.red, QBR.COLORS.blue, QBR.COLORS.orange, QBR.COLORS.gray],
    DL("donut", pctLabels(Object.values(dist))));

  // control effectiveness: MFA / SSPR / OTP / SMS counts
  const ctl = { MFA: sec.filter(s => s.mfa === "YES").length, SSPR: sec.filter(s => s.sspr === "YES").length,
    "Email OTP": sec.filter(s => s.emailOtp === "YES").length, SMS: sec.filter(s => s.sms === "YES").length };
  QBR.chart.bar("ch-sec-control", Object.keys(ctl), [{ label: "Schools", data: Object.values(ctl), backgroundColor: QBR.paletteColors(Object.keys(ctl).length) }],
    DL("vbar", Object.values(ctl).map(v => fmt(v)), { plugins: { legend: { display: false } } }));

  // disabled schools ranking by risky
  const disRank = joined.filter(j => j.securityDefault === "DISABLED").sort(byDesc(j => j.riskyTotal)).slice(0, 10);
  QBR.chart.hbar("ch-sec-disabled", disRank.map(j => j.schoolRaw), disRank.map(j => j.riskyTotal), QBR.COLORS.red, DL("hbar", disRank.map(j => fmt(j.riskyTotal))));

  // correlation note
  $("sec-corr").innerHTML = (enabledRisk != null && disabledRisk != null)
    ? `Schools with Security Defaults <b>Not Enabled</b> average <b>${disabledRisk.toFixed(1)}</b> risky users vs <b>${enabledRisk.toFixed(1)}</b> for tenants with them <b>Enabled</b> — a ${disabledRisk > enabledRisk ? ((disabledRisk - enabledRisk) / (enabledRisk || 1) * 100).toFixed(0) + "% higher" : "lower"} identity-risk exposure.`
    : "Insufficient overlap between security snapshot and risky-user data for the current filter.";

  // table (expandable) — the Auth-method filter narrows THIS table only (KPIs and
  // charts above stay portfolio-wide, like the "Show all" toggle is table-local).
  const wantMethod = APP.filters.authMethod || "ALL";
  const tableRows = joined.filter(j => matchAuthMethod(j.methods, wantMethod));
  const secSorted = tableRows.slice().sort(byDesc(j => j.riskyTotal));
  $("tbl-sec").innerHTML = expandSlice("sec", secSorted, 25).map(j =>
    `<tr><td>${esc(j.schoolRaw)}</td><td>${esc(SD_LABEL[j.securityDefault] || j.securityDefault)}</td><td>${esc(j.authMethods || "—")}</td><td class="text-end">${fmt(j.riskyTotal)}</td></tr>`
  ).join("") || `<tr><td colspan="4" class="text-muted">No schools match this authentication-method filter.</td></tr>`;
  setExpandBtn("sec", 25);
}

/* ==================== GDAP ACCESS ==================== */
// Parse an optional GDAP expiry value into {ms,label,cls}. Blank → neutral.
function gdapExpiry(raw) {
  if (raw == null || String(raw).trim() === "") return { ms: null, label: "—", cls: "" };
  let ms = null;
  if (raw instanceof Date) ms = raw.getTime();
  else if (typeof raw === "number") ms = raw; // (epoch ms if ever provided)
  else { const d = new Date(String(raw)); if (!isNaN(d)) ms = d.getTime(); }
  if (ms == null) return { ms: null, label: esc(String(raw)), cls: "" };
  const days = Math.floor((ms - Date.now()) / 86400000);
  const dt = new Date(ms); const lbl = dt.getFullYear() + "-" + String(dt.getMonth() + 1).padStart(2, "0") + "-" + String(dt.getDate()).padStart(2, "0");
  if (days < 0) return { ms, label: lbl, cls: "dr-expired", tag: "Expired" };
  if (days <= 60) return { ms, label: lbl, cls: "dr-expiring", tag: "Expiring Soon" };
  return { ms, label: lbl, cls: "dr-registered", tag: "Active" };
}

function renderGdap() {
  const m = APP.model, f = APP.filters;
  const filtered = m.security
    .filter(s => f.school === "ALL" || s.schoolRaw === f.school)
    .filter(s => f.org === "ALL" || orgOf(s.key) === f.org);
  // Dedupe safety: collapse any repeated school (by normalized key) to one row.
  // Keep the strongest record — GDAP present > more relationships > longer name.
  const byKey = new Map();
  filtered.forEach(s => {
    const cur = byKey.get(s.key);
    if (!cur) { byKey.set(s.key, s); return; }
    const rank = (s.hasGdap ? 1 : 0) - (cur.hasGdap ? 1 : 0)
      || ((s.gdapIds || []).length - (cur.gdapIds || []).length)
      || (s.schoolRaw.length - cur.schoolRaw.length);
    if (rank > 0) byKey.set(s.key, s);
  });
  const rows = [...byKey.values()];
  const dupCollapsed = filtered.length - rows.length;
  const withG = rows.filter(s => s.hasGdap);
  const without = rows.filter(s => !s.hasGdap);
  const multi = withG.filter(s => (s.gdapIds || []).length > 1);
  const cov = rows.length ? (withG.length / rows.length * 100) : 0;

  $("kpi-gdap").innerHTML =
    kpi("With GDAP", fmt(withG.length), "green") +
    kpi("Without GDAP", fmt(without.length), "red") +
    kpi("Coverage %", cov.toFixed(1) + "%", "blue") +
    kpi("Multiple Relationships", fmt(multi.length), "purple");

  QBR.chart.stacked100("ch-gdap-dist", [
    { label: "Granted", value: withG.length, color: QBR.COLORS.green },
    { label: "No GDAP", value: without.length, color: QBR.COLORS.red },
  ]);

  APP._gdapRows = rows.slice().sort((a, b) => a.schoolRaw.localeCompare(b.schoolRaw));
  APP._gdapDup = dupCollapsed;
  renderGdapTable("");
}

function renderGdapTable(q) {
  const sf = ($("gdap-status") && $("gdap-status").value) || "GRANTED";
  const list = (APP._gdapRows || []).filter(s =>
    (!q || s.schoolRaw.toLowerCase().includes(q.toLowerCase())) &&
    (sf === "ALL" || (sf === "GRANTED" ? s.hasGdap : !s.hasGdap)));
  const badge = (has) => has ? '<span class="risk-badge risk-none">Granted</span>' : '<span class="risk-badge risk-critical">None</span>';
  $("tbl-gdap").innerHTML = list.map(s => {
    const ids = (s.gdapIds || []);
    // Multiple relationships with no expiry data → can't auto-pick the active one.
    const multiTag = (ids.length > 1 && (s.gdapExpiryRaw == null || String(s.gdapExpiryRaw).trim() === ""))
      ? '<div class="small text-warning">Multiple — verify which is active</div>' : "";
    const idHtml = ids.length ? `<span class="gdap-id">${ids.map(esc).join("<br>")}</span>${multiTag}` : '<span class="text-muted">—</span>';
    const exp = gdapExpiry(s.gdapExpiryRaw);
    const expHtml = exp.ms == null ? `<span class="text-muted">${exp.label}</span>`
      : `<span class="dr-badge ${exp.cls}">${exp.label}${exp.tag ? " · " + exp.tag : ""}</span>`;
    return `<tr><td>${esc(s.schoolRaw)}</td><td>${esc(orgOf(s.key))}</td><td>${badge(s.hasGdap)}</td>
      <td class="text-end">${ids.length || 0}</td><td class="small">${idHtml}</td>
      <td class="small">${s.gdapRoles ? esc(s.gdapRoles) : '<span class="text-muted">—</span>'}</td>
      <td>${expHtml}</td></tr>`;
  }).join("") || '<tr><td colspan="7" class="text-muted">No schools match.</td></tr>';
  const g = (APP._gdapRows || []);
  const dup = APP._gdapDup || 0;
  $("gdap-note").textContent = `${g.filter(s => s.hasGdap).length} of ${g.length} unique tenants have a GDAP delegated-admin relationship`
    + (dup ? ` · ${dup} duplicate row${dup === 1 ? "" : "s"} collapsed` : "") + ".";
}

/* ==================== DASHBOARD 3: CANVA ADOPTION ==================== */
function renderCanva() {
  const rows = dedupeBySchool(applyFilters(APP.model.canva).filter(c => c.users != null), c => c.users);
  const totalUsers = sum(rows, c => c.users);
  const active = rows.filter(c => c.users > 0);
  const certs = rows.filter(c => /active/i.test(String(c.certStatus))).length;
  const top = rows.slice().sort(byDesc(c => c.users)).slice(0, 10);

  $("kpi-canva").innerHTML =
    kpi("Total Canva Users", fmt(totalUsers), "blue") +
    kpi("Active Schools", fmt(active.length), "green") +
    kpi("Certificates Active", fmt(certs), "teal") +
    kpi("Largest School", top[0] ? esc(top[0].schoolRaw) : "—", "purple");

  QBR.chart.hbar("ch-canva-top", top.map(c => c.schoolRaw), top.map(c => c.users), QBR.paletteColors(top.length), DL("hbar", top.map(c => shareLabel(fmt(c.users), c.users, totalUsers))));
  setNote("canva-top-note", concentrationNote(rows.filter(c => c.users > 0).map(c => c.users).sort((a, b) => b - a), totalUsers, fmt(totalUsers) + " Canva users", "school"));

  // Adoption by quarter — respects Org + School filters; the quarter axis shows
  // the SELECTED quarters (all four when "All").
  const f = APP.filters;
  const orgOf = c => c.org || (APP.model.orgByKey && APP.model.orgByKey[c.key]) || "—";
  const base = APP.model.canva.filter(c => c.users != null
    && (f.org === "ALL" || orgOf(c) === f.org)
    && (f.school === "ALL" || c.schoolRaw === f.school));
  const axis = selQuarters();
  const qUsers = axis.map(q => sum(base.filter(c => c.quarter === q), c => c.users));
  QBR.chart.line("ch-canva-quarter", axis, [{ label: "Canva Users", data: qUsers, color: QBR.COLORS.green, fill: true }],
    DL("line", qUsers.map(v => v ? fmt(v) : null)));

  // Active Schools by Quarter (Quadrant IV) — distinct schools with Canva users > 0
  // per quarter. SNAPSHOT metric: each quarter stands alone, never summed. Same
  // Org + School scope as "Adoption by Quarter", driven by the selected quarter
  // chips (`axis` = selQuarters()).
  //
  // Reporting reality: the CURRENT/future quarter is only partially entered (Canva
  // rows trickle in), so plotting its tiny count as a comparable column — and a
  // QoQ % against it — would read as a collapse rather than "not reported yet".
  // A quarter is therefore classified by how much of its own Canva sheet is filled
  // in: `ok` (majority recorded → a real column + in the QoQ line), `partial`
  // (some rows, but <50% → "in progress", not plotted among complete quarters),
  // or `none` (nothing recorded → "no data yet"). Threshold is data-driven, not a
  // hardcoded "current quarter", so it self-corrects as the quarter fills in.
  const scoped = APP.model.canva.filter(c =>
    (f.org === "ALL" || orgOf(c) === f.org) &&
    (f.school === "ALL" || c.schoolRaw === f.school));
  const qStat = axis.map(q => {
    const all = scoped.filter(c => c.quarter === q);
    const pop = all.filter(c => c.users != null);
    const keys = new Set(); pop.forEach(c => { if (c.users > 0) keys.add(c.key); });
    const cov = all.length ? pop.length / all.length : 0;
    return { q, count: keys.size, populated: pop.length, cov,
             state: pop.length === 0 ? "none" : (cov >= 0.5 ? "ok" : "partial") };
  });
  const single = axis.length === 1;
  // Which quarters get a column: complete ones normally; a single explicitly-
  // selected quarter is always shown if it has any data; if nothing is complete
  // (e.g. only partial quarters selected) fall back to the partial ones so the
  // tile is never needlessly blank.
  let plot = qStat.filter(s => s.state === "ok");
  if (single) plot = qStat.filter(s => s.populated > 0);
  else if (!plot.length) plot = qStat.filter(s => s.state === "partial");
  const plotSet = new Set(plot.map(s => s.q));
  QBR.chart.bar("ch-canva-active", plot.map(s => s.q),
    [{ label: "Active Schools", data: plot.map(s => s.count), backgroundColor: plot.map(s => QBR.QUARTER_COLORS[s.q]) }],
    DL("vbar", plot.map(s => fmt(s.count)), { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } }));
  // Caption. Single quarter → a stat callout. Otherwise a QoQ growth line across
  // consecutive plotted quarters, then a trailing status note for any quarter that
  // was in progress or had no data.
  let actNote = "";
  if (single && plot.length) {
    const s = plot[0];
    actNote = `${fmt(s.count)} active school${s.count === 1 ? "" : "s"} · ${s.q}${s.state === "partial" ? " (in progress)" : ""}`;
  } else {
    const seg = [];
    for (let i = 1; i < plot.length; i++) {
      const prev = plot[i - 1].count, cur = plot[i].count;
      if (prev > 0) {
        const d = cur - prev, pct = Math.round(d / prev * 100), sgn = d >= 0 ? "+" : "−";
        seg.push(`${plot[i - 1].q} → ${plot[i].q}: ${sgn}${fmt(Math.abs(d))} school${Math.abs(d) === 1 ? "" : "s"} (${sgn}${Math.abs(pct)}%)`);
      }
    }
    actNote = seg.join("  ·  ");
  }
  const prog = qStat.filter(s => !plotSet.has(s.q) && s.state === "partial").map(s => s.q);
  const none = qStat.filter(s => !plotSet.has(s.q) && s.state === "none").map(s => s.q);
  if (prog.length) actNote += (actNote ? "  ·  " : "") + `${prog.join(", ")}: in progress`;
  if (none.length) actNote += (actNote ? "  ·  " : "") + `${none.join(", ")}: no data yet`;
  const actNoteEl = $("canva-active-note"); if (actNoteEl) actNoteEl.textContent = actNote;

  const low = active.slice().sort((a, b) => a.users - b.users).slice(0, 10);
  QBR.chart.hbar("ch-canva-low", low.map(c => c.schoolRaw), low.map(c => c.users), QBR.COLORS.orange, DL("hbar", low.map(c => fmt(c.users))));
}

/* ==================== DASHBOARD 4: OFFICE 365 USAGE ==================== */
// Per-service field map (Active / Inactive / Total) used by the usage detail.
const USAGE_SVCS = [
  { label: "Exchange", a: "exchangeActive", i: "exchangeInactive", t: "exchangeTotal", color: QBR.COLORS.teal },
  { label: "OneDrive", a: "onedriveActive", i: "onedriveInactive", t: "onedriveTotal", color: QBR.COLORS.blue },
  { label: "SharePoint", a: "sharepointActive", i: "sharepointInactive", t: "sharepointTotal", color: QBR.COLORS.green },
  { label: "Teams", a: "teamsActive", i: "teamsInactive", t: "teamsTotal", color: QBR.COLORS.orange },
  { label: "Office 365", a: "office365Active", i: "office365Inactive", t: "office365Total", color: QBR.COLORS.purple },
];

function selectUsageTenant(key) {
  APP.usageTenant = key;
  const nm = (APP.model.usage.find(u => u.key === key) || {}).schoolRaw;
  if ($("usage-tenant") && nm) $("usage-tenant").value = nm;
  renderTenantDetail(key);
}

// Admin-center-style per-tenant drill-down: snapshot KPIs + Active/Inactive by
// service + a quarter trend of active users. Data is quarterly (not daily).
function renderTenantDetail(key) {
  if (!key || !APP.model) { $("kpi-td").innerHTML = ""; $("td-note").textContent = "Select a tenant to see its detailed usage."; return; }
  const recs = APP.model.usage.filter(u => u.key === key);
  if (!recs.length) { $("kpi-td").innerHTML = ""; $("td-note").textContent = "No usage rows for this tenant."; return; }
  const QIDX = { Q1: 0, Q2: 1, Q3: 2, Q4: 3 };
  // "populated" = a quarter with real activity (avoids empty/uncollected quarters
  // that carry a 0 or blank snapshot, e.g. Q3 not yet gathered for some tenants).
  const populated = u => (u.usagePct > 0) || (u.office365Active > 0) || USAGE_SVCS.some(s => u[s.a] > 0);
  const withData = recs.filter(populated);
  const L = withData.slice().sort((a, b) => (QIDX[b.quarter] || 0) - (QIDX[a.quarter] || 0))[0] || recs[0];
  const name = (recs[0] || {}).schoolRaw || key;
  $("td-note").innerHTML = `Detailed Office 365 usage for <b>${esc(name)}</b>` +
    (L.quarter ? ` · snapshot <b>${esc(L.quarter)}</b>` : "") +
    (L.dateExtracted ? ` · extracted ${esc(String(L.dateExtracted))}` : "") +
    `. Reflects the Microsoft 365 admin-center usage export (quarterly snapshot; Viva Engage isn't tracked in the workbook).`;
  $("kpi-td").innerHTML =
    kpi("Assigned Licenses", fmt(L.assignedLicenses), "blue") +
    kpi("Activated Apps", fmt(L.activated), "teal") +
    kpi("Total Active (O365)", fmt(L.office365Active), "green") +
    kpi("Usage %", L.usagePct == null ? "—" : L.usagePct.toFixed(1) + "%", "purple");
  // Active vs Inactive by service (stacked horizontal = Total).
  const act = USAGE_SVCS.map(s => L[s.a] || 0);
  const inact = USAGE_SVCS.map(s => { const iv = L[s.i]; if (iv != null) return iv; const t = L[s.t], a = L[s.a]; return (t != null && a != null) ? Math.max(0, t - a) : 0; });
  QBR.chart.bar("ch-td-adoption", USAGE_SVCS.map(s => s.label), [
    { label: "Active", data: act, backgroundColor: QBR.COLORS.green },
    { label: "Inactive", data: inact, backgroundColor: "#C8C6C4" },
  ], { indexAxis: "y", scales: { x: { stacked: true }, y: { stacked: true } }, dataLabels: { mode: "stack" } });
  // Per-quarter Office 365 Active vs Inactive — stacked bar with actual numbers
  // (admin-center "Active users" design). Only quarters with real activity.
  const qs = ["Q1", "Q2", "Q3", "Q4"].filter(q => recs.some(u => u.quarter === q && populated(u)));
  const byQ = {}; recs.forEach(u => { byQ[u.quarter] = u; });
  const actQ = qs.map(q => (byQ[q] || {}).office365Active || 0);
  const inactQ = qs.map(q => { const u = byQ[q] || {}; const iv = u.office365Inactive; if (iv != null) return iv; const t = u.office365Total, a = u.office365Active; return (t != null && a != null) ? Math.max(0, t - a) : 0; });
  QBR.chart.bar("ch-td-trend", qs.length ? qs : ["—"], [
    { label: "Active", data: actQ, backgroundColor: QBR.COLORS.green },
    { label: "Inactive", data: inactQ, backgroundColor: "#C8C6C4" },
  ], { scales: { x: { stacked: true }, y: { stacked: true } }, dataLabels: { mode: "stack" } });
}

function renderUsage() {
  const rows = dedupeBySchool(applyFilters(APP.model.usage), u => u.usagePct);
  const withPct = rows.filter(u => u.usagePct != null);
  const usageAvg = avg(withPct.map(u => u.usagePct));
  const sorted = withPct.slice().sort(byDesc(u => u.usagePct));

  $("kpi-usage").innerHTML =
    kpi("Average Usage %", usageAvg == null ? "—" : usageAvg.toFixed(1) + "%", "blue") +
    kpi("Most Active School", sorted[0] ? esc(sorted[0].schoolRaw) : "—", "green") +
    kpi("Least Active School", sorted.length ? esc(sorted[sorted.length - 1].schoolRaw) : "—", "orange") +
    kpi("Total Active (O365)", fmt(sum(rows, u => u.office365Active)), "teal");

  const top = sorted.slice(0, 10);
  QBR.chart.hbar("ch-usage-top", top.map(u => u.schoolRaw), top.map(u => +u.usagePct.toFixed(1)), QBR.paletteColors(top.length),
    DL("hbar", top.map(u => u.usagePct.toFixed(1) + "%"), {
      onClick: (evt, els) => { if (els && els.length) { const u = top[els[0].index]; if (u) selectUsageTenant(u.key); } }
    }));

  const svc = (field) => rows.filter(u => u[field] != null).sort(byDesc(u => u[field])).slice(0, 8);
  const drawSvc = (id, field, color) => { const d = svc(field); QBR.chart.hbar(id, d.map(u => u.schoolRaw), d.map(u => u[field]), color, DL("hbar", d.map(u => fmt(u[field])))); };
  drawSvc("ch-usage-onedrive", "onedriveActive", QBR.COLORS.blue);
  drawSvc("ch-usage-sharepoint", "sharepointActive", QBR.COLORS.teal);
  drawSvc("ch-usage-teams", "teamsActive", QBR.COLORS.purple);
  drawSvc("ch-usage-exchange", "exchangeActive", QBR.COLORS.orange);

  // Average adoption by service (portfolio): sum(active) / sum(licensed) per service.
  const adoption = USAGE_SVCS.map(s => {
    const base = rows.filter(u => u[s.t]);
    const a = sum(base, u => u[s.a]), t = sum(base, u => u[s.t]);
    return t ? Math.round(a / t * 1000) / 10 : 0;
  });
  QBR.chart.bar("ch-usage-adoption", USAGE_SVCS.map(s => s.label),
    [{ label: "Adoption %", data: adoption, backgroundColor: USAGE_SVCS.map(s => s.color) }],
    DL("vbar", adoption.map(v => v ? v.toFixed(1) + "%" : null),
      { plugins: { legend: { display: false } }, scales: { y: { max: 100, ticks: { callback: v => v + "%" } } } }));

  // by organization (stacked-ish: O365 active)
  const org = {}; rows.forEach(u => { const o = u.org || orgOf(u.key); org[o] = (org[o] || 0) + (u.office365Active || 0); });
  const orgE = Object.entries(org).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const activeTot = Object.values(org).reduce((a, b) => a + b, 0);
  QBR.chart.bar("ch-usage-org", orgE.map(x => x[0]), [{ label: "O365 Active Users", data: orgE.map(x => x[1]), backgroundColor: QBR.paletteColors(orgE.length) }],
    DL("vbar", orgE.map(x => sharePct(x[1], activeTot) || fmt(x[1])), { plugins: { legend: { display: false } } }));
  const actDesc = rows.map(u => u.office365Active || 0).filter(v => v > 0).sort((a, b) => b - a);
  const topAct = rows.filter(u => u.office365Active > 0).sort(byDesc(u => u.office365Active))[0];
  setNote("usage-org-note", activeTot ? `Bar labels show each organization's share of ${fmt(activeTot)} active O365 users. ` +
    (topAct ? `Most active school: ${esc(topAct.schoolRaw)} with ${fmt(topAct.office365Active)} (${sharePct(topAct.office365Active, activeTot)}). ` : "") +
    concentrationNote(actDesc, activeTot, "active users", "school").replace(/^The largest school alone holds <b>[^<]*<\/b>\. /, "") : "");

  // top-5 table w/ conditional coloring vs average
  const band = (p) => usageAvg == null ? "" : (p >= usageAvg * 1.1 ? "cell-green" : p <= usageAvg * 0.9 ? "cell-red" : "cell-yellow");
  const usageRow = (u, i) =>
    `<tr><td>${i + 1}</td><td>${esc(u.schoolRaw)}</td><td class="text-end ${band(u.usagePct)}">${u.usagePct.toFixed(1)}%</td>
     <td class="text-end">${fmt(u.onedriveActive)}</td><td class="text-end">${fmt(u.sharepointActive)}</td>
     <td class="text-end">${fmt(u.teamsActive)}</td><td class="text-end">${fmt(u.office365Active)}</td><td class="text-end">${sharePct(u.office365Active, sum(rows, r => r.office365Active)) || "—"}</td></tr>`;
  $("tbl-usage").innerHTML = expandSlice("usage", sorted, 5).map(usageRow).join("") || '<tr><td colspan="8" class="text-muted">No data</td></tr>';
  setExpandBtn("usage", 5);
  // Bottom 5 = lowest utilization (ascending), for adoption targeting
  const low5 = sorted.slice(-5).reverse();
  $("tbl-usage-low").innerHTML = low5.map(usageRow).join("") || '<tr><td colspan="8" class="text-muted">No data</td></tr>';

  // Per-tenant detail: fill the tenant picker and render the current/default tenant.
  const tenants = [...new Map(APP.model.usage.map(u => [u.key, u.schoolRaw])).entries()]
    .map(([k, name]) => ({ k, name })).sort((a, b) => a.name.localeCompare(b.name));
  if ($("dl-usage-tenant")) $("dl-usage-tenant").innerHTML = tenants.map(t => `<option value="${escAttr(t.name)}"></option>`).join("");
  const dk = (APP.usageTenant && APP.model.usage.some(u => u.key === APP.usageTenant)) ? APP.usageTenant : (sorted[0] && sorted[0].key);
  APP.usageTenant = dk || null;
  const nm = (tenants.find(t => t.k === dk) || {}).name;
  if ($("usage-tenant") && nm && $("usage-tenant").value !== nm) $("usage-tenant").value = nm;
  renderTenantDetail(dk);
}

/* ==================== DASHBOARD 5: STORAGE ==================== */
// Per-service storage fields for the storage detail (GB; converted to TB in UI).
const STOR_SVCS = [
  { label: "OneDrive", f: "onedriveGB", color: QBR.COLORS.blue },
  { label: "Exchange", f: "exchangeGB", color: QBR.COLORS.orange },
  { label: "SharePoint", f: "sharepointGB", color: QBR.COLORS.teal },
];

function selectStorageTenant(key) {
  APP.storageTenant = key;
  const nm = (APP.model.storage.find(s => s.key === key) || {}).schoolRaw;
  if ($("stor-tenant") && nm) $("stor-tenant").value = nm;
  renderStorageDetail(key);
}

// Per-tenant storage drill-down: capacity KPIs + composition-vs-capacity stacked
// bar + per-quarter stacked bar by service (TB). Quarterly snapshots (Q1/Q2).
function renderStorageDetail(key) {
  if (!key || !APP.model) { $("kpi-sd").innerHTML = ""; $("sd-note").textContent = "Select a tenant to see its storage breakdown."; return; }
  // Respect the global Quarter chips: the snapshot/KPIs and the per-quarter bars are
  // scoped to the selected quarters (matches every other tab). "All" → all quarters.
  const recs = APP.model.storage.filter(s => s.key === key && qActive(s.quarter));
  if (!recs.length) {
    $("kpi-sd").innerHTML = "";
    $("sd-note").textContent = isAllQ()
      ? "No storage rows for this tenant."
      : `No storage rows for this tenant in ${quarterLabel()}.`;
    QBR.chart.bar("ch-sd-comp", ["—"], [{ label: "", data: [0] }], {});
    QBR.chart.bar("ch-sd-quarter", ["—"], [{ label: "", data: [0] }], {});
    if ($("sd-growth")) $("sd-growth").innerHTML = "";
    return;
  }
  const QIDX = { Q1: 0, Q2: 1, Q3: 2, Q4: 3 };
  const populated = s => (s.usedGB > 0) || (s.totalGB > 0);
  const withData = recs.filter(populated);
  const L = withData.slice().sort((a, b) => (QIDX[b.quarter] || 0) - (QIDX[a.quarter] || 0))[0] || recs[0];
  const name = (recs[0] || {}).schoolRaw || key;
  const TB = v => (v == null ? 0 : v / 1024);
  const total = L.totalGB || 0;
  const used = L.usedGB != null ? L.usedGB : ((L.onedriveGB || 0) + (L.exchangeGB || 0) + (L.sharepointGB || 0));
  const free = Math.max(0, total - used);
  $("sd-note").innerHTML = `Storage breakdown for <b>${esc(name)}</b>` + (L.quarter ? ` · snapshot <b>${esc(L.quarter)}</b>` : "") +
    `. OneDrive + Exchange + SharePoint used vs total pooled capacity (quarterly snapshot).`;
  $("kpi-sd").innerHTML =
    kpi("Total Capacity", tb(total), "blue") +
    kpi("Used", tb(used), "orange") +
    kpi("Free", tb(free), "green") +
    (() => { const up = L.pct != null ? L.pct : (total ? used / total * 100 : null); const b = capacityBand(up);
      return kpi("Utilization %" + (b ? " · " + b.label : ""), up == null ? "—" : (Math.round(up * 10) / 10).toFixed(1) + "%", b ? (b.tone === "bad" ? "red" : "orange") : "purple"); })();
  // Composition vs capacity: single horizontal stacked bar (TB). Services + any
  // unclassified "Other used" + Free sum exactly to total capacity.
  const svcSum = (L.onedriveGB || 0) + (L.exchangeGB || 0) + (L.sharepointGB || 0);
  const otherUsed = Math.max(0, used - svcSum);
  const seg = [
    { label: "OneDrive", v: TB(L.onedriveGB), color: QBR.COLORS.blue },
    { label: "Exchange", v: TB(L.exchangeGB), color: QBR.COLORS.orange },
    { label: "SharePoint", v: TB(L.sharepointGB), color: QBR.COLORS.teal },
    { label: "Other", v: TB(otherUsed), color: "#8A8886" },
    { label: "Free", v: TB(free), color: "#C8C6C4" },
  ].filter(s => s.label === "Free" || s.v > 0.01);
  QBR.chart.bar("ch-sd-comp", [""], seg.map(s => ({ label: s.label, data: [+s.v.toFixed(2)], backgroundColor: s.color })),
    { indexAxis: "y", layout: { padding: { right: 74 } }, scales: { x: { stacked: true, title: { display: true, text: "TB" } }, y: { stacked: true } },
      dataLabels: { mode: "stack", showTotal: true, totalSuffix: " TB" } });
  // Per-quarter stacked bar by service (TB), with the used total above each column.
  const qs = ["Q1", "Q2", "Q3", "Q4"].filter(q => recs.some(s => s.quarter === q && populated(s)));
  const byQ = {}; recs.forEach(s => { byQ[s.quarter] = s; });
  const ds = STOR_SVCS.map(s => ({ label: s.label, backgroundColor: s.color, data: qs.map(q => +TB((byQ[q] || {})[s.f] || 0).toFixed(2)) }));
  QBR.chart.bar("ch-sd-quarter", qs.length ? qs : ["—"], ds,
    { layout: { padding: { top: 22 } }, scales: { x: { stacked: true }, y: { stacked: true, title: { display: true, text: "TB" } } },
      dataLabels: { mode: "stack", showTotal: true, totalSuffix: " TB" } });
  // Growth caption (latest vs previous populated quarter).
  const qTot = qs.map(q => STOR_SVCS.reduce((a, s) => a + TB((byQ[q] || {})[s.f] || 0), 0));
  let g = "";
  if (qs.length >= 2) {
    const cur = qTot[qs.length - 1], prev = qTot[qs.length - 2], diff = cur - prev, pctG = prev ? diff / prev * 100 : 0;
    g = `${qs[qs.length - 2]} → ${qs[qs.length - 1]}: <b>${diff >= 0 ? "+" : ""}${diff.toFixed(2)} TB</b> (${pctG >= 0 ? "+" : ""}${pctG.toFixed(1)}%) storage ${diff >= 0 ? "growth" : "decline"}.`;
  }
  if ($("sd-growth")) $("sd-growth").innerHTML = g;
}

function renderStorage() {
  const rows = dedupeBySchool(applyFilters(APP.model.storage).filter(s => s.usedGB != null), s => s.usedGB);
  const totalGB = sum(rows, s => s.usedGB);
  const avgPct = avg(rows.filter(s => s.pct != null).map(s => s.pct));
  const sorted = rows.slice().sort(byDesc(s => s.usedGB));

  const cap = storageCapacityList(), capCrit = cap.filter(x => x.band.tone === "bad").length;
  $("kpi-storage").innerHTML =
    kpi("Total Storage", tb(totalGB), "blue") +
    kpi("Average Usage %", avgPct == null ? "—" : avgPct.toFixed(1) + "%", "teal") +
    kpi(CAP_KPI_LABEL(), fmt(cap.length), capCrit ? "red" : cap.length ? "orange" : "green") +
    kpi("Largest Consumer", sorted[0] ? esc(sorted[0].schoolRaw) : "—", "orange") +
    kpi("Lowest Consumer", sorted.length ? esc(sorted[sorted.length - 1].schoolRaw) : "—", "green");
  renderCapacityWatch(cap);

  const top = sorted.slice(0, 10);
  QBR.chart.hbar("ch-stor-top", top.map(s => s.schoolRaw), top.map(s => +(s.usedGB / 1024).toFixed(2)), QBR.paletteColors(top.length),
    DL("hbar", top.map(s => shareLabel((s.usedGB / 1024).toFixed(1) + " TB", s.usedGB, totalGB)), {
      onClick: (evt, els) => { if (els && els.length) { const s = top[els[0].index]; if (s) selectStorageTenant(s.key); } }
    }));

  const svc = (field, color, id) => { const d = rows.filter(s => s[field] != null).sort(byDesc(s => s[field])).slice(0, 8);
    QBR.chart.hbar(id, d.map(s => s.schoolRaw), d.map(s => +(s[field] / 1024).toFixed(2)), color, DL("hbar", d.map(s => (s[field] / 1024).toFixed(1) + " TB"))); };
  svc("onedriveGB", QBR.COLORS.blue, "ch-stor-onedrive");
  svc("sharepointGB", QBR.COLORS.teal, "ch-stor-sharepoint");
  svc("exchangeGB", QBR.COLORS.orange, "ch-stor-exchange");


  setNote("stor-top-note", concentrationNote(sorted.filter(s => s.usedGB > 0).map(s => s.usedGB), totalGB, tb(totalGB) + " of managed storage", "tenant"));
  const storRow = (s, i) =>
    `<tr><td>${i + 1}</td><td>${esc(s.schoolRaw)}</td><td class="text-end">${tb(s.usedGB)}</td>
     <td class="text-end">${sharePct(s.usedGB, totalGB) || "—"}</td>
     <td class="text-end">${tb(s.totalGB)}</td>
     <td class="text-end">${storagePct(s) == null ? "—" : storagePct(s).toFixed(1) + "%"}${capBadge(storagePct(s))}</td></tr>`;
  $("tbl-storage").innerHTML = expandSlice("storage", sorted, 10).map(storRow).join("") || '<tr><td colspan="6" class="text-muted">No data</td></tr>';
  setExpandBtn("storage", 10);
  // Bottom 10 = lowest Used (ascending); rank 1 = lowest. Zeros included.
  const low10 = sorted.slice().reverse().slice(0, 10);
  $("tbl-storage-low").innerHTML = low10.map(storRow).join("") || '<tr><td colspan="6" class="text-muted">No data</td></tr>';

  // Per-tenant detail: fill picker + render current/default tenant.
  const tenants = [...new Map(APP.model.storage.map(s => [s.key, s.schoolRaw])).entries()]
    .map(([k, name]) => ({ k, name })).sort((a, b) => a.name.localeCompare(b.name));
  if ($("dl-stor-tenant")) $("dl-stor-tenant").innerHTML = tenants.map(t => `<option value="${escAttr(t.name)}"></option>`).join("");
  const dk = (APP.storageTenant && APP.model.storage.some(s => s.key === APP.storageTenant)) ? APP.storageTenant : (sorted[0] && sorted[0].key);
  APP.storageTenant = dk || null;
  const nm = (tenants.find(t => t.k === dk) || {}).name;
  if ($("stor-tenant") && nm && $("stor-tenant").value !== nm) $("stor-tenant").value = nm;
  renderStorageDetail(dk);
}

// Capacity watchlist card (Storage page): every tenant whose latest snapshot in
// scope is at/above the warning level, worst first. Rows open the tenant detail.
function renderCapacityWatch(cap) {
  const body = $("tbl-stor-cap"), note = $("stor-cap-note");
  if (!body) return;
  if (note) note.innerHTML = `Tenants at or above <b>${storWarn()}%</b> (near capacity) or <b>${storCrit()}%</b> (at capacity) of their pooled storage, ` +
    `using each tenant's latest snapshot in the current filter. If the Data Quality audit flags capacity for that quarter, verify the figure before acting.`;
  body.innerHTML = cap.length ? cap.map(x => {
    const r = x.r, w = Math.min(100, x.pct);
    return `<tr class="cap-row" data-cap-key="${escAttr(r.key)}" tabindex="0" title="Open ${escAttr(r.schoolRaw)} storage detail">` +
      `<td>${esc(r.schoolRaw)}</td><td>${esc(orgOf(r.key))}</td><td>${esc(r.quarter || "—")}</td>` +
      `<td class="text-end">${tb(r.usedGB)}</td><td class="text-end">${tb(r.totalGB)}</td>` +
      `<td class="text-end"><span class="cap-meter cap-${x.band.tone}" aria-hidden="true"><span style="width:${w.toFixed(1)}%"></span></span>${x.pct.toFixed(1)}%</td>` +
      `<td><span class="cap-badge cap-${x.band.tone}">${x.band.label}</span></td></tr>`;
  }).join("") : `<tr><td colspan="7" class="text-muted">No tenant is at or above ${storWarn()}% of its pooled storage for the current filter.</td></tr>`;
  body.querySelectorAll(".cap-row").forEach(tr => {
    const go = () => { selectStorageTenant(tr.dataset.capKey); const d = $("kpi-sd"); if (d) d.scrollIntoView({ block: "center" }); };
    tr.addEventListener("click", go);
    tr.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); } });
  });
}

/* ==================== GOOGLE POSTMASTER ==================== */
const REP_ORDER = ["HIGH", "MEDIUM", "LOW", "BAD"];               // Google's *rated* tiers
// Actionable live-status states carried by the current tracker (not rated tiers).
const REP_ACTION = ["Issues detected", "Verify to see health"];
// Full display order, worst / most-actionable first, for the distribution chart.
const REP_DISPLAY = ["BAD", "Issues detected", "LOW", "Verify to see health", "MEDIUM",
  "HIGH", "Not enough data", "No data to display", "No entry"];
const REP_COLOR = { HIGH: QBR.COLORS.green, MEDIUM: QBR.COLORS.yellow, LOW: QBR.COLORS.orange, BAD: QBR.COLORS.red,
  "Issues detected": QBR.COLORS.red, "Verify to see health": QBR.COLORS.orange,
  "No data to display": QBR.COLORS.gray, "Not enough data": "#A19F9D", "No entry": "#C8C6C4" };
const VERIFY_ORDER = ["Verified", "Managed", "Unverified", "Not Managed", "Unknown"];
const VERIFY_COLOR = { "Verified": QBR.COLORS.green, "Managed": QBR.COLORS.blue, "Unverified": QBR.COLORS.orange, "Not Managed": QBR.COLORS.red, "Unknown": QBR.COLORS.gray };
const DNS_ORDER = ["Records added", "Pending", "Not Managed", "Unknown"];
const DNS_COLOR = { "Records added": QBR.COLORS.green, "Pending": QBR.COLORS.yellow, "Not Managed": QBR.COLORS.red, "Unknown": QBR.COLORS.gray };

function orgOf(key) { return APP.model.orgByKey[key] || "Unspecified"; }

function renderPostmaster() {
  const m = APP.model, f = APP.filters;
  const hasVerify = m.pmVerify && m.pmVerify.length > 0;

  /* ---------- Verification (per-domain; respects Org + School only) ---------- */
  const vlist = Object.values(m.pmVerifyByKey || {}).filter(v =>
    (f.org === "ALL" || (v.org || orgOf(v.key)) === f.org) &&
    (f.school === "ALL" || v.schoolRaw === f.school))
    .sort((a, b) => a.schoolRaw.localeCompare(b.schoolRaw));
  const vCounts = { "Verified": 0, "Managed": 0, "Unverified": 0, "Not Managed": 0, "Unknown": 0 };
  const dnsCounts = { "Records added": 0, "Pending": 0, "Not Managed": 0, "Unknown": 0 };
  vlist.forEach(v => { if (vCounts[v.pmTool] != null) vCounts[v.pmTool]++; if (dnsCounts[v.dns] != null) dnsCounts[v.dns]++; });
  const totalDomains = vlist.length;
  const verified = vCounts["Verified"];
  const notVerified = vCounts["Unverified"] + vCounts["Not Managed"];   // domains not verified in Postmaster
  const pctVerified = totalDomains ? Math.round(verified / totalDomains * 100) : 0;
  const hasDns = ["Records added", "Pending", "Not Managed"].some(k => dnsCounts[k]);   // real DNS data (not just "Unknown")

  /* ---------- Reputation (monthly; respects Quarter + Month + Org + School) ---------- */
  const rows = m.postmaster.filter(p =>
    (f.org === "ALL" || (p.org || orgOf(p.key)) === f.org) &&
    (f.school === "ALL" || p.schoolRaw === f.school) &&
    qActive(p.quarter) &&
    (f.month === "ALL" || p.month === f.month));
  // Representative record per school = the LATEST month that carries a real status.
  // Later months are often still unfilled ("No entry"); picking the plain latest
  // month would let those blanks mask a real "Issues detected" from an earlier month.
  const perSchool = new Map();
  rows.forEach(p => {
    const e = perSchool.get(p.key);
    if (!e) { perSchool.set(p.key, p); return; }
    const pReal = p.reputation !== "No entry", eReal = e.reputation !== "No entry";
    if (pReal !== eReal) { if (pReal) perSchool.set(p.key, p); }   // a real status beats a blank
    else if (p.monthIdx >= e.monthIdx) perSchool.set(p.key, p);    // else the latest month wins
  });
  const list = [...perSchool.values()].sort((a, b) => a.schoolRaw.localeCompare(b.schoolRaw));
  // Count EVERY reputation state present (not just the rated tiers — the tracker's
  // live states "Issues detected" / "Verify to see health" / "Not enough data" are
  // the real signal and were previously collapsed and hidden).
  const repCounts = {};
  REP_DISPLAY.forEach(r => repCounts[r] = 0);
  list.forEach(e => { repCounts[e.reputation] = (repCounts[e.reputation] || 0) + 1; });
  const ratedTotal = REP_ORDER.reduce((a, r) => a + (repCounts[r] || 0), 0);
  const issues = repCounts["Issues detected"] || 0;
  const needVerifyRep = repCounts["Verify to see health"] || 0;

  /* ---------- KPI cards ---------- */
  $("kpi-pm").innerHTML =
    kpi("Total Domains", fmt(totalDomains || m.master.size), "blue") +
    kpi("Verified in Postmaster", hasVerify ? `${fmt(verified)} · ${pctVerified}%` : "—", "green") +
    kpi("Not Verified", hasVerify ? fmt(notVerified) : "—", notVerified ? "orange" : "teal") +
    kpi("Issues Detected", fmt(issues + (repCounts.BAD || 0)), (issues + (repCounts.BAD || 0)) ? "red" : "green") +
    kpi("Verify to See Health", fmt(needVerifyRep), needVerifyRep ? "orange" : "teal");

  $("pm-note").innerHTML = '<b>Issues Detected</b> = Google flags a deliverability problem (review SPF/DKIM/DMARC &amp; sending practices). <b>Verify to See Health</b> / <b>Not Verified</b> = the domain must be verified in Postmaster before Google will report reputation. <b>Not enough data</b> is normal for low-volume domains, not a fault.';

  /* ---------- Charts ---------- */
  // Verification coverage donut
  const vShown = VERIFY_ORDER.filter(k => vCounts[k]);
  QBR.chart.doughnut("ch-pm-verify", vShown, vShown.map(k => vCounts[k]), vShown.map(k => VERIFY_COLOR[k]),
    DL("donut", pctLabels(vShown.map(k => vCounts[k]))));
  // Cloudflare DNS status donut — only meaningful when a DNS column was supplied
  // (standalone Postmaster export). The combined tracker has none, so show a note.
  if (hasDns) {
    const dShown = DNS_ORDER.filter(k => dnsCounts[k]);
    QBR.chart.doughnut("ch-pm-dns", dShown, dShown.map(k => dnsCounts[k]), dShown.map(k => DNS_COLOR[k]),
      DL("donut", pctLabels(dShown.map(k => dnsCounts[k]))));
    if ($("pm-dns-note")) $("pm-dns-note").textContent = "";
  } else {
    QBR.chart.doughnut("ch-pm-dns", [], [], []);
    if ($("pm-dns-note")) $("pm-dns-note").textContent = "No Cloudflare/DNS data in this workbook — it ships with the standalone GOOGLE POSTMASTERTOOLS export.";
  }
  // Reputation distribution across ALL present states (worst-first)
  const repShown = REP_DISPLAY.filter(r => repCounts[r]);
  QBR.chart.stacked100("ch-pm-dist",
    repShown.map(r => ({ label: r, value: repCounts[r], color: REP_COLOR[r] })));

  // Verification by Organization (stacked)
  const orgs = [...new Set(vlist.map(v => v.org || orgOf(v.key)))].sort();
  if (orgs.length) {
    const ds = VERIFY_ORDER.map(k => ({ label: k, backgroundColor: VERIFY_COLOR[k],
      data: orgs.map(o => vlist.filter(v => (v.org || orgOf(v.key)) === o && v.pmTool === k).length) }));
    QBR.chart.bar("ch-pm-org", orgs, ds, { scales: { x: { stacked: true }, y: { stacked: true } } });
  } else { QBR.chart.bar("ch-pm-org", [], [], {}); }

  /* ---------- Reputation watchlist (any actionable state across the loaded period) ---------- */
  // Actionable = a rated tier OR a live problem status. Benign states (Not enough
  // data / No data / No entry) are excluded so the watchlist stays a to-do list.
  const WORST = { BAD: 6, "Issues detected": 5, LOW: 4, "Verify to see health": 3, MEDIUM: 2, HIGH: 1 };
  const isActionable = r => WORST[r] != null;
  const watch = new Map(); // key -> {schoolRaw, domain, ratings, months}
  rows.forEach(p => {
    if (!isActionable(p.reputation)) return;
    const e = watch.get(p.key) || { schoolRaw: p.schoolRaw, domain: p.domain, ratings: {}, months: [] };
    e.ratings[p.reputation] = (e.ratings[p.reputation] || 0) + 1;
    e.months.push({ m: p.month, r: p.reputation, i: p.monthIdx });
    watch.set(p.key, e);
  });
  const watchRows = [...watch.values()].map(e => {
    const worst = Object.keys(e.ratings).sort((a, b) => WORST[b] - WORST[a])[0];
    return { schoolRaw: e.schoolRaw, domain: e.domain, worst,
      months: e.months.sort((a, b) => a.i - b.i).map(x => `${x.m.slice(0, 3)}:${x.r}`).join(" · ") };
  }).sort((a, b) => WORST[b.worst] - WORST[a.worst] || a.schoolRaw.localeCompare(b.schoolRaw));
  const repBadge = (r) => {
    const cls = (r === "HIGH") ? "bg-success" : (r === "MEDIUM") ? "bg-info text-dark"
      : (r === "LOW" || r === "Verify to see health") ? "bg-warning text-dark" : "bg-danger";
    return `<span class="badge ${cls}">${esc(r)}</span>`;
  };
  $("tbl-pm-watch").innerHTML = watchRows.map(w =>
    `<tr><td>${esc(w.schoolRaw)}</td><td class="small">${esc(w.domain || "—")}</td><td>${repBadge(w.worst)}</td><td class="small text-muted">${esc(w.months)}</td></tr>`
  ).join("") || `<tr><td colspan="4" class="text-muted">No domains with reputation issues in this period — nothing flagged.</td></tr>`;

  /* ---------- Onboarding action list (not verified / DNS not managed) ---------- */
  const dnsBadge = (d) => `<span class="badge ${d === "Records added" ? "bg-success" : d === "Pending" ? "bg-warning text-dark" : d === "Not Managed" ? "bg-danger" : "bg-secondary"}">${esc(d)}</span>`;
  const pmBadge = (p) => `<span class="badge ${p === "Verified" ? "bg-success" : p === "Managed" ? "bg-info text-dark" : p === "Unverified" ? "bg-warning text-dark" : p === "Not Managed" ? "bg-danger" : "bg-secondary"}">${esc(p)}</span>`;
  const action = vlist.filter(v => v.pmTool !== "Verified");
  $("tbl-pm-action").innerHTML = action.map(v =>
    `<tr><td>${esc(v.schoolRaw)}</td><td class="small">${esc(v.org || orgOf(v.key))}</td><td class="small">${esc(v.domain || "—")}</td><td>${dnsBadge(v.dns)}</td><td>${pmBadge(v.pmTool)}</td></tr>`
  ).join("") || (hasVerify
    ? `<tr><td colspan="5" class="text-muted">Every domain is verified in Postmaster.</td></tr>`
    : `<tr><td colspan="5" class="text-muted">No verification data — upload GOOGLE POSTMASTERTOOLS.xlsx (with its master "…Data" sheet) to populate this.</td></tr>`);

  $("pm-period").textContent = (hasVerify ? `${totalDomains} domains tracked. ` : "") +
    (f.month !== "ALL" ? `Reputation: ${f.month[0] + f.month.slice(1).toLowerCase()}.` :
      (!isAllQ() ? `Reputation: latest scored month in ${quarterLabel()}.` : "Reputation: each domain's most recent scored month."));
  APP._pmList = list;
  renderPmTable("");
}

function renderPmTable(q) {
  const rf = ($("pm-rep") && $("pm-rep").value) || "ALL";
  const vBy = APP.model.pmVerifyByKey || {};
  const list = (APP._pmList || []).filter(e => {
    if (q && !e.schoolRaw.toLowerCase().includes(q.toLowerCase())) return false;
    if (rf === "ALL") return true;
    if (VERIFY_ORDER.includes(rf)) return (vBy[e.key] ? vBy[e.key].pmTool : "Unknown") === rf;
    return e.reputation === rf;                       // reputation value
  });
  const badge = (r) => {
    if (!r || r === "No entry") return '<span class="text-muted">—</span>';
    if (r === "Issues detected" || r === "BAD") return `<span class="badge bg-danger">${esc(r)}</span>`;
    if (r === "Verify to see health" || r === "LOW") return `<span class="badge bg-warning text-dark">${esc(r)}</span>`;
    if (r === "MEDIUM") return `<span class="badge bg-info text-dark">MEDIUM</span>`;
    if (r === "HIGH") return `<span class="badge bg-success">HIGH</span>`;
    return `<span class="text-muted">${esc(r)}</span>`;   // Not enough data / No data to display
  };
  const vBadge = (key) => {
    const v = vBy[key]; if (!v) return '<span class="text-muted">—</span>';
    const p = v.pmTool;
    return `<span class="badge ${p === "Verified" ? "bg-success" : p === "Managed" ? "bg-info text-dark" : p === "Unverified" ? "bg-warning text-dark" : p === "Not Managed" ? "bg-danger" : "bg-secondary"}">${esc(p)}</span>`;
  };
  const dBadge = (key) => {
    const v = vBy[key]; if (!v) return '<span class="text-muted">—</span>';
    const d = v.dns;
    return `<span class="badge ${d === "Records added" ? "bg-success" : d === "Pending" ? "bg-warning text-dark" : d === "Not Managed" ? "bg-danger" : "bg-secondary"}">${esc(d)}</span>`;
  };
  $("tbl-pm").innerHTML = list.map(e =>
    `<tr><td>${esc(e.schoolRaw)}</td><td class="small">${esc(e.org || orgOf(e.key))}</td><td class="small">${esc(e.domain || "—")}</td>
     <td>${vBadge(e.key)}</td><td>${dBadge(e.key)}</td><td>${badge(e.reputation)}</td>
     <td>${e.spamRaw != null ? esc(e.spamRaw) : '<span class="text-muted">—</span>'}</td></tr>`
  ).join("") || '<tr><td colspan="7" class="text-muted">No schools match.</td></tr>';
}

/* ==================== USER MANAGEMENT ==================== */
function umReadinessScore(key, status) {
  // 0-100 composite: Updated 40 + SecDefault enabled 20 + low risk 20 + usage>=40 10 + canva active 10
  let s = status === "Updated" ? 40 : 0;
  const sec = APP.model.security.find(x => x.key === key);
  if (sec && sec.securityDefault === "ENABLED") s += 20;
  const riskyTotal = APP.model.risky.filter(r => r.key === key && r.risky != null).reduce((a, r) => a + r.risky, 0);
  if (riskyTotal <= QBR.THRESH.HIGH_RISK) s += 20;
  const u = APP.model.usage.find(x => x.key === key && x.usagePct != null);
  if (u && u.usagePct >= QBR.THRESH.USAGE_LOW) s += 10;
  const c = APP.model.canva.find(x => x.key === key && x.users > 0);
  if (c) s += 10;
  return s;
}

function renderUserManagement() {
  const m = APP.model, f = APP.filters;
  const syList = [...new Set(m.usermgmt.map(u => u.sy))];
  const sy = (f.sy && f.sy !== "ALL" && syList.includes(f.sy)) ? f.sy : syList[0];
  let rows = m.usermgmt.filter(u => u.sy === sy)
    .filter(u => f.org === "ALL" || orgOf(u.key) === f.org)
    .filter(u => f.school === "ALL" || u.schoolRaw === f.school)
    .filter(u => f.status === "ALL" || u.status === f.status);

  const total = rows.length;
  const updated = rows.filter(u => u.status === "Updated").length;
  const pending = total - updated;
  const completion = total ? (updated / total * 100) : 0;

  $("um-topcard").innerHTML = `<b>School Year Readiness</b> — ${updated} / ${total} schools updated for <b>${esc(sy)}</b> &nbsp;·&nbsp; <b>${completion.toFixed(1)}%</b> ready`;
  $("kpi-um").innerHTML =
    kpi("Total Schools", fmt(total), "blue") +
    kpi("Schools Updated", fmt(updated), "green") +
    kpi("Schools Pending", fmt(pending), "orange") +
    kpi("Completion %", completion.toFixed(1) + "%", "teal") +
    kpi("SY Readiness %", completion.toFixed(1) + "%", "purple");

  QBR.chart.stacked100("ch-um-readiness", [
    { label: "Updated", value: updated, color: QBR.COLORS.green },
    { label: "Pending", value: pending, color: QBR.COLORS.orange },
  ]);

  // updated schools by organization (stacked Updated/Pending)
  const orgs = [...new Set(rows.map(u => orgOf(u.key)))].sort();
  QBR.chart.bar("ch-um-org", orgs, [
    { label: "Updated", backgroundColor: QBR.COLORS.green, data: orgs.map(o => rows.filter(u => orgOf(u.key) === o && u.status === "Updated").length) },
    { label: "Pending", backgroundColor: QBR.COLORS.orange, data: orgs.map(o => rows.filter(u => orgOf(u.key) === o && u.status === "Pending").length) },
  ], { scales: { x: { stacked: true }, y: { stacked: true } } });

  // readiness by school year (completion % of every SY, unfiltered by status)
  const syPct = syList.map(s => {
    const r = m.usermgmt.filter(u => u.sy === s);
    return r.length ? r.filter(u => u.status === "Updated").length / r.length * 100 : 0;
  });
  QBR.chart.bar("ch-um-sy", syList, [{ label: "Completion %", data: syPct.map(x => +x.toFixed(1)), backgroundColor: QBR.paletteColors(syList.length) }],
    DL("vbar", syPct.map(x => x.toFixed(0) + "%"), { plugins: { legend: { display: false } } }));

  APP._umRows = rows;
  renderUmTable("");

  // correlation table (top by readiness score)
  const corr = rows.map(u => {
    const sec = m.security.find(x => x.key === u.key);
    const riskyTotal = m.risky.filter(r => r.key === u.key && r.risky != null).reduce((a, r) => a + r.risky, 0);
    const usage = m.usage.find(x => x.key === u.key && x.usagePct != null);
    const canva = m.canva.find(x => x.key === u.key && x.users > 0);
    return { u, sec: sec ? sec.securityDefault : "—", risky: riskyTotal,
      usage: usage ? usage.usagePct : null, canva: canva ? canva.users : null,
      score: umReadinessScore(u.key, u.status) };
  }).sort((a, b) => b.score - a.score);
  $("tbl-um-corr").innerHTML = expandSlice("umcorr", corr, 25).map(c => {
    const cls = c.score >= 70 ? "cell-green" : c.score >= 40 ? "cell-yellow" : "cell-red";
    return `<tr><td>${esc(c.u.schoolRaw)}</td><td>${c.u.status === "Updated" ? '<span class="badge bg-success">Updated</span>' : '<span class="badge bg-warning text-dark">Pending</span>'}</td>
      <td>${esc(c.sec)}</td><td class="text-end">${fmt(c.risky)}</td><td class="text-end">${c.usage == null ? "—" : c.usage.toFixed(1) + "%"}</td>
      <td class="text-end">${c.canva == null ? "—" : fmt(c.canva)}</td><td class="text-end ${cls}">${c.score}</td></tr>`;
  }).join("") || '<tr><td colspan="7" class="text-muted">No data</td></tr>';
  setExpandBtn("umcorr", 25);
}

function renderUmTable(q) {
  const sf = ($("um-tbl-status") && $("um-tbl-status").value) || "ALL";
  const rows = (APP._umRows || []).filter(u =>
    (!q || u.schoolRaw.toLowerCase().includes(q.toLowerCase())) &&
    (sf === "ALL" || u.status === sf))
    .sort((a, b) => a.schoolRaw.localeCompare(b.schoolRaw));
  const st = (s) => s === "Updated" ? '<span class="badge bg-success">Updated</span>' : '<span class="badge bg-warning text-dark">Pending</span>';
  // Grade-Level / Extract columns: Done = green badge, Pending = bold plain text (no fill)
  const done = (b) => b ? '<span class="badge bg-success">Done</span>' : '<span class="um-pending">Pending</span>';
  $("tbl-um").innerHTML = rows.map(u => {
    const refs = [u.gradeFile, u.extractFile].filter(Boolean).map(esc).join("<br>") || "—";
    return `<tr><td>${esc(u.schoolRaw)}</td><td>${esc(orgOf(u.key))}</td><td>${esc(u.sy)}</td>
      <td>${done(u.gradeDone)}</td><td>${done(u.extractDone)}</td><td>${st(u.status)}</td>
      <td class="small text-muted">${refs}</td></tr>`;
  }).join("") || '<tr><td colspan="7" class="text-muted">No schools match.</td></tr>';
}

/* ==================== DOMAIN REGISTRATION ==================== */
// View metadata for each canonical domain status (the scoring/vocabulary lives
// in QBR.domainStatus; colors + badge class stay here as the view concern).
const DR_META = {
  "Active":         { cls: "dr-registered", color: QBR.COLORS.green,  kpi: "green" },
  "Expiring Soon":  { cls: "dr-expiring",   color: QBR.COLORS.orange, kpi: "orange" },
  "Expired":        { cls: "dr-expired",    color: QBR.COLORS.red,    kpi: "red" },
  "For Renewal":    { cls: "dr-expired",    color: "#c94f00",         kpi: "red" },
  "For Deletion":   { cls: "dr-expired",    color: QBR.COLORS.red,    kpi: "red" },
  "Deleted":        { cls: "dr-nodata",     color: "#8a8886",         kpi: "gray" },
  "End Contract":   { cls: "dr-nodata",     color: QBR.COLORS.gray,   kpi: "gray" },
  "Invalid Domain": { cls: "dr-expiring",   color: QBR.COLORS.yellow, kpi: "orange" },
  "Pending":        { cls: "dr-expiring",   color: QBR.COLORS.blue,   kpi: "blue" },
  "Error":          { cls: "dr-nodata",     color: QBR.COLORS.purple, kpi: "gray" },
  "Not Registered": { cls: "dr-nodata",     color: "#c8c6c4",         kpi: "gray" },
};
function drMeta(status) { return DR_META[status] || { cls: "dr-nodata", color: QBR.COLORS.gray, kpi: "gray" }; }
function drDate(ms) { if (ms == null) return "—"; const d = new Date(ms);
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }

function renderDomainReg() {
  const f = APP.filters;
  const rows = APP.model.domainreg
    .filter(d => f.org === "ALL" || orgOf(d.key) === f.org)
    .filter(d => f.school === "ALL" || d.schoolRaw === f.school);
  APP._drRows = rows;
  const n = rows.length, by = s => rows.filter(d => d.status === s).length;
  const active = by("Active");
  const expSoon = by("Expiring Soon"), urgent = rows.filter(d => d.status === "Expiring Soon" && d.tier === "Urgent").length;
  const expired = by("Expired");
  const renew = by("For Renewal") + by("For Deletion");
  const issues = by("Error") + by("Invalid Domain") + by("Pending");
  const actionNeeded = rows.filter(d => d.group === "critical" || d.group === "attention").length;

  $("kpi-dr").innerHTML =
    kpi("Total Domains", fmt(n), "blue") +
    kpi("Active", fmt(active), "green") +
    kpi("Expiring ≤60d", fmt(expSoon), "orange") +
    kpi("Expired", fmt(expired), "red") +
    kpi("Renew / Delete", fmt(renew), "red") +
    kpi("Errors / Invalid / Pending", fmt(issues), "purple");

  // status distribution — horizontal bar, one bar per status, most-common first
  const present = QBR.DR_STATUS_ORDER.filter(s => by(s) > 0);
  const distRows = present.map(s => ({ s, v: by(s) })).sort((a, b) => b.v - a.v);
  if ($("ch-dr-dist")) QBR.chart.hbar("ch-dr-dist", distRows.map(x => x.s), distRows.map(x => x.v),
    distRows.map(x => drMeta(x.s).color), DL("hbar", distRows.map(x => fmt(x.v))));

  $("dr-note").innerHTML = `${n} domains · <b>${active}</b> active · <b>${actionNeeded}</b> need action · ${expired} expired`
    + (urgent ? ` · <b class="text-danger">${urgent}</b> expiring in ≤30 days` : "")
    + ` — read from the tracker's own Status / Remaining Days (validity where stated).`;

  // dynamic status filter: All + Action-Needed + each status present
  const sel = $("dr-status");
  if (sel) {
    const cur = sel.value || "ALL";
    sel.innerHTML = `<option value="ALL">All Statuses</option><option value="__ACTION">Action Needed</option>`
      + present.map(s => `<option value="${escAttr(s)}">${esc(s)}</option>`).join("");
    sel.value = [...sel.options].some(o => o.value === cur) ? cur : "ALL";
  }
  renderDrTable(($("dr-search") || {}).value || "");
}

function renderDrTable(q) {
  const ql = (q || "").toLowerCase();
  const sf = ($("dr-status") && $("dr-status").value) || "ALL";
  const rank = {}; QBR.DR_STATUS_ORDER.forEach((s, i) => rank[s] = i);
  const rows = (APP._drRows || [])
    .filter(d => !ql || d.schoolRaw.toLowerCase().includes(ql) || (d.domain || "").toLowerCase().includes(ql))
    .filter(d => sf === "ALL" ? true : sf === "__ACTION" ? (d.group === "critical" || d.group === "attention") : d.status === sf)
    // worst / most-actionable first (status rank), then soonest / most-overdue (signed days)
    .sort((a, b) => (rank[a.status] - rank[b.status]) || ((a.days ?? 1e9) - (b.days ?? 1e9)) || a.schoolRaw.localeCompare(b.schoolRaw));
  $("tbl-dr").innerHTML = rows.map(d => {
    const meta = drMeta(d.status);
    const badge = `<span class="dr-badge ${meta.cls}">${esc(d.status)}${d.tier ? " · " + esc(d.tier) : ""}</span>`;
    const daysCell = d.days == null ? "—" : (d.days < 0 ? `<span class="text-danger">${d.days}</span>` : d.days);
    const raw = d.action || d.remarks || "";
    const short = raw.length > 64 ? raw.slice(0, 61) + "…" : raw;
    const act = raw ? `<span class="${d.action ? "" : "text-muted"}" title="${escAttr(raw)}">${esc(short)}</span>` : "—";
    return `<tr><td>${esc(d.schoolRaw)}</td><td class="small">${esc(d.domain || "—")}</td>
      <td>${drDate(d.regMs)}</td><td>${drDate(d.expMs)}</td>
      <td class="text-end">${daysCell}</td><td>${badge}</td><td class="small">${act}</td></tr>`;
  }).join("") || '<tr><td colspan="7" class="text-muted">No domains match.</td></tr>';
}

/* ==================== EXECUTIVE REPORT ==================== */
function renderReport() {
  const agg = computeAggregates();
  const rep = QBR.generateReport(APP.model, agg);
  const secHtml = rep.sections.map(s => `<div class="rep-sec"><h5>${esc(s.title)}</h5><p>${s.body}</p></div>`).join("");
  const toneMap = { Critical: "danger", High: "warning", Medium: "info", Low: "secondary" };
  const recHtml = rep.recommendations.map(r =>
    `<li><span class="badge bg-${toneMap[r.level]} me-2">${r.level}</span>${esc(r.text)}</li>`).join("");
  $("report-body").innerHTML = secHtml +
    `<div class="rep-sec"><h5>Automated Recommendations</h5><ul class="rec-list">${recHtml}</ul></div>`;
}

/* ==================== EXPAND HELPERS (Option A) ==================== */
// Return either the top-N slice or the full array, per the table's expand flag.
function expandSlice(key, arr, n) { return APP.expand[key] ? arr : arr.slice(0, n); }
// Sync the toggle button's label to the current expand state.
function setExpandBtn(key, n) {
  const b = document.querySelector('.expand-btn[data-exp="' + key + '"]');
  if (b) b.textContent = APP.expand[key] ? ("Show top " + n + " ▲") : ("Show all ▾");
}

/* ==================== FULL DATA EXPLORER (Option B) ==================== */
function orgFor(r) { return r.org || (APP.model.orgByKey && APP.model.orgByKey[r.key]) || "—"; }
function drLabel(r) { return r.status; }
function tbCell(gb) { return gb == null ? "" : (gb / 1024).toFixed(2); }

// Column registry per dataset: { key, label, get(row), num?, right? }
const FD_DATASETS = {
  risky: { label: "Risky Users & Domain", rows: m => m.risky, cols: [
    { l: "Month", g: r => r.month }, { l: "School", g: r => r.schoolRaw }, { l: "Organization", g: r => orgFor(r) },
    { l: "Quarter", g: r => r.quarter }, { l: "Risky Users", g: r => r.risky, num: 1 },
    { l: "Domain Health", g: r => r.health }, { l: "Error Cause", g: r => r.errorCause } ] },
  security: { label: "Security Data", rows: m => m.security, cols: [
    { l: "School", g: r => r.schoolRaw }, { l: "GDAP", g: r => r.hasGdap ? "Yes" : "No" },
    { l: "Security Default", g: r => r.securityDefault }, { l: "MFA", g: r => r.mfa }, { l: "SSPR", g: r => r.sspr },
    { l: "Email OTP", g: r => r.emailOtp }, { l: "SMS", g: r => r.sms }, { l: "Auth Methods", g: r => r.authMethods },
    { l: "Tenant Status", g: r => r.domainStatus } ] },
  storage: { label: "Storage Data", rows: m => m.storage, cols: [
    { l: "Quarter", g: r => r.quarter }, { l: "School", g: r => r.schoolRaw }, { l: "Organization", g: r => orgFor(r) },
    { l: "OneDrive (TB)", g: r => tbCell(r.onedriveGB), num: 1 }, { l: "SharePoint (TB)", g: r => tbCell(r.sharepointGB), num: 1 },
    { l: "Exchange (TB)", g: r => tbCell(r.exchangeGB), num: 1 }, { l: "Used (TB)", g: r => tbCell(r.usedGB), num: 1 },
    { l: "Pooled (TB)", g: r => tbCell(r.totalGB), num: 1 }, { l: "Utilization %", g: r => r.pct == null ? "" : r.pct.toFixed(1), num: 1 } ] },
  usage: { label: "Office 365 Usage", rows: m => m.usage, cols: [
    { l: "Quarter", g: r => r.quarter }, { l: "School", g: r => r.schoolRaw }, { l: "Organization", g: r => orgFor(r) },
    { l: "Usage %", g: r => r.usagePct == null ? "" : r.usagePct.toFixed(1), num: 1 },
    { l: "OneDrive", g: r => r.onedriveActive, num: 1 }, { l: "SharePoint", g: r => r.sharepointActive, num: 1 },
    { l: "Teams", g: r => r.teamsActive, num: 1 }, { l: "Exchange", g: r => r.exchangeActive, num: 1 },
    { l: "O365 Active", g: r => r.office365Active, num: 1 }, { l: "O365 Total", g: r => r.office365Total, num: 1 } ] },
  canva: { label: "Canva Status", rows: m => m.canva, cols: [
    { l: "Quarter", g: r => r.quarter }, { l: "School", g: r => r.schoolRaw }, { l: "Canva Users", g: r => r.users, num: 1 },
    { l: "Certificate Status", g: r => r.certStatus }, { l: "Certificate Expiry", g: r => r.certExpiry } ] },
  postmaster: { label: "Google Postmaster", rows: m => m.postmaster, cols: [
    { l: "Month", g: r => r.month }, { l: "School", g: r => r.schoolRaw }, { l: "Organization", g: r => orgFor(r) },
    { l: "Domain", g: r => r.domain }, { l: "Reputation", g: r => r.reputation }, { l: "Spam Rate", g: r => r.spamRaw } ] },
  usermgmt: { label: "User Management", rows: m => m.usermgmt, cols: [
    { l: "School", g: r => r.schoolRaw }, { l: "Organization", g: r => orgFor(r) }, { l: "School Year", g: r => r.sy },
    { l: "Status", g: r => r.status }, { l: "Grade-Level Update", g: r => r.gradeDone ? "Done" : "Pending" },
    { l: "Extract Users", g: r => r.extractDone ? "Done" : "Pending" }, { l: "Remarks", g: r => r.remarks } ] },
  domainreg: { label: "Domain Registration", rows: m => m.domainreg, cols: [
    { l: "School", g: r => r.schoolRaw }, { l: "Domain", g: r => r.domain }, { l: "Registration", g: r => drDate(r.regMs) },
    { l: "Expiration", g: r => drDate(r.expMs) }, { l: "Days Remaining", g: r => r.days == null ? "" : r.days, num: 1 },
    { l: "Status", g: r => r.status }, { l: "Remarks", g: r => r.remarks || "" } ] },
};

// Rows for the current dataset, after optional dashboard-filter + search.
function fdRows(ds) {
  const m = APP.model, f = APP.filters;
  let rows = ds.rows(m).slice();
  if (APP.fd.applyFilters) {
    rows = rows.filter(r => {
      if (r.quarter && !qActive(r.quarter)) return false;
      if (f.org !== "ALL" && orgFor(r) !== f.org) return false;
      if (f.school !== "ALL" && r.schoolRaw !== f.school) return false;
      return true;
    });
  }
  const q = APP.fd.search.trim().toLowerCase();
  if (q) rows = rows.filter(r => ds.cols.some(c => { const v = c.g(r); return v != null && String(v).toLowerCase().includes(q); }));
  if (APP.fd.sort) {
    const c = ds.cols[APP.fd.sort.col], dir = APP.fd.sort.dir;
    rows.sort((a, b) => {
      let va = c.g(a), vb = c.g(b);
      if (c.num) { va = parseFloat(va) || 0; vb = parseFloat(vb) || 0; return dir * (va - vb); }
      va = String(va == null ? "" : va).toLowerCase(); vb = String(vb == null ? "" : vb).toLowerCase();
      return dir * (va < vb ? -1 : va > vb ? 1 : 0);
    });
  }
  return rows;
}

function renderFullData() {
  const sel = $("fd-dataset");
  if (sel && !sel.options.length) {
    sel.innerHTML = Object.entries(FD_DATASETS).map(([k, d]) => `<option value="${k}">${esc(d.label)}</option>`).join("");
    sel.value = APP.fd.dataset;
  }
  const ds = FD_DATASETS[APP.fd.dataset]; if (!ds) return;
  const rows = fdRows(ds);
  const size = APP.fd.pageSize;
  const pages = size ? Math.max(1, Math.ceil(rows.length / size)) : 1;
  if (APP.fd.page > pages) APP.fd.page = pages;
  const start = size ? (APP.fd.page - 1) * size : 0;
  const view = size ? rows.slice(start, start + size) : rows;

  const sortMark = (i) => APP.fd.sort && APP.fd.sort.col === i ? (APP.fd.sort.dir > 0 ? " ▲" : " ▼") : "";
  $("fd-head").innerHTML = "<tr>" + ds.cols.map((c, i) =>
    `<th class="fd-th ${c.num ? "text-end" : ""}" data-col="${i}">${esc(c.l)}${sortMark(i)}</th>`).join("") + "</tr>";
  $("fd-body").innerHTML = view.map(r => "<tr>" + ds.cols.map(c => {
    const v = c.g(r); return `<td class="${c.num ? "text-end" : ""}">${v == null || v === "" ? '<span class="text-muted">—</span>' : esc(v)}</td>`;
  }).join("") + "</tr>").join("") || `<tr><td colspan="${ds.cols.length}" class="text-muted">No rows match.</td></tr>`;

  const shownFrom = rows.length ? start + 1 : 0, shownTo = size ? Math.min(start + size, rows.length) : rows.length;
  $("fd-count").textContent = `Showing ${shownFrom}–${shownTo} of ${fmt(rows.length)}` +
    (APP.fd.applyFilters ? " (filtered)" : " (all)");
  $("fd-page").textContent = size ? `Page ${APP.fd.page} / ${pages}` : "All rows";
  $("fd-prev").disabled = !size || APP.fd.page <= 1;
  $("fd-next").disabled = !size || APP.fd.page >= pages;
}

// CSV of the current Full-Data view (all filtered rows, not just the page).
function fdExportCsv() {
  const ds = FD_DATASETS[APP.fd.dataset]; if (!ds) return;
  const rows = fdRows(ds);
  const esc2 = v => { v = v == null ? "" : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  const lines = [ds.cols.map(c => esc2(c.l)).join(",")]
    .concat(rows.map(r => ds.cols.map(c => esc2(c.g(r))).join(",")));
  downloadBlob(`QBR_${APP.fd.dataset}.csv`, lines.join("\r\n"), "text/csv;charset=utf-8");
}

// Full workbook of every dataset (normalized), via SheetJS (already loaded).
function exportAllData() {
  if (!APP.model || typeof XLSX === "undefined") { alert("Load a workbook first."); return; }
  const wb = XLSX.utils.book_new();
  Object.entries(FD_DATASETS).forEach(([k, ds]) => {
    const rows = ds.rows(APP.model);
    if (!rows.length) return;
    const aoa = [ds.cols.map(c => c.l)].concat(rows.map(r => ds.cols.map(c => { const v = c.g(r); return v == null ? "" : v; })));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), ds.label.slice(0, 31));
  });
  XLSX.writeFile(wb, "QBR_Full_Data.xlsx");
}

// Clean, corrected per-tenant × quarter dataset (one tidy table) — the analytical
// feed for the executive snapshot. Uses QBR.buildCleanRows (loader).
function exportCleanData() {
  if (!APP.model || typeof XLSX === "undefined" || typeof QBR.buildCleanRows !== "function") { alert("Load a workbook first."); return; }
  const rows = QBR.buildCleanRows(APP.model), cols = QBR.CLEAN_COLS;
  const aoa = [cols].concat(rows.map(r => cols.map(c => { const v = r[c]; return v == null ? "" : v; })));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = cols.map((c, i) => ({ wch: i === 0 ? 36 : Math.max(9, c.length + 2) }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Clean Dataset");
  XLSX.writeFile(wb, "QBR_Clean_Dataset.xlsx");
}

function downloadBlob(name, text, type) {
  const b = new Blob([text], { type: type }), u = URL.createObjectURL(b);
  const a = document.createElement("a"); a.href = u; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(u), 1000);
}

/* ==================== DATA QUALITY ==================== */
function renderQuality() {
  const m = APP.model;
  if (!m || typeof QBR.audit !== "function") { $("quality-body").innerHTML = "<p class='text-muted'>Upload a workbook to run the data-quality audit.</p>"; return; }
  const a = QBR.audit(m);
  APP._audit = a;
  const sc = a.score;
  const banner = a.banners.map(b =>
    `<div class="dq-banner"><span class="dq-banner-t">${esc(b.title)}</span><span class="dq-banner-d">${esc(b.detail)}</span></div>`).join("");
  const scoreCard = `<div class="dq-score dq-grade-${sc.grade}"><div class="dq-score-num">${sc.value}<span>/100</span></div>
      <div class="dq-score-lbl">Source hygiene · ${esc(sc.label)} · grade ${sc.grade}</div></div>`;
  const kpis =
    kpi("Master Tenants", fmt(a.coverage.master), "blue") +
    kpi("Source Cells to Fix", fmt(a.counts.badCells), a.counts.badCells ? "orange" : "green") +
    kpi("Vocabulary Typos", fmt(a.counts.vocab), a.counts.vocab ? "orange" : "green") +
    kpi("Tenants Missing Org", fmt(a.counts.orgGaps), a.counts.orgGaps ? "orange" : "green");
  const cov = a.coverage.sheets.map(s => `<tr><td>${esc(s.label)}</td><td class="text-end">${fmt(s.schools)}</td></tr>`).join("");
  const missBlock = ["security", "storage", "usage"].map(k => {
    const list = a.coverage.missing[k]; return list.length ? `<div class="small mb-1"><b>${k}</b>: ${list.length} domain tenant(s) missing — <span class="text-muted">${list.slice(0, 10).map(esc).join(", ")}${list.length > 10 ? " …" : ""}</span></div>` : "";
  }).join("");
  const canvaNote = a.coverage.canvaOnly.length ? `<div class="small text-muted mt-1">${a.coverage.canvaOnly.length} Canva-only tenants (no Microsoft footprint) — excluded from Tenant Status by design.</div>` : "";
  const bc = a.badCells.length
    ? `<table class="table table-sm dq-tbl"><thead><tr><th>Cell</th><th>Tenant</th><th>Q</th><th>Column</th><th>Current</th><th>Suggested fix</th><th>Type</th></tr></thead><tbody>${
      a.badCells.map(b => `<tr><td><code>${esc(b.cell)}</code></td><td>${esc(b.tenant)}</td><td>${esc(b.quarter)}</td><td>${esc(b.column)}</td><td class="dq-cur">${esc(b.current)}</td><td class="dq-fix">${esc(b.fix)}</td><td><span class="dq-tag">${esc(b.type)}</span></td></tr>`).join("")
    }</tbody></table>` : `<p class="text-muted">No storage cell issues detected.</p>`;
  const vocab = a.vocab.length
    ? `<table class="table table-sm dq-tbl"><thead><tr><th>Sheet</th><th>Column</th><th>Unexpected value</th><th class="text-end">Count</th></tr></thead><tbody>${
      a.vocab.map(v => `<tr><td>${esc(v.sheet)}</td><td>${esc(v.column)}</td><td><code>${esc(v.value)}</code></td><td class="text-end">${v.count}</td></tr>`).join("")
    }</tbody></table>` : `<p class="text-muted">No controlled-vocabulary violations.</p>`;
  const sin = a.statusInNumber.map(s =>
    `<p class="dq-callout"><b>${esc(s.sheet)} · ${esc(s.column)}:</b> ${s.total} cells hold a status word instead of a number (${Object.entries(s.kinds).map(([k, v]) => esc(k) + " ×" + v).join(", ")}) — these months read as “no data”, not 0.</p>`).join("");
  const sent = a.sentinels.map(s => `<tr><td>${esc(s.sheet)}</td><td class="text-end">${fmt(s.count)}</td><td class="text-end">${fmt(s.ofRows)}</td></tr>`).join("");
  const varNote = a.nameVariants.length ? `<div class="small text-muted mt-2">Name variants (auto-merged by key): ${a.nameVariants.map(v => "“" + v.variants.map(esc).join("” / “") + "”").join("; ")}</div>` : "";
  const encNote = (a.nameEncoding && a.nameEncoding.length) ? `<div class="small mt-2"><span class="badge bg-warning text-dark">Encoding</span> ${a.nameEncoding.length} source name(s) have mojibake (UTF-8 read as Latin-1) — auto-repaired for display, but fix at source: ${a.nameEncoding.map(esc).join("; ")}</div>` : "";
  const og = a.orgGaps.count
    ? `<details><summary>${a.orgGaps.count} tenants have no Organization after the join</summary><div class="small text-muted mt-1">${a.orgGaps.list.map(esc).join(", ")}</div></details>`
    : `<p class="text-muted">Every tenant resolves to an organization.</p>`;

  $("quality-body").innerHTML =
    `<div class="dq-head">${scoreCard}
       <div class="dq-note">This audit runs <b>live on every upload</b>. The dashboard already <b>corrects</b> these at read time (median storage, capacity guards, name normalization); the score reflects the <b>raw workbook</b> so the source can be cleaned. Generated ${esc(a.generatedAt.toLocaleString())}.</div></div>
     ${banner}
     <div class="row row-cols-2 row-cols-md-4 g-3 kpi-row mb-3">${kpis}</div>
     <h6 class="dq-h">Coverage — distinct schools per sheet (master union ${fmt(a.coverage.master)})</h6>
     <div class="row"><div class="col-md-5"><table class="table table-sm dq-tbl"><thead><tr><th>Sheet</th><th class="text-end">Schools</th></tr></thead><tbody>${cov}</tbody></table></div>
       <div class="col-md-7">${missBlock || "<span class='text-muted small'>All domain tenants covered across the MS sheets.</span>"}${canvaNote}</div></div>
     <div class="d-flex align-items-center justify-content-between mt-2 flex-wrap gap-2"><h6 class="dq-h mb-0">Source cells to fix (${a.badCells.length})</h6>
       <button id="dq-export" class="btn btn-sm btn-primary">Export Fix List (.xlsx)</button></div>
     ${bc}
     <h6 class="dq-h mt-3">Controlled-vocabulary violations (auto-normalized, but dirty at source)</h6>${vocab}${varNote}${encNote}
     ${sin ? `<h6 class="dq-h mt-3">Status words in a numeric column</h6>${sin}` : ""}
     <h6 class="dq-h mt-3">Organization gaps</h6>${og}
     <h6 class="dq-h mt-3">Placeholder / sentinel rows (excluded from counts)</h6>
     <table class="table table-sm dq-tbl"><thead><tr><th>Sheet</th><th class="text-end">Sentinel rows</th><th class="text-end">of rows</th></tr></thead><tbody>${sent}</tbody></table>`;

  const btn = $("dq-export"); if (btn) btn.onclick = exportAuditFixList;
}

function exportAuditFixList() {
  const a = APP._audit; if (!a || !window.XLSX) return;
  const ws = XLSX.utils.aoa_to_sheet(QBR.auditFixRows(a));
  ws["!cols"] = [{ wch: 8 }, { wch: 9 }, { wch: 44 }, { wch: 4 }, { wch: 22 }, { wch: 18 }, { wch: 36 }, { wch: 18 }];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Source Fix List");
  XLSX.writeFile(wb, "QBR_Source_Fix_List.xlsx");
}

/* ==================== filters + orchestration ==================== */
function opt(v, label) { return `<option value="${esc(v)}">${esc(label == null ? v : label)}</option>`; }

// Render the multi-select quarter chips ([All] + each quarter present in data).
function renderQuarterChips() {
  const host = $("f-quarter-chips"); if (!host) return;
  const present = (APP.model ? quartersPresent() : []);
  const qs = present.length ? present : ALL_Q;
  const sel = APP.filters.quarters || [];
  const all = isAllQ();
  const chip = (val, label, active) => `<button type="button" class="qchip${active ? " active" : ""}" data-q="${val}">${esc(label)}</button>`;
  host.innerHTML = chip("ALL", "All", all) + qs.map(q => chip(q, q, !all && sel.includes(q))).join("");
}

function populateFilters() {
  const m = APP.model;
  const schools = [...m.master.values()].map(x => x.name).sort();
  const orgs = [...new Set([].concat(
    m.risky.map(r => r.org), m.storage.map(s => s.org), m.usage.map(u => u.org), m.postmaster.map(p => p.org),
    Object.values(m.orgByKey || {})
  ).filter(Boolean))].sort();
  APP._schools = schools; APP._orgs = orgs; // valid-value sets for type-to-search

  renderQuarterChips();
  $("f-month").innerHTML = opt("ALL", "All Months") + QBR.util.MONTHS.map(mo => opt(mo, mo[0] + mo.slice(1).toLowerCase())).join("");
  // type-to-search datalists; empty input == ALL
  $("dl-org").innerHTML = orgs.map(o => `<option value="${esc(o)}"></option>`).join("");
  $("dl-school").innerHTML = schools.map(s => `<option value="${esc(s)}"></option>`).join("");
  $("f-org").value = ""; $("f-school").value = "";
  $("f-secdef").innerHTML = opt("ALL", "All Security Defaults") + ["ENABLED", "DISABLED", "CONDITIONAL ACCESS", "NOT MANAGED"].map(s => opt(s, SD_LABEL[s] || s)).join("");
  // Authentication-method filter for the Security & Risk Detail table (methods
  // actually present in the data + posture presets).
  const amEl = $("f-authmethod");
  if (amEl) {
    const present = (QBR.AUTH_CANON || []).filter(mth => m.security.some(s => s.methods && s.methods.indexOf(mth) >= 0));
    amEl.innerHTML = opt("ALL", "All methods")
      + '<optgroup label="Has method">' + present.map(mth => opt(mth, mth)).join("") + '</optgroup>'
      + '<optgroup label="Posture">' + opt("__PHISH", "Phishing-resistant") + opt("__WEAKONLY", "Weak-only (no strong factor)")
        + opt("__OTPONLY", "Email OTP only") + opt("__NOMFA", "No MFA") + '</optgroup>';
  }
  const syList = [...new Set(m.usermgmt.map(u => u.sy))];
  $("f-sy").innerHTML = (syList.length ? syList : ["—"]).map(s => opt(s)).join("");
  APP.filters.sy = syList[0] || "ALL";
  $("f-status").innerHTML = opt("ALL", "All Statuses") + ["Updated", "Pending"].map(s => opt(s)).join("");
}

function renderAll() {
  if (!APP.model) return;
  captureKpiPrev();
  renderOverview(); renderRisky(); renderDomainHealth(); renderSecurity(); renderGdap(); renderCanva(); renderUsage(); renderStorage();
  renderPostmaster(); renderUserManagement(); renderDomainReg(); renderReport(); renderQuality();
  renderSchool360();
  if (typeof renderInventory === "function") renderInventory();
  if (typeof renderAsset360Panel === "function") renderAsset360Panel();
  if (typeof renderScan === "function") renderScan();
  if (APP.activeTab === "dash-fulldata") { if (APP.fd.applyFilters) APP.fd.page = 1; renderFullData(); }
  makeTablesResizable();
  updateNavBadges();
  refreshThiPane();
}

// Previous quarter for page-KPI deltas — same anchor as the Overview hero strip
// (latest selected quarter vs the quarter before it). No deltas on "All" or
// when a Month is selected (a month is not comparable with a whole quarter).
function kpiPrevQuarter() {
  if (isAllQ() || (APP.filters.month && APP.filters.month !== "ALL")) return null;
  const qs = quartersPresent(), cur = selQuarters().slice(-1)[0], i = qs.indexOf(cur);
  return i > 0 ? qs[i - 1] : null;
}
// Silent previous-quarter pass over the time-scoped pages: charts suppressed,
// kpi() records values; the DOM it writes is overwritten by the real pass.
function captureKpiPrev() {
  const prevQ = kpiPrevQuarter();
  if (!prevQ) { APP._kpi = null; return; }
  const K = APP._kpi = { mode: "capture", prev: {}, prevQ };
  const pq = APP.filters.quarters, pm = APP.filters.month, draw = QBR.draw;
  QBR.draw = function () { return null; };
  APP.filters.quarters = [prevQ]; APP.filters.month = "ALL";
  try {
    [renderRisky, renderDomainHealth, renderCanva, renderUsage, renderStorage, renderPostmaster]
      .forEach(fn => { try { fn(); } catch (e) { /* a page without data just contributes no deltas */ } });
  } finally {
    QBR.draw = draw; APP.filters.quarters = pq; APP.filters.month = pm;
    K.mode = "apply";
  }
}

// Which shared (global-bar) filters each tab actually uses. Tab-specific
// filters (Security Default, School Year, Status, DR status) now live inside
// their own panels, so they are not listed here.
const TAB_FILTERS = {
  "dash-overview":   ["quarter", "month", "org", "school"],
  "dash-risky":      ["quarter", "month", "org", "school"],
  "dash-health":     ["quarter", "month", "org", "school"],
  "dash-sec":        ["quarter", "month", "org", "school"],
  "dash-gdap":       ["org", "school"],
  "dash-canva":      ["quarter", "org", "school"],
  "dash-usage":      ["quarter", "org", "school"],
  "dash-storage":    ["quarter", "org", "school"],
  "dash-postmaster": ["quarter", "month", "org", "school"],
  "dash-usermgmt":   ["org", "school"],
  "dash-domainreg":  ["org", "school"],
  "dash-report":     ["quarter", "month", "org", "school"],
  "dash-quality":    [],
  "dash-fulldata":   ["quarter", "month", "org", "school"],
  "dash-school":     ["quarter"],                 // has its own school picker; org/school filters don't apply
  "dash-inventory":  [],                          // own client/type/status filters inside the panel
  "dash-asset360":   [],                          // single-asset page; has its own serial search
  "dash-scan":       [],                          // scan page; own upload UI, no shared filters
};
const GLOBAL_FILTER_IDS = { quarter: "f-quarter-chips", month: "f-month", org: "f-org", school: "f-school" };

// Show only the shared filters that apply to the active tab.
function applyFilterVisibility(tabId) {
  const applicable = TAB_FILTERS[tabId] || Object.keys(GLOBAL_FILTER_IDS);
  Object.entries(GLOBAL_FILTER_IDS).forEach(([key, id]) => {
    const el = $(id);
    if (el) el.style.display = applicable.includes(key) ? "" : "none";
  });
}

function wireFilters() {
  // dropdown/select filters
  const map = { "f-month": "month", "f-secdef": "secDefault", "f-sy": "sy", "f-status": "status", "f-authmethod": "authMethod" };
  Object.entries(map).forEach(([id, key]) => {
    const el = $(id); if (el) el.addEventListener("change", e => { APP.filters[key] = e.target.value; renderAll(); });
  });
  // quarter multi-select chips ([All] clears; each quarter toggles; selecting
  // all/none normalizes back to "All")
  const qhost = $("f-quarter-chips");
  if (qhost) qhost.addEventListener("click", e => {
    const b = e.target.closest(".qchip"); if (!b) return;
    const q = b.dataset.q;
    let s = (APP.filters.quarters || []).slice();
    if (q === "ALL") s = [];
    else {
      const i = s.indexOf(q);
      if (i >= 0) s.splice(i, 1); else s.push(q);
      if (s.length === 0 || s.length >= 4) s = [];
    }
    APP.filters.quarters = s;
    renderQuarterChips();
    renderAll();
  });
  // type-to-search inputs: empty or non-matching value resolves to ALL
  const bindSearch = (id, key, validList) => {
    const el = $(id); if (!el) return;
    const resolve = () => {
      const v = el.value.trim();
      APP.filters[key] = (v === "" || !validList().includes(v)) ? "ALL" : v;
      renderAll();
    };
    el.addEventListener("change", resolve);
  };
  bindSearch("f-org", "org", () => APP._orgs || []);
  bindSearch("f-school", "school", () => APP._schools || []);

  $("btn-reset").addEventListener("click", () => {
    APP.filters = { quarters: [], month: "ALL", org: "ALL", school: "ALL", secDefault: "ALL", usageCat: "ALL", sy: "ALL", status: "ALL", authMethod: "ALL" };
    populateFilters();
    ["f-month", "f-secdef", "f-status", "f-authmethod"].forEach(id => { if ($(id)) $(id).value = "ALL"; });
    APP.filters.sy = [...new Set(APP.model.usermgmt.map(u => u.sy))][0] || "ALL";
    renderAll(); applyFilterVisibility(APP.activeTab);
  });
}

/* ---- file upload + offline persistence ----
 * The parse pipeline is untouched: every path (drag-drop, picker, cache-restore)
 * ends in the SAME processBuffers() → QBR.loadWorkbooks(buffers). Persistence
 * only caches the raw file bytes and replays them; it never changes parsing.
 * ==========================================================================*/
function readU8(blob) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(new Uint8Array(r.result));
    r.onerror = rej;
    r.readAsArrayBuffer(blob);
  });
}

/* ---- IndexedDB session cache (survives reload on file:// too) ---- */
const IDB_NAME = "qbr-cache", IDB_STORE = "session", IDB_KEY = "workbooks";
function idbOpen() {
  return new Promise((res, rej) => {
    if (!window.indexedDB) return rej(new Error("no-indexeddb"));
    let rq;
    try { rq = indexedDB.open(IDB_NAME, 1); } catch (e) { return rej(e); }
    rq.onupgradeneeded = () => { try { rq.result.createObjectStore(IDB_STORE); } catch (e) {} };
    rq.onsuccess = () => res(rq.result);
    rq.onerror = () => rej(rq.error);
  });
}
function idbPut(val) {
  return idbOpen().then(db => new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(val, IDB_KEY);
    tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
  }));
}
function idbGet() {
  return idbOpen().then(db => new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const rq = tx.objectStore(IDB_STORE).get(IDB_KEY);
    rq.onsuccess = () => res(rq.result || null); rq.onerror = () => rej(rq.error);
  }));
}
function idbDel() {
  return idbOpen().then(db => new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).delete(IDB_KEY);
    tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
  }));
}
function rememberOn() { try { return localStorage.getItem("qbr-remember") !== "0"; } catch (e) { return true; } }
function setRemember(v) { try { localStorage.setItem("qbr-remember", v ? "1" : "0"); } catch (e) {} }

// Reflect saved-state in the header (saved date + Clear button).
function updateSavedState(savedAt) {
  const el = $("saved-state"); if (!el) return;
  if (!savedAt) { el.innerHTML = ""; return; }
  const d = new Date(savedAt);
  el.innerHTML = ` · <span class="saved-note" title="Restored automatically from this browser">Saved ${d.toLocaleDateString()}</span>` +
    ` <button id="btn-clearsaved" class="loaded-clear" title="Remove the saved copy from this browser">Clear</button>`;
  const c = $("btn-clearsaved");
  if (c) c.addEventListener("click", () => { idbDel().catch(() => {}).then(() => updateSavedState(null)); });
}

// Normalize a filename to a dedup key (re-uploading the same name refreshes it).
function fileKey(name) { return String(name || "workbook.xlsx").trim().toLowerCase(); }

// Render the accumulated-workbook list (with per-file refresh ↻ and remove ×).
function renderFileList() {
  const host = $("loaded-files"); if (!host) return;
  const files = APP.files || [];
  if (!files.length) { host.innerHTML = ""; return; }
  host.innerHTML =
    `<div class="loaded-files-head">Workbooks (${files.length})</div>` +
    files.map(it => {
      const k = escAttr(fileKey(it.name));
      return `<div class="loaded-file"><span class="loaded-file-name" title="${escAttr(it.name)}">${esc(it.name)}</span>` +
        `<button type="button" class="loaded-file-r" data-fk="${k}" title="Refresh this workbook — re-pick the file to load its latest data">↻</button>` +
        `<button type="button" class="loaded-file-x" data-fk="${k}" title="Remove this workbook">×</button></div>`;
    }).join("");
}

// Refresh ONE workbook: re-open the picker, then swap the chosen file's bytes
// into that slot — every OTHER loaded workbook (and its position) is preserved.
function promptRefresh(targetKey) {
  const inp = document.createElement("input");
  inp.type = "file"; inp.accept = ".xlsx,.xls,.xlsm"; inp.style.display = "none";
  document.body.appendChild(inp);
  inp.addEventListener("change", () => {
    const f = inp.files && inp.files[0];
    if (f) refreshFileWith(targetKey, f);
    if (inp.parentNode) inp.parentNode.removeChild(inp);
  });
  inp.click();
}
function refreshFileWith(targetKey, file) {
  const cur = APP.files || [];
  const idx = cur.findIndex(it => fileKey(it.name) === targetKey);
  if (idx < 0) return;                     // slot vanished — nothing to refresh
  const newItem = { name: file.name || "workbook.xlsx", blob: file };
  const nk = fileKey(newItem.name);
  // Keep the target's position; drop any OTHER slot that shares the new name.
  const items = [];
  cur.forEach((it, i) => {
    if (i === idx) { items.push(newItem); return; }
    if (fileKey(it.name) === nk) return;   // renamed onto an existing name → dedup
    items.push(it);
  });
  $("upload-status").textContent = "Updating…";
  loadItems(items, {}).then(ok => { if (ok) cacheSession(); })
    .catch(err => { $("upload-status").textContent = "Error: " + err.message; });
}

// THE single processing path — identical for upload and restore.
// Renders from APP.model (parsed) and APP.files (accumulated list), so both
// must be set to their intended values BEFORE this is called.
function processBuffers(buffers, meta) {
  meta = meta || {};
  try { APP.model = QBR.loadWorkbooks(buffers); }
  catch (e) { $("upload-status").textContent = "Error: " + e.message; return false; }
  try { APP.model.inventory = (typeof QBR.parseInventoryBuffers === "function") ? QBR.parseInventoryBuffers(buffers) : null; }
  catch (e) { console.warn("[QBR] inventory parse failed:", e && e.message); APP.model.inventory = null; }
  const s = APP.model.sources;
  const n = Object.values(s).filter(Boolean).length;
  const fc = (APP.files || []).length;
  const chips = Object.entries(s).map(([k, v]) => `<span class="badge bg-${v ? "success" : "secondary"}">${k}</span>`).join("");
  $("upload-status").innerHTML =
    `<button id="loaded-toggle" class="loaded-toggle" aria-expanded="false">Loaded: <b>${n}</b> source${n === 1 ? "" : "s"}` +
      `${fc ? ` · <b>${fc}</b> file${fc === 1 ? "" : "s"}` : ""} <span class="loaded-caret">▾</span></button>` +
    `<span id="loaded-chips" class="loaded-chips d-none">${chips}<span id="loaded-files" class="loaded-files"></span></span>` +
    `<span id="saved-state" class="saved-state"></span>`;
  const lt = $("loaded-toggle");
  if (lt) lt.addEventListener("click", () => {
    const box = $("loaded-chips"), open = !box.classList.contains("d-none");
    box.classList.toggle("d-none", open);
    lt.setAttribute("aria-expanded", String(!open));
    lt.querySelector(".loaded-caret").textContent = open ? "▾" : "▴";
  });
  renderFileList();
  updateSavedState(meta.savedAt || null);
  $("app-body").classList.remove("d-none");
  $("empty-state").classList.add("d-none");
  // NB: filters are wired once in initShell(), not here.
  populateFilters(); renderAll(); applyFilterVisibility(APP.activeTab);
  return true;
}

// Parse + render a full item set, committing APP.files only on a clean parse
// (rolls back on failure so a bad file never wipes the loaded session).
function loadItems(items, meta) {
  const prev = APP.files || [];
  APP.files = items;
  return Promise.all(items.map(it => readU8(it.blob))).then(buffers => {
    const ok = processBuffers(buffers, meta || {});
    if (!ok) APP.files = prev;
    return ok;
  }).catch(err => { APP.files = prev; throw err; });
}

function cacheSession() {
  if (!rememberOn()) return;
  const savedAt = Date.now();
  idbPut({ savedAt, items: (APP.files || []).map(it => ({ name: it.name, blob: it.blob })) })
    .then(() => updateSavedState(savedAt))
    .catch(e => console.warn("[QBR] cache failed:", e && e.message));
}

// Upload handler — ACCUMULATES across separate uploads. New files merge with
// everything already loaded this session (dedup by filename, newest wins), then
// the FULL combined set is re-parsed through the same multi-file path.
function handleFiles(fileList) {
  const incoming = [...fileList];
  if (!incoming.length) return;
  $("upload-status").textContent = "Parsing…";
  const map = new Map();
  (APP.files || []).forEach(it => map.set(fileKey(it.name), it));
  incoming.forEach(f => map.set(fileKey(f.name), { name: f.name || "workbook.xlsx", blob: f }));
  const items = [...map.values()];
  loadItems(items, {}).then(ok => { if (ok) cacheSession(); })
    .catch(err => { $("upload-status").textContent = "Error: " + err.message; });
}

// Remove one accumulated workbook; re-parse the remainder (or reset to empty).
function removeFile(key) {
  const items = (APP.files || []).filter(it => fileKey(it.name) !== key);
  if (!items.length) {
    APP.files = [];
    APP.model = null;
    idbDel().catch(() => {});
    $("upload-status").textContent = "";
    $("app-body").classList.add("d-none");
    $("empty-state").classList.remove("d-none");
    return;
  }
  loadItems(items, {}).then(ok => { if (ok) cacheSession(); })
    .catch(err => { $("upload-status").textContent = "Error: " + err.message; });
}

// On load, replay any cached session through the same accumulation path so
// subsequent uploads add to (not replace) the restored workbooks.
function restoreSession() {
  if (!rememberOn()) return Promise.resolve(false);
  return idbGet().then(sess => {
    if (!sess || !sess.items || !sess.items.length) return false;
    const items = sess.items.map(it => ({ name: it.name || "workbook.xlsx", blob: it.blob }));
    return loadItems(items, { savedAt: sess.savedAt });
  }).catch(e => { console.warn("[QBR] restore failed:", e && e.message); return false; });
}

/* ==================== EXPORT (full deck + image gallery) ====================
 * Both exports need every dashboard rendered at full size, but ten of the
 * eleven panels carry `d-none` at any moment — a hidden canvas has no layout
 * box, so Chart.js leaves it zero-sized and it captures blank. prepareExport()
 * therefore reveals every panel, forces the light theme (dark charts are
 * unreadable on a slide or on paper), resizes every chart, and hands back a
 * restore() that puts the UI back exactly as the user left it.
 * ==========================================================================*/

// Deck order — the single source of truth for both the printed deck and the
// image gallery. Executive Report leads, then risk → security → email →
// adoption → storage → readiness, with Data Quality last as an appendix.
const DECK_ORDER = {
  "dash-overview": 0,
  "dash-report": 1, "dash-risky": 2, "dash-health": 3, "dash-sec": 4, "dash-gdap": 4.5,
  "dash-postmaster": 5, "dash-usage": 6, "dash-canva": 7, "dash-storage": 8,
  "dash-usermgmt": 9, "dash-domainreg": 10, "dash-quality": 11, "dash-fulldata": 12,
};

// Wait n animation frames so layout + Chart.js resize actually settle.
// (A fixed setTimeout races on slow machines; frames are the real signal.)
// Browsers stop firing rAF in a hidden/backgrounded tab, which would strand the
// export mid-flight with every panel revealed — so race a timeout as a floor.
function afterFrames(n, capMs) {
  return new Promise(resolve => {
    let done = false;
    const finish = () => { if (!done) { done = true; resolve(); } };
    let i = 0;
    const tick = () => { if (++i >= n) finish(); else requestAnimationFrame(tick); };
    requestAnimationFrame(tick);
    setTimeout(finish, capMs || 1200);
  });
}

function eachChart(fn) { Object.values(QBR._charts).forEach(c => { try { fn(c); } catch (e) {} }); }

// Temporarily force the light palette into Chart.defaults. Dark charts are
// unreadable on paper and on a projector. Returns a restore function.
// Deliberately writes `data-theme` directly instead of calling setTheme(),
// which would persist to localStorage and clobber the user's preference.
function forceLightTheme() {
  const prev = document.body.getAttribute("data-theme") || "light";
  if (prev === "light") return { changed: false, restore() {} };
  document.body.setAttribute("data-theme", "light");
  QBR.applyChartTheme();
  return {
    changed: true,
    restore() {
      document.body.setAttribute("data-theme", prev);
      QBR.applyChartTheme();
    },
  };
}

// Reveal every dashboard for the printed deck.
// NB: do NOT set fixed pixel width/height on the live canvases. Screen-derived
// pixel sizes are wider than the printed page, which overflows the sheet and
// collides with whatever sits below the chart. Chart.js's `devicePixelRatio`
// option raises the backing-store resolution while leaving the CSS size
// responsive, which is what print needs.
async function prepareExport(opts) {
  const o = opts || {};
  const theme = forceLightTheme();
  const prevTab = APP.activeTab;
  const rehide = [];

  if (theme.changed) renderAll();      // charts bake colors in at construction
  document.querySelectorAll(".dash-panel").forEach(p => {
    if (p.classList.contains("d-none")) { rehide.push(p); p.classList.remove("d-none"); }
  });

  // Reorder BEFORE the charts are painted. Moving a canvas in the DOM makes
  // Chart.js's ResizeObserver fire, which clears the backing store and defers
  // the repaint to rAF — reordering after painting leaves blank charts.
  const restoreOrder = o.deckOrder ? applyDeckOrder() : function () {};

  await afterFrames(2);                // let the revealed/moved panels lay out
  eachChart(c => { c.options.devicePixelRatio = o.dpr || 2; c.resize(); });
  await afterFrames(2);                // let those resizes land
  // Final paint, synchronously: update()/draw() do not wait on rAF, so the
  // canvases are guaranteed painted whether or not the tab is compositing.
  eachChart(c => { c.update("none"); c.draw(); });

  return {
    restore() {
      eachChart(c => { delete c.options.devicePixelRatio; });
      restoreOrder();
      rehide.forEach(p => p.classList.add("d-none"));
      theme.restore();
      APP.activeTab = prevTab;
      renderAll();                     // rebuild every chart at normal scale
    },
  };
}

/* ---- print snapshots (v1.9.0 Phase 3) --------------------------------------
 * Printing live canvases was the root cause of the overlapping charts in
 * Export Tab / Export Deck: Chart.js sizes each canvas to its SCREEN container,
 * then the print layout (≈718px wide, columns stacked) rescales that bitmap
 * with height:auto while the card still reserves the screen height — so big,
 * blurry charts spilled over the cards below them. Instead, every chart in
 * scope is re-rendered offscreen at print size (QBR.renderChartPng — the same
 * path Export Images uses) and shown as a fixed-aspect <img>; the live canvas
 * is hidden only in print. Remove with the returned function (afterprint). */
const PRINT_CHART_W = 700;
function swapChartsForPrint(scope) {
  const added = [];
  (scope || document).querySelectorAll(".chart-wrap canvas").forEach(cv => {
    if (!cv.id || cv.closest("[hidden]") || !QBR._configs || !QBR._configs[cv.id]) return;
    const wrap = cv.closest(".chart-wrap");
    const screenH = (wrap && wrap.clientHeight) || cv.clientHeight || 260;
    const h = Math.max(200, Math.min(420, Math.round(screenH)));
    const url = QBR.renderChartPng(cv.id, { width: PRINT_CHART_W, height: h, dpr: 2 });
    if (!url) return;
    const img = document.createElement("img");
    img.className = "print-chart"; img.alt = ""; img.src = url;
    img.width = PRINT_CHART_W; img.height = h;
    cv.insertAdjacentElement("afterend", img);
    added.push(img);
  });
  document.body.classList.add("print-snap");
  return function removePrintSnapshots() {
    added.forEach(img => img.remove());
    document.body.classList.remove("print-snap");
  };
}

// Human-readable snapshot of the filters an export was taken under, so a deck
// is never ambiguous out of context.
function exportMeta() {
  const f = APP.filters, m = APP.model;
  const nice = v => (v == null || v === "ALL") ? "All" : v;
  const title = s => s[0] + s.slice(1).toLowerCase();
  const rows = [
    ["Quarter", quarterLabel()],
    ["Month", f.month === "ALL" ? "All" : title(f.month)],
    ["Organization", nice(f.org)],
    ["School", nice(f.school)],
    ["Schools in scope", String(m ? m.master.size : 0)],
    ["Generated", new Date().toLocaleString()],
  ];
  if (m && m.usermgmt.length) rows.splice(4, 0, ["School Year", nice(f.sy)]);
  return rows;
}

// Physically reorder the panels into deck order for printing, and return a
// function that puts the DOM back exactly as it was.
// This is deliberately DOM manipulation rather than CSS `order`: flex
// containers do not fragment across printed pages, so ordering with flexbox
// makes Chrome overlap the panels (the Executive Report printing on top of the
// next dashboard) instead of breaking between them.
function applyDeckOrder() {
  const body = $("app-body");
  if (!body) return function () {};
  const original = Array.from(body.children);      // exact pre-export order
  const cover = $("print-cover");
  const panels = original.filter(el => el.classList.contains("dash-panel"));
  const sorted = panels.slice().sort((a, b) =>
    (DECK_ORDER[a.id] || 99) - (DECK_ORDER[b.id] || 99));

  if (cover) body.appendChild(cover);             // cover first…
  sorted.forEach(p => body.appendChild(p));       // …then panels, in deck order

  return function restoreOrder() {
    original.forEach(el => body.appendChild(el)); // re-append in original order
  };
}

function buildPrintCover() {
  const el = $("cover-meta");
  if (el) el.innerHTML = exportMeta().map(([k, v]) =>
    `<tr><td>${esc(k)}</td><td>${esc(v)}</td></tr>`).join("");
}

/* ---- Tier 1: whole-deck PDF via the print stylesheet ---- */
async function exportDeck() {
  const btn = $("btn-deck");
  if (btn) { btn.disabled = true; btn.textContent = "Preparing…"; }
  // 2x keeps chart text crisp on paper without bloating the print spool
  buildPrintCover();
  const { restore } = await prepareExport({ dpr: 2, deckOrder: true });
  const unsnap = swapChartsForPrint(document.getElementById("app-body"));

  const done = () => {
    window.removeEventListener("afterprint", done);
    document.body.removeAttribute("data-print");
    unsnap();
    restore();
    if (btn) { btn.disabled = false; btn.textContent = "Export Deck"; }
  };
  window.addEventListener("afterprint", done);
  document.body.setAttribute("data-print", "ALL");
  window.print();
  // Safety net: some browsers on file:// never fire afterprint.
  setTimeout(() => { if (document.body.getAttribute("data-print") === "ALL") done(); }, 60000);
}

/* ---- Tier 2: chart PNGs ---- */
// 16:9 at 2x — drops onto a PowerPoint slide at native resolution.
const SLIDE_W = 960, SLIDE_H = 540, SLIDE_DPR = 2;

function slugify(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, ""); }

// Re-renders every chart offscreen from its stored config, so hidden panels and
// the live on-screen layout are both irrelevant. Returns {drawn, skipped}
// alongside the grouped images.
function collectChartImages() {
  const groups = [];
  let skipped = 0;
  document.querySelectorAll(".dash-panel").forEach(panel => {
    if (panel.dataset.deck === "skip") return;           // per-school pages (School 360) stay out of the portfolio gallery
    const h = panel.querySelector(".dash-h");
    const title = h ? h.textContent.trim() : panel.id;
    const items = [];
    panel.querySelectorAll("canvas").forEach(cv => {
      const box = cv.closest(".card-box");
      const cap = box && box.querySelector("h6");
      const url = QBR.renderChartPng(cv.id, { width: SLIDE_W, height: SLIDE_H, dpr: SLIDE_DPR });
      if (url) items.push({ caption: cap ? cap.textContent.trim() : cv.id, url, id: cv.id });
      else skipped++;
    });
    if (items.length) groups.push({ title, items, order: DECK_ORDER[panel.id] || 99 });
  });
  groups.sort((a, b) => a.order - b.order);
  groups.skipped = skipped;
  return groups;
}

const GALLERY_CSS = `
  .qbr-gal{font-family:Segoe UI,system-ui,sans-serif;background:#F3F2F1;color:#201F1E;padding:28px 32px}
  .qbr-gal h1{font-size:22px;margin:0 0 4px}
  .qbr-gal .sub{color:#605E5C;font-size:13px;margin-bottom:18px}
  .qbr-gal .tip{background:#EFF6FC;border:1px solid #0078D4;border-radius:8px;padding:10px 14px;font-size:13px;margin-bottom:24px}
  .qbr-gal table.meta{border-collapse:collapse;font-size:12px;margin-bottom:22px;background:#fff;border:1px solid #E1DFDD;border-radius:8px}
  .qbr-gal table.meta td{padding:6px 18px;border-bottom:1px solid #F3F2F1}
  .qbr-gal table.meta td:first-child{color:#605E5C}
  .qbr-gal table.meta td:last-child{font-weight:600}
  .qbr-gal h2{font-size:15px;margin:26px 0 10px;padding-bottom:6px;border-bottom:2px solid #0078D4}
  .qbr-gal .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(420px,1fr));gap:16px}
  .qbr-gal figure{margin:0;background:#fff;border:1px solid #E1DFDD;border-radius:10px;padding:10px}
  .qbr-gal figure img{width:100%;height:auto;display:block;border-radius:4px}
  .qbr-gal figcaption{display:flex;justify-content:space-between;align-items:center;gap:10px;
    font-size:12px;color:#201F1E;margin-top:8px;font-weight:600}
  .qbr-gal figcaption a{font-weight:400;color:#0078D4;text-decoration:none;white-space:nowrap}
  .qbr-gal figcaption a:hover{text-decoration:underline}
  .qbr-gal .closebar{position:sticky;top:0;background:#F3F2F1;padding:8px 0 12px;margin:-28px 0 0;z-index:2}
  @media print{.qbr-gal .tip,.qbr-gal .closebar{display:none}}
`;

// Inner markup shared by the popup window and the in-page fallback overlay.
function galleryBodyHtml(groups, opts) {
  const total = groups.reduce((n, g) => n + g.items.length, 0);
  const meta = exportMeta().map(([k, v]) =>
    `<tr><td>${esc(k)}</td><td>${esc(v)}</td></tr>`).join("");
  const sections = groups.map(g => `
    <h2>${esc(g.title)}</h2>
    <div class="grid">${g.items.map(it => `
      <figure>
        <img src="${it.url}" alt="${esc(it.caption)}" loading="lazy" decoding="async"
             width="${SLIDE_W}" height="${SLIDE_H}" />
        <figcaption>
          <span>${esc(it.caption)}</span>
          <a href="${it.url}" download="qbr-${slugify(g.title)}-${slugify(it.caption)}.png">Download</a>
        </figcaption>
      </figure>`).join("")}</div>`).join("");
  return `
  ${(opts && opts.closable) ? '<div class="closebar"><button id="qbr-gal-close" class="btn btn-sm btn-outline-secondary">Close gallery</button></div>' : ""}
  <h1>Microsoft 365 Tenant QBR — Chart Library</h1>
  <div class="sub">${total} charts across ${groups.length} dashboards, rendered at 3x for projection.</div>
  <div class="tip"><b>To build your deck:</b> right-click any chart → <b>Copy image</b> → paste directly onto a PowerPoint slide.
    Use <b>Download</b> to save a PNG instead. Images are already composited on white, so they sit cleanly on any slide background.</div>
  <table class="meta">${meta}</table>
  ${sections}`;
}

// Prefer a separate window (lets the user keep the dashboard open), but fall
// back to an in-page overlay — pop-ups are commonly blocked, and the app's
// normal home is a file:// page where blocking is even more likely.
function openGallery(groups) {
  const w = (() => { try { return window.open("", "_blank"); } catch (e) { return null; } })();
  if (w && w.document) {
    w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8" />
<title>QBR Charts — copy into your deck</title><style>body{margin:0}${GALLERY_CSS}</style></head>
<body><div class="qbr-gal">${galleryBodyHtml(groups)}</div></body></html>`);
    w.document.close();
    return "window";
  }
  // fallback: overlay inside the app
  document.getElementById("qbr-gallery-overlay")?.remove();
  const host = document.createElement("div");
  host.id = "qbr-gallery-overlay";
  host.style.cssText = "position:fixed;inset:0;z-index:2000;overflow:auto;background:#F3F2F1";
  host.innerHTML = `<style>${GALLERY_CSS}</style><div class="qbr-gal">${galleryBodyHtml(groups, { closable: true })}</div>`;
  document.body.appendChild(host);
  const close = host.querySelector("#qbr-gal-close");
  if (close) close.addEventListener("click", () => host.remove());
  return "overlay";
}

async function exportImages() {
  const btn = $("btn-images");
  if (btn) { btn.disabled = true; btn.textContent = "Rendering…"; }
  // Charts are rebuilt offscreen, so no panel needs revealing and the live
  // dashboard is never disturbed — only the palette has to be forced light.
  const theme = forceLightTheme();
  try {
    await afterFrames(1);
    const groups = collectChartImages();
    if (!groups.length) { alert("No charts available to export. Upload a workbook first."); return; }
    if (groups.skipped) console.warn("[QBR] " + groups.skipped + " chart(s) could not be rendered.");
    openGallery(groups);
  } catch (err) {
    alert("Image export failed: " + err.message);
  } finally {
    theme.restore();
    if (btn) { btn.disabled = false; btn.textContent = "Export Images"; }
  }
}

/* ---- tabs + PDF export ---- */
function setTheme(theme) {
  document.body.setAttribute("data-theme", theme);
  const btn = $("btn-theme");
  if (btn) btn.textContent = theme === "dark" ? "Light mode" : "Dark mode";
  try { localStorage.setItem("qbr-theme", theme); } catch (e) {}
  QBR.applyChartTheme();
  if (APP.model) renderAll(); // redraw charts with new axis/grid colors
}

function initTheme() {
  let saved = "light";
  try { saved = localStorage.getItem("qbr-theme") || "light"; } catch (e) {}
  document.body.setAttribute("data-theme", saved);
  const btn = $("btn-theme");
  if (btn) btn.textContent = saved === "dark" ? "Light mode" : "Dark mode";
  QBR.applyChartTheme();
}

/* ---- sticky header / filter bar / tab strip ---- */
// The three bars stack, so each one's `top` is the summed height of the bars
// above it. Those heights are NOT fixed: both the filter bar and the tab strip
// are flex-wrap, so they grow to two rows on a narrow window, and switching tabs
// hides some filters. Hardcoding the offsets leaves a gap or an overlap, so
// measure them into CSS variables instead.
// Above this share of the viewport, pinning the bars leaves too little room for
// the dashboard, so we let them scroll away instead.
const STICKY_MAX_VIEWPORT_FRACTION = 0.4;

function syncStickyOffsets() {
  const hdr = document.querySelector(".app-header");
  const filt = document.querySelector(".filter-bar");
  const tabs = document.querySelector(".tab-nav");
  const h = el => (el ? el.offsetHeight : 0);
  const s = document.documentElement.style;
  s.setProperty("--hdr-h", h(hdr) + "px");
  // 0 while #app-body is still hidden (pre-upload) — the bar is hidden too.
  s.setProperty("--filter-h", h(filt) + "px");

  // Toggling this only changes `position`, never the bars' heights, so the
  // measurement it depends on stays stable and cannot oscillate.
  const total = h(hdr) + h(filt) + h(tabs);
  const tooTall = window.innerHeight > 0 &&
    total > window.innerHeight * STICKY_MAX_VIEWPORT_FRACTION;
  document.body.classList.toggle("no-sticky", tooTall);
}

function initStickyBars() {
  syncStickyOffsets();
  if (typeof ResizeObserver !== "undefined") {
    const ro = new ResizeObserver(syncStickyOffsets);
    [".app-header", ".filter-bar"].forEach(sel => {
      const el = document.querySelector(sel);
      if (el) ro.observe(el);
    });
  } else {
    window.addEventListener("resize", syncStickyOffsets);  // fallback
  }
}

function initShell() {
  initTheme();
  initStickyBars();
  if ($("app-version") && QBR.VERSION) $("app-version").textContent = "v" + QBR.VERSION;  // header ← single version constant
  wireFilters();   // once — the controls are static in index.html
  const tb = $("btn-theme");
  if (tb) tb.addEventListener("click", () => setTheme(document.body.getAttribute("data-theme") === "dark" ? "light" : "dark"));
  $("file-input").addEventListener("change", e => { handleFiles(e.target.files); e.target.value = ""; });
  // Per-workbook refresh (↻) + remove (×) — delegated on the stable container.
  const us = $("upload-status");
  if (us) us.addEventListener("click", e => {
    const r = e.target.closest(".loaded-file-r");
    if (r) { e.stopPropagation(); promptRefresh(r.dataset.fk); return; }
    const x = e.target.closest(".loaded-file-x");
    if (x) { e.stopPropagation(); removeFile(x.dataset.fk); }
  });
  const drop = $("drop-zone");
  ["dragover", "dragenter"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("drag"); }));
  ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("drag"); }));
  drop.addEventListener("drop", e => handleFiles(e.dataTransfer.files));

  // Remember-data preference + auto-restore of any cached session
  const rem = $("remember-data");
  if (rem) {
    rem.checked = rememberOn();
    rem.addEventListener("change", e => {
      setRemember(e.target.checked);
      if (!e.target.checked) idbDel().catch(() => {}).then(() => updateSavedState(null));
    });
  }
  restoreSession();

  // Microsoft ▾ dropdown grouping (opens on hover, and on click for touch)
  const MS_TABS = ["dash-risky", "dash-health", "dash-sec", "dash-gdap", "dash-usage", "dash-storage", "dash-usermgmt", "dash-domainreg"];
  const msDd = $("ms-dd"), msMenu = $("ms-menu"), msWrap = msDd && msDd.closest(".tab-dd");
  let msHideTimer = null;
  function syncMsDd() { if (msDd) msDd.classList.toggle("active", MS_TABS.includes(APP.activeTab)); }
  function positionMsMenu() { const r = msDd.getBoundingClientRect(); msMenu.style.top = (r.bottom + 2) + "px"; msMenu.style.left = r.left + "px"; }
  function openMsMenu() { if (!msMenu) return; clearTimeout(msHideTimer); positionMsMenu(); msMenu.classList.remove("d-none"); }
  function closeMsMenu() { if (msMenu) msMenu.classList.add("d-none"); }
  function scheduleHide() { clearTimeout(msHideTimer); msHideTimer = setTimeout(closeMsMenu, 180); }
  if (msDd) msDd.addEventListener("click", e => {
    e.stopPropagation();
    if (msMenu.classList.contains("d-none")) openMsMenu(); else closeMsMenu();
  });
  if (msWrap) {
    msWrap.addEventListener("mouseenter", openMsMenu);
    msWrap.addEventListener("mouseleave", scheduleHide);
    msMenu.addEventListener("mouseenter", () => clearTimeout(msHideTimer));
    msMenu.addEventListener("mouseleave", scheduleHide);
  }
  document.addEventListener("click", e => {
    if (msMenu && !msMenu.classList.contains("d-none") && !msMenu.contains(e.target) && !msDd.contains(e.target)) closeMsMenu();
  });

  document.querySelectorAll("[data-tab]").forEach(btn => btn.addEventListener("click", () => {
    document.querySelectorAll("[data-tab]").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    document.querySelectorAll(".dash-panel").forEach(p => p.classList.add("d-none"));
    $(btn.dataset.tab).classList.remove("d-none");
    APP.activeTab = btn.dataset.tab;
    applyFilterVisibility(btn.dataset.tab);
    closeMsMenu(); syncMsDd();
    if (APP.activeTab === "dash-fulldata" && APP.model) renderFullData();
  }));

  // Per-table "Show all" toggles (Option A)
  document.addEventListener("click", e => {
    const eb = e.target.closest(".expand-btn"); if (!eb) return;
    const k = eb.dataset.exp; APP.expand[k] = !APP.expand[k];
    if (APP.model) renderAll();
  });

  // Tenant Health Index drill-down: toggle panel + band-filter chips.
  document.addEventListener("click", e => {
    const tog = e.target.closest(".thi-drill-btn");
    if (tog) { APP.thiDrill.open = !APP.thiDrill.open; if (APP.model) renderOverview(); return; }
    const chip = e.target.closest(".thi-band-btn");
    if (chip && !chip.disabled) { setThiBand(chip.dataset.band); }
  });

  // Click a segment of the health-distribution bar to filter the drill-down.
  const thiCv = $("ch-ov-thi");
  if (thiCv && !thiCv._thiWired) {
    thiCv._thiWired = true; thiCv.style.cursor = "pointer";
    thiCv.addEventListener("click", evt => {
      const ch = QBR._charts && QBR._charts["ch-ov-thi"];
      if (!ch || !ch.getElementsAtEventForMode) return;
      const pts = ch.getElementsAtEventForMode(evt, "nearest", { intersect: true }, false);
      if (!pts || !pts.length) return;
      const ds = ch.data.datasets[pts[0].datasetIndex];
      if (ds && ds.label) setThiBand(ds.label);
    });
  }

  // Full Data Explorer controls (Option B)
  const fdReset = () => { APP.fd.page = 1; if (APP.model) renderFullData(); };
  const fds = $("fd-dataset");
  if (fds) fds.addEventListener("change", e => { APP.fd.dataset = e.target.value; APP.fd.sort = null; fdReset(); });
  const fss = $("fd-search");
  if (fss) fss.addEventListener("input", e => { APP.fd.search = e.target.value; fdReset(); });
  const faf = $("fd-applyfilters");
  if (faf) faf.addEventListener("change", e => { APP.fd.applyFilters = e.target.checked; fdReset(); });
  const fps = $("fd-pagesize");
  if (fps) fps.addEventListener("change", e => { APP.fd.pageSize = +e.target.value; fdReset(); });
  const fprev = $("fd-prev");
  if (fprev) fprev.addEventListener("click", () => { if (APP.fd.page > 1) { APP.fd.page--; renderFullData(); } });
  const fnext = $("fd-next");
  if (fnext) fnext.addEventListener("click", () => { APP.fd.page++; renderFullData(); });
  const fexp = $("fd-export");
  if (fexp) fexp.addEventListener("click", fdExportCsv);
  const fdHead = $("fd-head");
  if (fdHead) fdHead.addEventListener("click", e => {
    const th = e.target.closest(".fd-th"); if (!th) return;
    const col = +th.dataset.col;
    const s = APP.fd.sort;
    APP.fd.sort = (s && s.col === col) ? { col, dir: -s.dir } : { col, dir: 1 };
    renderFullData();
  });
  const bd = $("btn-data");
  if (bd) bd.addEventListener("click", exportAllData);
  const bcl = $("btn-clean");
  if (bcl) bcl.addEventListener("click", exportCleanData);

  const hs = $("health-search");
  if (hs) hs.addEventListener("input", e => renderHealthTable(e.target.value));
  const hst = $("health-status");
  if (hst) hst.addEventListener("change", () => renderHealthTable(($("health-search") || {}).value || ""));
  const pm = $("pm-search");
  if (pm) pm.addEventListener("input", e => renderPmTable(e.target.value));
  const pmr = $("pm-rep");
  if (pmr) pmr.addEventListener("change", () => renderPmTable(($("pm-search") || {}).value || ""));
  const rsv = $("risky-secdef-view");
  if (rsv) rsv.addEventListener("change", () => renderRisky());
  const ut = $("usage-tenant");
  if (ut) ut.addEventListener("change", () => {
    const v = ut.value.trim();
    const rec = (APP.model ? APP.model.usage : []).find(u => u.schoolRaw === v);
    if (rec) selectUsageTenant(rec.key);
  });
  const st = $("stor-tenant");
  if (st) st.addEventListener("change", () => {
    const v = st.value.trim();
    const rec = (APP.model ? APP.model.storage : []).find(s => s.schoolRaw === v);
    if (rec) selectStorageTenant(rec.key);
  });
  const um = $("um-search");
  if (um) um.addEventListener("input", e => renderUmTable(e.target.value));
  const umt = $("um-tbl-status");
  if (umt) umt.addEventListener("change", () => renderUmTable(($("um-search") || {}).value || ""));
  const gs = $("gdap-search");
  if (gs) gs.addEventListener("input", e => renderGdapTable(e.target.value));
  const gst = $("gdap-status");
  if (gst) gst.addEventListener("change", () => renderGdapTable(($("gdap-search") || {}).value || ""));
  const dr = $("dr-search");
  if (dr) dr.addEventListener("input", e => renderDrTable(e.target.value));
  const drs = $("dr-status");
  if (drs) drs.addEventListener("change", () => renderDrTable(($("dr-search") || {}).value || ""));

  $("btn-pdf").addEventListener("click", () => {
    // print only the active panel: mark body with the active tab id; charts
    // print from fixed-size snapshots (see swapChartsForPrint) and are removed
    // on afterprint, or by the safety net if the browser never fires it.
    document.body.setAttribute("data-print", APP.activeTab);
    const theme = forceLightTheme();                // paper is white: print light-theme charts
    if (theme.changed) renderAll();
    const unsnap = swapChartsForPrint($(APP.activeTab));
    let cleaned = false;
    const done = () => {
      if (cleaned) return; cleaned = true; window.removeEventListener("afterprint", done); unsnap();
      if (theme.changed) { theme.restore(); renderAll(); }
    };
    window.addEventListener("afterprint", done);
    window.print();
    setTimeout(done, 60000);                         // safety net: some browsers on file:// never fire afterprint
  });
  const deck = $("btn-deck");
  if (deck) deck.addEventListener("click", exportDeck);
  const imgs = $("btn-images");
  if (imgs) imgs.addEventListener("click", exportImages);
}

/* ==================== SCHOOL 360 (v1.9.0 · 2026-10-03) ====================
 * One page per school: every module's facts for the selected tenant, its
 * Tenant Health Index, a risky-sign-in trend, and rule-based recommended
 * actions. Time-scoped facts follow the Quarter chips only — the page has its
 * own school picker, so the global Organization / School filters are ignored
 * here (TAB_FILTERS hides them). Portfolio comparisons (share, rank, average)
 * are computed across all schools for the same quarters. */
const S360_QIDX = { Q1: 0, Q2: 1, Q3: 2, Q4: 3 };
function s360Latest(rows) {
  return rows.slice().sort((a, b) => (S360_QIDX[b.quarter] == null ? -1 : S360_QIDX[b.quarter]) - (S360_QIDX[a.quarter] == null ? -1 : S360_QIDX[a.quarter]))[0] || null;
}
function s360Date(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") { const d = new Date(Math.round((v - 25569) * 864e5)); return isNaN(d) ? null : d; }
  const d = new Date(v); return isNaN(d) ? null : d;
}
function s360Fmt(d) { return d ? d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" }) : "—"; }
function s360Pill(text, tone) { return `<span class="s360-pill s360-${tone || "neutral"}">${esc(text)}</span>`; }
function s360YesNo(v) {
  if (v == null || v === "") return s360Pill("No data", "none");
  return /^(y|yes|true|enabled|on)$/i.test(String(v).trim()) ? s360Pill("Yes", "ok") : s360Pill("No", "warn");
}
function s360Facts(pairs) {
  return `<dl class="s360-facts">` + pairs.filter(Boolean).map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v == null || v === "" ? '<span class="text-muted">—</span>' : v}</dd>`).join("") + `</dl>`;
}
// Run fn with the Org / School filters cleared (portfolio-wide), restoring them after.
function s360Portfolio(fn) {
  const f = APP.filters, po = f.org, ps = f.school;
  f.org = "ALL"; f.school = "ALL";
  try { return fn(); } finally { f.org = po; f.school = ps; }
}
function openSchool360(key) {
  APP.school360 = key;
  goToTab("dash-school");
  renderSchool360();
}
function s360DefaultKey(thi) {
  const m = APP.model, f = APP.filters;
  if (APP.school360 && m.master.has(APP.school360)) return APP.school360;
  if (f.school && f.school !== "ALL") {
    for (const [k, v] of m.master) if (v.name === f.school || (v.aliases || []).includes(f.school)) return k;
  }
  return thi.perSchool[0] ? thi.perSchool[0].key : (m.master.keys().next().value || null);
}
function renderSchool360() {
  const host = $("s360-body"), m = APP.model;
  if (!host || !m) return;
  const thi = s360Portfolio(() => computeHealthIndex());
  const key = s360DefaultKey(thi);
  APP.school360 = key;
  const names = [...m.master.entries()].map(([k, v]) => ({ k, name: v.name })).sort((a, b) => a.name.localeCompare(b.name));
  const dl = $("dl-s360"); if (dl && dl.childElementCount !== names.length) dl.innerHTML = names.map(n => `<option value="${escAttr(n.name)}"></option>`).join("");
  if (!key) { host.innerHTML = '<p class="text-muted">No schools in the loaded workbooks.</p>'; return; }
  const ms = m.master.get(key) || {}, name = ms.name || key, org = orgOf(key);
  const inp = $("s360-school"); if (inp && document.activeElement !== inp) inp.value = name;
  const inQ = r => r.key === key && qActive(r.quarter);
  const scope = quarterLabel() === "All" ? "all quarters" : quarterLabel();

  // ---- Tenant Health Index ----
  const tp = thi.perSchool.find(x => x.key === key);
  const band = tp ? thiBand(tp.thi) : null;
  let weakDim = null, weakVal = Infinity;
  if (tp) THI_DIMS.forEach(([d]) => { const v = tp.scores[d]; if (v != null && v < weakVal) { weakVal = v; weakDim = d; } });
  const weakName = weakDim ? (THI_DIMS.find(d => d[0] === weakDim) || [])[1] : null;
  const thiRank = tp ? thi.perSchool.length - thi.perSchool.indexOf(tp) : null;   // perSchool is weakest-first

  // ---- Risky sign-ins ----
  const riskyRows = m.risky.filter(inQ);
  const riskyHas = riskyRows.some(r => r.risky != null);
  const riskyTotal = riskyRows.reduce((a, r) => a + (r.risky || 0), 0);
  const portRisky = {}; m.risky.filter(r => qActive(r.quarter) && r.risky != null).forEach(r => { portRisky[r.key] = (portRisky[r.key] || 0) + r.risky; });
  const portTotal = Object.values(portRisky).reduce((a, b) => a + b, 0);
  const riskyRanked = Object.entries(portRisky).filter(e => e[1] > 0).sort((a, b) => b[1] - a[1]);
  const riskyRank = riskyTotal > 0 ? riskyRanked.findIndex(e => e[0] === key) + 1 : null;
  const rb = riskyHas ? QBR.riskBand(riskyTotal) : null;
  const months = QBR.util.MONTHS.map((mo, i) => ({ mo, i, v: riskyRows.filter(r => r.monthIdx === i && r.risky != null) }))
    .filter(x => x.v.length).map(x => ({ mo: x.mo, v: x.v.reduce((a, r) => a + r.risky, 0) }));
  const healthRow = riskyRows.filter(r => r.health).sort((a, b) => b.monthIdx - a.monthIdx)[0];
  const peak = months.slice().sort((a, b) => b.v - a.v)[0];

  // ---- Security ----
  const sec = m.security.find(x => x.key === key) || null;
  const methods = sec ? (sec.methods || []) : [];
  const strong = methods.some(x => QBR.AUTH_STRONG && QBR.AUTH_STRONG[x]);
  const phish = methods.some(x => QBR.AUTH_PHISH_RESISTANT && QBR.AUTH_PHISH_RESISTANT[x]);
  const sdTone = !sec ? "none" : sec.securityDefault === "ENABLED" || sec.securityDefault === "CONDITIONAL ACCESS" ? "ok" : sec.securityDefault === "DISABLED" ? "bad" : "warn";

  // ---- Usage ----
  // latest quarter that actually carries usage numbers (trackers pre-create blank rows for upcoming quarters)
  const uRows = m.usage.filter(inQ);
  const uLatest = s360Latest(uRows.filter(u => u.usagePct != null || u.office365Active != null || u.assignedLicenses != null)) || s360Latest(uRows);
  const portUsage = s360Portfolio(() => dedupeBySchool(applyFilters(m.usage), u => u.usagePct).filter(u => u.usagePct != null));
  const portUsageAvg = avg(portUsage.map(u => u.usagePct));
  const usageLow = QBR.THRESH && QBR.THRESH.USAGE_LOW != null ? QBR.THRESH.USAGE_LOW : 40;
  const svcLine = (a, t) => (uLatest && uLatest[t]) ? `${fmt(uLatest[a] || 0)} of ${fmt(uLatest[t])} <span class="text-muted">(${Math.round((uLatest[a] || 0) / uLatest[t] * 100)}%)</span>` : null;

  // ---- Storage ----
  const stRows = m.storage.filter(inQ).sort((a, b) => (S360_QIDX[a.quarter] || 0) - (S360_QIDX[b.quarter] || 0));
  const stLatest = s360Latest(stRows.filter(r => r.usedGB != null || r.totalGB != null));
  const stPct = storagePct(stLatest), stBand = capacityBand(stPct);

  // ---- Canva ----
  const cvRows = m.canva.filter(inQ);
  const cvLatest = s360Latest(cvRows.filter(c => c.users != null || c.certStatus)) || s360Latest(cvRows);
  const cvUsers = s360Latest(cvRows.filter(c => c.users != null));
  const cvExp = cvLatest ? s360Date(cvLatest.certExpiry) : null;
  const cvDays = cvExp ? Math.round((cvExp - new Date()) / 864e5) : null;

  // ---- Email (Postmaster) ----
  const pmRows = m.postmaster.filter(inQ);
  const pmL = latestScorableReputation(pmRows)[key];
  const pmAny = pmRows.filter(r => r.reputation).sort((a, b) => b.monthIdx - a.monthIdx)[0];
  const pmFlagged = new Set(pmRows.filter(r => r.reputation === "BAD" || r.reputation === "Issues detected").map(r => r.monthIdx)).size;
  const pmVer = (m.pmVerifyByKey || {})[key];
  const repTone = r => !r ? "none" : /^(HIGH)$/i.test(r) ? "ok" : /^(MEDIUM)$/i.test(r) ? "warn" : /^(BAD|LOW|Issues detected)$/i.test(r) ? "bad" : "neutral";

  // ---- Domain registration + user management ----
  const dr = (m.domainreg || []).filter(d => d.key === key);
  const um = (m.usermgmt || []).filter(u => u.key === key);

  // ---- Recommended actions (rule-based, worst first) ----
  const acts = [];
  const act = (pri, text, tab) => acts.push({ pri, text, tab });
  if (sec && sec.securityDefault === "DISABLED") act("Critical", "Enable Security Defaults (or equivalent Conditional Access policies) — this tenant has neither.", "dash-sec");
  if (rb && riskyTotal > (QBR.THRESH.CRITICAL || 100)) act("Critical", `Escalate risky sign-ins: ${fmt(riskyTotal)} detections in ${scope} (${sharePct(riskyTotal, portTotal)} of the portfolio). Review flagged accounts and reset compromised credentials.`, "dash-risky");
  else if (rb && riskyTotal > (QBR.THRESH.HIGH_RISK || 50)) act("High", `Remediate risky sign-ins this week: ${fmt(riskyTotal)} detections in ${scope}.`, "dash-risky");
  if (stBand && stBand.tone === "bad") act("Critical", `Storage is at ${stPct.toFixed(1)}% of capacity — add capacity or clean up OneDrive/SharePoint now.`, "dash-storage");
  else if (stBand) act("High", `Storage is at ${stPct.toFixed(1)}% of capacity — plan cleanup or additional storage.`, "dash-storage");
  dr.filter(d => d.group === "critical").forEach(d => act("High", `Domain ${d.domain || ""}: ${d.status}${d.action ? " — " + d.action : ""}.`, "dash-domainreg"));
  if (pmL && /^(BAD|Issues detected)$/i.test(pmL.rep)) act("High", `Email reputation shows "${pmL.rep}" — check SPF, DKIM and DMARC and review outbound mail.`, "dash-postmaster");
  if (sec && !sec.hasGdap) act("High", "No GDAP relationship on record — set up delegated admin access so the team can support this tenant.", "dash-gdap");
  if (sec && methods.length && !strong) act("Medium", "Sign-in relies on weak methods only (Email OTP / SMS) — roll out Microsoft Authenticator or passkeys.", "dash-sec");
  if (sec && /^no$/i.test(String(sec.sspr || ""))) act("Medium", "Self-service password reset is off — enable SSPR to cut lockout tickets.", "dash-sec");
  if (uLatest && uLatest.usagePct != null && uLatest.usagePct < usageLow) act("Medium", `Microsoft 365 usage is ${uLatest.usagePct.toFixed(1)}% (benchmark ${usageLow}%) — schedule adoption training.`, "dash-usage");
  if (healthRow && /not managed|not connected|incomplete|no services/i.test(healthRow.health)) act("High", `Tenant status is "${healthRow.health}" — restore access / complete setup.`, "dash-health");
  if (cvDays != null && cvDays < 60) act(cvDays < 0 ? "High" : "Medium", cvDays < 0 ? "Canva Education certificate has expired — renew it." : `Canva Education certificate expires in ${cvDays} days — start renewal.`, "dash-canva");
  um.filter(u => u.status !== "Updated").forEach(u => act("Low", `User management for ${u.sy} is still pending.`, "dash-usermgmt"));
  const PRI = { Critical: 0, High: 1, Medium: 2, Low: 3 };
  acts.sort((a, b) => PRI[a.pri] - PRI[b.pri]);

  // ---- Coverage: which sources have this school ----
  const cov = [["Risky / tenant status", riskyRows.length], ["Security", sec ? 1 : 0], ["M365 usage", uLatest ? 1 : 0], ["Storage", stLatest ? 1 : 0],
    ["Canva", cvLatest ? 1 : 0], ["Postmaster", pmRows.length], ["Domain registration", dr.length], ["User management", um.length]];

  // ---- render ----
  const kpis =
    `<div class="row row-cols-2 row-cols-md-3 row-cols-xl-6 g-3 kpi-row">` +
      kpi("Tenant Health", tp ? String(tp.thi) : "—", !band ? "blue" : tp.thi >= 80 ? "green" : tp.thi >= 60 ? "orange" : "red") +
      kpi("Risky Users", riskyHas ? fmt(riskyTotal) : "—", !rb ? "blue" : riskyTotal > (QBR.THRESH.CRITICAL || 100) ? "red" : riskyTotal > (QBR.THRESH.HIGH_RISK || 50) ? "orange" : "green") +
      kpi("Security Defaults", sec ? (SD_LABEL[sec.securityDefault] || sec.securityDefault || "—") : "—", sdTone === "ok" ? "green" : sdTone === "bad" ? "red" : sdTone === "warn" ? "orange" : "blue") +
      kpi("M365 Usage", uLatest && uLatest.usagePct != null ? uLatest.usagePct.toFixed(1) + "%" : "—", uLatest && uLatest.usagePct != null ? (uLatest.usagePct < usageLow ? "orange" : "green") : "blue") +
      kpi("Storage Used", stPct != null ? stPct.toFixed(1) + "%" : "—", stBand ? (stBand.tone === "bad" ? "red" : "orange") : stPct != null ? "green" : "blue") +
      kpi("Email Reputation", pmL ? esc(pmL.rep) : pmAny ? esc(pmAny.reputation) : "—", !pmL ? "blue" : repTone(pmL.rep) === "bad" ? "red" : repTone(pmL.rep) === "ok" ? "green" : "orange") +
    `</div>`;

  const head =
    `<div class="s360-head card-box"${band ? ` style="border-left:4px solid ${band.color}"` : ""}>` +
      `<div class="s360-id"><div class="s360-over">${esc(org && org !== "Unspecified" ? org : "No organization on record")}</div>` +
        `<h4 class="s360-name">${esc(name)}</h4>` +
        `<div class="s360-sub">Facts for <b>${esc(scope)}</b> · ${tp ? `Health rank ${thiRank} of ${thi.perSchool.length} (1 = healthiest)` : "Not enough data for a Tenant Health score"}</div></div>` +
      (tp ? `<div class="s360-score"><div class="s360-score-num">${tp.thi}<span>/100</span></div>` +
        `<span class="badge" style="background:${band.color};color:#fff">${esc(band.label)}</span>` +
        (weakName ? `<div class="s360-weak">Weakest: <b>${esc(weakName)}</b> (${weakVal})</div>` : "") + `</div>` : "") +
    `</div>`;

  const card = (title, body, tab, extraCls) =>
    `<div class="col-lg-4 col-md-6${extraCls ? " " + extraCls : ""}"><div class="card-box s360-card"><div class="d-flex justify-content-between align-items-center"><h6 class="mb-0">${esc(title)}</h6>` +
    (tab ? `<button type="button" class="s360-go" data-s360-tab="${tab}">Open page ›</button>` : "") + `</div>${body}</div></div>`;

  const secBody = !sec ? '<p class="text-muted s360-empty">Not in SECURITY_DATA.</p>' : s360Facts([
    ["Security Defaults", s360Pill(SD_LABEL[sec.securityDefault] || sec.securityDefault || "Unknown", sdTone)],
    ["MFA", s360YesNo(sec.mfa)], ["SSPR", s360YesNo(sec.sspr)],
    ["Sign-in methods", methods.length ? methods.map(x => `<span class="s360-chip">${esc(x)}</span>`).join(" ") : '<span class="text-muted">None recorded</span>'],
    ["Strength", !methods.length ? null : phish ? s360Pill("Phishing-resistant method present", "ok") : strong ? s360Pill("Strong method present", "ok") : s360Pill("Weak methods only", "warn")],
    ["GDAP", sec.hasGdap ? s360Pill(`Granted${(sec.gdapIds || []).length > 1 ? " · " + sec.gdapIds.length + " relationships" : ""}`, "ok") : s360Pill("None", "bad")],
  ]);
  const riskyBody = !riskyHas ? '<p class="text-muted s360-empty">No risky sign-in data for this school in ' + esc(scope) + '.</p>' :
    s360Facts([
      ["Total detections", `<b>${fmt(riskyTotal)}</b> ${QBR.riskBadge(riskyTotal)}`],
      ["Share of portfolio", portTotal ? `${sharePct(riskyTotal, portTotal) || "0%"} of ${fmt(portTotal)}` : null],
      ["Rank", riskyRank ? `#${riskyRank} of ${riskyRanked.length} schools with risky users` : "No risky users"],
      ["Peak month", peak && peak.v ? `${esc(peak.mo.charAt(0) + peak.mo.slice(1).toLowerCase())} (${fmt(peak.v)})` : null],
    ]) + `<div class="chart-wrap s360-chart"><canvas id="ch-s360-risky"></canvas></div>`;
  const tenantBody = s360Facts([
    ["Tenant status", healthRow ? s360Pill(healthRow.health, /^healthy$/i.test(healthRow.health) ? "ok" : /not managed|not connected/i.test(healthRow.health) ? "bad" : "warn") : null],
    ["As of", healthRow ? esc(healthRow.month.charAt(0) + healthRow.month.slice(1).toLowerCase()) : null],
    ["Domain registration", dr.length ? dr.map(d => `${esc(d.domain || "")} ${s360Pill(d.status, d.group === "ok" ? "ok" : d.group === "critical" ? "bad" : d.group === "attention" ? "warn" : "neutral")}${d.days != null ? ` <span class="text-muted">${d.days >= 0 ? d.days + " days left" : Math.abs(d.days) + " days overdue"}</span>` : ""}`).join("<br>") : '<span class="text-muted">No domain workbook row</span>'],
  ]);
  const usageBody = !uLatest ? '<p class="text-muted s360-empty">Not in USAGE_REPORT for ' + esc(scope) + '.</p>' : s360Facts([
    ["Usage %", uLatest.usagePct != null ? `<b>${uLatest.usagePct.toFixed(1)}%</b> <span class="text-muted">· portfolio avg ${portUsageAvg == null ? "—" : portUsageAvg.toFixed(1) + "%"}</span>` : null],
    ["Snapshot", esc(uLatest.quarter || "—")],
    ["Assigned licenses", uLatest.assignedLicenses != null ? fmt(uLatest.assignedLicenses) : null],
    ["Office 365 active", svcLine("office365Active", "office365Total")],
    ["Teams active", svcLine("teamsActive", "teamsTotal")], ["OneDrive active", svcLine("onedriveActive", "onedriveTotal")],
    ["SharePoint active", svcLine("sharepointActive", "sharepointTotal")], ["Exchange active", svcLine("exchangeActive", "exchangeTotal")],
  ]);
  const storBody = !stLatest ? '<p class="text-muted s360-empty">Not in STORAGE_DATA for ' + esc(scope) + '.</p>' :
    s360Facts([
      ["Used / capacity", `<b>${tb(stLatest.usedGB)}</b> of ${tb(stLatest.totalGB)}`],
      ["Utilization", stPct != null ? `<span class="cap-meter${stBand ? " cap-" + stBand.tone : ""}" aria-hidden="true"><span style="width:${Math.min(100, stPct).toFixed(1)}%"></span></span>${stPct.toFixed(1)}%${capBadge(stPct)}` : null],
      ["OneDrive · SharePoint · Exchange", [stLatest.onedriveGB, stLatest.sharepointGB, stLatest.exchangeGB].map(v => v == null ? "—" : tb(v)).join(" · ")],
      ["Snapshot", esc(stLatest.quarter || "—")],
    ]) + (stRows.length > 1 ? `<table class="table table-sm s360-mini"><thead><tr><th>Quarter</th><th class="text-end">Used</th><th class="text-end">Utilization</th></tr></thead><tbody>` +
      stRows.map(r => `<tr><td>${esc(r.quarter || "—")}</td><td class="text-end">${tb(r.usedGB)}</td><td class="text-end">${storagePct(r) == null ? "—" : storagePct(r).toFixed(1) + "%"}</td></tr>`).join("") + `</tbody></table>` : "");
  const canvaBody = !cvLatest ? '<p class="text-muted s360-empty">Not in CANVA_STATUS for ' + esc(scope) + '.</p>' : s360Facts([
    ["Canva users", cvUsers ? `<b>${fmt(cvUsers.users)}</b>${cvUsers !== cvLatest ? ` <span class="text-muted">(${esc(cvUsers.quarter || "—")})</span>` : ""}` : null],
    ["Certificate", cvLatest.certStatus ? s360Pill(cvLatest.certStatus, /active/i.test(cvLatest.certStatus) ? "ok" : "warn") : null],
    ["Expires", cvExp ? `${s360Fmt(cvExp)}${cvDays != null ? ` <span class="text-muted">(${cvDays >= 0 ? cvDays + " days" : "expired"})</span>` : ""}` : null],
    ["Snapshot", esc(cvLatest.quarter || "—")],
  ]);
  const pmBody = !pmRows.length ? '<p class="text-muted s360-empty">Not in GOOGLE_POSTMASTERTOOLS for ' + esc(scope) + '.</p>' : s360Facts([
    ["Domain", pmAny && pmAny.domain ? esc(pmAny.domain) : (pmVer && pmVer.domain ? esc(pmVer.domain) : null)],
    ["Latest reputation", pmL ? s360Pill(pmL.rep, repTone(pmL.rep)) + ` <span class="text-muted">${esc((QBR.util.MONTHS[pmL.mi] || "").slice(0, 3))}</span>` : (pmAny ? s360Pill(pmAny.reputation, "neutral") : null)],
    ["Months with issues", pmFlagged ? `${pmFlagged}` : "0"],
    ["Verification", pmVer ? s360Pill(pmVer.pmTool || "Unknown", /verified/i.test(pmVer.pmTool || "") && !/un/i.test(pmVer.pmTool || "") ? "ok" : "warn") : null],
  ]);
  const umBody = !um.length ? '<p class="text-muted s360-empty">No User Management workbook row.</p>' : s360Facts(um.map(u =>
    [u.sy, `${s360Pill(u.status, u.status === "Updated" ? "ok" : "warn")} <span class="text-muted">Grade file ${u.gradeDone ? "✓" : "—"} · Extract ${u.extractDone ? "✓" : "—"}</span>`]));
  const actBody = acts.length
    ? `<ol class="s360-acts">` + acts.map(a => `<li class="s360-act s360-act-${a.pri.toLowerCase()}"><span class="s360-pri">${a.pri}</span><span class="s360-act-t">${esc(a.text)}</span>` +
        (a.tab ? `<button type="button" class="s360-go" data-s360-tab="${a.tab}" aria-label="Open the related page">›</button>` : "") + `</li>`).join("") + `</ol>`
    : '<p class="text-muted s360-empty">No actions flagged for this school in ' + esc(scope) + '.</p>';
  const covBody = `<ul class="s360-cov">` + cov.map(([k, n]) => `<li class="${n ? "on" : "off"}">${esc(k)}<span>${n ? "Data" : "No data"}</span></li>`).join("") + `</ul>`;

  host.innerHTML = head + kpis +
    `<div class="row g-3 mt-1">` +
      `<div class="col-lg-7"><div class="card-box s360-card"><div class="d-flex justify-content-between align-items-center"><h6 class="mb-0">Risky sign-ins</h6><button type="button" class="s360-go" data-s360-tab="dash-risky">Open page ›</button></div>${riskyBody}</div></div>` +
      `<div class="col-lg-5"><div class="card-box s360-card"><h6>Recommended actions</h6>${actBody}</div></div>` +
      card("Security & identity", secBody, "dash-sec") +
      card("Tenant & domain", tenantBody, "dash-health") +
      card("Microsoft 365 usage", usageBody, "dash-usage") +
      card("Storage", storBody, "dash-storage") +
      card("Canva Education", canvaBody, "dash-canva") +
      card("Email reputation", pmBody, "dash-postmaster") +
      card("User management", umBody, "dash-usermgmt") +
      card("Data coverage", covBody, null, "col-lg-8") +
    `</div>`;

  if (riskyHas && $("ch-s360-risky")) QBR.chart.line("ch-s360-risky", months.map(x => x.mo.slice(0, 3)),
    [{ label: "Risky users", data: months.map(x => x.v), color: QBR.COLORS.red, fill: true }],
    DL("line", months.map(x => x.v ? fmt(x.v) : null)));
  host.querySelectorAll("[data-s360-tab]").forEach(b => b.addEventListener("click", () => goToTab(b.dataset.s360Tab)));
  const ns = $("s360-note"); if (ns) ns.textContent = `Showing ${name}. Pick any school above; Organization and School filters don't apply on this page.`;
}

/* ---- v1.9.0 Phase 3 wiring ------------------------------------------------ */
// Page scope for kpi() delta lookup: each page renderer runs inside
// withKpiPage(), tenant-detail renderers inside noKpiPage() so their per-tenant
// cards never borrow a portfolio delta. Rebinding the function declarations
// keeps every existing call site (renderAll, filters, selects) unchanged.
function withKpiPage(page, fn) {
  return function () { const pg = APP._kpiPage; APP._kpiPage = page; try { return fn.apply(this, arguments); } finally { APP._kpiPage = pg; } };
}
function noKpiPage(fn) { return withKpiPage(null, fn); }
renderRisky = withKpiPage("dash-risky", renderRisky);
renderDomainHealth = withKpiPage("dash-health", renderDomainHealth);
renderSecurity = withKpiPage("dash-sec", renderSecurity);
renderGdap = withKpiPage("dash-gdap", renderGdap);
renderCanva = withKpiPage("dash-canva", renderCanva);
renderUsage = withKpiPage("dash-usage", renderUsage);
renderStorage = withKpiPage("dash-storage", renderStorage);
renderPostmaster = withKpiPage("dash-postmaster", renderPostmaster);
renderUserManagement = withKpiPage("dash-usermgmt", renderUserManagement);
renderDomainReg = withKpiPage("dash-domainreg", renderDomainReg);
renderQuality = withKpiPage("dash-quality", renderQuality);
renderTenantDetail = noKpiPage(renderTenantDetail);
renderStorageDetail = noKpiPage(renderStorageDetail);

// Detail pane: open from a Per-Tenant Health row (click / Enter / Space);
// close via ×, scrim or Escape. Delegated on document (rows are re-rendered).
document.addEventListener("click", e => {
  const row = e.target.closest && e.target.closest(".thi-row[data-thi-key]");
  if (row) { openThiPane(row.dataset.thiKey, row); return; }
  if (e.target.closest && (e.target.closest("#dp-close") || e.target.closest("#dp-scrim"))) closeThiPane();
});
document.addEventListener("keydown", e => {
  const row = e.target.closest && e.target.closest(".thi-row[data-thi-key]");
  if (row && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); openThiPane(row.dataset.thiKey, row); return; }
  if (e.key === "Escape" && APP.thiPane) { e.stopPropagation(); closeThiPane(); }
}, true);

// School 360 picker (type-to-search like the global School filter; commits on change).
document.addEventListener("change", e => {
  if (!e.target || e.target.id !== "s360-school" || !APP.model) return;
  const v = e.target.value.trim(); let hit = null;
  for (const [k, ms] of APP.model.master) if (ms.name === v) { hit = k; break; }
  if (!hit) for (const [k, ms] of APP.model.master) if (ms.name.toLowerCase() === v.toLowerCase() || Array.from(ms.aliases || []).some(a => String(a).toLowerCase() === v.toLowerCase())) { hit = k; break; }
  // same school already shown → don't rebuild (a rebuild on blur would swallow the click that caused the blur)
  if (hit && hit === APP.school360 && $("s360-body") && $("s360-body").children.length) { const n = $("s360-note"); if (n) n.textContent = `Showing ${APP.model.master.get(hit).name}. Pick any school above; Organization and School filters don't apply on this page.`; }
  else if (hit) { APP.school360 = hit; renderSchool360(); }
  else { const n = $("s360-note"); if (n) n.textContent = `No school named "${v}" — pick one from the list.`;
    const cur = APP.school360 && APP.model.master.get(APP.school360); if (cur) e.target.value = cur.name; }
});
document.addEventListener("click", e => {
  if (e.target && e.target.id === "s360-print") { const b = $("btn-pdf"); if (b) b.click(); }
  if (e.target && e.target.id === "s360-filter" && APP.school360 && APP.model) {
    const inp = $("f-school"), ms = APP.model.master.get(APP.school360);
    if (inp && ms) { inp.value = ms.name; inp.dispatchEvent(new Event("change", { bubbles: true })); goToTab("dash-overview"); }
  }
});
renderSchool360 = withKpiPage("dash-school", renderSchool360);

document.addEventListener("DOMContentLoaded", initShell);
