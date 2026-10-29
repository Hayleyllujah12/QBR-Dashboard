/* RCT OpsDesk — Risky Investigation upload validation (js/soc-validate.js)
 *
 * Pure, DOM-free rules for the SOC page's Source 01–05 cards (Node-testable):
 *   - classify()   : which Entra / script export a file is, from its header row + file name
 *   - slotCheck()  : may this file go into this card? (reject + "Move to Source 0X" when it fits another)
 *   - fileStats()  : what a parsed file covers (dates, users, domains, token types, rows)
 *   - crossCheck() : warnings across the loaded files (date range, tenant, month label, 100k cap,
 *                    coverage, AuthDetails pairing, interactive vs non-interactive, Source 03 age/status)
 * Header samples are the real Entra download headers (2026-10-29); see claude/SCOPE_SOC_UPLOAD_VALIDATION.md.
 * Wiring (instant header read, card notes, status bar) lives in js/soc.js.
 */
(function (root) {
  "use strict";

  var SLOT_NAME = { risky: "Source 01", all: "Source 02", auth: "Source 03", noni: "Source 04", noniad: "Source 05" };
  var SLOT_WANTS = {
    risky: "the Risky sign-ins download (Identity Protection › Risky sign-ins): Date (UTC), User, IP address, Location, Risk state",
    all: "the Interactive sign-ins download (Sign-in logs › Download, first file)",
    auth: "the AuthMethodLogs file from the Auth Methods script (UPN, Password, … MethodCount, DefaultSignInMethod)",
    noni: "the Non-interactive sign-ins download (Sign-in logs › Download)",
    noniad: "the Non-interactive sign-ins AuthDetails download (same date range as Source 04)"
  };
  var CAP = 100000;              // Entra: "up to 100,000 sign-in … records per file"
  var COVERAGE_MIN = 0.8;        // below this share → amber
  var TOKEN_MIN_ROWS = 20;       // need this many token-type values before judging 02 vs 04
  var TOKEN_SHARE = 0.6;         // ≥60 % pointing the other way → amber
  var AGE_DAYS = 7;              // Source 03 extracted more than this after the window end → amber

  function norm(h) { return String(h == null ? "" : h).replace(/^﻿/, "").toLowerCase().trim().replace(/\s+/g, " "); }

  /* ---------------- header row ---------------- */
  // First line of a CSV/TSV text → header cells. Handles a UTF-8 BOM, quotes ("a,b" / "a""b"),
  // and comma / tab / semicolon delimiters (whichever splits the line into most cells).
  function parseHeaderLine(text) {
    var t = String(text || "").replace(/^﻿/, "");
    var end = t.search(/\r?\n/); var line = end < 0 ? t : t.slice(0, end);
    function split(d) {
      var out = [], cur = "", q = false;
      for (var i = 0; i < line.length; i++) {
        var c = line[i];
        if (q) { if (c === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += c; }
        else if (c === '"') q = true;
        else if (c === d) { out.push(cur); cur = ""; }
        else cur += c;
      }
      out.push(cur); return out;
    }
    var best = [line];
    [",", "\t", ";"].forEach(function (d) { var s = split(d); if (s.length > best.length) best = s; });
    return best.map(function (s) { return s.trim(); }).filter(function (s) { return s !== ""; });
  }
  // Header cells of parsed rows (SheetJS objects): union of keys of the first rows; drops the "_1"
  // suffix SheetJS gives a duplicated column ("Incoming token type" appears twice in Entra exports).
  function headersOfRows(rows) {
    var seen = {}, out = [];
    for (var i = 0; i < Math.min(rows ? rows.length : 0, 50); i++) for (var k in rows[i]) {
      var base = String(k).replace(/_\d+$/, "");
      if (!seen[base]) { seen[base] = 1; out.push(base); }
    }
    return out;
  }

  /* ---------------- file name ---------------- */
  var NAME_KINDS = [
    [/NonInteractive[ _-]?SignIns[ _-]?AuthDetails/i, "authDetails", "non"],
    [/Interactive[ _-]?SignIns[ _-]?AuthDetails/i, "authDetails", "int"],
    [/NonInteractive[ _-]?SignIns/i, "signins", "non"],
    [/Interactive[ _-]?SignIns/i, "signins", "int"],
    [/(Application|ServicePrincipal)[ _-]?SignIns/i, "appSignins", null],
    [/(ManagedIdentity|MSI)[ _-]?SignIns|ManagedIdentity/i, "appSignins", null],
    [/Risky[ _-]?Sign[ _-]?ins/i, "riskySignins", null],
    [/Risky[ _-]?Users/i, "riskyUsers", null],
    [/AuthMethodLogs?/i, "authMethods", null]
  ];
  var MONTHS = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11 };
  var MON_RE = /^(JAN(UARY)?|FEB(RUARY)?|MAR(CH)?|APR(IL)?|MAY|JUNE?|JULY?|AUG(UST)?|SEPT?(EMBER)?|OCT(OBER)?|NOV(EMBER)?|DEC(EMBER)?)$/i;
  // "SCHOOL_MON_NonInteractiveSignIns_2026-10-22_2026-10-29.csv" →
  //   { kind:"signins", variant:"non", school:"SCHOOL", mon:"MON"→month index, from:"2026-10-22", to:"2026-10-29" }
  function parseName(filename) {
    var name = String(filename || "").replace(/^.*[\\/]/, "").replace(/\.(csv|xlsx|xls|txt)$/i, "");
    var out = { kind: null, variant: null, school: null, mon: null, month: null, from: null, to: null };
    var at = -1;
    for (var i = 0; i < NAME_KINDS.length; i++) {
      var m = NAME_KINDS[i][0].exec(name);
      if (m) { out.kind = NAME_KINDS[i][1]; out.variant = NAME_KINDS[i][2]; at = m.index; break; }
    }
    var d = /(\d{4}-\d{2}-\d{2})[ _-]+(\d{4}-\d{2}-\d{2})/.exec(name);
    if (d) { out.from = d[1]; out.to = d[2]; }
    var prefix = at > 0 ? name.slice(0, at) : (at < 0 && d ? name.slice(0, d.index) : "");
    var toks = prefix.split(/[_\s\-()]+/).filter(Boolean);
    toks.forEach(function (tk) {
      if (out.mon == null && MON_RE.test(tk)) { out.mon = tk.toUpperCase(); out.month = MONTHS[tk.slice(0, 3).toUpperCase()]; }
      else if (out.school == null && /^[A-Za-z][A-Za-z0-9.]*$/.test(tk) && !/^(SOC|signins?|risky|logs?)$/i.test(tk)) out.school = tk.toUpperCase();
    });
    return out;
  }

  /* ---------------- what is this file? ---------------- */
  var KIND_LABEL = {
    riskySignins: "Risky sign-ins download",
    riskyUsers: "Risky users list (Identity Protection › Risky users)",
    signins: "sign-in log download",
    signins_int: "Interactive sign-ins download",
    signins_non: "Non-interactive sign-ins download",
    authDetails: "AuthDetails download",
    authDetails_int: "Interactive sign-ins AuthDetails download",
    authDetails_non: "Non-interactive sign-ins AuthDetails download",
    appSignins: "app / managed-identity sign-in download (no user)",
    authMethods: "AuthMethodLogs file (Auth Methods script)",
    unknown: "file"
  };
  function kindOfHeaders(headers) {
    var H = {}; (headers || []).forEach(function (h) { H[norm(h)] = 1; });
    var has = function () { for (var i = 0; i < arguments.length; i++) if (H[arguments[i]]) return true; return false; };
    if (has("service principal id", "service principal name", "credential key id")) return "appSignins";
    if (has("risk processing state", "risk last updated (utc)")) return "riskyUsers";
    if (has("risk state", "riskstate")) return "riskySignins";
    if (has("upn", "userprincipalname", "user principal name") &&
        has("methodcount", "defaultsigninmethod", "authenticatorapp", "fido2", "windowshello", "temporaryaccesspass")) return "authMethods";
    if (has("request id") && has("authentication method") && has("succeeded", "result detail", "requirement") &&
        !has("user agent", "username", "user")) return "authDetails";
    if (has("user", "username", "user display name", "name") && has("date (utc)", "date", "datetime")) return "signins";
    return "unknown";
  }
  // → { kind, variant: "int"|"non"|null, label, name: parseName(), headers }
  function classify(headers, filename) {
    var name = parseName(filename), kind = kindOfHeaders(headers);
    var variant = (name.kind === kind) ? name.variant : null;
    var label = KIND_LABEL[kind + (variant ? "_" + variant : "")] || KIND_LABEL[kind];
    return { kind: kind, variant: variant, label: label, name: name, headers: headers || [] };
  }
  // Columns the investigation engine needs per card (an unrecognised layout with these still loads).
  var REQUIRED = {
    risky: [["risk state", "riskstate"], ["user", "name", "user display name"]],
    all: [["user", "username", "user display name", "name"], ["date (utc)", "date", "datetime"]],
    auth: [["upn", "userprincipalname", "user principal name"]],
    noni: [["user", "name", "user display name"], ["date (utc)", "date", "datetime"]],
    noniad: [["request id"], ["authentication method", "result detail"]]
  };
  function hasRequired(slot, headers) {
    var H = {}; (headers || []).forEach(function (h) { H[norm(h)] = 1; });
    return (REQUIRED[slot] || []).every(function (alts) { return alts.some(function (a) { return H[a]; }); });
  }
  // Which card a classified file belongs in (null = none).
  function homeSlot(c) {
    switch (c.kind) {
      case "riskySignins": return "risky";
      case "authMethods": return "auth";
      case "signins": return c.variant === "non" ? "noni" : "all";
      case "authDetails": return c.variant === "int" ? null : "noniad";
      default: return null;
    }
  }
  // → { ok, reject?, moveTo?, warn?, message }
  function slotCheck(slot, c, info) {
    info = info || {};
    var here = SLOT_NAME[slot];
    if (info.size === 0) return { ok: false, reject: true, message: "The file is empty (0 bytes) — download it again." };
    if (!c.headers.length) return { ok: false, reject: true, message: "No header row found — this doesn't look like a CSV or Excel export." };
    var accepted = {
      risky: c.kind === "riskySignins",
      all: c.kind === "signins" && c.variant !== "non",
      auth: c.kind === "authMethods",
      noni: c.kind === "signins" && c.variant !== "int",
      noniad: c.kind === "authDetails" && c.variant !== "int"
    }[slot];
    if (accepted) return { ok: true, message: "" };
    if (c.kind === "unknown") {
      if (hasRequired(slot, c.headers)) return { ok: true, warn: true, message: "Not a layout this page recognises — it loaded, but check it is " + SLOT_WANTS[slot] + "." };
      return { ok: false, reject: true, message: "Not recognised. " + here + " needs " + SLOT_WANTS[slot] + "." };
    }
    var home = homeSlot(c);
    var msg = "This is the " + c.label + ". " + here + " needs " + SLOT_WANTS[slot] + ".";
    if (c.kind === "riskyUsers") msg = "This is the Risky users list — " + here + " needs the Risky sign-ins download (Identity Protection › Risky sign-ins › Download).";
    if (c.kind === "authDetails" && c.variant === "int") msg = "This is the Interactive sign-ins AuthDetails download — Risky Investigation doesn't use it (Source 05 takes the Non-interactive AuthDetails).";
    if (c.kind === "appSignins") msg = "This is an app / managed-identity sign-in download (no user) — Risky Investigation doesn't use it.";
    if (slot === "auth" && c.kind === "authDetails") msg = "This is a per-sign-in AuthDetails download. " + here + " needs the per-user AuthMethodLogs from the Auth Methods script.";
    return { ok: false, reject: true, moveTo: home && home !== slot ? home : null, message: msg };
  }

  /* ---------------- what a parsed file covers ---------------- */
  function keyMap(rows) {
    var m = {};
    for (var i = 0; i < Math.min(rows.length, 50); i++) for (var k in rows[i]) { var n = norm(String(k).replace(/_\d+$/, "")); if (!(n in m)) m[n] = k; }
    return m;
  }
  function pick(m) { for (var i = 1; i < arguments.length; i++) if (m[arguments[i]] != null) return m[arguments[i]]; return null; }
  function toDate(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var d = new Date(v); return isNaN(d.getTime()) ? null : d;
  }
  function domainOf(v) { var s = String(v || "").trim().toLowerCase(), i = s.lastIndexOf("@"); return i > 0 ? s.slice(i + 1) : null; }
  // rows: parsed objects; meta: { totalRows } (Source 04 is pre-filtered to risky users — pass the full count)
  function fileStats(kind, rows, meta) {
    rows = rows || []; var m = keyMap(rows);
    var kDate = pick(m, "date (utc)", "date", "datetime"), kTs = pick(m, "timestamp");
    var kUser = pick(m, "user", "name", "user display name"), kUpn = pick(m, "username", "user principal name", "upn", "userprincipalname");
    var kTok = pick(m, "incoming token type"), kReq = pick(m, "request id"), kRisk = pick(m, "risk state", "riskstate");
    var kStatus = pick(m, "status");
    var s = { rows: rows.length, totalRows: (meta && meta.totalRows != null) ? meta.totalRows : rows.length,
      min: null, max: null, users: {}, atRisk: {}, userUpn: {}, domains: {}, upns: {}, reqIds: null,
      tokNone: 0, tokRefresh: 0, statusBad: {}, tsMax: null };
    if (kReq && (kind === "signins")) s.reqIds = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var d = kDate ? toDate(r[kDate]) : null;
      if (d) { var t = d.getTime(); if (s.min == null || t < s.min) s.min = t; if (s.max == null || t > s.max) s.max = t; }
      var u = kUser ? r[kUser] : null, uk = u ? String(u).toLowerCase().trim() : null;
      if (uk) s.users[uk] = 1;
      var up = kUpn ? r[kUpn] : null;
      if (up) { var dm = domainOf(up); if (dm) s.domains[dm] = 1; if (uk && !s.userUpn[uk]) s.userUpn[uk] = String(up).trim().toLowerCase(); }
      if (kind === "riskySignins" && uk && kRisk && /at risk/i.test(String(r[kRisk] || ""))) s.atRisk[uk] = 1;
      if (s.reqIds && r[kReq]) s.reqIds[String(r[kReq]).trim()] = 1;
      if (kTok && r[kTok] != null && r[kTok] !== "") {
        var tv = String(r[kTok]).toLowerCase();
        if (/^none$/.test(tv)) s.tokNone++; else if (/refresh/.test(tv)) s.tokRefresh++;
      }
      if (kind === "authMethods") {
        if (up) s.upns[String(up).trim().toLowerCase()] = 1;
        var st = kStatus ? String(r[kStatus] || "").trim() : "";
        if (st && !/^success$/i.test(st)) s.statusBad[st] = (s.statusBad[st] || 0) + 1;
        var ts = kTs ? toDate(r[kTs]) : null;
        if (ts && (s.tsMax == null || ts.getTime() > s.tsMax)) s.tsMax = ts.getTime();
      }
    }
    return s;
  }

  /* ---------------- across the loaded files ---------------- */
  var DAY = 86400000;
  function ymd(t) { return t == null ? "?" : new Date(t).toISOString().slice(0, 10); }
  function nice(t) { if (t == null) return "?"; var d = new Date(t); return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][d.getUTCMonth()] + " " + d.getUTCDate(); }
  function span(s) { return nice(s.min) + "–" + nice(s.max); }
  function keys(o) { return Object.keys(o || {}); }
  function pct(a, b) { return b ? Math.round(a / b * 100) : 0; }
  // files: { slot: { name: parseName(), kind, variant, stats } }; extra: { authMapIds: {requestId:1} } (Source 05)
  // → { slot: ["warning", …] }
  function crossCheck(files, extra) {
    extra = extra || {};
    var W = {}; var add = function (slot, msg) { (W[slot] = W[slot] || []).push(msg); };
    var ref = files.all || files.risky; var refSlot = files.all ? "all" : (files.risky ? "risky" : null);
    Object.keys(files).forEach(function (slot) {
      var f = files[slot], s = f.stats || {}, n = f.name || {};
      // the name's own dates vs the rows (more than a day outside)
      if (n.from && n.to && s.min != null && s.max != null && slot !== "auth") {
        var nf = Date.parse(n.from + "T00:00:00Z"), nt = Date.parse(n.to + "T23:59:59Z");
        if (s.min < nf - DAY || s.max > nt + DAY) add(slot, "Its rows cover " + span(s) + " but the file name says " + n.from + " to " + n.to + ".");
      }
      // month label vs the name's dates
      if (n.month != null && (n.from || n.to)) {
        // the label may name the month covered OR the month being reported (a "SEP" audit run on Oct 2–9)
        var ms = [n.from, n.to].filter(Boolean).map(function (x) { return +x.slice(5, 7) - 1; });
        ms.push((ms[0] + 11) % 12);
        if (ms.indexOf(n.month) < 0) add(slot, "The name says " + n.mon + " but its dates are " + n.from + " to " + n.to + ".");
      }
      // 100,000-row cap (sign-in log downloads)
      if ((f.kind === "signins" || f.kind === "authDetails") && (s.totalRows || 0) >= CAP)
        add(slot, "It has " + (s.totalRows).toLocaleString() + " rows — the Entra download limit — so it is probably cut off" +
          (s.min != null ? "; it only reaches back to " + nice(s.min) : "") + ". Narrow the date range or download in parts.");
      if (!ref || slot === refSlot) return;
      var r = ref.stats || {}, rn = ref.name || {};
      // date range vs the reference file (Source 02, else 01)
      if (slot !== "auth" && s.min != null && r.min != null && (s.max < r.min - DAY || s.min > r.max + DAY))
        add(slot, "Covers " + span(s) + "; " + SLOT_NAME[refSlot] + " covers " + span(r) + " — a different week?");
      else if (slot !== "auth" && n.from && rn.from && (n.from !== rn.from || n.to !== rn.to))
        add(slot, "Dated " + n.from + " to " + n.to + "; " + SLOT_NAME[refSlot] + " is " + rn.from + " to " + rn.to + ".");
      // same school code
      if (n.school && rn.school && n.school !== rn.school)
        add(slot, "School code " + n.school + " doesn't match " + SLOT_NAME[refSlot] + " (" + rn.school + ") — a file from another tenant?");
    });
    // shared email domain (tenant) between the files that carry UPNs
    var withDom = Object.keys(files).filter(function (sl) { return keys((files[sl].stats || {}).domains).length; });
    if (withDom.length > 1) {
      var base = withDom.indexOf("all") >= 0 ? "all" : withDom[0], bd = (files[base].stats || {}).domains;
      withDom.forEach(function (sl) {
        if (sl === base) return;
        var ds = keys(files[sl].stats.domains), shared = ds.some(function (d) { return bd[d]; });
        if (!shared) add(sl, "No email domain in common with " + SLOT_NAME[base] + " (" + keys(bd).slice(0, 2).join(", ") + " vs " + ds.slice(0, 2).join(", ") + ") — another tenant?");
      });
    }
    // risky users (Source 01) found in the interactive log (Source 02)
    if (files.risky && files.all) {
      var ar = keys(files.risky.stats.atRisk), inAll = ar.filter(function (u) { return files.all.stats.users[u]; });
      if (ar.length && inAll.length / ar.length < COVERAGE_MIN)
        add("all", "Only " + inAll.length + " of " + ar.length + " at-risk users from Source 01 appear in this log — is it the same tenant and date range?");
    }
    // interactive vs non-interactive by content (token type)
    ["all", "noni"].forEach(function (sl) {
      var f = files[sl]; if (!f) return; var s = f.stats, n = s.tokNone + s.tokRefresh;
      if (n < TOKEN_MIN_ROWS) return;
      if (sl === "all" && s.tokRefresh / n >= TOKEN_SHARE)
        add(sl, pct(s.tokRefresh, n) + "% of rows reuse a refresh token — this looks like a Non-interactive export (Source 04).");
      if (sl === "noni" && s.tokNone / n >= TOKEN_SHARE)
        add(sl, pct(s.tokNone, n) + "% of rows have no incoming token — this looks like an Interactive export (Source 02).");
    });
    // Source 05 pairs with Source 04 (and not with Source 02)
    if (files.noniad && extra.authMapIds) {
      var ids = extra.authMapIds;
      var hit = function (set) { var t = 0, h = 0; for (var k in set) { t++; if (ids[k]) h++; } return { t: t, h: h }; };
      var h4 = files.noni && files.noni.stats.reqIds ? hit(files.noni.stats.reqIds) : null;
      var h2 = files.all && files.all.stats.reqIds ? hit(files.all.stats.reqIds) : null;
      if (h2 && h2.t && h2.h / h2.t >= 0.5 && (!h4 || !h4.t || h4.h / h4.t < 0.1))
        add("noniad", "Its Request IDs match the interactive sign-ins (Source 02), not Source 04 — this is probably the Interactive AuthDetails file.");
      else if (h4 && h4.t && h4.h / h4.t < COVERAGE_MIN)
        add("noniad", "Auth details found for only " + pct(h4.h, h4.t) + "% of the Source 04 sign-ins — same date range as Source 04?");
    }
    // Source 03: coverage of the at-risk users, failed lookups, extraction age
    if (files.auth) {
      var a = files.auth.stats;
      var bad = keys(a.statusBad).map(function (k) { return a.statusBad[k] + " " + k; });
      if (bad.length) add("auth", bad.join(", ") + " — those users have no auth-method data.");
      if (files.risky) {
        var upnOf = (files.all && files.all.stats.userUpn) || {};
        var risky = keys(files.risky.stats.atRisk), known = risky.filter(function (u) { return upnOf[u] || files.risky.stats.userUpn[u]; });
        var found = known.filter(function (u) { return a.upns[upnOf[u] || files.risky.stats.userUpn[u]]; });
        if (known.length && found.length / known.length < COVERAGE_MIN)
          add("auth", "Only " + found.length + " of " + known.length + " at-risk users have a row here — run the Auth Methods script for the current risky users.");
      }
      var end = ref && ref.stats && ref.stats.max;
      if (a.tsMax != null && end != null && a.tsMax - end > AGE_DAYS * DAY)
        add("auth", "Extracted " + Math.round((a.tsMax - end) / DAY) + " days after the investigation window (" + nice(end) + ") — methods may have changed since.");
    }
    return W;
  }

  var api = { parseHeaderLine: parseHeaderLine, headersOfRows: headersOfRows, parseName: parseName, classify: classify,
    slotCheck: slotCheck, homeSlot: homeSlot, fileStats: fileStats, crossCheck: crossCheck,
    SLOT_NAME: SLOT_NAME, CAP: CAP, _ymd: ymd };
  root.QBR = root.QBR || {}; root.QBR.socValidate = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
