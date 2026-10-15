/* ============================================================================
 * RCT OpsDesk — Discreet Admin Panel: per-module feature flags (js/admin.js)
 * v1.32.0 (2026-10-09)
 *
 * Hidden by default. Double-click the version label ("v1.21.0" in the top bar or the sidebar
 * footer) to reveal the Admin section (v1.32.0; keyboard shortcuts removed at the user's request).
 * Optional password gate: first use prompts to set one; only the SHA-256
 * hash is stored (localStorage "qbr-admin-pw"). Deterrent against casual
 * snoopers — not real security (client-side, bypassable via devtools).
 * Each module can be toggled on/off; hidden modules keep all their data.
 * Flags persist in localStorage ("qbr-feature-flags").
 * Appearance (from Mars's patch 4/4, reworked at merge): light/dark mode, page
 * background, card, accent and heading colours, body-text and heading size
 * (90–150 %). Stored in localStorage ("qbr-theme-custom"), applied on <body>
 * so it works in both themes; text sizes use css/type-scale.css (generated).
 * Custom appearance is removed while printing/exporting, then restored.
 * ============================================================================ */
(function () {
  "use strict";
  const QBR = (window.QBR = window.QBR || {});

  const STORE_KEY = "qbr-feature-flags";
  const PW_KEY = "qbr-admin-pw";

  /* ---------- password (SHA-256 hash, never plaintext) -------------------- */
  // crypto.subtle needs a secure context (https/localhost); file:// falls back
  // to cyrb53 (non-crypto but fine for a casual-snooper deterrent).
  function cyrb53(str, seed) {
    let h1 = 0xdeadbeef ^ seed, h2 = 0x41c6ce57 ^ seed;
    for (let i = 0; i < str.length; i++) {
      const ch = str.charCodeAt(i);
      h1 = Math.imul(h1 ^ ch, 2654435761);
      h2 = Math.imul(h2 ^ ch, 1597334677);
    }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (h2 >>> 0).toString(16).padStart(8, "0") + (h1 >>> 0).toString(16).padStart(8, "0");
  }

  async function hashPw(plain) {
    try {
      if (window.crypto && window.crypto.subtle && window.isSecureContext !== false) {
        const buf = await window.crypto.subtle.digest("SHA-256",
          new TextEncoder().encode("qbr-admin::" + plain));
        return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
      }
    } catch (e) {}
    return "cyrb53:" + cyrb53("qbr-admin::" + plain, 0x9e3779b9);
  }

  QBR.adminHasPassword = function () {
    try { return !!localStorage.getItem(PW_KEY); } catch (e) { return false; }
  };

  QBR.adminSetPassword = async function (plain) {
    if (!plain) return false;
    const h = await hashPw(plain);
    try { localStorage.setItem(PW_KEY, h); } catch (e) { return false; }
    return true;
  };

  QBR.adminVerifyPassword = async function (plain) {
    let stored = null;
    try { stored = localStorage.getItem(PW_KEY); } catch (e) {}
    if (!stored) return false;
    const h = await hashPw(plain || "");
    return h === stored;
  };

  QBR.adminClearPassword = function () {
    try { localStorage.removeItem(PW_KEY); } catch (e) {}
  };

  /* ---------- appearance (theme, colours, text sizes) ---------------------- */
  const THEME_KEY = "qbr-theme-custom";
  const THEME_DEFAULTS = { pageBg: "", cardBg: "", accent: "", headerText: "", textScale: 100, headScale: 100, chartScale: 100 };
  const COLOR_VARS = { pageBg: ["--bg"], cardBg: ["--surface", "--card"], accent: ["--brand", "--brand-fg"] };
  const clampScale = (v, max) => { v = parseInt(v, 10); return isFinite(v) ? Math.min(max || 150, Math.max(90, v)) : 100; };
  const isHex = v => /^#[0-9a-f]{6}$/i.test(String(v || ""));

  QBR.adminThemeLoad = function () {
    let t = {};
    try { t = JSON.parse(localStorage.getItem(THEME_KEY) || "{}") || {}; } catch (e) {}
    const o = Object.assign({}, THEME_DEFAULTS, t);
    Object.keys(COLOR_VARS).concat("headerText").forEach(k => { if (!isHex(o[k])) o[k] = ""; });
    o.textScale = clampScale(o.textScale); o.headScale = clampScale(o.headScale); o.chartScale = clampScale(o.chartScale, 200);
    return o;
  };
  QBR.adminThemeSave = function (patch) {
    const t = Object.assign(QBR.adminThemeLoad(), patch || {});
    try { localStorage.setItem(THEME_KEY, JSON.stringify(t)); } catch (e) {}
    QBR.adminThemeApply();
    return t;
  };
  QBR.adminThemeReset = function () {
    try { localStorage.removeItem(THEME_KEY); } catch (e) {}
    QBR.adminThemeApply();
  };
  function themeClear() {
    const b = document.body, r = document.documentElement;
    if (!b) return;
    Object.keys(COLOR_VARS).forEach(k => COLOR_VARS[k].forEach(v => b.style.removeProperty(v)));
    ["--qbr-head-ink", "--fs-t", "--fs-h", "--fs-c"].forEach(v => b.style.removeProperty(v));
    r.removeAttribute("data-qbr-fs"); r.removeAttribute("data-qbr-headink");
  }
  // Overrides go on <body>: the dark theme redefines its tokens on body, so :root would lose there.
  QBR.adminThemeApply = function () {
    const b = document.body, r = document.documentElement;
    if (!b) return;
    themeClear();
    if (QBR._adminPrinting) return;
    const t = QBR.adminThemeLoad();
    Object.keys(COLOR_VARS).forEach(k => { if (t[k]) COLOR_VARS[k].forEach(v => b.style.setProperty(v, t[k])); });
    if (t.headerText) { b.style.setProperty("--qbr-head-ink", t.headerText); r.setAttribute("data-qbr-headink", ""); }
    if (t.textScale !== 100 || t.headScale !== 100) {
      b.style.setProperty("--fs-t", String(t.textScale / 100));
      b.style.setProperty("--fs-h", String(t.headScale / 100));
      r.setAttribute("data-qbr-fs", "");
    }
    // Chart text (canvas, so CSS can't reach it): chart-generator.js reads --fs-c on screen only.
    if (t.chartScale !== 100) b.style.setProperty("--fs-c", String(t.chartScale / 100));
    chartRedraw();
  };
  // Charts bake their font sizes in when drawn → redraw once the slider settles (not on every tick).
  let chartScaleDrawn = 1, chartTimer = null;
  function chartRedraw() {
    const k = (typeof QBR.chartFontScale === "function") ? QBR.chartFontScale() : 1;
    if (k === chartScaleDrawn) return;
    clearTimeout(chartTimer);
    chartTimer = setTimeout(() => {
      chartScaleDrawn = k;
      try { if (typeof renderAll === "function" && typeof APP !== "undefined" && APP.model) renderAll(); } catch (e) {}
    }, 250);
  }
  // Exports (Export Tab / Deck / print) always use the standard look.
  window.addEventListener("beforeprint", () => { QBR._adminPrinting = true; themeClear(); });
  window.addEventListener("afterprint", () => { QBR._adminPrinting = false; QBR.adminThemeApply(); });

  // WCAG contrast of two CSS colours (rgb()/hex) → ratio, or null.
  function rgbOf(c) {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c || "");
    if (m) return [+m[1], +m[2], +m[3]];
    const h = /^#([0-9a-f]{6})$/i.exec(c || "");
    return h ? [0, 2, 4].map(i => parseInt(h[1].substr(i, 2), 16)) : null;
  }
  function lum(rgb) { return rgb.map(v => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); }).reduce((s, v, i) => s + v * [0.2126, 0.7152, 0.0722][i], 0); }
  QBR.adminContrast = function (a, b) {
    const x = rgbOf(a), y = rgbOf(b); if (!x || !y) return null;
    const l1 = lum(x), l2 = lum(y); return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
  };
  function contrastWarnings() {
    const cs = getComputedStyle(document.body), out = [];
    const bg = cs.getPropertyValue("--bg").trim(), card = cs.getPropertyValue("--card").trim(), ink = cs.getPropertyValue("--ink").trim();
    const head = cs.getPropertyValue("--qbr-head-ink").trim() || ink;
    const chk = (fg, bgc, what) => { const c = QBR.adminContrast(fg, bgc); if (c != null && c < 4.5) out.push(`${what}: contrast ${c.toFixed(1)}:1 (needs 4.5:1)`); };
    chk(ink, bg, "Text on page background"); chk(ink, card, "Text on cards"); chk(head, bg, "Headings on page background"); chk(head, card, "Headings on cards");
    return out;
  }

  /* ---------- masked password modal (replaces window.prompt) --------------- */
  function pwModal(title, confirmMode) {
    return new Promise(resolve => {
      // Overlay
      const ov = document.createElement("div");
      ov.className = "admin-pw-overlay";
      ov.innerHTML =
        `<div class="admin-pw-box" role="dialog" aria-modal="true" aria-label="${esc(title)}">` +
        `<div class="admin-pw-title">${esc(title)}</div>` +
        `<input type="password" class="admin-pw-input" id="admin-pw-1" autocomplete="new-password" placeholder="Password">` +
        (confirmMode ? `<input type="password" class="admin-pw-input" id="admin-pw-2" autocomplete="new-password" placeholder="Confirm password" style="margin-top:8px">` : "") +
        `<div class="admin-pw-err" id="admin-pw-err"></div>` +
        `<div class="admin-pw-btns"><button type="button" class="btn btn-sm btn-outline-secondary" id="admin-pw-cancel">Cancel</button> ` +
        `<button type="button" class="btn btn-sm btn-primary" id="admin-pw-ok">OK</button></div></div>`;
      document.body.appendChild(ov);
      const inp1 = ov.querySelector("#admin-pw-1");
      const inp2 = confirmMode ? ov.querySelector("#admin-pw-2") : null;
      const err = ov.querySelector("#admin-pw-err");
      const done = val => { try { ov.remove(); } catch (e) {} resolve(val); };
      ov.querySelector("#admin-pw-cancel").addEventListener("click", () => done(null));
      ov.addEventListener("keydown", ev => {
        if (ev.key === "Escape") done(null);
        if (ev.key === "Enter") submit();
      });
      // Don't let the admin shortcut re-trigger while the modal is open.
      ov.addEventListener("keydown", ev => ev.stopPropagation());
      function submit() {
        const v1 = inp1.value;
        if (!v1) { err.textContent = "Password cannot be empty."; return; }
        if (confirmMode) {
          if (inp2.value !== v1) { err.textContent = "Passwords did not match."; return; }
        }
        done(v1);
      }
      ov.querySelector("#admin-pw-ok").addEventListener("click", submit);
      setTimeout(() => { try { inp1.focus(); } catch (e) {} }, 30);
    });
  }

  // Module registry: key -> { label, desc, groupId }
  const MODULES = {
    security: { label: "Security & identity", desc: "Risky sign-ins, Security defaults, GDAP", groupId: "sbg-sec" },
    health:   { label: "Tenant health",       desc: "Tenant status, User management",           groupId: "sbg-health" },
    adoption: { label: "Adoption & capacity", desc: "M365 usage, Canva Education, Storage",     groupId: "sbg-adopt" },
    mail:     { label: "Domains & email",     desc: "Domain registration, Google Postmaster",  groupId: "sbg-mail" },
    reporting:{ label: "Reporting",           desc: "Executive report",                         groupId: "sbg-rep" },
    inventory:{ label: "Inventory",           desc: "Inventory, Scan",                          groupId: "sbg-inv" },
    audit:    { label: "Audit",               desc: "Risky sign-ins editor",                    groupId: "sbg-audit" },
    data:     { label: "Data",                desc: "Data quality audit, Data explorer",       groupId: "sbg-data" },
  };

  function loadFlags() {
    let flags = {};
    try { flags = JSON.parse(localStorage.getItem(STORE_KEY) || "{}") || {}; }
    catch (e) { flags = {}; }
    // Default: all modules on.
    Object.keys(MODULES).forEach(k => { if (typeof flags[k] !== "boolean") flags[k] = true; });
    return flags;
  }

  function saveFlags(flags) {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(flags)); } catch (e) {}
  }

  QBR.adminFlags = loadFlags;

  QBR.adminSetFlag = function (key, on) {
    if (!MODULES[key]) return false;
    const flags = loadFlags();
    flags[key] = !!on;
    saveFlags(flags);
    QBR.adminApplyFlags();
    return true;
  };

  QBR.adminResetFlags = function () {
    try { localStorage.removeItem(STORE_KEY); } catch (e) {}
    QBR.adminApplyFlags();
    renderAdminPanel();
  };

  // Hide sidebar groups for disabled modules. The Admin group itself is
  // only visible when body.show-admin is set (via the admin shortcut).
  QBR.adminApplyFlags = function () {
    const flags = loadFlags();
    Object.keys(MODULES).forEach(key => {
      const m = MODULES[key];
      const head = document.getElementById(m.groupId);
      if (!head) return;
      const group = head.closest(".sb-group");
      if (!group) return;
      // v1.36.0: remote flags always win over this browser's toggle (QBR.adminIsOn = remote > local > on).
      let on = !!flags[key];
      try { if (typeof QBR.adminIsOn === "function") on = !!QBR.adminIsOn(key); } catch (e) {}
      group.style.display = on ? "" : "none";
    });
    // If the active tab's module was just disabled, fall back to the first visible tab.
    try {
      const active = document.querySelector(".tab-btn.active");
      if (active && active.closest(".sb-group") && active.closest(".sb-group").style.display === "none") {
        const first = document.querySelector('.sb-group:not([style*="none"]) .tab-btn');
        if (first) first.click();
      }
    } catch (e) {}
    // v1.36.0: also enforce remote sub-feature flags (data-feature nav markers).
    try { if (QBR.remoteFlagsApplyNav) QBR.remoteFlagsApplyNav(); } catch (e) {}
  };

  QBR.adminIsModuleOn = function (key) {
    try { if (typeof QBR.adminIsOn === "function") return !!QBR.adminIsOn(key); } catch (e) {}
    return !!loadFlags()[key];
  };
  // v1.36.0: true when the hosted flags file decides this key (the local toggle then has no effect).
  function remoteOwns(key) {
    try {
      const st = QBR.remoteFlagsStatus ? QBR.remoteFlagsStatus() : null;
      return !!(st && st.remote && st.remote.flags && Object.prototype.hasOwnProperty.call(st.remote.flags, key));
    } catch (e) { return false; }
  }

  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function renderAdminPanel() {
    const host = document.getElementById("admin-body");
    if (!host) return;
    const flags = loadFlags();
    // 2026-10-09: keep the PowerShell-scripts section alive across re-renders (flag/theme changes
    // re-render the panel; rebuilding it would drop a half-typed script in the Add/Edit form).
    const keepScripts = host.querySelector(".admin-scripts");
    if (keepScripts && keepScripts.parentNode) keepScripts.parentNode.removeChild(keepScripts);
    const rows = Object.keys(MODULES).map(key => {
      const m = MODULES[key];
      const rem = remoteOwns(key);
      const on = rem ? QBR.adminIsOn(key) : !!flags[key];
      return `<div class="admin-row">` +
        `<span class="admin-row-txt"><span class="admin-row-label">${esc(m.label)}</span>` +
        (rem ? ` <span class="badge bg-secondary" title="Set by the hosted flags file — this browser's toggle has no effect">REMOTE</span>` : "") +
        `<br><span class="admin-row-desc">${esc(m.desc)}</span></span>` +
        `<button type="button" class="btn btn-sm ${on ? "btn-success" : "btn-outline-secondary"}" ` +
        `data-admin-toggle="${esc(key)}" aria-pressed="${on}"` +
        (rem ? ` disabled title="Controlled remotely (Remote flags below)"` : "") +
        `>${on ? "ON" : "OFF"}</button></div>`;
    }).join("");
    host.innerHTML =
      `<div class="admin-head"><div><h4 class="mb-1">Admin &middot; Feature flags</h4>` +
      `<div class="small text-muted">Toggle modules on or off. Changes apply instantly. Hidden modules keep all their data.</div></div>` +
      `<div class="audit-actions"><button type="button" class="btn btn-sm btn-outline-secondary" id="admin-reset">Reset to defaults</button> ` +
      `<button type="button" class="btn btn-sm btn-outline-secondary" id="admin-pw">Change password</button></div></div>` +
      `<div class="admin-list">${rows}</div>` +
      `<div class="small text-muted mt-2">Stored in this browser only. Double-click the version label again to hide this panel.</div>`;

    host.querySelectorAll("[data-admin-toggle]").forEach(b =>
      b.addEventListener("click", () => {
        const key = b.getAttribute("data-admin-toggle");
        const cur = !!loadFlags()[key];
        QBR.adminSetFlag(key, !cur);
        renderAdminPanel();
      }));
    const rs = document.getElementById("admin-reset");
    if (rs) rs.addEventListener("click", () => QBR.adminResetFlags());
    const pw = document.getElementById("admin-pw");
    if (pw) pw.addEventListener("click", async () => {
      const p1 = await pwModal("New admin password", true);
      if (p1 == null) return;
      await QBR.adminSetPassword(p1);
    });
    renderRemoteSection(host); // v1.36.0
    renderThemeSection(host);
    if (keepScripts) host.appendChild(keepScripts); else renderScriptsSection(host);
  }

  // 2026-10-09 patch (Pedro, 2026-10-09; Mars's feature-admin-scripts): the script inventory lives in the
  // Admin panel. The wizard keeps its contextual "Copy … Extraction Script" / "Copy Tenant ID".
  function renderScriptsSection(host) {
    const sec = document.createElement("div");
    sec.className = "admin-scripts";
    sec.innerHTML =
      `<h4 class="mb-1 mt-4">Admin &middot; PowerShell scripts</h4>` +
      `<div class="small text-muted mb-2">Script inventory for report extraction. Built-in scripts ship with the app; ` +
      `custom scripts and the GDAP mapping are stored in this browser only. The dashboard never runs PowerShell.</div>` +
      `<div id="admin-scripts-host"></div>`;
    host.appendChild(sec);
    const box = sec.querySelector("#admin-scripts-host");
    if (box && typeof QBR.scriptsRenderInline === "function") QBR.scriptsRenderInline(box);
    else if (box) box.innerHTML = `<div class="small text-muted">Script library not loaded.</div>`;
  }

  function renderThemeSection(host) {
    const t = QBR.adminThemeLoad();
    let mode = "light";
    try { mode = document.body.getAttribute("data-theme") || "light"; } catch (e) {}
    const colorRow = (key, label, desc) => {
      let cur = t[key];
      if (!cur) { try { const v = getComputedStyle(document.body).getPropertyValue(key === "headerText" ? "--ink" : COLOR_VARS[key][0]).trim(); const rgb = rgbOf(v); if (rgb) cur = "#" + rgb.map(x => x.toString(16).padStart(2, "0")).join(""); } catch (e) {} }
      return `<div class="admin-row"><span class="admin-row-txt"><span class="admin-row-label">${esc(label)}</span><br><span class="admin-row-desc">${t[key] ? "Custom " + esc(t[key]) : esc(desc)}</span></span>` +
        `<input type="color" class="admin-color" data-theme-color="${esc(key)}" value="${esc(cur || "#ffffff")}" aria-label="${esc(label)}">` +
        (t[key] ? `<button type="button" class="btn btn-sm btn-outline-secondary" data-theme-clear="${esc(key)}">Default</button>` : "") + `</div>`;
    };
    const sizeRow = (key, label, max) =>
      `<div class="admin-row"><span class="admin-row-txt"><span class="admin-row-label">${esc(label)}</span><br>` +
      `<span class="admin-row-desc" data-size-out="${esc(key)}">${t[key]}%${t[key] === 100 ? " (default)" : ""}</span></span>` +
      `<input type="range" class="admin-range" data-theme-size="${esc(key)}" min="90" max="${max || 150}" step="5" value="${t[key]}" aria-label="${esc(label)}"></div>`;
    const warns = contrastWarnings();
    const sec = document.createElement("div");
    sec.className = "admin-theme";
    sec.innerHTML =
      `<h4 class="mb-1 mt-4">Admin &middot; Appearance</h4>` +
      `<div class="small text-muted mb-2">Applies to this browser only. Exports and printing always use the standard look.</div>` +
      `<div class="admin-list">` +
      `<div class="admin-row"><span class="admin-row-txt"><span class="admin-row-label">Theme</span></span>` +
      `<button type="button" class="btn btn-sm ${mode !== "dark" ? "btn-primary" : "btn-outline-secondary"}" data-theme-mode="light" aria-pressed="${mode !== "dark"}">Light</button> ` +
      `<button type="button" class="btn btn-sm ${mode === "dark" ? "btn-primary" : "btn-outline-secondary"}" data-theme-mode="dark" aria-pressed="${mode === "dark"}">Dark</button></div>` +
      colorRow("pageBg", "Page background", "Theme default") +
      colorRow("cardBg", "Cards / tiles", "Theme default") +
      colorRow("accent", "Accent colour", "Brand blue") +
      colorRow("headerText", "Heading text", "Same as body text") +
      sizeRow("textScale", "Paragraph & body text size (paragraphs, tables, menus, labels)") +
      sizeRow("headScale", "Heading & tile number size") +
      sizeRow("chartScale", "Chart & graph text size (axes, legends, value labels)", 200) +
      `</div>` +
      (warns.length ? `<div class="merge-warn mt-2" role="status"><b>Hard to read:</b><ul>${warns.map(w => `<li>${esc(w)}</li>`).join("")}</ul></div>` : "") +
      `<div class="mt-2"><button type="button" class="btn btn-sm btn-outline-secondary" id="admin-theme-reset">Reset appearance</button></div>`;
    host.appendChild(sec);

    sec.querySelectorAll("[data-theme-mode]").forEach(b =>
      b.addEventListener("click", () => {
        const m = b.getAttribute("data-theme-mode");
        try { if (typeof setTheme === "function") setTheme(m); else document.body.setAttribute("data-theme", m); }
        catch (e) { document.body.setAttribute("data-theme", m); }
        QBR.adminThemeApply(); renderAdminPanel();
      }));
    sec.querySelectorAll("[data-theme-color]").forEach(inp => {
      inp.addEventListener("input", () => QBR.adminThemeSave({ [inp.getAttribute("data-theme-color")]: inp.value }));
      inp.addEventListener("change", () => renderAdminPanel());
    });
    sec.querySelectorAll("[data-theme-clear]").forEach(b =>
      b.addEventListener("click", () => { QBR.adminThemeSave({ [b.getAttribute("data-theme-clear")]: "" }); renderAdminPanel(); }));
    sec.querySelectorAll("[data-theme-size]").forEach(inp => {
      const out = sec.querySelector(`[data-size-out="${inp.getAttribute("data-theme-size")}"]`);
      inp.addEventListener("input", () => { if (out) out.textContent = inp.value + "%"; QBR.adminThemeSave({ [inp.getAttribute("data-theme-size")]: parseInt(inp.value, 10) }); });
      inp.addEventListener("change", () => renderAdminPanel());
    });
    const tr = sec.querySelector("#admin-theme-reset");
    if (tr) tr.addEventListener("click", () => { QBR.adminThemeReset(); renderAdminPanel(); });
  }

  /* ---------- Remote flags section (v1.36.0) --------------------------------
   * UI control for the hosted app's feature flags. Toggles edit a working
   * copy; Publish commits config/features.json to the repo via the GitHub
   * Contents API. The personal access token lives ONLY in this browser's
   * localStorage ("qbr-gh-pat") — Pedro's machine. No token here → the
   * toggles are read-only and Publish is unavailable.
   * ------------------------------------------------------------------------ */
  QBR._rfUI = QBR._rfUI || {};
  try {
    document.addEventListener("qbr:flags-ready", function () {
      if (!QBR._rfUI.dirty) QBR._rfUI.working = null;
      try { if (document.body.classList.contains("show-admin")) renderAdminPanel(); } catch (e) {}
    });
  } catch (e) {}

  function rfWorkingFlags() {
    var ui = QBR._rfUI;
    if (!ui.working) {
      var base = {};
      try {
        var st = QBR.remoteFlagsStatus ? QBR.remoteFlagsStatus() : null;
        if (st && st.remote && st.remote.flags) base = st.remote.flags;
      } catch (e) {}
      // Publishing to another branch (e.g. main from a preview): start from THAT branch's current values.
      var tgt = rfTarget();
      if (tgt && tgt.loaded) base = tgt.doc ? tgt.doc.flags : {};
      var w = {};
      var feats = (QBR.features ? QBR.features.list() : []);
      feats.forEach(function (f) {
        w[f.key] = Object.prototype.hasOwnProperty.call(base, f.key) ? !!base[f.key] : true;
      });
      ui.working = w;
    }
    return ui.working;
  }

  // Target branch differs from the site this tab runs on → its current flags are fetched (needs the token).
  function rfNeedsTarget() {
    var b = (rfRepoCfg().branch || "main");
    var host = QBR.remoteFlagsHostBranch ? QBR.remoteFlagsHostBranch() : b;
    var can = false;
    try { can = QBR.remoteFlagsToken && QBR.remoteFlagsToken.present(); } catch (e) {}
    return can && !!QBR.remoteFlagsReadTarget && b !== host ? b : null;
  }
  function rfTarget() {
    var b = rfNeedsTarget(), t = QBR._rfUI.target;
    return b && t && t.branch === b ? t : null;
  }
  function rfLoadTarget() {
    var b = rfNeedsTarget(), ui = QBR._rfUI;
    if (!b || (ui.target && ui.target.branch === b)) return;
    ui.target = { branch: b, loaded: false, loading: true, doc: null, err: "" };
    QBR.remoteFlagsReadTarget(b).then(function (r) {
      if (!ui.target || ui.target.branch !== b) return;
      ui.target = { branch: b, loaded: true, loading: false, doc: r.doc, err: "" };
      if (!ui.dirty) ui.working = null;
      renderAdminPanel();
    }, function (err) {
      if (!ui.target || ui.target.branch !== b) return;
      ui.target = { branch: b, loaded: false, loading: false, doc: null, err: String((err && err.message) || err) };
      renderAdminPanel();
    });
  }
  function rfBaseFlags() {
    var t = rfTarget();
    if (t && t.loaded) return t.doc ? t.doc.flags : {};
    try { var st = QBR.remoteFlagsStatus(); if (st && st.remote && st.remote.flags) return st.remote.flags; } catch (e) {}
    return {};
  }

  function rfRepoCfg() {
    var ui = QBR._rfUI;
    if (!ui.repo) {
      try { ui.repo = QBR.remoteFlagsRepo ? QBR.remoteFlagsRepo.get() : { owner: "", repo: "", branch: "main" }; }
      catch (e) { ui.repo = { owner: "", repo: "", branch: "main" }; }
    }
    return ui.repo;
  }

  function renderRemoteSection(host) {
    var sec = document.createElement("div");
    sec.className = "admin-remote";
    var hasApi = !!(QBR.remoteFlagsStatus && QBR.remoteFlagsPublish && QBR.features);
    var st = { ready: false, source: "none", remote: null };
    try { if (QBR.remoteFlagsStatus) st = QBR.remoteFlagsStatus(); } catch (e) {}
    var T = QBR.remoteFlagsToken || { present: function () { return false; } };
    var canPublish = false, locked = false, remembered = false, canRemember = false, purged = false;
    try {
      canPublish = hasApi && T.present(); locked = hasApi && T.locked(); remembered = hasApi && T.remembered();
      canRemember = hasApi && T.canRemember(); purged = hasApi && T.purgedPlain() && !canPublish && !remembered;
    } catch (e) {}
    var working = rfWorkingFlags();
    var repo = rfRepoCfg();
    var ui = QBR._rfUI;
    var base = repo.branch || "main";
    var prot = !!(QBR.remoteFlagsIsProtected && QBR.remoteFlagsIsProtected(base));
    rfLoadTarget();
    var tgtNeed = rfNeedsTarget(), tgt = ui.target && ui.target.branch === tgtNeed ? ui.target : null;
    var tgtNote = "";
    if (tgtNeed && tgt) {
      if (tgt.loading) tgtNote = "Loading the current settings on " + tgtNeed + "…";
      else if (tgt.err) tgtNote = "Couldn't read the settings on " + tgtNeed + " (" + tgt.err + ") — the switches show this site's values.";
      else if (tgt.doc) tgtNote = "The switches show the CURRENT settings on " + tgtNeed + " (v" + tgt.doc.version + ", updated " + (tgt.doc.updated || "?") + "), not this site's.";
      else tgtNote = tgtNeed + " has no flags file yet — everything is on there.";
    }

    var statusTxt;
    if (!hasApi) statusTxt = "Remote flags module not loaded.";
    else if (st.source === "live" && st.remote) statusTxt = "Remote flags v" + st.remote.version + " · updated " + (st.remote.updated || "?") + " · loaded from the live site.";
    else if (st.source === "cache") statusTxt = "Showing the last-known remote flags (offline — cached copy).";
    else statusTxt = "Remote flags not loaded yet — showing defaults (all on).";

    var feats = QBR.features ? QBR.features.list() : [];
    var groups = {};
    feats.forEach(function (f) {
      var g = QBR.features.moduleLabel(f.module);
      (groups[g] = groups[g] || []).push(f);
    });
    var rows = Object.keys(groups).sort().map(function (g) {
      return '<div class="admin-row-group"><div class="small fw-bold mt-2 mb-1">' + esc(g) + "</div>" +
        groups[g].map(function (f) {
          var on = !!working[f.key];
          return '<div class="admin-row">' +
            '<span class="admin-row-txt"><span class="admin-row-label">' + esc(f.label) + "</span> " +
            '<span class="badge bg-secondary" title="Managed by the remote flags file">REMOTE</span><br>' +
            '<span class="admin-row-desc">' + esc(f.desc) + "</span></span>" +
            '<button type="button" class="btn btn-sm ' + (on ? "btn-success" : "btn-outline-secondary") + '" ' +
            'data-rf-toggle="' + esc(f.key) + '" aria-pressed="' + on + '"' +
            (canPublish ? "" : ' disabled title="Add or unlock your GitHub token to change remote flags"') +
            ">" + (on ? "ON" : "OFF") + "</button></div>";
        }).join("") + "</div>";
    }).join("");

    sec.innerHTML =
      '<h4 class="mb-1 mt-4">Admin &middot; Remote flags</h4>' +
      '<div class="small text-muted mb-2">Turn features on/off on the <b>hosted</b> app for every machine. ' +
      "Publishing updates <code>config/features.json</code> in the repo (through a pull request for main and develop); the site rebuilds in a minute or two and machines pick it up on next load.</div>" +
      '<div class="small mb-2">' + esc(statusTxt) + "</div>" +
      (tgtNote ? '<div class="small mb-2 fw-semibold" id="rf-target-note">' + esc(tgtNote) + "</div>" : "") +
      (feats.length ? '<div class="admin-list">' + rows + "</div>" : '<div class="small text-muted">No remotely-managed features registered.</div>') +
      '<div class="card-box mt-3"><h6 class="mb-2">Publish to GitHub</h6>' +
      '<div class="small mb-2">This tab can publish: <b>' + (canPublish ? "Yes" : "No") + "</b>" +
      (canPublish ? ' <span class="text-muted">(locks after ' + T.IDLE_MIN + " min without use)</span>" : "") +
      (locked ? ' · <span class="text-muted">a saved token is <b>locked</b> on this machine</span>' : "") + "</div>" +
      (purged ? '<div class="small text-danger mb-2">A token saved in plain text by an earlier build was deleted from this browser. Add it again below.</div>' : "") +
      '<div class="d-flex flex-wrap gap-2 align-items-end mb-2">' +
      '<label class="small mb-0">Owner<br><input class="form-control form-control-sm" style="max-width:180px" data-rf-repo="owner" value="' + esc(repo.owner || "") + '"></label>' +
      '<label class="small mb-0">Repo<br><input class="form-control form-control-sm" style="max-width:200px" data-rf-repo="repo" value="' + esc(repo.repo || "") + '"></label>' +
      '<label class="small mb-0">Branch<br><input class="form-control form-control-sm" style="max-width:110px" data-rf-repo="branch" value="' + esc(repo.branch || "main") + '"></label>' +
      "</div>" +
      (locked
        ? '<div class="d-flex flex-wrap gap-2 align-items-center mb-2"><button type="button" class="btn btn-sm btn-primary" id="rf-unlock">Unlock saved token</button>' +
          '<button type="button" class="btn btn-sm btn-outline-secondary" id="rf-pat-forget">Forget token</button></div>'
        : '<div class="d-flex flex-wrap gap-2 align-items-end mb-2">' +
          '<label class="small mb-0">GitHub fine-grained token<br><input type="password" id="rf-pat" class="form-control form-control-sm" style="max-width:280px" autocomplete="off" spellcheck="false" placeholder="github_pat_… (your own machine only)"></label>' +
          '<label class="small mb-0"' + (canRemember ? "" : ' title="Needs the https site — not available when opened as a file"') + '><input type="checkbox" id="rf-pat-remember"' + (canRemember ? "" : " disabled") + "> Remember on this machine (encrypted with a passphrase)</label>" +
          '<button type="button" class="btn btn-sm btn-outline-primary" id="rf-pat-save">Save token</button>' +
          (canPublish || remembered ? ' <button type="button" class="btn btn-sm btn-outline-secondary" id="rf-pat-forget">Forget token</button>' : "") +
          (canPublish ? ' <button type="button" class="btn btn-sm btn-outline-secondary" id="rf-lock">Lock now</button>' : "") +
          "</div>") +
      '<div class="small text-muted mb-2">Enter the token <b>only on your own machine</b>. Use a <b>fine-grained</b> token for this repository only — ' +
      "permissions <b>Contents</b> and <b>Pull requests</b>: read &amp; write, nothing else — with an expiry of 90 days or less. " +
      "Classic tokens (ghp_…) are refused. It is only ever sent to api.github.com and never written to the repo.</div>" +
      (prot
        ? '<div class="small mb-2" id="rf-target">Target <b>' + esc(base) + "</b> — Publish opens a <b>pull request</b>. Nothing changes until someone reviews and merges it on GitHub.</div>"
        : '<div class="small text-muted mb-2" id="rf-target">Target <b>' + esc(base) + "</b> (preview branch) — Publish commits directly; only that branch's preview site is affected.</div>") +
      '<div class="d-flex flex-wrap gap-2 align-items-center">' +
      '<button type="button" class="btn btn-sm btn-primary" id="rf-publish"' + (canPublish ? "" : " disabled") + ">" + (prot ? "Open pull request" : "Publish flags") + "</button>" +
      '<span class="small" id="rf-msg" role="status">' + (ui.msg ? esc(ui.msg) : "") +
      (ui.prUrl ? ' <a href="' + esc(ui.prUrl) + '" target="_blank" rel="noopener noreferrer">Open the pull request</a>' : "") + "</span>" +
      "</div></div>";

    host.appendChild(sec);

    // Toggles edit the working copy only — nothing leaves the browser until Publish.
    sec.querySelectorAll("[data-rf-toggle]").forEach(function (b) {
      b.addEventListener("click", function () {
        var key = b.getAttribute("data-rf-toggle");
        working[key] = !working[key];
        ui.dirty = true;
        renderAdminPanel();
      });
    });
    // Repo fields survive re-renders via QBR._rfUI.repo; a branch change re-renders the target note.
    sec.querySelectorAll("[data-rf-repo]").forEach(function (inp) {
      inp.addEventListener("input", function () {
        repo[inp.getAttribute("data-rf-repo")] = inp.value.trim();
        try { if (QBR.remoteFlagsRepo) QBR.remoteFlagsRepo.set(repo); } catch (e) {}
      });
      inp.addEventListener("change", function () { ui.target = null; if (!ui.dirty) ui.working = null; renderAdminPanel(); });
    });
    var say = function (m) { ui.msg = m; ui.prUrl = ""; renderAdminPanel(); };
    var patInp = sec.querySelector("#rf-pat");
    var saveBtn = sec.querySelector("#rf-pat-save");
    if (saveBtn) saveBtn.addEventListener("click", async function () {
      var v = patInp ? patInp.value.trim() : "";
      var bad = T.formatError(v);
      if (bad) { if (patInp) patInp.value = ""; say(bad); return; }
      var remember = !!(sec.querySelector("#rf-pat-remember") || {}).checked;
      var pass = null;
      if (remember) {
        pass = await pwModal("Passphrase to encrypt the token (8+ characters, not your GitHub password)", true);
        if (pass == null) return;
      }
      if (patInp) patInp.value = "";
      try {
        await T.set(v, { remember: remember, passphrase: pass });
      } catch (err) { say(err.message || String(err)); return; }
      var note = remember ? "Token saved, encrypted on this machine." : "Token kept for this tab only.";
      try {
        var chk = await QBR.remoteFlagsCheckToken();
        if (chk.daysLeft == null) note += " Couldn't read its expiry — make sure it expires within 90 days.";
        else if (chk.daysLeft > 90) note += " Warning: it expires in " + chk.daysLeft + " days — use 90 days or less.";
        else note += " It can reach the repository and expires in " + chk.daysLeft + " day" + (chk.daysLeft === 1 ? "" : "s") + ".";
      } catch (err) {
        var m = String((err && err.message) || err);
        if (/^Token check failed/.test(m)) { T.clear(); say("Token not accepted and removed: " + m); return; }
        note += " Couldn't reach GitHub to check it (" + m + ") — it will be checked when you publish.";
      }
      say(note);
    });
    var unlockBtn = sec.querySelector("#rf-unlock");
    if (unlockBtn) unlockBtn.addEventListener("click", async function () {
      var pass = await pwModal("Passphrase for the saved GitHub token", false);
      if (pass == null) return;
      try { await T.unlock(pass); say("Token unlocked for this tab."); }
      catch (err) { say(err.message || String(err)); }
    });
    var lockBtn = sec.querySelector("#rf-lock");
    if (lockBtn) lockBtn.addEventListener("click", function () { T.lock(); say(remembered ? "Locked — the saved token needs its passphrase again." : "Token removed from this tab."); });
    var forgetBtn = sec.querySelector("#rf-pat-forget");
    if (forgetBtn) forgetBtn.addEventListener("click", function () {
      try { T.clear(); } catch (e) {}
      say("Token forgotten on this machine. You can also revoke it in GitHub › Settings › Developer settings.");
    });
    var pubBtn = sec.querySelector("#rf-publish");
    if (pubBtn) pubBtn.addEventListener("click", async function () {
      var msg = sec.querySelector("#rf-msg");
      // Show exactly what changes on the target before anything is sent.
      var changes = QBR.remoteFlagsDiff ? QBR.remoteFlagsDiff(rfBaseFlags(), Object.assign({}, rfBaseFlags(), working)) : [];
      var q = (prot ? "Open a pull request into " : "Publish directly to ") + base + "?\n\n" +
        (changes.length ? "Changes:\n" + changes.join("\n") : "No switch changes (the version number only).");
      if (!window.confirm(q)) return;
      pubBtn.disabled = true;
      if (msg) msg.textContent = prot ? "Opening a pull request…" : "Publishing…";
      try {
        var res = await QBR.remoteFlagsPublish(Object.assign({}, working));
        ui.dirty = false;
        ui.target = null;
        if (res.pr) {
          ui.working = null;   // the live flags haven't changed yet — show them again
          ui.msg = "Pull request #" + res.pr.number + " opened for v" + res.doc.version + ". Merge it on GitHub to apply; the site rebuilds after the merge.";
          ui.prUrl = res.pr.url || "";
        } else {
          ui.working = Object.assign({}, res.doc.flags);
          ui.msg = "Published to " + base + " — that site rebuilds in ~2 min, then machines pick it up on next load.";
          ui.prUrl = "";
        }
      } catch (err) {
        ui.msg = "Publish failed: " + (err && err.message ? err.message : err);
        ui.prUrl = "";
      }
      renderAdminPanel();
    });
  }

  QBR.renderAdmin = renderAdminPanel;

  // Admin trigger (v1.32.0, user's choice): DOUBLE-CLICK the version label only.
  // Two copies of the label exist: the sidebar footer (".sb-ver", hidden while the sidebar
  // is collapsed to a rail — the default below 1400 px) and the top bar ("#app-version",
  // always visible). Either one opens the password prompt. No keyboard shortcut.
  QBR.adminTriggerHit = function (target) {
    return !!(target && target.closest && target.closest(".sb-ver, #app-version"));
  };
  // Password gate, then toggle the Admin section.
  async function onTrigger(e) {
    try {
      if (QBR.adminTriggerHit(e.target)) {
        e.preventDefault();
        try { const sel = window.getSelection && window.getSelection(); if (sel) sel.removeAllRanges(); } catch (x) {}
        const showing = document.body.classList.contains("show-admin");
        if (showing) {
          document.body.classList.remove("show-admin");
          return;
        }
        // Gate: set password on first use, verify afterwards.
        if (!QBR.adminHasPassword()) {
          const p1 = await pwModal("Set admin password", true);
          if (p1 == null) return;
          await QBR.adminSetPassword(p1);
        } else {
          const p = await pwModal("Admin password", false);
          if (p == null) return;
          if (!await QBR.adminVerifyPassword(p)) return; // wrong → stay hidden
        }
        document.body.classList.add("show-admin");
        renderAdminPanel();
        const btn = document.querySelector('[data-tab="dash-admin"]');
        if (btn) btn.click();
      }
    } catch (err) {}
  }

  function init() {
    document.addEventListener("dblclick", onTrigger);
    // Apply flags + appearance on load.
    const apply = () => { QBR.adminApplyFlags(); QBR.adminThemeApply(); };
    if (document.readyState === "complete" || document.readyState === "interactive") apply();
    else document.addEventListener("DOMContentLoaded", apply);
  }

  QBR.adminInit = init;
  try { init(); } catch (e) {}
})();
