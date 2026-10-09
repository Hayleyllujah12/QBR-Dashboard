/* =========================================================================
 * QBR Dashboard — shell.js  (v1.9.0, Phase 1)
 * App-shell chrome ONLY: left-nav state, rail/drawer, breadcrumb, Export
 * menu, theme-button label, deep links. It holds no data logic and never
 * calls into the model.
 *
 * Contract with app.js (unchanged):
 *   - app.js binds a click handler to every [data-tab] element in initShell()
 *     and toggles .active + the matching .dash-panel. Sidebar items ARE those
 *     elements (button.tab-btn[data-tab]); this file only mirrors state
 *     (aria-current, breadcrumb, focus) after app.js has handled the click.
 *   - app.js shows #app-body / hides #empty-state once data is loaded; this
 *     file observes that to reveal the sidebar (body.has-data).
 *   - Export buttons keep their ids (btn-pdf/deck/images/data/clean) and
 *     their app.js listeners; they simply live inside the Export menu.
 * Loaded after app.js. Plain ES2015, no modules, offline.
 * ========================================================================= */
(function () {
  "use strict";
  var LS_RAIL = "qbr-sidebar-rail";
  var DESKTOP = window.matchMedia("(min-width: 1024px)");
  var $ = function (id) { return document.getElementById(id); };
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }

  function init() {
    var body = document.body;
    var sidebar = $("app-sidebar"), scrim = $("sb-scrim");
    var navToggle = $("nav-toggle"), collapse = $("sb-collapse");
    var appBody = $("app-body");
    if (!sidebar || !appBody) return;
    var items = Array.prototype.slice.call(sidebar.querySelectorAll(".sb-item[data-tab]"));

    /* ---- version label in the sidebar footer (single source: QBR.VERSION) ---- */
    if (window.QBR && QBR.VERSION && $("sb-version")) $("sb-version").textContent = "v" + QBR.VERSION;

    /* ---- sidebar only once data is loaded (mirrors app.js empty-state toggle) ---- */
    function syncHasData() {
      var has = !appBody.classList.contains("d-none");
      body.classList.toggle("has-data", has);
      if (!has) closeDrawer();
      if (has) applyHash();
    }
    new MutationObserver(syncHasData).observe(appBody, { attributes: true, attributeFilter: ["class"] });

    /* ---- rail (desktop) ---- */
    function setRail(on, persist) {
      body.classList.toggle("sb-rail", on);
      if (collapse) {
        collapse.setAttribute("aria-expanded", String(!on));
        collapse.setAttribute("aria-label", on ? "Expand navigation" : "Collapse navigation");
        collapse.title = on ? "Expand navigation" : "Collapse navigation";
        var lbl = collapse.querySelector(".sb-label"); if (lbl) lbl.textContent = on ? "Expand" : "Collapse";
      }
      if (persist) lsSet(LS_RAIL, on ? "1" : "0");
    }
    var saved = lsGet(LS_RAIL);
    setRail(saved === null ? window.innerWidth < 1400 : saved === "1", false);
    if (collapse) collapse.addEventListener("click", function () { setRail(!body.classList.contains("sb-rail"), true); });

    /* ---- drawer (< 1024px) ---- */
    function openDrawer() {
      body.classList.add("sb-open"); if (scrim) scrim.hidden = false;
      if (navToggle) { navToggle.setAttribute("aria-expanded", "true"); navToggle.setAttribute("aria-label", "Close navigation"); }
      var a = sidebar.querySelector(".sb-item.active") || items[0]; if (a) a.focus();
    }
    function closeDrawer(returnFocus) {
      if (!body.classList.contains("sb-open")) return;
      body.classList.remove("sb-open"); if (scrim) scrim.hidden = true;
      if (navToggle) { navToggle.setAttribute("aria-expanded", "false"); navToggle.setAttribute("aria-label", "Open navigation"); if (returnFocus) navToggle.focus(); }
    }
    if (navToggle) navToggle.addEventListener("click", function () { body.classList.contains("sb-open") ? closeDrawer(true) : openDrawer(); });
    if (scrim) scrim.addEventListener("click", function () { closeDrawer(true); });
    var onMq = function () { if (DESKTOP.matches) closeDrawer(); };
    if (DESKTOP.addEventListener) DESKTOP.addEventListener("change", onMq); else if (DESKTOP.addListener) DESKTOP.addListener(onMq);

    /* ---- rail tooltips (labels are visually hidden in rail mode) ---- */
    items.forEach(function (b) { var l = b.querySelector(".sb-label"); if (l) b.title = l.textContent; });

    /* ---- nav state: runs AFTER app.js's own [data-tab] handler ---- */
    function markActive(btn, moveFocus) {
      items.forEach(function (b) { if (b === btn) b.setAttribute("aria-current", "page"); else b.removeAttribute("aria-current"); });
      var g = btn.getAttribute("data-group") || "", l = btn.querySelector(".sb-label");
      var cg = $("crumb-group"), cp = $("crumb-page");
      if (cg) { cg.innerHTML = g; cg.hidden = !g; }
      if (cp && l) cp.textContent = l.textContent;
      if (moveFocus) {
        var panel = $(btn.getAttribute("data-tab")), h = panel && panel.querySelector(".dash-h");
        window.scrollTo(0, 0);
        if (h) { h.setAttribute("tabindex", "-1"); h.focus({ preventScroll: true }); }
      }
    }
    sidebar.addEventListener("click", function (e) {
      var b = e.target.closest(".sb-item[data-tab]"); if (!b) return;
      setTimeout(function () {
        markActive(b, true);
        setHash(b.getAttribute("data-tab"));
        closeDrawer();
      }, 0);
    });
    // goToTab() (Overview quick-jumps) calls btn.click() → same path as above.

    /* ---- deep links: #risky, #sec, … (bare tokens only) ---- */
    function setHash(tab) {
      var t = String(tab).replace(/^dash-/, "");
      try { history.replaceState(null, "", t === "overview" ? location.pathname + location.search : "#" + t); } catch (e) {}
    }
    var hashApplied = false;
    function applyHash() {
      if (hashApplied) return; hashApplied = true;
      var h = (location.hash || "").replace(/^#/, "");
      if (!/^[a-z]+$/.test(h)) return;
      var b = sidebar.querySelector('.sb-item[data-tab="dash-' + h + '"]');
      if (b && !b.classList.contains("active")) b.click();
    }

    /* ---- Export menu (disclosure; buttons inside keep their app.js listeners) ---- */
    var eBtn = $("btn-export"), eMenu = $("export-menu");
    function menuItems() { return Array.prototype.slice.call(eMenu.querySelectorAll(".menu-item")); }
    function openMenu(focusFirst) {
      eMenu.hidden = false; eBtn.setAttribute("aria-expanded", "true");
      if (focusFirst) { var m = menuItems()[0]; if (m) m.focus(); }
    }
    function closeMenu(returnFocus) {
      if (eMenu.hidden) return;
      eMenu.hidden = true; eBtn.setAttribute("aria-expanded", "false");
      if (returnFocus) eBtn.focus();
    }
    if (eBtn && eMenu) {
      eBtn.addEventListener("click", function (e) { e.stopPropagation(); eMenu.hidden ? openMenu(e.detail === 0) : closeMenu(); });
      eBtn.addEventListener("keydown", function (e) { if (e.key === "ArrowDown") { e.preventDefault(); openMenu(true); } });
      eMenu.addEventListener("keydown", function (e) {
        var list = menuItems(), i = list.indexOf(document.activeElement);
        if (e.key === "ArrowDown") { e.preventDefault(); list[(i + 1) % list.length].focus(); }
        else if (e.key === "ArrowUp") { e.preventDefault(); list[(i - 1 + list.length) % list.length].focus(); }
        else if (e.key === "Home") { e.preventDefault(); list[0].focus(); }
        else if (e.key === "End") { e.preventDefault(); list[list.length - 1].focus(); }
        else if (e.key === "Tab") { closeMenu(); }
      });
      // Close after the item's own app.js handler has run (print/export start synchronously).
      eMenu.addEventListener("click", function (e) { if (e.target.closest(".menu-item")) setTimeout(function () { closeMenu(); }, 0); });
      document.addEventListener("click", function (e) { if (!eMenu.hidden && !eMenu.contains(e.target) && e.target !== eBtn) closeMenu(); });
    }

    /* ---- Escape closes the menu first, then the drawer ---- */
    document.addEventListener("keydown", function (e) {
      if (e.key !== "Escape") return;
      if (eMenu && !eMenu.hidden) { closeMenu(true); return; }
      if (body.classList.contains("sb-open")) closeDrawer(true);
    });

    /* ---- theme button: clean accessible name (app.js writes an emoji label) ---- */
    var tb = $("btn-theme");
    function syncThemeLabel() {
      if (!tb) return;
      var dark = body.getAttribute("data-theme") === "dark";
      var t = dark ? "Switch to light mode" : "Switch to dark mode";
      tb.setAttribute("aria-label", t); tb.title = t;
    }
    new MutationObserver(syncThemeLabel).observe(body, { attributes: true, attributeFilter: ["data-theme"] });
    syncThemeLabel();

    /* ---- sortable tables (v1.9.0 Phase 3) ----
     * Every data table inside a dashboard panel gets clickable headers
     * (mouse, Enter or Space) with aria-sort. app.js rebuilds table bodies on
     * each render, so the chosen sort is remembered per table and re-applied
     * whenever the rows change. The Data explorer keeps its own sorter. */
    var sortState = {};
    function tableKey(t) {
      var tb = t.tBodies[0];
      if (tb && tb.id) return tb.id;
      if (t.id) return t.id;
      var panel = t.closest(".dash-panel");
      var all = panel ? Array.prototype.slice.call(panel.querySelectorAll("table")) : [];
      return (panel ? panel.id : "x") + "#" + all.indexOf(t);
    }
    function sortable(t) {
      // data-nosort on the <table>: opt out (e.g. grouped rows with paired detail rows — v1.33.0)
      return t.closest(".dash-panel") && !t.closest(".fd-tablewrap") && !t.hasAttribute("data-nosort") && t.tHead && t.tBodies[0] && t.tHead.rows.length;
    }
    function setupTable(t) {
      if (t.dataset.sortable || !sortable(t)) return;
      t.dataset.sortable = "1";
      var row = t.tHead.rows[t.tHead.rows.length - 1];
      Array.prototype.forEach.call(row.cells, function (th) {
        if (!th.textContent.trim() || th.textContent.trim() === "#" || th.colSpan > 1) return;   // skip blank + rank columns
        th.classList.add("th-sort"); th.tabIndex = 0; th.setAttribute("aria-sort", "none");
        if (!th.title) th.title = "Sort by " + th.textContent.trim();
      });
    }
    function cellValue(td) {
      if (!td) return null;
      var t = td.textContent.replace(/\s+/g, " ").trim();
      if (!t || t === "—" || t === "-" || /^no data$/i.test(t)) return null;
      var m = t.replace(/,/g, "").match(/^[-+]?\d*\.?\d+/);
      return m ? { n: parseFloat(m[0]) } : { s: t.toLowerCase() };
    }
    function applySort(t, col, dir) {
      var tb = t.tBodies[0], rows = Array.prototype.slice.call(tb.rows);
      if (rows.length < 2 || rows.some(function (r) { return r.cells.length === 1 && r.cells[0].colSpan > 1; })) return;
      var keyed = rows.map(function (r, i) { return { r: r, i: i, v: cellValue(r.cells[col]) }; });
      keyed.sort(function (a, b) {
        if (a.v == null && b.v == null) return a.i - b.i;
        if (a.v == null) return 1;                       // blanks / "—" always last
        if (b.v == null) return -1;
        var c;
        if (a.v.n != null && b.v.n != null) c = a.v.n - b.v.n;
        else c = String(a.v.s != null ? a.v.s : a.v.n).localeCompare(String(b.v.s != null ? b.v.s : b.v.n), undefined, { numeric: true });
        return (dir === "ascending" ? c : -c) || (a.i - b.i);
      });
      keyed.forEach(function (k) { tb.appendChild(k.r); });
    }
    function markHeaders(t, col, dir) {
      var row = t.tHead.rows[t.tHead.rows.length - 1];
      Array.prototype.forEach.call(row.cells, function (th, i) {
        if (th.classList.contains("th-sort")) th.setAttribute("aria-sort", i === col ? dir : "none");
      });
    }
    var mo = null;
    function withObserverPaused(fn) {
      if (mo) mo.disconnect();
      try { fn(); } finally { if (mo) { mo.takeRecords(); mo.observe(appBody, { childList: true, subtree: true }); } }
    }
    function sweep() {
      withObserverPaused(function () {
        Array.prototype.forEach.call(appBody.querySelectorAll(".dash-panel table"), function (t) {
          setupTable(t);
          var st = sortState[tableKey(t)];
          if (st && t.dataset.sortable) { markHeaders(t, st.col, st.dir); applySort(t, st.col, st.dir); }
        });
      });
    }
    function onSortTrigger(th) {
      var t = th.closest("table"), row = th.parentNode, col = Array.prototype.indexOf.call(row.cells, th);
      var key = tableKey(t), cur = sortState[key];
      var firstNum = Array.prototype.some.call(t.tBodies[0].rows, function (r) { var v = cellValue(r.cells[col]); return v && v.n != null; });
      var dir = cur && cur.col === col ? (cur.dir === "ascending" ? "descending" : "ascending") : (firstNum ? "descending" : "ascending");
      sortState[key] = { col: col, dir: dir };
      withObserverPaused(function () { markHeaders(t, col, dir); applySort(t, col, dir); });
    }
    appBody.addEventListener("click", function (e) {
      if (e.target.closest(".col-resizer")) return;
      var th = e.target.closest("th.th-sort"); if (th) onSortTrigger(th);
    });
    appBody.addEventListener("keydown", function (e) {
      var th = e.target.closest && e.target.closest("th.th-sort");
      if (th && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); onSortTrigger(th); }
    });
    var sweepTimer = null;
    mo = new MutationObserver(function () { clearTimeout(sweepTimer); sweepTimer = setTimeout(sweep, 30); });
    mo.observe(appBody, { childList: true, subtree: true });

    /* ---- initial state ---- */
    var act = sidebar.querySelector(".sb-item.active"); if (act) markActive(act, false);
    syncHasData();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
