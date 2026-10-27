/* RCT OpsDesk — branding.js (2026-10-05)
 * Custom logo by manual upload. Click the top-bar logo → "Upload new logo…" / "Reset to default".
 * The image is stored as a data: URL in localStorage ("rct-logo-v1") — this browser only, never sent anywhere.
 * Default logo = assets/rakso-ed-logo.png (replace that file to change the logo for everyone).
 * Shown via <img>, so an uploaded SVG cannot run scripts. No data logic; safe to remove. */
(function () {
  "use strict";
  var KEY = "rct-logo-v1", MAX = 512 * 1024, DEFAULT = "assets/rakso-ed-logo.png";
  var OK_TYPES = ["image/svg+xml", "image/png", "image/jpeg", "image/webp"];
  function $(id) { return document.getElementById(id); }
  function load() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function save(v) { try { if (v) localStorage.setItem(KEY, v); else localStorage.removeItem(KEY); return true; } catch (e) { return false; } }
  function favicon(src) {
    var l = document.querySelector('link[rel="icon"]');
    if (!l) { l = document.createElement("link"); l.rel = "icon"; document.head.appendChild(l); }
    l.type = /^data:image\/svg|\.svg$/i.test(src) ? "image/svg+xml" : (src.match(/^data:([^;]+)/) || [, "image/png"])[1];
    l.href = src;
  }
  function apply(src) {
    var img = $("brand-logo"); if (!img) return;
    var url = src || DEFAULT;
    img.onerror = function () { if (src) { img.onerror = null; save(null); apply(null); } };
    img.src = url; favicon(url);
    var r = $("brand-reset"); if (r) r.disabled = !src;
  }
  function note(msg, err) {
    var n = document.querySelector("#brand-menu .brand-menu-n"); if (!n) return;
    if (!n.dataset.base) n.dataset.base = n.textContent;
    n.textContent = msg || n.dataset.base; n.classList.toggle("err", !!err);
  }
  window.RCT_BRAND = { apply: apply, reset: function () { save(null); apply(null); }, KEY: KEY };

  // Environment badge (2026-10-06, v1.31.0). Detected from where the app runs, so the
  // same code is safe on every branch (nothing shows on the live folder / GitHub Pages):
  //   Beta version — Netlify is the beta/testing host (live = GitHub Pages), so the project's main
  //                  address "<site>.netlify.app" and "develop--<site>" are beta; also the
  //                  QBR-Dashboard_dev test folder, or ?beta (?dev kept as an alias)
  //   Preview      — any other Netlify branch / deploy-preview URL ("<branch>--<site>.netlify.app")
  //   Main         — live: GitHub Pages, the live OneDrive folder, or ?main
  //   (2026-10-19: the Azure Static Web App deploys develop → "Beta version"; its PR staging sites → "Staging · PR n")
  //   (2026-10-26: a stamped branch label wins over hostname labels; explicit ?main/?beta/?dev flags still win over the stamp.)
  function envInfo() {
    var path = "", host = "", q = "";
    try { path = decodeURIComponent(location.pathname); } catch (e) { path = location.pathname || ""; }
    try { host = String(location.hostname || "").toLowerCase(); q = location.search || ""; } catch (e) {}
    /* Precedence (2026-10-26 hosted-branch fix):
     * 1. Explicit query flags win: ?beta/?dev -> Beta, ?main -> Main.
     * 2. Branch label wins over hostname-derived labels: ?branch=<name>,
     *    else the QBR_BRANCH stamp in index.html. Previously the stamp only
     *    filled the null case, so a stamped branch build hosted on GitHub
     *    Pages showed "Main" and on the Azure SWA showed "Beta".
     * 3. Hostname-derived labels (unchanged): Netlify, Azure staging/Beta,
     *    GitHub Pages / versioned-folder path -> Main.
     * Main/develop builds carry no stamp, so their labels are unchanged.
     * WARNING: the stamp is now load-bearing — it MUST be dropped on merge
     * to main (see the index.html comment), or main would mislabel itself
     * as a branch. */
    if (/[?&](beta|dev)\b/i.test(q))
      return { kind: "beta", text: "Beta version", tag: "[BETA]", tip: "Beta (develop branch) — for testing before release. Changes here are not live." };
    if (/[?&]main\b/i.test(q))
      return { kind: "main", text: "Main", tag: "[MAIN]", tip: "Main (live) — the production build everyone uses." };
    var label = "";
    try { var bm = /[?&]branch=([^&#]*)/i.exec(q); label = bm ? decodeURIComponent(bm[1] || "").trim() : ""; } catch (e) { label = ""; }
    if (!label && typeof window.QBR_BRANCH === "string") label = window.QBR_BRANCH.trim();
    if (label) return { kind: "preview", text: "Branch · " + label, tag: "[BRANCH]", tip: "Branch build (" + label + ") — not live." };
    var m = /^([a-z0-9-]+)--[a-z0-9-]+\.netlify\.app$/.exec(host);
    var netlifyMain = !m && /\.netlify\.app$/.test(host);
    if (netlifyMain || (m && m[1] === "develop") || /QBR-Dashboard_dev/i.test(path))
      return { kind: "beta", text: "Beta version", tag: "[BETA]", tip: "Beta (develop branch) — for testing before release. Changes here are not live." };
    if (m) return { kind: "preview", text: "Preview · " + m[1], tag: "[PREVIEW]", tip: "Feature preview (" + m[1] + ") — not live, not beta." };
    // 2026-10-14: "Main" = the live build (main branch): GitHub Pages, the Azure production site, the live
    // OneDrive folder (QBR-Dashboard_<x.y.z>). Azure PR staging sites
    // ("<name>-<PR>.<region>.azurestaticapps.net") are labelled Staging, not Main.
    var az = /^([a-z0-9-]+)\.(?:[a-z0-9-]+\.)*azurestaticapps\.net$/.exec(host);
    var azPr = az && /^[a-z]+-[a-z]+-[0-9a-f]+-(\d+)$/.exec(az[1]);
    if (azPr) return { kind: "preview", text: "Staging · PR " + azPr[1], tag: "[STAGING]", tip: "Azure staging site for pull request #" + azPr[1] + " — not live." };
    // 2026-10-19: the Azure Static Web App deploys the develop branch → Beta, not Main
    if (az) return { kind: "beta", text: "Beta version", tag: "[BETA]", tip: "Beta on Azure (develop branch) — for testing before release. Changes here are not live." };
    if (/\.github\.io$/.test(host) || /QBR-Dashboard_\d+\.\d+\.\d+[\\\/]/i.test(path))
      return { kind: "main", text: "Main", tag: "[MAIN]", tip: "Main (live) — the production build everyone uses." };
    return null;
  }
  function devBadge() {
    var env = envInfo(); if (!env) return;
    var sub = document.querySelector(".app-sub"); if (!sub || document.getElementById("dev-badge")) return;
    var b = document.createElement("span");
    b.id = "dev-badge"; b.className = "dev-badge env-" + env.kind; b.textContent = env.text; b.title = env.tip;
    sub.insertBefore(b, sub.firstChild);
    if (document.title.indexOf(env.tag) !== 0) document.title = env.tag + " " + document.title;
  }
  window.RCT_ENV = envInfo;
  function init() {
    var btn = $("brand-btn"), menu = $("brand-menu"), file = $("brand-file");
    devBadge();
    apply(load());
    if (!btn || !menu || !file) return;
    function open(v) {
      menu.hidden = !v; btn.setAttribute("aria-expanded", v ? "true" : "false");
      if (v) { note(); var f = menu.querySelector(".brand-menu-i:not(:disabled)"); if (f) f.focus(); }
    }
    btn.addEventListener("click", function (e) { e.stopPropagation(); open(menu.hidden); });
    document.addEventListener("click", function (e) { if (!menu.hidden && !menu.contains(e.target)) open(false); });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape" && !menu.hidden) { open(false); btn.focus(); } });
    $("brand-upload").addEventListener("click", function () { file.value = ""; file.click(); });
    $("brand-reset").addEventListener("click", function () { save(null); apply(null); open(false); btn.focus(); });
    file.addEventListener("change", function () {
      var f = file.files && file.files[0]; if (!f) return;
      var type = f.type || (/\.svg$/i.test(f.name) ? "image/svg+xml" : "");
      if (OK_TYPES.indexOf(type) < 0) { note("Not an image we can use — choose SVG, PNG, JPG or WebP.", true); return; }
      if (f.size > MAX) { note("That file is " + Math.round(f.size / 1024) + " KB — please use one under 512 KB.", true); return; }
      var rd = new FileReader();
      rd.onload = function () {
        var url = String(rd.result);
        if (type === "image/svg+xml" && url.indexOf("data:image/svg+xml") !== 0) url = "data:image/svg+xml;base64," + url.split(",")[1];
        var probe = new Image();
        probe.onload = function () {
          if (!save(url)) { note("Couldn't save in this browser (storage full or blocked).", true); return; }
          apply(url); open(false); btn.focus();
        };
        probe.onerror = function () { note("That image couldn't be read — try another file.", true); };
        probe.src = url;
      };
      rd.onerror = function () { note("That file couldn't be read.", true); };
      rd.readAsDataURL(f);
    });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
