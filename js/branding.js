/* RCT OpsDesk — branding.js (2026-10-05)
 * Custom logo by manual upload. Click the top-bar logo → "Upload new logo…" / "Reset to default".
 * The image is stored as a data: URL in localStorage ("rct-logo-v1") — this browser only, never sent anywhere.
 * Default logo = assets/rct-opsdesk-logo.svg (replace that file to change the logo for everyone).
 * Shown via <img>, so an uploaded SVG cannot run scripts. No data logic; safe to remove. */
(function () {
  "use strict";
  var KEY = "rct-logo-v1", MAX = 512 * 1024, DEFAULT = "assets/rct-opsdesk-logo.svg";
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

  function init() {
    var btn = $("brand-btn"), menu = $("brand-menu"), file = $("brand-file");
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
