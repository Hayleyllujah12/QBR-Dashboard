/* ============================================================================
 * RCT OpsDesk — Netlify Identity gate (js/identity.js)
 * v1.31.0 — from Mars's patch 5/5, reworked at merge (2026-10-06)
 *
 * Online login for the Netlify-hosted preview/beta sites (*.netlify.app). Invite-only:
 * no login = the dashboard stays hidden. Everywhere else (local file://, the live folder,
 * GitHub Pages) this file does nothing and loads nothing.
 *
 * Merge changes vs the patch:
 *  - The Netlify widget script is loaded HERE, only on Netlify hosts. The patch put a
 *    <script src="https://identity.netlify.com/..."> tag in index.html, which would make the
 *    offline/live app fetch a third-party script on every load (breaks the no-CDN rule).
 *  - The gate is raised immediately (before the widget arrives) and hides the real
 *    containers (.app-header, #app-sidebar, main, #detail-pane — the patch's .topbar/.sidebar
 *    don't exist). Theme tokens instead of hard-coded light colours.
 *
 * NOTE: this hides the UI; it is not server-side access control — the site's files remain
 * downloadable. Workbook data is never part of the site (users load their own files).
 * Netlify settings needed: Identity enabled, Registration = Invite only.
 * ============================================================================ */
(function () {
  "use strict";
  const QBR = (window.QBR = window.QBR || {});
  const WIDGET = "https://identity.netlify.com/v1/netlify-identity-widget.js";

  function isNetlifyHost() {
    try { const h = window.location.hostname || ""; return /\.netlify\.app$/i.test(h) || /\.netlify\.com$/i.test(h); }
    catch (e) { return false; }
  }
  QBR.identityRequired = isNetlifyHost;

  function showLogin(note) {
    document.body.classList.add("identity-gated");
    document.body.classList.remove("identity-authed");
    let ov = document.getElementById("identity-gate");
    if (!ov) {
      ov = document.createElement("div");
      ov.id = "identity-gate"; ov.className = "identity-gate";
      ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true"); ov.setAttribute("aria-labelledby", "identity-title");
      ov.innerHTML =
        `<div class="identity-box">` +
        `<div class="identity-logo" id="identity-title">RCT OpsDesk</div>` +
        `<div class="identity-sub">Beta site — sign in to continue</div>` +
        `<button type="button" class="btn btn-primary" id="identity-login-btn">Log in</button>` +
        `<div class="identity-note" id="identity-note">Invite-only. Contact the administrator for access.</div>` +
        `</div>`;
      document.body.appendChild(ov);
      ov.querySelector("#identity-login-btn").addEventListener("click", () => {
        try { window.netlifyIdentity.open("login"); } catch (e) {}
      });
    }
    ov.style.display = "flex";
    if (note) ov.querySelector("#identity-note").textContent = note;
    const b = ov.querySelector("#identity-login-btn"); if (b) { b.disabled = !window.netlifyIdentity; }
  }
  function hideLogin() {
    const ov = document.getElementById("identity-gate");
    if (ov) ov.style.display = "none";
    document.body.classList.add("identity-authed");
    document.body.classList.remove("identity-gated");
  }
  function addLogout() {
    if (document.getElementById("identity-logout-btn")) return;
    const btn = document.createElement("button");
    btn.id = "identity-logout-btn"; btn.type = "button";
    btn.className = "btn btn-sm btn-outline-secondary identity-logout";
    btn.textContent = "Log out";
    btn.addEventListener("click", () => { try { window.netlifyIdentity.logout(); } catch (e) {} });
    const anchor = document.getElementById("btn-theme");
    if (anchor && anchor.parentNode) anchor.parentNode.insertBefore(btn, anchor.nextSibling);
    else document.body.appendChild(btn);
  }
  function removeLogout() { const b = document.getElementById("identity-logout-btn"); if (b) b.remove(); }

  function loadWidget(timeoutMs) {
    if (window.netlifyIdentity) return Promise.resolve(true);
    return new Promise(resolve => {
      const s = document.createElement("script");
      let done = false; const fin = ok => { if (!done) { done = true; resolve(ok && !!window.netlifyIdentity); } };
      s.src = WIDGET; s.async = true;
      s.onload = () => fin(true); s.onerror = () => fin(false);
      setTimeout(() => fin(false), timeoutMs || 10000);
      document.head.appendChild(s);
    });
  }

  QBR.identityInit = async function () {
    if (!isNetlifyHost()) return false;                    // local / live / GitHub Pages: no-op
    showLogin("Checking sign-in…");                        // gate first, before anything renders
    const ok = await loadWidget(QBR._identityTimeout);
    if (!ok) { showLogin("Login service unavailable. Check your connection, then reload."); return true; }
    const id = window.netlifyIdentity;
    // The widget may initialise itself before our handlers exist (its own "init" can fire first),
    // so don't wait for "init": enable the button now and check the current user directly.
    const sync = () => { let u = null; try { u = id.currentUser(); } catch (e) {} if (u) { hideLogin(); addLogout(); } else showLogin("Invite-only. Contact the administrator for access."); };
    id.on("init", sync);
    id.on("login", () => { hideLogin(); addLogout(); try { id.close(); } catch (e) {} });
    id.on("logout", () => { removeLogout(); showLogin("You have signed out."); });
    id.on("error", err => { console.warn("[QBR] Netlify Identity:", err && err.message); });
    try { id.init(); } catch (e) {}
    sync();
    return true;
  };

  try {
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", () => QBR.identityInit());
    else QBR.identityInit();
  } catch (e) {}
})();
