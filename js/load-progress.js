/* Load progress bar (2026-10-25): one shared, non-modal indicator above the
 * app chrome for every app-level load (upload, folder link, reload, restore).
 * Determinate fill + phase label + live "i/N — filename" text, reusing the
 * Bootstrap .progress/.progress-bar convention from the MOS load card.
 * Sessions nest: only the outermost session drives the bar — a per-file
 * reload inside Reload-all stays silent so the outer "Reloading i/N" text
 * never fights the inner parse detail. Bar only: no Cancel button, by design
 * (the requester's pick). Always hidden on session end, including error paths: the
 * caller wraps its load in try/finally around begin()/end().
 * Exposed on QBR: QBR.loadProgress. */
(function () {
  var depth = 0, curFrac = 0, el = null, fill = null, track = null, label = null, detail = null;

  function ensure() {
    if (el) return;
    el = document.createElement("div");
    el.id = "qbr-loadbar";
    el.className = "qbr-loadbar";
    el.setAttribute("role", "status");
    el.setAttribute("aria-live", "polite");
    el.hidden = true;
    el.innerHTML = '<div class="qbr-loadbar-inner"><span class="qbr-loadbar-label"></span>' +
      '<div class="progress qbr-loadbar-track" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">' +
      '<div class="progress-bar qbr-loadbar-fill" style="width:0%"></div></div>' +
      '<span class="qbr-loadbar-detail"></span></div>';
    document.body.insertBefore(el, document.body.firstChild);
    fill = el.querySelector(".qbr-loadbar-fill");
    track = el.querySelector(".qbr-loadbar-track");
    label = el.querySelector(".qbr-loadbar-label");
    detail = el.querySelector(".qbr-loadbar-detail");
  }

  /* Generic setter. No-op unless a session is open (depth > 0); the caller
   * that owns the outermost session is responsible for calling it. */
  function set(labelText, frac, detailText, indet) {
    if (depth <= 0) return;
    ensure();
    label.textContent = labelText;
    detail.textContent = detailText || "";
    if (indet) {
      fill.style.width = "100%";
      fill.classList.add("progress-bar-striped", "progress-bar-animated");
      track.setAttribute("aria-valuenow", "0");
    } else {
      curFrac = frac || 0;
      var pct = Math.max(0, Math.min(100, Math.round((frac || 0) * 100)));
      fill.style.width = pct + "%";
      fill.classList.remove("progress-bar-striped", "progress-bar-animated");
      track.setAttribute("aria-valuenow", String(pct));
    }
  }

  function fracText(i, n, name) {
    return i + "/" + n + (name ? " — " + name : "");
  }

  QBR.loadProgress = {
    /* Returns true when this call opened the outermost session (the caller
     * then owns the bar); false for nested sessions, which stay silent. */
    begin: function () {
      ensure();
      depth++;
      if (depth === 1) { el.hidden = false; set("Loading…", 0, "", false); }
      return depth === 1;
    },
    end: function () {
      if (depth > 0) depth--;
      if (depth === 0 && el) el.hidden = true;
    },
    /* Yield so the browser can paint the bar between files. The parse is
     * synchronous otherwise, and a bar that never paints is no better than
     * the static "Parsing…" text it replaces. */
    yield: function () { return new Promise(function (r) { setTimeout(r, 0); }); },
    /* Flip the label (e.g. Reading -> Parsing) without resetting progress. */
    phase: function (labelText) { set(labelText, curFrac, "", false); },
    scan: function (entries) {
      set("Scanning…", 0, entries + " entr" + (entries === 1 ? "y" : "ies") + " found", true);
    },
    read: function (i, n, name) {
      set("Reading", n ? i / n : 0, n ? fracText(i, n, name) : "Reading…", false);
    },
    parse: function (done, total, i, n, name) {
      set("Parsing", total ? done / total : 0, n ? fracText(i, n, name) : "Parsing…", false);
    },
    render: function () { set("Rendering…", 1, "", false); },
    reload: function (i, n, name) {
      set("Reloading", n ? i / n : 0, n ? fracText(i, n, name) : "Reloading…", false);
    }
  };
})();
