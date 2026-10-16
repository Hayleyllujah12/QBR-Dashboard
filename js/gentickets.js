/* QBR Dashboard — general ticketing (js/gentickets.js)
 *
 * Separate queue from Lenovo hardware tickets (Robin, 2026-10-12).
 * Per-user: tickets live in the current user's data namespace.
 * Numbering: REQ-YYYYMMDD-## (daily sequence).
 * Fields mirror Ticket 360 (priority, statuses, follow-ups/notes) minus serial;
 * an optional asset serial may link to Asset 360.
 * Import: Excel/CSV file, preview-then-confirm, dedup by fingerprint
 * (date + time + title + requester) — mirrors the Lenovo Forms import.
 * Every open general ticket auto-appears in My To-Do (see js/todo.js).
 */

(function () {
  "use strict";

  var STATUSES = ["Open", "In Progress", "Waiting", "Escalated", "Resolved"];
  var PRIORITIES = ["High", "Medium", "Low"];
  var CLOSED_RE = /^(resolved|completed|closed)$/i;

  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function $(id) { return document.getElementById(id); }

  function isOpen(t) { return !CLOSED_RE.test(String((t && t.status) || "Open")); }

  function list() {
    var d = QBR.users.myData();
    return (d.tickets || []).slice();
  }

  function saveAll(arr) {
    var d = QBR.users.myData();
    d.tickets = arr;
    QBR.users.saveMyData(d);
  }

  function todayStamp() {
    var d = new Date();
    return d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
  }

  function nextTno() {
    var stamp = todayStamp(), prefix = "REQ-" + stamp + "-";
    var n = 0;
    list().forEach(function (t) {
      if (t.tno && t.tno.indexOf(prefix) === 0) {
        var k = parseInt(t.tno.slice(prefix.length), 10);
        if (k > n) n = k;
      }
    });
    return prefix + String(n + 1).padStart(2, "0");
  }

  function create(f) {
    f = f || {};
    var title = String(f.title || "").trim();
    if (!title) return { ok: false, error: "Give the ticket a title." };
    var arr = list();
    var t = {
      tno: nextTno(),
      title: title,
      requester: String(f.requester || "").trim() || null,
      client: String(f.client || "").trim() || null,
      priority: PRIORITIES.indexOf(f.priority) >= 0 ? f.priority : "Medium",
      status: "Open",
      due: f.due || null,
      serial: String(f.serial || "").trim().toUpperCase() || null,
      notes: [],
      createdAt: new Date().toISOString()
    };
    if (f.notes) t.notes.push({ d: new Date().toISOString().slice(0, 10), text: String(f.notes).trim() });
    arr.unshift(t);
    saveAll(arr);
    return { ok: true, ticket: t };
  }

  function find(tno) {
    return list().filter(function (t) { return t.tno === tno; })[0] || null;
  }

  function setStatus(tno, status) {
    if (STATUSES.indexOf(status) < 0) return false;
    var arr = list(), t = arr.filter(function (x) { return x.tno === tno; })[0];
    if (!t) return false;
    t.status = status;
    saveAll(arr);
    return true;
  }

  function addNote(tno, text) {
    text = String(text || "").trim();
    if (!text) return false;
    var arr = list(), t = arr.filter(function (x) { return x.tno === tno; })[0];
    if (!t) return false;
    if (!Array.isArray(t.notes)) t.notes = t.notes ? [{ d: "", text: String(t.notes) }] : [];   // older string notes
    t.notes.push({ d: localDay(new Date().toISOString()), text: text });
    saveAll(arr);
    return true;
  }

  /* --- import: fingerprint + preview ---------------------------------- */
  function fingerprint(row) {
    var parts = [row.date, row.time, row.title, row.requester].map(function (c) {
      return String(c == null ? "" : c).trim().replace(/\s+/g, " ").toLowerCase();
    });
    var s = parts.join(""), h = 5381;
    for (var i = 0; i < s.length; i++) h = (((h << 5) + h) ^ s.charCodeAt(i)) >>> 0;
    return h.toString(16);
  }

  /* Local YYYY-MM-DD (toISOString is UTC and can be a day off). */
  function localDay(iso) {
    var d = iso ? new Date(iso) : new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  }

  /* Import ledger per user (was one global key: a second user importing the same file got 0 tickets). */
  function seenKey() {
    var u = (QBR.users && QBR.users.current) ? QBR.users.current() : null;
    return "qbr-gtix-import-v1:" + ((u && u.id) || "nouser");
  }
  /* Excel date cells / common date text -> local YYYY-MM-DD; "" when not a date. */
  function normDate(v) {
    if (v instanceof Date && !isNaN(v)) return v.getFullYear() + "-" + String(v.getMonth() + 1).padStart(2, "0") + "-" + String(v.getDate()).padStart(2, "0");
    var t = String(v == null ? "" : v).trim(), m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t))) return m[1] + "-" + m[2].padStart(2, "0") + "-" + m[3].padStart(2, "0");
    if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t))) return m[3] + "-" + m[1].padStart(2, "0") + "-" + m[2].padStart(2, "0");   // m/d/yyyy (Excel US)
    return "";
  }

  function existingFingerprints() {
    var set = {};
    list().forEach(function (t) { set[fingerprint({ date: localDay(t.createdAt), time: "", title: t.title, requester: t.requester })] = 1; });
    try {
      var seen = JSON.parse(localStorage.getItem(seenKey()) || "{}");
      Object.keys(seen).forEach(function (k) { set[k] = 1; });
    } catch (e) {}
    return set;
  }

  /* Parse an Excel/CSV file into candidate rows. Expected headers (case-insensitive):
   * Title*, Requester, Client, Priority, Due date, Date, Time, Notes, Serial. */
  function parseFile(file) {
    return new Promise(function (res, rej) {
      if (typeof XLSX === "undefined") return rej(new Error("Excel library not loaded."));
      var rd = new FileReader();
      rd.onload = function () {
        try {
          var wb = XLSX.read(rd.result, { type: "array", cellDates: true });
          var ws = wb.Sheets[wb.SheetNames[0]];
          var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
          if (!rows.length) return res([]);
          var heads = rows[0].map(function (h) { return String(h).trim().toLowerCase(); });
          var ci = function () {
            for (var i = 0; i < arguments.length; i++) {
              var x = heads.indexOf(arguments[i]);
              if (x >= 0) return x;
            }
            return -1;
          };
          var c = {
            title: ci("title", "request", "subject"), requester: ci("requester", "requested by", "from"),
            client: ci("client", "school", "organization"), priority: ci("priority"),
            due: ci("due date", "due"), date: ci("date", "created", "submitted"),
            time: ci("time"), notes: ci("notes", "description", "details"), serial: ci("serial", "s/n")
          };
          if (c.title < 0) return rej(new Error("No Title column found."));
          var out = [];
          for (var r = 1; r < rows.length; r++) {
            var row = rows[r], g = function (i) {
              if (i < 0 || row[i] == null) return "";
              var v = row[i];
              if (v instanceof Date) return (i === c.time) ? String(v.getHours()).padStart(2, "0") + ":" + String(v.getMinutes()).padStart(2, "0") : normDate(v);
              return String(v).trim();
            };
            if (!g(c.title)) continue;
            out.push({
              title: g(c.title), requester: g(c.requester), client: g(c.client),
              priority: g(c.priority), due: normDate(g(c.due)), date: normDate(g(c.date)) || g(c.date), time: g(c.time),
              notes: g(c.notes), serial: g(c.serial)
            });
          }
          res(out);
        } catch (e) { rej(e); }
      };
      rd.onerror = function () { rej(new Error("Could not read file.")); };
      rd.readAsArrayBuffer(file);
    });
  }

  function previewImport(cands) {
    var seen = existingFingerprints(), inFile = {};
    return cands.map(function (r) {
      var fp = fingerprint(r);
      var dupe = !!seen[fp] || !!inFile[fp];   // also the same row twice in one file
      inFile[fp] = 1;
      return { row: r, fp: fp, dupe: dupe };
    });
  }

  function confirmImport(items) {
    var seen = {};
    try { seen = JSON.parse(localStorage.getItem(seenKey()) || "{}"); } catch (e) {}
    var n = 0, skipped = 0;
    items.forEach(function (it) {
      if (it.dupe || seen[it.fp]) { skipped++; return; }
      var r = create({
        title: it.row.title, requester: it.row.requester, client: it.row.client,
        priority: it.row.priority, due: it.row.due || null,
        serial: it.row.serial, notes: it.row.notes
      });
      if (r.ok) { n++; seen[it.fp] = 1; }
    });
    try { localStorage.setItem(seenKey(), JSON.stringify(seen)); } catch (e) {}
    return { imported: n, skipped: skipped };
  }

  /* --- render ----------------------------------------------------------- */
  function statusBadge(s) {
    var map = { "Open": "warning", "In Progress": "primary", "Waiting": "secondary", "Escalated": "danger", "Resolved": "success" };
    return '<span class="badge bg-' + (map[s] || "secondary") + '">' + esc(s) + "</span>";
  }

  function render() {
    var host = $("dash-gtickets");
    if (!host) return;
    var u = QBR.users.ensure();
    if (!u) {
      host.innerHTML = '<div class="card-box text-center py-5"><h4>Pick your profile first</h4>' +
        '<p class="text-muted">General tickets are per-user. Choose your profile on the My To-Do page.</p></div>';
      return;
    }
    var arr = list();
    var rows = arr.map(function (t) {
      return '<tr><td class="font-monospace small">' + esc(t.tno) + "</td>" +
        "<td>" + esc(t.title) +
        (t.notes && t.notes.length ? '<div class="small text-muted">' + esc(t.notes[t.notes.length - 1].text).slice(0, 80) + "</div>" : "") + "</td>" +
        "<td>" + esc(t.requester || "—") + "</td>" +
        "<td>" + esc(t.client || "—") + "</td>" +
        "<td>" + esc(t.priority || "—") + "</td>" +
        "<td>" + statusBadge(t.status || "Open") + "</td>" +
        "<td class=\"small\">" + esc(t.due || "—") + "</td>" +
        '<td><div class="d-flex gap-1">' +
        '<select class="form-select form-select-sm w-auto" data-gt-status="' + esc(t.tno) + '">' +
        STATUSES.map(function (s) { return '<option' + (s === t.status ? " selected" : "") + ">" + s + "</option>"; }).join("") +
        "</select>" +
        '<button class="btn btn-sm btn-outline-primary" data-gt-note="' + esc(t.tno) + '">+ Note</button>' +
        "</div></td></tr>";
    }).join("");

    host.innerHTML =
      '<div class="d-flex flex-wrap align-items-center gap-2 mb-3"><h4 class="mb-0">General tickets</h4>' +
      '<span class="badge bg-primary">' + esc(u.name) + "</span>" +
      '<button class="btn btn-sm btn-primary" data-gt-act="new">+ New ticket</button>' +
      '<label class="btn btn-sm btn-outline-secondary mb-0">Import file<input type="file" id="gt-import-file" accept=".xlsx,.xls,.csv" hidden></label>' +
      "</div>" +
      '<div class="card-box"><div class="table-responsive"><table class="table table-sm">' +
      "<thead><tr><th>Ticket #</th><th>Title</th><th>Requester</th><th>Client</th><th>Priority</th><th>Status</th><th>Due</th><th></th></tr></thead>" +
      "<tbody>" + (rows || '<tr><td colspan="8" class="text-muted">No tickets yet.</td></tr>') + "</tbody></table></div></div>" +
      '<div id="gt-modal-host"></div>';

    host.querySelectorAll("[data-gt-act]").forEach(function (b) {
      if (b.dataset.gtAct === "new") b.addEventListener("click", newDialog);
    });
    host.querySelectorAll("[data-gt-status]").forEach(function (sel) {
      sel.addEventListener("change", function () { setStatus(sel.dataset.gtStatus, sel.value); render(); });
    });
    host.querySelectorAll("[data-gt-note]").forEach(function (b) {
      b.addEventListener("click", function () {
        var v = prompt("Add a follow-up note to " + b.dataset.gtNote + ":");
        if (v) { addNote(b.dataset.gtNote, v); render(); }
      });
    });
    var fi = $("gt-import-file");
    if (fi) fi.addEventListener("change", function () {
      if (fi.files.length) importDialog(fi.files[0]);
      fi.value = "";
    });
  }

  function field(id, label, inner) {
    return '<label class="form-label small">' + label + "</label>" + inner;
  }

  function newDialog() {
    var host = $("gt-modal-host");
    host.innerHTML = '<div class="modal-overlay open"><div class="scan-dialog card-box" role="dialog" aria-modal="true">' +
      "<h6>New general ticket</h6>" +
      field(0, "Title *", '<input id="gt-n-title" class="form-control form-control-sm mb-2">') +
      '<div class="row g-2 mb-2"><div class="col-6">' +
      field(0, "Requester", '<input id="gt-n-req" class="form-control form-control-sm">') + "</div><div class=\"col-6\">" +
      field(0, "Client / school", '<input id="gt-n-client" class="form-control form-control-sm">') + "</div></div>" +
      '<div class="row g-2 mb-2"><div class="col-4">' +
      field(0, "Priority", '<select id="gt-n-prio" class="form-select form-select-sm">' +
        PRIORITIES.map(function (p) { return '<option' + (p === "Medium" ? " selected" : "") + ">" + p + "</option>"; }).join("") + "</select>") +
      '</div><div class="col-4">' + field(0, "Due date", '<input id="gt-n-due" type="date" class="form-control form-control-sm">') +
      '</div><div class="col-4">' + field(0, "Asset serial (optional)", '<input id="gt-n-serial" class="form-control form-control-sm" style="font-family:monospace">') + "</div></div>" +
      field(0, "Notes", '<textarea id="gt-n-notes" class="form-control form-control-sm mb-3" rows="2"></textarea>') +
      '<div class="d-flex gap-2"><button class="btn btn-sm btn-primary" id="gt-n-save">Create</button>' +
      '<button class="btn btn-sm btn-outline-secondary" id="gt-n-cancel">Cancel</button></div>' +
      '<div id="gt-n-msg" class="small text-danger mt-2"></div></div></div>';
    var close = function () { host.innerHTML = ""; };
    $("gt-n-cancel").addEventListener("click", close);
    host.querySelector(".modal-overlay").addEventListener("click", function (e) {
      if (e.target.classList.contains("modal-overlay")) close();
    });
    $("gt-n-save").addEventListener("click", function () {
      var r = create({
        title: $("gt-n-title").value, requester: $("gt-n-req").value,
        client: $("gt-n-client").value, priority: $("gt-n-prio").value,
        due: $("gt-n-due").value || null, serial: $("gt-n-serial").value,
        notes: $("gt-n-notes").value
      });
      if (!r.ok) { $("gt-n-msg").textContent = r.error; return; }
      close(); render();
    });
  }

  function importDialog(file) {
    var host = $("gt-modal-host");
    host.innerHTML = '<div class="modal-overlay open"><div class="scan-dialog scan-dialog-wide card-box" role="dialog" aria-modal="true">' +
      "<h6>Import general tickets</h6><p class=\"small text-muted\">Reading file…</p></div></div>";
    parseFile(file).then(function (cands) {
      var prev = previewImport(cands);
      var newOnes = prev.filter(function (p) { return !p.dupe; });
      var rows = prev.map(function (p, i) {
        return "<tr class=\"" + (p.dupe ? "table-warning" : "") + "\"><td>" + (i + 1) + "</td>" +
          "<td>" + esc(p.row.title) + "</td><td>" + esc(p.row.requester || "—") + "</td>" +
          "<td class=\"small\">" + esc([p.row.date, p.row.time].filter(Boolean).join(" ")) + "</td>" +
          "<td>" + (p.dupe ? '<span class="badge bg-warning text-dark">duplicate — skipped</span>' : '<span class="badge bg-success">new</span>') + "</td></tr>";
      }).join("");
      host.querySelector(".card-box").innerHTML = "<h6>Import general tickets</h6>" +
        '<p class="small">' + newOnes.length + " new, " + (prev.length - newOnes.length) + " duplicates skipped.</p>" +
        '<div class="table-responsive" style="max-height:300px;overflow:auto"><table class="table table-sm">' +
        "<thead><tr><th>#</th><th>Title</th><th>Requester</th><th>Date</th><th></th></tr></thead><tbody>" + rows + "</tbody></table></div>" +
        '<div class="d-flex gap-2 mt-2"><button class="btn btn-sm btn-primary" id="gt-i-go" ' +
        (newOnes.length ? "" : "disabled") + ">Import " + newOnes.length + "</button>" +
        '<button class="btn btn-sm btn-outline-secondary" id="gt-i-cancel">Cancel</button></div>';
      $("gt-i-cancel").addEventListener("click", function () { host.innerHTML = ""; });
      var go = $("gt-i-go");
      if (go) go.addEventListener("click", function () {
        var r = confirmImport(newOnes);
        host.innerHTML = "";
        render();
        alert("Imported " + r.imported + ", skipped " + r.skipped + " duplicates.");
      });
    }).catch(function (e) {
      host.querySelector(".card-box").innerHTML = "<h6>Import general tickets</h6>" +
        '<p class="text-danger small">' + esc(e.message) + "</p>" +
        '<button class="btn btn-sm btn-outline-secondary" id="gt-i-cancel">Close</button>';
      $("gt-i-cancel").addEventListener("click", function () { host.innerHTML = ""; });
    });
  }

  /* Re-render when the current user or their data changes (user switch,
   * register) so this panel never shows another user's tickets. */
  if (typeof window.addEventListener === "function") {
    var _rerenderPending = false;
    window.addEventListener("qbr:userdata", function () {
      if (_rerenderPending) return;
      _rerenderPending = true;
      setTimeout(function () {
        _rerenderPending = false;
        try { render(); } catch (e) {}
      }, 0);
    });
  }

  /* Public API */
  window.QBR = window.QBR || {};
  QBR.gtickets = {
    list: list, create: create, find: find, setStatus: setStatus, addNote: addNote,
    isOpen: isOpen, nextTno: nextTno, fingerprint: fingerprint,
    parseFile: parseFile, previewImport: previewImport, confirmImport: confirmImport,
    render: render, STATUSES: STATUSES
  };
})();
