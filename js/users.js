/* QBR Dashboard — user profiles (js/users.js)
 *
 * Per-user workspaces for the My To-Do list and general ticketing.
 * Each user gets their own data namespace (todos, tickets, history);
 * inventory and audits stay shared, exactly as before.
 *
 * Design (owner, 2026-10-12): per-user WORKBOOKS, not sheets in a shared
 * file — each user exports their own file, so concurrent edits can never
 * collide. No per-user PIN (a locally-checked PIN is security theater when
 * anyone holding the file can bypass it); the admin gate is unchanged.
 *
 * Storage (all local, no network):
 *   qbr-users-v1        [{ id, name, createdAt }]
 *   qbr-user-current    <id>
 *   qbr-userdata-<id>   { todos: [], tickets: [], history: [] }
 */

(function () {
  "use strict";

  var LS_USERS = "qbr-users-v1";
  var LS_CURRENT = "qbr-user-current";
  var LS_DATA_PREFIX = "qbr-userdata-";

  function lsGet(k, fb) {
    try {
      var v = localStorage.getItem(k);
      return v == null ? fb : JSON.parse(v);
    } catch (e) { return fb; }
  }
  function lsSet(k, v) {
    try { localStorage.setItem(k, JSON.stringify(v)); return true; }
    catch (e) { return false; }
  }

  function newId() {
    return "u" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  /* Registry ------------------------------------------------------------ */
  function listUsers() {
    var a = lsGet(LS_USERS, []);
    return Array.isArray(a) ? a : [];
  }

  function saveUsers(a) { lsSet(LS_USERS, a); }

  function currentUserId() {
    return lsGet(LS_CURRENT, null);
  }

  function currentUser() {
    var id = currentUserId();
    if (!id) return null;
    var u = listUsers().filter(function (x) { return x.id === id; })[0];
    return u || null;
  }

  /* Name -> safe token for filenames/keys. Excel sheet names forbid
   * []:*?/\ and cap at 31 chars; we apply the same rule everywhere. */
  function safeName(name) {
    var s = String(name == null ? "" : name).trim().replace(/[\[\]:*?\/\\]/g, "").replace(/\s+/g, " ");
    if (s.length > 31) s = s.slice(0, 31).trim();
    return s;
  }

  function registerUser(name) {
    var clean = safeName(name);
    if (!clean) return { ok: false, error: "Enter a name." };
    var users = listUsers();
    var dup = users.some(function (u) { return u.name.toLowerCase() === clean.toLowerCase(); });
    if (dup) return { ok: false, error: "That name is already registered." };
    var u = { id: newId(), name: clean, createdAt: new Date().toISOString() };
    users.push(u);
    saveUsers(users);
    blankData(u.id);
    lsSet(LS_CURRENT, u.id);
    notifyChange();
    return { ok: true, user: u };
  }

  function switchUser(id) {
    var u = listUsers().filter(function (x) { return x.id === id; })[0];
    if (!u) return false;
    blankData(id); // ensure namespace exists
    lsSet(LS_CURRENT, id);
    notifyChange();
    return true;
  }

  /* Ensure at least one user exists (first run -> no default user is assumed;
   * the user picks or registers their own name). */
  function ensureUser() {
    var u = currentUser();
    if (u) return u;
    var users = listUsers();
    if (users.length === 1) {
      lsSet(LS_CURRENT, users[0].id);
      return users[0];
    }
    return null; // picker must choose (0 or 2+ users)
  }

  /* Per-user data ------------------------------------------------------- */
  function dataKey(id) { return LS_DATA_PREFIX + id; }

  function blankData(id) {
    var k = dataKey(id);
    if (lsGet(k, null) == null) lsSet(k, { todos: [], tickets: [], history: [] });
  }

  function getData(id) {
    blankData(id);
    var d = lsGet(dataKey(id), null) || {};
    if (!Array.isArray(d.todos)) d.todos = [];
    if (!Array.isArray(d.tickets)) d.tickets = [];
    if (!Array.isArray(d.history)) d.history = [];
    return d;
  }

  function setData(id, d) { return lsSet(dataKey(id), d); }

  function myData() {
    var u = ensureUser();
    if (!u) return { todos: [], tickets: [], history: [], _nouser: true };
    return getData(u.id);
  }

  function saveMyData(d) {
    var u = ensureUser();
    if (!u) return false;
    var ok = setData(u.id, d);
    if (ok) notifyChange();
    return ok;
  }

  /* Tell the views (My To-Do, General tickets) that the current user or their
   * data changed so they re-render instead of showing a stale/other user's list. */
  function notifyChange() {
    try {
      if (typeof window.dispatchEvent === "function" && typeof CustomEvent === "function") {
        window.dispatchEvent(new CustomEvent("qbr:userdata"));
      }
    } catch (e) {}
  }

  /* Ticket notes are [{d, text}] (General tickets); older data may be a string. */
  function ticketNotesText(n) {
    if (Array.isArray(n)) {
      return n.map(function (x) { return (x && x.d ? x.d + ": " : "") + (x && x.text != null ? x.text : x); }).join("\n");
    }
    return n || "";
  }

  /* Per-user workbook export: my todos + my tickets, filename from my name. */
  function exportWorkbook() {
    if (typeof XLSX === "undefined") return { ok: false, error: "Excel library not loaded." };
    var u = ensureUser();
    if (!u) return { ok: false, error: "Pick or register a user first." };
    var d = getData(u.id);
    var wb = XLSX.utils.book_new();

    var th = [["ID", "Title", "Product", "Due date", "Priority", "Status", "Notes", "Created"]];
    var rows = [th];
    d.todos.forEach(function (t) {
      rows.push([t.id, t.title, t.product, t.due || "", t.priority || "auto",
        t.done ? "Done" : "Open", t.notes || "", (t.createdAt || "").slice(0, 10)]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "My Todos");

    var gh = [["Ticket #", "Title", "Requester", "Client", "Priority", "Status", "Due date", "Notes"]];
    var grows = [gh];
    d.tickets.forEach(function (t) {
      grows.push([t.tno, t.title, t.requester || "", t.client || "", t.priority || "",
        t.status || "Open", t.due || "", ticketNotesText(t.notes)]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(grows), "My Tickets");

    var hh = [["Completed", "Type", "Title", "Product"]];
    var hrows = [hh];
    d.history.forEach(function (h) {
      hrows.push([(h.doneAt || "").slice(0, 10), h.kind || "todo", h.title || "", h.product || ""]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(hrows), "History");

    var fname = safeName(u.name).replace(/\s+/g, "-") + "-todo.xlsx";
    XLSX.writeFile(wb, fname);
    return { ok: true, file: fname };
  }

  /* 2026-10-27: re-open a previously exported "<Name>-todo.xlsx" (sheets My Todos / My Tickets / History,
   * exactly the layout exportWorkbook writes). MERGE only — nothing in the current list is deleted:
   * to-dos match by ID, tickets by Ticket #, history by date+title; a match is updated from the file,
   * anything new is added. With no user picked yet, the name in the file name is registered first. */
  function parseExport(buf) {
    var wb = XLSX.read(buf, { type: "array", cellDates: false });
    var sh = function (n) {
      var name = wb.SheetNames.filter(function (x) { return x.trim().toLowerCase() === n; })[0];
      return name ? XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: false }) : null;
    };
    var T = sh("my todos"), K = sh("my tickets"), H = sh("history");
    if (!T && !K) return { ok: false, error: "This isn't a My To-Do export (no \"My Todos\" or \"My Tickets\" sheet)." };
    var s = function (v) { return v == null ? "" : String(v).trim(); };
    var ymd = function (v) { var x = s(v); var m = /^(\d{4}-\d{2}-\d{2})/.exec(x); return m ? m[1] : (x || null); };
    var notesBack = function (txt) {
      return s(txt).split(/\r?\n/).filter(Boolean).map(function (l) {
        var m = /^(\d{4}-\d{2}-\d{2}):\s?(.*)$/.exec(l); return m ? { d: m[1], text: m[2] } : { d: "", text: l };
      });
    };
    var todos = (T || []).slice(1).filter(function (r) { return s(r[1]); }).map(function (r) {
      return { id: s(r[0]) || ("t" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)), title: s(r[1]),
        product: s(r[2]) || "general", due: ymd(r[3]), priority: s(r[4]) || "auto", done: /^done$/i.test(s(r[5])),
        notes: s(r[6]), createdAt: ymd(r[7]) ? ymd(r[7]) + "T00:00:00.000Z" : new Date().toISOString() };
    });
    var tickets = (K || []).slice(1).filter(function (r) { return s(r[0]) || s(r[1]); }).map(function (r) {
      return { tno: s(r[0]), title: s(r[1]), requester: s(r[2]) || null, client: s(r[3]) || null, priority: s(r[4]) || "Medium",
        status: s(r[5]) || "Open", due: ymd(r[6]), serial: null, notes: notesBack(r[7]), createdAt: new Date().toISOString() };
    });
    var history = (H || []).slice(1).filter(function (r) { return s(r[2]); }).map(function (r) {
      return { doneAt: ymd(r[0]) ? ymd(r[0]) + "T00:00:00.000Z" : "", kind: s(r[1]) || "todo", title: s(r[2]), product: s(r[3]) };
    });
    return { ok: true, todos: todos, tickets: tickets, history: history };
  }
  function nameFromFile(fileName) {
    var m = /^(.*?)-todo(?:\s*\(\d+\))?\.xlsx$/i.exec(String(fileName || "").split(/[\\/]/).pop());
    return m ? safeName(m[1].replace(/-/g, " ")) : "";
  }
  function importWorkbook(buf, fileName) {
    if (typeof XLSX === "undefined") return { ok: false, error: "Excel library not loaded." };
    var p;
    try { p = parseExport(buf); } catch (e) { return { ok: false, error: "Couldn't read the file: " + (e && e.message) }; }
    if (!p.ok) return p;
    var u = ensureUser();
    if (!u) {
      var nm = nameFromFile(fileName);
      if (!nm) return { ok: false, error: "Pick or register a user first." };
      var ex = listUsers().filter(function (x) { return x.name.toLowerCase() === nm.toLowerCase(); })[0];
      if (ex) switchUser(ex.id); else { var r = registerUser(nm); if (!r.ok) return r; }
      u = currentUser();
    }
    var d = getData(u.id), n = { todosAdded: 0, todosUpdated: 0, ticketsAdded: 0, ticketsUpdated: 0, history: 0 };
    p.todos.forEach(function (t) {
      var cur = d.todos.filter(function (x) { return x.id === t.id; })[0];
      if (cur) { Object.assign(cur, t, { createdAt: cur.createdAt || t.createdAt }); n.todosUpdated++; } else { d.todos.push(t); n.todosAdded++; }
    });
    p.tickets.forEach(function (t) {
      var cur = t.tno ? d.tickets.filter(function (x) { return String(x.tno) === t.tno; })[0] : null;
      if (cur) { Object.assign(cur, t, { serial: cur.serial || null, createdAt: cur.createdAt || t.createdAt, notes: t.notes.length ? t.notes : (cur.notes || []) }); n.ticketsUpdated++; }
      else { d.tickets.push(t); n.ticketsAdded++; }
    });
    var hk = function (h) { return (h.doneAt || "").slice(0, 10) + "|" + h.title; };
    var have = {}; d.history.forEach(function (h) { have[hk(h)] = 1; });
    p.history.forEach(function (h) { if (!have[hk(h)]) { d.history.push(h); have[hk(h)] = 1; n.history++; } });
    setData(u.id, d); notifyChange();
    return { ok: true, user: u, counts: n, fileUser: nameFromFile(fileName) };
  }
  // Ask for a file, confirm, import, report. Shared by My To-Do and General tickets.
  function importDialog() {
    var inp = document.createElement("input"); inp.type = "file"; inp.accept = ".xlsx";
    inp.addEventListener("change", function () {
      var f = inp.files && inp.files[0]; if (!f) return;
      var u = currentUser(), fu = nameFromFile(f.name);
      var q = "Open \"" + f.name + "\"" + (u ? " into " + u.name + "'s list?" : "?") +
        (u && fu && fu.toLowerCase() !== u.name.toLowerCase() ? "\n\nNote: this file looks like " + fu + "'s export." : "") +
        "\n\nNothing is deleted: items with the same ID / Ticket # are updated from the file, new ones are added.";
      if (!confirm(q)) return;
      f.arrayBuffer().then(function (buf) {
        var r = importWorkbook(new Uint8Array(buf), f.name);
        if (!r.ok) { alert(r.error); return; }
        var c = r.counts;
        alert("Imported into " + r.user.name + ": " + c.todosAdded + " new + " + c.todosUpdated + " updated to-dos, " +
          c.ticketsAdded + " new + " + c.ticketsUpdated + " updated tickets, " + c.history + " history rows.");
      });
    });
    inp.click();
  }

  /* Public API */
  window.QBR = window.QBR || {};
  QBR.users = {
    list: listUsers,
    current: currentUser,
    currentId: currentUserId,
    ensure: ensureUser,
    register: registerUser,
    switch: switchUser,
    safeName: safeName,
    myData: myData,
    saveMyData: saveMyData,
    exportWorkbook: exportWorkbook,
    importWorkbook: importWorkbook,
    importDialog: importDialog
  };
})();
