/* QBR Dashboard — user profiles (js/users.js)
 *
 * Per-user workspaces for the My To-Do list and general ticketing.
 * Each user gets their own data namespace (todos, tickets, history);
 * inventory and audits stay shared, exactly as before.
 *
 * Design (Robin, 2026-10-12): per-user WORKBOOKS, not sheets in a shared
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

  /* Ensure at least one user exists (first run -> "Robin" is NOT assumed;
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
    exportWorkbook: exportWorkbook
  };
})();
