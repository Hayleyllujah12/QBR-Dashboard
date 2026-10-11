/* QBR Dashboard — My To-Do list (js/todo.js)
 *
 * Overview of current tasks + unfinished responsibilities, per user.
 * Sources: manual entries AND auto-pulled items (open tickets, unaudited
 * schools). Tiles per product. Filter by urgency + deadlines.
 * Done items disappear into a per-user history view.
 *
 * Tile rule (one tile per task, no double counting):
 *   manual todos  -> their product tile (Microsoft / Lenovo / General)
 *   open tickets  -> "Pending tickets" tile (Lenovo + general tickets)
 *   unaudited     -> "Unaudited schools" tile
 *
 * Urgency: overdue -> high, due within 7 days -> medium, else low.
 * A manual priority override always wins.
 */

(function () {
  "use strict";

  var PRODUCTS = [
    { id: "microsoft", label: "Microsoft" },
    { id: "lenovo", label: "Lenovo" },
    { id: "tickets", label: "Pending tickets" },
    { id: "schools", label: "Unaudited schools" },
    { id: "general", label: "General" }
  ];
  var MANUAL_PRODUCTS = ["microsoft", "lenovo", "general"];

  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function $(id) { return document.getElementById(id); }

  /* --- urgency --------------------------------------------------------- */
  function daysUntil(due) {
    if (!due) return null;
    var t = new Date(due + "T00:00:00"), now = new Date();
    now.setHours(0, 0, 0, 0);
    return Math.round((t - now) / 86400000);
  }

  function autoUrgency(due) {
    var d = daysUntil(due);
    if (d == null) return "low";
    if (d < 0) return "high";
    if (d <= 7) return "medium";
    return "low";
  }

  function urgencyOf(item) {
    if (item.priority && item.priority !== "auto") return item.priority;
    return autoUrgency(item.due);
  }

  /* --- task assembly --------------------------------------------------- */
  function manualTasks() {
    var d = QBR.users.myData();
    return (d.todos || []).filter(function (t) { return !t.done; }).map(function (t) {
      return {
        id: "m:" + t.id, auto: false, title: t.title, product: t.product,
        due: t.due || null, priority: t.priority || "auto", notes: t.notes || "",
        ref: t, urgency: urgencyOf(t)
      };
    });
  }

  function openLenovoTickets() {
    var out = [];
    try {
      var inv = (typeof APP !== "undefined" && APP.model && APP.model.inventory) || null;
      if (!inv || !inv.tickets) return out;
      inv.tickets.forEach(function (t) {
        if (typeof invTixOpen === "function" ? !invTixOpen(t) : /^(resolved|completed|closed)$/i.test(String(t.status || ""))) return;
        out.push({
          id: "auto:lenovo:" + (t.tno || t.key || Math.random().toString(36).slice(2)),
          auto: true, kind: "ticket",
          title: (t.tno || "Ticket") + " — " + (t.issue || t.subject || "Lenovo ticket"),
          product: "tickets", due: t.followUp || t.due || null,
          notes: t.sn ? "Serial " + t.sn : "", urgency: "low"
        });
      });
    } catch (e) {}
    out.forEach(function (x) { x.urgency = autoUrgency(x.due); });
    return out;
  }

  function openGeneralTickets() {
    var out = [];
    try {
      var d = QBR.users.myData();
      (d.tickets || []).forEach(function (t) {
        if (/^(resolved|completed|closed)$/i.test(String(t.status || "Open"))) return;
        out.push({
          id: "auto:gen:" + t.tno, auto: true, kind: "genticket",
          title: t.tno + " — " + t.title, product: "tickets",
          due: t.due || null, notes: t.requester ? "From " + t.requester : "",
          urgency: autoUrgency(t.due)
        });
      });
    } catch (e) {}
    return out;
  }

  function unauditedSchools() {
    var out = [];
    try {
      var a = (typeof QBR !== "undefined" && QBR._audit) || null;
      if (!a || !a.months) return out;
      var months = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
                    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
      var m = months[new Date().getMonth()];
      var ms = a.months[m];
      if (!ms || !ms.rows) return out;
      ms.rows.forEach(function (rec) {
        if (rec.exempt) return;
        var hasRisky = rec.risky != null && String(rec.risky).trim() !== "";
        if (hasRisky) return;
        out.push({
          id: "auto:school:" + m + ":" + String(rec.school || "").toLowerCase(),
          auto: true, kind: "school",
          title: (rec.school || "Unknown school") + " — " + m.charAt(0) + m.slice(1).toLowerCase() + " audit missing",
          product: "schools", due: null, notes: rec.org || "", urgency: "medium"
        });
      });
    } catch (e) {}
    return out;
  }

  function allTasks() {
    return manualTasks().concat(openLenovoTickets(), openGeneralTickets(), unauditedSchools());
  }

  /* --- filters ---------------------------------------------------------- */
  var _filters = { urgency: "all", deadline: "all", view: "tiles" };

  function passFilters(t) {
    if (_filters.urgency !== "all" && t.urgency !== _filters.urgency) return false;
    var dl = _filters.deadline, d = daysUntil(t.due);
    if (dl === "overdue" && !(d != null && d < 0)) return false;
    if (dl === "week" && !(d != null && d >= 0 && d <= 7)) return false;
    if (dl === "later" && !(d != null && d > 7)) return false;
    if (dl === "nodate" && d != null) return false;
    return true;
  }

  /* --- mutations (manual todos only; auto items resolve at the source) --- */
  function addTodo(f) {
    var d = QBR.users.myData();
    if (d._nouser) return { ok: false, error: "Pick or register a user first." };
    var t = {
      id: "t" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      title: String(f.title || "").trim(),
      product: MANUAL_PRODUCTS.indexOf(f.product) >= 0 ? f.product : "general",
      due: f.due || null, priority: f.priority || "auto",
      notes: String(f.notes || "").trim(), done: false,
      createdAt: new Date().toISOString()
    };
    if (!t.title) return { ok: false, error: "Give the task a title." };
    d.todos.push(t);
    QBR.users.saveMyData(d);
    return { ok: true, todo: t };
  }

  function completeTodo(manualId) {
    var d = QBR.users.myData();
    var i = (d.todos || []).findIndex(function (t) { return t.id === manualId; });
    if (i < 0) return false;
    var t = d.todos[i];
    d.history.unshift({ doneAt: new Date().toISOString(), kind: "todo", title: t.title, product: t.product });
    d.todos.splice(i, 1);
    QBR.users.saveMyData(d);
    return true;
  }

  function setPriority(manualId, p) {
    var d = QBR.users.myData();
    var t = (d.todos || []).filter(function (x) { return x.id === manualId; })[0];
    if (!t) return false;
    t.priority = ["auto", "high", "medium", "low"].indexOf(p) >= 0 ? p : "auto";
    QBR.users.saveMyData(d);
    return true;
  }

  /* --- badge: open high-urgency / overdue count for the nav -------------- */
  function badgeCount() {
    try {
      return allTasks().filter(function (t) {
        return t.urgency === "high" || daysUntil(t.due) < 0;
      }).length;
    } catch (e) { return 0; }
  }

  /* --- render ------------------------------------------------------------ */
  var URG = { high: ["●", "HIGH", "danger"], medium: ["▲", "MED", "warning"], low: ["○", "LOW", "secondary"] };

  function dueLabel(t) {
    if (!t.due) return '<span class="text-muted">No date</span>';
    var d = daysUntil(t.due);
    var cls = d < 0 ? "text-danger fw-bold" : (d <= 7 ? "text-warning" : "text-muted");
    var extra = d < 0 ? " — OVERDUE" : "";
    return '<span class="' + cls + '">' + esc(t.due) + extra + "</span>";
  }

  function taskRow(t) {
    var u = URG[t.urgency] || URG.low;
    var doneBtn = t.auto
      ? '<span class="badge bg-info" title="Resolves automatically when the underlying work is done">auto</span>'
      : '<button class="btn btn-sm btn-outline-success" data-todo-act="done" data-id="' + esc(t.id) + '">✓ Done</button>';
    var prioCtl = t.auto ? "" :
      '<select class="form-select form-select-sm d-inline-block w-auto" data-todo-prio="' + esc(t.id) + '" title="Urgency override">' +
      ["auto", "high", "medium", "low"].map(function (p) {
        return '<option value="' + p + '"' + ((t.priority || "auto") === p ? " selected" : "") + ">" +
          (p === "auto" ? "Auto" : p.charAt(0).toUpperCase() + p.slice(1)) + "</option>";
      }).join("") + "</select>";
    return '<div class="todo-task d-flex align-items-center gap-2 py-1 border-bottom">' +
      '<span class="badge bg-' + u[2] + '">' + u[0] + " " + u[1] + "</span>" +
      '<div class="flex-grow-1"><div>' + esc(t.title) + "</div>" +
      (t.notes ? '<div class="small text-muted">' + esc(t.notes) + "</div>" : "") + "</div>" +
      "<div class=\"small\">" + dueLabel(t) + "</div>" + prioCtl + doneBtn + "</div>";
  }

  function render() {
    var host = $("dash-todo");
    if (!host) return;
    var u = QBR.users.ensure();
    if (!u) { renderPicker(host); return; }

    var tasks = allTasks().filter(passFilters);
    var tiles = PRODUCTS.map(function (p) {
      var items = tasks.filter(function (t) { return t.product === p.id; });
      var body = items.length
        ? items.map(taskRow).join("")
        : '<div class="text-muted small px-2 py-1">Nothing here.</div>';
      return '<div class="col-md-6 col-lg-4 mb-3"><div class="card-box h-100">' +
        '<h6 class="mb-2">' + esc(p.label) + ' <span class="badge bg-secondary">' + items.length + "</span></h6>" +
        body + "</div></div>";
    }).join("");

    host.innerHTML =
      '<div class="d-flex flex-wrap align-items-center gap-2 mb-3">' +
        "<h4 class=\"mb-0\">My To-Do</h4>" +
        '<span class="badge bg-primary">' + esc(u.name) + "</span>" +
        '<button class="btn btn-sm btn-outline-secondary" data-todo-act="switch">Switch user</button>' +
        '<button class="btn btn-sm btn-primary" data-todo-act="add">+ Add task</button>' +
        '<button class="btn btn-sm btn-outline-secondary" data-todo-act="history">History</button>' +
        '<button class="btn btn-sm btn-outline-secondary" data-todo-act="export">Export my file</button>' +
        '<span class="ms-auto d-flex gap-2">' +
          '<select id="todo-f-urg" class="form-select form-select-sm w-auto" title="Filter by urgency">' +
            [["all", "Urgency: all"], ["high", "High"], ["medium", "Medium"], ["low", "Low"]].map(function (o) {
              return '<option value="' + o[0] + '"' + (_filters.urgency === o[0] ? " selected" : "") + ">" + o[1] + "</option>";
            }).join("") + "</select>" +
          '<select id="todo-f-due" class="form-select form-select-sm w-auto" title="Filter by deadline">' +
            [["all", "Deadline: all"], ["overdue", "Overdue"], ["week", "Due this week"], ["later", "Due later"], ["nodate", "No date"]].map(function (o) {
              return '<option value="' + o[0] + '"' + (_filters.deadline === o[0] ? " selected" : "") + ">" + o[1] + "</option>";
            }).join("") + "</select>" +
        "</span></div>" +
      '<div class="row">' + tiles + "</div>" +
      '<div id="todo-modal-host"></div>';

    bind(host);
  }

  function renderPicker(host) {
    var users = QBR.users.list();
    var opts = users.map(function (x) {
      return '<button class="btn btn-outline-primary m-1" data-todo-user="' + esc(x.id) + '">' + esc(x.name) + "</button>";
    }).join("");
    host.innerHTML = '<div class="card-box text-center py-5"><h4>Who\'s using the dashboard?</h4>' +
      '<p class="text-muted">Pick your profile — you\'ll only see your own tasks.</p>' +
      "<div>" + (opts || '<span class="text-muted">No users yet — register below.</span>') + "</div>" +
      '<div class="mt-3 d-flex justify-content-center gap-2">' +
      '<input id="todo-reg-name" class="form-control w-auto" placeholder="Your name" maxlength="31">' +
      '<button class="btn btn-primary" id="todo-reg-go">Register</button></div>' +
      '<div id="todo-reg-msg" class="small text-danger mt-2"></div></div>';
    host.querySelectorAll("[data-todo-user]").forEach(function (b) {
      b.addEventListener("click", function () { QBR.users.switch(b.dataset.todoUser); render(); });
    });
    $("todo-reg-go").addEventListener("click", function () {
      var r = QBR.users.register($("todo-reg-name").value);
      if (!r.ok) { $("todo-reg-msg").textContent = r.error; return; }
      render();
    });
  }

  function renderHistory() {
    var host = $("dash-todo");
    var d = QBR.users.myData();
    var rows = (d.history || []).map(function (h) {
      return '<div class="py-1 border-bottom"><span class="text-success">✓</span> ' + esc(h.title || "") +
        ' <span class="small text-muted">' + esc((h.doneAt || "").slice(0, 10)) + " · " + esc(h.product || "") + "</span></div>";
    }).join("") || '<div class="text-muted">Nothing completed yet.</div>';
    host.innerHTML = '<div class="d-flex align-items-center gap-2 mb-3"><h4 class="mb-0">History</h4>' +
      '<button class="btn btn-sm btn-outline-secondary" data-todo-act="back">← Back to To-Do</button></div>' +
      '<div class="card-box">' + rows + "</div>";
    host.querySelector('[data-todo-act="back"]').addEventListener("click", render);
  }

  function addDialog() {
    var host = $("todo-modal-host");
    host.innerHTML = '<div class="modal-overlay open"><div class="scan-dialog card-box" role="dialog" aria-modal="true">' +
      "<h6>Add task</h6>" +
      '<label class="form-label small">Title *</label><input id="todo-a-title" class="form-control form-control-sm mb-2">' +
      '<div class="row g-2 mb-2"><div class="col-6"><label class="form-label small">Product</label>' +
      '<select id="todo-a-prod" class="form-select form-select-sm">' +
      MANUAL_PRODUCTS.map(function (p) { return '<option value="' + p + '">' + p.charAt(0).toUpperCase() + p.slice(1) + "</option>"; }).join("") +
      '</select></div><div class="col-6"><label class="form-label small">Due date</label>' +
      '<input id="todo-a-due" type="date" class="form-control form-control-sm"></div></div>' +
      '<label class="form-label small">Priority</label>' +
      '<select id="todo-a-prio" class="form-select form-select-sm mb-2">' +
      '<option value="auto">Auto (from due date)</option><option value="high">High</option>' +
      '<option value="medium">Medium</option><option value="low">Low</option></select>' +
      '<label class="form-label small">Notes</label><textarea id="todo-a-notes" class="form-control form-control-sm mb-3" rows="2"></textarea>' +
      '<div class="d-flex gap-2"><button class="btn btn-sm btn-primary" id="todo-a-save">Save</button>' +
      '<button class="btn btn-sm btn-outline-secondary" id="todo-a-cancel">Cancel</button></div>' +
      '<div id="todo-a-msg" class="small text-danger mt-2"></div></div></div>';
    $("todo-a-cancel").addEventListener("click", function () { host.innerHTML = ""; });
    host.querySelector(".modal-overlay").addEventListener("click", function (e) {
      if (e.target.classList.contains("modal-overlay")) host.innerHTML = "";
    });
    $("todo-a-save").addEventListener("click", function () {
      var r = addTodo({
        title: $("todo-a-title").value, product: $("todo-a-prod").value,
        due: $("todo-a-due").value || null, priority: $("todo-a-prio").value,
        notes: $("todo-a-notes").value
      });
      if (!r.ok) { $("todo-a-msg").textContent = r.error; return; }
      render();
    });
  }

  function bind(host) {
    var fu = $("todo-f-urg"), fd = $("todo-f-due");
    if (fu) fu.addEventListener("change", function () { _filters.urgency = fu.value; render(); });
    if (fd) fd.addEventListener("change", function () { _filters.deadline = fd.value; render(); });
    host.querySelectorAll("[data-todo-act]").forEach(function (b) {
      b.addEventListener("click", function () {
        var act = b.dataset.todoAct, id = b.dataset.id;
        if (act === "done" && id) { completeTodo(id.replace(/^m:/, "")); render(); }
        else if (act === "add") addDialog();
        else if (act === "history") renderHistory();
        else if (act === "switch") { renderPicker(host); }
        else if (act === "export") {
          var r = QBR.users.exportWorkbook();
          alert(r.ok ? "Saved " + r.file : r.error);
        }
      });
    });
    host.querySelectorAll("[data-todo-prio]").forEach(function (sel) {
      sel.addEventListener("change", function () {
        setPriority(sel.dataset.todoPrio.replace(/^m:/, ""), sel.value);
        render();
      });
    });
  }

  /* Re-render when the current user or their data changes (user switch,
   * register, a general ticket added/closed) so the view is never stale. */
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
  QBR.todo = {
    render: render,
    all: allTasks,
    urgencyOf: urgencyOf,
    autoUrgency: autoUrgency,
    badgeCount: badgeCount,
    add: addTodo,
    complete: completeTodo,
    setPriority: setPriority
  };
})();
