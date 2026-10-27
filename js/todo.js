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
    { id: "general", label: "General" },
    { id: "microsoft", label: "Microsoft" },
    { id: "lenovo", label: "Lenovo" },
    { id: "tickets", label: "Pending tickets" },
    { id: "schools", label: "Unaudited schools" }
  ];
  /* Tighten pass: capped tiles (TILE_LIMIT rows + "Show all"), jump chips,
   * clickable rows (deep links), manual tiles first. */
  var TILE_LIMIT = 10;
  /* Urgency rules (2026-10-13, ops lead): an open Lenovo ticket is HIGH after 30 days, MED after 14
   * (or HIGH when the sheet's Priority says High/Critical/Urgent); the monthly risky sign-in audit
   * is due on AUDIT_DUE_DAY, so a missing audit is LOW before the last week, MED in it, HIGH after. */
  var TICKET_HIGH_DAYS = 30, TICKET_MED_DAYS = 14, AUDIT_DUE_DAY = 25;
  var RANK = { high: 1, medium: 2, low: 3 };
  var _expanded = {};
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
        var age = ticketAgeDays(t);
        out.push({
          id: "auto:lenovo:" + (t.tno || t.key || Math.random().toString(36).slice(2)),
          auto: true, kind: "ticket",
          title: (t.issue || t.subject || "Lenovo ticket"),
          product: "tickets", due: t.followUp || t.due || null, age: age,
          client: t.client || "", sheetPriority: t.priority || "",
          tno: t.tno || null, sn: t.sn || null, urgency: "low"
        });
      });
    } catch (e) {}
    out.forEach(function (x) { x.urgency = ticketUrgency(x); });
    return out;
  }

  /* Days a Lenovo ticket has been open: the ticket number carries its date (<SERIAL>-YYYYMMDD);
   * else the date it was reported to EdTech / Lenovo. null when unknown. */
  function ticketAgeDays(t) {
    /* Only "<SERIAL>-YYYYMMDD[-n]" with a real calendar date counts (a numeric Lenovo case
     * number like "PF12345678" must not read as the year 1234). */
    var d = null, m = /-(\d{4})(\d{2})(\d{2})(?:-\d+)?$/.exec(String(t.tno || ""));
    if (m && +m[1] >= 2000 && +m[1] <= 2100) {
      d = new Date(+m[1], +m[2] - 1, +m[3]);
      if (d.getMonth() !== +m[2] - 1 || d.getDate() !== +m[3]) d = null;
    }
    if ((!d || isNaN(d)) && t.repEdtech instanceof Date) d = t.repEdtech;
    if ((!d || isNaN(d)) && t.repLenovo instanceof Date) d = t.repLenovo;
    if (!d || isNaN(d)) return null;
    var now = new Date(); now.setHours(0, 0, 0, 0);
    var n = Math.round((now - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 86400000);
    return n >= 0 ? n : 0;
  }
  function ticketUrgency(x) {
    var u = [];
    if (/^(high|critical|urgent|p1)/i.test(String(x.sheetPriority || "").trim())) u.push("high");
    if (x.age != null) u.push(x.age > TICKET_HIGH_DAYS ? "high" : x.age >= TICKET_MED_DAYS ? "medium" : "low");
    if (x.due) u.push(autoUrgency(x.due));
    if (!u.length) return "low";
    return u.sort(function (a, b) { return RANK[a] - RANK[b]; })[0];
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
          tno: t.tno, urgency: autoUrgency(t.due)
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
      var now = new Date(), m = months[now.getMonth()];
      var ms = a.months[m];
      if (!ms || !ms.rows) return out;
      var due = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0") + "-" + String(AUDIT_DUE_DAY).padStart(2, "0");
      /* One entry per school: the audit sheet can hold the same school under two spellings
       * ("Dr Sample's Colleges Inc" / "DR. SAMPLE'S COLLEGES INC."). Same join key as the rest of
       * the dashboard; audited if ANY of its rows has a value; display = master-list name. */
      var keyOf = typeof schoolKey === "function" ? schoolKey : function (n) { return String(n || "").toUpperCase().replace(/[^A-Z0-9 ]/g, " ").replace(/\s+/g, " ").trim(); };
      var master = (typeof APP !== "undefined" && APP.model && APP.model.master) || null;
      var groups = {}, order = [];
      ms.rows.forEach(function (rec) {
        var k = keyOf(rec.school);
        if (!k) return;
        if (!groups[k]) { groups[k] = { recs: [], done: false, exempt: false }; order.push(k); }
        var g = groups[k];
        g.recs.push(rec);
        if (rec.exempt) g.exempt = true;
        if (rec.risky != null && String(rec.risky).trim() !== "") g.done = true;
      });
      order.forEach(function (k) {
        var g = groups[k], rec = g.recs[0];
        if (g.done || g.exempt) return;
        var mm = master && master.get ? master.get(k) : null;
        var name = (mm && mm.name) || rec.school || "Unknown school";
        out.push({
          id: "auto:school:" + m + ":" + k.toLowerCase(),
          auto: true, kind: "school",
          title: name + " — " + m.charAt(0) + m.slice(1).toLowerCase() + " audit missing",
          product: "schools", due: due, notes: rec.org || "", urgency: autoUrgency(due),
          month: m, school: rec.school || "", dupes: g.recs.length
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

  function updateTodo(manualId, f) {
    var d = QBR.users.myData();
    var t = (d.todos || []).filter(function (x) { return x.id === manualId; })[0];
    if (!t) return false;
    if (String(f.title || "").trim()) t.title = String(f.title).trim();
    if (MANUAL_PRODUCTS.indexOf(f.product) >= 0) t.product = f.product;
    t.due = f.due || null;
    t.priority = ["auto", "high", "medium", "low"].indexOf(f.priority) >= 0 ? f.priority : "auto";
    t.notes = String(f.notes || "").trim();
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
        return isOverdue(t) || (!t.auto && t.urgency === "high");
      }).length;
    } catch (e) { return 0; }
  }

  /* --- render ------------------------------------------------------------ */
  var URG = { high: ["●", "HIGH", "danger"], medium: ["▲", "MED", "warning"], low: ["○", "LOW", "secondary"] };

  function isOverdue(t) { var d = daysUntil(t.due); return d != null && d < 0; }
  /* Visual level: overdue > high > medium > low (colour + word, never colour alone). */
  function levelOf(t) { return isOverdue(t) ? "overdue" : (t.urgency || "low"); }
  var LEVEL = {
    overdue: { word: "Overdue", icon: "⚠" },
    high: { word: "High", icon: "●" },
    medium: { word: "Due soon", icon: "▲" },
    low: { word: "Planned", icon: "○" }
  };
  function sortTasks(list) {
    var r = { overdue: 0, high: 1, medium: 2, low: 3 };
    return list.slice().sort(function (a, b) {
      var x = r[levelOf(a)] - r[levelOf(b)];
      if (x) return x;
      var da = daysUntil(a.due), db = daysUntil(b.due);
      if (da != null || db != null) { if (da == null) return 1; if (db == null) return -1; if (da !== db) return da - db; }
      return (b.age || 0) - (a.age || 0);
    });
  }
  var MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  /* One-line time chip: "5 days overdue", "today", "in 3 days", "Oct 16"; tickets: "open 207 days". */
  function timeChip(t) {
    var lv = levelOf(t), d = daysUntil(t.due), txt = "";
    if (d != null) {
      if (d < -60) { var q = String(t.due).split("-"); txt = "overdue since " + MON[(+q[1] || 1) - 1] + " " + q[0]; }
      else if (d < 0) txt = (-d) + (d === -1 ? " day" : " days") + " overdue";
      else if (d === 0) txt = "due today";
      else if (d <= 7) txt = "in " + d + (d === 1 ? " day" : " days");
      else { var p = String(t.due).split("-"); txt = MON[(+p[1] || 1) - 1] + " " + (+p[2] || ""); }
    }
    if (t.kind === "ticket" && t.age != null) txt = (txt ? txt + " · " : "") + "open " + t.age + (t.age === 1 ? " day" : " days");
    if (!txt) return "";
    return '<span class="todo-chip lv-' + lv + '" title="' + esc(t.due || "") + '">' + esc(txt) + "</span>";
  }

  /* Title opens the item (row click does the same): lenovo ticket -> #ticket/<tno> (Ticket 360),
   * general ticket -> General tickets tab, school -> audit editor (month preselected),
   * manual task -> edit dialog. Plain text look; underline on hover/focus only. */
  function titleLink(t) {
    var label = esc(t.title);
    if (t.kind === "ticket" && t.tno) {
      return '<a class="todo-title" href="#ticket/' + esc(encodeURIComponent(t.tno)) + '" data-todo-hash title="Open in Ticket 360">' + label + "</a>";
    }
    if (t.kind === "genticket") {
      return '<a class="todo-title" href="#" data-todo-goto="gtickets" title="Open in General tickets">' + label + "</a>";
    }
    if (t.kind === "school") {
      return '<a class="todo-title" href="#" data-todo-goto="audit" data-month="' + esc(t.month || "") + '" title="Open in audit editor">' + label + "</a>";
    }
    if (!t.auto) {
      return '<a class="todo-title" href="#" data-todo-goto="edit" data-id="' + esc(t.id) + '" title="Edit task">' + label + "</a>";
    }
    return '<span class="todo-title">' + label + "</span>";
  }

  function subLine(t) {
    var bits = [];
    if (t.kind === "ticket") {
      if (t.tno) bits.push(esc(t.tno));
      if (t.sn) {
        /* Asset 360 routes on the normalized serial key (upper-case, no invisible chars). */
        var akey = (QBR.invSerialKey ? QBR.invSerialKey(t.sn) : String(t.sn).toUpperCase().trim());
        bits.push('Serial <a href="#asset/' + esc(encodeURIComponent(akey)) + '" data-todo-hash title="Open in Asset 360">' + esc(t.sn) + "</a>");
      }
      if (t.client) bits.push(esc(t.client));
    } else if (t.notes) bits.push(esc(t.notes));
    if (t.auto) bits.push('<span class="todo-auto" title="Clears automatically when the work is done at the source">auto</span>');
    return bits.length ? '<span class="todo-sub">' + bits.join(" · ") + "</span>" : "";
  }

  function taskRow(t) {
    var lv = levelOf(t), L = LEVEL[lv];
    var check = t.auto ? '<span class="todo-check-ph" aria-hidden="true"></span>'
      : '<input type="checkbox" class="form-check-input todo-check" data-todo-act="done" data-id="' + esc(t.id) + '" aria-label="Mark this task done">';
    var prioCtl = t.auto ? "" :
      '<select class="form-select form-select-sm todo-prio" data-todo-prio="' + esc(t.id) + '" aria-label="Urgency override">' +
      ["auto", "high", "medium", "low"].map(function (p) {
        return '<option value="' + p + '"' + ((t.priority || "auto") === p ? " selected" : "") + ">" +
          (p === "auto" ? "Auto" : p.charAt(0).toUpperCase() + p.slice(1)) + "</option>";
      }).join("") + "</select>";
    return '<div class="todo-task todo-row lv-' + lv + '">' + check +
      '<div class="todo-main"><span class="visually-hidden">' + L.word + ": </span>" + titleLink(t) +
      '<div class="todo-meta">' + timeChip(t) + subLine(t) + "</div></div>" +
      (prioCtl ? '<span class="todo-ctl">' + prioCtl + "</span>" : "") + "</div>";
  }

  function render() {
    var host = $("dash-todo");
    if (!host) return;
    var u = QBR.users.ensure();
    if (!u) { renderPicker(host); return; }

    var tasks = allTasks().filter(passFilters);
    var counts = {};
    PRODUCTS.forEach(function (p) {
      counts[p.id] = tasks.filter(function (t) { return t.product === p.id; }).length;
    });
    var chips = PRODUCTS.map(function (p) {
      return '<button class="btn btn-sm btn-outline-secondary" data-todo-jump="' + p.id + '">' +
        esc(p.label) + ' <span class="badge bg-secondary">' + counts[p.id] + "</span></button>";
    }).join("");
    var tiles = PRODUCTS.map(function (p) {
      var items = sortTasks(tasks.filter(function (t) { return t.product === p.id; }));
      var expanded = !!_expanded[p.id];
      var shown = expanded ? items : items.slice(0, TILE_LIMIT);
      var body = shown.length
        ? shown.map(taskRow).join("")
        : '<div class="todo-empty">Nothing here.</div>';
      var more = "";
      if (items.length > TILE_LIMIT) {
        more = '<button class="btn btn-sm btn-link w-100" data-todo-more="' + p.id + '">' +
          (expanded ? "Show less ↑" : "Show all " + items.length + " ↓") + "</button>";
      }
      var manual = MANUAL_PRODUCTS.indexOf(p.id) >= 0;
      var hot = items.filter(function (t) { var l = levelOf(t); return l === "overdue" || l === "high"; }).length;
      return '<div class="' + (manual ? "col-md-6 col-lg-4" : "col-xl-6") + ' mb-3" id="todo-tile-' + p.id + '" style="scroll-margin-top:72px">' +
        '<div class="card-box todo-tile todo-p-' + p.id + '">' +
        '<h6 class="mb-2">' + esc(p.label) + ' <span class="badge bg-secondary">' + items.length + "</span>" +
        (hot ? ' <span class="todo-hot">' + hot + " urgent</span>" : "") + "</h6>" +
        '<div class="todo-tile-body" style="max-height:520px;overflow-y:auto;">' + body + "</div>" + more +
        "</div></div>";
    }).join("");

    host.innerHTML =
      '<div class="d-flex flex-wrap align-items-center gap-2 mb-3">' +
        "<h4 class=\"mb-0\">My To-Do</h4>" +
        '<span class="badge bg-primary">' + esc(u.name) + "</span>" +
        '<button class="btn btn-sm btn-outline-secondary" data-todo-act="switch">Switch user</button>' +
        '<button class="btn btn-sm btn-primary" data-todo-act="add">+ Add task</button>' +
        '<button class="btn btn-sm btn-outline-secondary" data-todo-act="history">History</button>' +
        '<button class="btn btn-sm btn-outline-secondary" data-todo-act="export">Export my file</button>' +
        '<button class="btn btn-sm btn-outline-secondary" data-todo-act="import" title="Re-open a file saved with Export my file (merges into this list)">Open saved file…</button>' +
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
      '<div class="d-flex flex-wrap gap-2 mb-2">' + chips + "</div>" +
      '<div class="todo-legend mb-3" aria-hidden="true">' + ["overdue", "high", "medium", "low"].map(function (k) {
        return '<span class="lv-' + k + '"><i></i>' + LEVEL[k].word + "</span>";
      }).join("") + '<span class="todo-legend-note">Tickets: High after ' + TICKET_HIGH_DAYS + " days open, due soon after " + TICKET_MED_DAYS +
      " · Audit due on the " + AUDIT_DUE_DAY + "th</span></div>" +
      '<div class="row align-items-start">' + tiles + "</div>" +
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

  function addDialog(editId) {
    var host = $("todo-modal-host");
    var d = QBR.users.myData();
    var ex = editId ? (d.todos || []).filter(function (t) { return t.id === editId; })[0] : null;
    var ev = function (v) { return esc(v == null ? "" : v); };
    host.innerHTML = '<div class="modal-overlay open"><div class="scan-dialog card-box" role="dialog" aria-modal="true">' +
      "<h6>" + (ex ? "Edit task" : "Add task") + "</h6>" +
      '<label class="form-label small">Title *</label><input id="todo-a-title" class="form-control form-control-sm mb-2" value="' + ev(ex && ex.title) + '">' +
      '<div class="row g-2 mb-2"><div class="col-6"><label class="form-label small">Product</label>' +
      '<select id="todo-a-prod" class="form-select form-select-sm">' +
      MANUAL_PRODUCTS.map(function (p) {
        return '<option value="' + p + '"' + (ex && ex.product === p ? " selected" : "") + ">" +
          p.charAt(0).toUpperCase() + p.slice(1) + "</option>";
      }).join("") +
      '</select></div><div class="col-6"><label class="form-label small">Due date</label>' +
      '<input id="todo-a-due" type="date" class="form-control form-control-sm" value="' + ev(ex && ex.due) + '"></div></div>' +
      '<label class="form-label small">Priority</label>' +
      '<select id="todo-a-prio" class="form-select form-select-sm mb-2">' +
      [["auto", "Auto (from due date)"], ["high", "High"], ["medium", "Medium"], ["low", "Low"]].map(function (o) {
        return '<option value="' + o[0] + '"' + (ex && (ex.priority || "auto") === o[0] ? " selected" : "") + ">" + o[1] + "</option>";
      }).join("") + "</select>" +
      '<label class="form-label small">Notes</label><textarea id="todo-a-notes" class="form-control form-control-sm mb-3" rows="2">' +
      ev(ex && ex.notes) + "</textarea>" +
      '<div class="d-flex gap-2"><button class="btn btn-sm btn-primary" id="todo-a-save">Save</button>' +
      '<button class="btn btn-sm btn-outline-secondary" id="todo-a-cancel">Cancel</button></div>' +
      '<div id="todo-a-msg" class="small text-danger mt-2"></div></div></div>';
    $("todo-a-cancel").addEventListener("click", function () { host.innerHTML = ""; });
    host.querySelector(".modal-overlay").addEventListener("click", function (e) {
      if (e.target.classList.contains("modal-overlay")) host.innerHTML = "";
    });
    $("todo-a-save").addEventListener("click", function () {
      var f = {
        title: $("todo-a-title").value, product: $("todo-a-prod").value,
        due: $("todo-a-due").value || null, priority: $("todo-a-prio").value,
        notes: $("todo-a-notes").value
      };
      if (ex) {
        if (!String(f.title || "").trim()) { $("todo-a-msg").textContent = "Give the task a title."; return; }
        if (!updateTodo(ex.id, f)) { $("todo-a-msg").textContent = "Task not found."; return; }
      } else {
        var r = addTodo(f);
        if (!r.ok) { $("todo-a-msg").textContent = r.error; return; }
      }
      render();
    });
  }

  function gotoTabSafe(tabId) {
    if (typeof goToTab === "function") { goToTab(tabId); return; }
    var btn = document.querySelector('.tab-btn[data-tab="' + tabId + '"]');
    if (btn) btn.click();
  }

  function bind(host) {
    var fu = $("todo-f-urg"), fd = $("todo-f-due");
    if (fu) fu.addEventListener("change", function () { _filters.urgency = fu.value; render(); });
    if (fd) fd.addEventListener("change", function () { _filters.deadline = fd.value; render(); });
    host.querySelectorAll("[data-todo-jump]").forEach(function (b) {
      b.addEventListener("click", function () {
        var el = document.getElementById("todo-tile-" + b.dataset.todoJump);
        if (el && el.scrollIntoView) el.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    });
    host.querySelectorAll("[data-todo-more]").forEach(function (b) {
      b.addEventListener("click", function () {
        var k = b.dataset.todoMore;
        _expanded[k] = !_expanded[k];
        render();
      });
    });
    /* Ticket 360 / Asset 360 links: openTicket360/openAsset360 leave the hash on
     * #ticket/<tno>, so clicking the same row again after coming back would not
     * fire hashchange (dead link), and the 360 handlers skip a ticket they think
     * is already open. Clear that memo and re-apply the route ourselves. */
    host.querySelectorAll("a[data-todo-hash]").forEach(function (a) {
      a.addEventListener("click", function (e) {
        var h = a.getAttribute("href");
        try { QBR._invT360 = null; QBR._invA360 = null; } catch (err) {}
        if (location.hash === h && typeof invApplyHash === "function") { e.preventDefault(); invApplyHash(); }
      });
    });
    host.querySelectorAll("[data-todo-goto]").forEach(function (b) {
      b.addEventListener("click", function (e) {
        e.preventDefault();
        var g = b.dataset.todoGoto;
        if (g === "gtickets") { gotoTabSafe("dash-gtickets"); }
        else if (g === "audit") {
          try {
            var m = b.dataset.month;
            if (QBR._auditUI) { if (m) QBR._auditUI.month = m; }
            else QBR._auditUI = { month: m || null, editing: null, adding: false, pasteFor: null };
          } catch (err) {}
          gotoTabSafe("dash-audit");
        }
        else if (g === "edit" && b.dataset.id) { addDialog(b.dataset.id.replace(/^m:/, "")); }
      });
    });
    host.querySelectorAll("input.todo-check").forEach(function (c) {
      c.addEventListener("change", function () { if (c.checked) { completeTodo(c.dataset.id.replace(/^m:/, "")); render(); } });
    });
    /* Whole row opens the item, except clicks on its own controls/links. */
    host.querySelectorAll(".todo-row").forEach(function (row) {
      row.addEventListener("click", function (e) {
        if (e.target.closest("a, button, input, select, label")) return;
        var a = row.querySelector("a.todo-title"); if (a) a.click();
      });
    });
    host.querySelectorAll("[data-todo-act]:not(input)").forEach(function (b) {
      b.addEventListener("click", function () {
        var act = b.dataset.todoAct, id = b.dataset.id;
        if (act === "done" && id) { completeTodo(id.replace(/^m:/, "")); render(); }
        else if (act === "add") addDialog();
        else if (act === "history") renderHistory();
        else if (act === "switch") { renderPicker(host); }
        else if (act === "import") QBR.users.importDialog();
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
    update: updateTodo,
    complete: completeTodo,
    setPriority: setPriority
  };
})();
