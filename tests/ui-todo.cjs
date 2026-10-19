// 2026-10-13 My To-Do tighten pass (contributor) + review fixes. Headless Chromium, file://.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-todo.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ viewport: { width: 1500, height: 900 } })).newPage();
  const errs = []; p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(300);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx"), path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx")]); await p.waitForTimeout(2500);
  await p.evaluate(() => {
    if (!QBR.users.ensure()) QBR.users.register("Tester");
    for (let i = 0; i < 12; i++) QBR.todo.add({ title: "G task " + i, product: "general" });
    QBR.todo.add({ title: "MS renewal", product: "microsoft", due: "2020-01-01" });
    const M = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"][new Date().getMonth()];
    QBR._audit = QBR._audit || { months: {} }; QBR._audit.months = QBR._audit.months || {};
    QBR._audit.months[M] = QBR._audit.months[M] || { rows: [] };
    QBR._audit.months[M].rows.push({ school: "Demo School", org: "DEMO" });
    QBR._audit.months[M].rows.push({ school: "Twin College Inc", org: "DEMO" }, { school: "TWIN COLLEGE INC.", org: "DEMO" });
    QBR._audit.months[M].rows.push({ school: "Done Academy", org: "DEMO", risky: "" }, { school: "DONE ACADEMY", org: "DEMO", risky: 3 });
    QBR.todo.add({ title: "Lenovo RMA follow-up", product: "lenovo", priority: "high" });
    goToTab("dash-todo");
  });
  await p.waitForTimeout(300);
  const st = await p.evaluate(() => ({
    order: [...document.querySelectorAll("[id^=todo-tile-]")].map(e => e.id.replace("todo-tile-", "")),
    chips: [...document.querySelectorAll("[data-todo-jump]")].length,
    gRows: document.querySelectorAll("#todo-tile-general .todo-task").length,
    more: (document.querySelector('[data-todo-more="general"]') || {}).textContent || "",
    tickets: document.querySelectorAll("#todo-tile-tickets a[href^=\"#ticket/\"]").length,
    noDate: document.getElementById("dash-todo").innerHTML.includes(">No date</span>"),
  }));
  ok(st.order.join(",") === "general,microsoft,lenovo,tickets,schools", "manual tiles first  " + st.order);
  ok(st.chips === 5, "5 jump chips");
  ok(st.gRows === 10 && /Show all 12/.test(st.more), "General tile capped at 10 + 'Show all 12'");
  ok(!st.noDate, "no 'No date' row noise");
  await p.click('[data-todo-more="general"]'); await p.waitForTimeout(150);
  ok(await p.evaluate(() => document.querySelectorAll("#todo-tile-general .todo-task").length === 12), "Show all expands");
  ok(st.tickets > 0, "Lenovo ticket rows link to Ticket 360 (" + st.tickets + " links)");

  console.log("== ticket deep link, twice");
  const href = await p.evaluate(() => document.querySelector("#todo-tile-tickets a[href^=\"#ticket/\"]").getAttribute("href"));
  await p.click(`#todo-tile-tickets a[href="${href}"]`); await p.waitForTimeout(300);
  ok(await p.evaluate(() => APP.activeTab === "dash-ticket360"), "first click opens Ticket 360  " + href);
  await p.evaluate(() => goToTab("dash-todo")); await p.waitForTimeout(200);
  await p.click(`#todo-tile-tickets a[href="${href}"]`); await p.waitForTimeout(300);
  ok(await p.evaluate(() => APP.activeTab === "dash-ticket360"), "same ticket again after coming back still opens (was a dead link)");

  console.log("== serial link");
  await p.evaluate(() => goToTab("dash-todo")); await p.waitForTimeout(200);
  const sref = await p.evaluate(() => { const a = [...document.querySelectorAll("#todo-tile-tickets a[href^=\"#asset/\"]")].find(x => /#asset\//.test(x.getAttribute("href"))); return a && a.getAttribute("href"); });
  if (sref) { await p.click(`#todo-tile-tickets a[href="${sref}"]`); await p.waitForTimeout(300); }
  ok(sref && await p.evaluate(() => APP.activeTab === "dash-asset360" && !/not in (the )?inventory/i.test(document.getElementById("dash-asset360").textContent)), "serial opens Asset 360 for a known unit  " + sref);

  console.log("== school + manual edit");
  await p.evaluate(() => goToTab("dash-todo")); await p.waitForTimeout(200);
  await p.click('#todo-tile-schools a[data-todo-goto="audit"]'); await p.waitForTimeout(300);
  ok(await p.evaluate(() => APP.activeTab === "dash-audit"), "unaudited school opens the audit editor");
  await p.evaluate(() => goToTab("dash-todo")); await p.waitForTimeout(200);
  await p.click('#todo-tile-microsoft a[data-todo-goto="edit"]'); await p.waitForTimeout(150);
  ok(await p.inputValue("#todo-a-title") === "MS renewal" && await p.inputValue("#todo-a-due") === "2020-01-01", "manual task opens Edit, prefilled");
  await p.fill("#todo-a-title", ""); await p.click("#todo-a-save"); await p.waitForTimeout(100);
  ok(/title/i.test(await p.textContent("#todo-a-msg")), "blank title on edit is refused with a message");
  await p.fill("#todo-a-title", "MS renewal (edited)"); await p.click("#todo-a-save"); await p.waitForTimeout(150);
  ok(await p.evaluate(() => /MS renewal \(edited\)/.test(document.getElementById("todo-tile-microsoft").textContent)), "edit saved");
  console.log("== Phase A/B: urgency signals + colour coding");
  const v = await p.evaluate(() => {
    const rows = [...document.querySelectorAll("#dash-todo .todo-row")];
    const lv = r => (r.className.match(/lv-(\w+)/) || [])[1];
    const tix = [...document.querySelectorAll("#todo-tile-tickets .todo-row")];
    const sch = [...document.querySelectorAll("#todo-tile-schools .todo-row")].map(r => r.querySelector(".todo-title").textContent);
    const chip = r => (r.querySelector(".todo-chip") || {}).textContent || "";
    const ms = [...document.querySelectorAll("#todo-tile-microsoft .todo-row")][0];
    const bar = getComputedStyle(document.querySelector(".filter-bar")).display;
    const rank = { overdue: 0, high: 1, medium: 2, low: 3 };
    const tl = tix.map(lv);
    return {
      tixLevels: [...new Set(tl)], tixSorted: tl.every((x, i) => !i || rank[tl[i - 1]] <= rank[x]),
      tixChip: chip(tix[0]), msLv: lv(ms), msChip: chip(ms), msBar: getComputedStyle(ms).borderLeftColor, msBg: getComputedStyle(ms).backgroundColor,
      twin: sch.filter(t => /twin college/i.test(t)).length, done: sch.some(t => /done academy/i.test(t)),
      schoolChip: chip([...document.querySelectorAll("#todo-tile-schools .todo-row")][0]),
      lenovoHigh: lv(document.querySelector("#todo-tile-lenovo .todo-row")),
      titlesUnderlined: getComputedStyle(document.querySelector("#dash-todo a.todo-title")).textDecorationLine,
      legend: document.querySelectorAll("#dash-todo .todo-legend span[class^=lv-]").length,
      filterBar: bar, badge: QBR.todo.badgeCount(),
      wrapDate: [...document.querySelectorAll("#dash-todo .todo-chip")].every(c => c.getClientRects().length === 1),
    };
  });
  ok(v.tixLevels.length > 1 && v.tixSorted, "tickets no longer all LOW, sorted worst first  " + JSON.stringify(v.tixLevels));
  ok(/open \d+ days/.test(v.tixChip), "ticket chip shows its age  " + v.tixChip);
  ok(v.msLv === "overdue" && /overdue/.test(v.msChip) && v.msBg !== "rgba(0, 0, 0, 0)", "overdue task: red bar + tint + 'N days overdue'  " + v.msChip);
  ok(v.twin === 1 && !v.done, "audit: same school under two spellings listed once; audited-in-either counts as done");
  ok(/in \d+ days|due today|overdue|[A-Z][a-z]{2} \d+/.test(v.schoolChip), "school rows carry the audit deadline  " + v.schoolChip);
  ok(v.lenovoHigh === "high", "manual High task coloured high");
  ok(v.titlesUnderlined === "none", "titles are plain text (underline only on hover)");
  ok(v.legend === 4, "legend: Overdue / High / Due soon / Planned");
  ok(v.filterBar === "none", "shared filter bar hidden on My To-Do");
  ok(v.wrapDate, "time chips stay on one line");
  await p.click("#todo-tile-microsoft .todo-row .todo-sub, #todo-tile-microsoft .todo-row .todo-main"); await p.waitForTimeout(150);
  ok(await p.evaluate(() => !!document.getElementById("todo-a-title")), "clicking the row (not the title) opens the item");
  await p.evaluate(() => { const c = document.getElementById("todo-a-cancel"); if (c) c.click(); });
  await p.click("#todo-tile-lenovo input.todo-check"); await p.waitForTimeout(150);
  ok(await p.evaluate(() => !document.querySelector("#todo-tile-lenovo .todo-row") && QBR.users.myData().history.some(h => /RMA/.test(h.title))), "checkbox marks done → History");
  await p.evaluate(() => goToTab("dash-overview")); await p.waitForTimeout(150);
  ok(await p.evaluate(() => getComputedStyle(document.querySelector(".filter-bar")).display !== "none"), "filter bar back on other tabs");
  await p.evaluate(() => goToTab("dash-todo")); await p.waitForTimeout(150);
  await p.screenshot({ path: path.join(__dirname, "..", "todo_light.png"), fullPage: true });
  await p.evaluate(() => { document.body.setAttribute("data-theme", "dark"); QBR.todo.render(); }); await p.waitForTimeout(150);
  await p.screenshot({ path: path.join(__dirname, "..", "todo_dark.png"), fullPage: true });
  ok(errs.length === 0, "no page errors  " + JSON.stringify(errs.slice(0, 3)));
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
