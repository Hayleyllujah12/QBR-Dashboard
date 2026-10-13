/* General ticketing tests (Node). Run: node tools/test-gtickets.js */
"use strict";
const fs = require("fs"), vm = require("vm"), path = require("path");

const sb = {
  console, XLSX: undefined, APP: { model: {} },
  document: { getElementById: () => null, createElement: () => ({ style: {} }), querySelectorAll: () => [] },
  localStorage: { _d: {}, getItem(k) { return this._d[k] || null; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } },
};
sb.window = sb; sb.globalThis = sb;
vm.createContext(sb);
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "js", "users.js"), "utf8"), sb, { filename: "users.js" });
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "js", "todo.js"), "utf8"), sb, { filename: "todo.js" });
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "js", "gentickets.js"), "utf8"), sb, { filename: "gentickets.js" });

let pass = 0, fail = 0;
function t(name, cond) {
  if (cond) { pass++; }
  else { fail++; console.error("FAIL:", name); }
}
const eq = (a, b) => a === b;

vm.runInContext('QBR.users.register("Robin")', sb);
const stamp = (() => { const d = new Date(); return d.getFullYear() +
  String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0"); })();
const todayLocal = stamp.slice(0,4) + "-" + stamp.slice(4,6) + "-" + stamp.slice(6,8); // local YYYY-MM-DD ("today")

/* numbering */
let r1 = vm.runInContext('QBR.gtickets.create({title:"Fix wifi",requester:"Ana"})', sb);
t("create ok", r1.ok === true);
t("numbering REQ-YYYYMMDD-01", eq(r1.ticket.tno, `REQ-${stamp}-01`));
let r2 = vm.runInContext('QBR.gtickets.create({title:"Install Office"})', sb);
t("sequence increments", eq(r2.ticket.tno, `REQ-${stamp}-02`));
t("default status Open", eq(r1.ticket.status, "Open"));
t("blank title rejected", vm.runInContext('QBR.gtickets.create({title:"  "})', sb).ok === false);

/* status flow */
t("setStatus ok", vm.runInContext(`QBR.gtickets.setStatus("${r1.ticket.tno}", "In Progress")`, sb) === true);
t("status applied", eq(vm.runInContext(`QBR.gtickets.find("${r1.ticket.tno}").status`, sb), "In Progress"));
t("bad status rejected", vm.runInContext(`QBR.gtickets.setStatus("${r1.ticket.tno}", "Nope")`, sb) === false);
t("isOpen true", vm.runInContext(`QBR.gtickets.isOpen(QBR.gtickets.find("${r1.ticket.tno}"))`, sb) === true);
vm.runInContext(`QBR.gtickets.setStatus("${r1.ticket.tno}", "Resolved")`, sb);
t("isOpen false when resolved", vm.runInContext(`QBR.gtickets.isOpen(QBR.gtickets.find("${r1.ticket.tno}"))`, sb) === false);

/* notes */
t("addNote ok", vm.runInContext(`QBR.gtickets.addNote("${r2.ticket.tno}", "called vendor")`, sb) === true);
t("note stored", eq(vm.runInContext(`QBR.gtickets.find("${r2.ticket.tno}").notes.length`, sb), 1));

/* fingerprint dedup */
const fp1 = vm.runInContext(
  'QBR.gtickets.fingerprint({date:"' + todayLocal + '",time:"09:00",title:"Fix wifi",requester:"Ana"})', sb);
const fp2 = vm.runInContext(
  'QBR.gtickets.fingerprint({date:"' + todayLocal + '",time:"09:00",title:"  fix WIFI ",requester:"ana"})', sb);
t("fingerprint case/space-insensitive", eq(fp1, fp2));
const fp3 = vm.runInContext(
  'QBR.gtickets.fingerprint({date:"' + todayLocal + '",time:"10:00",title:"Fix wifi",requester:"Ana"})', sb);
t("different time -> different fp", fp1 !== fp3);

/* import preview: existing ticket detected as dupe */
const prev = vm.runInContext(`QBR.gtickets.previewImport([
  {date:"${todayLocal}",time:"",title:"Install Office",requester:""},
  {date:"${todayLocal}",time:"11:00",title:"New printer",requester:"Bo"}
])`, sb);
t("existing flagged dupe", prev[0].dupe === true);
t("new one not dupe", prev[1].dupe === false);

/* confirm import */
const cr = vm.runInContext(`QBR.gtickets.confirmImport(${JSON.stringify(prev.filter(p => !p.dupe))})`, sb);
t("imported count", eq(cr.imported, 1));
t("ticket created from import",
  vm.runInContext("QBR.gtickets.list().some(t => t.title === 'New printer')", sb));
const prev2 = vm.runInContext(`QBR.gtickets.previewImport([
  {date:"${todayLocal}",time:"11:00",title:"New printer",requester:"Bo"}
])`, sb);
t("re-import detected via seen-set", prev2[0].dupe === true);

/* auto -> To-Do */
const auto = vm.runInContext("QBR.todo.all().filter(t => t.kind === 'genticket')", sb);
t("open general tickets in To-Do", auto.length === 2); // Install Office + New printer (Fix wifi resolved)
t("resolved excluded from To-Do",
  !vm.runInContext("QBR.todo.all().some(t => t.title.indexOf('Fix wifi') >= 0)", sb));

/* per-user isolation */
vm.runInContext('QBR.users.register("Ana")', sb);
t("ana sees no tickets", eq(vm.runInContext("QBR.gtickets.list().length", sb), 0));
vm.runInContext('QBR.users.switch(QBR.users.list()[0].id)', sb);
t("robin keeps tickets", vm.runInContext("QBR.gtickets.list().length", sb) === 3);

/* export: ticket notes are [{d,text}] and must not become "[object Object]" */
sb.XLSX = { utils: { book_new: () => ({}), aoa_to_sheet: a => a, book_append_sheet: (wb, ws, n) => { wb[n] = ws; wb._last = wb; } }, writeFile: (wb) => { sb._wb = wb; } };
const ex = vm.runInContext("QBR.users.exportWorkbook()", sb);
t("export ok", ex.ok === true);
const tixRows = sb._wb["My Tickets"];
t("export notes not [object Object]", !JSON.stringify(tixRows).includes("object Object"));
t("export notes contain text", JSON.stringify(tixRows).indexOf("called vendor") >= 0);

/* views are told when user/data changes (so panels re-render) */
let fired = 0;
sb.CustomEvent = function (name) { this.type = name; };
sb.dispatchEvent = e => { if (e.type === "qbr:userdata") fired++; return true; };
vm.runInContext('QBR.users.register("Cy")', sb);
t("register notifies views", fired >= 1);
const f1 = fired;
vm.runInContext("QBR.users.switch(QBR.users.list()[0].id)", sb);
t("switch notifies views", fired > f1);
const f2 = fired;
vm.runInContext('QBR.gtickets.create({title:"Notify me"})', sb);
t("saving data notifies views", fired > f2);

console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
