/* My To-Do + user profiles tests (Node). Run: node tools/test-todo.js */
"use strict";
const fs = require("fs"), vm = require("vm"), path = require("path");

function makeSb() {
  const sb = {
    console, XLSX: undefined, APP: { model: {} },
    document: { getElementById: () => null, createElement: () => ({ style: {} }), querySelectorAll: () => [] },
    localStorage: { _d: {}, getItem(k) { return this._d[k] || null; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } },
  };
  sb.window = sb; sb.globalThis = sb;
  vm.createContext(sb);
  return sb;
}

const sb = makeSb();
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "js", "users.js"), "utf8"), sb, { filename: "users.js" });
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "js", "todo.js"), "utf8"), sb, { filename: "todo.js" });

let pass = 0, fail = 0;
function t(name, cond) {
  if (cond) { pass++; }
  else { fail++; console.error("FAIL:", name); }
}
const eq = (a, b) => a === b;

/* users --------------------------------------------------------------- */
t("safeName strips forbidden chars",
  eq(vm.runInContext('QBR.users.safeName("a/b:c*d?e[f]g")', sb), "abcdefg"));
t("safeName caps at 31 chars",
  vm.runInContext('QBR.users.safeName("12345678901234567890123456789012345")', sb).length <= 31);
t("safeName trims", eq(vm.runInContext('QBR.users.safeName("  Ana  ")', sb), "Ana"));

let r = vm.runInContext('QBR.users.register("Tester")', sb);
t("register ok", r.ok === true && r.user.name === "Tester");
t("register duplicate rejected",
  vm.runInContext('QBR.users.register("tester")', sb).ok === false);
t("register blank rejected",
  vm.runInContext('QBR.users.register("   ")', sb).ok === false);
t("current user set", eq(vm.runInContext('QBR.users.current().name', sb), "Tester"));

r = vm.runInContext('QBR.users.register("Ana")', sb);
t("second user ok", r.ok === true);
t("two users listed", eq(vm.runInContext('QBR.users.list().length', sb), 2));
vm.runInContext('QBR.users.switch(QBR.users.list()[0].id)', sb);
t("switch back", eq(vm.runInContext('QBR.users.current().name', sb), "Tester"));

/* per-user data isolation */
vm.runInContext('var _d=QBR.users.myData();_d.todos.push({id:"x",title:"t",product:"general"});QBR.users.saveMyData(_d)', sb);
t("tester has todo", eq(vm.runInContext('QBR.users.myData().todos.length', sb), 1));
vm.runInContext('QBR.users.switch(QBR.users.list()[1].id)', sb);
t("ana isolated", eq(vm.runInContext('QBR.users.myData().todos.length', sb), 0));
vm.runInContext('QBR.users.switch(QBR.users.list()[0].id)', sb);

/* urgency -------------------------------------------------------------- */
const today = new Date(), iso = d => d.toISOString().slice(0, 10);
const past = new Date(today); past.setDate(past.getDate() - 2);
const soon = new Date(today); soon.setDate(soon.getDate() + 3);
const later = new Date(today); later.setDate(later.getDate() + 30);
t("overdue -> high", eq(vm.runInContext(`QBR.todo.autoUrgency("${iso(past)}")`, sb), "high"));
t("3 days -> medium", eq(vm.runInContext(`QBR.todo.autoUrgency("${iso(soon)}")`, sb), "medium"));
t("30 days -> low", eq(vm.runInContext(`QBR.todo.autoUrgency("${iso(later)}")`, sb), "low"));
t("no date -> low", eq(vm.runInContext("QBR.todo.autoUrgency(null)", sb), "low"));
t("override wins", eq(
  vm.runInContext(`QBR.todo.urgencyOf({due:"${iso(past)}",priority:"low"})`, sb), "low"));
t("auto priority defers", eq(
  vm.runInContext(`QBR.todo.urgencyOf({due:"${iso(past)}",priority:"auto"})`, sb), "high"));

/* todo add/complete/history */
let a = vm.runInContext(
  `QBR.todo.add({title:"Renew GDAP",product:"microsoft",due:"${iso(soon)}",priority:"auto",notes:""})`, sb);
t("add ok", a.ok === true);
t("add appears in all()", vm.runInContext("QBR.todo.all().some(t => t.title === 'Renew GDAP')", sb));
t("manual tile = microsoft",
  vm.runInContext("QBR.todo.all().filter(t => t.title === 'Renew GDAP')[0].product", sb) === "microsoft");
const mid = vm.runInContext("QBR.users.myData().todos.filter(t=>t.title==='Renew GDAP')[0].id", sb);
vm.runInContext(`QBR.todo.complete("${mid}")`, sb);
t("complete removes", !vm.runInContext("QBR.todo.all().some(t => t.title === 'Renew GDAP')", sb));
t("complete -> history", vm.runInContext("QBR.users.myData().history.length", sb) === 1);

/* auto: open Lenovo tickets */
sb.APP.model.inventory = { tickets: [
  { tno: "PF1-20260101", issue: "Screen", status: "Open", followUp: iso(past) },
  { tno: "PF2-20260102", issue: "Battery", status: "Closed" }
]};
const autos = vm.runInContext("QBR.todo.all().filter(t => t.auto && t.kind === 'ticket')", sb);
t("open ticket auto-pulled", autos.length === 1);
t("closed ticket excluded", !vm.runInContext("QBR.todo.all().some(t => t.title.indexOf('Battery') >= 0)", sb));
t("ticket tile = tickets", autos[0] && autos[0].product === "tickets");
t("ticket urgency from follow-up", autos[0] && autos[0].urgency === "high");

/* badge */
t("badge counts high/overdue", vm.runInContext("QBR.todo.badgeCount()", sb) >= 1);

/* tighten pass: render with a capturing host -------------------------- */
const sb2 = makeSb();
const hostStub = {
  _html: "",
  set innerHTML(v) { this._html = String(v); },
  get innerHTML() { return this._html; },
  querySelectorAll() { return []; },
  querySelector() { return null; },
};
sb2.document.getElementById = (id) => id === "dash-todo" ? hostStub : null;
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "js", "users.js"), "utf8"), sb2, { filename: "users.js" });
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "js", "todo.js"), "utf8"), sb2, { filename: "todo.js" });
vm.runInContext('QBR.users.register("Tester")', sb2);
for (let i = 0; i < 12; i++) {
  vm.runInContext(`QBR.todo.add({title:"G task ${i}",product:"general"})`, sb2);
}
sb2.APP.model.inventory = { tickets: [
  { tno: "PF9-20260101", sn: "PF9SN01", issue: "Wifi", status: "Open" }
]};
const _mnames = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
const _mkey = _mnames[new Date().getMonth()];
vm.runInContext(`QBR._audit = { months: { "${_mkey}": { rows: [{ school: "Demo School", org: "DEMO" }] } } }`, sb2);
vm.runInContext("QBR.todo.render()", sb2);
const html = hostStub._html;

t("tile order: General first", html.indexOf('id="todo-tile-general"') < html.indexOf('id="todo-tile-tickets"'));
t("tile order: manual before auto", html.indexOf('id="todo-tile-lenovo"') < html.indexOf('id="todo-tile-schools"'));
t("jump chips render", html.includes('data-todo-jump="schools"') && html.includes('data-todo-jump="general"'));
t("capped tile: Show all footer", html.includes("Show all 12"));
t("capped tile: 10 rows shown", (html.match(/G task /g) || []).length === 10);
t("no row-level No-date noise", !html.includes(">No date</span>"));
t("ticket title deep-links", html.includes('#ticket/PF9-20260101'));
t("serial deep-links to asset", html.includes('#asset/PF9SN01'));
t("school row links to audit", html.includes('data-todo-goto="audit"') && html.includes('data-month="' + _mkey + '"'));
t("manual row links to edit", html.includes('data-todo-goto="edit"'));

const mid2 = vm.runInContext("QBR.todo.add({title:'Old name',product:'microsoft'})", sb2).todo.id;
t("update ok", vm.runInContext(`QBR.todo.update("${mid2}", {title:"New name",product:"general"})`, sb2) === true);
t("update applied", vm.runInContext("QBR.todo.all().some(t => t.title === 'New name' && t.product === 'general')", sb2));
t("update missing -> false", vm.runInContext('QBR.todo.update("nope", {title:"x"})', sb2) === false);

console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
