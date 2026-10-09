// Core-feature checklist (2026-10-06). Fast static check that every shipped feature is still wired in:
// each entry = feature name + file + markers that must exist. A new build that drops a feature fails here
// even before the UI suites run. Add a line whenever a feature ships.
// Usage: node tests/feature-manifest.cjs            (app folder auto-detected: qbr-app/ or repo root)
const fs = require("fs"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const APP = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const read = f => { try { return fs.readFileSync(path.join(APP, f), "utf8"); } catch (e) { return null; } };
const F = [
  // [feature, file, [markers…]]
  ["Tracker engine + version", "js/excel-loader.js", ["QBR.VERSION", "QBR.scoreTenantHealth", "QBR.parseAuthMethods", "QBR.domainStatus", "QBR.buildCleanRows"]],
  ["Tenant Health Index + drill-down", "js/app.js", ["computeHealthIndex", "thi-drill-btn"]],
  ["School 360", "index.html", ["dash-school", "s360-body"]],
  ["Data Quality audit", "js/data-quality.js", ["QBR.audit"]],
  ["Exports (deck / images / clean)", "js/app.js", ["exportCleanData", "collectChartImages", "swapChartsForPrint"]],
  ["Welcome file guide", "index.html", ["dz-choose", "ALL TENANT AUTOMATED TRACKER.xlsx", "dz-foot"]],
  ["Rebrand + logo upload", "index.html", ["RCT OpsDesk", "brand-logo", "js/branding.js"]],
  ["Logo upload logic", "js/branding.js", ["rct-logo-v1", "RCT_BRAND"]],
  ["Inventory core", "js/inventory.js", ["QBR.invIntake", "QBR.invDeploy", "QBR.invAddTicket", "QBR.invUpdateAsset"]],
  ["Unregistered-serial links", "js/inventory.js", ["invSerialCell", "QBR.invRegisterSerial", "Unit not in inventory", "QBR.invSuggestSerials"]],
  ["Known warranty end", "js/inventory.js", ["QBR.invSetWarrantyEnd", "in-we"]],
  ["Batch codes + Batches view", "js/inventory.js", ["QBR.invNextBatchCode", "QBR.invBatches", "data-inv-batch-edit", "inv-f-batch"]],
  ["Bulk select / bulk edit + preview", "js/inventory.js", ["QBR.invBulkPlan", "QBR.invBulkApply", "inv-sel-all", "bk-preview", "bk-addr", "bk-wy"]],
  ["SQ deployment rows", "js/inventory.js", ["QBR.invAddDeployment", "QBR.invCurrentSq"]],
  ["Assigned-not-delivered flag", "js/inventory.js", ["assignedNoDel"]],
  ["Look up by DR / SQ / batch", "js/inventory.js", ["QBR.invGroupSearch", "invApplyGroup"]],
  ["Asset 360 edit: batch + SQ", "js/inventory.js", ["ae-batch", "ae-sq"]],
  ["Scan: batch review + stocktake", "js/scan.js", ["scanBatchTagDialog", "scanBatchClassify", "scan-stocktake", "sb-target", "sb-batch", "sb-wmode"]],
  ["Scan engine (ZXing + offline OCR)", "js/scan.js", ["QBR.SCAN_VERSION"]],
  ["Safe Save to Excel (.xlsm)", "js/persist.js", ["fsBookType", "fsEnsureVba", "fsVerifyBytes", "download-unsafe", "download-changed", "bookVBA"]],
  ["Patch-in-place save ops", "js/patch.js", ["invAddDeployment: patchInvAddDeployment", "pEnsureCol", "invUpdateAsset: patchInvUpdateAsset"]],
  ["Supplies module", "js/supplies.js", ["QBR.SUP_VERSION"]],
  ["Journal safety (v1.30.0)", "js/journal.js", ["QBR.journalPanel", "QBR.journalApplyOrphan", "QBR.journalImport", "QBR.journalPark", "beforeunload", "qbr-inv-journal-v2:"]],
  ["Journal script wired", "index.html", ["js/journal.js"]],
  ["Admin panel: flags, password, appearance (v1.31.0; dblclick v1.32.0)", "js/admin.js", ["QBR.adminSetFlag", "QBR.adminVerifyPassword", "QBR.adminThemeApply", "QBR.adminTriggerHit"]],
  ["Text-size stylesheet wired", "index.html", ["css/type-scale.css", "js/admin.js", "js/identity.js"]],
  ["Netlify Identity gate", "js/identity.js", ["QBR.identityInit", "isNetlifyHost"]],
  ["Environment label (Beta / Preview)", "js/branding.js", ["RCT_ENV", "Beta version"]],
  ["SOC risky investigation (v1.32.0)", "js/soc.js", ["QBR.socInit", "socImportReport", "socFilePrefix"]],
  ["Bulk user generator (v1.32.0)", "js/usergen.js", ["QBR.usergenInit", "ugFilePrefix", "O365_Users_Export.csv"]],
  ["Script library + GDAP", "js/scripts.js", ["QBR.scriptCopyText", "QBR.scriptsOpen"]],
  ["GDAP mapping", "js/gdap.js", ["QBR.gdapParseCsv", "QBR.gdapLookup"]],
  ["Scan save all images", "js/scan.js", ["scanDownloadAll", "dupOf"]],
  ["New sections wired", "index.html", ["dash-soc", "um-generator", "js/soc.js", "js/usergen.js", "js/scripts.js", "js/gdap.js"]],
  ["Offline libraries referenced", "index.html", ["libs/xlsx.full.min.js", "libs/chart.umd.min.js", "libs/tesseract/tesseract.min.js"]],
  ["SheetJS library file", "libs/xlsx.full.min.js", [""]],
  ["ZXing barcode library file", "libs/zxing/zxing.min.js", [""]],
  ["Tesseract offline OCR files", "libs/tesseract/offline/worker.js", [""]],
];
// Optional modules: checked only when present (so the manifest runs on builds from before they existed)
const OPTIONAL = [
  ["Audit editor + wizard + exempt (v1.28.0)", "js/audit.js", ["QBR.parseAuditBuffers", "QBR.renderAudit", "QBR.sanitizeHyperlinkTargets", "auditSetExempt"]],
];
let pass = 0, fail = 0;
const check = ([name, file, marks], opt) => {
  const src = read(file);
  if (src == null) { if (opt) { console.log("  · " + name + " (not in this build)"); return; } fail++; console.log("  ✗ " + name + " — missing file " + file); return; }
  const miss = marks.filter(m => src.indexOf(m) < 0);
  if (miss.length) { fail++; console.log("  ✗ " + name + " — missing: " + miss.join(", ")); }
  else { pass++; console.log("  ✓ " + name); }
};
console.log("== core feature checklist (" + path.relative(process.cwd(), APP) + ")");
F.forEach(f => check(f, false)); OPTIONAL.forEach(f => check(f, true));
console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
