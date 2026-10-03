#!/usr/bin/env node
/* =============================================================================
 * embed-domainreg.cjs — bake a domain-registration workbook into the app as the
 * built-in DEFAULT, so the offline app opens with the Domain Registration tab
 * already populated (no separate upload). A user-uploaded domain workbook still
 * overrides this seed at runtime.
 *
 * The app is an offline file:// page and browsers can't fetch() a loose local
 * file, so the bytes are embedded as base64 in qbr-app/data/domainreg-seed.js.
 * This makes the baked copy a POINT-IN-TIME SNAPSHOT — re-run this script to
 * refresh it whenever the domain data changes.
 *
 * USAGE:
 *   node tests/embed-domainreg.cjs "PATH\TO\QBR_domainreg.xlsx"
 *   (defaults to ./QBR_domainreg.xlsx next to this script if omitted)
 * ===========================================================================*/
const fs = require("fs"), path = require("path");
const base = __dirname;
function findDir(cands) { for (const c of cands) { try { if (fs.statSync(c).isDirectory()) return c; } catch (e) {} } return null; }
const appDir = findDir([base + "/../qbr-app", base + "/../app", base + "/.."]);
if (!appDir) { console.error("Could not locate qbr-app/ (or app/) next to this script."); process.exit(2); }

const src = process.argv[2] || (base + "/QBR_domainreg.xlsx");
if (!fs.existsSync(src)) { console.error("Domain workbook not found: " + src + "\nPass the path: node tests/embed-domainreg.cjs \"PATH\\TO\\QBR_domainreg.xlsx\""); process.exit(2); }

const bytes = fs.readFileSync(src);
const b64 = bytes.toString("base64");
const dataDir = appDir + "/data";
try { fs.mkdirSync(dataDir, { recursive: true }); } catch (e) {}
const out = dataDir + "/domainreg-seed.js";
const today = new Date().toISOString().slice(0, 10);
const js =
  "/* AUTO-GENERATED — do not edit by hand.\n" +
  " * Built-in DEFAULT domain-registration workbook, embedded so the offline app\n" +
  " * opens with the Domain Registration tab populated (no separate upload). This\n" +
  " * is a point-in-time snapshot; refresh it with:\n" +
  " *   node tests/embed-domainreg.cjs \"PATH\\\\TO\\\\QBR_domainreg.xlsx\"\n" +
  " * A user-uploaded domain workbook overrides this seed at runtime. */\n" +
  "window.QBR = window.QBR || {};\n" +
  "QBR.SEED_DOMAINREG = { name: " + JSON.stringify(path.basename(src)) +
  ", generated: " + JSON.stringify(today) +
  ", b64: \"" + b64 + "\" };\n";
fs.writeFileSync(out, js);
console.log("wrote " + out + "  (source " + path.basename(src) + ", " + bytes.length + " bytes → " + js.length + " B JS)");
