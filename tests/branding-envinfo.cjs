// Hosted branch badge (2026-10-26) — Node vm harness for js/branding.js envInfo().
// The browser suite (tests/branding.browser.cjs) can only run file://, where the
// hostname is empty, so hosted cases are exercised here with stubbed location:
// a stamped branch build must show "Branch" on GitHub Pages / Azure SWA / Netlify
// hosts that previously labelled it Main / Beta / Preview, while explicit
// ?main/?beta/?dev flags still win and unstamped builds are unchanged.
// No real hosts, no real people. Usage: node tests/branding-envinfo.cjs
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

function envFor(opts) {
  const sub = { firstChild: null, insertBefore() {} };
  const sandbox = {
    console, setTimeout, clearTimeout, Promise, String, Array, Object, Date, Math, JSON,
    location: { pathname: opts.path || "/index.html", hostname: opts.host || "", search: opts.query || "" },
    document: {
      readyState: "complete",
      getElementById: () => null,
      querySelector: (sel) => (sel === ".app-sub" ? sub : null),
      createElement: () => ({ className: "", setAttribute() {}, classList: { toggle() {} } }),
      body: { appendChild() {} },
      addEventListener() {}, removeEventListener() {},
      title: "RCT OpsDesk",
    },
    localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    window: {},
  };
  sandbox.window.QBR = {};
  if (opts.stamp !== undefined) sandbox.window.QBR_BRANCH = opts.stamp;
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  const src = fs.readFileSync(path.join(__dirname, "..", "js", "branding.js"), "utf8");
  vm.runInContext(src, sandbox);
  return sandbox.window.RCT_ENV();
}

function check(label, opts, wantText, wantTag) {
  const r = envFor(opts);
  if (wantText === null) {
    ok(r === null, label + ": envInfo() null (got " + JSON.stringify(r) + ")");
  } else {
    ok(!!r && r.text === wantText && r.tag === wantTag,
      label + ": '" + (r && r.text) + "' " + (r && r.tag) + " (want '" + wantText + "' " + wantTag + ")");
  }
}

// Hosted fix cases: stamp beats hostname-derived labels.
check("github.io + stamp", { host: "hayleyllujah12.github.io", stamp: "MOS-RECON" }, "Branch · MOS-RECON", "[BRANCH]");
check("Azure SWA + stamp", { host: "ambitious-grass-0e4be480f.6.azurestaticapps.net", stamp: "MOS-RECON" }, "Branch · MOS-RECON", "[BRANCH]");
check("Netlify branch host + stamp", { host: "mos-recon--rct-ops-desk.netlify.app", stamp: "MOS-RECON" }, "Branch · MOS-RECON", "[BRANCH]");
check("Netlify main host + stamp", { host: "rct-ops-desk.netlify.app", stamp: "MOS-RECON" }, "Branch · MOS-RECON", "[BRANCH]");
check("versioned folder path + stamp", { path: "/QBR-Dashboard_1.21.0/index.html", stamp: "MOS-RECON" }, "Branch · MOS-RECON", "[BRANCH]");
check("?branch= beats github.io + stamp", { host: "hayleyllujah12.github.io", query: "?branch=TEST-X", stamp: "MOS-RECON" }, "Branch · TEST-X", "[BRANCH]");
// Explicit flags still win over the stamp.
check("?main + stamp on github.io", { host: "hayleyllujah12.github.io", query: "?main", stamp: "MOS-RECON" }, "Main", "[MAIN]");
check("?beta + stamp on github.io", { host: "hayleyllujah12.github.io", query: "?beta", stamp: "MOS-RECON" }, "Beta version", "[BETA]");
check("?dev + ?branch=", { host: "hayleyllujah12.github.io", query: "?dev&branch=TEST-X", stamp: "MOS-RECON" }, "Beta version", "[BETA]");
// Unstamped builds unchanged.
check("github.io, no stamp", { host: "hayleyllujah12.github.io" }, "Main", "[MAIN]");
check("Azure SWA, no stamp", { host: "ambitious-grass-0e4be480f.6.azurestaticapps.net" }, "Beta version", "[BETA]");
check("Netlify branch host, no stamp", { host: "mos-recon--rct-ops-desk.netlify.app" }, "Preview · mos-recon", "[PREVIEW]");
check("Azure PR staging, no stamp", { host: "rct-ops-abc123-42.6.azurestaticapps.net" }, "Staging · PR 42", "[STAGING]");
check("Netlify develop host, no stamp", { host: "develop--rct-ops-desk.netlify.app" }, "Beta version", "[BETA]");
// Local regression: file:// + stamp, and no label at all.
check("file:// + stamp", { host: "", stamp: "MOS-RECON" }, "Branch · MOS-RECON", "[BRANCH]");
check("file://, no label", { host: "" }, null, null);

console.log(pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
