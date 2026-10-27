// Folder per-file skip — browser wiring check (file:// half; http half runs on the requester's machine).
// Loads the real app and confirms the edited fsfolder.js loads cleanly and
// QBR.fsOpenFolder is present. The read-loop behaviour itself is exercised by
// tests/fs-perfile-skip.cjs (Node vm, fake File System Access handles),
// because showDirectoryPicker cannot be driven headlessly here.
// Usage: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/ui-fs-perfile-skip.browser.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path");
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }
(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME, args: ["--no-sandbox"] });
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await ctx.newPage();
  const errs = [];
  page.on("pageerror", (e) => errs.push(String(e.message)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(1000);
  ok(!errs.some((m) => /fsfolder/i.test(m)), "no fsfolder page error (errors: " + JSON.stringify(errs) + ")");
  const hasFn = await page.evaluate(() => typeof window.QBR !== "undefined" && typeof window.QBR.fsOpenFolder === "function");
  ok(hasFn, "QBR.fsOpenFolder is a function in the loaded app");
  const hasReload = await page.evaluate(() => typeof window.QBR.fsFolderReloadAll === "function");
  ok(hasReload, "QBR.fsFolderReloadAll still present (no collateral break)");
  await browser.close();
  console.log(pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error("HARNESS ERROR", e); process.exit(1); });
