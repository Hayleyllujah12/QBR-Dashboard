// Official logo (2026-10-26): RAKSO Ed transparent PNG is the default logo.
// Header logo must render wide (height 36, width well over 36), the favicon
// must point at the same PNG, and Reset to default must return to it.
// Usage: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/ui-official-logo.browser.cjs
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
  ok(errs.length === 0, "no page errors (errors: " + JSON.stringify(errs) + ")");
  const img = await page.evaluate(() => {
    const el = document.getElementById("brand-logo");
    if (!el) return null;
    const r = el.getBoundingClientRect();
    return { src: el.getAttribute("src"), alt: el.getAttribute("alt"), w: Math.round(r.width), h: Math.round(r.height), naturalW: el.naturalWidth };
  });
  ok(!!img && /rakso-ed-logo\.png$/.test(img.src || ""), "brand-logo src is rakso-ed-logo.png (got " + (img && img.src) + ")");
  ok(!!img && img.h === 36, "brand-logo renders 36px tall (got " + (img && img.h) + ")");
  ok(!!img && img.w > 100, "brand-logo renders wide, not squashed to a square (got " + (img && img.w) + "px wide)");
  ok(!!img && img.naturalW > 500, "logo image loaded at full resolution (naturalWidth " + (img && img.naturalW) + ")");
  const fav = await page.evaluate(() => {
    const l = document.querySelector('link[rel="icon"]');
    return l ? { href: l.getAttribute("href"), type: l.type } : null;
  });
  ok(!!fav && /rakso-ed-logo\.png$/.test(fav.href || ""), "favicon points at rakso-ed-logo.png (got " + (fav && fav.href) + ")");
  ok(!!fav && fav.type === "image/png", "favicon type is image/png (got " + (fav && fav.type) + ")");
  const afterReset = await page.evaluate(() => {
    window.RCT_BRAND.reset();
    const el = document.getElementById("brand-logo");
    return el ? el.getAttribute("src") : null;
  });
  ok(/rakso-ed-logo\.png$/.test(afterReset || ""), "Reset to default returns to rakso-ed-logo.png (got " + afterReset + ")");
  await browser.close();
  console.log(pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error("HARNESS ERROR", e); process.exit(1); });
