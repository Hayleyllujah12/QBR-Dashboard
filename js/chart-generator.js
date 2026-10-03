/* =============================================================================
 * chart-generator.js
 * Thin, reusable wrappers over Chart.js using the Microsoft / Power BI palette.
 * Every helper destroys any prior chart bound to the same <canvas> before
 * drawing, so dashboards can be re-rendered on every filter change safely.
 * ===========================================================================*/

var QBR = (window.QBR = window.QBR || {});

QBR.COLORS = {
  blue: "#0078D4", green: "#107C10", orange: "#FF8C00", red: "#D13438",
  teal: "#038387", purple: "#8764B8", gray: "#605E5C", yellow: "#FFB900",
};
// ordered categorical palette for multi-series charts
QBR.PALETTE = ["#0078D4","#107C10","#FF8C00","#D13438","#8764B8","#038387","#FFB900","#605E5C"];

QBR._charts = {};  // canvasId -> Chart instance
QBR._configs = {}; // canvasId -> pristine config, cloned before Chart.js mutates it
                   // (export re-renders from this instead of touching live canvases)

// Read current theme colors from CSS vars and push them into Chart.js globals
// so axis text, ticks, and grid lines are legible in both light and dark mode.
QBR.applyChartTheme = function () {
  if (typeof Chart === "undefined") return;
  const cs = getComputedStyle(document.body);
  const axis = (cs.getPropertyValue("--axis") || "#605E5C").trim();
  const grid = (cs.getPropertyValue("--grid") || "rgba(0,0,0,.08)").trim();
  Chart.defaults.color = axis;
  Chart.defaults.borderColor = grid;
  if (Chart.defaults.scale) {
    Chart.defaults.scale.grid = Object.assign({}, Chart.defaults.scale.grid, { color: grid });
    Chart.defaults.scale.ticks = Object.assign({}, Chart.defaults.scale.ticks, { color: axis });
  }
};

// N distinct colors cycled from the categorical palette (for single-series bars).
QBR.paletteColors = function (n) {
  const out = [];
  for (let i = 0; i < n; i++) out.push(QBR.PALETTE[i % QBR.PALETTE.length]);
  return out;
};
// Fixed quarter colors so Q1–Q4 read the same across the app.
QBR.QUARTER_COLORS = { Q1: QBR.COLORS.blue, Q2: QBR.COLORS.green, Q3: QBR.COLORS.orange, Q4: QBR.COLORS.red };

function baseOpts(extra) {
  return Object.assign({
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 300 },
    plugins: {
      legend: { labels: { boxWidth: 12, font: { size: 11 } } },
      tooltip: { enabled: true },
    },
  }, extra || {});
}

// Destroy + (re)create a chart on a canvas id.
QBR.draw = function (canvasId, config) {
  const el = document.getElementById(canvasId);
  if (!el) return null;
  if (QBR._charts[canvasId]) { QBR._charts[canvasId].destroy(); delete QBR._charts[canvasId]; }
  // Snapshot the config first — Chart.js mutates what it is handed, and the
  // export path needs a clean copy to re-render from. Our configs are plain
  // data (no callbacks), so a JSON round-trip is a safe deep clone.
  try { QBR._configs[canvasId] = JSON.parse(JSON.stringify(config)); }
  catch (e) { QBR._configs[canvasId] = null; }
  const c = new Chart(el.getContext("2d"), config);
  QBR._charts[canvasId] = c;
  return c;
};

// Render a chart to a standalone PNG at slide resolution WITHOUT touching the
// live canvas. This sidesteps two traps at once: a chart inside a hidden panel
// has a 0x0 backing store, and re-sizing a live canvas fights Chart.js's own
// rAF-driven resize (which blanks the capture and, with fixed pixel styles,
// overflows the printed page). Default 960x540 @2x = a full 1920x1080 slide.
QBR.renderChartPng = function (canvasId, opts) {
  const cfg = QBR._configs[canvasId];
  if (!cfg) return null;
  const o = opts || {};
  const w = o.width || 960, h = o.height || 540, dpr = o.dpr || 2;

  // Offscreen but still laid out — Chart.js reads computed style on construct.
  const holder = document.createElement("div");
  holder.style.cssText = "position:fixed;left:-10000px;top:0;width:" + w + "px;height:" + h + "px;";
  const cv = document.createElement("canvas");
  // Logical size only. Chart.js multiplies this by devicePixelRatio below to
  // size the backing store — pre-multiplying here too would square the scale
  // (960x540 @2x became 3840x2160), and 33 canvases that large exhaust memory
  // and silently fail to decode, which renders as blank cards in the gallery.
  cv.width = w; cv.height = h;
  cv.style.width = w + "px"; cv.style.height = h + "px";
  holder.appendChild(cv);
  document.body.appendChild(holder);

  let url = null, tmp = null;
  try {
    tmp = new Chart(cv.getContext("2d"), {
      type: cfg.type,
      data: cfg.data,
      options: Object.assign({}, cfg.options, {
        responsive: false, maintainAspectRatio: false,
        animation: false, devicePixelRatio: dpr,
      }),
    });
    tmp.update("none");
    tmp.draw();                       // synchronous — never waits on rAF
    // Composite onto white: Chart.js canvases are transparent, and a
    // transparent PNG on a coloured slide loses all its axis text.
    const out = document.createElement("canvas");
    out.width = cv.width; out.height = cv.height;
    const ctx = out.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, out.width, out.height);
    ctx.drawImage(cv, 0, 0);
    url = out.toDataURL("image/png");
  } catch (e) { url = null; }
  if (tmp) { try { tmp.destroy(); } catch (e) {} }
  holder.remove();
  return url;
};

/* =============================================================================
 * Offline data-label + leader-line plugin (no external dependency).
 * Only acts on charts whose options carry `plugins.qbrLabels = {mode, labels}`,
 * so the ten deep-dive dashboards are untouched — only the Overview opts in.
 *   mode "donut" : value/% centered in each slice; a slice too thin for its
 *                  text gets a leader line to a collision-nudged outside label.
 *   mode "hbar"  : value at the bar end (inside if it fits, else just outside).
 *   mode "line"  : value floated above each point (null entries are skipped).
 * Labels are pre-formatted strings passed from app.js, so nothing here needs a
 * formatter callback — which also means the config survives the JSON clone the
 * export path makes, and labels render in Export Deck / Images automatically.
 * ===========================================================================*/
function _qbrTheme() {
  const cs = getComputedStyle(document.body);
  const g = (v, d) => (cs.getPropertyValue(v) || d).trim();
  return { ink: g("--ink", "#201F1E"), surface: g("--card", "#ffffff"), line: g("--axis", "#9a9a9a") };
}
function _qbrText(ctx, txt, x, y, fill, align, halo, haloW) {
  ctx.textAlign = align || "center";
  ctx.lineJoin = "round"; ctx.lineWidth = haloW || 3; ctx.strokeStyle = halo;
  ctx.strokeText(txt, x, y);
  ctx.fillStyle = fill; ctx.fillText(txt, x, y);
}
function _qbrDonut(chart, meta, spec, ctx, th) {
  const L = spec.labels, data = chart.data.datasets[0].data, out = { left: [], right: [] };
  meta.data.forEach((arc, i) => {
    if (!data[i] || L[i] == null) return;
    const p = arc.getProps(["x", "y", "startAngle", "endAngle", "innerRadius", "outerRadius"], true);
    const ang = (p.startAngle + p.endAngle) / 2, size = p.endAngle - p.startAngle;
    const rMid = (p.innerRadius + p.outerRadius) / 2, cos = Math.cos(ang), sin = Math.sin(ang);
    const tw = ctx.measureText(L[i]).width;
    if (size > 0.45 && size * rMid > tw + 10 && (p.outerRadius - p.innerRadius) > 18) {
      _qbrText(ctx, L[i], p.x + cos * rMid, p.y + sin * rMid, "#fff", "center", "rgba(0,0,0,.55)", th.halo);
    } else {
      out[cos >= 0 ? "right" : "left"].push({ p, cos, sin, txt: L[i], y: p.y + sin * (p.outerRadius + 2) });
    }
  });
  const W = chart.width, ca = chart.chartArea, gap = Math.max(15, Math.round(th.fpx * 1.35));
  const lead = th.isExport ? th.fpx * 1.1 : 9;
  ["left", "right"].forEach(side => {
    const arr = out[side].sort((a, b) => a.y - b.y), dir = side === "right" ? 1 : -1;
    if (!arr.length) return;
    // Initial Y = the slice's projected position just outside the ring.
    arr.forEach(o => { o.y = o.p.y + o.sin * (o.p.outerRadius + lead); });
    arr.sort((a, b) => a.y - b.y);
    // Spread apart so none overlap, then clamp the whole stack inside the area.
    for (let k = 1; k < arr.length; k++) if (arr[k].y - arr[k - 1].y < gap) arr[k].y = arr[k - 1].y + gap;
    const overBottom = arr[arr.length - 1].y - (ca.bottom - th.fpx);
    if (overBottom > 0) arr.forEach(o => { o.y -= overBottom; });
    const overTop = (ca.top + th.fpx) - arr[0].y;
    if (overTop > 0) arr.forEach(o => { o.y += overTop; });

    if (th.isExport) {
      // Ring-hugging column: anchor labels just OUTSIDE the ring on their side
      // (not the canvas edge) so the horizontal leader stays short for every
      // slice — including top-clustered ones. Elbow leader = radial out, then
      // a short run to the aligned column.
      const cx = arr[0].p.x, cy = arr[0].p.y, R = arr[0].p.outerRadius;
      const maxTw = Math.max.apply(null, arr.map(o => ctx.measureText(o.txt).width));
      const hgap = Math.round(th.fpx * 0.7);
      let colX = side === "right" ? (cx + R + hgap) : (cx - R - hgap);
      if (side === "right") colX = Math.min(colX, W - maxTw - 4);   // keep text on-canvas
      else colX = Math.max(colX, maxTw + 4);
      const bendX = side === "right" ? colX - 6 : colX + 6;
      arr.forEach(o => {
        const x0 = cx + o.cos * R, y0 = cy + o.sin * R;
        const x1 = cx + o.cos * (R + lead), y1 = cy + o.sin * (R + lead);
        ctx.strokeStyle = th.line; ctx.lineWidth = 1; ctx.beginPath();
        ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); ctx.lineTo(bendX, o.y); ctx.stroke();
        _qbrText(ctx, o.txt, colX, o.y, th.ink, side === "right" ? "left" : "right", th.surface, th.halo);
      });
    } else {
      // On-screen cards have no gutter room — keep labels beside their slice.
      arr.forEach(o => {
        const tw = ctx.measureText(o.txt).width;
        const x0 = o.p.x + o.cos * o.p.outerRadius, y0 = o.p.y + o.sin * o.p.outerRadius;
        const x1 = o.p.x + o.cos * (o.p.outerRadius + 9), y1 = o.p.y + o.sin * (o.p.outerRadius + 9);
        let tx = x1 + dir * 7;
        if (side === "right") tx = Math.min(tx, W - tw - 3); else tx = Math.max(tx, tw + 3);
        ctx.strokeStyle = th.line; ctx.lineWidth = 1; ctx.beginPath();
        ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); ctx.lineTo(tx, o.y); ctx.stroke();
        _qbrText(ctx, o.txt, tx, o.y, th.ink, side === "right" ? "left" : "right", th.surface, th.halo);
      });
    }
  });
}
function _qbrHbar(chart, meta, spec, ctx, th) {
  const L = spec.labels;
  meta.data.forEach((el, i) => {
    if (L[i] == null) return;
    const p = el.getProps(["x", "y", "base"], true), tw = ctx.measureText(L[i]).width;
    if (Math.abs(p.x - p.base) > tw + 14) _qbrText(ctx, L[i], p.x - 7, p.y, "#fff", "right", "rgba(0,0,0,.45)", th.halo);
    else _qbrText(ctx, L[i], p.x + 7, p.y, th.ink, "left", th.surface, th.halo);
  });
}
// Point labels. `spec.labels` may be a flat array (single series -> dataset 0,
// original behaviour) or a 2D array [datasetIndex][pointIndex] to label every
// series. With `spec.colored`, each label takes its line's colour; multi-series
// labels alternate above/below the point so overlapping lines don't collide.
function _qbrLine(chart, meta, spec, ctx, th) {
  const L = spec.labels;
  const twoD = Array.isArray(L) && Array.isArray(L[0]);
  const idxs = twoD ? chart.data.datasets.map((_, i) => i) : [0];
  idxs.forEach(di => {
    const dm = chart.getDatasetMeta(di);
    if (dm.hidden) return;
    const labs = twoD ? L[di] : L;
    if (!labs) return;
    const ds = chart.data.datasets[di];
    const col = spec.colored ? (ds.borderColor || ds.backgroundColor || th.ink) : th.ink;
    const dir = (twoD && di % 2 === 1) ? 1 : -1;         // series 0 above, series 1 below
    const dy = dir * (th.isExport ? th.fpx : 12);
    dm.data.forEach((pt, i) => {
      if (labs[i] == null) return;
      const p = pt.getProps(["x", "y"], true);
      _qbrText(ctx, labs[i], p.x, p.y + dy, col, "center", th.surface, th.halo);
    });
  });
}
function _qbrVbar(chart, meta, spec, ctx, th) {
  const L = spec.labels, ca = chart.chartArea;
  meta.data.forEach((el, i) => {
    if (L[i] == null) return;
    const p = el.getProps(["x", "y"], true);
    // above the bar top, unless too close to the top edge — then just inside
    if (p.y - 14 < ca.top) _qbrText(ctx, L[i], p.x, p.y + 11, "#fff", "center", "rgba(0,0,0,.5)", th.halo);
    else _qbrText(ctx, L[i], p.x, p.y - 9, th.ink, "center", th.surface, th.halo);
  });
}
// Stacked bars: label every segment (across all datasets) at its centre, but
// only when the segment is tall/wide enough to hold the number — tiny slivers
// are skipped so labels never collide. Reads values straight off the datasets.
function _qbrStack(chart, spec, ctx, th) {
  // spec.labels (optional) is a 2D array [datasetIndex][pointIndex] of strings;
  // when absent, the raw value is drawn.
  chart.data.datasets.forEach((ds, di) => {
    const meta = chart.getDatasetMeta(di);
    if (meta.hidden) return;
    meta.data.forEach((el, i) => {
      const v = ds.data[i];
      if (!v) return;
      const lab = (spec.labels && spec.labels[di] && spec.labels[di][i] != null) ? spec.labels[di][i] : String(v);
      if (lab == null || lab === "") return;
      const p = el.getProps(["x", "y", "base", "horizontal"], true);
      const extent = p.horizontal ? Math.abs(p.x - p.base) : Math.abs(p.base - p.y);
      if (extent < ctx.measureText(lab).width + 8) return;   // too small to hold the label
      const cx = p.horizontal ? (p.x + p.base) / 2 : p.x;
      const cy = p.horizontal ? p.y : (p.y + p.base) / 2;
      _qbrText(ctx, lab, cx, cy, "#fff", "center", "rgba(0,0,0,.5)", th.halo);
    });
  });
}
// Column/bar TOTAL for a stacked chart: sum of all datasets at each index, drawn
// above the column (vertical) or just past the bar end (horizontal). spec.totalSuffix
// appends a unit (e.g. " TB"). Enabled with spec.showTotal.
function _qbrStackTotals(chart, spec, ctx, th) {
  const dss = chart.data.datasets, meta0 = chart.getDatasetMeta(0);
  if (!meta0.data.length) return;
  const horizontal = !!meta0.data[0].horizontal, suffix = spec.totalSuffix || "";
  const bold = '700 ' + th.fpx + 'px "Segoe UI",system-ui,sans-serif';
  ctx.save(); ctx.font = bold;
  for (let i = 0; i < meta0.data.length; i++) {
    let total = 0; dss.forEach(ds => { const v = ds.data[i]; if (typeof v === "number") total += v; });
    if (!total) continue;
    const lab = _qbrNum(total) + suffix, el = meta0.data[i];
    if (horizontal) {
      const y = el.getProps(["y"], true).y, x = chart.scales.x.getPixelForValue(total) + (th.isExport ? 10 : 7);
      _qbrText(ctx, lab, x, y, th.ink, "left", th.surface, th.halo);
    } else {
      const x = el.getProps(["x"], true).x, y = chart.scales.y.getPixelForValue(total) - (th.isExport ? th.fpx * 0.9 : 10);
      _qbrText(ctx, lab, x, y, th.ink, "center", th.surface, th.halo);
    }
  }
  ctx.restore();
}
const QBRLabelsPlugin = {
  id: "qbrLabels",
  afterDatasetsDraw(chart) {
    const spec = chart.options.plugins && chart.options.plugins.qbrLabels;
    if (!spec) return;
    if (spec.mode !== "stack" && !spec.labels) return;   // stack reads values directly
    const ctx = chart.ctx, meta = chart.getDatasetMeta(0), th = _qbrTheme();
    // Export renders offscreen with responsive:false — scale the font up there so
    // labels aren't tiny in the 1920x1080 PNG. On-screen (responsive:true) → 11px.
    th.isExport = chart.options && chart.options.responsive === false;
    th.fpx = th.isExport ? Math.max(14, Math.round(chart.height / 28)) : 11;
    th.halo = th.isExport ? 4 : 3;
    ctx.save();
    ctx.font = '700 ' + th.fpx + 'px "Segoe UI",system-ui,sans-serif';
    ctx.textBaseline = "middle";
    try {
      if (spec.mode === "donut") _qbrDonut(chart, meta, spec, ctx, th);
      else if (spec.mode === "hbar") _qbrHbar(chart, meta, spec, ctx, th);
      else if (spec.mode === "vbar") _qbrVbar(chart, meta, spec, ctx, th);
      else if (spec.mode === "line") _qbrLine(chart, meta, spec, ctx, th);
      else if (spec.mode === "stack") { _qbrStack(chart, spec, ctx, th); if (spec.showTotal) _qbrStackTotals(chart, spec, ctx, th); }
    } catch (e) {}
    ctx.restore();
  },
};
// Draws a value + sub-label in the hole of a doughnut/gauge, ON the canvas — so
// it survives the Export Images PNG re-render (an HTML overlay would not).
// Greedy word-wrap into lines that each fit maxW at the current ctx.font.
function _qbrWrap(ctx, text, maxW) {
  const words = String(text).split(/\s+/), lines = [];
  let cur = "";
  words.forEach(w => {
    const t = cur ? cur + " " + w : w;
    if (ctx.measureText(t).width <= maxW || !cur) cur = t;
    else { lines.push(cur); cur = w; }
  });
  if (cur) lines.push(cur);
  return lines;
}
const QBRCenterPlugin = {
  id: "qbrCenter",
  afterDraw(chart) {
    const o = chart.options.plugins && chart.options.plugins.qbrCenter;
    if (!o || (o.value == null && !o.title)) return;
    const ctx = chart.ctx, ca = chart.chartArea, th = _qbrTheme();
    const cx = (ca.left + ca.right) / 2;
    const half = chart.options.circumference === 180;
    const cy = half ? ca.bottom - (ca.bottom - ca.top) * 0.30 : (ca.top + ca.bottom) / 2;
    ctx.save();
    ctx.textAlign = "center"; ctx.textBaseline = "middle";

    if (o.value != null) {
      // Big value + sub-label (gauge / KPI use).
      const g = ctx.createLinearGradient(cx - 34, cy - 18, cx + 34, cy + 18);
      g.addColorStop(0, o.color || "#107C10"); g.addColorStop(1, o.color2 || "#13a10e");
      ctx.fillStyle = g;
      const isExport = chart.options && chart.options.responsive === false;
      const vpx = isExport ? Math.max(30, Math.round(chart.height / 12)) : 30;
      const spx = isExport ? Math.max(11, Math.round(vpx * 0.36)) : 11;
      ctx.font = '800 ' + vpx + 'px "Segoe UI",system-ui,sans-serif';
      ctx.fillText(String(o.value), cx, cy - (o.label ? vpx * 0.24 : 0));
      if (o.label) {
        ctx.fillStyle = th.line;
        ctx.font = '600 ' + spx + 'px "Segoe UI",system-ui,sans-serif';
        ctx.fillText(String(o.label).toUpperCase(), cx, cy + vpx * 0.55);
      }
    } else if (o.title && chart.options && chart.options.responsive === false) {
      // Center title is EXPORT-ONLY (responsive:false = offscreen PNG render).
      // On the live dashboard it is intentionally not drawn.
      // Needs a real hole — read the inner radius; skip on true pies (inner ~ 0).
      const meta = chart.getDatasetMeta(0), arc = meta && meta.data && meta.data[0];
      const inner = arc ? arc.getProps(["innerRadius"], true).innerRadius : 0;
      if (inner < 18) { ctx.restore(); return; }
      const maxW = inner * 1.7;
      let fs = Math.max(11, Math.round(inner * 0.36)), lines;
      for (; ; fs--) {
        ctx.font = "700 " + fs + 'px "Segoe UI",system-ui,sans-serif';
        lines = _qbrWrap(ctx, o.title, maxW);
        if (fs <= 10 || (lines.length * fs * 1.18 <= inner * 1.7 && lines.every(l => ctx.measureText(l).width <= maxW))) break;
      }
      const lh = fs * 1.18, y0 = cy - (lines.length - 1) * lh / 2;
      ctx.fillStyle = th.ink;
      lines.forEach((l, i) => ctx.fillText(l, cx, y0 + i * lh));
    }
    ctx.restore();
  },
};
// Draws categorical labels under a scatter's numeric x positions (0..n-1).
// options.plugins.qbrXcats = ["Enabled", ...]. Export-safe (plain data).
const QBRXCatsPlugin = {
  id: "qbrXcats",
  afterDraw(chart) {
    const spec = chart.options.plugins && chart.options.plugins.qbrXcats;
    const cats = spec && spec.cats;
    if (!Array.isArray(cats) || !chart.scales.x) return;
    const ctx = chart.ctx, xs = chart.scales.x, ca = chart.chartArea, th = _qbrTheme();
    const isExport = chart.options && chart.options.responsive === false;
    const fpx = isExport ? Math.max(13, Math.round(chart.height / 34)) : 11;
    ctx.save();
    ctx.fillStyle = th.axis || th.ink; ctx.textAlign = "center"; ctx.textBaseline = "top";
    ctx.font = "600 " + fpx + 'px "Segoe UI",system-ui,sans-serif';
    cats.forEach((c, i) => { const x = xs.getPixelForValue(i); if (x >= ca.left - 20 && x <= ca.right + 20) ctx.fillText(c, x, ca.bottom + 6); });
    ctx.restore();
  },
};
if (typeof Chart !== "undefined") { Chart.register(QBRLabelsPlugin); Chart.register(QBRCenterPlugin); Chart.register(QBRXCatsPlugin); }

// number formatter for chart labels
function _qbrNum(n) { return Number(n).toLocaleString(undefined, { maximumFractionDigits: 1 }); }

// Pull an optional `dataLabels` spec out of an opts object and fold it into
// Chart.js plugin options (leaving the rest of opts intact).
function withLabels(options, dl, isDonut) {
  if (!dl) return options;
  options.plugins = options.plugins || {};
  options.plugins.qbrLabels = dl;
  if (isDonut) options.layout = Object.assign({ padding: { left: 52, right: 52, top: 6, bottom: 6 } }, options.layout);
  return options;
}

QBR.chart = {
  line(canvasId, labels, datasets, opts) {
    const dl = opts && opts.dataLabels; if (opts) delete opts.dataLabels;
    return QBR.draw(canvasId, {
      type: "line",
      data: { labels, datasets: datasets.map((d, i) => Object.assign({
        borderColor: d.color || QBR.PALETTE[i % QBR.PALETTE.length],
        backgroundColor: (d.color || QBR.PALETTE[i % QBR.PALETTE.length]) + "22",
        tension: 0.3, fill: !!d.fill, pointRadius: 3, borderWidth: 2,
      }, d)) },
      options: withLabels(baseOpts(opts), dl),
    });
  },
  bar(canvasId, labels, datasets, opts) {
    const dl = opts && opts.dataLabels; if (opts) delete opts.dataLabels;
    const options = withLabels(baseOpts(opts), dl);
    // Show EVERY category label. Chart.js `autoSkip` silently drops labels that would
    // crowd/overlap (long rotated org names lost RCAMES/SPCEM/MAPSA/CEAP etc.). Force
    // all of them on; a caller's explicit x-tick options still win via Object.assign.
    options.scales = options.scales || {};
    options.scales.x = options.scales.x || {};
    // Show EVERY label (autoSkip off) but bound its length so long, rotated category
    // names (e.g. full org titles like "SCHOOLS DIVISION OFFICE (SDO) OF TAGUIG…")
    // can't overlap their neighbours. Full text stays in the tooltip (which reads
    // chart.data.labels, not this callback). Only strings > 16 chars are trimmed, so
    // short labels (Q1–Q4, service names, numeric axes) are untouched. The callback is
    // dropped from the export clone by JSON round-trip — harmless, the hi-res PNG has room.
    options.scales.x.ticks = Object.assign({
      autoSkip: false, maxRotation: 90,
      callback: function (v) { var s = String(this.getLabelForValue(v)); return s.length > 16 ? s.slice(0, 15) + "…" : s; },
    }, options.scales.x.ticks);
    return QBR.draw(canvasId, {
      type: "bar",
      data: { labels, datasets: datasets.map((d, i) => Object.assign({
        backgroundColor: d.color || QBR.PALETTE[i % QBR.PALETTE.length],
        borderRadius: 4,
      }, d)) },
      options,
    });
  },
  hbar(canvasId, labels, values, color, opts) {
    const dl = opts && opts.dataLabels; if (opts) delete opts.dataLabels;
    return QBR.draw(canvasId, {
      type: "bar",
      data: { labels, datasets: [{ data: values, backgroundColor: color || QBR.COLORS.blue, borderRadius: 4 }] },
      options: withLabels(baseOpts(Object.assign({ indexAxis: "y", plugins: { legend: { display: false } } }, opts)), dl),
    });
  },
  pie(canvasId, labels, values, colors, opts) {
    const dl = opts && opts.dataLabels; if (opts) delete opts.dataLabels;
    return QBR.draw(canvasId, {
      type: "pie",
      data: { labels, datasets: [{ data: values, backgroundColor: colors || QBR.PALETTE }] },
      options: withLabels(baseOpts(opts), dl, true),
    });
  },
  doughnut(canvasId, labels, values, colors, opts) {
    const dl = opts && opts.dataLabels; if (opts) delete opts.dataLabels;
    const center = opts && opts.center; if (opts) delete opts.center;
    const options = withLabels(baseOpts(Object.assign({ cutout: "60%" }, opts)), dl, true);
    if (center) { options.plugins = options.plugins || {}; options.plugins.qbrCenter = Object.assign({}, options.plugins.qbrCenter, { title: center }); }
    return QBR.draw(canvasId, {
      type: "doughnut",
      data: { labels, datasets: [{ data: values, backgroundColor: colors || QBR.PALETTE }] },
      options,
    });
  },
  // Single horizontal 100%-composition bar. segments: [{label,value,color}].
  // Each segment is labelled "value · pct%" when it's wide enough.
  stacked100(canvasId, segments) {
    const total = segments.reduce((a, s) => a + (s.value || 0), 0) || 1;
    const datasets = segments.map(s => ({ label: s.label, data: [s.value || 0], backgroundColor: s.color, maxBarThickness: 60, borderWidth: 0 }));
    const labels2d = segments.map(s => [s.value ? `${_qbrNum(s.value)} · ${Math.round(s.value / total * 100)}%` : null]);
    const options = baseOpts({
      indexAxis: "y",
      scales: { x: { stacked: true, display: false, min: 0, max: total, grid: { display: false } },
                y: { stacked: true, display: false, grid: { display: false } } },
      plugins: { legend: { position: "top", labels: { boxWidth: 12, font: { size: 11 } } } },
    });
    options.plugins.qbrLabels = { mode: "stack", labels: labels2d };
    return QBR.draw(canvasId, { type: "bar", data: { labels: [""], datasets }, options });
  },
  // Scatter with a categorical x-axis (positions 0..n-1 labelled by `cats`).
  // points: [{x,y,color}]. Used for the security-posture vs risky correlation.
  scatter(canvasId, points, cats, opts) {
    const options = baseOpts(Object.assign({
      scales: {
        x: { min: -0.5, max: cats.length - 0.5, ticks: { display: false }, grid: { display: false }, offset: false },
        y: { beginAtZero: true, title: { display: true, text: "Risky Users" } },
      },
      plugins: { legend: { display: false } },
      layout: { padding: { bottom: 22 } },
    }, opts || {}));
    options.plugins = options.plugins || {};
    options.plugins.qbrXcats = { cats: cats };
    return QBR.draw(canvasId, {
      type: "scatter",
      data: { datasets: [{ data: points.map(p => ({ x: p.x, y: p.y })), pointBackgroundColor: points.map(p => p.color), pointBorderColor: "rgba(0,0,0,.25)", pointRadius: 4, pointHoverRadius: 6 }] },
      options,
    });
  },
};

if (typeof module !== "undefined" && module.exports) module.exports = QBR;
