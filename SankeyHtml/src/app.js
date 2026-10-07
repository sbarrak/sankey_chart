import { select, pointer } from "d3-selection";
import { drag } from "d3-drag";
import { sankey, sankeyLinkHorizontal, sankeyJustify, sankeyLeft, sankeyRight, sankeyCenter } from "d3-sankey";
import { parseNum, parseDelimited, readXlsx, detectTable } from "./io.js";

const $ = (s, r = document) => r.querySelector(s);
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

const PALETTES = {
  corporate: ["#1F4E79", "#2E75B6", "#0096A0", "#70AD47", "#ED7D31", "#FFC000", "#5B9BD5", "#A5A5A5", "#44546A", "#C0504D"],
  vibrant: ["#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F", "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC"],
  pastel: ["#8CAABE", "#F096A0", "#96C8EB", "#F5C896", "#96E1AF", "#CDA0D7", "#EBE18C", "#AFCD8C", "#E6A5A5", "#BEBEB4"],
  earth: ["#8C5632", "#BF9000", "#70803C", "#A0522D", "#5E7279", "#CD853F", "#556B2F", "#8B7355", "#B26646", "#787864"],
  ocean: ["#005F73", "#0A9396", "#6EBEAA", "#D2BE78", "#EE9B00", "#CA6702", "#BB3E03", "#AE2012", "#9B2226", "#003C50"]
};

const SAMPLE = {
  headers: ["From", "To", "Value"],
  rows: [["Salary", "Budget", 5000], ["Freelance", "Budget", 1200], ["Investments", "Budget", 800],
    ["Budget", "Housing", 2400], ["Budget", "Food", 1100], ["Budget", "Transport", 600], ["Budget", "Savings", 1800],
    ["Budget", "Leisure", 1100], ["Savings", "Emergency Fund", 600], ["Savings", "Retirement", 1200],
    ["Housing", "Rent", 2000], ["Housing", "Utilities", 400]]
};
const SAMPLE_TITLE = "Monthly Budget Flow";
const SAMPLE_SUB = "Where the money comes from and where it goes";

const state = {
  table: { headers: [], rows: [] },
  map: { from: 0, to: 1, value: 2 },
  reverse: false,
  sheets: null,
  fileLabel: "Sample data",
  s: {
    title: SAMPLE_TITLE, subtitle: SAMPLE_SUB, palette: "corporate", customPalette: "", linkStyle: "gradient",
    opacity: 50, bg: "white", nodeWidth: 14, nodePadding: 12, align: "justify", width: 0, height: 0,
    showLabels: true, showValues: true, fontSize: 12, labelPos: "auto", numFmt: "auto", unit: "",
    minValue: 0, scale: "2", fileName: "sankey"
  },
  nodeColors: {},     // node name -> #hex
  manual: {},         // node name -> {dx, dy} (dragged offsets from the automatic layout)
  sel: null,          // selected node name
  crop: null,         // {x, y, w, h} in chart px
  cropMode: false
};

// ---------------------------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------------------------
let toastTimer;
function toast(msg) {
  const t = $("#toast");
  t.textContent = msg; t.classList.add("show");
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove("show"), 2600);
}
const raf = (() => { let p = false; return (fn) => { if (p) return; p = true; requestAnimationFrame(() => { p = false; fn(); }); }; })();

function fmt(v) {
  const s = state.s;
  let t;
  if (s.numFmt === "compact") t = new Intl.NumberFormat(undefined, { notation: "compact", maximumFractionDigits: 1 }).format(v);
  else if (s.numFmt === "auto") t = new Intl.NumberFormat(undefined, { maximumFractionDigits: Math.abs(v - Math.round(v)) < 1e-6 ? 0 : 2 }).format(v);
  else t = new Intl.NumberFormat(undefined, { minimumFractionDigits: +s.numFmt, maximumFractionDigits: +s.numFmt }).format(v);
  return s.unit ? `${t} ${s.unit}` : t;
}

function parseHexList(str) {
  const out = [];
  for (const m of (str || "").matchAll(/(?<![0-9a-fA-F])#?([0-9a-fA-F]{6})(?![0-9a-fA-F])/g)) out.push("#" + m[1].toUpperCase());
  return out;
}

function palette() {
  if (state.s.palette === "custom") {
    const c = parseHexList(state.s.customPalette);
    if (c.length) return c;
  }
  return PALETTES[state.s.palette] || PALETTES.corporate;
}
const colorOf = (name, index) => state.nodeColors[name] || palette()[index % palette().length];

// ---------------------------------------------------------------------------------------------
// data -> links
// ---------------------------------------------------------------------------------------------
function deriveLinks() {
  const { from, to, value } = state.map;
  const agg = new Map();
  let maxFlow = 0;
  state.table.rows.forEach((r, ri) => {
    let s = String(r[from] ?? "").trim(), t = String(r[to] ?? "").trim();
    const v = parseNum(r[value]);
    if (state.reverse) [s, t] = [t, s];
    if (!s || !t || s === t || !(v > 0)) return;
    const key = s + "\u0000" + t;
    let l = agg.get(key);
    if (!l) { l = { source: s, target: t, value: 0, rows: [] }; agg.set(key, l); }
    l.value += v; l.rows.push(ri);
  });
  const all = [...agg.values()];
  all.forEach((l) => { maxFlow = Math.max(maxFlow, l.value); });
  const thr = (state.s.minValue / 100) * maxFlow;
  return { links: all.filter((l) => l.value >= thr), total: all.length };
}

// ---------------------------------------------------------------------------------------------
// layout
// ---------------------------------------------------------------------------------------------
let G = null;   // current graph + helpers

function chartSize() {
  const stage = $("#stage");
  const W = state.s.width > 0 ? state.s.width : Math.max(560, stage.clientWidth - 36);
  const H = state.s.height > 0 ? state.s.height : Math.max(420, stage.clientHeight - 36);
  return { W: Math.round(W), H: Math.round(H) };
}

function layout() {
  const { links, total } = deriveLinks();
  const { W, H } = chartSize();
  G = { links, total, W, H, error: null };
  if (!links.length) { G.error = "No valid flows. Check the column mapping and that values are positive numbers."; return; }

  const names = [], idx = new Map();
  const id = (n) => { if (!idx.has(n)) { idx.set(n, names.length); names.push(n); } return idx.get(n); };
  const raw = links.map((l) => ({ source: id(l.source), target: id(l.target), value: l.value, ref: l }));
  const hasIn = names.map(() => false), hasOut = names.map(() => false);
  const inSum = names.map(() => 0), outSum = names.map(() => 0);
  raw.forEach((l) => { hasOut[l.source] = true; hasIn[l.target] = true; outSum[l.source] += l.value; inSum[l.target] += l.value; });
  Object.assign(G, { names, hasIn, hasOut, inSum, outSum });

  const s = state.s, fs = clamp(s.fontSize, 6, 40);
  let mL = 16, mR = 16;
  if (s.showLabels) {
    names.forEach((n, i) => {
      const w = (n.length + (s.showValues ? 2 + fmt(Math.max(inSum[i], outSum[i])).length : 0)) * fs * 0.58 + 14;
      const leftSide = s.labelPos === "left" || (s.labelPos === "auto" && !hasIn[i]);
      if (leftSide) mL = Math.max(mL, w); else if (!hasOut[i]) mR = Math.max(mR, w);
    });
    mL = Math.min(mL, W * 0.32); mR = Math.min(mR, W * 0.32);
  }
  const mT = 14 + (s.title ? 36 : 0) + (s.subtitle ? 22 : 0), mB = 14;
  const aligns = { justify: sankeyJustify, left: sankeyLeft, right: sankeyRight, center: sankeyCenter };
  const gen = sankey().nodeWidth(clamp(s.nodeWidth, 2, 100)).nodePadding(clamp(s.nodePadding, 0, 100))
    .nodeAlign(aligns[s.align] || sankeyJustify)
    .extent([[mL, mT], [Math.max(mL + 60, W - mR), Math.max(mT + 60, H - mB)]]);
  try {
    G.graph = gen({ nodes: names.map((name) => ({ name })), links: raw.map((l) => ({ ...l })) });
  } catch (e) {
    G.error = "Circular flows (A → B → A) are not supported. Check your data."; return;
  }
  G.gen = gen;
  G.graph.nodes.forEach((n) => {
    n.ax0 = n.x0; n.ay0 = n.y0;
    const m = state.manual[n.name];
    if (m) { const w = n.x1 - n.x0, h = n.y1 - n.y0; n.x0 = clamp(n.x0 + m.dx, 0, W - w); n.y0 = clamp(n.y0 + m.dy, 0, H - h); n.x1 = n.x0 + w; n.y1 = n.y0 + h; }
  });
  gen.update(G.graph);
}

// ---------------------------------------------------------------------------------------------
// drawing
// ---------------------------------------------------------------------------------------------
let linkSel, nodeSel, labelSel, gradSel, cropRect;
const linkPath = sankeyLinkHorizontal();
const baseOpacity = () => clamp(state.s.opacity, 5, 100) / 100;

function build() {
  const svg = select("#chart");
  svg.selectAll("*").remove();
  const { W, H } = G;
  svg.attr("width", W).attr("height", H).attr("viewBox", `0 0 ${W} ${H}`).classed("crop", state.cropMode);
  const s = state.s;
  if (s.bg === "white") svg.append("rect").attr("width", W).attr("height", H).attr("fill", "#ffffff");

  if (G.error) {
    svg.append("text").attr("x", W / 2).attr("y", H / 2).attr("text-anchor", "middle").attr("fill", "#6B7280").attr("font-size", 14).text(G.error);
    linkSel = nodeSel = labelSel = gradSel = null;
    return;
  }
  if (s.title) svg.append("text").attr("x", 24).attr("y", 36).attr("font-size", 21).attr("font-weight", 650).attr("fill", "#1F2937").text(s.title);
  if (s.subtitle) svg.append("text").attr("x", 24).attr("y", s.title ? 58 : 36).attr("font-size", 12.5).attr("fill", "#6B7280").text(s.subtitle);

  const { graph } = G;
  const col = (n) => colorOf(n.name, n.index);
  const defs = svg.append("defs");
  gradSel = null;
  if (s.linkStyle === "gradient") {
    gradSel = defs.selectAll("linearGradient").data(graph.links).join("linearGradient")
      .attr("gradientUnits", "userSpaceOnUse");
    gradSel.each(function (l) {
      const g = select(this);
      g.append("stop").attr("offset", "0%").attr("stop-color", col(l.source));
      g.append("stop").attr("offset", "100%").attr("stop-color", col(l.target));
    });
    gradSel.attr("id", (_, i) => `grad${i}`);
  }

  linkSel = svg.append("g").attr("fill", "none").selectAll("path").data(graph.links).join("path")
    .attr("stroke", (l, i) => s.linkStyle === "gradient" ? `url(#grad${i})` : s.linkStyle === "gray" ? "#B8BEC7" : col(l.source))
    .attr("stroke-opacity", baseOpacity()).style("cursor", "pointer")
    .on("mousemove", (ev, l) => {
      showTip(ev, `${l.source.name} → ${l.target.name}\n${fmt(l.value)}\n${(100 * l.value / Math.max(G.outSum[l.source.index], 1e-12)).toFixed(1)}% of ${l.source.name} · ${(100 * l.value / Math.max(G.inSum[l.target.index], 1e-12)).toFixed(1)}% of ${l.target.name}`);
      highlight((x) => x === l);
    })
    .on("mouseleave", () => { hideTip(); highlight(null); });

  nodeSel = svg.append("g").selectAll("rect").data(graph.nodes).join("rect")
    .attr("fill", col).style("cursor", "grab")
    .on("mousemove", (ev, n) => {
      showTip(ev, `${n.name}\nTotal: ${fmt(Math.max(G.inSum[n.index], G.outSum[n.index]))}\nIn: ${fmt(G.inSum[n.index])} · Out: ${fmt(G.outSum[n.index])}`);
      highlight((l) => l.source === n || l.target === n);
    })
    .on("mouseleave", () => { hideTip(); highlight(null); })
    .on("click", (ev, n) => { ev.stopPropagation(); selectNode(n.name); })
    .call(drag().filter((ev) => !state.cropMode && !ev.button)
      .on("start", function () { select(this).style("cursor", "grabbing"); hideTip(); })
      .on("drag", (ev, n) => {
        const w = n.x1 - n.x0, h = n.y1 - n.y0;
        const nx = clamp(n.x0 + ev.dx, 0, G.W - w), ny = clamp(n.y0 + ev.dy, 0, G.H - h);
        n.x0 = nx; n.x1 = nx + w; n.y0 = ny; n.y1 = ny + h;
        state.manual[n.name] = { dx: n.x0 - n.ax0, dy: n.y0 - n.ay0 };
        G.gen.update(G.graph); position();
      })
      .on("end", function () { select(this).style("cursor", "grab"); }));
  svg.on("click", () => { if (state.sel) selectNode(null); });

  if (s.showLabels) {
    labelSel = svg.append("g").attr("font-size", s.fontSize).attr("pointer-events", "none").selectAll("text").data(graph.nodes).join("text")
      .attr("dy", "0.35em")
      .attr("text-anchor", (n) => labelLeft(n) ? "end" : "start")
      .attr("fill", "#374151")
      .style("paint-order", "stroke").style("stroke", "#ffffff").style("stroke-width", "3px").style("stroke-linejoin", "round");
    labelSel.each(function (n) {
      const t = select(this);
      t.append("tspan").attr("font-weight", 600).text(n.name);
      if (s.showValues) t.append("tspan").attr("dx", 6).attr("fill", "#6B7280").text(fmt(Math.max(G.inSum[n.index], G.outSum[n.index])));
    });
  } else labelSel = null;

  // UI layer (not exported)
  const ui = svg.append("g").attr("class", "ui");
  cropRect = ui.append("rect").attr("fill", "rgba(31,111,181,.12)").attr("stroke", "#1f6fb5").attr("stroke-width", 1.5)
    .attr("stroke-dasharray", "6 4").attr("pointer-events", "none").style("display", "none");
  svg.on("pointerdown", cropStart).on("pointermove", cropMove).on("pointerup", cropEnd);
  drawCrop();
  position();
  applySelection();
}

function labelLeft(n) {
  const p = state.s.labelPos;
  return p === "left" || (p === "auto" && !G.hasIn[n.index]);
}

function position() {
  if (!linkSel) return;
  linkSel.attr("d", linkPath).attr("stroke-width", (l) => Math.max(1, l.width));
  if (gradSel) gradSel.attr("x1", (l) => l.source.x1).attr("x2", (l) => l.target.x0);
  nodeSel.attr("x", (n) => n.x0).attr("y", (n) => n.y0).attr("width", (n) => n.x1 - n.x0).attr("height", (n) => Math.max(1, n.y1 - n.y0));
  if (labelSel) labelSel.attr("x", (n) => labelLeft(n) ? n.x0 - 6 : n.x1 + 6).attr("y", (n) => (n.y0 + n.y1) / 2);
}

function highlight(pred) {
  if (!linkSel) return;
  const b = baseOpacity();
  linkSel.attr("stroke-opacity", (l) => !pred ? b : pred(l) ? Math.min(1, b * 1.5 + 0.15) : b * 0.15);
}

function applySelection() {
  if (!nodeSel) return;
  nodeSel.attr("stroke", (n) => n.name === state.sel ? "#111827" : "none").attr("stroke-width", 2);
}

const tip = $("#tip");
function showTip(ev, text) {
  if (state.cropMode) return;
  tip.textContent = text; tip.style.display = "block";
  const r = tip.getBoundingClientRect();
  tip.style.left = Math.min(ev.clientX + 14, innerWidth - r.width - 8) + "px";
  tip.style.top = Math.min(ev.clientY + 14, innerHeight - r.height - 8) + "px";
}
function hideTip() { tip.style.display = "none"; }

// ---------------------------------------------------------------------------------------------
// selected node panel
// ---------------------------------------------------------------------------------------------
function selectNode(name) {
  state.sel = name;
  const panel = $("#nodePanel");
  panel.hidden = !name;
  if (name) {
    $("#nodeName").textContent = name;
    const n = G && G.graph && G.graph.nodes.find((x) => x.name === name);
    $("#nodeColor").value = n ? toHex6(colorOf(name, n.index)) : "#1F4E79";
  }
  applySelection();
}
const toHex6 = (c) => /^#[0-9a-f]{6}$/i.test(c) ? c.toLowerCase() : "#1f4e79";

// ---------------------------------------------------------------------------------------------
// select area (crop)
// ---------------------------------------------------------------------------------------------
let cropDrag = null;
function cropStart(ev) {
  if (!state.cropMode) return;
  const [x, y] = pointer(ev, this);
  cropDrag = { x0: clamp(x, 0, G.W), y0: clamp(y, 0, G.H) };
  this.setPointerCapture(ev.pointerId);
  state.crop = null; drawCrop(cropDrag.x0, cropDrag.y0, 0, 0);
}
function cropMove(ev) {
  if (!cropDrag) return;
  const [x, y] = pointer(ev, this);
  const x1 = clamp(x, 0, G.W), y1 = clamp(y, 0, G.H);
  drawCrop(Math.min(cropDrag.x0, x1), Math.min(cropDrag.y0, y1), Math.abs(x1 - cropDrag.x0), Math.abs(y1 - cropDrag.y0));
}
function cropEnd(ev) {
  if (!cropDrag) return;
  const [x, y] = pointer(ev, this);
  const x1 = clamp(x, 0, G.W), y1 = clamp(y, 0, G.H);
  const r = { x: Math.min(cropDrag.x0, x1), y: Math.min(cropDrag.y0, y1), w: Math.abs(x1 - cropDrag.x0), h: Math.abs(y1 - cropDrag.y0) };
  cropDrag = null;
  state.crop = r.w > 12 && r.h > 12 ? r : null;
  drawCrop();
  setCropMode(false);
  if (state.crop) toast(`Area selected (${Math.round(r.w)}×${Math.round(r.h)} px). Copy or export now uses this area.`);
}
function drawCrop(x, y, w, h) {
  if (!cropRect) return;
  const c = x === undefined ? state.crop : { x, y, w, h };
  if (!c) { cropRect.style("display", "none"); } else cropRect.style("display", null).attr("x", c.x).attr("y", c.y).attr("width", c.w).attr("height", c.h);
  $("#btnCropClear").disabled = !state.crop;
}
function setCropMode(on) {
  state.cropMode = on;
  $("#btnCrop").classList.toggle("on", on);
  select("#chart").classed("crop", on);
  if (on) { hideTip(); toast("Drag a rectangle on the chart to select the area."); }
}

// ---------------------------------------------------------------------------------------------
// export
// ---------------------------------------------------------------------------------------------
function svgMarkup(scale) {
  const el = $("#chart").cloneNode(true);
  el.querySelectorAll(".ui").forEach((e) => e.remove());
  el.removeAttribute("class");
  el.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  el.setAttribute("font-family", "Segoe UI, Arial, sans-serif");
  el.setAttribute("width", G.W * scale); el.setAttribute("height", G.H * scale);
  return new XMLSerializer().serializeToString(el);
}

async function chartBlob(scale) {
  if (!G || G.error) throw new Error("Nothing to export");
  const img = new Image();
  await new Promise((res, rej) => {
    img.onload = res; img.onerror = () => rej(new Error("render failed"));
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgMarkup(scale));
  });
  const c = state.crop ? { x: state.crop.x, y: state.crop.y, w: state.crop.w, h: state.crop.h } : { x: 0, y: 0, w: G.W, h: G.H };
  const cv = document.createElement("canvas");
  cv.width = Math.round(c.w * scale); cv.height = Math.round(c.h * scale);
  cv.getContext("2d").drawImage(img, c.x * scale, c.y * scale, c.w * scale, c.h * scale, 0, 0, cv.width, cv.height);
  return new Promise((res, rej) => cv.toBlob((b) => b ? res(b) : rej(new Error("png failed")), "image/png"));
}

function download(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
const baseName = () => (state.s.fileName || "sankey").replace(/[\\/:*?"<>|]+/g, "_");

async function copyPng() {
  try {
    const item = new ClipboardItem({ "image/png": chartBlob(+state.s.scale || 2) });
    await navigator.clipboard.write([item]);
    toast("Copied to clipboard ✓ — paste it into PowerPoint, Word, Teams...");
  } catch (e) {
    console.error(e);
    toast("Copy is blocked in this browser/context. Use Export PNG instead.");
  }
}
async function exportPng() {
  try { download(await chartBlob(+state.s.scale || 2), baseName() + ".png"); toast("PNG exported"); }
  catch (e) { toast("Could not export: " + e.message); }
}
function exportSvg() {
  if (!G || G.error) return toast("Nothing to export");
  download(new Blob([svgMarkup(1)], { type: "image/svg+xml" }), baseName() + ".svg");
  toast("SVG exported");
}

// ---------------------------------------------------------------------------------------------
// data loading + editor
// ---------------------------------------------------------------------------------------------
function setTable(rows2d, label, keepText) {
  if (!rows2d || rows2d.length < 2) { toast("The file needs a header row and at least one data row."); return false; }
  const d = detectTable(rows2d);
  if (!d.rows.length) { toast("No data rows found."); return false; }
  state.table = { headers: d.headers, rows: d.rows };
  state.map = d.map;
  state.reverse = false; state.manual = {}; state.nodeColors = {}; state.crop = null; state.sel = null;
  state.s.minValue = 0;
  if (!keepText && state.s.title === SAMPLE_TITLE) { state.s.title = ""; state.s.subtitle = ""; }
  state.fileLabel = label;
  syncControls();
  refreshData();
  return true;
}

function loadSample() {
  state.table = { headers: [...SAMPLE.headers], rows: SAMPLE.rows.map((r) => [...r]) };
  state.map = { from: 0, to: 1, value: 2 };
  state.reverse = false; state.manual = {}; state.nodeColors = {}; state.crop = null; state.sel = null; state.sheets = null;
  state.s.title = SAMPLE_TITLE; state.s.subtitle = SAMPLE_SUB; state.s.minValue = 0;
  state.fileLabel = "Sample data";
  syncControls(); refreshData();
}

async function loadFile(file) {
  try {
    const ext = (file.name.split(".").pop() || "").toLowerCase();
    if (ext === "xlsx" || ext === "xlsm") {
      state.sheets = await readXlsx(await file.arrayBuffer());
      const best = Math.max(0, state.sheets.findIndex((s) => s.rows.length > 1));
      fillSheets(best);
      setTable(state.sheets[best].rows, file.name);
    } else if (ext === "xls") {
      toast("Old .xls files are not supported. In Excel use Save As > .xlsx or .csv.");
    } else {
      state.sheets = null;
      setTable(parseDelimited(await file.text()), file.name);
    }
    if (state.sheets) $("#sheetRow").hidden = state.sheets.length < 2; else $("#sheetRow").hidden = true;
  } catch (e) {
    console.error(e);
    toast("Could not read the file: " + e.message);
  }
}
function fillSheets(sel) {
  const s = $("#sheetSel");
  s.innerHTML = "";
  state.sheets.forEach((sh, i) => s.add(new Option(`${sh.name} (${sh.rows.length} rows)`, i)));
  s.value = sel;
}

function fillMapping() {
  for (const [id, key] of [["#mapFrom", "from"], ["#mapTo", "to"], ["#mapVal", "value"]]) {
    const sel = $(id);
    sel.innerHTML = "";
    state.table.headers.forEach((h, i) => sel.add(new Option(h, i)));
    sel.value = state.map[key];
  }
}

function renderTable() {
  const t = $("#dataTable");
  t.innerHTML = "";
  const head = t.createTHead().insertRow();
  head.insertCell().outerHTML = "<th></th>";
  state.table.headers.forEach((h) => { const th = document.createElement("th"); th.textContent = h; head.appendChild(th); });
  const body = t.createTBody();
  const limit = 300;
  state.table.rows.slice(0, limit).forEach((r, ri) => {
    const tr = body.insertRow();
    const x = tr.insertCell(); x.className = "x"; x.textContent = "×"; x.title = "Delete row";
    x.onclick = () => { state.table.rows.splice(ri, 1); refreshData(); };
    state.table.headers.forEach((_, ci) => {
      const td = tr.insertCell();
      td.contentEditable = "true"; td.textContent = r[ci] ?? "";
      td.addEventListener("input", () => { state.table.rows[ri][ci] = td.textContent; refreshChart(); });
    });
  });
  $("#tblNote").textContent = state.table.rows.length > limit ? `Showing first ${limit} of ${state.table.rows.length} rows` : `${state.table.rows.length} rows`;
}

// ---------------------------------------------------------------------------------------------
// refresh
// ---------------------------------------------------------------------------------------------
function updateInfo() {
  const n = G && G.names ? G.names.length : 0;
  const hidden = G && G.total ? G.total - G.links.length : 0;
  $("#dataInfo").textContent = `${state.fileLabel} · ${state.table.rows.length} rows · ${G ? G.links.length : 0} flows · ${n} nodes` + (hidden ? ` · ${hidden} hidden` : "");
}

function refreshChart() {
  raf(() => {
    layout(); build(); updateInfo();
    if (state.sel && !(G.graph && G.graph.nodes.some((n) => n.name === state.sel))) selectNode(null);
  });
}
function refreshData() { fillMapping(); renderTable(); refreshChart(); $("#btnSwap").classList.toggle("on", state.reverse); }

function syncControls() {
  document.querySelectorAll("[data-k]").forEach((inp) => {
    const v = state.s[inp.dataset.k];
    if (inp.type === "checkbox") inp.checked = !!v; else inp.value = v;
    syncOut(inp);
  });
  $("#customRow").hidden = state.s.palette !== "custom";
  $("#nodePanel").hidden = !state.sel;
}
function syncOut(inp) {
  const o = document.querySelector(`output[data-for="${inp.dataset.k}"]`);
  if (!o) return;
  const k = inp.dataset.k;
  o.textContent = k === "minValue" || k === "opacity" ? inp.value + "%" : inp.value;
}

// ---------------------------------------------------------------------------------------------
// wiring
// ---------------------------------------------------------------------------------------------
function init() {
  document.querySelectorAll("[data-k]").forEach((inp) => {
    inp.addEventListener("input", () => {
      const k = inp.dataset.k;
      state.s[k] = inp.type === "checkbox" ? inp.checked : (inp.type === "range" || inp.type === "number") ? +inp.value : inp.value;
      syncOut(inp);
      if (k === "palette") $("#customRow").hidden = inp.value !== "custom";
      if (k === "scale" || k === "fileName") return;
      refreshChart();
    });
  });
  $("#mapFrom").onchange = (e) => { state.map.from = +e.target.value; refreshChart(); };
  $("#mapTo").onchange = (e) => { state.map.to = +e.target.value; refreshChart(); };
  $("#mapVal").onchange = (e) => { state.map.value = +e.target.value; refreshChart(); };
  $("#btnSwap").onclick = () => { state.reverse = !state.reverse; state.manual = {}; $("#btnSwap").classList.toggle("on", state.reverse); refreshChart(); };
  $("#sheetSel").onchange = (e) => { const i = +e.target.value; setTable(state.sheets[i].rows, state.fileLabel.split(" · ")[0], true); };

  const fileInput = $("#file");
  fileInput.onchange = () => { if (fileInput.files[0]) loadFile(fileInput.files[0]); fileInput.value = ""; };
  $("#btnUpload").onclick = () => fileInput.click();
  const drop = $("#drop");
  ["dragenter", "dragover"].forEach((t) => window.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((t) => window.addEventListener(t, (e) => { e.preventDefault(); if (t === "drop" || e.target === document.documentElement) drop.classList.remove("over"); }));
  window.addEventListener("drop", (e) => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f) loadFile(f); });

  $("#btnSample").onclick = loadSample;
  const dlg = $("#pasteDlg");
  $("#btnPaste").onclick = () => { $("#pasteText").value = ""; dlg.showModal(); $("#pasteText").focus(); };
  $("#pasteCancel").onclick = () => dlg.close();
  $("#pasteOk").onclick = () => { if (setTable(parseDelimited($("#pasteText").value), "Pasted data")) { state.sheets = null; $("#sheetRow").hidden = true; dlg.close(); } };
  $("#btnAddRow").onclick = () => { state.table.rows.push(state.table.headers.map(() => "")); renderTable(); refreshChart(); };

  $("#btnResetPos").onclick = () => { state.manual = {}; refreshChart(); toast("Positions reset"); };
  $("#nodeColor").oninput = (e) => { if (state.sel) { state.nodeColors[state.sel] = e.target.value.toUpperCase(); refreshChart(); } };
  $("#btnNodeColorReset").onclick = () => { if (state.sel) { delete state.nodeColors[state.sel]; refreshChart(); setTimeout(() => selectNode(state.sel), 60); } };
  $("#btnNodePosReset").onclick = () => { if (state.sel) { delete state.manual[state.sel]; refreshChart(); } };
  $("#btnNodeDeselect").onclick = () => selectNode(null);

  $("#btnCrop").onclick = () => setCropMode(!state.cropMode);
  $("#btnCropClear").onclick = () => { state.crop = null; drawCrop(); toast("Selection cleared"); };
  for (const id of ["#btnCopy", "#btnCopy2"]) $(id).onclick = copyPng;
  for (const id of ["#btnPng", "#btnPng2"]) $(id).onclick = exportPng;
  for (const id of ["#btnSvg", "#btnSvg2"]) $(id).onclick = exportSvg;

  window.addEventListener("keydown", (e) => {
    const a = document.activeElement, editing = a && (a.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName));
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "c" && !editing && !String(getSelection())) { e.preventDefault(); copyPng(); }
    if (e.key === "Escape") { if (state.cropMode) setCropMode(false); else if (state.crop) { state.crop = null; drawCrop(); } else if (state.sel) selectNode(null); }
  });
  let rt; window.addEventListener("resize", () => { clearTimeout(rt); rt = setTimeout(() => { if (!state.s.width || !state.s.height) refreshChart(); }, 150); });

  loadSample();
}

init();
window.SankeyApp = { state, loadFile, setTable, parseDelimited, readXlsx, get G() { return G; }, refresh: refreshChart };
