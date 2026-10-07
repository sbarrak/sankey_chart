// File readers: delimited text and XLSX (no external libraries; uses the browser's DecompressionStream).

export function parseNum(x) {
  if (typeof x === "number") return x;
  let s = String(x ?? "").trim().replace(/[\s €$£]/g, "");
  if (!s) return NaN;
  const c = s.lastIndexOf(","), d = s.lastIndexOf(".");
  if (c > -1 && d > -1) {
    s = c > d ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  } else if (c > -1) {
    s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, "") : s.replace(",", ".");
  }
  return /^-?\d*\.?\d+(e[+-]?\d+)?$/i.test(s) ? parseFloat(s) : NaN;
}

/** Delimited text -> 2D array of strings. Delimiter (comma, semicolon, tab) is auto-detected. */
export function parseDelimited(text) {
  text = text.replace(/^﻿/, "");
  const firstLine = text.split(/\r?\n/, 1)[0] || "";
  const count = (ch) => firstLine.split(ch).length - 1;
  const cand = [["\t", count("\t")], [";", count(";")], [",", count(",")]].sort((a, b) => b[1] - a[1]);
  const delim = cand[0][1] > 0 ? cand[0][0] : ",";
  const rows = [];
  let row = [], cell = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cell += '"'; i++; } else q = false;
      } else cell += ch;
    } else if (ch === '"') q = true;
    else if (ch === delim) { row.push(cell); cell = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(cell); cell = "";
      if (row.some((c) => c.trim() !== "")) rows.push(row);
      row = [];
    } else cell += ch;
  }
  row.push(cell);
  if (row.some((c) => c.trim() !== "")) rows.push(row);
  return rows;
}

// ---- XLSX -------------------------------------------------------------------------------------
async function unzip(buf) {
  const dv = new DataView(buf), u8 = new Uint8Array(buf);
  let e = u8.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error("This is not an .xlsx file (if it is an old .xls, save it as .xlsx or .csv).");
  const n = dv.getUint16(e + 10, true), cd = dv.getUint32(e + 16, true);
  const files = {};
  let p = cd;
  for (let i = 0; i < n; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    const lho = dv.getUint32(p + 42, true);
    files[new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nl))] = { method, csize, lho };
    p += 46 + nl + el + cl;
  }
  const read = async (name) => {
    const f = files[name];
    if (!f) return null;
    const nl = dv.getUint16(f.lho + 26, true), el = dv.getUint16(f.lho + 28, true);
    const start = f.lho + 30 + nl + el;
    const data = u8.subarray(start, start + f.csize);
    if (f.method === 0) return new TextDecoder().decode(data);
    const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new TextDecoder().decode(await new Response(stream).arrayBuffer());
  };
  return { read, has: (n2) => !!files[n2] };
}

const xml = (s) => new DOMParser().parseFromString(s, "application/xml");
const colIndex = (ref) => {
  let n = 0;
  for (const ch of ref.replace(/[0-9]/g, "")) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};

/** -> [{name, rows}] (rows = 2D array of strings/numbers) */
export async function readXlsx(buf) {
  const z = await unzip(buf);
  const wbXml = await z.read("xl/workbook.xml");
  if (!wbXml) throw new Error("Could not find a workbook inside this file.");
  const rels = {};
  const relXml = await z.read("xl/_rels/workbook.xml.rels");
  if (relXml) for (const r of xml(relXml).getElementsByTagName("Relationship")) rels[r.getAttribute("Id")] = r.getAttribute("Target");

  const sst = [];
  const sstXml = await z.read("xl/sharedStrings.xml");
  if (sstXml) for (const si of xml(sstXml).getElementsByTagName("si")) {
    sst.push([...si.getElementsByTagName("t")].map((t) => t.textContent).join(""));
  }

  const sheets = [];
  for (const sh of xml(wbXml).getElementsByTagName("sheet")) {
    const rid = sh.getAttribute("r:id") || sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    let target = rels[rid];
    if (!target) continue;
    target = target.replace(/^\//, "");
    if (!target.startsWith("xl/")) target = "xl/" + target;
    const sx = await z.read(target);
    if (!sx) continue;
    const rows = [];
    for (const rEl of xml(sx).getElementsByTagName("row")) {
      const r = [];
      for (const c of rEl.getElementsByTagName("c")) {
        const t = c.getAttribute("t");
        const vEl = c.getElementsByTagName("v")[0];
        let val = "";
        if (t === "inlineStr") val = [...c.getElementsByTagName("t")].map((x) => x.textContent).join("");
        else if (vEl) {
          val = vEl.textContent;
          if (t === "s") val = sst[+val] ?? "";
          else if (t === "b" || t === "str" || t === "e") val = String(val);
          else if (val !== "" && !isNaN(+val)) val = +val;
        }
        r[colIndex(c.getAttribute("r") || "A1")] = val;
      }
      for (let i = 0; i < r.length; i++) if (r[i] === undefined) r[i] = "";
      if (r.some((x) => String(x).trim() !== "")) rows.push(r);
    }
    sheets.push({ name: sh.getAttribute("name"), rows });
  }
  if (!sheets.length) throw new Error("No worksheets found in this file.");
  return sheets;
}

/** Guess header row + which columns are From / To / Value. */
export function detectTable(rows2d) {
  const width = Math.max(...rows2d.map((r) => r.length));
  const rows = rows2d.map((r) => Array.from({ length: width }, (_, i) => r[i] ?? ""));
  const body = rows.slice(1);
  const numRatio = (col, list) => {
    const v = list.map((r) => r[col]).filter((x) => String(x).trim() !== "");
    return v.length ? v.filter((x) => !isNaN(parseNum(x))).length / v.length : 0;
  };
  let valueCol = -1;
  for (let c = width - 1; c >= 0; c--) if (numRatio(c, body) >= 0.8) { valueCol = c; break; }
  if (valueCol < 0) valueCol = Math.min(2, width - 1);
  const hasHeader = isNaN(parseNum(rows[0][valueCol]));
  const headers = hasHeader
    ? rows[0].map((h, i) => String(h).trim() || `Column ${i + 1}`)
    : rows[0].map((_, i) => `Column ${i + 1}`);
  const data = hasHeader ? body : rows;

  const find = (re) => headers.findIndex((h) => re.test(h));
  let from = find(/^(from|source|origin|start|src|de)\b/i);
  let to = find(/^(to|target|dest|destination|end|tgt|para)\b/i);
  const val = find(/^(value|amount|weight|flow|qty|quantity|count|sum|total|size)\b/i);
  if (val >= 0) valueCol = val;
  const others = headers.map((_, i) => i).filter((i) => i !== valueCol);
  if (from < 0 || from === valueCol) from = others[0] ?? 0;
  if (to < 0 || to === valueCol || to === from) to = others.find((i) => i !== from) ?? Math.min(1, width - 1);
  return { headers, rows: data, map: { from, to, value: valueCol } };
}
