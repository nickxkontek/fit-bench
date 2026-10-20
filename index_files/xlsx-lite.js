/* xlsx-lite: a small, dependency-free Excel (.xlsx) writer and reader for Fit Bench.
   Writer: builds a valid workbook (numbers, text, formulas, bold header row, column widths).
   Reader: returns the first worksheet as rows of strings (uses the browser's built-in deflate). */
(function (global) {
  "use strict";

  /* ---------- ZIP (store, no compression) ---------- */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(bytes) { let c = 0xFFFFFFFF; for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function zip(files) { // files: [{name, data:Uint8Array}]
    const enc = new TextEncoder(), parts = [], central = []; let offset = 0;
    for (const f of files) {
      const name = enc.encode(f.name), crc = crc32(f.data), size = f.data.length;
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, 0, true); h.setUint16(12, 0x21, true); h.setUint32(14, crc, true); h.setUint32(18, size, true);
      h.setUint32(22, size, true); h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, f.data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint16(10, 0, true); c.setUint16(12, 0, true); c.setUint16(14, 0x21, true); c.setUint32(16, crc, true);
      c.setUint32(20, size, true); c.setUint32(24, size, true); c.setUint16(28, name.length, true);
      c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + size;
    }
    const cdSize = central.reduce((s, p) => s + p.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cdSize, true); e.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(e.buffer)], out = new Uint8Array(all.reduce((s, p) => s + p.length, 0));
    let p = 0; for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  /* ---------- Writer ---------- */
  const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  function colName(i) { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }

  /* sheet = {name, rows:[[cell]], widths:[n], headerRow:bool}
     cell = number | string | null | {v:number, f:"FORMULA"} */
  function sheetXML(sheet) {
    const rows = sheet.rows.map((row, r) => {
      const cells = row.map((v, c) => {
        if (v === null || v === undefined || v === "") return "";
        const ref = colName(c) + (r + 1), st = sheet.headerRow && r === 0 ? ' s="1"' : "";
        if (typeof v === "object" && v.f) {
          const val = typeof v.v === "number" && isFinite(v.v) ? `<v>${v.v}</v>` : "";
          return `<c r="${ref}"${st}><f>${xmlEsc(v.f)}</f>${val}</c>`;
        }
        if (typeof v === "number") return isFinite(v) ? `<c r="${ref}"${st}><v>${v}</v></c>` : "";
        return `<c r="${ref}" t="inlineStr"${st}><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
      }).join("");
      return `<row r="${r + 1}">${cells}</row>`;
    }).join("");
    const cols = sheet.widths && sheet.widths.length ? `<cols>${sheet.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>` : "";
    const freeze = sheet.headerRow ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${freeze}${cols}<sheetData>${rows}</sheetData></worksheet>`;
  }
  function write(sheets) {
    const enc = new TextEncoder(), F = [], add = (name, s) => F.push({ name, data: enc.encode(s) });
    add("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`);
    add("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    add("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xmlEsc(s.name.slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
    add("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    add("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
    sheets.forEach((s, i) => add(`xl/worksheets/sheet${i + 1}.xml`, sheetXML(s)));
    return zip(F);
  }

  /* ---------- Reader ---------- */
  async function unzip(buf) {
    const b = new Uint8Array(buf), dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    let e = -1; for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
    if (e < 0) throw new Error("Not a valid .xlsx file.");
    const count = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true);
    const dec = new TextDecoder(), out = {};
    for (let n = 0; n < count; n++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), lo = dv.getUint32(p + 42, true);
      const name = dec.decode(b.subarray(p + 46, p + 46 + nlen));
      p += 46 + nlen + xlen + clen;
      const start = lo + 30 + dv.getUint16(lo + 26, true) + dv.getUint16(lo + 28, true);
      out[name] = { method, raw: b.subarray(start, start + csize) };
    }
    return {
      async text(name) {
        const f = out[name]; if (!f) return null;
        if (f.method === 0) return dec.decode(f.raw);
        if (f.method !== 8 || typeof DecompressionStream === "undefined") throw new Error("This .xlsx uses a compression this app cannot read. Save it as CSV instead.");
        const s = new Blob([f.raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
        return await new Response(s).text();
      }, names: Object.keys(out)
    };
  }
  const colIndex = ref => { const m = /^([A-Z]+)/.exec(ref); let n = 0; for (const ch of m[1]) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  /* Returns {name, rows} for the first sheet that holds at least two columns of numbers
     (falls back to the first sheet), so a workbook with a summary sheet in front still imports its data. */
  async function readDataSheet(buf) {
    const z = await unzip(buf), P = new DOMParser(), xml = s => P.parseFromString(s, "application/xml");
    const wbTxt = await z.text("xl/workbook.xml"); if (!wbTxt) throw new Error("Not a valid .xlsx file.");
    const sheets = [...xml(wbTxt).getElementsByTagName("sheet")]; if (!sheets.length) throw new Error("The workbook has no sheets.");
    const rels = {}, relsTxt = await z.text("xl/_rels/workbook.xml.rels");
    if (relsTxt) for (const r of xml(relsTxt).getElementsByTagName("Relationship")) { const t = r.getAttribute("Target"); rels[r.getAttribute("Id")] = t.startsWith("/") ? t.slice(1) : "xl/" + t.replace(/^\.\//, ""); }
    const shared = []; const ssTxt = await z.text("xl/sharedStrings.xml");
    if (ssTxt) for (const si of xml(ssTxt).getElementsByTagName("si")) shared.push([...si.getElementsByTagName("t")].map(t => t.textContent).join(""));
    let firstResult = null, firstData = null;
    for (let i = 0; i < sheets.length; i++) {
      const sh = sheets[i];
      const rid = sh.getAttribute("r:id") || sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      const path = rels[rid] || `xl/worksheets/sheet${i + 1}.xml`;
      const txt = await z.text(path); if (!txt) continue;
      const rows = sheetRows(xml(txt), shared), res = { name: sh.getAttribute("name") || `Sheet ${i + 1}`, rows };
      if (!firstResult) firstResult = res;
      const numericRows = rows.filter(r => r.filter(c => c !== "" && isFinite(Number(c))).length >= 2).length;
      if (numericRows >= 2) { if (/^data$/i.test(res.name)) return res; if (!firstData) firstData = res; }
    }
    if (!firstResult) throw new Error("The workbook has no readable sheets.");
    return firstData || firstResult;
  }
  function sheetRows(doc, shared) {
    const rows = [];
    for (const row of doc.getElementsByTagName("row")) {
      const r = []; let next = 0;
      for (const c of row.getElementsByTagName("c")) {
        const ref = c.getAttribute("r"), i = ref ? colIndex(ref) : next; next = i + 1;
        const t = c.getAttribute("t"), v = c.getElementsByTagName("v")[0];
        let val = "";
        if (t === "s" && v) val = shared[+v.textContent] ?? "";
        else if (t === "inlineStr") val = [...c.getElementsByTagName("t")].map(x => x.textContent).join("");
        else if (v) val = v.textContent;
        r[i] = val;
      }
      rows.push(Array.from(r, x => x ?? ""));
    }
    return rows;
  }
  async function readFirstSheet(buf) { return (await readDataSheet(buf)).rows; }

  global.XlsxLite = { write, readDataSheet, readFirstSheet, colName };
})(typeof self !== "undefined" ? self : this);
