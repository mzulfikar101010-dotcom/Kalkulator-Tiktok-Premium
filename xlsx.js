/* MiniXlsx — membuat dan membaca file Excel (.xlsx) tanpa library luar, jalan offline. */
const MiniXlsx = (function () {
  const enc = new TextEncoder(), dec = new TextDecoder();

  /* ---------- Bantuan ---------- */
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  const esc = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const unesc = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, "&");
  const attrs = (tag) => { const o = {}; tag.replace(/([\w:.-]+)\s*=\s*"([^"]*)"/g, (m, k, v) => { o[k] = unesc(v); return m; }); return o; };
  function colName(i) { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }

  /* ---------- Menulis ZIP (tanpa kompresi) ---------- */
  function zip(files) {
    const parts = [], cd = []; let off = 0;
    const d = new Date();
    const dosDate = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    const dosTime = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    files.forEach(f => {
      const name = enc.encode(f.name), data = f.data, crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), name, data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true);
      ch.setUint16(12, dosTime, true); ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true);
      ch.setUint32(42, off, true);
      cd.push(new Uint8Array(ch.buffer), name);
      off += 30 + name.length + data.length;
    });
    let cdSize = 0; cd.forEach(x => cdSize += x.length);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, off, true);
    const all = [...parts, ...cd, new Uint8Array(end.buffer)];
    let total = 0; all.forEach(x => total += x.length);
    const out = new Uint8Array(total); let p = 0; all.forEach(x => { out.set(x, p); p += x.length; });
    return out;
  }

  /* ---------- Membuat .xlsx ----------
     sheets: [{ name, cols: [lebar...], rows: [[sel, ...], ...] }]
     sel: teks | angka | { v: teks/angka, s: gaya }   gaya: 1 tebal, 2 rupiah, 3 tebal+rupiah, 4 judul kolom, 5 persen */
  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;Rp&quot; #,##0;[Red]-&quot;Rp&quot; #,##0"/><numFmt numFmtId="165" formatCode="0.0&quot;%&quot;"/></numFmts>' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF3ECD9"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function sheetXml(sh) {
    let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    if (sh.cols) x += "<cols>" + sh.cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("") + "</cols>";
    x += "<sheetData>";
    sh.rows.forEach((row, ri) => {
      let cells = "";
      row.forEach((c, ci) => {
        if (c == null || c === "") return;
        const o = typeof c === "object" ? c : { v: c };
        if (o.v == null || o.v === "") return;
        const ref = colName(ci) + (ri + 1), st = o.s ? ` s="${o.s}"` : "";
        if (typeof o.v === "number" && isFinite(o.v)) cells += `<c r="${ref}"${st}><v>${o.v}</v></c>`;
        else cells += `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${esc(o.v)}</t></is></c>`;
      });
      x += `<row r="${ri + 1}">${cells}</row>`;
    });
    return x + "</sheetData></worksheet>";
  }

  function buat(sheets) {
    const H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const files = [
      { name: "[Content_Types].xml", text: H + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: "_rels/.rels", text: H + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: "xl/workbook.xml", text: H + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
        sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") + "</sheets></workbook>" },
      { name: "xl/_rels/workbook.xml.rels", text: H + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") +
        `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
      { name: "xl/styles.xml", text: STYLES }
    ];
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, text: sheetXml(s) }));
    return zip(files.map(f => ({ name: f.name, data: enc.encode(f.text) })));
  }

  /* ---------- Membaca .xlsx ---------- */
  async function inflate(raw) {
    if (typeof DecompressionStream === "undefined") throw new Error("Browser ini belum bisa membaca Excel. Perbarui browser/iOS-mu.");
    const stream = new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function bukaZip(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(buf);
    let e = u8.length - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error("File ini bukan Excel (.xlsx) yang valid.");
    const n = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true);
    const map = {};
    for (let i = 0; i < n; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      const start = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
      map[name] = { method, raw: u8.subarray(start, start + csize) };
      p += 46 + nlen + xlen + clen;
    }
    return map;
  }

  async function ambilTeks(map, name) {
    const f = map[name];
    if (!f) return null;
    if (f.method === 0) return dec.decode(f.raw);
    if (f.method === 8) return dec.decode(await inflate(f.raw));
    throw new Error("Format kompresi file Excel tidak didukung.");
  }

  function uraiSheet(xml, shared) {
    const rows = [], re = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = re.exec(xml))) {
      const a = attrs(m[1]);
      const mm = /^([A-Z]+)(\d+)$/.exec(a.r || "");
      if (!mm) continue;
      let ci = 0; for (const ch of mm[1]) ci = ci * 26 + (ch.charCodeAt(0) - 64);
      ci--; const ri = +mm[2] - 1;
      const inner = m[3] || ""; let val = null;
      if (a.t === "inlineStr") {
        val = [...inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x => unesc(x[1])).join("");
      } else {
        const v = /<v>([\s\S]*?)<\/v>/.exec(inner);
        if (v) {
          if (a.t === "s") val = shared[+v[1]];
          else if (a.t === "str" || a.t === "e") val = unesc(v[1]);
          else if (a.t === "b") val = v[1] === "1";
          else val = Number(v[1]);
        }
      }
      if (val === null || val === undefined || val === "") continue;
      (rows[ri] = rows[ri] || [])[ci] = val;
    }
    return Array.from(rows, r => r ? Array.from(r, v => v === undefined ? null : v) : []);
  }

  async function baca(buf) {
    const map = bukaZip(buf);
    const wbXml = await ambilTeks(map, "xl/workbook.xml");
    if (!wbXml) throw new Error("File ini bukan Excel (.xlsx) yang valid.");
    const relXml = (await ambilTeks(map, "xl/_rels/workbook.xml.rels")) || "";
    const rels = {};
    (relXml.match(/<Relationship\b[^>]*>/g) || []).forEach(t => { const a = attrs(t); if (a.Id) rels[a.Id] = a.Target; });
    const ssXml = await ambilTeks(map, "xl/sharedStrings.xml");
    const shared = [];
    if (ssXml) {
      (ssXml.match(/<si\b[^>]*>[\s\S]*?<\/si>/g) || []).forEach(si => {
        const clean = si.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
        shared.push([...clean.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x => unesc(x[1])).join(""));
      });
    }
    const sheets = [];
    const tags = wbXml.match(/<sheet\b[^>]*>/g) || [];
    for (let i = 0; i < tags.length; i++) {
      const a = attrs(tags[i]);
      let target = rels[a["r:id"]] || `worksheets/sheet${i + 1}.xml`;
      target = target.startsWith("/") ? target.slice(1) : "xl/" + target;
      const xml = await ambilTeks(map, target);
      sheets.push({ name: a.name || "Sheet" + (i + 1), rows: xml ? uraiSheet(xml, shared) : [] });
    }
    return { sheets };
  }

  return { buat, baca };
})();
if (typeof module !== "undefined") module.exports = MiniXlsx;
