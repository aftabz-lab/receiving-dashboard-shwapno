/* ---------------------------------------------------------------------------
   xlsx-lite.js
   Minimal, dependency-free multi-sheet workbook writer for the receiving
   dashboard. A CSV file can only ever hold one sheet, so the incident export
   ("Visible CSV" on the user-incident table) is written as a real .xlsx
   workbook with one sheet per table.

   No external library, no CDN: the file is a ZIP container built here, using
   the browser's own CompressionStream when it exists and stored entries when
   it does not.
--------------------------------------------------------------------------- */

const textEncoder = new TextEncoder();

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    table[index] = value >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    crc = CRC_TABLE[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;")
    // Excel rejects most control characters inside a shared or inline string.
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}

function columnLetter(index) {
  let letter = "";
  let cursor = index + 1;
  while (cursor > 0) {
    const remainder = (cursor - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    cursor = Math.floor((cursor - remainder) / 26);
  }
  return letter;
}

function sanitizeSheetName(name, fallback) {
  const cleaned = String(name || fallback).replace(/[\\/*?:[\]]/g, " ").trim() || fallback;
  return cleaned.slice(0, 31);
}

function cellXml(value, reference, headerRow) {
  const style = headerRow ? ' s="1"' : "";
  if (value == null || value === "") return `<c r="${reference}"${style}/>`;
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${reference}"${style}><v>${value}</v></c>`;
  }
  return `<c r="${reference}"${style} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function columnWidths(rows) {
  const widths = [];
  rows.slice(0, 400).forEach(row => {
    row.forEach((value, index) => {
      const length = String(value ?? "").length;
      widths[index] = Math.max(widths[index] || 10, Math.min(48, length + 3));
    });
  });
  return widths;
}

function sheetXml(rows) {
  const widths = columnWidths(rows);
  const cols = widths.length
    ? `<cols>${widths.map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`).join("")}</cols>`
    : "";
  const body = rows.map((row, rowIndex) => {
    const cells = row.map((value, columnIndex) => cellXml(value, `${columnLetter(columnIndex)}${rowIndex + 1}`, rowIndex === 0)).join("");
    return `<row r="${rowIndex + 1}">${cells}</row>`;
  }).join("");
  const freeze = rows.length > 1
    ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${freeze}${cols}<sheetData>${body}</sheetData><autoFilter ref="A1:${columnLetter(Math.max(0, (rows[0]?.length || 1) - 1))}${Math.max(1, rows.length)}"/></worksheet>`;
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF405AA7"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

async function deflateRaw(bytes) {
  if (typeof CompressionStream !== "function") return null;
  try {
    const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream("deflate-raw"));
    const compressed = new Uint8Array(await new Response(stream).arrayBuffer());
    return compressed.length < bytes.length ? compressed : null;
  } catch {
    return null;
  }
}

function dosDateTime(date) {
  const time = ((date.getHours() & 0x1f) << 11) | ((date.getMinutes() & 0x3f) << 5) | ((date.getSeconds() / 2) & 0x1f);
  const day = (((date.getFullYear() - 1980) & 0x7f) << 9) | (((date.getMonth() + 1) & 0x0f) << 5) | (date.getDate() & 0x1f);
  return { time, day };
}

async function zip(entries) {
  const stamp = dosDateTime(new Date());
  const locals = [];
  const central = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = textEncoder.encode(entry.name);
    const raw = textEncoder.encode(entry.content);
    const compressed = await deflateRaw(raw);
    const payload = compressed || raw;
    const method = compressed ? 8 : 0;
    const checksum = crc32(raw);

    const localHeader = new DataView(new ArrayBuffer(30));
    localHeader.setUint32(0, 0x04034b50, true);
    localHeader.setUint16(4, 20, true);
    localHeader.setUint16(6, 0x0800, true);
    localHeader.setUint16(8, method, true);
    localHeader.setUint16(10, stamp.time, true);
    localHeader.setUint16(12, stamp.day, true);
    localHeader.setUint32(14, checksum, true);
    localHeader.setUint32(18, payload.length, true);
    localHeader.setUint32(22, raw.length, true);
    localHeader.setUint16(26, nameBytes.length, true);
    localHeader.setUint16(28, 0, true);
    locals.push(new Uint8Array(localHeader.buffer), nameBytes, payload);

    const centralHeader = new DataView(new ArrayBuffer(46));
    centralHeader.setUint32(0, 0x02014b50, true);
    centralHeader.setUint16(4, 20, true);
    centralHeader.setUint16(6, 20, true);
    centralHeader.setUint16(8, 0x0800, true);
    centralHeader.setUint16(10, method, true);
    centralHeader.setUint16(12, stamp.time, true);
    centralHeader.setUint16(14, stamp.day, true);
    centralHeader.setUint32(16, checksum, true);
    centralHeader.setUint32(20, payload.length, true);
    centralHeader.setUint32(24, raw.length, true);
    centralHeader.setUint16(28, nameBytes.length, true);
    centralHeader.setUint32(42, offset, true);
    central.push(new Uint8Array(centralHeader.buffer), nameBytes);

    offset += 30 + nameBytes.length + payload.length;
  }

  const centralSize = central.reduce((total, part) => total + part.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, entries.length, true);
  end.setUint16(10, entries.length, true);
  end.setUint32(12, centralSize, true);
  end.setUint32(16, offset, true);

  return new Blob([...locals, ...central, new Uint8Array(end.buffer)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

/**
 * sheets: [{ name, rows }] where rows[0] is the header row and every cell is a
 * string, a number or null.
 */
export async function buildWorkbook(sheets) {
  const usable = (sheets || []).filter(sheet => Array.isArray(sheet?.rows) && sheet.rows.length);
  if (!usable.length) throw new Error("The workbook has no rows to write.");

  const names = [];
  usable.forEach((sheet, index) => {
    let name = sanitizeSheetName(sheet.name, `Sheet${index + 1}`);
    let suffix = 2;
    while (names.includes(name)) name = `${sanitizeSheetName(sheet.name, `Sheet${index + 1}`).slice(0, 28)} ${suffix++}`;
    names.push(name);
  });

  const entries = [
    {
      name: "[Content_Types].xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${usable.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`,
    },
    {
      name: "_rels/.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    },
    {
      name: "xl/workbook.xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((name, index) => `<sheet name="${escapeXml(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets></workbook>`,
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}<Relationship Id="rId${names.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    },
    { name: "xl/styles.xml", content: STYLES_XML },
    ...usable.map((sheet, index) => ({ name: `xl/worksheets/sheet${index + 1}.xml`, content: sheetXml(sheet.rows) })),
  ];

  return zip(entries);
}

export async function downloadWorkbook(sheets, filename) {
  const blob = await buildWorkbook(sheets);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 4000);
}
