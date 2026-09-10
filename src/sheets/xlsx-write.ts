// .xlsx writer: regenerates only edited worksheets, shared strings and styles;
// every other part of the package is kept byte for byte. calcChain is dropped
// and fullCalcOnLoad set so Excel/LibreOffice recalculate on open.
import { Workbook, Sheet, Cell, Xf, Font, Fill, Border, BorderSide, Color, key, rowOf, colOf, cellRef, colName, isError, rangeRef } from "./model";
import { escapeXml, escapeXmlText } from "../docx/xml";
import { Package } from "../docx/zip";

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_SST = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_WB = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const REL_SHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SST = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_CALC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

const SHEET_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];

function colorXml(tag: string, c: Color | null): string {
  if (!c) return "";
  const at: string[] = [];
  if (c.auto) at.push('auto="1"');
  if (c.rgb) at.push(`rgb="${c.rgb.length === 6 ? "FF" + c.rgb : c.rgb}"`);
  if (c.theme !== undefined) at.push(`theme="${c.theme}"`);
  if (c.indexed !== undefined) at.push(`indexed="${c.indexed}"`);
  if (c.tint !== undefined && c.tint !== 0) at.push(`tint="${c.tint}"`);
  return `<${tag} ${at.join(" ")}/>`;
}
function fontXml(f: Font): string {
  if (f.raw) return f.raw;
  let s = "<font>";
  if (f.bold) s += "<b/>"; if (f.italic) s += "<i/>"; if (f.strike) s += "<strike/>";
  if (f.underline) s += f.underline === "single" ? "<u/>" : `<u val="${f.underline}"/>`;
  s += `<sz val="${f.size}"/>`;
  s += colorXml("color", f.color);
  s += `<name val="${escapeXml(f.name)}"/>`;
  if (f.family !== undefined) s += `<family val="${f.family}"/>`;
  if (f.charset !== undefined) s += `<charset val="${f.charset}"/>`;
  if (f.scheme) s += `<scheme val="${f.scheme}"/>`;
  return s + "</font>";
}
function fillXml(f: Fill): string {
  if (f.raw) return f.raw;
  if (f.patternType === "none" || f.patternType === "gray125") return `<fill><patternFill patternType="${f.patternType}"/></fill>`;
  return `<fill><patternFill patternType="${f.patternType}">${colorXml("fgColor", f.fg)}${colorXml("bgColor", f.bg)}</patternFill></fill>`;
}
function sideXml(tag: string, s: BorderSide | null): string {
  if (!s) return `<${tag}/>`;
  return `<${tag} style="${s.style}">${colorXml("color", s.color)}</${tag}>`;
}
function borderXml(b: Border): string {
  if (b.raw) return b.raw;
  const at = (b.diagonalUp ? ' diagonalUp="1"' : "") + (b.diagonalDown ? ' diagonalDown="1"' : "");
  return `<border${at}>${sideXml("left", b.left)}${sideXml("right", b.right)}${sideXml("top", b.top)}${sideXml("bottom", b.bottom)}${sideXml("diagonal", b.diagonal)}</border>`;
}
function xfXml(x: Xf, cellXf: boolean): string {
  const at: string[] = [`numFmtId="${x.numFmtId}"`, `fontId="${x.fontId}"`, `fillId="${x.fillId}"`, `borderId="${x.borderId}"`];
  if (cellXf) at.push(`xfId="${x.xfId}"`);
  if (x.quotePrefix) at.push('quotePrefix="1"');
  for (const k of ["applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment", "applyProtection"] as const) if ((x as any)[k] !== undefined) at.push(`${k}="${(x as any)[k] ? 1 : 0}"`);
  let inner = "";
  if (x.alignment) {
    const a = x.alignment; const aa: string[] = [];
    if (a.horizontal) aa.push(`horizontal="${a.horizontal}"`); if (a.vertical) aa.push(`vertical="${a.vertical}"`); if (a.wrapText) aa.push('wrapText="1"');
    if (a.indent) aa.push(`indent="${a.indent}"`); if (a.textRotation) aa.push(`textRotation="${a.textRotation}"`); if (a.shrinkToFit) aa.push('shrinkToFit="1"'); if (a.readingOrder) aa.push(`readingOrder="${a.readingOrder}"`);
    inner += `<alignment ${aa.join(" ")}/>`;
  }
  if (x.protection) { const pa: string[] = []; if (x.protection.locked !== undefined) pa.push(`locked="${x.protection.locked ? 1 : 0}"`); if (x.protection.hidden) pa.push('hidden="1"'); inner += `<protection ${pa.join(" ")}/>`; }
  return inner ? `<xf ${at.join(" ")}>${inner}</xf>` : `<xf ${at.join(" ")}/>`;
}

/**
 * Parts written by some tools use a prefix for the main namespace (`<x:worksheet xmlns:x=…>`). We emit
 * unprefixed elements, so the root must declare the default namespace and raw fragments copied from
 * the original must lose that prefix.
 */
function mainPrefix(rootAttrs: string): string | null {
  const m = new RegExp(`xmlns:([A-Za-z0-9_]+)="${NS_SS}"`).exec(rootAttrs);
  return m ? m[1] : null;
}
function rootAttrsWithDefault(rootAttrs: string, fallback: string): string {
  if (!rootAttrs) return fallback;
  return /(^|\s)xmlns="/.test(rootAttrs) ? rootAttrs : `xmlns="${NS_SS}" ` + rootAttrs;
}
function stripMainPrefix(xml: string, rootAttrs: string): string {
  const p = mainPrefix(rootAttrs);
  if (!p) return xml;
  return xml.replace(new RegExp(`<(/?)${p}:`, "g"), "<$1");
}

export function stylesXml(wb: Workbook): string {
  const st = wb.styles;
  return stripMainPrefix(stylesXmlInner(wb), st.rootAttrs);
}
function stylesXmlInner(wb: Workbook): string {
  const st = wb.styles;
  let s = XML + `<styleSheet ${rootAttrsWithDefault(st.rootAttrs, `xmlns="${NS_SS}"`)}>`;
  if (st.numFmts.size) { s += `<numFmts count="${st.numFmts.size}">`; for (const [id, code] of st.numFmts) s += `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`; s += "</numFmts>"; }
  s += `<fonts count="${st.fonts.length}">${st.fonts.map(fontXml).join("")}</fonts>`;
  s += `<fills count="${st.fills.length}">${st.fills.map(fillXml).join("")}</fills>`;
  s += `<borders count="${st.borders.length}">${st.borders.map(borderXml).join("")}</borders>`;
  s += `<cellStyleXfs count="${st.cellStyleXfs.length}">${st.cellStyleXfs.map((x) => xfXml(x, false)).join("")}</cellStyleXfs>`;
  s += `<cellXfs count="${st.xfs.length}">${st.xfs.map((x) => xfXml(x, true)).join("")}</cellXfs>`;
  const order = ["cellStyles", "dxfs", "tableStyles", "colors", "extLst"];
  const tail = [...st.rawTail].sort((a, b) => order.indexOf(a.name) - order.indexOf(b.name));
  if (!tail.some((t) => t.name === "cellStyles")) s += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
  for (const t of tail) s += t.xml;
  return s + "</styleSheet>";
}

const ST_ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"]);

function cellXml(cell: Cell, r: number, c: number, sst: SstBuilder): string {
  const ref = cellRef(r, c);
  const sAttr = cell.s ? ` s="${cell.s}"` : "";
  let f = "";
  if (cell.f !== undefined) f = cell.arr ? `<f t="array" ref="${cell.arr}">${escapeXmlText(cell.f)}</f>` : `<f>${escapeXmlText(cell.f)}</f>`;
  const v = cell.v;
  if (v === null || v === undefined) return f ? `<c r="${ref}"${sAttr}>${f}</c>` : `<c r="${ref}"${sAttr}/>`;
  if (typeof v === "number") return `<c r="${ref}"${sAttr}>${f}<v>${isFinite(v) ? String(v) : "0"}</v></c>`;
  if (typeof v === "boolean") return `<c r="${ref}"${sAttr} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
  if (isError(v)) {
    // Excel only accepts ST_Error values; our internal #CIRC! (Excel stores 0) and #ERROR! (parse failure) must not reach the file.
    if (v.e === "#CIRC!") return `<c r="${ref}"${sAttr}>${f}<v>0</v></c>`;
    const e = ST_ERRORS.has(v.e) ? v.e : "#VALUE!";
    return `<c r="${ref}"${sAttr} t="e">${f}<v>${escapeXmlText(e)}</v></c>`;
  }
  // string
  if (cell.f !== undefined) return `<c r="${ref}"${sAttr} t="str">${f}<v>${escapeXmlText(v)}</v></c>`;
  if (cell.rich) return `<c r="${ref}"${sAttr} t="inlineStr">${cell.rich}</c>`;
  return `<c r="${ref}"${sAttr} t="s"><v>${sst.index(v)}</v></c>`;
}

class SstBuilder {
  constructor(private wb: Workbook) {}
  index(s: string): number {
    let i = this.wb.sstIndex.get(s);
    if (i === undefined) { i = this.wb.sst.length; this.wb.sst.push(s); this.wb.sstRaw.push(""); this.wb.sstIndex.set(s, i); }
    return i;
  }
}

export function sheetXml(sheet: Sheet, wb: Workbook, sst: SstBuilder): string {
  const parts = new Map<string, string>();
  for (const rb of sheet.rawBlocks) parts.set(rb.name, (parts.get(rb.name) || "") + rb.xml);
  // sheetPr: tab colour (children order: tabColor, outlinePr, pageSetUpPr)
  {
    let sp = parts.get("sheetPr") || "";
    if (!sheet.tabColorChanged) { /* keep the original <tabColor> (theme / indexed colours are not modelled) */ }
    else if (sheet.tabColor) {
      const tc = `<tabColor rgb="FF${sheet.tabColor.replace(/^#/, "").toUpperCase().slice(-6)}"/>`;
      if (!sp) sp = `<sheetPr>${tc}</sheetPr>`;
      else if (/<tabColor/.test(sp)) sp = sp.replace(/<tabColor[^>]*\/>|<tabColor[^>]*>[\s\S]*?<\/tabColor>/, tc);
      else if (/<sheetPr[^>]*\/>/.test(sp)) sp = sp.replace(/<sheetPr([^>]*)\/>/, `<sheetPr$1>${tc}</sheetPr>`);
      else sp = sp.replace(/(<sheetPr[^>]*>)/, `$1${tc}`);
      parts.set("sheetPr", sp);
    } else if (sp && /<tabColor/.test(sp)) {
      sp = sp.replace(/<tabColor[^>]*\/>|<tabColor[^>]*>[\s\S]*?<\/tabColor>/, "");
      parts.set("sheetPr", /<sheetPr[^>]*>\s*<\/sheetPr>/.test(sp) ? "" : sp);
      if (!parts.get("sheetPr")) parts.delete("sheetPr");
    }
  }
  // dimension
  const dim = sheet.maxRow >= 0 ? `A1:${cellRef(sheet.maxRow, Math.max(0, sheet.maxCol))}` : "A1";
  parts.set("dimension", `<dimension ref="${dim}"/>`);
  // sheetViews
  const v = sheet.view;
  let sv = `<sheetView workbookViewId="0"${v.tabSelected ? ' tabSelected="1"' : ""}${!v.showGridLines ? ' showGridLines="0"' : ""}${v.zoom && v.zoom !== 100 ? ` zoomScale="${Math.round(v.zoom)}" zoomScaleNormal="${Math.round(v.zoom)}"` : ""}${v.rightToLeft ? ' rightToLeft="1"' : ""}${v.topLeft ? ` topLeftCell="${v.topLeft}"` : ""}>`;
  const active = v.active || "A1", sqref = v.selection || active;
  if (sheet.freeze && !sheet.autoFreeze && (sheet.freeze.rows || sheet.freeze.cols)) {
    const f = sheet.freeze;
    const pane = f.rows && f.cols ? "bottomRight" : f.rows ? "bottomLeft" : "topRight";
    sv += `<pane${f.cols ? ` xSplit="${f.cols}"` : ""}${f.rows ? ` ySplit="${f.rows}"` : ""} topLeftCell="${cellRef(f.rows, f.cols)}" activePane="${pane}" state="frozen"/>`;
    if (f.rows && f.cols) sv += `<selection pane="topRight"/><selection pane="bottomLeft"/>`;
    sv += `<selection pane="${pane}" activeCell="${active}" sqref="${sqref}"/>`;
  } else sv += `<selection activeCell="${active}" sqref="${sqref}"/>`;
  sv += "</sheetView>";
  parts.set("sheetViews", `<sheetViews>${sv}</sheetViews>`);
  parts.set("sheetFormatPr", `<sheetFormatPr defaultRowHeight="${sheet.defaultRowHeight}"${Math.abs(sheet.defaultColWidth - 8.43) > 0.01 ? ` defaultColWidth="${sheet.defaultColWidth}"` : ""}/>`);
  if (sheet.cols.length) {
    parts.set("cols", "<cols>" + sheet.cols.map((c) => `<col min="${c.min + 1}" max="${c.max + 1}"${c.width !== undefined ? ` width="${c.width}"` : ""}${c.style ? ` style="${c.style}"` : ""}${c.hidden ? ' hidden="1"' : ""}${c.customWidth || c.width !== undefined ? ' customWidth="1"' : ""}${c.bestFit ? ' bestFit="1"' : ""}${c.level ? ` outlineLevel="${c.level}"` : ""}/>`).join("") + "</cols>");
  } else parts.delete("cols");
  // sheetData
  const rowsMap = new Map<number, { r: number; cells: [number, Cell][] }>();
  for (const [k, cell] of sheet.cells) { const r = rowOf(k), c = colOf(k); let row = rowsMap.get(r); if (!row) { row = { r, cells: [] }; rowsMap.set(r, row); } row.cells.push([c, cell]); }
  for (const r of sheet.rows.keys()) if (!rowsMap.has(r)) rowsMap.set(r, { r, cells: [] });
  const rowNums = Array.from(rowsMap.keys()).sort((a, b) => a - b);
  let sd = "<sheetData>";
  for (const r of rowNums) {
    const row = rowsMap.get(r)!;
    const info = sheet.rows.get(r);
    row.cells.sort((a, b) => a[0] - b[0]);
    let at = ` r="${r + 1}"`;
    if (row.cells.length) at += ` spans="${row.cells[0][0] + 1}:${row.cells[row.cells.length - 1][0] + 1}"`;
    if (info) {
      if (info.height !== undefined) at += ` ht="${info.height}"${info.customHeight === true ? ' customHeight="1"' : ""}`;
      if (info.hidden) at += ' hidden="1"';
      if (info.style) at += ` s="${info.style}" customFormat="1"`;
      if (info.level) at += ` outlineLevel="${info.level}"`;
    }
    if (!row.cells.length && !info) continue;
    sd += `<row${at}>${row.cells.map(([c, cell]) => cellXml(cell, r, c, sst)).join("")}</row>`;
  }
  sd += "</sheetData>";
  parts.set("sheetData", sd);
  if (sheet.autoFilter) {
    const raw = parts.get("autoFilter");
    const ref = rangeRef(sheet.autoFilter);
    const rawRef = raw ? /\bref="([^"]*)"/.exec(raw)?.[1] : null;
    if (!raw || rawRef !== ref) parts.set("autoFilter", `<autoFilter ref="${ref}"/>`);
  } else parts.delete("autoFilter");
  if (sheet.merges.length) parts.set("mergeCells", `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${rangeRef(m)}"/>`).join("")}</mergeCells>`); else parts.delete("mergeCells");
  if (sheet.hyperlinks.length) parts.set("hyperlinks", "<hyperlinks>" + sheet.hyperlinks.map((h) => `<hyperlink ref="${h.ref}"${h.rId ? ` r:id="${h.rId}"` : ""}${h.location ? ` location="${escapeXml(h.location)}"` : ""}${h.display ? ` display="${escapeXml(h.display)}"` : ""}${h.tooltip ? ` tooltip="${escapeXml(h.tooltip)}"` : ""}/>`).join("") + "</hyperlinks>"); else parts.delete("hyperlinks");
  let out = XML + `<worksheet ${rootAttrsWithDefault(sheet.rawRootAttrs, `xmlns="${NS_SS}" xmlns:r="${NS_R}"`)}>`;
  if (sheet.hyperlinks.some((h) => h.rId) && !/xmlns:r=/.test(out)) out = out.replace(/>$/, ` xmlns:r="${NS_R}">`);
  for (const name of SHEET_ORDER) { const p = parts.get(name); if (p) out += p; }
  return stripMainPrefix(out + "</worksheet>", sheet.rawRootAttrs);
}

function sstXml(wb: Workbook): string {
  let s = XML + `<sst xmlns="${NS_SS}" count="${wb.sst.length}" uniqueCount="${wb.sst.length}">`;
  for (let i = 0; i < wb.sst.length; i++) {
    const raw = wb.sstRaw[i];
    // rich-text items copied from a prefixed source ("<x:si>") are re-emitted unprefixed
    s += raw ? raw.replace(/<(\/?)[A-Za-z0-9_]+:(si|t|r|rPr|rFont|b|i|u|sz|color|family|charset|scheme|strike|vertAlign|outline|shadow|condense|extend|rPh|phoneticPr)\b/g, "<$1$2") : `<si><t xml:space="preserve">${escapeXmlText(wb.sst[i])}</t></si>`;
  }
  return s + "</sst>";
}

function escapeRe(s: string): string { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

function ensureContentType(pkg: Package, partName: string, contentType: string) {
  let ct = pkg.text("[Content_Types].xml") || XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>';
  if (!ct.includes(`PartName="/${partName}"`)) ct = ct.replace(/<\/Types>\s*$/, `<Override PartName="/${partName}" ContentType="${contentType}"/></Types>`);
  pkg.setText("[Content_Types].xml", ct);
}
function removeContentType(pkg: Package, partName: string) {
  const ct = pkg.text("[Content_Types].xml");
  if (!ct) return;
  pkg.setText("[Content_Types].xml", ct.replace(new RegExp(`<Override PartName="/${partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*/>`), ""));
}

export function writeXlsx(wb: Workbook): Uint8Array {
  const pkg = wb.pkg;
  const isNew = !pkg.has(wb.workbookPart);
  const wbDir = wb.workbookPart.split("/").slice(0, -1).join("/") || "xl";
  const relsPart = wbDir + "/_rels/" + wb.workbookPart.split("/").pop() + ".rels";
  let rels = pkg.text(relsPart) || XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  let nextRid = 1;
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) nextRid = Math.max(nextRid, parseInt(m[1], 10) + 1);
  const addRel = (type: string, target: string): string => { const id = "rId" + nextRid++; rels = rels.replace(/<\/Relationships>\s*$/, `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`); return id; };
  const sst = new SstBuilder(wb);
  // sheets
  wb.sheets.forEach((sheet, i) => {
    const registered = !!sheet.rId && new RegExp(`\\sId="${escapeRe(sheet.rId)}"`).test(rels) && pkg.has(sheet.part);
    if (!registered) {
      if (!sheet.part || pkg.has(sheet.part)) { sheet.part = `${wbDir}/worksheets/sheet${i + 1}.xml`; let n = i + 1; while (pkg.has(sheet.part)) { n++; sheet.part = `${wbDir}/worksheets/sheet${n}.xml`; } }
      sheet.rId = addRel(REL_SHEET, "worksheets/" + sheet.part.split("/").pop());
      sheet.dirty = true;
      ensureContentType(pkg, sheet.part, CT_SHEET);
    }
    if (sheet.dirty || isNew) { pkg.setText(sheet.part, sheetXml(sheet, wb, sst)); sheet.dirty = false; }
  });
  // sheets deleted in this session: drop their parts, rels and content types
  for (const rm of wb.removed || []) {
    pkg.delete(rm.part);
    const relsOfSheet = rm.part.replace(/([^/]+)$/, "_rels/$1.rels");
    pkg.delete(relsOfSheet);
    rels = rels.replace(new RegExp(`<Relationship\\s[^>]*\\sId="${escapeRe(rm.rId)}"[^>]*/>`), "");
    removeContentType(pkg, rm.part);
  }
  wb.removed = [];
  // shared strings
  let sstPart = wb.sstPart;
  if (wb.sst.length || sstPart) {
    if (!sstPart) { sstPart = wbDir + "/sharedStrings.xml"; wb.sstPart = sstPart; addRel(REL_SST, "sharedStrings.xml"); ensureContentType(pkg, sstPart, CT_SST); }
    pkg.setText(sstPart, sstXml(wb));
  }
  // styles
  if (!pkg.has(wb.stylesPart)) { addRel(REL_STYLES, "styles.xml"); ensureContentType(pkg, wb.stylesPart, CT_STYLES); }
  pkg.setText(wb.stylesPart, stylesXml(wb));
  // workbook.xml: sheets list, calcPr
  let wbXml = wb.workbookXml;
  if (!wbXml) {
    wbXml = XML + `<workbook xmlns="${NS_SS}" xmlns:r="${NS_R}"><workbookPr${wb.date1904 ? ' date1904="1"' : ""}/><bookViews><workbookView activeTab="0"/></bookViews><sheets></sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>`;
    ensureContentType(pkg, wb.workbookPart, CT_WB);
    if (!pkg.has("_rels/.rels")) pkg.setText("_rels/.rels", XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${wb.workbookPart}"/></Relationships>`);
  }
  const sheetsEl = "<sheets>" + wb.sheets.map((s) => `<sheet name="${escapeXml(s.name)}" sheetId="${s.sheetId}"${s.state !== "visible" ? ` state="${s.state}"` : ""} r:id="${s.rId}"/>`).join("") + "</sheets>";
  const prefix = /<(\w+:)?sheets[\s>]/.exec(wbXml)?.[1] || "";
  wbXml = wbXml.replace(new RegExp(`<${prefix}sheets>[\\s\\S]*?</${prefix}sheets>|<${prefix}sheets/>`), prefix ? sheetsEl.replace(/<(\/?)sheet/g, `<$1${prefix}sheet`) : sheetsEl);
  if (!/xmlns:r=/.test(wbXml.slice(0, 600))) wbXml = wbXml.replace(/<(\w+:)?workbook\b/, (m) => m + ` xmlns:r="${NS_R}"`);
  const hasFormulas = wb.sheets.some((s) => { for (const c of s.cells.values()) if (c.f) return true; return false; });
  if (hasFormulas) {
    if (/<(\w+:)?calcPr\b[^>]*fullCalcOnLoad=/.test(wbXml)) wbXml = wbXml.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"');
    else if (/<(\w+:)?calcPr\b/.test(wbXml)) wbXml = wbXml.replace(/<((\w+:)?calcPr\b)/, '<$1 fullCalcOnLoad="1"');
    else wbXml = wbXml.replace(/<\/(\w+:)?workbook>\s*$/, (m) => `<${prefix}calcPr calcId="0" fullCalcOnLoad="1"/>` + m);
  }
  // defined names: drop names scoped to deleted sheets, renumber localSheetId, keep other attributes
  {
    const names = wb.definedNames.filter((d) => !(d.sheet && !wb.sheets.includes(d.sheet)));
    wb.definedNames = names;
    const dnEl = names.length ? `<${prefix}definedNames>` + names.map((d) => {
      const m = /^<(?:\w+:)?definedName\b([^>]*)>/.exec(d.raw || "");
      let attrs = (m ? m[1] : ` name="${escapeXml(d.name)}"${d.hidden ? ' hidden="1"' : ""}`).replace(/\s+localSheetId="[^"]*"/, "");
      if (d.sheet) { const li = wb.sheets.indexOf(d.sheet); attrs += ` localSheetId="${li}"`; d.localSheetId = li; }
      return `<${prefix}definedName${attrs}>${escapeXmlText(d.ref)}</${prefix}definedName>`;
    }).join("") + `</${prefix}definedNames>` : "";
    const dnRe = new RegExp(`<${prefix}definedNames>[\\s\\S]*?</${prefix}definedNames>|<${prefix}definedNames/>`);
    if (dnRe.test(wbXml)) wbXml = wbXml.replace(dnRe, dnEl);
    else if (dnEl) wbXml = wbXml.replace(new RegExp(`</${prefix}sheets>`), `</${prefix}sheets>` + dnEl);
  }
  // active tab
  wbXml = wbXml.replace(/activeTab="\d+"/, `activeTab="${wb.active}"`);
  pkg.setText(wb.workbookPart, wbXml);
  wb.workbookXml = wbXml;
  // calc chain
  if (wb.calcChainPart) {
    pkg.delete(wb.calcChainPart);
    rels = rels.replace(new RegExp(`<Relationship [^>]*Type="${REL_CALC}"[^>]*/>`), "");
    removeContentType(pkg, wb.calcChainPart);
    wb.calcChainPart = null;
  }
  pkg.setText(relsPart, rels);
  wb.workbookRels = rels;
  if (!pkg.has("docProps/app.xml")) {
    pkg.setText("docProps/app.xml", XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>OfficeMini</Application></Properties>');
    ensureContentType(pkg, "docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
    let root = pkg.text("_rels/.rels")!;
    if (!root.includes("docProps/app.xml")) { root = root.replace(/<\/Relationships>\s*$/, '<Relationship Id="rIdApp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'); pkg.setText("_rels/.rels", root); }
  }
  return pkg.toBytes();
}

export { colName, key };
