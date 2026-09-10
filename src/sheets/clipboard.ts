// Copy / cut / paste. Internal pastes carry cells (values, formulas, styles)
// and behave like Google Sheets: relative references shift with the paste
// offset, cut keeps them. External pastes use TSV (and HTML tables).
import { Workbook, Sheet, Cell, Range, key, Value, isError, MAXR, MAXC } from "./model";
import { shiftFormula } from "./formula/tokens";
import { CellChange, boundRange, xfWith, inputCell, StylePatch } from "./edit";
import { formatValue, formatCodeFor, editText } from "./numfmt";
import { StyleResolver } from "./render-style";

export interface ClipCell { r: number; c: number; cell: Cell | null; }
export interface ClipData { sheet: string; range: Range; cells: ClipCell[]; cut: boolean; styles: { xf: any; font: any; fill: any; border: any; numFmt: string }[]; text?: string; }

let internal: ClipData | null = null;
let internalToken = "";

export function internalClip(): ClipData | null { return internal; }

/** Snapshot the selection into the internal clipboard and the system clipboard. */
export function copyRange(wb: Workbook, sheet: Sheet, range0: Range, cut: boolean, styles: StyleResolver): { text: string; html: string } {
  const range = boundRange(sheet, range0);
  const cells: ClipCell[] = [];
  const styleIdx = new Map<number, number>();
  const styleList: ClipData["styles"] = [];
  // Rows hidden by a filter are not copied (Excel / Sheets copy the visible cells only); the clip is compacted.
  const rows: number[] = [];
  for (let r = range.r1; r <= range.r2; r++) if (!(sheet.hiddenRowsByFilter.size && sheet.hiddenRowsByFilter.has(r))) rows.push(r);
  const rowIdx = new Map(rows.map((r, i) => [r, i]));
  for (const r of rows) for (let c = range.c1; c <= range.c2; c++) {
    const cell = sheet.cells.get(key(r, c));
    if (cell && !styleIdx.has(cell.s)) {
      const xf = wb.styles.xfs[cell.s] || wb.styles.xfs[0];
      styleList.push({ xf, font: wb.styles.fonts[xf.fontId], fill: wb.styles.fills[xf.fillId], border: wb.styles.borders[xf.borderId], numFmt: formatCodeFor(xf.numFmtId, wb.styles.numFmts) });
      styleIdx.set(cell.s, styleList.length - 1);
    }
    cells.push({ r: rowIdx.get(r)!, c: c - range.c1, cell: cell ? { ...cell } : null });
  }
  const clipRange: Range = rows.length && rows.length !== range.r2 - range.r1 + 1 ? { r1: range.r1, c1: range.c1, r2: range.r1 + rows.length - 1, c2: range.c2 } : { ...range };
  internal = { sheet: sheet.name, range: clipRange, cells, cut: cut && rows.length === range.r2 - range.r1 + 1, styles: styleList };
  internalToken = "om-sheets:" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  // text/plain: TSV of displayed values; text/html: table with basic styles
  const lines: string[] = [];
  let html = `<!-- ${internalToken} --><table>`;
  for (const r of rows) {
    const fields: string[] = [];
    html += "<tr>";
    for (let c = range.c1; c <= range.c2; c++) {
      const cell = sheet.cells.get(key(r, c));
      const cs = styles.get(cell ? cell.s : 0);
      const text = cell ? styles.render(cell, cs).text : "";
      fields.push(text.replace(/\t/g, " ").replace(/\r?\n/g, " "));
      const st: string[] = [];
      if (cs.bold) st.push("font-weight:bold"); if (cs.italic) st.push("font-style:italic"); if (cs.color) st.push("color:" + cs.color); if (cs.fill) st.push("background:" + cs.fill);
      if (cs.halign && cs.halign !== "general") st.push("text-align:" + cs.halign);
      html += `<td${st.length ? ` style="${st.join(";")}"` : ""}>${escapeHtml(text)}</td>`;
    }
    html += "</tr>";
    lines.push(fields.join("\t"));
  }
  html += "</table>";
  internal.text = lines.join("\n");
  return { text: internal.text, html };
}

function escapeHtml(s: string) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

export function clearInternalClip() { internal = null; }
/** Esc after Ctrl+X: the data stays copied but the source is no longer deleted on paste. */
export function cancelCut() { if (internal) internal.cut = false; }

/** Is the HTML on the system clipboard the one we wrote last? */
export function isOurHtml(html: string | null): boolean {
  return !!html && !!internalToken && html.includes(internalToken);
}

export type PasteMode = "all" | "values" | "formats" | "formulas" | "transpose";

/**
 * Compute cell changes for pasting the internal clipboard at `at` (top-left).
 * If the target selection is larger than the clip and a multiple of it, the
 * clip is tiled (like Sheets); a single-cell clip fills the whole selection.
 */
/**
 * Rows a paste lands on, starting at `from`. Rows the filter has hidden (or that were hidden
 * by hand) are skipped, so pasting a column of values into a filtered list fills the rows you
 * can see instead of the ones in between. `count` rows always come back; if the sheet runs out
 * of visible rows the remainder continues past the end.
 */
function rowIsHidden(sheet: Sheet, r: number): boolean {
  return !!sheet.rows.get(r)?.hidden || sheet.hiddenRowsByFilter.has(r);
}

function pasteRows(sheet: Sheet, from: number, count: number): number[] {
  const out: number[] = [];
  let r = from;
  while (out.length < count && r < MAXR) { if (!rowIsHidden(sheet, r)) out.push(r); r++; }
  while (out.length < count) out.push(MAXR - 1);
  return out;
}

export function pasteInternal(wb: Workbook, sheet: Sheet, target: Range, mode: PasteMode): { changes: CellChange[]; resultRange: Range; cutSource: { sheet: string; range: Range } | null } {
  const clip = internal!;
  const h = clip.range.r2 - clip.range.r1 + 1, w = clip.range.c2 - clip.range.c1 + 1;
  const th = target.r2 - target.r1 + 1, tw = target.c2 - target.c1 + 1;
  const transpose = mode === "transpose";
  const ch = transpose ? w : h, cw = transpose ? h : w;
  const repR = th >= ch && th % ch === 0 && (th > ch || tw > cw) && !clip.cut ? th / ch : 1;
  const repC = tw >= cw && tw % cw === 0 && (th > ch || tw > cw) && !clip.cut ? tw / cw : 1;
  const changes: CellChange[] = [];
  const rows = pasteRows(sheet, target.r1, repR * ch);
  const styleMap = new Map<number, number>();
  const mapStyle = (cell: Cell): number => {
    // Same workbook: style indices are valid as-is (clipboard survives sheet switches).
    return cell.s;
  };
  const srcSheetIsThis = clip.sheet.toLowerCase() === sheet.name.toLowerCase();
  for (let rr = 0; rr < repR; rr++) for (let cc = 0; cc < repC; cc++) {
    for (const cc0 of clip.cells) {
      const pr = transpose ? cc0.c : cc0.r, pc = transpose ? cc0.r : cc0.c;
      const r = rows[rr * ch + pr], c = target.c1 + cc * cw + pc;
      if (r === undefined || r >= MAXR || c >= MAXC) continue;
      const existing = sheet.cells.get(key(r, c));
      const src = cc0.cell;
      if (!src) { if (mode !== "formats" && existing) changes.push({ r, c, cell: existing.s && mode === "values" ? { v: null, s: existing.s } : null }); continue; }
      const dr = r - (clip.range.r1 + cc0.r), dc = c - (clip.range.c1 + cc0.c);
      let cell: Cell;
      if (mode === "values") cell = { v: src.v, s: existing ? existing.s : 0 };
      else if (mode === "formats") { if (!existing) { changes.push({ r, c, cell: { v: null, s: mapStyle(src) } }); continue; } cell = { ...existing, s: mapStyle(src) }; }
      else if (mode === "formulas") cell = { v: src.f ? null : src.v, f: src.f ? (clip.cut ? src.f : shiftFormula(src.f, dr, dc)) : undefined, s: existing ? existing.s : 0 };
      else { cell = { ...src, s: mapStyle(src) }; if (src.f) { cell.f = clip.cut ? src.f : shiftFormula(src.f, dr, dc); if (!clip.cut) cell.v = null; } }
      if (cell.f === undefined) delete cell.f;
      delete cell.sh; delete cell.arr;
      changes.push({ r, c, cell });
    }
  }
  const resultRange: Range = { r1: target.r1, c1: target.c1, r2: rows[rows.length - 1] ?? target.r1, c2: target.c1 + repC * cw - 1 };
  void styleMap; void srcSheetIsThis;
  return { changes, resultRange, cutSource: clip.cut ? { sheet: clip.sheet, range: clip.range } : null };
}

/** Parse external clipboard content (HTML table preferred, else TSV) into a value grid. */
export function parseExternal(html: string | null, text: string): string[][] {
  if (html && /<table/i.test(html)) {
    try {
      const doc = new DOMParser().parseFromString(html, "text/html");
      const table = doc.querySelector("table");
      if (table) {
        const rows: string[][] = [];
        for (const tr of Array.from(table.querySelectorAll("tr"))) {
          const cells: string[] = [];
          for (const td of Array.from(tr.children)) {
            if (!/^t[dh]$/i.test(td.tagName)) continue;
            const span = parseInt(td.getAttribute("colspan") || "1", 10) || 1;
            cells.push((td as HTMLElement).innerText.replace(/ /g, " ").trim());
            for (let i = 1; i < span; i++) cells.push("");
          }
          rows.push(cells);
        }
        if (rows.length) return rows;
      }
    } catch { /* fall through */ }
  }
  const t = text.replace(/\r\n?/g, "\n");
  const lines = t.endsWith("\n") ? t.slice(0, -1).split("\n") : t.split("\n");
  // Quoted TSV fields with embedded newlines
  const rows: string[][] = [];
  let cur: string[] | null = null, field = "", q = false;
  for (const line of lines) {
    if (!cur) cur = [];
    let i = 0;
    while (i < line.length) {
      const ch = line[i];
      if (q) { if (ch === '"') { if (line[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; } field += ch; i++; continue; }
      if (ch === '"' && field === "") { q = true; i++; continue; }
      if (ch === "\t") { cur.push(field); field = ""; i++; continue; }
      field += ch; i++;
    }
    if (q) { field += "\n"; continue; }
    cur.push(field); field = "";
    rows.push(cur); cur = null;
  }
  if (cur) { cur.push(field); rows.push(cur); }
  return rows;
}

export function pasteExternal(wb: Workbook, sheet: Sheet, at: { r: number; c: number }, grid: string[][], target: Range | null): { changes: CellChange[]; resultRange: Range } {
  const changes: CellChange[] = [];
  let maxC = 0;
  const single = grid.length === 1 && grid[0].length === 1;
  const reps = single && target ? target : null;
  const put = (r: number, c: number, text: string) => {
    if (r >= MAXR || c >= MAXC) return;
    const cell = inputCell(wb, sheet, r, c, text);
    changes.push({ r, c, cell });
  };
  if (reps) {
    for (let r = reps.r1; r <= reps.r2; r++) for (let c = reps.c1; c <= reps.c2; c++) if (!rowIsHidden(sheet, r)) put(r, c, grid[0][0]);
    return { changes, resultRange: reps };
  }
  // Same rule as an internal paste: rows the filter hides are skipped, not written through.
  const rows = pasteRows(sheet, at.r, grid.length);
  grid.forEach((row, i) => { row.forEach((text, j) => { put(rows[i], at.c + j, text); }); maxC = Math.max(maxC, row.length); });
  return { changes, resultRange: { r1: at.r, c1: at.c, r2: rows[rows.length - 1] ?? at.r, c2: at.c + Math.max(0, maxC - 1) } };
}

export { formatValue, editText, xfWith, isError };
export type { StylePatch, Value };
