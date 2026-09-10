// Printing a worksheet: the used range (or a selection) is rendered as an HTML
// table with the cell styles, frozen rows become a repeating <thead>, and the
// page is scaled to fit the paper width when asked.
import { Sheet, Range, key, colWidthChars, colHidden, charsToPx, rowHeightPx, mergeAt, inRange } from "./model";
import { StyleResolver } from "./render-style";
import { fontFamilyCss } from "./render-style";
import { el } from "../ui/widgets";
import { showDialog } from "../ui/dialog-core";
import { setDarkMode, isDarkMode } from "../docx/props";

export interface PrintOptions { range: "sheet" | "selection"; orientation: "portrait" | "landscape"; fit: boolean; gridlines: boolean; headers: boolean; paper: "A4" | "Letter"; repeatHeader: boolean; }

let lastOptions: PrintOptions = { range: "sheet", orientation: "portrait", fit: true, gridlines: false, headers: false, paper: "A4", repeatHeader: true };

export function printDialog(hasSelection: boolean): Promise<PrintOptions | null> {
  return new Promise((resolve) => {
    const o = { ...lastOptions };
    const sel = el("select", null, el("option", { value: "sheet" }, "Whole sheet"), el("option", { value: "selection", disabled: !hasSelection || undefined }, "Selected cells"));
    sel.value = hasSelection ? o.range : "sheet";
    const orient = el("select", null, el("option", { value: "portrait" }, "Portrait"), el("option", { value: "landscape" }, "Landscape"));
    orient.value = o.orientation;
    const paper = el("select", null, el("option", { value: "A4" }, "A4"), el("option", { value: "Letter" }, "Letter"));
    paper.value = o.paper;
    const cb = (label: string, v: boolean) => { const c = el("input", { type: "checkbox" }); c.checked = v; return { el: el("label", { style: { display: "flex", alignItems: "center", gap: "6px", color: "var(--ui-fg)" } }, c, label), c }; };
    const fit = cb("Fit to page width", o.fit), grid = cb("Print gridlines", o.gridlines), heads = cb("Print row and column headers", o.headers), rep = cb("Repeat frozen rows on every page", o.repeatHeader);
    const body = el("div", null,
      el("div", { class: "grid2" }, el("label", null, "Print"), sel, el("label", null, "Orientation"), orient, el("label", null, "Paper"), paper),
      el("div", { style: { marginTop: "10px", display: "flex", flexDirection: "column", gap: "4px" } }, fit.el, grid.el, heads.el, rep.el));
    let done = false;
    showDialog("Print", body, [
      { label: "Cancel" },
      { label: "Print", primary: true, action: () => { done = true; lastOptions = { range: sel.value as any, orientation: orient.value as any, paper: paper.value as any, fit: fit.c.checked, gridlines: grid.c.checked, headers: heads.c.checked, repeatHeader: rep.c.checked }; resolve(lastOptions); } },
    ], { onClose: () => { if (!done) resolve(null); } });
  });
}

function esc(s: string) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br>"); }

export function printSheet(sheet: Sheet, styles: StyleResolver, range: Range | null, o: PrintOptions, date1904: boolean): Promise<void> {
  const wasDark = isDarkMode();
  if (wasDark) { setDarkMode(false); styles.invalidate(); }
  const rg: Range = range || { r1: 0, c1: 0, r2: Math.max(sheet.maxRow, 0), c2: Math.max(sheet.maxCol, 0) };
  const cols: number[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) if (!colHidden(sheet, c)) cols.push(c);
  const totalW = cols.reduce((a, c) => a + charsToPx(colWidthChars(sheet, c)), 0) + (o.headers ? 40 : 0);
  const paperW = (o.paper === "A4" ? 210 : 215.9) - 20, paperH = (o.paper === "A4" ? 297 : 279.4) - 20; // mm minus margins
  const availW = (o.orientation === "landscape" ? paperH : paperW) * 96 / 25.4;
  const scale = o.fit && totalW > availW ? availW / totalW : 1;
  const frozen = o.repeatHeader && sheet.freeze && !sheet.autoFreeze && rg.r1 === 0 ? Math.min(sheet.freeze.rows, rg.r2 - rg.r1) : 0;
  const skip = new Set<number>();
  const cellHtml = (r: number, c: number): string => {
    const k = key(r, c);
    if (skip.has(k)) return "";
    const m = mergeAt(sheet, r, c);
    let span = "";
    let cell = sheet.cells.get(k);
    if (m) {
      // the first printed cell of the merge acts as its anchor (the real anchor may lie outside the printed range)
      const ar = Math.max(m.r1, rg.r1), ac = cols.find((cc) => cc >= m.c1) ?? c;
      if (ar !== r || ac !== c) return "";
      const rowsIn: number[] = []; for (let rr = r; rr <= Math.min(m.r2, rg.r2); rr++) if (rowHeightPx(sheet, rr) > 0) rowsIn.push(rr);
      const rs = Math.max(1, rowsIn.length), cs = cols.filter((cc) => cc >= c && cc <= m.c2).length;
      span = `${rs > 1 ? ` rowspan="${rs}"` : ""}${cs > 1 ? ` colspan="${cs}"` : ""}`;
      for (let rr = r; rr <= Math.min(m.r2, rg.r2); rr++) for (const cc of cols) if (cc >= c && cc <= m.c2 && (rr !== r || cc !== c)) skip.add(key(rr, cc));
      cell = sheet.cells.get(key(m.r1, m.c1));
    }
    const cs = styles.get(cell ? cell.s : 0);
    const rd = cell ? styles.render(cell, cs) : { text: "", color: null, align: null };
    const st: string[] = [];
    st.push(`font:${cs.italic ? "italic " : ""}${cs.bold ? "bold " : ""}${cs.fontSize}pt ${fontFamilyCss(cs.fontName)}`);
    if (rd.color) st.push("color:" + rd.color); else if (cs.color) st.push("color:" + cs.color);
    if (cs.fill) st.push("background:" + cs.fill);
    const align = rd.align || (cs.halign === "general" ? (cell && (typeof cell.v === "number" || typeof cell.v === "boolean") ? "right" : "left") : cs.halign);
    st.push("text-align:" + (align === "centerContinuous" ? "center" : align));
    st.push("vertical-align:" + (cs.valign === "center" ? "middle" : cs.valign));
    if (cs.wrap) st.push("white-space:pre-wrap"); else st.push("white-space:pre;overflow:hidden");
    if (cs.underline) st.push("text-decoration:underline"); if (cs.strike) st.push("text-decoration:line-through");
    if (cs.indent) st.push(`padding-left:${2 + cs.indent * 8}px`);
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const b = cs.borders[side];
      if (b) st.push(`border-${side}:${b.style === "medium" || b.style === "thick" || b.style === "double" ? (b.style === "double" ? "3px double" : b.style === "thick" ? "2.5px solid" : "1.5px solid") : b.style === "dashed" || b.style === "mediumDashed" ? "1px dashed" : b.style === "dotted" || b.style === "hair" ? "1px dotted" : "1px solid"} ${b.color}`);
    }
    return `<td${span} style="${st.join(";")}">${esc(rd.text)}</td>`;
  };
  const rowHtml = (r: number) => {
    if (rowHeightPx(sheet, r) === 0) return "";
    let h = `<tr style="height:${rowHeightPx(sheet, r)}px">`;
    if (o.headers) h += `<th class="rh">${r + 1}</th>`;
    for (const c of cols) h += cellHtml(r, c);
    return h + "</tr>";
  };
  let html = `<table class="ps" style="width:${totalW}px">`;
  html += "<colgroup>" + (o.headers ? '<col style="width:40px">' : "") + cols.map((c) => `<col style="width:${charsToPx(colWidthChars(sheet, c))}px">`).join("") + "</colgroup>";
  if (o.headers || frozen) {
    html += "<thead>";
    if (o.headers) html += "<tr>" + '<th class="rh"></th>' + cols.map((c) => `<th class="ch">${colLetter(c)}</th>`).join("") + "</tr>";
    for (let r = rg.r1; r < rg.r1 + frozen; r++) html += rowHtml(r);
    html += "</thead>";
  }
  html += "<tbody>";
  for (let r = rg.r1 + frozen; r <= rg.r2; r++) html += rowHtml(r);
  html += "</tbody></table>";
  const wrap = document.createElement("div");
  wrap.id = "print-sheet";
  wrap.innerHTML = html;
  wrap.style.zoom = String(scale);
  const style = document.createElement("style");
  style.id = "print-sheet-style";
  style.textContent = `
    #print-sheet { display: none; }
    @media print {
      @page { size: ${o.paper} ${o.orientation}; margin: 10mm; }
      body > #app, #overlay, #tooltip { display: none !important; }
      html, body { overflow: visible !important; height: auto !important; background: #fff !important; }
      #print-sheet { display: block; color: #000; }
      #print-sheet table.ps { border-collapse: collapse; table-layout: fixed; }
      #print-sheet td, #print-sheet th { padding: 1px 3px; box-sizing: border-box; ${o.gridlines ? "border: 1px solid #c8c8c8;" : ""} }
      #print-sheet th { font: 9pt ${fontFamilyCss("Calibri")}; color: #555; background: #f0f0f0; border: 1px solid #c8c8c8; text-align: center; }
      #print-sheet thead { display: table-header-group; }
      #print-sheet tr { page-break-inside: avoid; }
      #print-sheet, #print-sheet * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }`;
  document.head.appendChild(style);
  document.body.appendChild(wrap);
  return new Promise<void>((resolve) => {
    let finished = false;
    const done = () => {
      if (finished) return; finished = true;
      window.removeEventListener("afterprint", done);
      wrap.remove(); style.remove();
      if (wasDark) { setDarkMode(true); styles.invalidate(); }
      resolve();
    };
    window.addEventListener("afterprint", done);
    // Some webviews never fire afterprint, and a print dialog can stay open for a long time: tear down only when
    // the print media query goes back to screen, or after a long safety timeout.
    const mq = window.matchMedia("print");
    let wasPrint = false;
    const onMq = (ev: MediaQueryListEvent) => { if (ev.matches) wasPrint = true; else if (wasPrint) { mq.removeEventListener("change", onMq); done(); } };
    mq.addEventListener("change", onMq);
    requestAnimationFrame(() => requestAnimationFrame(() => { try { window.print(); } catch { done(); } setTimeout(done, 10 * 60 * 1000); }));
  });
  void date1904; void inRange;
}

function colLetter(c: number): string { let s = ""; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
