// Toolbar: buttons with tooltips (name + shortcut), font/size combos, colors.
import { EditorState, Command } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { undoDepth, redoDepth } from "prosemirror-history";
import { el, icon, tooltip, showPopup, showMenu, PopupHandle } from "./widgets";
import * as C from "../editor/commands";
import { Shortcut, keyLabel, AppActions } from "../editor/keymap";
import { ctx } from "../docx/styles";
import { schema } from "../schema";

export interface ToolbarCtx {
  view(): EditorView | null;
  run(cmd: Command, focus?: boolean): void;
  actions: AppActions;
  shortcuts: Shortcut[];
}

export interface ToolbarHandle {
  update(state: EditorState): void;
  focusFont(): void;
  focusSize(): void;
}

const FONTS = ["Aptos", "Arial", "Calibri", "Calibri Light", "Cambria", "Consolas", "Courier New", "Garamond", "Georgia", "Helvetica", "Open Sans", "Segoe UI", "Tahoma", "Times New Roman", "Verdana"];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

export const THEME_COLORS = [
  ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"],
  ["F2F2F2", "808080", "D0CECE", "D6DCE4", "D9E2F3", "FBE5D5", "EDEDED", "FFF2CC", "DEEBF6", "E2EFD9"],
  ["D8D8D8", "595959", "AEABAB", "ADB9CA", "B4C6E7", "F7CBAC", "DBDBDB", "FFE599", "BDD7EE", "C5E0B3"],
  ["BFBFBF", "3F3F3F", "757070", "8496B0", "8EAADB", "F4B183", "C9C9C9", "FFD965", "9DC3E6", "A8D08D"],
  ["A5A5A5", "262626", "3A3838", "323F4F", "2F5496", "C55A11", "7B7B7B", "BF9000", "2E75B5", "538135"],
  ["7F7F7F", "0C0C0C", "171616", "222A35", "1F3864", "833C0B", "525252", "7F6000", "1E4E79", "375623"],
];
export const STANDARD_COLORS = ["C00000", "FF0000", "FFC000", "FFFF00", "92D050", "00B050", "00B0F0", "0070C0", "002060", "7030A0"];
export const HIGHLIGHTS: [string, string][] = [["yellow", "#ffff00"], ["green", "#00ff00"], ["cyan", "#00ffff"], ["magenta", "#ff00ff"], ["blue", "#0000ff"], ["red", "#ff0000"], ["darkBlue", "#000080"], ["darkCyan", "#008080"], ["darkGreen", "#008000"], ["darkMagenta", "#800080"], ["darkRed", "#800000"], ["darkYellow", "#808000"], ["darkGray", "#808080"], ["lightGray", "#c0c0c0"], ["black", "#000000"]];

export function colorPopup(anchor: HTMLElement, current: string | null, onPick: (hex: string | null) => void, opts: { auto?: string; highlight?: boolean } = {}): PopupHandle {
  return showPopup(anchor, (popup, close) => {
    const pick = (hex: string | null) => { close(); onPick(hex); };
    popup.appendChild(el("div", { class: "item", onclick: () => pick(null) }, opts.auto || "Automatic"));
    if (opts.highlight) {
      const grid = el("div", { class: "color-grid" });
      for (const [name, css] of HIGHLIGHTS) grid.appendChild(el("div", { class: "color-cell", style: { background: css }, title: name, onclick: () => pick(name) }));
      popup.appendChild(grid);
      return;
    }
    const grid = el("div", { class: "color-grid" });
    for (const row of THEME_COLORS) for (const c of row) grid.appendChild(el("div", { class: "color-cell", style: { background: "#" + c }, title: "#" + c, onclick: () => pick(c) }));
    popup.appendChild(grid);
    popup.appendChild(el("div", { class: "item", style: { color: "var(--ui-muted)", fontSize: "11px", padding: "2px 10px" } }, "Standard colors"));
    const std = el("div", { class: "color-grid" });
    for (const c of STANDARD_COLORS) std.appendChild(el("div", { class: "color-cell", style: { background: "#" + c }, title: "#" + c, onclick: () => pick(c) }));
    popup.appendChild(std);
    const input = el("input", { type: "color", value: "#" + (current && /^[0-9A-Fa-f]{6}$/.test(current) ? current : "000000") });
    input.addEventListener("input", () => onPick(input.value.slice(1).toUpperCase()));
    input.addEventListener("change", () => close());
    popup.appendChild(el("div", { class: "row" }, el("span", null, "More colors…"), input));
  });
}

export function buildToolbar(container: HTMLElement, t: ToolbarCtx): ToolbarHandle {
  container.innerHTML = "";
  const key = (id: string) => { const s = t.shortcuts.find((x) => x.id === id); return s ? keyLabel(s.keys[0]) : undefined; };
  const buttons = new Map<string, HTMLButtonElement>();

  const btn = (id: string, ic: string, label: string, onClick: () => void, opts: { split?: () => void; text?: string } = {}) => {
    const b = el("button", { class: "tb-btn", type: "button", tabindex: -1 }, opts.text ? el("span", { style: { fontWeight: "600" } }, opts.text) : icon(ic));
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep editor focus
    b.addEventListener("click", onClick);
    tooltip(b, label, key(id));
    buttons.set(id, b);
    container.appendChild(b);
    return b;
  };
  const sep = () => container.appendChild(el("div", { class: "tb-sep" }));
  const cmdBtn = (id: string, ic: string, label: string, cmd: Command) => btn(id, ic, label, () => t.run(cmd));

  btn("new", "new", "New document", () => t.actions.newDocument());
  btn("open", "open", "Open", () => t.actions.openFile());
  btn("save", "save", "Save", () => t.actions.save());
  btn("print", "print", "Print", () => t.actions.print());
  sep();
  cmdBtn("undo", "undo", "Undo", (s, d) => { const { undo } = require_history(); return undo(s, d); });
  cmdBtn("redo", "redo", "Redo", (s, d) => { const { redo } = require_history(); return redo(s, d); });
  sep();

  // Paragraph style
  const styleSel = el("select", { class: "tb-select", title: "Paragraph style", tabindex: -1 });
  styleSel.addEventListener("mousedown", () => refreshStyles());
  styleSel.addEventListener("change", () => { const v = styleSel.value; t.run(C.setParaStyle(v === "__normal" ? null : v)); });
  tooltip(styleSel, "Paragraph style");
  container.appendChild(styleSel);
  const refreshStyles = () => {
    const cur = styleSel.value;
    styleSel.innerHTML = "";
    const add = (id: string, name: string) => styleSel.appendChild(el("option", { value: id }, name));
    add("__normal", "Normal");
    const wanted = /^(Heading[1-6]|Title|Subtitle|Quote|IntenseQuote|ListParagraph|Code|NoSpacing|Caption)$/i;
    const seen = new Set<string>();
    for (const s of ctx.styles.values()) {
      if (s.type !== "paragraph" || s.id === ctx.defaultPara) continue;
      if (wanted.test(s.id) || /^(heading \d|title|subtitle|quote|list paragraph|no spacing|caption)$/i.test(s.name)) { add(s.id, s.name.replace(/^heading (\d)/i, "Heading $1")); seen.add(s.id); }
    }
    for (const id of ["Heading1", "Heading2", "Heading3"]) if (!seen.has(id) && !ctx.styleIdByName("heading " + id.slice(-1))) add(id, "Heading " + id.slice(-1));
    styleSel.value = cur;
  };
  refreshStyles();

  // Font family combo
  const fontInput = el("input", { type: "text", spellcheck: "false", autocomplete: "off" });
  const fontBtn = el("button", { type: "button", tabindex: -1 }, "▾");
  const fontCombo = el("div", { class: "tb-combo", id: "font-family" }, fontInput, fontBtn);
  tooltip(fontCombo, "Font", key("fontbox"));
  container.appendChild(fontCombo);
  const applyFont = () => { const v = fontInput.value.trim(); if (v) t.run(C.setFont(v)); };
  fontInput.addEventListener("keydown", (e) => { if (e.key === "Enter") { applyFont(); e.preventDefault(); } else if (e.key === "Escape") { t.view()?.focus(); } });
  fontInput.addEventListener("focus", () => fontInput.select());
  const fontList = () => {
    const used = new Set<string>(FONTS);
    const v = t.view(); if (v) v.state.doc.descendants((n) => { for (const m of n.marks) if (m.type === schema.marks.rpr && m.attrs.props.font) used.add(m.attrs.props.font); return true; });
    for (const s of ctx.styles.values()) if (s.rPr.font) used.add(s.rPr.font);
    if (ctx.docRPr.font) used.add(ctx.docRPr.font);
    return Array.from(used).sort((a, b) => a.localeCompare(b));
  };
  fontBtn.addEventListener("mousedown", (e) => e.preventDefault());
  fontBtn.addEventListener("click", () => showPopup(fontCombo, (popup, close) => {
    for (const f of fontList()) popup.appendChild(el("div", { class: "item", style: { fontFamily: `"${f}"` }, onclick: () => { close(); t.run(C.setFont(f)); } }, f));
  }));

  // Font size combo
  const sizeInput = el("input", { type: "text", spellcheck: "false", autocomplete: "off" });
  const sizeBtn = el("button", { type: "button", tabindex: -1 }, "▾");
  const sizeCombo = el("div", { class: "tb-combo", id: "font-size" }, sizeInput, sizeBtn);
  tooltip(sizeCombo, "Font size", key("sizebox"));
  container.appendChild(sizeCombo);
  const applySize = () => { const v = parseFloat(sizeInput.value.replace(",", ".")); if (v > 0 && v <= 400) t.run(C.setFontSize(v)); };
  sizeInput.addEventListener("keydown", (e) => { if (e.key === "Enter") { applySize(); e.preventDefault(); } else if (e.key === "Escape") t.view()?.focus(); });
  sizeInput.addEventListener("focus", () => sizeInput.select());
  sizeBtn.addEventListener("mousedown", (e) => e.preventDefault());
  sizeBtn.addEventListener("click", () => showPopup(sizeCombo, (popup, close) => {
    for (const s of SIZES) popup.appendChild(el("div", { class: "item", onclick: () => { close(); t.run(C.setFontSize(s)); } }, String(s)));
  }));
  btn("sizeup", "", "Grow font", () => t.run(C.fontSizeStep(1)), { text: "A↑" });
  btn("sizedown", "", "Shrink font", () => t.run(C.fontSizeStep(-1)), { text: "A↓" });
  sep();

  cmdBtn("bold", "bold", "Bold", C.toggleBold);
  cmdBtn("italic", "italic", "Italic", C.toggleItalic);
  cmdBtn("underline", "underline", "Underline", C.toggleUnderline);
  cmdBtn("strike", "strike", "Strikethrough", C.toggleStrike);
  cmdBtn("sup", "sup", "Superscript", C.toggleSuperscript);
  cmdBtn("sub", "sub", "Subscript", C.toggleSubscript);

  // Text color (split button: apply last color / pick)
  let lastColor = "C00000";
  const colorB = btn("color", "color", "Font color", () => t.run(C.setColor(lastColor)));
  colorB.classList.add("tb-color");
  const colorSwatch = el("span", { class: "swatch", style: { background: "#" + lastColor } });
  colorB.appendChild(colorSwatch);
  const colorDD = el("span", { class: "tb-dd" }, "▾");
  colorDD.addEventListener("mousedown", (e) => e.preventDefault());
  colorDD.addEventListener("click", () => colorPopup(colorB, lastColor, (hex) => { if (hex) { lastColor = hex; colorSwatch.style.background = "#" + hex; } t.run(C.setColor(hex)); }));
  tooltip(colorDD, "Choose font color");
  container.appendChild(colorDD);

  let lastHl = "yellow";
  const hlB = btn("highlight", "highlight", "Text highlight", () => t.run(C.setHighlight(lastHl)));
  hlB.classList.add("tb-color");
  const hlSwatch = el("span", { class: "swatch", style: { background: "#ffff00" } });
  hlB.appendChild(hlSwatch);
  const hlDD = el("span", { class: "tb-dd" }, "▾");
  hlDD.addEventListener("mousedown", (e) => e.preventDefault());
  hlDD.addEventListener("click", () => colorPopup(hlB, null, (name) => { if (name) { lastHl = name; hlSwatch.style.background = HIGHLIGHTS.find((h) => h[0] === name)?.[1] || "#ffff00"; } t.run(C.setHighlight(name)); }, { auto: "No color", highlight: true }));
  tooltip(hlDD, "Choose highlight color");
  container.appendChild(hlDD);
  cmdBtn("clear", "clear", "Clear formatting", C.clearFormatting);
  sep();

  cmdBtn("left", "alignLeft", "Align left", C.setAlign("left"));
  cmdBtn("center", "alignCenter", "Center", C.setAlign("center"));
  cmdBtn("right", "alignRight", "Align right", C.setAlign("right"));
  cmdBtn("justify", "alignJustify", "Justify", C.setAlign("both"));
  sep();
  cmdBtn("bullets", "bullets", "Bulleted list", C.toggleList("bullet"));
  cmdBtn("numbers", "numbers", "Numbered list", C.toggleList("decimal"));
  cmdBtn("outdent", "outdent", "Decrease indent", C.indentParagraphs(-1));
  cmdBtn("indent", "indent", "Increase indent", C.indentParagraphs(1));
  const spB = btn("spacing", "spacing", "Line and paragraph spacing", () => showMenu(spB, [
    { label: "1.0", action: () => t.run(C.setLineSpacing(1)), key: key("single") },
    { label: "1.15", action: () => t.run(C.setLineSpacing(1.15)) },
    { label: "1.5", action: () => t.run(C.setLineSpacing(1.5)), key: key("onehalf") },
    { label: "2.0", action: () => t.run(C.setLineSpacing(2)), key: key("double") },
    { sep: true },
    { label: "Remove space before paragraph", action: () => t.run(C.setSpaceBefore(0)) },
    { label: "Add space before paragraph (12 pt)", action: () => t.run(C.setSpaceBefore(12)) },
    { label: "Remove space after paragraph", action: () => t.run(C.setSpaceAfter(0)) },
    { label: "Add space after paragraph (8 pt)", action: () => t.run(C.setSpaceAfter(8)) },
    { sep: true },
    { label: "Paragraph…", action: () => (t.actions as any).paragraphDialog?.() },
  ]));
  sep();
  btn("image", "image", "Insert image", () => t.actions.insertImage());
  btn("table", "table", "Insert table", () => t.actions.insertTable());
  btn("link", "link", "Insert link", () => t.actions.insertLink());
  cmdBtn("pagebreak", "pagebreak", "Page break", C.insertPageBreak);
  sep();
  btn("find", "find", "Find and replace", () => t.actions.find());

  const setActive = (id: string, on: boolean) => buttons.get(id)?.classList.toggle("active", on);

  return {
    update(state) {
      const rp = C.selectionRunProps(state);
      setActive("bold", !!rp.b); setActive("italic", !!rp.i); setActive("underline", !!rp.u && rp.u !== "none");
      setActive("strike", !!rp.strike); setActive("sup", rp.vertAlign === "superscript"); setActive("sub", rp.vertAlign === "subscript");
      if (document.activeElement !== fontInput) fontInput.value = rp.font || ctx.docRPr.font || "";
      if (document.activeElement !== sizeInput) sizeInput.value = rp.size ? String(rp.size / 2) : "";
      const pp = C.selectionParaProps(state).eff;
      const jc = pp.jc || "left";
      setActive("left", jc === "left" || jc === "start"); setActive("center", jc === "center"); setActive("right", jc === "right" || jc === "end"); setActive("justify", jc === "both" || jc === "distribute");
      setActive("bullets", !!pp.numId && ctx.isBulletList(pp.numId)); setActive("numbers", !!pp.numId && !ctx.isBulletList(pp.numId));
      const st = pp.pStyle && pp.pStyle !== ctx.defaultPara ? pp.pStyle : "__normal";
      if (styleSel.value !== st) { if (!Array.from(styleSel.options).some((o) => o.value === st)) styleSel.appendChild(el("option", { value: st }, ctx.style(st)?.name || st)); styleSel.value = st; }
      buttons.get("undo")!.disabled = undoDepth(state) === 0;
      buttons.get("redo")!.disabled = redoDepth(state) === 0;
    },
    focusFont() { fontInput.focus(); fontInput.select(); },
    focusSize() { sizeInput.focus(); sizeInput.select(); },
  };
}

import { undo, redo } from "prosemirror-history";
function require_history() { return { undo, redo }; }
