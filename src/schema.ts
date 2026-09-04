// ProseMirror schema modelled closely on WordprocessingML so documents round-trip.
import { Schema, NodeSpec, MarkSpec, Node as PMNode, Mark } from "prosemirror-model";
import { ctx } from "./docx/styles";
import { ParaProps, RunProps, TableProps, RowProps, CellProps, Border, runCss, paraCss, merge, borderCss, cssColor, cssBgColor, cssTextColor, fontStack } from "./docx/props";
import { twipsToPx, fmt, twipsToPt } from "./docx/units";

export interface SectProps {
  pgW: number; pgH: number; // twips
  marT: number; marR: number; marB: number; marL: number;
  header: number; footer: number; gutter: number;
  orient: string | null;
  titlePg: boolean;
  headers: { default?: string; first?: string; even?: string };
  footers: { default?: string; first?: string; even?: string };
  cols: number;
  xml: string | null;
}

export const DEFAULT_SECT: SectProps = {
  pgW: 11906, pgH: 16838, marT: 1417, marR: 1417, marB: 1417, marL: 1417,
  header: 708, footer: 708, gutter: 0, orient: null, titlePg: false, headers: {}, footers: {}, cols: 1, xml: null,
};

// ---------------------------------------------------------------------------
// CSS caches (cleared when a new document context is installed)
let runCssCache = new Map<string, string>();
export function clearCssCaches() { runCssCache = new Map(); }

export function effectiveRunProps(props: RunProps): RunProps {
  if (props.rStyle) return merge(ctx.resolveCharStyle(props.rStyle), props);
  return props;
}

function runStyle(mark: Mark): string {
  const props = mark.attrs.props as RunProps;
  const key = mark.attrs.xml ?? ("j:" + JSON.stringify(props));
  let css = runCssCache.get(key);
  if (css === undefined) {
    css = runCss(effectiveRunProps(props));
    runCssCache.set(key, css);
  }
  return css;
}

export function paragraphStyle(node: PMNode): { css: string; pPr: ParaProps; rPr: RunProps } {
  const a = node.attrs;
  const eff = ctx.effectivePara(a.props as ParaProps, a.tblStyle, !!a.inTable);
  let rPr = { ...eff.rPr };
  if (node.content.size === 0 && (a.props as ParaProps).rPr) rPr = merge(rPr, (a.props as ParaProps).rPr!);
  delete rPr.vanish; delete rPr.highlight; delete rPr.shd; delete rPr.vertAlign;
  const css = runCss(rPr) + paraCss(eff.pPr, rPr.font, rPr.size || 22, !!a.inTable);
  return { css, pPr: eff.pPr, rPr };
}

// ---------------------------------------------------------------------------
// Nodes

const paragraph: NodeSpec = {
  content: "inline*",
  group: "block",
  attrs: {
    pPr: { default: null },        // original w:pPr XML (string) for round-trip
    props: { default: {} },        // parsed direct ParaProps
    tblStyle: { default: null },   // table style id when inside a table
    inTable: { default: false },
    sdt: { default: null },        // {id, pr} block content-control wrapper
    sectPr: { default: null },     // raw w:sectPr for a section break in this paragraph
    src: { default: null },        // index into LoadedDoc.origEls: the original w:p this came from
  },
  parseDOM: [
    { tag: "p", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement) },
    { tag: "h1", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement, "Heading1") },
    { tag: "h2", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement, "Heading2") },
    { tag: "h3", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement, "Heading3") },
    { tag: "h4", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement, "Heading4") },
    { tag: "h5", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement, "Heading5") },
    { tag: "h6", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement, "Heading6") },
    { tag: "li", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement) },
    { tag: "div", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement) },
    { tag: "blockquote", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement, "Quote") },
    { tag: "pre", preserveWhitespace: "full", getAttrs: (dom) => attrsFromHtmlBlock(dom as HTMLElement, "Code") },
  ],
  toDOM(node) {
    const { css, pPr } = paragraphStyle(node);
    const attrs: Record<string, string> = { class: "om-p", style: css };
    if (pPr.numId) { attrs.class += " om-num"; attrs["data-num"] = `${pPr.numId}/${pPr.ilvl || 0}`; }
    if (pPr.keepNext) attrs["data-keepnext"] = "1";
    if (pPr.keepLines) attrs["data-keeplines"] = "1";
    if (pPr.pageBreakBefore) attrs["data-pbb"] = "1";
    if (pPr.contextual) attrs["data-ctx"] = "1";
    if (pPr.pStyle) attrs["data-style"] = pPr.pStyle;
    if (node.attrs.sectPr) attrs.class += " om-sectbreak";
    if (pPr.tabs && pPr.tabs.length) attrs["data-tabs"] = pPr.tabs.map((t) => `${t.pos}:${t.val}`).join(";");
    if (pPr.indHanging && pPr.indLeft) attrs["data-hang"] = String(pPr.indLeft);
    if (pPr.indRight) attrs["data-indr"] = String(pPr.indRight);
    return ["p", attrs, 0];
  },
};

function attrsFromHtmlBlock(dom: HTMLElement, style?: string) {
  const props: ParaProps = {};
  if (style) props.pStyle = style;
  const ps = dom.getAttribute("data-pstyle");
  if (ps) props.pStyle = ctx.styles.has(ps) ? ps : (ctx.styleIdByName(ps.replace(/(\d)$/, " $1")) || ctx.styleIdByName(ps) || ps);
  if (dom.getAttribute("data-hr")) props.bdrBottom = { val: "single", sz: 6, color: "A0A0A0", space: 1 };
  const num = dom.getAttribute("data-num");
  if (num) {
    const [id, lvl] = num.split("/");
    props.numId = parseInt(id, 10); props.ilvl = parseInt(lvl || "0", 10);
    const ls = ctx.styleIdByName("List Paragraph");
    if (ls && !props.pStyle) props.pStyle = ls;
  }
  const align = dom.style?.textAlign || dom.getAttribute("align");
  if (align === "center") props.jc = "center";
  else if (align === "right") props.jc = "right";
  else if (align === "justify") props.jc = "both";
  return { props };
}

const table: NodeSpec = {
  content: "table_row+",
  group: "block",
  tableRole: "table",
  isolating: true,
  attrs: {
    tblPr: { default: null },   // raw XML
    props: { default: {} },     // TableProps
    grid: { default: null },    // number[] twips
    sdt: { default: null },
    src: { default: null },     // index into LoadedDoc.origEls (original w:tbl)
  },
  parseDOM: [{ tag: "table", getAttrs: () => ({ props: { tblStyle: ctx.styleIdByName("Table Grid") || null }, grid: null }) }],
  toDOM(node) {
    const props = node.attrs.props as TableProps;
    const st = ctx.resolveTableStyle(props.tblStyle);
    const borders = { ...(st.tblPr.borders || {}), ...(props.borders || {}) };
    const mar = { ...(st.tblPr.cellMar || {}), ...(props.cellMar || {}) };
    let style = "";
    const bv = (b: Border | undefined) => { const c = borderCss(b); return c && c !== "none" ? c : "none"; };
    style += `--bt:${bv(borders.top)};--bb:${bv(borders.bottom)};--bl:${bv(borders.left)};--br:${bv(borders.right)};`;
    style += `--bih:${bv(borders.insideH)};--biv:${bv(borders.insideV)};`;
    style += `--cpt:${fmt(twipsToPt(mar.top ?? 0))}pt;--cpb:${fmt(twipsToPt(mar.bottom ?? 0))}pt;--cpl:${fmt(twipsToPt(mar.left ?? 108))}pt;--cpr:${fmt(twipsToPt(mar.right ?? 108))}pt;`;
    const grid: number[] | null = node.attrs.grid;
    const total = grid ? grid.reduce((a, b) => a + b, 0) : 0;
    if (props.width && props.width.type === "pct") style += `width:${fmt(props.width.w / 50)}%;`;
    else if (props.width && props.width.type === "dxa" && props.width.w > 0) style += `width:${fmt(twipsToPx(props.width.w))}px;`;
    else if (total) style += `width:${fmt(twipsToPx(total))}px;`;
    if (props.jc === "center") style += "margin-left:auto;margin-right:auto;";
    else if (props.jc === "right" || props.jc === "end") style += "margin-left:auto;";
    else {
      const ind = (props.indent ?? st.tblPr.indent ?? 0) - (mar.left ?? 108);
      if (ind) style += `margin-left:${fmt(twipsToPx(ind))}px;`;
    }
    if (props.shd) style += `background-color:${cssBgColor(props.shd)};`;
    const cols: any[] = ["colgroup"];
    if (grid) for (const w of grid) cols.push(["col", { style: `width:${fmt(twipsToPx(w))}px` }]);
    const layout = props.layout === "fixed" || !!grid ? "fixed" : "auto";
    return ["table", { class: "om-tbl", style: style + `table-layout:${layout};` }, cols, ["tbody", 0]];
  },
};

const table_row: NodeSpec = {
  content: "table_cell+",
  tableRole: "row",
  attrs: {
    trPr: { default: null },
    props: { default: {} }, // RowProps
    src: { default: null }, // index into LoadedDoc.origEls (original w:tr)
  },
  parseDOM: [{ tag: "tr", getAttrs: () => ({ props: {} }) }],
  toDOM(node) {
    const p = node.attrs.props as RowProps;
    let style = "";
    if (p.height) style += `height:${fmt(twipsToPx(p.height))}px;`;
    const attrs: Record<string, string> = { style };
    if (p.header) attrs["data-header"] = "1";
    if (p.cantSplit) attrs["data-cantsplit"] = "1";
    return ["tr", attrs, 0];
  },
};

const table_cell: NodeSpec = {
  content: "block+",
  tableRole: "cell",
  isolating: true,
  attrs: {
    colspan: { default: 1 },
    rowspan: { default: 1 },
    colwidth: { default: null },
    tcPr: { default: null },
    props: { default: {} },     // CellProps
    vmergeTcPr: { default: null }, // raw tcPr of the continuation cells swallowed by rowspan
  },
  parseDOM: [
    { tag: "td", getAttrs: (dom) => cellAttrsFromDom(dom as HTMLElement) },
    { tag: "th", getAttrs: (dom) => cellAttrsFromDom(dom as HTMLElement) },
  ],
  toDOM(node) {
    const p = node.attrs.props as CellProps;
    let style = "";
    if (p.shd) style += `background-color:${cssBgColor(p.shd)};`;
    if (p.vAlign === "center") style += "vertical-align:middle;";
    else if (p.vAlign === "bottom") style += "vertical-align:bottom;";
    if (p.borders) {
      for (const side of ["top", "bottom", "left", "right"] as const) {
        const b = p.borders[side];
        if (b) style += `border-${side}:${borderCss(b)};`;
      }
    }
    if (p.mar) {
      if (p.mar.top !== undefined) style += `padding-top:${fmt(twipsToPt(p.mar.top))}pt;`;
      if (p.mar.bottom !== undefined) style += `padding-bottom:${fmt(twipsToPt(p.mar.bottom))}pt;`;
      if (p.mar.left !== undefined) style += `padding-left:${fmt(twipsToPt(p.mar.left))}pt;`;
      if (p.mar.right !== undefined) style += `padding-right:${fmt(twipsToPt(p.mar.right))}pt;`;
    }
    if (p.noWrap) style += "white-space:nowrap;";
    const attrs: Record<string, any> = { style };
    if (node.attrs.colspan !== 1) attrs.colspan = node.attrs.colspan;
    if (node.attrs.rowspan !== 1) attrs.rowspan = node.attrs.rowspan;
    return ["td", attrs, 0];
  },
};

function cellAttrsFromDom(dom: HTMLElement) {
  const cs = parseInt(dom.getAttribute("colspan") || "1", 10) || 1;
  const rs = parseInt(dom.getAttribute("rowspan") || "1", 10) || 1;
  const props: CellProps = {};
  if (dom.tagName === "TH") { /* header cells: bold handled by run marks */ }
  const bg = dom.style?.backgroundColor;
  if (bg) { const hex = rgbToHex(bg); if (hex) props.shd = hex; }
  return { colspan: cs, rowspan: rs, props };
}

export function rgbToHex(c: string): string | null {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (m) return [m[1], m[2], m[3]].map((x) => parseInt(x, 10).toString(16).padStart(2, "0")).join("").toUpperCase();
  if (/^#[0-9a-f]{6}$/i.test(c)) return c.slice(1).toUpperCase();
  return null;
}

const image: NodeSpec = {
  inline: true,
  group: "inline",
  marks: "_",
  atom: true,
  draggable: true,
  attrs: {
    src: { default: "" },
    w: { default: 100 }, // css px
    h: { default: 100 },
    rId: { default: null },
    raw: { default: null },   // original w:drawing / w:pict XML
    kind: { default: "inline" }, // inline | anchor
    name: { default: "" },
    alt: { default: "" },
    wrap: { default: null },  // {type, hAlign, vAlign, distL...}
    media: { default: null }, // {ext, bytes} for newly inserted images
    origW: { default: null }, // px size when the drawing was parsed
    origH: { default: null },
    ext: { default: null },
  },
  parseDOM: [{
    tag: "img[src]",
    getAttrs: (dom) => {
      const el = dom as HTMLImageElement;
      const src = el.getAttribute("src") || "";
      let media: { ext: string; bytes: Uint8Array } | null = null;
      let ext: string | null = null;
      const m = /^data:image\/([a-z0-9+.-]+);base64,(.*)$/i.exec(src);
      if (m) {
        try {
          const bin = atob(m[2]);
          const bytes = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          ext = m[1].toLowerCase() === "svg+xml" ? "svg" : m[1].toLowerCase() === "jpeg" ? "jpeg" : m[1].toLowerCase();
          media = { ext, bytes };
        } catch { /* keep as plain src */ }
      }
      const w = parseFloat(el.getAttribute("width") || "") || el.width || 0;
      const h = parseFloat(el.getAttribute("height") || "") || el.height || 0;
      return { src, alt: el.getAttribute("alt") || "", w: w || 200, h: h || 150, kind: "inline", raw: null, media, ext, origW: null, origH: null, rId: null, name: el.getAttribute("alt") || "" };
    },
  }],
  toDOM(node) {
    const a = node.attrs;
    let cls = "om-img";
    let style = `width:${fmt(a.w)}px;height:${fmt(a.h)}px;`;
    const attrs: Record<string, string> = { src: a.src, alt: a.alt || "", draggable: "false" };
    if (a.kind === "anchor" && a.wrap) {
      if (a.wrap.float === "left") { cls += " om-float-left"; style += `margin-right:${a.wrap.distR ?? 8}px;`; }
      else if (a.wrap.float === "right") { cls += " om-float-right"; style += `margin-left:${a.wrap.distL ?? 8}px;`; }
      else if (a.wrap.float === "center") cls += " om-block-center";
      if (a.wrap.behind) cls += " om-behind";
      Object.assign(attrs, anchorAttrs(a.wrap));
      cls += " om-anchor";
    }
    attrs.class = cls; attrs.style = style;
    return ["img", attrs];
  },
};

/** Data attributes describing an anchored object's position; consumed by the layout pass. */
export function anchorAttrs(wrap: any): Record<string, string> {
  const o: Record<string, string> = { "data-wrap": wrap.type || "none" };
  if (wrap.relH) o["data-relh"] = wrap.relH;
  if (wrap.relV) o["data-relv"] = wrap.relV;
  if (wrap.offH !== null && wrap.offH !== undefined) o["data-offh"] = fmt(wrap.offH);
  if (wrap.offV !== null && wrap.offV !== undefined) o["data-offv"] = fmt(wrap.offV);
  if (wrap.alignH) o["data-alignh"] = wrap.alignH;
  if (wrap.alignV) o["data-alignv"] = wrap.alignV;
  if (wrap.float) o["data-float"] = wrap.float;
  if (wrap.behind) o["data-behind"] = "1";
  return o;
}

const textbox: NodeSpec = {
  inline: true,
  group: "inline",
  marks: "_",
  content: "block+",
  isolating: true,
  attrs: {
    raw: { default: null },
    w: { default: null },
    h: { default: null },
    kind: { default: "anchor" },
    style: { default: null }, // {fill, line, lineW, padL, padR, padT, padB}
    wrap: { default: null },
    contentPath: { default: "txbx" },
  },
  parseDOM: [{ tag: "span.om-textbox", getAttrs: () => ({}) }],
  toDOM(node) {
    const a = node.attrs;
    let style = "";
    if (a.w) style += `width:${fmt(a.w)}px;`;
    if (a.h) style += `min-height:${fmt(a.h)}px;`;
    const s = a.style || {};
    if (s.fill) style += `background-color:${cssBgColor(s.fill)};`;
    if (s.line) style += `border:${fmt(s.lineW || 1)}px solid ${cssTextColor(s.line)};`;
    style += `padding:${fmt(s.padT ?? 3.6)}pt ${fmt(s.padR ?? 7.2)}pt ${fmt(s.padB ?? 3.6)}pt ${fmt(s.padL ?? 7.2)}pt;`;
    let cls = "om-textbox";
    const attrs: Record<string, string> = {};
    if (a.wrap?.float === "left") cls += " om-float-left";
    else if (a.wrap?.float === "right") cls += " om-float-right";
    if (a.kind === "anchor" && a.wrap) { cls += " om-anchor"; Object.assign(attrs, anchorAttrs(a.wrap)); }
    attrs.class = cls; attrs.style = style;
    return ["span", attrs, ["div", { class: "om-textbox-content" }, 0]];
  },
};

const opaque_inline: NodeSpec = {
  inline: true,
  group: "inline",
  marks: "_",
  atom: true,
  selectable: false,
  attrs: {
    xml: { default: "" },
    text: { default: "" },
    kind: { default: "marker" }, // marker | field | sdt | object | shape
    w: { default: null },
    h: { default: null },
    field: { default: null },   // instruction for simple fields (PAGE, NUMPAGES...)
  },
  toDOM(node) {
    const a = node.attrs;
    if (a.kind === "marker") return ["span", { class: "om-marker" }];
    if (a.kind === "shape") {
      return ["span", { class: "om-shape", style: `width:${fmt(a.w || 40)}px;height:${fmt(a.h || 20)}px;`, title: a.text || "Drawing (preserved, not editable)" }];
    }
    const cls = a.kind === "field" ? "om-field" : "om-opaque-inline";
    const attrs: Record<string, string> = { class: cls, title: a.field ? "Field: " + a.field : "Preserved content (not editable)" };
    if (a.field) attrs["data-field"] = String(a.field).split(/\s+/)[0].toUpperCase();
    return ["span", attrs, a.text || ""];
  },
};

const opaque_block: NodeSpec = {
  group: "block",
  atom: true,
  selectable: true,
  attrs: { xml: { default: "" }, label: { default: "Unsupported content (preserved)" } },
  toDOM(node) { return ["div", { class: "om-opaque-block", contenteditable: "false" }, node.attrs.label]; },
};

const hard_break: NodeSpec = {
  inline: true,
  group: "inline",
  marks: "_",
  selectable: false,
  attrs: { kind: { default: "textWrapping" }, clear: { default: null } },
  parseDOM: [{ tag: "br", getAttrs: () => ({ kind: "textWrapping" }) }],
  toDOM(node) {
    if (node.attrs.kind === "page") return ["span", { class: "om-pagebreak", contenteditable: "false" }, ["span", { class: "om-pagebreak-label" }, "Page Break"]];
    if (node.attrs.kind === "column") return ["span", { class: "om-colbreak", contenteditable: "false" }];
    return ["br"];
  },
};

const tab: NodeSpec = {
  inline: true,
  group: "inline",
  marks: "_",
  selectable: false,
  parseDOM: [{ tag: "span.om-tab" }],
  toDOM() { return ["span", { class: "om-tab" }, "\t"]; },
};

const text: NodeSpec = { group: "inline" };

// ---------------------------------------------------------------------------
// Marks

const rpr: MarkSpec = {
  attrs: {
    xml: { default: null },   // original w:rPr XML
    props: { default: {} },   // RunProps
  },
  parseDOM: [
    { tag: "span[data-rpr]", getAttrs: (dom) => ({ props: JSON.parse((dom as HTMLElement).getAttribute("data-rpr") || "{}") }) },
  ],
  toDOM(mark) {
    return ["span", { class: "om-r", style: runStyle(mark) }, 0];
  },
};

const link: MarkSpec = {
  attrs: {
    href: { default: "" },
    rId: { default: null },
    anchor: { default: null },
    tooltip: { default: null },
    raw: { default: null }, // original attribute string of w:hyperlink
  },
  inclusive: false,
  parseDOM: [{ tag: "a[href]", getAttrs: (dom) => ({ href: (dom as HTMLElement).getAttribute("href") }) }],
  toDOM(mark) {
    const a = mark.attrs;
    return ["a", { href: a.href || (a.anchor ? "#" + a.anchor : "#"), class: "om-link", title: a.tooltip || a.href || "", target: "_blank", rel: "noopener" }, 0];
  },
};

export const schema = new Schema({
  nodes: {
    doc: { content: "block+", attrs: { sect: { default: DEFAULT_SECT } } },
    paragraph,
    table,
    table_row,
    table_cell,
    image,
    textbox,
    opaque_inline,
    opaque_block,
    hard_break,
    tab,
    text,
  },
  marks: { link, rpr },
});

export type OMSchema = typeof schema;

export const bulletFontStack = fontStack;
