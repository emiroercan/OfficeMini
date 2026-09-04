// Markdown <-> document, and HTML normalization used for paste and import.
import { marked } from "marked";
import { Node as PMNode, DOMParser as PMDOMParser } from "prosemirror-model";
import { schema } from "../schema";
import { ctx } from "../docx/styles";
import { RunProps, ParaProps } from "../docx/props";
import { paragraphStyle, effectiveRunProps } from "../schema";

// ---------------------------------------------------------------------------
// HTML -> normalized HTML (one <span data-rpr> per text run, flat paragraphs)

interface Inline { b?: boolean; i?: boolean; u?: string; strike?: boolean; font?: string; size?: number; color?: string; highlight?: string; shd?: string; vertAlign?: string; }

function parseColor(v: string): string | null {
  v = v.trim();
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(v);
  if (m) return [m[1], m[2], m[3]].map((x) => parseInt(x, 10).toString(16).padStart(2, "0")).join("").toUpperCase();
  if (/^#[0-9a-f]{6}$/i.test(v)) return v.slice(1).toUpperCase();
  if (/^#[0-9a-f]{3}$/i.test(v)) return v.slice(1).split("").map((c) => c + c).join("").toUpperCase();
  const named: Record<string, string> = { black: "000000", white: "FFFFFF", red: "FF0000", blue: "0000FF", green: "008000", yellow: "FFFF00", gray: "808080", grey: "808080", orange: "FFA500", purple: "800080" };
  return named[v.toLowerCase()] || null;
}

function sizeToHalfPt(v: string, base: number): number | null {
  const m = /^([\d.]+)(px|pt|em|rem|%)?$/.exec(v.trim());
  if (!m) {
    const kw: Record<string, number> = { "xx-small": 14, "x-small": 16, small: 20, medium: 22, large: 27, "x-large": 36, "xx-large": 48 };
    return kw[v.trim()] || null;
  }
  const n = parseFloat(m[1]);
  switch (m[2]) {
    case "pt": return Math.round(n * 2);
    case "em": case "rem": return Math.round(n * base);
    case "%": return Math.round((n / 100) * base);
    default: return Math.round(n * 1.5); // px -> half-points (96dpi)
  }
}

function inlineFrom(el: HTMLElement, inh: Inline): Inline {
  const out: Inline = { ...inh };
  const tag = el.tagName.toLowerCase();
  if (tag === "b" || tag === "strong") out.b = true;
  if (tag === "i" || tag === "em" || tag === "cite" || tag === "var") out.i = true;
  if (tag === "u" || tag === "ins") out.u = "single";
  if (tag === "s" || tag === "strike" || tag === "del") out.strike = true;
  if (tag === "sup") out.vertAlign = "superscript";
  if (tag === "sub") out.vertAlign = "subscript";
  if (tag === "code" || tag === "kbd" || tag === "samp" || tag === "tt") { out.font = "Consolas"; if (!el.closest("pre")) out.shd = "F2F2F2"; }
  if (tag === "mark") out.highlight = "yellow";
  if (tag === "font") {
    const face = el.getAttribute("face"); if (face) out.font = face.split(",")[0].replace(/["']/g, "").trim();
    const color = el.getAttribute("color"); const c = color && parseColor(color); if (c) out.color = c;
    const size = el.getAttribute("size"); if (size) { const map: Record<string, number> = { "1": 16, "2": 20, "3": 24, "4": 28, "5": 36, "6": 48, "7": 72 }; if (map[size]) out.size = map[size]; }
  }
  const st = el.style;
  if (st) {
    const fw = st.fontWeight; if (fw) out.b = fw === "bold" || fw === "bolder" || parseInt(fw, 10) >= 600;
    const fs = st.fontStyle; if (fs) out.i = fs === "italic" || fs === "oblique";
    const td = st.textDecorationLine || st.textDecoration; if (td) { if (/underline/.test(td)) out.u = "single"; if (/line-through/.test(td)) out.strike = true; if (td === "none") { out.u = undefined; } }
    const ff = st.fontFamily; if (ff) { const f = ff.split(",")[0].replace(/["']/g, "").trim(); if (f && !/^(serif|sans-serif|monospace|system-ui|inherit)$/i.test(f)) out.font = f; }
    const fsz = st.fontSize; if (fsz) { const s = sizeToHalfPt(fsz, out.size || 22); if (s) out.size = s; }
    const col = st.color; if (col) { const c = parseColor(col); if (c) out.color = c; }
    const bg = st.backgroundColor; if (bg && bg !== "transparent") { const c = parseColor(bg); if (c && c !== "FFFFFF") out.shd = c; }
    const va = st.verticalAlign; if (va === "super") out.vertAlign = "superscript"; else if (va === "sub") out.vertAlign = "subscript";
  }
  return out;
}

function rprJson(inl: Inline): string | null {
  const p: RunProps = {};
  if (inl.b) p.b = true; if (inl.i) p.i = true; if (inl.u) p.u = inl.u; if (inl.strike) p.strike = true;
  if (inl.font) p.font = inl.font; if (inl.size) p.size = inl.size; if (inl.color) p.color = inl.color;
  if (inl.highlight) p.highlight = inl.highlight; if (inl.shd) p.shd = inl.shd; if (inl.vertAlign) p.vertAlign = inl.vertAlign;
  return Object.keys(p).length ? JSON.stringify(p) : null;
}

const BLOCK_TAGS = new Set(["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "ul", "ol", "blockquote", "pre", "table", "tr", "td", "th", "tbody", "thead", "tfoot", "section", "article", "header", "footer", "main", "nav", "aside", "hr", "figure", "figcaption", "dl", "dt", "dd", "address", "center"]);

interface ListCtx { numId: number; level: number; }

/**
 * Rewrites HTML so that every text node is wrapped in a single <span data-rpr>
 * with its effective formatting, lists become flat numbered paragraphs, and
 * only structure the schema understands remains.
 */
export function normalizeHtml(html: string): string {
  const src = new DOMParser().parseFromString(html, "text/html");
  const out = src.createElement("div");
  const listCache = new Map<string, number>();
  const listId = (kind: "bullet" | "decimal", key: string) => {
    let id = listCache.get(key);
    if (!id) { id = ctx.createList(kind); listCache.set(key, id); }
    return id;
  };

  let curPara: HTMLElement | null = null;
  let pendingAttrs: Record<string, string> = {};
  const ensurePara = (attrs?: Record<string, string>) => {
    if (!curPara) {
      curPara = src.createElement("p");
      for (const [k, v] of Object.entries(attrs || pendingAttrs)) curPara.setAttribute(k, v);
      out.appendChild(curPara);
    }
    return curPara;
  };
  const endPara = () => { curPara = null; };

  const emitText = (text: string, inl: Inline, link: string | null, pre: boolean) => {
    if (!pre) text = text.replace(/\s+/g, " ");
    if (!text) return;
    // Whitespace between blocks must not create paragraphs.
    if (!curPara && !text.trim()) return;
    if (curPara && !curPara.childNodes.length && !text.trim()) return;
    const p = ensurePara();
    let target: HTMLElement = p;
    if (link) { const a = src.createElement("a"); a.setAttribute("href", link); p.appendChild(a); target = a; }
    if (pre) {
      const lines = text.split("\n");
      lines.forEach((line, i) => {
        if (i > 0) target.appendChild(src.createElement("br"));
        if (line) { const span = src.createElement("span"); const j = rprJson(inl); if (j) span.setAttribute("data-rpr", j); span.textContent = line; target.appendChild(span); }
      });
      return;
    }
    const span = src.createElement("span");
    const j = rprJson(inl);
    if (j) span.setAttribute("data-rpr", j);
    span.textContent = text;
    target.appendChild(span);
  };

  const walk = (node: Node, inl: Inline, link: string | null, list: ListCtx | null, pre: boolean, paraAttrs: Record<string, string>) => {
    if (node.nodeType === 3) { emitText(node.nodeValue || "", inl, link, pre); return; }
    if (node.nodeType !== 1) return;
    const el = node as HTMLElement;
    const tag = el.tagName.toLowerCase();
    if (tag === "script" || tag === "style" || tag === "head" || tag === "meta" || tag === "title" || tag === "o:p") return;
    if (tag === "br") { ensurePara(paraAttrs).appendChild(src.createElement("br")); return; }
    if (tag === "img") {
      const s = el.getAttribute("src"); if (!s) return;
      const img = src.createElement("img");
      img.setAttribute("src", s); img.setAttribute("alt", el.getAttribute("alt") || "");
      const w = parseFloat(el.getAttribute("width") || el.style.width || "0"), h = parseFloat(el.getAttribute("height") || el.style.height || "0");
      if (w) img.setAttribute("width", String(w)); if (h) img.setAttribute("height", String(h));
      ensurePara(paraAttrs).appendChild(img);
      return;
    }
    if (tag === "hr") { endPara(); const p = ensurePara({ "data-hr": "1" }); p.appendChild(src.createElement("span")); endPara(); return; }
    if (tag === "a") { link = el.getAttribute("href") || link; }
    if (tag === "table") {
      endPara();
      const t = src.createElement("table");
      const tbody = src.createElement("tbody");
      t.appendChild(tbody);
      for (const tr of Array.from(el.querySelectorAll("tr"))) {
        if (tr.closest("table") !== el) continue;
        const row = src.createElement("tr");
        for (const cell of Array.from(tr.children)) {
          const ct = cell.tagName.toLowerCase();
          if (ct !== "td" && ct !== "th") continue;
          const td = src.createElement("td");
          if (cell.getAttribute("colspan")) td.setAttribute("colspan", cell.getAttribute("colspan")!);
          if (cell.getAttribute("rowspan")) td.setAttribute("rowspan", cell.getAttribute("rowspan")!);
          const bg = (cell as HTMLElement).style?.backgroundColor; if (bg) td.style.backgroundColor = bg;
          // Cell content: render into a temporary container
          const saveOut = out.childNodes.length;
          const savedPara = curPara; curPara = null;
          const cellInl = ct === "th" ? { ...inl, b: true } : inl;
          for (const c of Array.from(cell.childNodes)) walk(c, inlineFrom(cell as HTMLElement, cellInl), link, null, pre, {});
          curPara = null;
          const produced = Array.from(out.childNodes).slice(saveOut);
          for (const p of produced) td.appendChild(p);
          if (!td.childNodes.length) td.appendChild(src.createElement("p"));
          row.appendChild(td);
          curPara = savedPara;
        }
        if (row.childNodes.length) tbody.appendChild(row);
      }
      if (tbody.childNodes.length) out.appendChild(t);
      endPara();
      return;
    }
    if (tag === "ul" || tag === "ol") {
      endPara();
      const kind = tag === "ul" ? "bullet" : "decimal";
      const level = list ? list.level + 1 : 0;
      const numId = list ? list.numId : listId(kind, kind + ":" + (out.childNodes.length));
      const ctxList: ListCtx = { numId, level: Math.min(8, level) };
      for (const li of Array.from(el.children)) {
        if (li.tagName.toLowerCase() !== "li") { walk(li, inl, link, ctxList, pre, paraAttrs); continue; }
        endPara();
        const attrs = { "data-num": `${numId}/${ctxList.level}` };
        ensurePara(attrs);
        for (const c of Array.from(li.childNodes)) {
          const ctag = c.nodeType === 1 ? (c as HTMLElement).tagName.toLowerCase() : "";
          if (ctag === "ul" || ctag === "ol") { walk(c, inl, link, ctxList, pre, attrs); endPara(); continue; }
          if (ctag === "p" || ctag === "div") { for (const cc of Array.from(c.childNodes)) walk(cc, inlineFrom(c as HTMLElement, inl), link, ctxList, pre, attrs); ensurePara(attrs); continue; }
          walk(c, inl, link, ctxList, pre, attrs);
        }
        endPara();
      }
      return;
    }
    const isBlock = BLOCK_TAGS.has(tag);
    let attrs = paraAttrs;
    if (isBlock) {
      endPara();
      attrs = {};
      // Plain containers inherit style/list attrs from the enclosing block (blockquote > p, li > p).
      if (tag === "p" || tag === "div" || tag === "section" || tag === "article" || tag === "center") {
        if (paraAttrs["data-style"]) attrs["data-style"] = paraAttrs["data-style"];
        if (paraAttrs["data-num"]) attrs["data-num"] = paraAttrs["data-num"];
      }
      if (/^h[1-6]$/.test(tag)) attrs["data-style"] = "Heading" + tag[1];
      else if (tag === "blockquote") attrs["data-style"] = "Quote";
      else if (tag === "pre") { attrs["data-style"] = "Code"; pre = true; inl = { ...inl, font: "Consolas" }; }
      const al = el.style?.textAlign || el.getAttribute("align") || paraAttrs["data-align"];
      if (al) attrs["data-align"] = al;
      pendingAttrs = attrs;
      // Leaf blocks always produce a paragraph (even when empty); containers create one lazily.
      if (/^h[1-6]$/.test(tag) || tag === "pre" || tag === "p") ensurePara(attrs);
    }
    const childInl = inlineFrom(el, inl);
    for (const c of Array.from(el.childNodes)) walk(c, childInl, link, list, pre, attrs);
    if (isBlock) { endPara(); pendingAttrs = paraAttrs; }
  };

  for (const c of Array.from(src.body.childNodes)) walk(c, {}, null, null, false, {});
  endPara();
  // Apply collected paragraph attrs as inline styles/attributes the schema parse rules understand.
  for (const p of Array.from(out.querySelectorAll("p"))) {
    const style = p.getAttribute("data-style");
    if (style) p.setAttribute("data-pstyle", style);
    const al = p.getAttribute("data-align");
    if (al) p.style.textAlign = al;
  }
  return out.innerHTML;
}

// ---------------------------------------------------------------------------
// Markdown -> document

export function markdownToDoc(md: string): PMNode {
  const html = marked.parse(md, { gfm: true, breaks: false }) as string;
  const norm = normalizeHtml(html);
  const dom = new DOMParser().parseFromString("<div>" + norm + "</div>", "text/html");
  const parser = PMDOMParser.fromSchema(schema);
  const doc = parser.parse(dom.body.firstElementChild || dom.body);
  return doc.childCount ? doc : schema.nodes.doc.create(doc.attrs, schema.nodes.paragraph.create());
}

// ---------------------------------------------------------------------------
// Document -> Markdown

export interface MarkdownResult { markdown: string; assets: { name: string; bytes: Uint8Array }[]; }

export function docToMarkdown(doc: PMNode, resolveImage: (node: PMNode) => { name: string; bytes: Uint8Array } | null): MarkdownResult {
  const assets: { name: string; bytes: Uint8Array }[] = [];
  const lines: string[] = [];
  let prevWasList = false;

  const inlineMd = (para: PMNode): string => {
    let s = "";
    para.forEach((n) => {
      const t = n.type.name;
      if (t === "text") {
        const m = n.marks.find((mk) => mk.type === schema.marks.rpr);
        const p = m ? effectiveRunProps(m.attrs.props as RunProps) : {};
        let txt = n.text || "";
        const link = n.marks.find((mk) => mk.type === schema.marks.link);
        const esc = txt.replace(/([*_`\\])/g, "\\$1");
        let w = esc;
        const code = /consolas|courier|mono/i.test(p.font || "");
        if (code) w = "`" + txt.replace(/`/g, "``") + "`";
        else {
          if (p.b && p.i) w = "***" + w + "***"; else if (p.b) w = "**" + w + "**"; else if (p.i) w = "*" + w + "*";
          if (p.strike) w = "~~" + w + "~~";
        }
        if (link && link.attrs.href) w = `[${w}](${link.attrs.href})`;
        s += w;
      } else if (t === "tab") s += "\t";
      else if (t === "hard_break") s += n.attrs.kind === "page" ? "\n\n<div style=\"page-break-after: always\"></div>\n\n" : "  \n";
      else if (t === "image") {
        const a = resolveImage(n);
        if (a) { assets.push(a); s += `![${n.attrs.alt || n.attrs.name || ""}](${a.name})`; }
        else if (/^https?:/.test(n.attrs.src)) s += `![${n.attrs.alt || ""}](${n.attrs.src})`;
      } else if (t === "opaque_inline") s += n.attrs.text || "";
      else if (t === "textbox") { n.forEach((b) => { if (b.type === schema.nodes.paragraph) s += inlineMd(b) + " "; }); }
    });
    return s;
  };

  const writePara = (p: PMNode, prefixIndent = "") => {
    const eff = paragraphStyle(p).pPr;
    const text = inlineMd(p);
    const style = (eff.pStyle || "").toLowerCase();
    const hm = /heading\s?(\d)/.exec(style) || /^(?:başlık|titre|überschrift)\s?(\d)/.exec(style);
    if (eff.numId) {
      const lvl = ctx.numLevel(eff.numId, eff.ilvl || 0);
      const bullet = !lvl || lvl.numFmt === "bullet";
      const indent = "  ".repeat(eff.ilvl || 0);
      lines.push(prefixIndent + indent + (bullet ? "- " : "1. ") + text);
      prevWasList = true;
      return;
    }
    if (prevWasList) { lines.push(""); prevWasList = false; }
    if (hm) { lines.push(prefixIndent + "#".repeat(Math.min(6, parseInt(hm[1], 10))) + " " + text); lines.push(""); return; }
    if (style === "title") { lines.push(prefixIndent + "# " + text); lines.push(""); return; }
    if (style === "quote" || style === "intensequote") { lines.push(prefixIndent + "> " + text); lines.push(""); return; }
    if (style === "code") { lines.push(prefixIndent + "    " + p.textContent); return; }
    if (eff.bdrBottom && !text.trim()) { lines.push(prefixIndent + "---"); lines.push(""); return; }
    lines.push(prefixIndent + text);
    lines.push("");
  };

  doc.forEach((node) => {
    if (node.type === schema.nodes.paragraph) writePara(node);
    else if (node.type === schema.nodes.table) {
      if (prevWasList) { lines.push(""); prevWasList = false; }
      const rows: string[][] = [];
      node.forEach((row) => {
        const cells: string[] = [];
        row.forEach((cell) => {
          const parts: string[] = [];
          cell.forEach((b) => { if (b.type === schema.nodes.paragraph) parts.push(inlineMd(b)); });
          let txt = parts.join("<br>").replace(/\|/g, "\\|").trim();
          cells.push(txt);
          for (let i = 1; i < (cell.attrs.colspan as number); i++) cells.push("");
        });
        rows.push(cells);
      });
      if (rows.length) {
        const width = Math.max(...rows.map((r) => r.length));
        const pad = (r: string[]) => { while (r.length < width) r.push(""); return r; };
        lines.push("| " + pad(rows[0]).join(" | ") + " |");
        lines.push("|" + new Array(width).fill(" --- ").join("|") + "|");
        for (let i = 1; i < rows.length; i++) lines.push("| " + pad(rows[i]).join(" | ") + " |");
        lines.push("");
      }
    } else if (node.type === schema.nodes.opaque_block) { /* skip */ }
  });
  // collapse triple blank lines
  const md = lines.join("\n").replace(/\n{3,}/g, "\n\n").trimEnd() + "\n";
  return { markdown: md, assets };
}

export type { ParaProps };
