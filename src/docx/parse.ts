// document.xml -> ProseMirror document. Everything we do not model is kept as
// raw XML on the nodes so it can be written back untouched.
import { Node as PMNode } from "prosemirror-model";
import { schema, SectProps, DEFAULT_SECT, clearCssCaches } from "../schema";
import { NS, parseXml, child, children, wattr, wint, attr, serialize, rootNamespaceDecls, runText, descendant, NO_BREAK_HYPHEN, SOFT_HYPHEN } from "./xml";
import { Package, parseRels, resolveTarget, Relationship, mimeForExt, REL_TYPES } from "./zip";
import { DocContext, setContext } from "./styles";
import { parsePPr, parseRPr, parseTblPr, parseTrPr, parseTcPr, ParaProps, RunProps, CellProps } from "./props";
import { emuToPx, cssLenToPx } from "./units";
import { placeholderImage, emfToDataUrl } from "./images";

export interface MediaEntry { part: string; url: string; ext: string; bytes: Uint8Array | null; deferred: boolean; }

/** Large media stays compressed until after first paint. */
export function shouldDeferMedia(name: string, size: number): boolean {
  return name.includes("/media/") && (size > 400_000 || /\.(emf|wmf)$/i.test(name));
}

/** Turn raw media bytes into something the webview can display. */
export function mediaUrl(bytes: Uint8Array, ext: string): string {
  if (ext === "emf" || ext === "wmf") return emfToDataUrl(bytes, ext) || placeholderImage(200, 120, ext.toUpperCase() + " image");
  return URL.createObjectURL(new Blob([bytes as BlobPart], { type: mimeForExt(ext) }));
}

/**
 * Inflate deferred media one entry per call; `onEach` receives the rId and its
 * final URL so image nodes can be updated. Resolves when everything is loaded.
 */
export function resolveDeferredMedia(loaded: LoadedDoc, onEach: (rId: string, entry: MediaEntry) => void): Promise<void> {
  const pending = Array.from(loaded.media.entries()).filter(([, m]) => m.deferred);
  return new Promise((resolve) => {
    const step = () => {
      const next = pending.shift();
      if (!next) { resolve(); return; }
      const [rId, m] = next;
      try {
        const bytes = loaded.pkg.get(m.part);
        if (bytes) { m.bytes = bytes; m.url = mediaUrl(bytes, m.ext); }
      } catch (e) { loaded.warnings.push("Failed to load " + m.part + ": " + (e as Error).message); }
      m.deferred = false;
      onEach(rId, m);
      setTimeout(step, 0);
    };
    setTimeout(step, 0);
  });
}

export interface HeaderFooter { rId: string; part: string; doc: PMNode; xml: string; }

export interface LoadedDoc {
  pkg: Package;
  ctx: DocContext;
  doc: PMNode;
  docPart: string;
  docXml: string;
  rels: Map<string, Relationship>;
  media: Map<string, MediaEntry>;
  headers: Map<string, HeaderFooter>;
  footers: Map<string, HeaderFooter>;
  rootDecls: Map<string, string>;
  warnings: string[];
}

const SYMBOL_MAP: Record<number, string> = {
  0xf0b7: "•", 0xf0a7: "▪", 0xf0a8: "□", 0xf0fc: "✓", 0xf0fe: "☒", 0xf0fd: "☑", 0xf0d8: "➢", 0xf076: "❖",
  0xf06c: "●", 0xf06e: "■", 0xf075: "◆", 0xf0a1: "○", 0xf09f: "•", 0xf0a0: " ", 0xf0d7: "×", 0xf0ae: "→",
  0xf0ac: "←", 0xf0a3: "≤", 0xf0b3: "≥", 0xf0b1: "±", 0xf0b0: "°", 0xf0bc: "…", 0xf0b9: "≠", 0xf0bb: "≈",
  0xf0d6: "√", 0xf0e5: "∑", 0xf0a5: "∞", 0xf0be: "−", 0xf02d: "−", 0xf0a2: "′", 0xf0a4: "⁄", 0xf0ab: "↔",
  0xf0e0: "◊", 0xf0e8: "⇒", 0xf0f0: "ð", 0xf0d5: "∏", 0xf0f8: "ø", 0xf0b2: "″", 0xf03e: "∼",
};

export function mapSymbolChar(font: string | null | undefined, ch: string): string {
  const code = ch.codePointAt(0) || 0;
  if (code >= 0xf000 && code <= 0xf0ff) {
    if (SYMBOL_MAP[code]) return SYMBOL_MAP[code];
    if (/symbol|wingdings|webdings/i.test(font || "")) return "•";
    return String.fromCodePoint(code - 0xf000);
  }
  return ch;
}

type Json = any;

interface FieldState { depth: number; buf: Element[]; instr: string; result: string; inResult: boolean; }

class BodyParser {
  private sdtCounter = 0;
  private rprCache = new Map<string, Json>();
  warnings: string[] = [];

  constructor(
    private ctx: DocContext,
    private rels: Map<string, Relationship>,
    private media: Map<string, MediaEntry>,
    private rootDecls: Map<string, string>,
  ) {}

  ser(el: Element): string { return serialize(el, this.rootDecls); }

  // ---- blocks -------------------------------------------------------------

  parseBlocks(container: Element, inTable: boolean, tblStyle: string | null, sdt: Json | null): Json[] {
    const out: Json[] = [];
    for (const el of children(container)) this.parseBlock(el, out, inTable, tblStyle, sdt);
    return out;
  }

  parseBlock(el: Element, out: Json[], inTable: boolean, tblStyle: string | null, sdt: Json | null) {
    if (el.namespaceURI === NS.w) {
      switch (el.localName) {
        case "p": out.push(this.parseParagraph(el, inTable, tblStyle, sdt)); return;
        case "tbl": out.push(this.parseTable(el, sdt)); return;
        case "sectPr": return; // handled by caller
        case "tcPr": case "trPr": case "tblPr": case "tblGrid": case "tblPrEx": case "pPr": return; // properties, handled by their owners
        case "sdt": {
          const content = child(el, NS.w, "sdtContent");
          const pr = child(el, NS.w, "sdtPr");
          const endPr = child(el, NS.w, "sdtEndPr");
          const info = { id: ++this.sdtCounter, pr: pr ? this.ser(pr) : "", endPr: endPr ? this.ser(endPr) : "" };
          if (content) {
            const before = out.length;
            for (const c of children(content)) this.parseBlock(c, out, inTable, tblStyle, info);
            if (out.length === before) out.push(this.emptyParagraph(inTable, tblStyle, info));
          }
          return;
        }
        case "customXml": {
          const before = out.length;
          for (const c of children(el)) this.parseBlock(c, out, inTable, tblStyle, sdt);
          if (out.length === before) out.push({ type: "opaque_block", attrs: { xml: this.ser(el), label: "Custom XML (preserved)" } });
          return;
        }
        case "bookmarkStart": case "bookmarkEnd": case "commentRangeStart": case "commentRangeEnd":
        case "permStart": case "permEnd": case "proofErr":
          // Block-level markers: keep as opaque, rendered invisibly.
          if (el.localName !== "proofErr") out.push({ type: "opaque_block", attrs: { xml: this.ser(el), label: "" } });
          return;
        case "altChunk":
          out.push({ type: "opaque_block", attrs: { xml: this.ser(el), label: "Embedded content (altChunk) preserved" } });
          return;
      }
    }
    if (el.namespaceURI === NS.mc && el.localName === "AlternateContent") {
      const choice = child(el, NS.mc, "Choice") || child(el, NS.mc, "Fallback");
      if (choice) for (const c of children(choice)) this.parseBlock(c, out, inTable, tblStyle, sdt);
      return;
    }
    out.push({ type: "opaque_block", attrs: { xml: this.ser(el), label: `Unsupported element <${el.nodeName}> (preserved)` } });
  }

  emptyParagraph(inTable: boolean, tblStyle: string | null, sdt: Json | null): Json {
    return { type: "paragraph", attrs: { pPr: null, props: {}, inTable, tblStyle, sdt, sectPr: null } };
  }

  parseParagraph(p: Element, inTable: boolean, tblStyle: string | null, sdt: Json | null): Json {
    const pPrEl = child(p, NS.w, "pPr");
    const props: ParaProps = pPrEl ? parsePPr(pPrEl, this.ctx.theme) : {};
    let sectPr: string | null = null;
    if (pPrEl) {
      const sp = child(pPrEl, NS.w, "sectPr");
      if (sp) sectPr = this.ser(sp);
    }
    const content: Json[] = [];
    const field: FieldState = { depth: 0, buf: [], instr: "", result: "", inResult: false };
    for (const el of children(p)) {
      if (el === pPrEl) continue;
      this.parseInline(el, content, [], field);
    }
    if (field.depth > 0) this.flushField(field, content);
    const attrs = { pPr: pPrEl ? this.ser(pPrEl) : null, props, inTable, tblStyle, sdt, sectPr };
    const node: Json = { type: "paragraph", attrs };
    if (content.length) node.content = content;
    return node;
  }

  // ---- inline -------------------------------------------------------------

  private marksFor(rPrEl: Element | null, extra: Json[]): Json[] {
    if (!rPrEl) return extra.length ? extra : [];
    const xml = this.ser(rPrEl);
    let m = this.rprCache.get(xml);
    if (!m) {
      m = { type: "rpr", attrs: { xml, props: parseRPr(rPrEl, this.ctx.theme) } };
      this.rprCache.set(xml, m);
    }
    return extra.length ? [...extra, m] : [m];
  }

  parseInline(el: Element, out: Json[], linkMarks: Json[], field: FieldState) {
    // Inside a complex field everything is buffered verbatim.
    if (field.depth > 0 && !this.containsFieldChar(el)) {
      field.buf.push(el);
      if (field.inResult) field.result += runText(el);
      else field.instr += this.instrText(el);
      return;
    }
    if (el.namespaceURI === NS.w) {
      switch (el.localName) {
        case "r": this.parseRun(el, out, linkMarks, field); return;
        case "hyperlink": {
          const rid = attr(el, NS.r, "id");
          const rel = rid ? this.rels.get(rid) : undefined;
          const anchor = wattr(el, "anchor");
          const attrsStr = Array.from(el.attributes).map((a) => `${a.name}="${a.value.replace(/"/g, "&quot;")}"`).join(" ");
          const mark = { type: "link", attrs: { href: rel ? rel.target : anchor ? "#" + anchor : "", rId: rid, anchor, tooltip: wattr(el, "tooltip"), raw: attrsStr } };
          for (const c of children(el)) this.parseInline(c, out, [mark], field);
          return;
        }
        case "fldSimple": {
          const instr = wattr(el, "instr") || "";
          out.push({ type: "opaque_inline", attrs: { xml: this.ser(el), text: runText(el), kind: "field", field: instr.trim() } });
          return;
        }
        case "sdt": {
          const content = child(el, NS.w, "sdtContent");
          out.push({ type: "opaque_inline", attrs: { xml: this.ser(el), text: content ? runText(content) : "", kind: "sdt" } });
          return;
        }
        case "smartTag": case "customXml": case "ins": case "moveTo": case "dir": case "bdo":
          for (const c of children(el)) this.parseInline(c, out, linkMarks, field);
          return;
        case "del": case "moveFrom":
          out.push({ type: "opaque_inline", attrs: { xml: this.ser(el), text: "", kind: "marker" } });
          return;
        case "proofErr": return;
        case "bookmarkStart": case "bookmarkEnd": case "commentRangeStart": case "commentRangeEnd":
        case "permStart": case "permEnd": case "moveFromRangeStart": case "moveFromRangeEnd":
        case "moveToRangeStart": case "moveToRangeEnd": case "customXmlInsRangeStart": case "customXmlInsRangeEnd":
          out.push({ type: "opaque_inline", attrs: { xml: this.ser(el), text: "", kind: "marker" } });
          return;
        case "pPr": return;
      }
    }
    if (el.namespaceURI === NS.mc && el.localName === "AlternateContent") {
      const choice = child(el, NS.mc, "Choice") || child(el, NS.mc, "Fallback");
      if (choice) for (const c of children(choice)) this.parseInline(c, out, linkMarks, field);
      return;
    }
    // Math (m:oMath, m:oMathPara) and anything else: opaque, show its text.
    out.push({ type: "opaque_inline", attrs: { xml: this.ser(el), text: runText(el), kind: "object" } });
  }

  private containsFieldChar(el: Element): boolean {
    if (el.namespaceURI !== NS.w || el.localName !== "r") return false;
    return !!child(el, NS.w, "fldChar");
  }

  private instrText(el: Element): string {
    let s = "";
    for (const t of el.getElementsByTagNameNS(NS.w, "instrText")) s += t.textContent || "";
    return s;
  }

  private flushField(field: FieldState, out: Json[]) {
    const xml = field.buf.map((e) => this.ser(e)).join("");
    out.push({ type: "opaque_inline", attrs: { xml, text: field.result, kind: "field", field: field.instr.trim() } });
    field.depth = 0; field.buf = []; field.instr = ""; field.result = ""; field.inResult = false;
  }

  parseRun(r: Element, out: Json[], linkMarks: Json[], field: FieldState) {
    const fld = child(r, NS.w, "fldChar");
    if (fld) {
      const type = wattr(fld, "fldCharType");
      if (type === "begin") {
        if (field.depth === 0) { field.buf = []; field.instr = ""; field.result = ""; field.inResult = false; }
        field.depth++;
        field.buf.push(r);
        return;
      }
      if (field.depth > 0) {
        field.buf.push(r);
        if (type === "separate" && field.depth === 1) field.inResult = true;
        if (type === "end") { field.depth--; if (field.depth === 0) this.flushField(field, out); }
        return;
      }
      // stray fldChar: keep as marker
      out.push({ type: "opaque_inline", attrs: { xml: this.ser(r), text: "", kind: "marker" } });
      return;
    }
    const rPrEl = child(r, NS.w, "rPr");
    const marks = this.marksFor(rPrEl, linkMarks);
    const withMarks = (n: Json) => { if (marks.length) n.marks = marks; return n; };
    for (const c of children(r)) {
      if (c === rPrEl) continue;
      if (c.namespaceURI === NS.w) {
        switch (c.localName) {
          case "t": { const t = c.textContent || ""; if (t) out.push(withMarks({ type: "text", text: t })); break; }
          case "tab": case "ptab": out.push(withMarks({ type: "tab" })); break;
          case "br": out.push({ type: "hard_break", attrs: { kind: wattr(c, "type") || "textWrapping", clear: wattr(c, "clear") } }); break;
          case "cr": out.push({ type: "hard_break", attrs: { kind: "textWrapping", clear: null } }); break;
          case "noBreakHyphen": out.push(withMarks({ type: "text", text: NO_BREAK_HYPHEN })); break;
          case "softHyphen": out.push(withMarks({ type: "text", text: SOFT_HYPHEN })); break;
          case "sym": {
            const font = wattr(c, "font");
            const code = parseInt(wattr(c, "char") || "0", 16);
            const ch = code ? mapSymbolChar(font, String.fromCodePoint(code)) : "";
            if (ch) out.push(withMarks({ type: "text", text: ch }));
            break;
          }
          case "drawing": { const n = this.parseDrawing(c, marks); if (n) out.push(n); break; }
          case "pict": { const n = this.parsePict(c, marks); if (n) out.push(n); break; }
          case "object": { const n = this.parseObject(c, r); if (n) out.push(n); break; }
          case "footnoteReference": case "endnoteReference": {
            const id = wattr(c, "id") || "";
            out.push({ type: "opaque_inline", attrs: { xml: this.ser(r), text: id, kind: "field", field: c.localName === "footnoteReference" ? "FOOTNOTE" : "ENDNOTE" } });
            return; // the whole run is the reference
          }
          case "commentReference": case "annotationRef": case "footnoteRef": case "endnoteRef":
            out.push({ type: "opaque_inline", attrs: { xml: this.ser(r), text: "", kind: "marker" } });
            return;
          case "lastRenderedPageBreak": case "instrText": case "delText": case "delInstrText": break;
          case "ruby": out.push(withMarks({ type: "text", text: runText(c) })); break;
          case "dayShort": case "dayLong": case "monthShort": case "monthLong": case "yearShort": case "yearLong":
            out.push({ type: "opaque_inline", attrs: { xml: this.ser(r), text: "", kind: "field", field: c.localName.toUpperCase() } });
            return;
          default:
            // Unknown run child: keep the entire run opaque.
            out.push({ type: "opaque_inline", attrs: { xml: this.ser(r), text: runText(c), kind: "object" } });
            return;
        }
      } else if (c.namespaceURI === NS.mc && c.localName === "AlternateContent") {
        const n = this.parseAlternate(c, marks);
        if (n) out.push(n);
      } else {
        out.push({ type: "opaque_inline", attrs: { xml: this.ser(r), text: runText(c), kind: "object" } });
        return;
      }
    }
  }

  private parseAlternate(ac: Element, marks: Json[]): Json | null {
    const choice = child(ac, NS.mc, "Choice");
    const fallback = child(ac, NS.mc, "Fallback");
    const raw = this.ser(ac);
    if (choice) {
      for (const c of children(choice)) {
        if (c.namespaceURI === NS.w && c.localName === "drawing") { const n = this.parseDrawing(c, marks, raw); if (n) return n; }
        if (c.namespaceURI === NS.w && c.localName === "pict") { const n = this.parsePict(c, marks, raw); if (n) return n; }
      }
    }
    if (fallback) {
      for (const c of children(fallback)) {
        if (c.namespaceURI === NS.w && c.localName === "pict") { const n = this.parsePict(c, marks, raw); if (n) return n; }
        if (c.namespaceURI === NS.w && c.localName === "drawing") { const n = this.parseDrawing(c, marks, raw); if (n) return n; }
      }
    }
    return { type: "opaque_inline", attrs: { xml: raw, text: runText(ac), kind: "object" } };
  }

  // ---- drawings -----------------------------------------------------------

  private mediaFor(rId: string | null): MediaEntry | undefined {
    return rId ? this.media.get(rId) : undefined;
  }

  parseDrawing(d: Element, marks: Json[], rawOverride?: string): Json | null {
    const raw = rawOverride ?? this.ser(d);
    const inline = child(d, NS.wp, "inline");
    const anchor = child(d, NS.wp, "anchor");
    const box = inline || anchor;
    if (!box) return { type: "opaque_inline", attrs: { xml: raw, text: "", kind: "shape", w: 40, h: 20 } };
    const extent = child(box, NS.wp, "extent");
    let w = emuToPx(parseInt(attr(extent, null, "cx") || "0", 10));
    let h = emuToPx(parseInt(attr(extent, null, "cy") || "0", 10));
    const docPr = child(box, NS.wp, "docPr");
    const name = attr(docPr, null, "name") || "";
    const alt = attr(docPr, null, "descr") || "";
    const wrap = anchor ? this.parseAnchorWrap(anchor, w) : null;
    const kind = anchor ? "anchor" : "inline";
    const graphic = child(box, NS.a, "graphic");
    const gd = child(graphic, NS.a, "graphicData");
    const uri = attr(gd, null, "uri") || "";
    if (uri.endsWith("/picture")) {
      const pic = child(gd, NS.pic, "pic");
      const blipFill = child(pic, NS.pic, "blipFill");
      const blip = child(blipFill, NS.a, "blip");
      let rId = attr(blip, NS.r, "embed") || attr(blip, NS.r, "link");
      // Prefer SVG when present (Word stores PNG fallback in blip, SVG in an extension).
      const svg = blip ? descendant(blip, NS.asvg, "svgBlip") : null;
      const svgId = svg ? attr(svg, NS.r, "embed") : null;
      let m = this.mediaFor(svgId) || this.mediaFor(rId);
      if (svgId && this.mediaFor(svgId)) rId = rId; // keep original rId for round-trip; url from svg
      // Cropping is ignored for now (srcRect).
      const src = m ? (m.deferred ? placeholderImage(w || 100, h || 60, "Loading…") : m.url) : placeholderImage(w || 100, h || 60, "Missing image");
      return { type: "image", attrs: { src, w: w || 100, h: h || 60, rId, raw, kind, name, alt, wrap, media: null, origW: w, origH: h, ext: m ? m.ext : null } };
    }
    if (uri.endsWith("/wordprocessingShape")) {
      const wsp = child(gd, NS.wps, "wsp");
      const txbx = child(wsp, NS.wps, "txbx");
      const content = txbx ? child(txbx, NS.w, "txbxContent") : null;
      const style = this.shapeStyle(wsp);
      if (content) {
        const blocks = this.parseBlocks(content, false, null, null);
        if (!blocks.length) blocks.push(this.emptyParagraph(false, null, null));
        return { type: "textbox", attrs: { raw, w: w || null, h: h || null, kind, style, wrap }, content: blocks };
      }
      return { type: "opaque_inline", attrs: { xml: raw, text: "", kind: "shape", w: w || 40, h: h || 20 } };
    }
    if (uri.endsWith("/wordprocessingGroup")) {
      // Group: render as a placeholder box but expose text of contained text boxes.
      const text = gd ? Array.from(gd.getElementsByTagNameNS(NS.w, "txbxContent")).map((t) => runText(t)).join(" ") : "";
      return { type: "opaque_inline", attrs: { xml: raw, text: text ? "Group: " + text.slice(0, 80) : "Group shape", kind: "shape", w: w || 40, h: h || 20 } };
    }
    return { type: "opaque_inline", attrs: { xml: raw, text: name || "Drawing", kind: "shape", w: w || 40, h: h || 20 } };
  }

  private shapeStyle(wsp: Element | null): Json {
    const s: Json = {};
    if (!wsp) return s;
    const spPr = child(wsp, NS.wps, "spPr");
    const colorOf = (fillEl: Element | null): string | null => {
      if (!fillEl) return null;
      const srgb = child(fillEl, NS.a, "srgbClr");
      if (srgb) return attr(srgb, null, "val");
      const sc = child(fillEl, NS.a, "schemeClr");
      if (sc) {
        const lm = child(sc, NS.a, "lumMod"), lo = child(sc, NS.a, "lumOff");
        return this.ctx.schemeColor(attr(sc, null, "val") || "accent1", lm ? parseInt(attr(lm, null, "val") || "100000", 10) : undefined, lo ? parseInt(attr(lo, null, "val") || "0", 10) : undefined);
      }
      const sys = child(fillEl, NS.a, "sysClr");
      if (sys) return attr(sys, null, "lastClr");
      return null;
    };
    const fillColor = colorOf(child(spPr, NS.a, "solidFill"));
    if (fillColor) s.fill = fillColor;
    const ln = child(spPr, NS.a, "ln");
    if (ln && !child(ln, NS.a, "noFill")) {
      const lc = colorOf(child(ln, NS.a, "solidFill"));
      if (lc) { s.line = lc; s.lineW = emuToPx(parseInt(attr(ln, null, "w") || "9525", 10)); }
    }
    const style = child(wsp, NS.wps, "style");
    if (!s.line && style && !ln) {
      const lnRef = child(style, NS.a, "lnRef");
      if (lnRef && attr(lnRef, null, "idx") !== "0") { s.line = colorOf(lnRef) || "000000"; s.lineW = 1; }
    }
    if (!s.fill && style && !child(spPr, NS.a, "noFill")) {
      const fillRef = child(style, NS.a, "fillRef");
      if (fillRef && attr(fillRef, null, "idx") !== "0") s.fill = colorOf(fillRef);
    }
    const bodyPr = child(wsp, NS.wps, "bodyPr");
    if (bodyPr) {
      const ins = (n: string, d: number) => { const v = attr(bodyPr, null, n); return v ? emuToPx(parseInt(v, 10)) * 0.75 : d; };
      s.padL = ins("lIns", 7.2); s.padR = ins("rIns", 7.2); s.padT = ins("tIns", 3.6); s.padB = ins("bIns", 3.6);
    }
    return s;
  }

  contentWidthPx = 600;

  private parseAnchorWrap(anchor: Element, w: number): Json {
    const wrap: Json = { type: "none", float: null, behind: attr(anchor, null, "behindDoc") === "1" };
    const posH = child(anchor, NS.wp, "positionH");
    const posV = child(anchor, NS.wp, "positionV");
    const relH = attr(posH, null, "relativeFrom") || "column";
    const relV = attr(posV, null, "relativeFrom") || "paragraph";
    const alignH = child(posH, NS.wp, "align")?.textContent || null;
    const alignV = child(posV, NS.wp, "align")?.textContent || null;
    const offH = child(posH, NS.wp, "posOffset")?.textContent;
    const offV = child(posV, NS.wp, "posOffset")?.textContent;
    const dist = (n: string) => { const v = attr(anchor, null, n); return v ? emuToPx(parseInt(v, 10)) : 0; };
    wrap.distL = dist("distL"); wrap.distR = dist("distR"); wrap.distT = dist("distT"); wrap.distB = dist("distB");
    for (const c of children(anchor, NS.wp)) {
      if (c.localName === "wrapSquare") wrap.type = "square";
      else if (c.localName === "wrapTight") wrap.type = "tight";
      else if (c.localName === "wrapThrough") wrap.type = "through";
      else if (c.localName === "wrapTopAndBottom") wrap.type = "topAndBottom";
      else if (c.localName === "wrapNone") wrap.type = "none";
    }
    wrap.relH = relH; wrap.relV = relV;
    wrap.alignH = alignH; wrap.alignV = alignV;
    wrap.offH = offH ? emuToPx(parseInt(offH, 10)) : null;
    wrap.offV = offV ? emuToPx(parseInt(offV, 10)) : null;
    // Float side for the flow fallback.
    let side: string | null = null;
    if (alignH === "right" || alignH === "outside") side = "right";
    else if (alignH === "left" || alignH === "inside") side = "left";
    else if (alignH === "center") side = "center";
    else if (wrap.offH !== null) side = wrap.offH + w / 2 > this.contentWidthPx / 2 ? "right" : "left";
    if (wrap.type === "topAndBottom") wrap.float = side === "right" ? "right" : side === "center" ? "center" : "left";
    else if (wrap.type === "none") wrap.float = null;
    else wrap.float = side === "right" ? "right" : "left";
    return wrap;
  }

  parsePict(p: Element, marks: Json[], rawOverride?: string): Json | null {
    const raw = rawOverride ?? this.ser(p);
    // Find the first shape-ish VML element.
    let shape: Element | null = null;
    for (const c of children(p)) if (c.namespaceURI === NS.v && c.localName !== "shapetype") { shape = c; break; }
    if (!shape) return { type: "opaque_inline", attrs: { xml: raw, text: "", kind: "shape", w: 40, h: 20 } };
    const style = attr(shape, null, "style") || "";
    const get = (k: string) => { const m = new RegExp("(?:^|;)\\s*" + k + ":([^;]+)").exec(style); return m ? m[1].trim() : null; };
    const w = cssLenToPx(get("width")) || 0;
    const h = cssLenToPx(get("height")) || 0;
    const imagedata = descendant(shape, NS.v, "imagedata");
    if (imagedata) {
      const rId = attr(imagedata, NS.r, "id") || attr(imagedata, NS.o, "relid");
      const m = this.mediaFor(rId);
      const src = m ? m.url : placeholderImage(w || 100, h || 60, "Missing image");
      return { type: "image", attrs: { src, w: w || 100, h: h || 60, rId, raw, kind: get("position") === "absolute" ? "anchor" : "inline", name: attr(shape, null, "alt") || "", alt: "", wrap: null, media: null, origW: w, origH: h, ext: m ? m.ext : null } };
    }
    const content = descendant(shape, NS.w, "txbxContent");
    if (content) {
      const blocks = this.parseBlocks(content, false, null, null);
      if (!blocks.length) blocks.push(this.emptyParagraph(false, null, null));
      const st: Json = {};
      const fill = attr(shape, null, "fillcolor");
      if (fill && !/^#?fff(fff)?$/i.test(fill)) st.fill = fill.replace("#", "");
      if (attr(shape, null, "stroked") !== "f") { st.line = (attr(shape, null, "strokecolor") || "#000000").replace("#", ""); st.lineW = 1; }
      return { type: "textbox", attrs: { raw, w: w || null, h: h || null, kind: "anchor", style: st, wrap: null }, content: blocks };
    }
    return { type: "opaque_inline", attrs: { xml: raw, text: "", kind: "shape", w: w || 40, h: h || 20 } };
  }

  parseObject(o: Element, run: Element): Json | null {
    const raw = this.ser(run);
    const shape = descendant(o, NS.v, "shape");
    const imagedata = shape ? descendant(shape, NS.v, "imagedata") : null;
    const style = attr(shape, null, "style") || "";
    const get = (k: string) => { const m = new RegExp("(?:^|;)\\s*" + k + ":([^;]+)").exec(style); return m ? m[1].trim() : null; };
    const w = cssLenToPx(get("width")) || 100, h = cssLenToPx(get("height")) || 60;
    const rId = imagedata ? attr(imagedata, NS.r, "id") : null;
    const m = this.mediaFor(rId);
    const src = m ? m.url : placeholderImage(w, h, "Embedded object");
    return { type: "image", attrs: { src, w, h, rId, raw, kind: "object", name: "Embedded object", alt: "", wrap: null, media: null, origW: w, origH: h, ext: m ? m.ext : null } };
  }

  // ---- tables -------------------------------------------------------------

  parseTable(tbl: Element, sdt: Json | null): Json {
    const tblPrEl = child(tbl, NS.w, "tblPr");
    const props = parseTblPr(tblPrEl);
    const gridEl = child(tbl, NS.w, "tblGrid");
    let grid: number[] | null = gridEl ? children(gridEl, NS.w, "gridCol").map((g) => wint(g, "w") ?? 0) : null;
    if (grid && !grid.length) grid = null;
    const tblStyle = props.tblStyle || null;
    const rowEls = children(tbl, NS.w, "tr");
    // Also rows wrapped in sdt/customXml.
    const allRows: Element[] = [];
    for (const c of children(tbl)) {
      if (c.namespaceURI !== NS.w) continue;
      if (c.localName === "tr") allRows.push(c);
      else if (c.localName === "sdt" || c.localName === "customXml") {
        const content = c.localName === "sdt" ? child(c, NS.w, "sdtContent") : c;
        if (content) for (const r of children(content, NS.w, "tr")) allRows.push(r);
      }
    }
    void rowEls;
    const gridCount = grid ? grid.length : Math.max(1, ...allRows.map((r) => this.rowGridWidth(r)));
    if (!grid) grid = new Array(gridCount).fill(Math.round(9000 / gridCount));
    // Build rows with vMerge resolution.
    const rows: Json[] = [];
    // open[col] = the cell json that started a vertical merge covering this column
    const open: (Json | null)[] = new Array(gridCount).fill(null);
    for (const tr of allRows) {
      const trPrEl = child(tr, NS.w, "trPr");
      const rprops = parseTrPr(trPrEl);
      const cells: Json[] = [];
      let col = 0;
      const gridBefore = wint(child(trPrEl, NS.w, "gridBefore"), "val") || 0;
      if (gridBefore > 0) { cells.push(this.padCell(gridBefore, grid, col)); col += gridBefore; }
      const cellEls: Element[] = [];
      for (const c of children(tr)) {
        if (c.namespaceURI !== NS.w) continue;
        if (c.localName === "tc") cellEls.push(c);
        else if (c.localName === "sdt" || c.localName === "customXml") {
          const content = c.localName === "sdt" ? child(c, NS.w, "sdtContent") : c;
          if (content) for (const tc of children(content, NS.w, "tc")) cellEls.push(tc);
        }
      }
      for (const tc of cellEls) {
        const tcPrEl = child(tc, NS.w, "tcPr");
        const cprops: CellProps = parseTcPr(tcPrEl);
        const span = Math.max(1, cprops.gridSpan || 1);
        if (cprops.vMerge === "continue" && open[col] && open[col].attrs.colspan === span) {
          const target = open[col];
          target.attrs.rowspan += 1;
          (target.attrs.vmergeTcPr as string[]).push(tcPrEl ? this.ser(tcPrEl) : "");
          col += span;
          continue;
        }
        const colwidth = grid.slice(col, col + span).map((w) => Math.round(emuToPx(w * 635)));
        const cell: Json = {
          type: "table_cell",
          attrs: { colspan: span, rowspan: 1, colwidth, tcPr: tcPrEl ? this.ser(tcPrEl) : null, props: cprops, vmergeTcPr: [] },
          content: this.parseBlocks(tc, true, tblStyle, null),
        };
        if (!cell.content.length) cell.content.push(this.emptyParagraph(true, tblStyle, null));
        cells.push(cell);
        for (let i = 0; i < span; i++) open[col + i] = cprops.vMerge === "restart" ? cell : null;
        col += span;
      }
      if (col < gridCount) {
        // Word allows short rows; pad so the table map is rectangular.
        let padSpan = gridCount - col;
        // Columns covered by an open rowspan from above must not be padded.
        let coveredTail = 0;
        for (let i = gridCount - 1; i >= col; i--) { if (open[i] && !cells.includes(open[i])) coveredTail++; else break; }
        padSpan -= coveredTail;
        if (padSpan > 0) cells.push(this.padCell(padSpan, grid, col));
      }
      if (!cells.length) cells.push(this.padCell(gridCount, grid, 0));
      rows.push({ type: "table_row", attrs: { trPr: trPrEl ? this.ser(trPrEl) : null, props: rprops }, content: cells });
    }
    if (!rows.length) rows.push({ type: "table_row", attrs: { trPr: null, props: {} }, content: [this.padCell(gridCount, grid, 0)] });
    return { type: "table", attrs: { tblPr: tblPrEl ? this.ser(tblPrEl) : null, props, grid, sdt }, content: rows };
  }

  private rowGridWidth(tr: Element): number {
    let n = 0;
    for (const tc of tr.getElementsByTagNameNS(NS.w, "tc")) {
      if (tc.parentNode !== tr && !(tc.parentNode && (tc.parentNode as Element).parentNode === tr)) continue;
      n += wint(child(child(tc, NS.w, "tcPr"), NS.w, "gridSpan"), "val") || 1;
    }
    return n;
  }

  private padCell(span: number, grid: number[], col: number): Json {
    const colwidth = grid.slice(col, col + span).map((w) => Math.round(emuToPx(w * 635)));
    return {
      type: "table_cell",
      attrs: { colspan: span, rowspan: 1, colwidth, tcPr: null, props: { pad: true }, vmergeTcPr: [] },
      content: [this.emptyParagraph(true, null, null)],
    };
  }
}

// ---------------------------------------------------------------------------

export function parseSectPr(sp: Element | null, ser: (e: Element) => string): SectProps {
  const s: SectProps = { ...DEFAULT_SECT, headers: {}, footers: {} };
  if (!sp) return s;
  const pgSz = child(sp, NS.w, "pgSz");
  if (pgSz) { s.pgW = wint(pgSz, "w") ?? s.pgW; s.pgH = wint(pgSz, "h") ?? s.pgH; s.orient = wattr(pgSz, "orient"); }
  const mar = child(sp, NS.w, "pgMar");
  if (mar) {
    s.marT = wint(mar, "top") ?? s.marT; s.marR = wint(mar, "right") ?? s.marR;
    s.marB = wint(mar, "bottom") ?? s.marB; s.marL = wint(mar, "left") ?? s.marL;
    s.header = wint(mar, "header") ?? s.header; s.footer = wint(mar, "footer") ?? s.footer;
    s.gutter = wint(mar, "gutter") ?? 0;
  }
  // Negative/absurd margins from odd producers: clamp.
  s.marT = Math.max(0, s.marT); s.marB = Math.max(0, s.marB);
  s.titlePg = !!child(sp, NS.w, "titlePg");
  for (const h of children(sp, NS.w, "headerReference")) {
    const t = wattr(h, "type") || "default"; const id = attr(h, NS.r, "id");
    if (id) (s.headers as any)[t] = id;
  }
  for (const f of children(sp, NS.w, "footerReference")) {
    const t = wattr(f, "type") || "default"; const id = attr(f, NS.r, "id");
    if (id) (s.footers as any)[t] = id;
  }
  const cols = child(sp, NS.w, "cols");
  s.cols = wint(cols, "num") || 1;
  s.xml = ser(sp);
  return s;
}

function loadMedia(pkg: Package, sourcePart: string, rels: Map<string, Relationship>, media: Map<string, MediaEntry>, warnings: string[]) {
  for (const rel of rels.values()) {
    if (rel.type !== REL_TYPES.image) continue;
    if (rel.mode === "External") continue;
    const part = resolveTarget(sourcePart, rel.target);
    const ext = (part.split(".").pop() || "").toLowerCase();
    if (pkg.isDeferred(part)) {
      media.set(rel.id, { part, url: placeholderImage(200, 120, "Loading…"), ext, bytes: null, deferred: true });
      continue;
    }
    const bytes = pkg.get(part);
    if (!bytes) { warnings.push("Missing media part " + part); continue; }
    media.set(rel.id, { part, url: mediaUrl(bytes, ext), ext, bytes, deferred: false });
  }
}

export function loadDocx(bytes: Uint8Array): LoadedDoc {
  const pkg = Package.fromBytes(bytes, shouldDeferMedia);
  const warnings: string[] = [];
  // Locate the main document part via package rels.
  const pkgRels = parseRels(pkg.text("_rels/.rels"));
  let docPart = "word/document.xml";
  for (const r of pkgRels.values()) if (r.type === REL_TYPES.officeDocument) { docPart = resolveTarget("", r.target); break; }
  if (!pkg.has(docPart)) throw new Error("Not a Word document (missing " + docPart + ")");
  const docDir = docPart.split("/").slice(0, -1).join("/");
  const relsPart = docDir + "/_rels/" + docPart.split("/").pop() + ".rels";
  const rels = parseRels(pkg.text(relsPart));
  const partFor = (type: string, fallback: string) => {
    for (const r of rels.values()) if (r.type === type) return resolveTarget(docPart, r.target);
    return fallback;
  };
  const ctx = DocContext.parse(
    pkg.text(partFor(REL_TYPES.styles, docDir + "/styles.xml")),
    pkg.text(partFor(REL_TYPES.theme, docDir + "/theme/theme1.xml")),
    pkg.text(partFor(REL_TYPES.numbering, docDir + "/numbering.xml")),
    pkg.text(partFor(REL_TYPES.settings, docDir + "/settings.xml")),
  );
  setContext(ctx);
  clearCssCaches();

  const media = new Map<string, MediaEntry>();
  loadMedia(pkg, docPart, rels, media, warnings);

  const docXml = pkg.text(docPart)!;
  const xml = parseXml(docXml);
  const root = xml.documentElement;
  const rootDecls = rootNamespaceDecls(root);
  const body = child(root, NS.w, "body");
  if (!body) throw new Error("Document has no body");
  const parser = new BodyParser(ctx, rels, media, rootDecls);
  const sect = parseSectPr(child(body, NS.w, "sectPr"), (e) => parser.ser(e));
  parser.contentWidthPx = Math.max(100, (sect.pgW - sect.marL - sect.marR) / 15);
  const blocks = parser.parseBlocks(body, false, null, null);
  if (!blocks.length) blocks.push(parser.emptyParagraph(false, null, null));
  const doc = schema.nodeFromJSON({ type: "doc", attrs: { sect }, content: blocks });
  warnings.push(...parser.warnings);

  // Headers & footers referenced by the final section.
  const headers = new Map<string, HeaderFooter>();
  const footers = new Map<string, HeaderFooter>();
  const loadHF = (rId: string, into: Map<string, HeaderFooter>) => {
    const rel = rels.get(rId);
    if (!rel) return;
    const part = resolveTarget(docPart, rel.target);
    const text = pkg.text(part);
    if (!text) return;
    try {
      const hfRels = parseRels(pkg.text(part.split("/").slice(0, -1).join("/") + "/_rels/" + part.split("/").pop() + ".rels"));
      const hfMedia = new Map<string, MediaEntry>();
      loadMedia(pkg, part, hfRels, hfMedia, warnings);
      const hx = parseXml(text);
      const p = new BodyParser(ctx, hfRels, hfMedia, rootNamespaceDecls(hx.documentElement));
      const hb = p.parseBlocks(hx.documentElement, false, null, null);
      if (!hb.length) hb.push(p.emptyParagraph(false, null, null));
      const hdoc = schema.nodeFromJSON({ type: "doc", attrs: { sect }, content: hb });
      into.set(rId, { rId, part, doc: hdoc, xml: text });
    } catch (e) {
      warnings.push("Failed to parse " + part + ": " + (e as Error).message);
    }
  };
  for (const id of Object.values(sect.headers)) if (id) loadHF(id, headers);
  for (const id of Object.values(sect.footers)) if (id) loadHF(id, footers);

  return { pkg, ctx, doc, docPart, docXml, rels, media, headers, footers, rootDecls, warnings };
}

/** Create an empty document with a built-in default template. */
export function emptyDocument(): PMNode {
  return schema.nodeFromJSON({
    type: "doc",
    attrs: { sect: { ...DEFAULT_SECT, headers: {}, footers: {} } },
    content: [{ type: "paragraph", attrs: { pPr: null, props: {}, inTable: false, tblStyle: null, sdt: null, sectPr: null } }],
  });
}

export { runText };
