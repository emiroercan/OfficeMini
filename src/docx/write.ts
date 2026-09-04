// ProseMirror document -> document.xml (+ package updates). Untouched XML is
// re-emitted verbatim; only properties the editor can change are patched, and
// they are patched into the original property elements (schema order kept).
import { Node as PMNode, Mark } from "prosemirror-model";
import { schema, SectProps } from "../schema";
import { NS, parseXml, child, children, wattr, serialize, escapeXml, escapeXmlText, NO_BREAK_HYPHEN, SOFT_HYPHEN } from "./xml";
import { Package, parseRels, resolveTarget, REL_TYPES } from "./zip";
import { ctx, NumLevel, AbstractNum, NumDef } from "./styles";
import { parsePPr, parseRPr, parseTcPr, parseTblPr, ParaProps, RunProps, CellProps, TableProps, Border } from "./props";
import { pxToEmu, pxToTwips, fmt } from "./units";
import type { LoadedDoc } from "./parse";
import { TableMap } from "prosemirror-tables";

const PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];
const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"];
const TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"];
const TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"];

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const REQUIRED_DECLS: Record<string, string> = {
  w: NS.w, r: NS.r, wp: NS.wp, a: NS.a, pic: NS.pic, mc: NS.mc, v: NS.v, o: NS.o,
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  wps: NS.wps, wpg: NS.wpg,
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
};

/** Fragment (re)parsing with the document's namespace declarations in scope. */
class Frag {
  private wrapperOpen: string;
  constructor(private decls: Map<string, string>) {
    let open = "<om-root";
    for (const [p, uri] of decls) open += p ? ` xmlns:${p}="${uri}"` : ` xmlns="${uri}"`;
    for (const [p, uri] of Object.entries(REQUIRED_DECLS)) if (!decls.has(p)) open += ` xmlns:${p}="${uri}"`;
    this.wrapperOpen = open + ">";
  }
  parse(xml: string): Element {
    const doc = parseXml(this.wrapperOpen + xml + "</om-root>");
    const el = doc.documentElement.firstElementChild;
    if (!el) throw new Error("empty fragment");
    return el;
  }
  ser(el: Element): string { return serialize(el, this.decls); }
}

function wEl(doc: Document, name: string): Element { return doc.createElementNS(NS.w, "w:" + name); }
function setW(el: Element, name: string, value: string | number) { el.setAttributeNS(NS.w, "w:" + name, String(value)); }

/** Insert/replace child `name` respecting the schema order. Returns the element (or null when removed). */
function setChild(parent: Element, name: string, order: string[], remove = false): Element | null {
  let existing: Element | null = null;
  for (const c of children(parent, NS.w)) if (c.localName === name) { existing = c; break; }
  if (remove) { if (existing) parent.removeChild(existing); return null; }
  if (existing) return existing;
  const el = wEl(parent.ownerDocument!, name);
  const idx = order.indexOf(name);
  let before: Element | null = null;
  for (const c of children(parent, NS.w)) {
    const ci = order.indexOf(c.localName);
    if (ci > idx) { before = c; break; }
  }
  parent.insertBefore(el, before);
  return el;
}

function setValEl(parent: Element, name: string, order: string[], val: string | number | null | undefined) {
  if (val === null || val === undefined) { setChild(parent, name, order, true); return; }
  const el = setChild(parent, name, order)!;
  setW(el, "val", val);
}

function setToggle(parent: Element, name: string, order: string[], val: boolean | undefined) {
  if (val === undefined) { setChild(parent, name, order, true); return; }
  const el = setChild(parent, name, order)!;
  for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
  if (!val) setW(el, "val", "0");
}

function borderEl(doc: Document, name: string, b: Border): Element {
  const el = wEl(doc, name);
  setW(el, "val", b.val); setW(el, "sz", b.sz); setW(el, "space", b.space); setW(el, "color", b.color);
  return el;
}

interface WriteOptions { docPart: string; }

export class DocxWriter {
  private frag: Frag;
  private rprCache = new Map<string, string>();
  private rels: Map<string, { id: string; type: string; target: string; mode: string | null }>;
  private newRels: string[] = [];
  private nextRelId: number;
  private nextDocPrId: number;
  private newMedia: { part: string; bytes: Uint8Array; ext: string }[] = [];
  private mediaCounter: number;
  private docPart: string;
  private docDir: string;
  private themeFonts = ctx.theme;

  constructor(private loaded: LoadedDoc) {
    this.frag = new Frag(loaded.rootDecls);
    this.docPart = loaded.docPart;
    this.docDir = this.docPart.split("/").slice(0, -1).join("/");
    this.rels = loaded.rels;
    let max = 0;
    for (const id of this.rels.keys()) { const m = /^rId(\d+)$/.exec(id); if (m) max = Math.max(max, parseInt(m[1], 10)); }
    this.nextRelId = max + 1;
    const ids = loaded.docXml.match(/docPr id="(\d+)"/g) || [];
    this.nextDocPrId = ids.reduce((m, s) => Math.max(m, parseInt(/(\d+)/.exec(s)![1], 10)), 0) + 1;
    let mc = 0;
    for (const name of loaded.pkg.parts.keys()) { const m = /media\/image(\d+)\./.exec(name); if (m) mc = Math.max(mc, parseInt(m[1], 10)); }
    this.mediaCounter = mc + 1;
  }

  // ---- package level ------------------------------------------------------

  write(doc: PMNode): Uint8Array {
    const pkg = this.loaded.pkg;
    const bodyXml = this.writeBlocks(doc.content);
    const sectXml = this.writeSectPr(doc.attrs.sect as SectProps);
    // Reassemble document.xml around the original root element.
    const original = this.loaded.docXml;
    const rootMatch = /<w:document\b[^>]*>/.exec(original);
    if (!rootMatch) throw new Error("document root not found");
    let rootOpen = rootMatch[0];
    for (const [p, uri] of Object.entries(REQUIRED_DECLS)) if (!new RegExp(`xmlns:${p}=`).test(rootOpen)) rootOpen = rootOpen.replace(/>$/, ` xmlns:${p}="${uri}">`);
    // Preserve non-body children of the root (e.g. w:background).
    const xmlDoc = parseXml(original);
    let before = "", after = "";
    let seenBody = false;
    for (const c of children(xmlDoc.documentElement)) {
      if (c.namespaceURI === NS.w && c.localName === "body") { seenBody = true; continue; }
      if (seenBody) after += this.frag.ser(c); else before += this.frag.ser(c);
    }
    const out = XML_DECL + rootOpen + before + "<w:body>" + bodyXml + sectXml + "</w:body>" + after + "</w:document>";
    pkg.setText(this.docPart, out);
    this.flushRels(pkg);
    this.flushMedia(pkg);
    this.flushNumbering(pkg);
    return pkg.toBytes();
  }

  private flushRels(pkg: Package) {
    if (!this.newRels.length) return;
    const relsPart = this.docDir + "/_rels/" + this.docPart.split("/").pop() + ".rels";
    let xml = pkg.text(relsPart);
    if (!xml) xml = XML_DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    xml = xml.replace(/<\/Relationships>\s*$/, this.newRels.join("") + "</Relationships>");
    pkg.setText(relsPart, xml);
    this.newRels = [];
  }

  private flushMedia(pkg: Package) {
    if (!this.newMedia.length) return;
    let ct = pkg.text("[Content_Types].xml") || "";
    for (const m of this.newMedia) {
      pkg.set(m.part, m.bytes);
      const re = new RegExp(`<Default\\s+Extension="${m.ext}"`, "i");
      if (!re.test(ct)) {
        const mime = m.ext === "jpg" || m.ext === "jpeg" ? "image/jpeg" : m.ext === "svg" ? "image/svg+xml" : "image/" + m.ext;
        ct = ct.replace(/<Types\b[^>]*>/, (s) => s + `<Default Extension="${m.ext}" ContentType="${mime}"/>`);
      }
    }
    pkg.setText("[Content_Types].xml", ct);
    this.newMedia = [];
  }

  private flushNumbering(pkg: Package) {
    if (!ctx.newAbstractNums.length && !ctx.newNums.length) return;
    let part = this.docDir + "/numbering.xml";
    for (const r of this.rels.values()) if (r.type === REL_TYPES.numbering) { part = resolveTarget(this.docPart, r.target); break; }
    let xml = pkg.text(part);
    if (!xml) {
      xml = XML_DECL + `<w:numbering xmlns:w="${NS.w}" xmlns:r="${NS.r}"></w:numbering>`;
      this.newRels.push(`<Relationship Id="rId${this.nextRelId++}" Type="${REL_TYPES.numbering}" Target="numbering.xml"/>`);
      let ct = pkg.text("[Content_Types].xml") || "";
      if (!/numbering\.xml/.test(ct)) ct = ct.replace(/<\/Types>\s*$/, `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`);
      pkg.setText("[Content_Types].xml", ct);
      this.flushRels(pkg);
    }
    const abs = ctx.newAbstractNums.map((a) => this.abstractNumXml(a)).join("");
    const nums = ctx.newNums.map((n) => this.numXml(n)).join("");
    // abstractNum elements must precede num elements.
    const numIdx = xml.search(/<w:num\b/);
    if (numIdx >= 0) xml = xml.slice(0, numIdx) + abs + xml.slice(numIdx);
    else xml = xml.replace(/<\/w:numbering>\s*$/, abs + "</w:numbering>");
    xml = xml.replace(/<\/w:numbering>\s*$/, nums + "</w:numbering>");
    pkg.setText(part, xml);
    ctx.newAbstractNums = [];
    ctx.newNums = [];
  }

  private abstractNumXml(a: AbstractNum): string {
    let s = `<w:abstractNum w:abstractNumId="${a.id}"><w:multiLevelType w:val="${a.multiLevelType || "hybridMultilevel"}"/>`;
    for (const [i, l] of a.levels) s += this.levelXml(i, l);
    return s + "</w:abstractNum>";
  }
  private levelXml(i: number, l: NumLevel): string {
    let s = `<w:lvl w:ilvl="${i}"><w:start w:val="${l.start}"/><w:numFmt w:val="${l.numFmt}"/><w:lvlText w:val="${escapeXml(l.lvlText)}"/><w:lvlJc w:val="${l.lvlJc}"/>`;
    s += `<w:pPr><w:ind w:left="${l.pPr.indLeft ?? 720}" w:hanging="${l.pPr.indHanging ?? 360}"/></w:pPr>`;
    if (l.rPr.font) s += `<w:rPr><w:rFonts w:ascii="${escapeXml(l.rPr.font)}" w:hAnsi="${escapeXml(l.rPr.font)}" w:hint="default"/></w:rPr>`;
    return s + "</w:lvl>";
  }
  private numXml(n: NumDef): string {
    return `<w:num w:numId="${n.numId}"><w:abstractNumId w:val="${n.abstractNumId}"/></w:num>`;
  }

  private writeSectPr(sect: SectProps): string {
    let el: Element;
    if (sect.xml) el = this.frag.parse(sect.xml);
    else el = this.frag.parse("<w:sectPr/>");
    const order = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"];
    const pgSz = setChild(el, "pgSz", order)!;
    setW(pgSz, "w", sect.pgW); setW(pgSz, "h", sect.pgH);
    if (sect.orient) setW(pgSz, "orient", sect.orient); else pgSz.removeAttributeNS(NS.w, "orient");
    const mar = setChild(el, "pgMar", order)!;
    setW(mar, "top", sect.marT); setW(mar, "right", sect.marR); setW(mar, "bottom", sect.marB); setW(mar, "left", sect.marL);
    setW(mar, "header", sect.header); setW(mar, "footer", sect.footer); setW(mar, "gutter", sect.gutter);
    return this.frag.ser(el);
  }

  // ---- blocks -------------------------------------------------------------

  writeBlocks(fragment: { forEach(f: (n: PMNode) => void): void }): string {
    let out = "";
    let sdtId: number | null = null;
    let sdtBuf = "";
    let sdtInfo: any = null;
    const flushSdt = () => {
      if (sdtId !== null) {
        out += `<w:sdt>${sdtInfo.pr || ""}${sdtInfo.endPr || ""}<w:sdtContent>${sdtBuf}</w:sdtContent></w:sdt>`;
        sdtId = null; sdtBuf = ""; sdtInfo = null;
      }
    };
    fragment.forEach((node) => {
      const xml = this.writeBlock(node);
      const sdt = node.attrs.sdt;
      if (sdt && typeof sdt.id === "number") {
        if (sdtId !== sdt.id) { flushSdt(); sdtId = sdt.id; sdtInfo = sdt; }
        sdtBuf += xml;
      } else { flushSdt(); out += xml; }
    });
    flushSdt();
    return out;
  }

  writeBlock(node: PMNode): string {
    switch (node.type) {
      case schema.nodes.paragraph: return this.writeParagraph(node);
      case schema.nodes.table: return this.writeTable(node);
      case schema.nodes.opaque_block: return node.attrs.xml || "";
      default: return "";
    }
  }

  writeParagraph(node: PMNode): string {
    const props = node.attrs.props as ParaProps;
    const pPr = this.buildPPr(node.attrs.pPr, props, node.attrs.sectPr);
    return `<w:p>${pPr}${this.writeInlines(node)}</w:p>`;
  }

  private buildPPr(raw: string | null, cur: ParaProps, sectPr: string | null): string {
    let el: Element;
    let orig: ParaProps = {};
    if (raw) { el = this.frag.parse(raw); orig = parsePPr(el, this.themeFonts); }
    else {
      const keys = Object.keys(cur).filter((k) => (cur as any)[k] !== undefined && (cur as any)[k] !== null && k !== "rPr");
      if (!keys.length && !sectPr) return "";
      el = this.frag.parse("<w:pPr/>");
    }
    this.patchPPr(el, orig, cur);
    const existingSect = child(el, NS.w, "sectPr");
    if (sectPr) { if (existingSect) el.removeChild(existingSect); el.appendChild(this.frag.parse(sectPr)); }
    else if (existingSect) el.removeChild(existingSect);
    if (!el.firstChild && !el.attributes.length) return "";
    return this.frag.ser(el);
  }

  private patchPPr(el: Element, o: ParaProps, c: ParaProps) {
    const O = PPR_ORDER;
    if ((c.pStyle || null) !== (o.pStyle || null)) setValEl(el, "pStyle", O, c.pStyle || null);
    if ((c.jc || null) !== (o.jc || null)) setValEl(el, "jc", O, c.jc || null);
    const n = (v: number | null | undefined) => (v === undefined ? null : v);
    if (n(c.indLeft) !== n(o.indLeft) || n(c.indRight) !== n(o.indRight) || n(c.indHanging) !== n(o.indHanging) || n(c.indFirstLine) !== n(o.indFirstLine)) {
      if (n(c.indLeft) === null && n(c.indRight) === null && n(c.indHanging) === null && n(c.indFirstLine) === null) setChild(el, "ind", O, true);
      else {
        const ind = setChild(el, "ind", O)!;
        const setOrRemove = (name: string, v: number | null) => { if (v === null) ind.removeAttributeNS(NS.w, name); else setW(ind, name, v); };
        setOrRemove("left", n(c.indLeft)); ind.removeAttributeNS(NS.w, "start");
        setOrRemove("right", n(c.indRight)); ind.removeAttributeNS(NS.w, "end");
        setOrRemove("hanging", n(c.indHanging));
        setOrRemove("firstLine", n(c.indFirstLine));
      }
    }
    if (n(c.spBefore) !== n(o.spBefore) || n(c.spAfter) !== n(o.spAfter) || n(c.spLine) !== n(o.spLine) || (c.spLineRule || null) !== (o.spLineRule || null) || !!c.spBeforeAuto !== !!o.spBeforeAuto || !!c.spAfterAuto !== !!o.spAfterAuto) {
      const sp = setChild(el, "spacing", O)!;
      const setOrRemove = (name: string, v: number | string | null) => { if (v === null) sp.removeAttributeNS(NS.w, name); else setW(sp, name, v); };
      setOrRemove("before", n(c.spBefore)); setOrRemove("after", n(c.spAfter));
      setOrRemove("line", n(c.spLine)); setOrRemove("lineRule", c.spLine !== undefined && c.spLine !== null ? (c.spLineRule || "auto") : null);
      setOrRemove("beforeAutospacing", c.spBeforeAuto ? "1" : null); setOrRemove("afterAutospacing", c.spAfterAuto ? "1" : null);
      if (!sp.attributes.length) el.removeChild(sp);
    }
    if (n(c.numId) !== n(o.numId) || n(c.ilvl) !== n(o.ilvl)) {
      if (c.numId === null || c.numId === undefined) setChild(el, "numPr", O, true);
      else {
        const np = setChild(el, "numPr", O)!;
        while (np.firstChild) np.removeChild(np.firstChild);
        const il = wEl(el.ownerDocument!, "ilvl"); setW(il, "val", c.ilvl ?? 0); np.appendChild(il);
        const ni = wEl(el.ownerDocument!, "numId"); setW(ni, "val", c.numId); np.appendChild(ni);
      }
    }
    const tog = (key: keyof ParaProps, name: string) => { if (!!(c as any)[key] !== !!(o as any)[key]) setToggle(el, name, O, (c as any)[key] ? true : undefined); };
    tog("keepNext", "keepNext"); tog("keepLines", "keepLines"); tog("pageBreakBefore", "pageBreakBefore");
    tog("contextual", "contextualSpacing"); tog("bidi", "bidi");
    if ((c.shd || null) !== (o.shd || null)) {
      if (!c.shd) setChild(el, "shd", O, true);
      else { const sh = setChild(el, "shd", O)!; setW(sh, "val", "clear"); setW(sh, "color", "auto"); setW(sh, "fill", c.shd); }
    }
    if (c.bdrBottom !== o.bdrBottom || c.bdrTop !== o.bdrTop) {
      const hasAny = c.bdrTop || c.bdrBottom || c.bdrLeft || c.bdrRight;
      if (!hasAny) setChild(el, "pBdr", O, true);
      else if (!child(el, NS.w, "pBdr")) {
        const pb = setChild(el, "pBdr", O)!;
        for (const [name, b] of [["top", c.bdrTop], ["left", c.bdrLeft], ["bottom", c.bdrBottom], ["right", c.bdrRight]] as const) if (b) pb.appendChild(borderEl(el.ownerDocument!, name, b));
      }
    }
  }

  // ---- inline -------------------------------------------------------------

  private rPrFor(mark: Mark | undefined): string {
    if (!mark) return "";
    const key = (mark.attrs.xml || "") + " " + JSON.stringify(mark.attrs.props);
    let s = this.rprCache.get(key);
    if (s !== undefined) return s;
    const props = mark.attrs.props as RunProps;
    let el: Element, orig: RunProps = {};
    if (mark.attrs.xml) { el = this.frag.parse(mark.attrs.xml); orig = parseRPr(el, this.themeFonts); }
    else el = this.frag.parse("<w:rPr/>");
    this.patchRPr(el, orig, props);
    s = el.firstChild || el.attributes.length ? this.frag.ser(el) : "";
    this.rprCache.set(key, s);
    return s;
  }

  private patchRPr(el: Element, o: RunProps, c: RunProps) {
    const O = RPR_ORDER;
    if ((c.rStyle || null) !== (o.rStyle || null)) setValEl(el, "rStyle", O, c.rStyle || null);
    if ((c.font ?? null) !== (o.font ?? null)) {
      if (!c.font) setChild(el, "rFonts", O, true);
      else {
        const rf = setChild(el, "rFonts", O)!;
        for (const a of ["asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"]) rf.removeAttributeNS(NS.w, a);
        setW(rf, "ascii", c.font); setW(rf, "hAnsi", c.font); setW(rf, "cs", c.font);
        if (rf.hasAttributeNS(NS.w, "eastAsia")) setW(rf, "eastAsia", c.font);
      }
    }
    if ((c.size ?? null) !== (o.size ?? null)) { setValEl(el, "sz", O, c.size ?? null); setValEl(el, "szCs", O, c.size ?? null); }
    const tog = (key: keyof RunProps, name: string, cs?: string) => {
      if ((c as any)[key] !== (o as any)[key]) {
        setToggle(el, name, O, (c as any)[key] === undefined ? undefined : !!(c as any)[key]);
        if (cs) setToggle(el, cs, O, (c as any)[key] === undefined ? undefined : !!(c as any)[key]);
      }
    };
    tog("b", "b", "bCs"); tog("i", "i", "iCs"); tog("caps", "caps"); tog("smallCaps", "smallCaps");
    tog("strike", "strike"); tog("dstrike", "dstrike"); tog("vanish", "vanish"); tog("rtl", "rtl");
    if ((c.u ?? null) !== (o.u ?? null)) setValEl(el, "u", O, c.u ?? null);
    if ((c.color ?? null) !== (o.color ?? null)) {
      if (!c.color) setChild(el, "color", O, true);
      else { const ce = setChild(el, "color", O)!; for (const a of ["themeColor", "themeTint", "themeShade"]) ce.removeAttributeNS(NS.w, a); setW(ce, "val", c.color); }
    }
    if ((c.highlight ?? null) !== (o.highlight ?? null)) setValEl(el, "highlight", O, c.highlight && c.highlight !== "none" ? c.highlight : null);
    if ((c.shd ?? null) !== (o.shd ?? null)) {
      if (!c.shd) setChild(el, "shd", O, true);
      else { const sh = setChild(el, "shd", O)!; setW(sh, "val", "clear"); setW(sh, "color", "auto"); setW(sh, "fill", c.shd); }
    }
    if ((c.vertAlign ?? null) !== (o.vertAlign ?? null)) setValEl(el, "vertAlign", O, c.vertAlign && c.vertAlign !== "baseline" ? c.vertAlign : null);
    if ((c.spacing ?? null) !== (o.spacing ?? null)) setValEl(el, "spacing", O, c.spacing ?? null);
  }

  writeInlines(node: PMNode): string {
    let out = "";
    // Group by hyperlink mark, then by run mark.
    let curLink: Mark | null = null;
    let linkBuf = "";
    let runMark: Mark | undefined | null = null;
    let runBuf = "";
    let runOpen = false;
    const emit = (s: string) => { if (curLink) linkBuf += s; else out += s; };
    const flushRun = () => { if (runOpen) { emit(`<w:r>${this.rPrFor(runMark || undefined)}${runBuf}</w:r>`); runOpen = false; runBuf = ""; runMark = null; } };
    const flushLink = () => {
      flushRun();
      if (curLink) {
        const attrs = this.linkAttrs(curLink);
        out += `<w:hyperlink ${attrs}>${linkBuf}</w:hyperlink>`;
        curLink = null; linkBuf = "";
      }
    };
    const rprType = schema.marks.rpr, linkType = schema.marks.link;
    node.forEach((child) => {
      const link = linkType.isInSet(child.marks) || null;
      if (!(link === curLink || (link && curLink && link.eq(curLink)))) { flushLink(); curLink = link; }
      const rm = rprType.isInSet(child.marks);
      const t = child.type;
      if (t === schema.nodes.text || t === schema.nodes.tab) {
        if (!runOpen || !(rm === runMark || (rm && runMark && rm.eq(runMark)))) { flushRun(); runOpen = true; runMark = rm; }
        runBuf += t === schema.nodes.tab ? "<w:tab/>" : this.textXml(child.text || "");
        return;
      }
      flushRun();
      if (t === schema.nodes.hard_break) {
        const kind = child.attrs.kind;
        const br = kind === "page" ? '<w:br w:type="page"/>' : kind === "column" ? '<w:br w:type="column"/>' : child.attrs.clear ? `<w:br w:clear="${child.attrs.clear}"/>` : "<w:br/>";
        emit(`<w:r>${this.rPrFor(rm)}${br}</w:r>`);
      } else if (t === schema.nodes.image) {
        emit(`<w:r>${this.rPrFor(rm)}${this.imageXml(child)}</w:r>`);
      } else if (t === schema.nodes.textbox) {
        emit(`<w:r>${this.rPrFor(rm)}${this.textboxXml(child)}</w:r>`);
      } else if (t === schema.nodes.opaque_inline) {
        emit(child.attrs.xml || "");
      }
    });
    flushLink();
    return out;
  }

  private textXml(text: string): string {
    let out = "", buf = "";
    const flush = () => { if (buf) { out += `<w:t xml:space="preserve">${escapeXmlText(buf)}</w:t>`; buf = ""; } };
    for (const ch of text) {
      if (ch === NO_BREAK_HYPHEN) { flush(); out += "<w:noBreakHyphen/>"; }
      else if (ch === SOFT_HYPHEN) { flush(); out += "<w:softHyphen/>"; }
      else if (ch === "\t") { flush(); out += "<w:tab/>"; }
      else if (ch === "\n") { flush(); out += "<w:br/>"; }
      else buf += ch;
    }
    flush();
    return out;
  }

  private linkAttrs(mark: Mark): string {
    const a = mark.attrs;
    if (a.raw) return a.raw;
    if (a.anchor) return `w:anchor="${escapeXml(a.anchor)}" w:history="1"`;
    const id = `rId${this.nextRelId++}`;
    this.newRels.push(`<Relationship Id="${id}" Type="${REL_TYPES.hyperlink}" Target="${escapeXml(a.href || "")}" TargetMode="External"/>`);
    return `r:id="${id}" w:history="1"`;
  }

  private imageXml(node: PMNode): string {
    const a = node.attrs;
    const cx = pxToEmu(a.w), cy = pxToEmu(a.h);
    if (a.raw) {
      let raw: string = a.raw;
      if (a.origW !== null && (Math.abs(a.w - a.origW) > 0.5 || Math.abs(a.h - a.origH) > 0.5)) {
        raw = raw.replace(/<wp:extent\s+cx="\d+"\s+cy="\d+"/, `<wp:extent cx="${cx}" cy="${cy}"`);
        raw = raw.replace(/<a:ext\s+cx="\d+"\s+cy="\d+"\/>/, `<a:ext cx="${cx}" cy="${cy}"/>`);
        raw = raw.replace(/(<v:shape\b[^>]*style=")([^"]*)"/, (_m, p1, style) => {
          const st = style.replace(/width:[^;]+/, `width:${fmt(a.w * 0.75)}pt`).replace(/height:[^;]+/, `height:${fmt(a.h * 0.75)}pt`);
          return `${p1}${st}"`;
        });
      }
      return raw;
    }
    // New image: add media + relationship.
    let rId = a.rId as string | null;
    if (!rId && a.media) {
      const ext = (a.media.ext || "png").toLowerCase().replace("jpeg", "jpeg");
      const part = `${this.docDir}/media/image${this.mediaCounter++}.${ext}`;
      this.newMedia.push({ part, bytes: a.media.bytes, ext });
      rId = `rId${this.nextRelId++}`;
      this.newRels.push(`<Relationship Id="${rId}" Type="${REL_TYPES.image}" Target="media/${part.split("/").pop()}"/>`);
    }
    if (!rId) return "";
    const id = this.nextDocPrId++;
    const name = escapeXml(a.name || `Picture ${id}`);
    return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${name}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  }

  private textboxXml(node: PMNode): string {
    const raw: string = node.attrs.raw || "";
    const content = this.writeBlocks(node.content);
    if (!raw) return "";
    // Replace every txbxContent body (DrawingML choice and VML fallback) with the edited content.
    return raw.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>|<w:txbxContent\/>/g, `<w:txbxContent>${content}</w:txbxContent>`);
  }

  // ---- tables -------------------------------------------------------------

  writeTable(node: PMNode): string {
    const props = node.attrs.props as TableProps;
    const map = TableMap.get(node);
    // Column widths: prefer live colwidth attrs (column resizing) over the parsed grid.
    const grid: number[] = (node.attrs.grid as number[] | null) ? [...(node.attrs.grid as number[])] : new Array(map.width).fill(Math.round(9000 / map.width));
    for (let col = 0; col < map.width; col++) {
      for (let row = 0; row < map.height; row++) {
        const cell = node.nodeAt(map.map[row * map.width + col]);
        if (cell && cell.attrs.colspan === 1 && cell.attrs.colwidth && cell.attrs.colwidth[0]) {
          const tw = pxToTwips(cell.attrs.colwidth[0]);
          if (Math.abs(tw - grid[col]) > 10) grid[col] = tw;
          break;
        }
      }
    }
    while (grid.length < map.width) grid.push(grid[grid.length - 1] || 1000);
    let out = "<w:tbl>" + this.buildTblPr(node.attrs.tblPr, props);
    out += "<w:tblGrid>" + grid.slice(0, map.width).map((w) => `<w:gridCol w:w="${w}"/>`).join("") + "</w:tblGrid>";
    // Track vertical merges: pending[col] = {remaining rows, span, tcPrs, index}
    const pending = new Map<number, { rows: number; span: number; tcPrs: string[]; i: number; width: number }>();
    node.forEach((row) => {
      out += "<w:tr>" + (row.attrs.trPr || "");
      let col = 0;
      const emitPending = () => {
        const p = pending.get(col);
        if (!p) return false;
        const rawTcPr = p.tcPrs[p.i];
        let tcPr: string;
        if (rawTcPr) {
          const el = this.frag.parse(rawTcPr);
          const cp = parseTcPr(el);
          if (cp.vMerge !== "continue") { const vm = setChild(el, "vMerge", TCPR_ORDER)!; for (const at of Array.from(vm.attributes)) vm.removeAttribute(at.name); }
          tcPr = this.frag.ser(el);
        } else {
          tcPr = `<w:tcPr><w:tcW w:w="${p.width}" w:type="dxa"/>${p.span > 1 ? `<w:gridSpan w:val="${p.span}"/>` : ""}<w:vMerge/></w:tcPr>`;
        }
        out += `<w:tc>${tcPr}<w:p/></w:tc>`;
        p.i++; p.rows--;
        if (p.rows <= 0) pending.delete(col);
        col += p.span;
        return true;
      };
      row.forEach((cell) => {
        while (emitPending()) { /* continuation cells before this one */ }
        const cprops = cell.attrs.props as CellProps;
        const span = cell.attrs.colspan as number;
        const width = grid.slice(col, col + span).reduce((a, b) => a + b, 0);
        if ((cprops as any).pad && cell.textContent === "" && cell.attrs.tcPr === null) {
          // Padding cell (not in the original): only write it if the row would otherwise be shorter than the grid
          // — Word tolerates short rows, so skip it.
          col += span;
          return;
        }
        out += "<w:tc>" + this.buildTcPr(cell.attrs.tcPr, cprops, span, cell.attrs.rowspan as number, width);
        const inner = this.writeBlocks(cell.content);
        out += (inner || "<w:p/>") + "</w:tc>";
        if ((cell.attrs.rowspan as number) > 1) pending.set(col, { rows: cell.attrs.rowspan - 1, span, tcPrs: cell.attrs.vmergeTcPr || [], i: 0, width });
        col += span;
      });
      while (col < map.width && emitPending()) { /* trailing continuation cells */ }
      // any remaining pending cells at columns beyond what this row emitted
      for (const [c] of Array.from(pending.entries()).sort((a, b) => a[0] - b[0])) {
        if (c >= col) { col = c; emitPending(); }
      }
      out += "</w:tr>";
    });
    return out + "</w:tbl>";
  }

  private buildTblPr(raw: string | null, cur: TableProps): string {
    let el: Element, orig: TableProps = {};
    if (raw) { el = this.frag.parse(raw); orig = parseTblPr(el); }
    else el = this.frag.parse("<w:tblPr/>");
    const O = TBLPR_ORDER;
    if ((cur.tblStyle || null) !== (orig.tblStyle || null)) setValEl(el, "tblStyle", O, cur.tblStyle || null);
    if (!child(el, NS.w, "tblW")) { const w = setChild(el, "tblW", O)!; setW(w, "w", cur.width?.w ?? 0); setW(w, "type", cur.width?.type ?? "auto"); }
    if ((cur.jc || null) !== (orig.jc || null)) setValEl(el, "jc", O, cur.jc || null);
    if (cur.borders !== orig.borders && cur.borders) {
      const tb = setChild(el, "tblBorders", O)!;
      while (tb.firstChild) tb.removeChild(tb.firstChild);
      for (const name of ["top", "left", "bottom", "right", "insideH", "insideV"] as const) { const b = cur.borders[name]; if (b) tb.appendChild(borderEl(el.ownerDocument!, name, b)); }
    }
    if (!raw && !child(el, NS.w, "tblLook")) { const lk = setChild(el, "tblLook", O)!; setW(lk, "val", "04A0"); setW(lk, "firstRow", "1"); setW(lk, "lastRow", "0"); setW(lk, "firstColumn", "1"); setW(lk, "lastColumn", "0"); setW(lk, "noHBand", "0"); setW(lk, "noVBand", "1"); }
    return this.frag.ser(el);
  }

  private buildTcPr(raw: string | null, cur: CellProps, span: number, rowspan: number, width: number): string {
    let el: Element, orig: CellProps = {};
    if (raw) { el = this.frag.parse(raw); orig = parseTcPr(el); }
    else el = this.frag.parse("<w:tcPr/>");
    const O = TCPR_ORDER;
    if (!child(el, NS.w, "tcW")) { const w = setChild(el, "tcW", O)!; setW(w, "w", width); setW(w, "type", "dxa"); }
    else if (orig.width && orig.width.type === "dxa" && Math.abs(orig.width.w - width) > 10) { const w = child(el, NS.w, "tcW")!; setW(w, "w", width); }
    if (span !== (orig.gridSpan || 1)) setValEl(el, "gridSpan", O, span > 1 ? span : null);
    if (rowspan > 1) { const vm = setChild(el, "vMerge", O)!; setW(vm, "val", "restart"); }
    else if (orig.vMerge) setChild(el, "vMerge", O, true);
    if ((cur.shd ?? null) !== (orig.shd ?? null)) {
      if (!cur.shd) setChild(el, "shd", O, true);
      else { const sh = setChild(el, "shd", O)!; setW(sh, "val", "clear"); setW(sh, "color", "auto"); setW(sh, "fill", cur.shd); }
    }
    if ((cur.vAlign ?? null) !== (orig.vAlign ?? null)) setValEl(el, "vAlign", O, cur.vAlign ?? null);
    return this.frag.ser(el);
  }
}

export function writeDocx(loaded: LoadedDoc, doc: PMNode): Uint8Array {
  return new DocxWriter(loaded).write(doc);
}

export { parseRels, wattr };
