// styles.xml / theme / numbering / settings parsing and effective-property resolution.
import { NS, child, children, wattr, wint, parseXml } from "./xml";
import {
  ParaProps, RunProps, TableProps, CellProps, parsePPr, parseRPr, parseTblPr, parseTcPr, merge,
} from "./props";

export interface StyleDef {
  id: string;
  type: string; // paragraph | character | table | numbering
  name: string;
  basedOn: string | null;
  next: string | null;
  link: string | null;
  isDefault: boolean;
  pPr: ParaProps;
  rPr: RunProps;
  tblPr: TableProps;
  tcPr: CellProps;
  numId: number | null; // for numbering styles (via pPr/numPr)
}

export interface NumLevel {
  ilvl: number;
  start: number;
  numFmt: string;
  lvlText: string;
  lvlJc: string;
  pPr: ParaProps;
  rPr: RunProps;
  suff: string;
  isLgl: boolean;
  lvlRestart: number | null;
}

export interface AbstractNum {
  id: number;
  levels: Map<number, NumLevel>;
  numStyleLink: string | null;
  styleLink: string | null;
  multiLevelType: string | null;
}

export interface NumDef {
  numId: number;
  abstractNumId: number;
  overrides: Map<number, { start: number | null; lvl: NumLevel | null }>;
}

export interface ThemeFonts { major: string; minor: string; }

export interface ResolvedPara { pPr: ParaProps; rPr: RunProps; }

/** Everything needed to render a document: styles, numbering, theme, settings. */
export class DocContext {
  styles = new Map<string, StyleDef>();
  defaultPara: string | null = null;
  defaultChar: string | null = null;
  defaultTable: string | null = null;
  docPPr: ParaProps = {};
  docRPr: RunProps = {};
  theme: ThemeFonts = { major: "Calibri Light", minor: "Calibri" };
  abstractNums = new Map<number, AbstractNum>();
  nums = new Map<number, NumDef>();
  defaultTabStop = 720;
  evenOddHeaders = false;
  private paraCache = new Map<string, ResolvedPara>();
  private charCache = new Map<string, RunProps>();
  private tableCache = new Map<string, { tblPr: TableProps; pPr: ParaProps; rPr: RunProps; tcPr: CellProps }>();

  static parse(stylesXml?: string, themeXml?: string, numberingXml?: string, settingsXml?: string): DocContext {
    const c = new DocContext();
    if (themeXml) c.parseTheme(themeXml);
    if (stylesXml) c.parseStyles(stylesXml);
    if (numberingXml) c.parseNumbering(numberingXml);
    if (settingsXml) c.parseSettings(settingsXml);
    return c;
  }

  themeColors: Record<string, string> = {
    dk1: "000000", lt1: "FFFFFF", dk2: "44546A", lt2: "E7E6E6", accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5",
    accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47", hlink: "0563C1", folHlink: "954F72",
  };

  parseTheme(xml: string) {
    try {
      const doc = parseXml(xml);
      const major = doc.getElementsByTagNameNS(NS.a, "majorFont")[0];
      const minor = doc.getElementsByTagNameNS(NS.a, "minorFont")[0];
      const latin = (el?: Element) => el && child(el, NS.a, "latin")?.getAttribute("typeface");
      const mj = latin(major), mn = latin(minor);
      if (mj) this.theme.major = mj;
      if (mn) this.theme.minor = mn;
      const scheme = doc.getElementsByTagNameNS(NS.a, "clrScheme")[0];
      if (scheme) {
        for (const c of children(scheme, NS.a)) {
          const srgb = child(c, NS.a, "srgbClr")?.getAttribute("val");
          const sys = child(c, NS.a, "sysClr")?.getAttribute("lastClr");
          const v = srgb || sys;
          if (v) this.themeColors[c.localName] = v.toUpperCase();
        }
      }
    } catch { /* keep defaults */ }
  }

  /** Resolve a DrawingML scheme colour (with lumMod/lumOff) to RRGGBB. */
  schemeColor(name: string, lumMod?: number, lumOff?: number): string {
    const alias: Record<string, string> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" };
    let hex = this.themeColors[alias[name] || name] || "000000";
    if (lumMod === undefined && lumOff === undefined) return hex;
    const r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0, l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
      else if (max === g) h = ((b - r) / d + 2) / 6;
      else h = ((r - g) / d + 4) / 6;
    }
    if (lumMod !== undefined) l = l * (lumMod / 100000);
    if (lumOff !== undefined) l = l + lumOff / 100000;
    l = Math.max(0, Math.min(1, l));
    const hue2rgb = (p: number, q: number, t: number) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    let R: number, G: number, B: number;
    if (s === 0) R = G = B = l;
    else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; R = hue2rgb(p, q, h + 1 / 3); G = hue2rgb(p, q, h); B = hue2rgb(p, q, h - 1 / 3); }
    return [R, G, B].map((v) => Math.round(v * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
  }

  parseStyles(xml: string) {
    const doc = parseXml(xml);
    const root = doc.documentElement;
    const dd = child(root, NS.w, "docDefaults");
    if (dd) {
      const rd = child(child(dd, NS.w, "rPrDefault"), NS.w, "rPr");
      const pd = child(child(dd, NS.w, "pPrDefault"), NS.w, "pPr");
      this.docRPr = parseRPr(rd, this.theme);
      this.docPPr = parsePPr(pd, this.theme);
    }
    if (this.docRPr.size === undefined) this.docRPr.size = 20;
    if (this.docRPr.font === undefined) this.docRPr.font = this.theme.minor;
    for (const s of children(root, NS.w, "style")) {
      const id = wattr(s, "styleId");
      if (!id) continue;
      const type = wattr(s, "type") || "paragraph";
      const def: StyleDef = {
        id, type,
        name: wattr(child(s, NS.w, "name"), "val") || id,
        basedOn: wattr(child(s, NS.w, "basedOn"), "val"),
        next: wattr(child(s, NS.w, "next"), "val"),
        link: wattr(child(s, NS.w, "link"), "val"),
        isDefault: wattr(s, "default") === "1" || wattr(s, "default") === "true",
        pPr: parsePPr(child(s, NS.w, "pPr"), this.theme),
        rPr: parseRPr(child(s, NS.w, "rPr"), this.theme),
        tblPr: parseTblPr(child(s, NS.w, "tblPr")),
        tcPr: parseTcPr(child(s, NS.w, "tcPr")),
        numId: null,
      };
      def.numId = def.pPr.numId ?? null;
      this.styles.set(id, def);
      if (def.isDefault) {
        if (type === "paragraph") this.defaultPara = id;
        else if (type === "character") this.defaultChar = id;
        else if (type === "table") this.defaultTable = id;
      }
    }
    if (!this.defaultPara && this.styles.has("Normal")) this.defaultPara = "Normal";
  }

  parseNumbering(xml: string) {
    const doc = parseXml(xml);
    const root = doc.documentElement;
    for (const an of children(root, NS.w, "abstractNum")) {
      const id = wint(an, "abstractNumId");
      if (id === null) continue;
      const a: AbstractNum = {
        id, levels: new Map(),
        numStyleLink: wattr(child(an, NS.w, "numStyleLink"), "val"),
        styleLink: wattr(child(an, NS.w, "styleLink"), "val"),
        multiLevelType: wattr(child(an, NS.w, "multiLevelType"), "val"),
      };
      for (const lvl of children(an, NS.w, "lvl")) {
        const l = this.parseLevel(lvl);
        if (l) a.levels.set(l.ilvl, l);
      }
      this.abstractNums.set(id, a);
    }
    for (const num of children(root, NS.w, "num")) {
      const numId = wint(num, "numId");
      const abs = wint(child(num, NS.w, "abstractNumId"), "val");
      if (numId === null || abs === null) continue;
      const d: NumDef = { numId, abstractNumId: abs, overrides: new Map() };
      for (const ov of children(num, NS.w, "lvlOverride")) {
        const ilvl = wint(ov, "ilvl");
        if (ilvl === null) continue;
        const start = wint(child(ov, NS.w, "startOverride"), "val");
        const lvlEl = child(ov, NS.w, "lvl");
        d.overrides.set(ilvl, { start, lvl: lvlEl ? this.parseLevel(lvlEl) : null });
      }
      this.nums.set(numId, d);
    }
  }

  private parseLevel(lvl: Element): NumLevel | null {
    const ilvl = wint(lvl, "ilvl");
    if (ilvl === null) return null;
    return {
      ilvl,
      start: wint(child(lvl, NS.w, "start"), "val") ?? 1,
      numFmt: wattr(child(lvl, NS.w, "numFmt"), "val") || "decimal",
      lvlText: wattr(child(lvl, NS.w, "lvlText"), "val") ?? "",
      lvlJc: wattr(child(lvl, NS.w, "lvlJc"), "val") || "left",
      pPr: parsePPr(child(lvl, NS.w, "pPr"), this.theme),
      rPr: parseRPr(child(lvl, NS.w, "rPr"), this.theme),
      suff: wattr(child(lvl, NS.w, "suff"), "val") || "tab",
      isLgl: !!child(lvl, NS.w, "isLgl"),
      lvlRestart: wint(child(lvl, NS.w, "lvlRestart"), "val"),
    };
  }

  parseSettings(xml: string) {
    try {
      const doc = parseXml(xml);
      const root = doc.documentElement;
      const dts = wint(child(root, NS.w, "defaultTabStop"), "val");
      if (dts !== null && dts > 0) this.defaultTabStop = dts;
      this.evenOddHeaders = !!child(root, NS.w, "evenAndOddHeaders");
    } catch { /* ignore */ }
  }

  style(id: string | null | undefined): StyleDef | undefined {
    return id ? this.styles.get(id) : undefined;
  }

  /** Style chain root-first (base styles first, the style itself last). */
  private chain(id: string | null | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur = id ? this.styles.get(id) : undefined;
    while (cur && !seen.has(cur.id)) {
      seen.add(cur.id);
      out.unshift(cur);
      cur = cur.basedOn ? this.styles.get(cur.basedOn) : undefined;
    }
    return out;
  }

  /** Paragraph style resolved with docDefaults (without table style or numbering). */
  resolveParaStyle(id: string | null | undefined): ResolvedPara {
    const key = id || "";
    let r = this.paraCache.get(key);
    if (r) return r;
    let pPr: ParaProps = { ...this.docPPr };
    let rPr: RunProps = { ...this.docRPr };
    const styleId = id && this.styles.has(id) ? id : this.defaultPara;
    for (const s of this.chain(styleId)) {
      pPr = merge(pPr, s.pPr);
      rPr = merge(rPr, s.rPr);
    }
    // The paragraph-mark rPr of the style chain is not text formatting; drop it.
    delete (pPr as any).rPr;
    r = { pPr, rPr };
    this.paraCache.set(key, r);
    return r;
  }

  resolveCharStyle(id: string | null | undefined): RunProps {
    if (!id) return {};
    let r = this.charCache.get(id);
    if (r) return r;
    r = {};
    for (const s of this.chain(id)) r = merge(r, s.rPr);
    delete (r as any).rStyle;
    this.charCache.set(id, r);
    return r;
  }

  resolveTableStyle(id: string | null | undefined) {
    const key = id || "";
    let r = this.tableCache.get(key);
    if (r) return r;
    let tblPr: TableProps = {}, pPr: ParaProps = {}, rPr: RunProps = {}, tcPr: CellProps = {};
    const ids = this.chain(id && this.styles.has(id) ? id : this.defaultTable);
    for (const s of ids) {
      tblPr = merge(tblPr, s.tblPr);
      pPr = merge(pPr, s.pPr);
      rPr = merge(rPr, s.rPr);
      tcPr = merge(tcPr, s.tcPr);
    }
    if (!tblPr.cellMar) tblPr.cellMar = { left: 108, right: 108, top: 0, bottom: 0 };
    r = { tblPr, pPr, rPr, tcPr };
    this.tableCache.set(key, r);
    return r;
  }

  /** Numbering level for a numId/ilvl, honoring overrides and style links. */
  numLevel(numId: number | null | undefined, ilvl: number | null | undefined): NumLevel | null {
    if (!numId) return null;
    const num = this.nums.get(numId);
    if (!num) return null;
    const lvl = ilvl ?? 0;
    const ov = num.overrides.get(lvl);
    if (ov?.lvl) return ov.lvl;
    let abs = this.abstractNums.get(num.abstractNumId);
    let guard = 0;
    while (abs && abs.numStyleLink && guard++ < 5) {
      const st = this.styles.get(abs.numStyleLink);
      const linkedNum = st?.numId ? this.nums.get(st.numId) : undefined;
      abs = linkedNum ? this.abstractNums.get(linkedNum.abstractNumId) : undefined;
    }
    return abs?.levels.get(lvl) || null;
  }

  /** Numbering definitions created in this session (to be appended to numbering.xml on save). */
  newAbstractNums: AbstractNum[] = [];
  newNums: NumDef[] = [];

  /** Find the style id whose (language independent) name matches. */
  styleIdByName(name: string): string | null {
    for (const s of this.styles.values()) if (s.name.toLowerCase() === name.toLowerCase() || s.id.toLowerCase() === name.toLowerCase()) return s.id;
    return null;
  }

  /** Create a new list definition (Word defaults) and return its numId. */
  createList(kind: "bullet" | "decimal"): number {
    let absId = 0;
    for (const id of this.abstractNums.keys()) absId = Math.max(absId, id + 1);
    let numId = 1;
    for (const id of this.nums.keys()) numId = Math.max(numId, id + 1);
    const levels = new Map<number, NumLevel>();
    const bulletChars = [String.fromCharCode(0xf0b7), "o", String.fromCharCode(0xf0a7)];
    const bulletFonts = ["Symbol", "Courier New", "Wingdings"];
    const fmts = ["decimal", "lowerLetter", "lowerRoman"];
    for (let i = 0; i < 9; i++) {
      const isBullet = kind === "bullet";
      levels.set(i, {
        ilvl: i,
        start: 1,
        numFmt: isBullet ? "bullet" : fmts[i % 3],
        lvlText: isBullet ? bulletChars[i % 3] : `%${i + 1}.`,
        lvlJc: !isBullet && fmts[i % 3] === "lowerRoman" ? "right" : "left",
        pPr: { indLeft: 720 * (i + 1), indHanging: 360 },
        rPr: isBullet ? { font: bulletFonts[i % 3] } : {},
        suff: "tab",
        isLgl: false,
        lvlRestart: null,
      });
    }
    const abs: AbstractNum = { id: absId, levels, numStyleLink: null, styleLink: null, multiLevelType: "hybridMultilevel" };
    const num: NumDef = { numId, abstractNumId: absId, overrides: new Map() };
    this.abstractNums.set(absId, abs);
    this.nums.set(numId, num);
    this.newAbstractNums.push(abs);
    this.newNums.push(num);
    return numId;
  }

  /** Is this numId a bullet list (level 0)? */
  isBulletList(numId: number | null | undefined): boolean {
    const lvl = this.numLevel(numId, 0);
    return !!lvl && lvl.numFmt === "bullet";
  }

  numStart(numId: number, ilvl: number): number {
    const num = this.nums.get(numId);
    const ov = num?.overrides.get(ilvl);
    if (ov && ov.start !== null) return ov.start;
    return this.numLevel(numId, ilvl)?.start ?? 1;
  }

  /**
   * Effective paragraph props: docDefaults -> table style -> paragraph style
   * -> numbering level indent -> direct formatting.
   */
  effectivePara(direct: ParaProps, tblStyle: string | null | undefined, inTable: boolean): ResolvedPara {
    let pPr: ParaProps, rPr: RunProps;
    if (inTable) {
      const ts = this.resolveTableStyle(tblStyle);
      pPr = merge(merge(this.docPPr, ts.pPr), {});
      rPr = merge(merge(this.docRPr, ts.rPr), {});
      const styleId = direct.pStyle && this.styles.has(direct.pStyle) ? direct.pStyle : this.defaultPara;
      for (const s of this.chain(styleId)) { pPr = merge(pPr, s.pPr); rPr = merge(rPr, s.rPr); }
      delete (pPr as any).rPr;
    } else {
      const base = this.resolveParaStyle(direct.pStyle);
      pPr = base.pPr; rPr = base.rPr;
    }
    // Numbering: direct numPr wins over style numPr; numId 0 disables.
    const numId = direct.numId !== undefined && direct.numId !== null ? direct.numId : pPr.numId ?? null;
    const ilvl = direct.numId !== undefined && direct.numId !== null ? (direct.ilvl ?? 0) : (pPr.ilvl ?? 0);
    if (numId) {
      const lvl = this.numLevel(numId, ilvl);
      if (lvl) {
        const ind: ParaProps = {};
        if (lvl.pPr.indLeft !== undefined) ind.indLeft = lvl.pPr.indLeft;
        if (lvl.pPr.indHanging !== undefined) { ind.indHanging = lvl.pPr.indHanging; ind.indFirstLine = null; }
        else if (lvl.pPr.indFirstLine !== undefined) { ind.indFirstLine = lvl.pPr.indFirstLine; ind.indHanging = null; }
        if (lvl.pPr.tabs) ind.tabs = lvl.pPr.tabs;
        pPr = merge(pPr, ind);
      }
    }
    pPr = merge(pPr, direct);
    pPr.numId = numId; pPr.ilvl = ilvl;
    return { pPr, rPr };
  }
}

/** The context of the document currently loaded in this window. */
export let ctx: DocContext = new DocContext();
export function setContext(c: DocContext) { ctx = c; }
