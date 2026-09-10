// Resolve a cell's style index into drawing properties, with caching.
import { Workbook, Xf, Font, Fill, Border, BorderSide, Cell, Value, isError } from "./model";
import { resolveColor } from "./colors";
import { formatCodeFor, formatValue, Rendered } from "./numfmt";
import { cssTextColor, cssBgColor, isDarkMode } from "../docx/props";

export interface SideStyle { style: string; color: string; }
export interface CellStyle {
  fontName: string; fontSize: number; bold: boolean; italic: boolean; underline: boolean; strike: boolean;
  color: string | null;          // CSS colour or null (auto)
  fill: string | null;           // CSS colour or null
  borders: { top: SideStyle | null; bottom: SideStyle | null; left: SideStyle | null; right: SideStyle | null };
  halign: "left" | "center" | "right" | "general" | "fill" | "justify" | "centerContinuous";
  valign: "top" | "center" | "bottom";
  wrap: boolean; indent: number; rotation: number; shrink: boolean;
  numFmt: string; numFmtId: number;
}

const FONT_FALLBACK: Record<string, string> = {
  calibri: '"Calibri","Carlito",sans-serif', arial: '"Arial","Liberation Sans",sans-serif', "times new roman": '"Times New Roman","Liberation Serif",serif',
  cambria: '"Cambria","Caladea",serif', aptos: '"Aptos","Calibri","Carlito",sans-serif', "segoe ui": '"Segoe UI","Noto Sans",sans-serif', verdana: '"Verdana","DejaVu Sans",sans-serif',
  tahoma: '"Tahoma","DejaVu Sans",sans-serif', "courier new": '"Courier New","Liberation Mono",monospace', consolas: '"Consolas","DejaVu Sans Mono",monospace',
};
export function fontFamilyCss(name: string): string {
  return FONT_FALLBACK[name.toLowerCase()] || `"${name.replace(/"/g, "")}","Calibri","Carlito",sans-serif`;
}

/** Canvas cannot use `var(--x)` colours: resolve them to the current theme's concrete value. */
const varCache = new Map<string, string>();
function concrete(css: string | null): string | null {
  if (!css || !css.startsWith("var(")) return css;
  let v = varCache.get(css);
  if (v === undefined) {
    const name = css.slice(4, -1).trim();
    v = getComputedStyle(document.documentElement).getPropertyValue(name).trim() || (isDarkMode() ? "#e4e4e6" : "#000000");
    varCache.set(css, v);
  }
  return v;
}

export class StyleResolver {
  private cache = new Map<number, CellStyle>();
  version = 0;
  constructor(public wb: Workbook) {}

  invalidate() { this.cache.clear(); this.version++; varCache.clear(); }

  private side(s: BorderSide | null): SideStyle | null {
    if (!s) return null;
    const hex = resolveColor(s.color, this.wb.themeColors);
    return { style: s.style, color: hex ? concrete(cssTextColor(hex))! : (isDarkMode() ? "#8a8b90" : "#000000") };
  }

  get(idx: number): CellStyle {
    let cs = this.cache.get(idx);
    if (cs) return cs;
    const st = this.wb.styles;
    const xf: Xf = st.xfs[idx] || st.xfs[0];
    const font: Font = st.fonts[xf.fontId] || st.fonts[0];
    const fill: Fill = st.fills[xf.fillId] || st.fills[0];
    const border: Border = st.borders[xf.borderId] || st.borders[0];
    const fontColor = resolveColor(font.color, this.wb.themeColors);
    let fillHex: string | null = null;
    if (fill.patternType === "solid") fillHex = resolveColor(fill.fg, this.wb.themeColors) || resolveColor(fill.bg, this.wb.themeColors);
    else if (fill.patternType === "gradient") fillHex = resolveColor(fill.fg, this.wb.themeColors);
    else if (fill.patternType && fill.patternType !== "none" && fill.patternType !== "gray125") fillHex = resolveColor(fill.fg, this.wb.themeColors) || resolveColor(fill.bg, this.wb.themeColors);
    const al = xf.alignment || {};
    cs = {
      fontName: font.name || "Calibri", fontSize: font.size || 11, bold: font.bold, italic: font.italic, underline: !!font.underline && font.underline !== "none", strike: font.strike,
      color: fontColor ? concrete(cssTextColor(fontColor)) : null,
      fill: fillHex ? concrete(cssBgColor(fillHex)) : null,
      borders: { top: this.side(border.top), bottom: this.side(border.bottom), left: this.side(border.left), right: this.side(border.right) },
      halign: (al.horizontal as CellStyle["halign"]) || "general",
      valign: al.vertical === "top" ? "top" : al.vertical === "center" ? "center" : "bottom",
      wrap: !!al.wrapText, indent: al.indent || 0, rotation: al.textRotation || 0, shrink: !!al.shrinkToFit,
      numFmt: formatCodeFor(xf.numFmtId, st.numFmts), numFmtId: xf.numFmtId,
    };
    this.cache.set(idx, cs);
    return cs;
  }

  /** Displayed text for a cell (number formatting applied). */
  render(cell: Cell | undefined, style: CellStyle): Rendered {
    if (!cell) return { text: "", color: null, align: null };
    return formatValue(cell.v, style.numFmt, this.wb.date1904);
  }
}

export function isNumberLike(v: Value): boolean { return typeof v === "number" || typeof v === "boolean" || isError(v); }
