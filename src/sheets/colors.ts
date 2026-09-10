// Colour resolution for SpreadsheetML: rgb / theme+tint / indexed / auto.
import { Color } from "./model";

// Excel's default indexed palette (0..63); 64 = system foreground, 65 = system background.
const INDEXED = ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"];

function hexToHsl(hex: string): [number, number, number] {
  const r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  return [h, s, l];
}
function hslToHex(h: number, s: number, l: number): string {
  const f = (p: number, q: number, t: number) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  let r: number, g: number, b: number;
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
  return [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** Apply an Excel tint (-1..1) to a colour. */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = hexToHsl(hex);
  const l2 = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return hslToHex(h, s, Math.max(0, Math.min(1, l2)));
}

/** Resolve to "RRGGBB" (uppercase) or null for automatic. themeColors are in scheme order dk1,lt1,dk2,lt2,accent1-6,hlink,folHlink. */
export function resolveColor(c: Color | null | undefined, themeColors: string[]): string | null {
  if (!c || c.auto) return null;
  if (c.rgb) {
    const hex = c.rgb.length === 8 ? c.rgb.slice(2) : c.rgb;
    return applyTint(hex.toUpperCase(), c.tint || 0);
  }
  if (c.theme !== undefined) {
    // Spreadsheet theme indexes swap the first two pairs: 0=lt1 1=dk1 2=lt2 3=dk2.
    const map = [1, 0, 3, 2];
    const idx = c.theme < 4 ? map[c.theme] : c.theme;
    const hex = themeColors[idx] || "000000";
    return applyTint(hex, c.tint || 0);
  }
  if (c.indexed !== undefined) {
    if (c.indexed === 64) return null;   // system foreground (auto)
    if (c.indexed === 65) return "FFFFFF";
    return INDEXED[c.indexed] || null;
  }
  return null;
}
