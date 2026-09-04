// Unit conversions. Layout is done at 96 dpi CSS pixels.
export const TWIP_PX = 96 / 1440;
export const EMU_PX = 96 / 914400;

export const twipsToPx = (tw: number) => tw * TWIP_PX;
export const pxToTwips = (px: number) => Math.round(px / TWIP_PX);
export const twipsToPt = (tw: number) => tw / 20;
export const ptToTwips = (pt: number) => Math.round(pt * 20);
export const emuToPx = (emu: number) => emu * EMU_PX;
export const pxToEmu = (px: number) => Math.round(px / EMU_PX);
export const halfPtToPt = (hp: number) => hp / 2;
export const eighthPtToPx = (e: number) => (e / 8) * (96 / 72);
export const ptToPx = (pt: number) => pt * (96 / 72);

export function fmt(n: number, digits = 2): string {
  const s = n.toFixed(digits);
  return s.indexOf(".") >= 0 ? s.replace(/\.?0+$/, "") : s;
}

/** Parse a CSS-ish length used in VML styles ("12pt", "1.5in", "96px", "2cm") to px. */
export function cssLenToPx(v: string | null | undefined): number | null {
  if (!v) return null;
  const m = /^\s*(-?[\d.]+)\s*(pt|px|in|cm|mm|em|%)?\s*$/.exec(v);
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch (m[2]) {
    case "pt": return ptToPx(n);
    case "in": return n * 96;
    case "cm": return (n / 2.54) * 96;
    case "mm": return (n / 25.4) * 96;
    case "em": return n * 16;
    case "%": return null;
    default: return n;
  }
}
