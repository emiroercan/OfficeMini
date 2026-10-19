// Reading themes for Markdown documents: paper and text colours, and the body font.
// Word documents keep their own formatting; these only restyle .md files, the way an e-reader
// does. Printing and PDF export always use white paper (the chosen font is kept).

export interface ReadingTheme { id: string; name: string; paper: string; text: string; desk: string; }

export const READING_THEMES: ReadingTheme[] = [
  { id: "light", name: "Light", paper: "#ffffff", text: "#1f1f1f", desk: "#e9e9ea" },
  { id: "sepia", name: "Sepia", paper: "#f4ecd8", text: "#5b4636", desk: "#e0d4b8" },
  { id: "vintage", name: "Vintage", paper: "#ebdcb9", text: "#3b2c1c", desk: "#c9b88f" },
  { id: "solarized", name: "Solarized", paper: "#fdf6e3", text: "#4a5a61", desk: "#eee8d5" },
  { id: "mist", name: "Mist", paper: "#eceef1", text: "#2b2f36", desk: "#d6d9de" },
  { id: "night", name: "Night", paper: "#1f2023", text: "#d9d9de", desk: "#131416" },
  { id: "dusk", name: "Dusk", paper: "#1c2431", text: "#c9d4e3", desk: "#111720" },
];

export const READING_FONTS: { id: string; name: string; stack: string | null }[] = [
  { id: "document", name: "Document fonts", stack: null },
  { id: "serif", name: "Serif", stack: '"Noto Serif", Georgia, "Liberation Serif", "Times New Roman", serif' },
  { id: "sans", name: "Sans-serif", stack: '"Noto Sans", "Segoe UI", Inter, "Liberation Sans", Arial, sans-serif' },
  { id: "book", name: "Book (Literata, Charter)", stack: 'Literata, "Bitstream Charter", Charter, "Iowan Old Style", "Palatino Linotype", Palatino, "Noto Serif", serif' },
  { id: "mono", name: "Monospace", stack: '"JetBrains Mono", "Cascadia Mono", Consolas, "DejaVu Sans Mono", monospace' },
];

export interface ReadingSettings { theme?: string; custom?: { paper: string; text: string }; font?: string; }

/** The colours to use, or null to follow the app's light/dark theme. */
export function readingColors(s: ReadingSettings | undefined): ReadingTheme | null {
  const id = s?.theme || "auto";
  if (id === "auto") return null;
  if (id === "custom" && s?.custom) {
    const { paper, text } = s.custom;
    return { id: "custom", name: "Custom", paper, text, desk: shade(paper, isDarkColor(paper) ? -0.35 : -0.1) };
  }
  return READING_THEMES.find((t) => t.id === id) || null;
}

export function readingFontStack(s: ReadingSettings | undefined): string | null {
  const id = s?.font || "document";
  const preset = READING_FONTS.find((f) => f.id === id);
  if (preset) return preset.stack;
  // Any other value is a font name the user typed.
  return `"${id.replace(/["\\]/g, "")}", ${READING_FONTS[1].stack}`;
}

export function isDarkColor(hex: string): boolean {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) return false;
  const n = parseInt(m[1], 16);
  const lin = (v: number) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
  const L = 0.2126 * lin((n >> 16) & 255) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255);
  return L < 0.18;
}

function shade(hex: string, f: number): string {
  const n = parseInt(hex.replace("#", ""), 16);
  const ch = (v: number) => Math.max(0, Math.min(255, Math.round(f < 0 ? v * (1 + f) : v + (255 - v) * f)));
  return "#" + [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => ch(v).toString(16).padStart(2, "0")).join("");
}

let styleEl: HTMLStyleElement | null = null;

/**
 * Apply (or clear, with active=false) the reading look. Returns whether the paper is dark, which
 * decides the document colour remapping (the same one dark mode uses).
 */
export function applyReading(s: ReadingSettings | undefined, use: { colors: boolean; font: boolean }, appDark: boolean): boolean {
  if (!styleEl) { styleEl = document.createElement("style"); styleEl.id = "reading-theme"; document.head.appendChild(styleEl); }
  const colors = use.colors ? readingColors(s) : null;
  const font = use.font ? readingFontStack(s) : null;
  const root = document.documentElement;
  root.toggleAttribute("data-reading", !!(colors || font));
  let css = "";
  if (colors) {
    // Printing turns the colours off (use.colors=false) so paper stays white and ink stays black.
    css += `html[data-reading] { --paper: ${colors.paper}; --paper-text: ${colors.text}; --page-bg: ${colors.desk}; }\n`;
  }
  if (font) {
    // Body text only: code (monospace runs and code blocks) and formulas keep their own fonts.
    css += `html[data-reading] .om-editor :is(.om-p, .om-r):not([style*="Consolas" i]):not([style*="Courier" i]):not([style*="mono" i]) { font-family: ${font} !important; }\n`;
  }
  styleEl.textContent = css;
  return colors ? isDarkColor(colors.paper) : appDark;
}
