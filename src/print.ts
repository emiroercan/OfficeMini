// Printing: the page view is already laid out page-by-page, so printing is
// mostly a matter of collapsing the inter-page gap and setting @page size.
import { SectProps } from "./schema";

let styleEl: HTMLStyleElement | null = null;

export function setPrintPageSize(sect: SectProps) {
  if (!styleEl) { styleEl = document.createElement("style"); styleEl.id = "print-page"; document.head.appendChild(styleEl); }
  const mm = (tw: number) => (tw / 1440) * 25.4;
  const w = mm(sect.pgW).toFixed(2), h = mm(sect.pgH).toFixed(2);
  // Everything printed must fit the paper width exactly: anything wider (the titlebar strip,
  // the window-control overlay, the page-number gutter) makes WebKit and Chromium shrink the
  // whole document to fit, and the pages then drift against the paper (0.89x on A4 here).
  styleEl.textContent = `@page { size: ${w}mm ${h}mm; margin: 0; }
@media print {
  html, body, #app, #workspace, #pagearea { width: ${w}mm !important; min-width: 0 !important; max-width: ${w}mm !important; overflow: hidden !important; }
  #doctitle, div[data-tauri-decorum-tb], .decorum-tb-btn { display: none !important; }
}`;
}

/**
 * Print the document. `prepare` must put the editor into page view at 100% and
 * resolve when layout has settled; `restore` undoes that afterwards.
 */
export async function printDocument(prepare: () => Promise<void>, restore: () => void) {
  await prepare();
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  const done = () => { window.removeEventListener("afterprint", done); restore(); };
  window.addEventListener("afterprint", done);
  try { window.print(); } catch { done(); }
  // Some webviews never fire afterprint; restore on a timer as a fallback.
  setTimeout(done, 1500);
}
