// Printing: the page view is already laid out page-by-page, so printing is
// mostly a matter of collapsing the inter-page gap and setting @page size.
import { SectProps } from "./schema";

let styleEl: HTMLStyleElement | null = null;

export function setPrintPageSize(sect: SectProps) {
  if (!styleEl) { styleEl = document.createElement("style"); styleEl.id = "print-page"; document.head.appendChild(styleEl); }
  const mm = (tw: number) => (tw / 1440) * 25.4;
  styleEl.textContent = `@page { size: ${mm(sect.pgW).toFixed(2)}mm ${mm(sect.pgH).toFixed(2)}mm; margin: 0; }`;
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
