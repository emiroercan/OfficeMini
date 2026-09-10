// Modal dialog primitives (shared by the Words and Sheets editors; keeps the
// Sheets bundle free of ProseMirror). Enter = OK, Esc = cancel.
import { el, closeAllPopups } from "./widgets";

export interface DialogButton { label: string; primary?: boolean; action?: () => boolean | void; }

let current: HTMLElement | null = null;
let currentClose: (() => void) | null = null;

export function showDialog(title: string, body: HTMLElement, buttons: DialogButton[], opts: { onClose?: () => void; width?: string } = {}): () => void {
  closeAllPopups();
  closeDialog();
  const overlay = document.getElementById("overlay")!;
  const dlg = el("div", { class: "dialog", role: "dialog", "aria-label": title }, el("h2", null, title), el("div", { class: "body" }, body));
  if (opts.width) dlg.style.width = opts.width;
  const btnRow = el("div", { class: "buttons" });
  const close = () => { if (current !== dlg) return; overlay.classList.remove("show"); overlay.innerHTML = ""; current = null; currentClose = null; document.removeEventListener("keydown", onKey, true); opts.onClose?.(); };
  const onKey = (e: KeyboardEvent) => {
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(); }
    else if (e.key === "Enter" && !(e.target instanceof HTMLTextAreaElement) && !(e.target instanceof HTMLButtonElement)) {
      const primary = buttons.find((b) => b.primary);
      if (primary) { e.preventDefault(); e.stopPropagation(); if (primary.action?.() !== false) close(); }
    }
  };
  for (const b of buttons) {
    const be = el("button", { type: "button", class: b.primary ? "primary" : "" }, b.label);
    be.addEventListener("click", () => { if (b.action?.() !== false) close(); });
    btnRow.appendChild(be);
  }
  dlg.appendChild(btnRow);
  overlay.innerHTML = "";
  overlay.appendChild(dlg);
  overlay.classList.add("show");
  overlay.onmousedown = (e) => { if (e.target === overlay) close(); };
  current = dlg;
  currentClose = close;
  document.addEventListener("keydown", onKey, true);
  setTimeout(() => { const f = dlg.querySelector<HTMLElement>("input, select, textarea, button.primary"); f?.focus(); if (f instanceof HTMLInputElement) f.select(); }, 0);
  return close;
}

export function closeDialog() {
  if (currentClose) { const c = currentClose; currentClose = null; c(); return; }
  if (!current) return;
  const overlay = document.getElementById("overlay")!;
  overlay.classList.remove("show"); overlay.innerHTML = ""; current = null;
}

export function dialogOpen(): boolean { return !!current; }

export function promptDialog(title: string, label: string, value = "", placeholder = ""): Promise<string | null> {
  return new Promise((resolve) => {
    const input = el("input", { type: "text", value, placeholder, style: { width: "100%" } });
    let done = false;
    showDialog(title, el("div", null, el("label", null, label), input), [
      { label: "Cancel" },
      { label: "OK", primary: true, action: () => { done = true; resolve(input.value); } },
    ], { onClose: () => { if (!done) resolve(null); } });
  });
}
