// Small UI primitives: element builder, tooltips, popups, menus, icons.

export function el<K extends keyof HTMLElementTagNameMap>(tag: K, attrs?: Record<string, any> | null, ...children: (Node | string | null | undefined)[]): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (attrs) for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === "class") e.className = v;
    else if (k === "style" && typeof v === "object") Object.assign(e.style, v);
    else if (k.startsWith("on") && typeof v === "function") e.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === "text") e.textContent = v;
    else e.setAttribute(k, String(v));
  }
  for (const c of children) if (c !== null && c !== undefined) e.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
  return e;
}

// ---------------------------------------------------------------------------
// Tooltip

let tipEl: HTMLElement | null = null;
let tipTimer = 0;
export function tooltip(target: HTMLElement, text: string, key?: string) {
  target.setAttribute("data-tip", text);
  if (key) target.setAttribute("data-tipkey", key);
  target.setAttribute("aria-label", text);
  target.addEventListener("mouseenter", () => {
    clearTimeout(tipTimer);
    tipTimer = window.setTimeout(() => showTip(target), 450);
  });
  target.addEventListener("mouseleave", hideTip);
  target.addEventListener("mousedown", hideTip);
}
function showTip(target: HTMLElement) {
  if (!document.body.contains(target)) return;
  hideTip();
  const text = target.getAttribute("data-tip") || "";
  const key = target.getAttribute("data-tipkey");
  tipEl = el("div", { id: "tooltip" }, text, key ? el("span", { class: "tt-key" }, key) : null);
  document.body.appendChild(tipEl);
  const r = target.getBoundingClientRect();
  const tr = tipEl.getBoundingClientRect();
  let x = r.left + r.width / 2 - tr.width / 2;
  x = Math.max(4, Math.min(window.innerWidth - tr.width - 4, x));
  let y = r.bottom + 6;
  if (y + tr.height > window.innerHeight) y = r.top - tr.height - 6;
  tipEl.style.left = x + "px";
  tipEl.style.top = y + "px";
}
export function hideTip() { clearTimeout(tipTimer); if (tipEl) { tipEl.remove(); tipEl = null; } }

// ---------------------------------------------------------------------------
// Popups

export interface PopupHandle { close(): void; el: HTMLElement; }
let openPopups: PopupHandle[] = [];

export function closeAllPopups() { for (const p of [...openPopups]) p.close(); }

export interface PopupOptions { className?: string; align?: "left" | "right"; below?: boolean; onClose?: () => void; keepOthers?: boolean; }

export function showPopup(anchor: HTMLElement | { x: number; y: number }, build: (popup: HTMLElement, close: () => void) => void, opts: PopupOptions = {}): PopupHandle {
  if (!opts.keepOthers) closeAllPopups();
  hideTip();
  const popup = el("div", { class: opts.className || "popup" });
  const handle: PopupHandle = { el: popup, close() {
    if (!popup.parentNode) return;
    popup.remove();
    openPopups = openPopups.filter((p) => p !== handle);
    document.removeEventListener("mousedown", onDown, true);
    document.removeEventListener("keydown", onKey, true);
    window.removeEventListener("blur", handle.close);
    opts.onClose?.();
  } };
  const onDown = (e: MouseEvent) => {
    if (popup.contains(e.target as Node)) return;
    if (anchor instanceof HTMLElement && anchor.contains(e.target as Node)) return;
    // Click on another popup keeps this one only if it is a child popup.
    if (openPopups.some((p) => p !== handle && p.el.contains(e.target as Node) && openPopups.indexOf(p) > openPopups.indexOf(handle))) return;
    handle.close();
  };
  const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); handle.close(); } };
  build(popup, handle.close);
  document.body.appendChild(popup);
  // position
  const pr = popup.getBoundingClientRect();
  let x: number, y: number;
  if (anchor instanceof HTMLElement) {
    const r = anchor.getBoundingClientRect();
    x = opts.align === "right" ? r.right - pr.width : r.left;
    y = opts.below === false ? r.top - pr.height - 2 : r.bottom + 2;
  } else { x = anchor.x; y = anchor.y; }
  x = Math.max(2, Math.min(window.innerWidth - pr.width - 2, x));
  if (y + pr.height > window.innerHeight - 2) y = Math.max(2, (anchor instanceof HTMLElement ? anchor.getBoundingClientRect().top : anchor.y) - pr.height - 2);
  if (y < 2) y = 2;
  popup.style.left = x + "px";
  popup.style.top = y + "px";
  setTimeout(() => {
    document.addEventListener("mousedown", onDown, true);
    document.addEventListener("keydown", onKey, true);
    window.addEventListener("blur", handle.close);
  }, 0);
  openPopups.push(handle);
  return handle;
}

// ---------------------------------------------------------------------------
// Menus

export interface MenuItem {
  label?: string;
  key?: string;          // shortcut label
  action?: () => void;
  disabled?: boolean;
  checked?: boolean;
  sep?: boolean;
  submenu?: MenuItem[];
  icon?: string;
}

export function showMenu(anchor: HTMLElement | { x: number; y: number }, items: MenuItem[], opts: PopupOptions = {}): PopupHandle {
  return showPopup(anchor, (popup, close) => {
    popup.className = "menu-popup";
    let sub: PopupHandle | null = null;
    for (const it of items) {
      if (it.sep) { popup.appendChild(el("div", { class: "menu-sep" })); continue; }
      const row = el("div", { class: "menu-item" + (it.disabled ? " disabled" : "") + (it.submenu ? " has-sub" : "") },
        el("span", { class: "mi-check" }, it.checked ? "✓" : ""),
        el("span", { class: "mi-label" }, it.label || ""),
        it.key ? el("span", { class: "mi-key" }, it.key) : null);
      if (it.submenu) {
        row.addEventListener("mouseenter", () => {
          sub?.close();
          sub = showMenu({ x: row.getBoundingClientRect().right - 2, y: row.getBoundingClientRect().top - 4 }, it.submenu!, { keepOthers: true });
        });
      } else {
        row.addEventListener("mouseenter", () => { sub?.close(); sub = null; });
        row.addEventListener("click", (e) => { e.stopPropagation(); if (it.disabled) return; closeAllPopups(); it.action?.(); });
      }
      popup.appendChild(row);
    }
    // keyboard navigation
    let idx = -1;
    const rows = () => Array.from(popup.querySelectorAll<HTMLElement>(".menu-item:not(.disabled)"));
    popup.tabIndex = -1;
    popup.addEventListener("keydown", (e) => {
      const rs = rows();
      if (e.key === "ArrowDown") { idx = (idx + 1) % rs.length; rs.forEach((r, i) => r.classList.toggle("sel", i === idx)); e.preventDefault(); }
      else if (e.key === "ArrowUp") { idx = (idx - 1 + rs.length) % rs.length; rs.forEach((r, i) => r.classList.toggle("sel", i === idx)); e.preventDefault(); }
      else if (e.key === "Enter" && idx >= 0) { rs[idx].click(); e.preventDefault(); }
    });
    setTimeout(() => popup.focus(), 0);
    void close;
  }, opts);
}

// ---------------------------------------------------------------------------
// Icons (16x16, currentColor)

const ICONS: Record<string, string> = {
  new: '<path d="M4 1h6l4 4v10H4z" fill="none" stroke="currentColor"/><path d="M10 1v4h4" fill="none" stroke="currentColor"/>',
  open: '<path d="M1.5 4.5h5l1.5 1.5h6.5v8h-13z" fill="none" stroke="currentColor"/>',
  save: '<path d="M2 2h10l2 2v10H2z" fill="none" stroke="currentColor"/><rect x="4.5" y="2.5" width="6" height="4" fill="none" stroke="currentColor"/><rect x="4.5" y="9.5" width="7" height="4" fill="none" stroke="currentColor"/>',
  print: '<path d="M4 6V2h8v4" fill="none" stroke="currentColor"/><rect x="1.5" y="6.5" width="13" height="6" rx="1" fill="none" stroke="currentColor"/><rect x="4.5" y="10.5" width="7" height="4" fill="#fff" stroke="currentColor"/>',
  undo: '<path d="M6 4L2 7.5 6 11" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M2.5 7.5H10a3.5 3.5 0 0 1 0 7H7" fill="none" stroke="currentColor" stroke-width="1.5"/>',
  redo: '<path d="M10 4l4 3.5-4 3.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M13.5 7.5H6a3.5 3.5 0 0 0 0 7h3" fill="none" stroke="currentColor" stroke-width="1.5"/>',
  find: '<circle cx="6.5" cy="6.5" r="4.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M10 10l4.5 4.5" stroke="currentColor" stroke-width="1.8"/>',
  bold: '<path d="M4 2h5a3 3 0 0 1 1.5 5.6A3.3 3.3 0 0 1 9.3 14H4z M6 4v3.5h3a1.75 1.75 0 0 0 0-3.5z M6 9.5V12h3.3a1.25 1.25 0 0 0 0-2.5z"/>',
  italic: '<path d="M6.5 2h6v1.8h-2.2L8 12.2h2.3V14h-6v-1.8h2.2L8.8 3.8H6.5z"/>',
  underline: '<path d="M4 2v6a4 4 0 0 0 8 0V2h-1.8v6a2.2 2.2 0 0 1-4.4 0V2z"/><rect x="3" y="13.2" width="10" height="1.4"/>',
  strike: '<path d="M8 2.5c-2.6 0-4 1.3-4 3 0 .8.3 1.4.8 1.9h2.9C6.5 7 6 6.4 6 5.6c0-.9.8-1.5 2-1.5 1 0 1.7.4 2.1 1.1l1.7-.8C11.2 3.2 9.8 2.5 8 2.5zM8.3 10.5H5.4c.1 1.9 1.5 3 3.6 3 2.7 0 4.1-1.4 4.1-3.1 0-.4-.1-.7-.2-1H9.5c.3.2.4.5.4.8 0 .8-.9 1.4-2 1.4-1 0-1.8-.4-2-1.1z"/><rect x="2" y="7.6" width="12" height="1.2"/>',
  sup: '<path d="M2 4h2.2l2 3 2-3h2.2L7.4 8.2 10.6 13H8.4L6.2 9.6 4 13H1.8l3.2-4.8z"/><path d="M11 2h3.5v1h-2.3v1h2.3v1H11v-1h2.3V3H11z"/>',
  sub: '<path d="M2 3h2.2l2 3 2-3h2.2L7.4 7.2 10.6 12H8.4L6.2 8.6 4 12H1.8l3.2-4.8z"/><path d="M11 11h3.5v1h-2.3v1h2.3v1H11v-1h2.3v-1H11z"/>',
  alignLeft: '<path d="M2 3h12v1.5H2zM2 6.5h8V8H2zM2 10h12v1.5H2zM2 13.5h8V15H2z"/>',
  alignCenter: '<path d="M2 3h12v1.5H2zM4 6.5h8V8H4zM2 10h12v1.5H2zM4 13.5h8V15H4z"/>',
  alignRight: '<path d="M2 3h12v1.5H2zM6 6.5h8V8H6zM2 10h12v1.5H2zM6 13.5h8V15H6z"/>',
  alignJustify: '<path d="M2 3h12v1.5H2zM2 6.5h12V8H2zM2 10h12v1.5H2zM2 13.5h12V15H2z"/>',
  bullets: '<circle cx="3" cy="4" r="1.4"/><circle cx="3" cy="8" r="1.4"/><circle cx="3" cy="12" r="1.4"/><path d="M6 3.2h8v1.5H6zM6 7.2h8v1.5H6zM6 11.2h8v1.5H6z"/>',
  numbers: '<path d="M1.6 2.2h1.3v3H1.6zM1.4 6.5h2.2v.8l-1.3 1.3h1.3v.8H1.4v-.8l1.3-1.3H1.4zM1.4 11h2.2v3.2H1.4v-.8h1.4v-.5H1.6v-.7h1.2v-.5H1.4z"/><path d="M5.5 3.2h9v1.5h-9zM5.5 7.2h9v1.5h-9zM5.5 11.2h9v1.5h-9z"/>',
  indent: '<path d="M2 3h12v1.5H2zM7 6.5h7V8H7zM7 10h7v1.5H7zM2 13.5h12V15H2z"/><path d="M2 6.5l3 2.2-3 2.2z"/>',
  outdent: '<path d="M2 3h12v1.5H2zM7 6.5h7V8H7zM7 10h7v1.5H7zM2 13.5h12V15H2z"/><path d="M5 6.5L2 8.7l3 2.2z"/>',
  image: '<rect x="1.5" y="2.5" width="13" height="11" rx="1" fill="none" stroke="currentColor"/><circle cx="5" cy="6" r="1.3"/><path d="M2 13l4-4 3 3 2-2 3 3z"/>',
  table: '<rect x="1.5" y="2.5" width="13" height="11" fill="none" stroke="currentColor"/><path d="M1.5 6.5h13M1.5 10h13M6 2.5v11M10.5 2.5v11" stroke="currentColor"/>',
  link: '<path d="M6.5 9.5a2.5 2.5 0 0 0 3.5 0l2-2a2.5 2.5 0 0 0-3.5-3.5l-1 1" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M9.5 6.5a2.5 2.5 0 0 0-3.5 0l-2 2a2.5 2.5 0 0 0 3.5 3.5l1-1" fill="none" stroke="currentColor" stroke-width="1.5"/>',
  color: '<path d="M8 2l4 10h-1.8l-1-2.6H6.8l-1 2.6H4z M7.3 8h1.4L8 4.4z"/><rect x="2" y="13" width="12" height="2" class="swatch-rect"/>',
  highlight: '<path d="M3 11l7.5-7.5 2 2L5 13H3z" fill="none" stroke="currentColor"/><path d="M9 5l2 2" stroke="currentColor"/><rect x="2" y="13" width="12" height="2" class="swatch-rect"/>',
  clear: '<path d="M4 12L10 3h2l1 1-6 9z" fill="none" stroke="currentColor"/><path d="M2 14h12" stroke="currentColor"/><path d="M7.5 5.5l3 3" stroke="currentColor"/>',
  pagebreak: '<path d="M3 2h10v4H3zM3 10h10v4H3z" fill="none" stroke="currentColor"/><path d="M1 8h2M5 8h2M9 8h2M13 8h2" stroke="currentColor"/>',
  spacing: '<path d="M6 3h8v1.5H6zM6 7.2h8v1.5H6zM6 11.5h8V13H6z"/><path d="M3 2v12" stroke="currentColor"/><path d="M1.5 4L3 2l1.5 2M1.5 12L3 14l1.5-2" fill="none" stroke="currentColor"/>',
  zoomIn: '<circle cx="6.5" cy="6.5" r="4.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M10 10l4.5 4.5" stroke="currentColor" stroke-width="1.8"/><path d="M4.5 6.5h4M6.5 4.5v4" stroke="currentColor"/>',
  marks: '<path d="M9 2v12h-1.5V2zM12 2v12h-1.5V2zM8.5 2H6a3 3 0 0 0 0 6h2.5z"/>',
};

export function icon(name: string): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("width", "16");
  svg.setAttribute("height", "16");
  svg.innerHTML = ICONS[name] || "";
  return svg;
}
