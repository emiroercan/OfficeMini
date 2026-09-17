// ============================================================================
// macOS ONLY - native menu bar
// ----------------------------------------------------------------------------
// On macOS a document app is expected to put its menus in the system menu bar at
// the top of the screen, not inside the window. This module mirrors the app's
// existing in-window menu definitions (the same `menus` array the in-window bar is
// built from) into the native macOS menu bar, so there is a single source of truth.
//
// How it works:
//   1. serialize() walks the current MenuItem[] for each top-level menu and turns
//      it into a plain JSON model (labels, accelerators, checked/enabled, submenus),
//      registering each clickable item's action under a STABLE, path-based id.
//   2. F.setAppMenu(model) hands that model to the Rust side (src-tauri/src/macos_menu.rs),
//      which builds the real NSMenu and sets it. Rust wraps the model with the native
//      App and Window menus and turns `nativeRole` items into predefined items.
//   3. When the user clicks a native item, Rust emits a "menu-action" event with the
//      item's id; we look the id up in the registry and run the action. "quit" is
//      special (App menu) and routes to onQuit so the unsaved-changes prompt runs.
//   4. refreshNativeMenu() rebuilds and re-pushes the model (debounced) whenever the
//      app's state changes, so checkmarks, enabled-state and Open Recent stay live.
//
// On every non-macOS platform (and in the browser) every export here is a no-op, so
// the in-window menu bar keeps working unchanged. Keyboard shortcuts continue to be
// handled by the app's own key handlers; the native accelerators only display them
// and act as a fallback, so nothing fires twice.
// ============================================================================

import { MenuItem } from "./ui/widgets";
import * as F from "./files";

export interface MenuSpec { title: string; items: () => MenuItem[]; }

interface Node {
  id?: string;
  label?: string;
  accel?: string;
  checked?: boolean;
  enabled?: boolean;
  role?: string;   // "cut" | "copy" | "paste" | "selectAll"
  sep?: boolean;
  submenu?: Node[];
}

const isMac = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform);
/** True only when the native menu bar should be used (macOS desktop build). */
export const nativeMenuActive = isMac && F.isTauri;

let getMenus: (() => MenuSpec[]) | null = null;
let actions = new Map<string, () => void>();
let debounce = 0;

/**
 * Convert a shortcut label ("⌘+Shift+S", "F9", "Delete") into a Tauri accelerator string, or ""
 * to leave the item without one. Only well-known keys are emitted; punctuation shortcuts (+, -, /,
 * ;, \) are skipped so a stray label can never produce an accelerator Tauri fails to parse - those
 * shortcuts still work through the app's own key handler, they just are not shown in the menu.
 */
function toAccel(key?: string): string {
  if (!key) return "";
  const mods: string[] = [];
  let main = "";
  for (const raw of key.split("+")) {
    const p = raw.trim();
    if (p === "⌘" || /^cmd$/i.test(p) || /^ctrl$/i.test(p)) mods.push("CmdOrCtrl");
    else if (p === "⌥" || /^alt$/i.test(p) || /^option$/i.test(p)) mods.push("Alt");
    else if (p === "⇧" || /^shift$/i.test(p)) mods.push("Shift");
    else if (p === "⌃") mods.push("Ctrl");
    else if (p) main = p;
  }
  let code = "";
  if (/^[A-Za-z]$/.test(main)) code = main.toUpperCase();
  else if (/^[0-9]$/.test(main)) code = main;
  else if (/^F([1-9]|1[0-9]|2[0-4])$/i.test(main)) code = main.toUpperCase();
  else if (["Delete", "Backspace", "Enter", "Tab", "Escape", "Space", "Home", "End", "PageUp", "PageDown"].includes(main)) code = main;
  else return "";   // unknown / punctuation key: skip (still works via the key handler)
  if (!code) return "";
  return [...mods, code].join("+");
}

function serialize(items: MenuItem[], path: string): Node[] {
  const out: Node[] = [];
  items.forEach((it, i) => {
    if (it.sep) { out.push({ sep: true }); return; }
    const id = path + "-" + i;
    const node: Node = { id, label: it.label || "", enabled: it.disabled !== true };
    if (it.checked !== undefined) node.checked = !!it.checked;
    if (it.nativeRole) node.role = it.nativeRole;
    const accel = toAccel(it.key);
    if (accel) node.accel = accel;
    if (it.submenu && it.submenu.length) node.submenu = serialize(it.submenu, id);
    else if (it.action) actions.set(id, it.action);
    out.push(node);
  });
  return out;
}

/** Build the JSON model from the current menu definitions, refreshing the action registry. */
function build(): { title: string; items: Node[] }[] {
  if (!getMenus) return [];
  actions = new Map();
  return getMenus().map((m, i) => ({ title: m.title, items: serialize(m.items(), String(i)) }));
}

/** Rebuild and re-push the native menu (debounced), so checkmarks / enabled-state stay current. */
export function refreshNativeMenu(): void {
  if (!nativeMenuActive || !getMenus) return;
  if (debounce) return;
  debounce = window.setTimeout(() => { debounce = 0; void F.setAppMenu(build()); }, 150);
}

/**
 * Turn the in-window menu bar into the native macOS menu bar. `getMenus` returns the current
 * top-level menus (the same array the in-window bar uses); `onQuit` runs when the App menu's Quit
 * is chosen (so the unsaved-changes flow stays in charge). No-op off macOS.
 */
export function installNativeMenu(provider: () => MenuSpec[], onQuit: () => void): void {
  if (!nativeMenuActive) return;
  getMenus = provider;
  document.documentElement.classList.add("native-menu");   // CSS drops the in-window menu row
  // No titlebar strip on macOS: the toolbar is the top row and doubles as the drag region (empty
  // areas move the window; the buttons, being children without the attribute, still click).
  document.getElementById("toolbar")?.setAttribute("data-tauri-drag-region", "");
  void F.onBackendEvent<string>("menu-action", (id) => {
    if (id === "quit") { onQuit(); return; }
    actions.get(id)?.();
  });
  void F.setAppMenu(build());
}
