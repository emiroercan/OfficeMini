import { el, tooltip } from "./widgets";

export interface StatusHandlers {
  onZoom(z: number): void;
  onViewMode(mode: "page" | "continuous"): void;
  onGoToPage(): void;
  onSource(): void;
  onTheme(): void;
}

export interface StatusHandle {
  update(s: { page: number; pages: number; words: number; zoom: number; mode: "page" | "continuous"; dirty: boolean; status?: string; isMd: boolean; source: boolean; dark: boolean }): void;
  flash(msg: string): void;
}

export function buildStatusbar(container: HTMLElement, h: StatusHandlers): StatusHandle {
  container.innerHTML = "";
  const pageEl = el("span", { class: "sb-btn", title: "Go to page (Ctrl+Shift+G)" }, "Page 1 of 1");
  pageEl.addEventListener("click", () => h.onGoToPage());
  const wordsEl = el("span", null, "0 words");
  const msgEl = el("span", { style: { color: "var(--ui-accent)" } }, "");
  const srcBtn = el("span", { class: "sb-btn", hidden: "hidden" }, "Markdown source");
  tooltip(srcBtn, "Switch between rendered view and raw Markdown", "Ctrl+Alt+M");
  srcBtn.addEventListener("click", () => h.onSource());
  const modePage = el("span", { class: "sb-btn active" }, "Page view");
  const modeCont = el("span", { class: "sb-btn" }, "Continuous");
  tooltip(modePage, "Show pages as they will print");
  tooltip(modeCont, "Single continuous column (faster for very long documents)");
  modePage.addEventListener("click", () => h.onViewMode("page"));
  modeCont.addEventListener("click", () => h.onViewMode("continuous"));
  const themeBtn = el("span", { class: "sb-btn" }, "☾");
  tooltip(themeBtn, "Dark mode on/off", "Ctrl+Shift+D");
  themeBtn.addEventListener("click", () => h.onTheme());
  const zoomOut = el("span", { class: "sb-btn" }, "−");
  const zoomIn = el("span", { class: "sb-btn" }, "+");
  const range = el("input", { type: "range", min: "50", max: "300", step: "5", value: "100" });
  const zoomLabel = el("span", { class: "sb-btn", style: { minWidth: "44px", textAlign: "center" }, title: "Reset zoom (Ctrl+0)" }, "100%");
  let zoom = 100;
  range.addEventListener("input", () => h.onZoom(parseInt(range.value, 10) / 100));
  zoomOut.addEventListener("click", () => h.onZoom(Math.max(0.5, (zoom - 10) / 100)));
  zoomIn.addEventListener("click", () => h.onZoom(Math.min(3, (zoom + 10) / 100)));
  zoomLabel.addEventListener("click", () => h.onZoom(1));
  tooltip(zoomOut, "Zoom out", "Ctrl+-");
  tooltip(zoomIn, "Zoom in", "Ctrl++");
  container.append(pageEl, wordsEl, msgEl, el("span", { class: "grow" }), srcBtn, modePage, modeCont, themeBtn, el("span", { class: "zoom" }, zoomOut, range, zoomIn, zoomLabel));
  let flashTimer = 0;
  return {
    update(s) {
      pageEl.textContent = s.source ? "Markdown source" : `Page ${s.page} of ${s.pages}`;
      wordsEl.textContent = `${s.words} word${s.words === 1 ? "" : "s"}`;
      zoom = Math.round(s.zoom * 100);
      range.value = String(zoom);
      zoomLabel.textContent = zoom + "%";
      modePage.classList.toggle("active", s.mode === "page" && !s.source);
      modeCont.classList.toggle("active", s.mode === "continuous" && !s.source);
      srcBtn.hidden = !s.isMd;
      srcBtn.classList.toggle("active", s.source);
      themeBtn.textContent = s.dark ? "☀" : "☾";
      if (s.status !== undefined) msgEl.textContent = s.status;
    },
    flash(msg) {
      msgEl.textContent = msg;
      clearTimeout(flashTimer);
      flashTimer = window.setTimeout(() => { msgEl.textContent = ""; }, 2500);
    },
  };
}
