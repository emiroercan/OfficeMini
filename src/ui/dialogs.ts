// Modal dialogs (all keyboard friendly: Enter = OK, Esc = cancel).
import { el, closeAllPopups } from "./widgets";
import { Shortcut, keyLabel, EXTRA_HELP } from "../editor/keymap";
import { SectProps } from "../schema";
import { twipsToPt, ptToTwips } from "../docx/units";

export interface DialogButton { label: string; primary?: boolean; action?: () => boolean | void; }

let current: HTMLElement | null = null;

export function showDialog(title: string, body: HTMLElement, buttons: DialogButton[], opts: { onClose?: () => void; width?: string } = {}): () => void {
  closeAllPopups();
  closeDialog();
  const overlay = document.getElementById("overlay")!;
  const dlg = el("div", { class: "dialog", role: "dialog", "aria-label": title }, el("h2", null, title), el("div", { class: "body" }, body));
  if (opts.width) dlg.style.width = opts.width;
  const btnRow = el("div", { class: "buttons" });
  const close = () => { if (current !== dlg) return; overlay.classList.remove("show"); overlay.innerHTML = ""; current = null; document.removeEventListener("keydown", onKey, true); opts.onClose?.(); };
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
  document.addEventListener("keydown", onKey, true);
  setTimeout(() => { const f = dlg.querySelector<HTMLElement>("input, select, textarea, button.primary"); f?.focus(); if (f instanceof HTMLInputElement) f.select(); }, 0);
  return close;
}

export function closeDialog() {
  if (!current) return;
  const overlay = document.getElementById("overlay")!;
  overlay.classList.remove("show"); overlay.innerHTML = ""; current = null;
}

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

export function linkDialog(href: string, text: string, hasSelection: boolean): Promise<{ href: string; text: string } | null> {
  return new Promise((resolve) => {
    const hrefIn = el("input", { type: "text", value: href, placeholder: "https://…", style: { width: "100%" } });
    const textIn = el("input", { type: "text", value: text, style: { width: "100%" }, disabled: hasSelection ? "disabled" : null });
    let done = false;
    showDialog("Insert link", el("div", null, el("label", null, "Address"), hrefIn, el("label", null, "Text to display"), textIn), [
      { label: "Cancel" },
      { label: "OK", primary: true, action: () => { let h = hrefIn.value.trim(); if (!h) return false; if (!/^(https?:|mailto:|file:|#|ftp:|tel:)/i.test(h) && /\./.test(h)) h = "https://" + h; done = true; resolve({ href: h, text: textIn.value }); } },
    ], { onClose: () => { if (!done) resolve(null); } });
  });
}

export function tableDialog(): Promise<{ rows: number; cols: number } | null> {
  return new Promise((resolve) => {
    const rows = el("input", { type: "number", value: "3", min: "1", max: "200", style: { width: "80px" } });
    const cols = el("input", { type: "number", value: "3", min: "1", max: "40", style: { width: "80px" } });
    let done = false;
    showDialog("Insert table", el("div", { class: "grid2" }, el("div", null, el("label", null, "Columns"), cols), el("div", null, el("label", null, "Rows"), rows)), [
      { label: "Cancel" },
      { label: "Insert", primary: true, action: () => { done = true; resolve({ rows: Math.max(1, parseInt(rows.value, 10) || 1), cols: Math.max(1, Math.min(40, parseInt(cols.value, 10) || 1)) }); } },
    ], { onClose: () => { if (!done) resolve(null); } });
  });
}

const PAPERS: [string, number, number][] = [["A4", 11906, 16838], ["Letter", 12240, 15840], ["Legal", 12240, 20160], ["A5", 8391, 11906], ["A3", 16838, 23811], ["B5", 9979, 14175], ["Executive", 10440, 15120]];

export function pageSetupDialog(sect: SectProps): Promise<SectProps | null> {
  return new Promise((resolve) => {
    const portraitW = Math.min(sect.pgW, sect.pgH), portraitH = Math.max(sect.pgW, sect.pgH);
    const paper = el("select", null, ...PAPERS.map(([n, w, h]) => el("option", { value: n, selected: Math.abs(w - portraitW) < 20 && Math.abs(h - portraitH) < 20 ? "selected" : null }, `${n} (${(w / 567).toFixed(1)} × ${(h / 567).toFixed(1)} cm)`)), el("option", { value: "custom" }, "Custom"));
    const landscape = sect.pgW > sect.pgH;
    const orient = el("select", null, el("option", { value: "portrait", selected: !landscape ? "selected" : null }, "Portrait"), el("option", { value: "landscape", selected: landscape ? "selected" : null }, "Landscape"));
    const num = (v: number) => el("input", { type: "number", step: "0.1", min: "0", value: (twipsToPt(v) / 28.35).toFixed(2), style: { width: "90px" } });
    const mT = num(sect.marT), mB = num(sect.marB), mL = num(sect.marL), mR = num(sect.marR);
    const cw = el("input", { type: "number", step: "0.1", value: (sect.pgW / 567).toFixed(2), style: { width: "90px" } });
    const ch = el("input", { type: "number", step: "0.1", value: (sect.pgH / 567).toFixed(2), style: { width: "90px" } });
    if (!Array.from(paper.options).some((o) => o.selected)) paper.value = "custom";
    const custom = el("div", { class: "grid2", style: { display: paper.value === "custom" ? "grid" : "none" } }, el("div", null, el("label", null, "Width (cm)"), cw), el("div", null, el("label", null, "Height (cm)"), ch));
    paper.addEventListener("change", () => { custom.style.display = paper.value === "custom" ? "grid" : "none"; });
    const body = el("div", null,
      el("div", { class: "grid2" }, el("div", null, el("label", null, "Paper size"), paper), el("div", null, el("label", null, "Orientation"), orient)),
      custom,
      el("label", { style: { marginTop: "10px" } }, "Margins (cm)"),
      el("div", { class: "grid2" }, el("div", null, el("label", null, "Top"), mT), el("div", null, el("label", null, "Bottom"), mB), el("div", null, el("label", null, "Left"), mL), el("div", null, el("label", null, "Right"), mR)),
    );
    let done = false;
    const cm = (inp: HTMLInputElement) => Math.round(Math.max(0, parseFloat(inp.value.replace(",", ".")) || 0) * 567);
    showDialog("Page setup", body, [
      { label: "Cancel" },
      { label: "OK", primary: true, action: () => {
        let w: number, h: number;
        if (paper.value === "custom") { w = cm(cw); h = cm(ch); }
        else { const p = PAPERS.find((x) => x[0] === paper.value)!; w = p[1]; h = p[2]; }
        if (orient.value === "landscape" && w < h) [w, h] = [h, w];
        if (orient.value === "portrait" && w > h) [w, h] = [h, w];
        done = true;
        resolve({ ...sect, pgW: w, pgH: h, orient: orient.value === "landscape" ? "landscape" : null, marT: cm(mT), marB: cm(mB), marL: cm(mL), marR: cm(mR) });
      } },
    ], { onClose: () => { if (!done) resolve(null); } });
  });
}

export function paragraphDialog(cur: { before: number; after: number; line: number; rule: string; left: number; right: number; firstLine: number; hanging: number }): Promise<Record<string, any> | null> {
  return new Promise((resolve) => {
    const n = (v: number, step = "1") => el("input", { type: "number", step, value: String(v), style: { width: "90px" } });
    const before = n(cur.before), after = n(cur.after);
    const lineSel = el("select", null, ...[["1", "Single"], ["1.15", "1.15"], ["1.5", "1.5 lines"], ["2", "Double"], ["exact", "Exactly"], ["atLeast", "At least"]].map(([v, l]) => el("option", { value: v }, l)));
    const lineVal = n(cur.rule === "auto" ? Math.round((cur.line / 240) * 100) / 100 : cur.line / 20, "0.5");
    if (cur.rule === "auto") { const m = (cur.line / 240).toFixed(2); lineSel.value = ["1.00", "1.15", "1.50", "2.00"].includes(m) ? String(parseFloat(m)) : "1"; }
    else lineSel.value = cur.rule;
    const left = n(cur.left), right = n(cur.right);
    const special = el("select", null, el("option", { value: "none" }, "(none)"), el("option", { value: "firstLine" }, "First line"), el("option", { value: "hanging" }, "Hanging"));
    special.value = cur.hanging ? "hanging" : cur.firstLine ? "firstLine" : "none";
    const specialBy = n(cur.hanging || cur.firstLine || 0);
    const body = el("div", null,
      el("label", null, "Indentation (pt)"),
      el("div", { class: "grid2" }, el("div", null, el("label", null, "Left"), left), el("div", null, el("label", null, "Right"), right), el("div", null, el("label", null, "Special"), special), el("div", null, el("label", null, "By"), specialBy)),
      el("label", { style: { marginTop: "10px" } }, "Spacing (pt)"),
      el("div", { class: "grid2" }, el("div", null, el("label", null, "Before"), before), el("div", null, el("label", null, "After"), after), el("div", null, el("label", null, "Line spacing"), lineSel), el("div", null, el("label", null, "At"), lineVal)),
    );
    let done = false;
    showDialog("Paragraph", body, [
      { label: "Cancel" },
      { label: "OK", primary: true, action: () => {
        const patch: Record<string, any> = {
          spBefore: ptToTwips(parseFloat(before.value) || 0), spAfter: ptToTwips(parseFloat(after.value) || 0), spBeforeAuto: false, spAfterAuto: false,
          indLeft: ptToTwips(parseFloat(left.value) || 0), indRight: ptToTwips(parseFloat(right.value) || 0),
        };
        const by = ptToTwips(parseFloat(specialBy.value) || 0);
        if (special.value === "hanging") { patch.indHanging = by; patch.indFirstLine = null; }
        else if (special.value === "firstLine") { patch.indFirstLine = by; patch.indHanging = null; }
        else { patch.indHanging = null; patch.indFirstLine = null; }
        const lv = lineSel.value;
        if (lv === "exact" || lv === "atLeast") { patch.spLine = ptToTwips(parseFloat(lineVal.value) || 12); patch.spLineRule = lv; }
        else { patch.spLine = Math.round(parseFloat(lv) * 240); patch.spLineRule = "auto"; }
        done = true; resolve(patch);
      } },
    ], { onClose: () => { if (!done) resolve(null); } });
  });
}

export function goToPageDialog(max: number): Promise<number | null> {
  return new Promise((resolve) => {
    const input = el("input", { type: "number", min: "1", max: String(max), value: "1", style: { width: "100px" } });
    let done = false;
    showDialog("Go to page", el("div", null, el("label", null, `Page number (1–${max})`), input), [
      { label: "Cancel" },
      { label: "Go", primary: true, action: () => { done = true; resolve(Math.max(1, Math.min(max, parseInt(input.value, 10) || 1))); } },
    ], { onClose: () => { if (!done) resolve(null); } });
  });
}

export function shortcutsDialog(list: Shortcut[]) {
  const groups = new Map<string, [string, string][]>();
  for (const s of list) {
    if (s.hidden) continue;
    const rows = groups.get(s.group) || [];
    rows.push([s.keys.map(keyLabel).filter((v, i, a) => a.indexOf(v) === i).join("  or  "), s.label]);
    groups.set(s.group, rows);
  }
  for (const g of EXTRA_HELP) groups.set(g.group, [...(groups.get(g.group) || []), ...g.rows]);
  const table = el("table", { class: "keys" });
  for (const [g, rows] of groups) {
    table.appendChild(el("tr", null, el("th", { colspan: "2" }, g)));
    for (const [k, l] of rows) table.appendChild(el("tr", null, el("td", null, k), el("td", null, l)));
  }
  const body = el("div", { style: { maxHeight: "70vh", overflow: "auto" } }, table);
  showDialog("Keyboard shortcuts and mouse", body, [{ label: "Close", primary: true }], { width: "560px" });
}

export function aboutDialog() {
  showDialog("About OfficeMini", el("div", null,
    el("p", null, "OfficeMini 0.1 — a small, fast editor for Word (.docx) and Markdown files."),
    el("p", null, "Files are saved with a round-trip strategy: everything the editor does not understand is preserved exactly as it was."),
    el("p", { style: { color: "var(--ui-muted)" } }, "Press Ctrl+/ for keyboard shortcuts."),
  ), [{ label: "Close", primary: true }]);
}
