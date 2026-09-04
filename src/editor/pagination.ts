// Page view: measures the rendered flow and inserts spacer decorations so
// content jumps from one page box to the next. Print uses the same breaks
// (the inter-page gap collapses to 0 in print CSS), so screen == paper.
import { Plugin, PluginKey, EditorState, Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import { DOMSerializer, Node as PMNode } from "prosemirror-model";
import { schema, SectProps } from "../schema";
import { twipsToPx } from "../docx/units";
import type { HeaderFooter } from "../docx/parse";

export const paginationKey = new PluginKey<PagState>("pagination");

export interface Break { pos: number; gap: number; inline: boolean; page: number; }
export interface PagState { breaks: Break[]; pageCount: number; mode: "page" | "continuous"; }

export interface PageGeom {
  pageW: number; pageH: number; marT: number; marR: number; marB: number; marL: number;
  headerDist: number; footerDist: number; contentW: number; contentH: number; effMarT: number; effMarB: number;
}

let zoomFactor = 1;
export function setZoomFactor(z: number) { zoomFactor = z; }
export const PAGE_GAP = 18;

let hfProvider: { headers: Map<string, HeaderFooter>; footers: Map<string, HeaderFooter> } | null = null;
export function setHeaderFooters(h: { headers: Map<string, HeaderFooter>; footers: Map<string, HeaderFooter> } | null) { hfProvider = h; }

export function geometry(sect: SectProps, effMarT?: number, effMarB?: number): PageGeom {
  const pageW = twipsToPx(sect.pgW), pageH = twipsToPx(sect.pgH);
  const marT = twipsToPx(sect.marT), marB = twipsToPx(sect.marB), marL = twipsToPx(sect.marL + sect.gutter), marR = twipsToPx(sect.marR);
  const t = effMarT ?? marT, b = effMarB ?? marB;
  return {
    pageW, pageH, marT, marR, marB, marL,
    headerDist: twipsToPx(sect.header), footerDist: twipsToPx(sect.footer),
    contentW: pageW - marL - marR, contentH: pageH - t - b, effMarT: t, effMarB: b,
  };
}

function spacer(gap: number, inline: boolean): HTMLElement {
  const el = document.createElement(inline ? "span" : "div");
  el.className = "om-pagegap" + (inline ? " inline" : "");
  el.style.setProperty("--gap", gap + "px");
  if (inline) { el.style.display = "inline-block"; el.style.width = "100%"; el.style.verticalAlign = "top"; }
  el.contentEditable = "false";
  return el;
}

function decorationsFor(doc: PMNode, breaks: Break[]): DecorationSet {
  return DecorationSet.create(doc, breaks.map((b, i) =>
    Decoration.widget(b.pos, () => spacer(b.gap, b.inline), { side: -1, key: "pg" + i + ":" + b.pos + ":" + Math.round(b.gap) + (b.inline ? "i" : "b"), ignoreSelection: true })));
}

// ---------------------------------------------------------------------------
// Measurement

/** A line box: `top/bottom` are raw (unpaginated) layout px; `vtop/vbottom` are visual px relative to the content origin. */
interface LineBox { top: number; bottom: number; vtop: number; vbottom: number; }

export interface MeasureResult { breaks: Break[]; pageCount: number; }

/** Compute page breaks for the current DOM. Must be called after layout. */
export function measure(view: EditorView, geom: PageGeom): MeasureResult {
  const dom = view.dom as HTMLElement;
  const rect = dom.getBoundingClientRect();
  const padTop = parseFloat(getComputedStyle(dom).paddingTop) || 0;
  const originTop = rect.top + padTop * zoomFactor; // content box top (visual client px)
  const contentH = geom.contentH;
  const gapExtra = geom.effMarB + geom.effMarT; // page gap itself is added via CSS var

  // Existing spacers add height; subtract them to get raw (unpaginated) coordinates.
  const spacers = Array.from(dom.querySelectorAll<HTMLElement>(".om-pagegap")).map((s) => {
    const r = s.getBoundingClientRect();
    return { top: (r.top - originTop) / zoomFactor, h: r.height / zoomFactor };
  });
  const rawY = (vy: number) => { let d = 0; for (const s of spacers) if (s.top + s.h <= vy + 0.5) d += s.h; return vy - d; };
  const rel = (clientY: number) => (clientY - originTop) / zoomFactor;

  const breaks: Break[] = [];
  let pageStart = 0; // raw y where the current page's content starts
  let page = 1;
  let forcedNext = false;

  const breakAt = (pos: number, rawTop: number, inline: boolean, marginTop = 0) => {
    const boundary = pageStart + contentH;
    const unused = Math.max(0, boundary - rawTop);
    breaks.push({ pos, gap: Math.max(0, unused + gapExtra - marginTop), inline, page });
    pageStart = rawTop;
    page++;
  };

  const lineBoxes = (el: HTMLElement): LineBox[] => {
    // Collect rects of text and inline objects, skipping absolutely positioned anchors.
    const rects: DOMRect[] = [];
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
      acceptNode: (node) => {
        if (node.nodeType === 1) {
          const e = node as HTMLElement;
          if (e.classList.contains("om-abs")) return NodeFilter.FILTER_REJECT;
          if (e.tagName === "IMG" || e.classList.contains("om-shape") || e.classList.contains("om-textbox") || e.classList.contains("om-pagebreak")) return NodeFilter.FILTER_ACCEPT;
          return NodeFilter.FILTER_SKIP;
        }
        return NodeFilter.FILTER_ACCEPT;
      },
    });
    const range = document.createRange();
    while (walker.nextNode()) {
      const n = walker.currentNode;
      if (n.nodeType === 1) { const r = (n as HTMLElement).getBoundingClientRect(); if (r.height > 0) rects.push(r); }
      else { range.selectNodeContents(n); for (const r of Array.from(range.getClientRects())) if (r.height > 0) rects.push(r); }
    }
    const lines: LineBox[] = [];
    const arr = rects.map((r) => ({ vtop: rel(r.top), vbottom: rel(r.bottom) })).sort((a, b) => a.vtop - b.vtop);
    for (const r of arr) {
      const last = lines[lines.length - 1];
      if (last && r.vtop < last.vbottom - 1) { last.vbottom = Math.max(last.vbottom, r.vbottom); last.vtop = Math.min(last.vtop, r.vtop); }
      else lines.push({ vtop: r.vtop, vbottom: r.vbottom, top: 0, bottom: 0 });
    }
    if (!lines.length) {
      const r = el.getBoundingClientRect();
      lines.push({ vtop: rel(r.top), vbottom: rel(r.bottom), top: 0, bottom: 0 });
    }
    for (const l of lines) { l.top = rawY(l.vtop); l.bottom = rawY(l.vbottom); }
    return lines;
  };

  /** Document position at the start of a line inside paragraph `el` (visual line box). */
  const posOfLine = (el: HTMLElement, line: LineBox, pos: number, node: PMNode): number | null => {
    const r = el.getBoundingClientRect();
    const padL = parseFloat(getComputedStyle(el).paddingLeft) || 0;
    const y = originTop + ((line.vtop + line.vbottom) / 2) * zoomFactor;
    const x = r.left + (padL + 1) * zoomFactor;
    const p = view.posAtCoords({ left: x, top: y });
    if (!p) return null;
    const start = pos + 1, end = pos + 1 + node.content.size;
    if (p.pos <= start || p.pos > end) return null;
    return p.pos;
  };

  const paginateLines = (pos: number, node: PMNode, el: HTMLElement, lines: LineBox[], marginTop: number) => {
    if (!lines.length) return;
    const keepLines = el.getAttribute("data-keeplines") === "1";
    let i = 0;
    let guard = 0;
    while (i < lines.length && guard++ < 200) {
      const boundary = pageStart + contentH;
      let k = i;
      while (k < lines.length && lines[k].bottom <= boundary + 0.5) k++;
      if (k >= lines.length) return; // everything fits
      const stay = k - i, remain = lines.length - k;
      let breakLine = k;
      if (keepLines) breakLine = i;
      else if (stay < 2) breakLine = i;                       // orphan: move all
      else if (remain < 2 && stay >= 3) breakLine = k - 1;    // widow: move one more line
      else if (remain < 2) breakLine = i;
      if (breakLine === i) {
        // Move the remainder (from line i) to the next page.
        if (lines[i].top > pageStart + 1) {
          if (i === 0) breakAt(pos, lines[0].top, false, marginTop);
          else { const bp = posOfLine(el, lines[i], pos, node); if (bp === null) return; breakAt(bp, lines[i].top, true); }
        } else if (i === 0 && stay === 0) {
          // Already at the top and the first line does not fit (taller than a page): overflow.
          pageStart = lines[0].top;
        }
        // Now at the top of a page: split without the orphan rule (we cannot move again).
        const b2 = pageStart + contentH;
        let k2 = i;
        while (k2 < lines.length && lines[k2].bottom <= b2 + 0.5) k2++;
        if (k2 >= lines.length) return;
        if (k2 === i) { // single line taller than a page
          if (i + 1 >= lines.length) return;
          const bp = posOfLine(el, lines[i + 1], pos, node); if (bp === null) return;
          breakAt(bp, lines[i + 1].top, true); i = i + 1; continue;
        }
        let bl = k2;
        if (!keepLines && lines.length - k2 < 2 && k2 - i >= 3) bl = k2 - 1;
        const bp = posOfLine(el, lines[bl], pos, node);
        if (bp === null) return;
        breakAt(bp, lines[bl].top, true);
        i = bl;
        continue;
      }
      const bp = posOfLine(el, lines[breakLine], pos, node);
      if (bp === null) { if (lines[0].top > pageStart + 1) breakAt(pos, lines[0].top, false, marginTop); return; }
      breakAt(bp, lines[breakLine].top, true);
      i = breakLine;
    }
  };

  const walkBlocks = (parentPos: number, parent: PMNode) => {
    parent.forEach((node, offset) => {
      const pos = parentPos + offset;
      const el = view.nodeDOM(pos) as HTMLElement | null;
      if (!el || !(el instanceof HTMLElement)) return;
      const r = el.getBoundingClientRect();
      const marginTop = parseFloat(getComputedStyle(el).marginTop) || 0;
      const top = rawY(rel(r.top));
      let bottom = rawY(rel(r.bottom));
      for (const f of Array.from(el.querySelectorAll<HTMLElement>(".om-float-left, .om-float-right"))) {
        bottom = Math.max(bottom, rawY(rel(f.getBoundingClientRect().bottom)));
      }
      const pbb = node.type === schema.nodes.paragraph && el.getAttribute("data-pbb") === "1";
      if ((forcedNext || pbb) && pos > 0 && top > pageStart + 1) { breakAt(pos, top, false, marginTop); }
      forcedNext = false;
      const fits = bottom <= pageStart + contentH + 0.5;
      if (node.type === schema.nodes.paragraph) {
        // Manual page breaks split the paragraph at the break element.
        const pbPositions: number[] = [];
        node.forEach((c, off) => { if (c.type === schema.nodes.hard_break && c.attrs.kind === "page") pbPositions.push(pos + 1 + off); });
        const pbEls = Array.from(el.querySelectorAll<HTMLElement>(".om-pagebreak"));
        if (pbPositions.length && pbEls.length === pbPositions.length) {
          const lines = lineBoxes(el);
          let lineFrom = 0;
          for (let i = 0; i < pbPositions.length; i++) {
            const pbTop = rawY(rel(pbEls[i].getBoundingClientRect().top));
            const seg = lines.filter((l, idx) => idx >= lineFrom && l.bottom <= pbTop + 1);
            paginateLines(pos, node, el, seg, marginTop);
            lineFrom += seg.length;
            const after = lines[lineFrom];
            const nextTop = after ? after.top : pbTop;
            breakAt(pbPositions[i] + 1, nextTop, true);
            // Skip the line box that is the break element itself if measured as a line.
            if (after && Math.abs(after.top - pbTop) < 1 && after.bottom - after.top < 6) lineFrom++;
          }
          const rest = lines.slice(lineFrom);
          if (rest.length) paginateLines(pos, node, el, rest, 0);
          return;
        }
        if (fits) return;
        paginateLines(pos, node, el, lineBoxes(el), marginTop);
        return;
      }
      if (fits) return;
      if (node.type === schema.nodes.table) {
        let rowIdx = 0;
        node.forEach((_row, roff) => {
          const rp = pos + 1 + roff;
          const rowEl = view.nodeDOM(rp) as HTMLElement | null;
          if (rowEl) {
            const rr = rowEl.getBoundingClientRect();
            const rtop = rawY(rel(rr.top)), rbottom = rawY(rel(rr.bottom));
            if (rbottom > pageStart + contentH + 0.5) {
              if (rowIdx === 0) { if (top > pageStart + 1) breakAt(pos, top, false, marginTop); }
              else if (rtop > pageStart + 1) breakAt(rp, rtop, false);
            }
          }
          rowIdx++;
        });
        return;
      }
      if (top > pageStart + 1) breakAt(pos, top, false, marginTop);
    });
  };

  walkBlocks(0, view.state.doc);
  return { breaks, pageCount: page };
}

// ---------------------------------------------------------------------------
// Layout passes run before pagination: tab stops and anchored objects.

/** Size tab elements so text after them lands on the next tab stop (Word rules). */
export function layoutTabs(view: EditorView) {
  const dom = view.dom as HTMLElement;
  const tabs = dom.querySelectorAll<HTMLElement>(".om-tab");
  if (!tabs.length) return;
  const editorRect = dom.getBoundingClientRect();
  const padL = parseFloat(getComputedStyle(dom).paddingLeft) || 0;
  const padR = parseFloat(getComputedStyle(dom).paddingRight) || 0;
  const contentLeft = editorRect.left + padL * zoomFactor;
  const contentW = (editorRect.width / zoomFactor) - padL - padR;
  const defaultTab = Math.max(8, twipsToPx(ctxDefaultTabStop()));
  let lastPara: HTMLElement | null = null;
  let stops: { pos: number; val: string }[] = [];
  let hang: number | null = null;
  let indR = 0;
  for (const tab of Array.from(tabs)) {
    const para = tab.closest<HTMLElement>(".om-p");
    if (!para) continue;
    // Tabs inside an absolutely positioned text box are relative to that box: skip (rare).
    if (tab.closest(".om-abs")) { tab.style.width = ""; continue; }
    if (para !== lastPara) {
      lastPara = para;
      stops = (para.getAttribute("data-tabs") || "").split(";").filter(Boolean).map((s) => { const [p, v] = s.split(":"); return { pos: twipsToPx(parseInt(p, 10)), val: v || "left" }; }).filter((s) => s.val !== "clear").sort((a, b) => a.pos - b.pos);
      const h = para.getAttribute("data-hang");
      hang = h ? twipsToPx(parseInt(h, 10)) : null;
      indR = twipsToPx(parseInt(para.getAttribute("data-indr") || "0", 10));
    }
    const r = tab.getBoundingClientRect();
    const x = (r.left - contentLeft) / zoomFactor;
    let stop: number | null = null, kind = "left";
    for (const s of stops) if (s.pos > x + 0.5) { stop = s.pos; kind = s.val; break; }
    if (hang !== null && hang > x + 0.5 && (stop === null || hang < stop)) { stop = hang; kind = "left"; }
    if (stop === null) {
      const lastCustom = stops.length ? stops[stops.length - 1].pos : 0;
      const start = Math.max(x, lastCustom);
      stop = (Math.floor((start + 0.5) / defaultTab) + 1) * defaultTab;
    }
    let width = stop - x;
    if (kind === "right" || kind === "center" || kind === "decimal") {
      const segW = measureAfterTab(tab, para, r);
      const adj = kind === "center" ? segW / 2 : segW;
      width = stop - x - adj;
      if (width < 0) width = stop - x;
    }
    const limit = contentW - indR - x - 1;
    if (width > limit) width = Math.max(0, limit);
    tab.style.width = Math.max(0, width) + "px";
  }
}

let cachedDefaultTab = 720;
export function setDefaultTabStop(tw: number) { cachedDefaultTab = tw || 720; }
function ctxDefaultTabStop() { return cachedDefaultTab; }

/** Width of the content following `tab` on the same line, up to the next tab or the paragraph end. */
function measureAfterTab(tab: HTMLElement, para: HTMLElement, tabRect: DOMRect): number {
  const range = document.createRange();
  range.setStartAfter(tab);
  let next: Element | null = null;
  const all = para.querySelectorAll(".om-tab");
  let found = false;
  for (const t of Array.from(all)) { if (found) { next = t; break; } if (t === tab) found = true; }
  if (next) range.setEndBefore(next); else range.setEnd(para, para.childNodes.length);
  let minL = Infinity, maxR = -Infinity;
  for (const rc of Array.from(range.getClientRects())) {
    if (rc.width <= 0) continue;
    if (rc.bottom <= tabRect.top || rc.top >= tabRect.bottom) continue; // other lines
    minL = Math.min(minL, rc.left); maxR = Math.max(maxR, rc.right);
  }
  if (minL === Infinity) return 0;
  return Math.max(0, (maxR - Math.max(minL, tabRect.right)) / zoomFactor);
}

/** Position anchored images / text boxes (absolute when the anchor paragraph carries no text or wrapping is "none"). */
export function layoutAnchors(view: EditorView, geom: PageGeom | null) {
  const dom = view.dom as HTMLElement;
  const anchors = dom.querySelectorAll<HTMLElement>(".om-anchor");
  if (!anchors.length) return;
  const editorRect = dom.getBoundingClientRect();
  const textLenCache = new Map<HTMLElement, number>();
  const textLen = (para: HTMLElement) => {
    let n = textLenCache.get(para);
    if (n !== undefined) return n;
    n = 0;
    const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
      acceptNode: (node) => node.nodeType === 1 ? ((node as HTMLElement).classList.contains("om-anchor") ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_SKIP) : NodeFilter.FILTER_ACCEPT,
    });
    while (walker.nextNode()) n += (walker.currentNode.nodeValue || "").trim().length;
    textLenCache.set(para, n);
    return n;
  };
  for (const el of Array.from(anchors)) {
    const para = el.closest<HTMLElement>(".om-p");
    if (!para || el.closest(".om-abs") !== el && el.closest(".om-abs")) continue;
    const wrap = el.getAttribute("data-wrap") || "none";
    const relV = el.getAttribute("data-relv") || "paragraph";
    const relH = el.getAttribute("data-relh") || "column";
    const hasPos = el.hasAttribute("data-offh") || el.hasAttribute("data-alignh") || el.hasAttribute("data-offv") || el.hasAttribute("data-alignv");
    const abs = hasPos && (wrap === "none" || textLen(para) === 0 || relV === "page" || relV === "margin");
    el.classList.remove("om-float-left", "om-float-right", "om-block-center");
    if (!abs) {
      el.classList.remove("om-abs");
      el.style.left = ""; el.style.top = "";
      const f = el.getAttribute("data-float");
      if (f === "left") el.classList.add("om-float-left"); else if (f === "right") el.classList.add("om-float-right"); else if (f === "center") el.classList.add("om-block-center");
      continue;
    }
    el.classList.add("om-abs");
    if (el.getAttribute("data-behind") === "1") el.classList.add("om-behind");
    const pr = para.getBoundingClientRect();
    const paraLeft = (pr.left - editorRect.left) / zoomFactor, paraTop = (pr.top - editorRect.top) / zoomFactor;
    const er = el.getBoundingClientRect();
    const w = er.width / zoomFactor, h = er.height / zoomFactor;
    const g = geom;
    const pageW = g ? g.pageW : editorRect.width / zoomFactor, pageH = g ? g.pageH : 0;
    const marL = g ? g.marL : parseFloat(getComputedStyle(dom).paddingLeft) || 0;
    const marT = g ? g.effMarT : parseFloat(getComputedStyle(dom).paddingTop) || 0;
    const marR = g ? g.marR : 0, marB = g ? g.effMarB : 0;
    const contentW = pageW - marL - marR, contentH = pageH - marT - marB;
    let pageTop = 0;
    if (g) {
      const pos = view.posAtDOM(para, 0);
      const page = pageAt(view.state, pos);
      pageTop = (page - 1) * (pageH + PAGE_GAP);
    }
    // Horizontal
    let baseX: number, areaW: number;
    switch (relH) {
      case "page": baseX = 0; areaW = pageW; break;
      case "leftMargin": case "insideMargin": baseX = 0; areaW = marL; break;
      case "rightMargin": case "outsideMargin": baseX = pageW - marR; areaW = marR; break;
      case "character": baseX = paraLeft; areaW = contentW; break;
      default: baseX = marL; areaW = contentW; // margin, column
    }
    const alignH = el.getAttribute("data-alignh");
    const offH = parseFloat(el.getAttribute("data-offh") || "");
    let x: number;
    if (alignH === "center") x = baseX + (areaW - w) / 2;
    else if (alignH === "right" || alignH === "outside") x = baseX + areaW - w;
    else if (alignH === "left" || alignH === "inside") x = baseX;
    else x = baseX + (isNaN(offH) ? 0 : offH);
    // Vertical
    let baseY: number, areaH: number;
    switch (relV) {
      case "page": baseY = pageTop; areaH = pageH; break;
      case "margin": case "insideMargin": case "outsideMargin": baseY = pageTop + marT; areaH = contentH; break;
      case "topMargin": baseY = pageTop; areaH = marT; break;
      case "bottomMargin": baseY = pageTop + pageH - marB; areaH = marB; break;
      default: baseY = paraTop; areaH = 0; // paragraph, line
    }
    const alignV = el.getAttribute("data-alignv");
    const offV = parseFloat(el.getAttribute("data-offv") || "");
    let y: number;
    if (alignV === "center") y = baseY + (areaH - h) / 2;
    else if (alignV === "bottom" || alignV === "outside") y = baseY + areaH - h;
    else if (alignV === "top" || alignV === "inside") y = baseY;
    else y = baseY + (isNaN(offV) ? 0 : offV);
    x = Math.max(0, Math.min(pageW - Math.min(w, pageW), x));
    el.style.left = (x - paraLeft) + "px";
    el.style.top = (y - paraTop) + "px";
  }
}

// ---------------------------------------------------------------------------
// Plugin

export interface PaginationOptions {
  pagesEl: HTMLElement;
  onPages?: (count: number) => void;
}

let hfHeights = { top: 0, bottom: 0 };

export function paginationPlugin(opts: PaginationOptions): Plugin<PagState> {
  let scheduled = 0;
  let lastSig = "";
  let passes = 0;

  const schedule = (view: EditorView) => {
    if (scheduled) return;
    scheduled = requestAnimationFrame(() => { scheduled = 0; run(view); });
  };

  const run = (view: EditorView) => {
    const st = paginationKey.getState(view.state)!;
    const sect: SectProps = view.state.doc.attrs.sect;
    if (st.mode !== "page") {
      lastSig = ""; // page mode must re-dispatch its breaks when we come back
      applyContinuous(view, sect);
      layoutTabs(view);
      layoutAnchors(view, null);
      if (st.breaks.length) view.dispatch(view.state.tr.setMeta(paginationKey, { breaks: [], pageCount: 1 }).setMeta("addToHistory", false));
      opts.pagesEl.innerHTML = "";
      opts.pagesEl.removeAttribute("data-key");
      opts.onPages?.(1);
      return;
    }
    const geom = applyPageStyle(view, sect);
    // Header/footer heights must be known before the body is laid out.
    ensureHFHeights(opts.pagesEl, geom, sect);
    const geom2 = applyPageStyle(view, sect);
    layoutTabs(view);
    layoutAnchors(view, geom2);
    const res = measure(view, geom2);
    const sig = res.breaks.map((b) => b.pos + ":" + Math.round(b.gap) + ":" + (b.inline ? 1 : 0)).join(",");
    if (sig !== lastSig && passes < 4) {
      lastSig = sig;
      passes++;
      view.dispatch(view.state.tr.setMeta(paginationKey, { breaks: res.breaks, pageCount: res.pageCount }).setMeta("addToHistory", false));
      schedule(view); // verify after re-render
      return;
    }
    passes = 0;
    renderPages(opts.pagesEl, geom2, res.pageCount, sect);
    opts.onPages?.(res.pageCount);
  };

  return new Plugin<PagState>({
    key: paginationKey,
    state: {
      init: () => ({ breaks: [], pageCount: 1, mode: "page" }),
      apply(tr, prev) {
        const meta = tr.getMeta(paginationKey) as Partial<PagState> | undefined;
        let next = meta ? { ...prev, ...meta } : prev;
        if (tr.docChanged && !(meta && meta.breaks)) {
          next = { ...next, breaks: next.breaks.map((b) => ({ ...b, pos: tr.mapping.map(b.pos) })) };
        }
        return next;
      },
    },
    props: {
      decorations(state: EditorState) {
        const st = paginationKey.getState(state)!;
        if (st.mode !== "page" || !st.breaks.length) return null;
        return decorationsFor(state.doc, st.breaks);
      },
    },
    view(view) {
      const onResize = () => { lastSig = ""; schedule(view); };
      window.addEventListener("resize", onResize);
      (document as any).fonts?.addEventListener?.("loadingdone", onResize);
      schedule(view);
      return {
        update(v, prev) {
          const st = paginationKey.getState(v.state)!;
          const pst = paginationKey.getState(prev)!;
          if (v.state.doc !== prev.doc || st !== pst) schedule(v);
        },
        destroy() { window.removeEventListener("resize", onResize); if (scheduled) cancelAnimationFrame(scheduled); },
      };
    },
  });
}

function applyPageStyle(view: EditorView, sect: SectProps): PageGeom {
  const base = geometry(sect);
  const effT = Math.max(base.marT, base.headerDist + hfHeights.top);
  const effB = Math.max(base.marB, base.footerDist + hfHeights.bottom);
  const geom = geometry(sect, effT, effB);
  const dom = view.dom as HTMLElement;
  dom.style.width = geom.pageW + "px";
  dom.style.padding = `${geom.effMarT}px ${geom.marR}px ${geom.effMarB}px ${geom.marL}px`;
  dom.style.minHeight = geom.pageH + "px";
  dom.style.background = "transparent";
  return geom;
}

function applyContinuous(view: EditorView, sect: SectProps) {
  const geom = geometry(sect);
  const dom = view.dom as HTMLElement;
  dom.style.width = geom.pageW + "px";
  dom.style.padding = `28px ${geom.marR}px 60px ${geom.marL}px`;
  dom.style.minHeight = "";
  dom.style.background = "var(--paper)";
}

function ensureHFHeights(pagesEl: HTMLElement, geom: PageGeom, sect: SectProps) {
  if (!hfProvider) { hfHeights = { top: 0, bottom: 0 }; return; }
  if (pagesEl.getAttribute("data-hf") === sect.xml) return;
  pagesEl.setAttribute("data-hf", sect.xml || "");
  // Render one probe page to measure header/footer heights.
  const probe = document.createElement("div");
  probe.className = "page-bg";
  probe.style.width = geom.pageW + "px";
  probe.style.height = geom.pageH + "px";
  probe.style.visibility = "hidden";
  renderHF(probe, 2, 2, geom, sect);
  pagesEl.appendChild(probe);
  const h = probe.querySelector<HTMLElement>(".page-hf.header");
  const f = probe.querySelector<HTMLElement>(".page-hf.footer");
  hfHeights = {
    top: h ? h.getBoundingClientRect().height / zoomFactor : 0,
    bottom: f ? f.getBoundingClientRect().height / zoomFactor : 0,
  };
  probe.remove();
  pagesEl.removeAttribute("data-key");
}

function renderPages(pagesEl: HTMLElement, geom: PageGeom, count: number, sect: SectProps) {
  const key = `${geom.pageW}x${geom.pageH}:${count}:${sect.xml?.length || 0}`;
  if (pagesEl.getAttribute("data-key") === key) { updatePageNumbers(pagesEl, count); return; }
  pagesEl.setAttribute("data-key", key);
  pagesEl.innerHTML = "";
  const frag = document.createDocumentFragment();
  for (let i = 0; i < count; i++) {
    const p = document.createElement("div");
    p.className = "page-bg";
    p.style.width = geom.pageW + "px";
    p.style.height = geom.pageH + "px";
    p.style.top = `calc(${i} * (${geom.pageH}px + var(--page-gap)))`;
    p.setAttribute("data-page", String(i + 1));
    const badge = document.createElement("div");
    badge.className = "page-num-badge";
    badge.textContent = `${i + 1}`;
    p.appendChild(badge);
    renderHF(p, i + 1, count, geom, sect);
    frag.appendChild(p);
  }
  pagesEl.appendChild(frag);
}

function updatePageNumbers(pagesEl: HTMLElement, count: number) {
  pagesEl.querySelectorAll<HTMLElement>("[data-field=NUMPAGES]").forEach((el) => (el.textContent = String(count)));
}

let cachedHF = new Map<string, HTMLElement>();
export function resetHeaderFooterCache() { cachedHF = new Map(); hfHeights = { top: 0, bottom: 0 }; }

function pickHF(map: { default?: string; first?: string; even?: string }, page: number, titlePg: boolean): string | undefined {
  if (page === 1 && titlePg && map.first !== undefined) return map.first;
  if (page % 2 === 0 && map.even) return map.even;
  return map.default;
}

function renderHF(pageEl: HTMLElement, page: number, count: number, geom: PageGeom, sect: SectProps) {
  if (!hfProvider) return;
  const hId = pickHF(sect.headers, page, sect.titlePg);
  const fId = pickHF(sect.footers, page, sect.titlePg);
  const make = (hf: HeaderFooter | undefined, kind: "header" | "footer") => {
    if (!hf) return;
    let tpl = cachedHF.get(hf.rId);
    if (!tpl) {
      tpl = document.createElement("div");
      tpl.appendChild(DOMSerializer.fromSchema(schema).serializeFragment(hf.doc.content));
      cachedHF.set(hf.rId, tpl);
    }
    const el = tpl.cloneNode(true) as HTMLElement;
    el.className = "page-hf " + kind;
    el.style.left = geom.marL + "px";
    el.style.width = geom.contentW + "px";
    if (kind === "header") el.style.top = geom.headerDist + "px";
    else el.style.bottom = geom.footerDist + "px";
    el.querySelectorAll<HTMLElement>("[data-field=PAGE]").forEach((f) => (f.textContent = String(page)));
    el.querySelectorAll<HTMLElement>("[data-field=NUMPAGES], [data-field=SECTIONPAGES]").forEach((f) => (f.textContent = String(count)));
    el.querySelectorAll<HTMLElement>(".om-field").forEach((f) => { f.style.background = "transparent"; });
    pageEl.appendChild(el);
  };
  make(hId ? hfProvider.headers.get(hId) : undefined, "header");
  make(fId ? hfProvider.footers.get(fId) : undefined, "footer");
}

/** Page number containing a document position (1-based). */
export function pageAt(state: EditorState, pos: number): number {
  const st = paginationKey.getState(state);
  if (!st || st.mode !== "page") return 1;
  let page = 1;
  for (const b of st.breaks) if (pos >= b.pos) page++; else break;
  return page;
}

/** Document position at the start of a page (1-based). */
export function pageStartPos(state: EditorState, page: number): number {
  const st = paginationKey.getState(state);
  if (!st || page <= 1 || !st.breaks.length) return 0;
  const b = st.breaks[Math.min(page - 2, st.breaks.length - 1)];
  return b.pos;
}

export function setViewMode(view: EditorView, mode: "page" | "continuous") {
  view.dispatch(view.state.tr.setMeta(paginationKey, { mode }).setMeta("addToHistory", false));
}

export function relayout(view: EditorView) {
  view.dispatch(view.state.tr.setMeta(paginationKey, {}).setMeta("addToHistory", false));
}

export type { Transaction };
