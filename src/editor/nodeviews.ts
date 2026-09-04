// Node views: images with resize handles.
import { NodeView, EditorView } from "prosemirror-view";
import { Node as PMNode } from "prosemirror-model";
import { NodeSelection } from "prosemirror-state";
import { fmt } from "../docx/units";
import { anchorAttrs } from "../schema";

export class ImageView implements NodeView {
  dom: HTMLElement;
  img: HTMLImageElement;
  private handles: HTMLElement[] = [];
  private dragging: { startX: number; startY: number; w: number; h: number; corner: string } | null = null;

  constructor(private node: PMNode, private view: EditorView, private getPos: () => number | undefined) {
    this.dom = document.createElement("span");
    this.dom.className = "om-img-wrap";
    this.img = document.createElement("img");
    this.img.className = "om-img";
    this.img.draggable = false;
    this.dom.appendChild(this.img);
    for (const c of ["nw", "ne", "sw", "se"]) {
      const h = document.createElement("span");
      h.className = "handle " + c;
      h.addEventListener("mousedown", (e) => this.startDrag(e, c));
      this.dom.appendChild(h);
      this.handles.push(h);
    }
    this.render();
  }

  private static alphaCache = new Map<string, boolean>();

  /** Detect transparency so dark mode can give the image a light plate instead of losing dark strokes. */
  private detectAlpha() {
    const a = this.node.attrs;
    const ext = String(a.ext || "").toLowerCase();
    if (ext === "jpeg" || ext === "jpg" || ext === "bmp") { this.dom.classList.remove("om-img-alpha"); return; }
    const src: string = a.src;
    const cached = ImageView.alphaCache.get(src);
    if (cached !== undefined) { this.dom.classList.toggle("om-img-alpha", cached); return; }
    const probe = new Image();
    probe.onload = () => {
      try {
        const w = Math.min(64, probe.naturalWidth || 64), h = Math.min(64, probe.naturalHeight || 64);
        const c = document.createElement("canvas");
        c.width = w; c.height = h;
        const g = c.getContext("2d", { willReadFrequently: true })!;
        g.drawImage(probe, 0, 0, w, h);
        const d = g.getImageData(0, 0, w, h).data;
        let alpha = false;
        for (let i = 3; i < d.length; i += 4) if (d[i] < 240) { alpha = true; break; }
        ImageView.alphaCache.set(src, alpha);
        this.dom.classList.toggle("om-img-alpha", alpha);
      } catch { /* tainted or unsupported: assume opaque */ }
    };
    probe.src = src;
  }

  private render() {
    const a = this.node.attrs;
    if (this.img.src !== a.src) { this.img.src = a.src; this.detectAlpha(); }
    this.img.alt = a.alt || "";
    this.img.style.width = fmt(a.w) + "px";
    this.img.style.height = fmt(a.h) + "px";
    this.img.title = a.name || "";
    let cls = "om-img-wrap";
    for (const at of Array.from(this.dom.attributes)) if (at.name.startsWith("data-")) this.dom.removeAttribute(at.name);
    if (a.kind === "anchor" && a.wrap) {
      if (a.wrap.float === "left") cls += " om-float-left";
      else if (a.wrap.float === "right") cls += " om-float-right";
      else if (a.wrap.float === "center") cls += " om-block-center";
      cls += " om-anchor";
      for (const [k, v] of Object.entries(anchorAttrs(a.wrap))) this.dom.setAttribute(k, v);
    }
    if (this.dom.classList.contains("selected")) cls += " selected";
    if (this.dom.classList.contains("om-img-alpha")) cls += " om-img-alpha";
    this.dom.className = cls;
  }

  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  selectNode() { this.dom.classList.add("selected"); }
  deselectNode() { this.dom.classList.remove("selected"); }

  stopEvent(e: Event) {
    // Handle drags are ours; everything else goes to ProseMirror.
    return this.dragging !== null || (e.target as HTMLElement)?.classList?.contains("handle");
  }

  ignoreMutation() { return true; }

  private startDrag(e: MouseEvent, corner: string) {
    e.preventDefault();
    e.stopPropagation();
    const pos = this.getPos();
    if (pos === undefined) return;
    this.view.dispatch(this.view.state.tr.setSelection(NodeSelection.create(this.view.state.doc, pos)));
    this.dragging = { startX: e.clientX, startY: e.clientY, w: this.node.attrs.w, h: this.node.attrs.h, corner };
    const zoom = parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--zoom")) || 1;
    const move = (ev: MouseEvent) => {
      if (!this.dragging) return;
      const dx = (ev.clientX - this.dragging.startX) / zoom;
      const dy = (ev.clientY - this.dragging.startY) / zoom;
      const sx = this.dragging.corner.includes("e") ? 1 : -1;
      const sy = this.dragging.corner.includes("s") ? 1 : -1;
      let w = Math.max(8, this.dragging.w + dx * sx);
      let h = Math.max(8, this.dragging.h + dy * sy);
      if (!ev.shiftKey) {
        // keep aspect ratio, driven by the larger relative change
        const ratio = this.dragging.w / this.dragging.h;
        if (Math.abs(dx) > Math.abs(dy)) h = w / ratio; else w = h * ratio;
      }
      this.img.style.width = fmt(w) + "px";
      this.img.style.height = fmt(h) + "px";
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      if (!this.dragging) return;
      const w = parseFloat(this.img.style.width), h = parseFloat(this.img.style.height);
      this.dragging = null;
      const p = this.getPos();
      if (p !== undefined && (Math.round(w) !== Math.round(this.node.attrs.w) || Math.round(h) !== Math.round(this.node.attrs.h))) {
        this.view.dispatch(this.view.state.tr.setNodeMarkup(p, undefined, { ...this.node.attrs, w: Math.round(w * 100) / 100, h: Math.round(h * 100) / 100 }));
      }
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }
}
