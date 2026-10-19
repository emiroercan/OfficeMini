// LaTeX math: rendering (KaTeX, loaded on first use), the node view and the equation dialog.
//
// A formula is one inline `math` node holding its TeX source. Display math ($$…$$) is the same
// node with display=true, sitting alone in a centered paragraph, so pagination, lists, tables
// and the DOCX writer keep dealing with ordinary paragraphs.
import type { Node as PMNode } from "prosemirror-model";
import type { EditorView, NodeView } from "prosemirror-view";
import { NodeSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";
import { schema } from "../schema";

type Katex = typeof import("katex").default;
let katex: Katex | null = null;
let loading: Promise<Katex> | null = null;
const waiting = new Set<MathView>();
let afterLoad: (() => void) | null = null;

/** Called once KaTeX (and its fonts) arrive, so late-rendered formulas can be re-paginated. */
export function onMathReady(cb: () => void) { afterLoad = cb; }

/** Load KaTeX and its stylesheet. Resolves when formulas can be rendered with final metrics. */
export function ensureKatex(): Promise<Katex> {
  if (katex) return Promise.resolve(katex);
  if (!loading) {
    loading = Promise.all([import("katex"), import("katex/dist/katex.min.css")]).then(async ([m]) => {
      katex = m.default;
      // Formula boxes are measured by pagination: wait for the math fonts they use.
      await Promise.all(["KaTeX_Main", "KaTeX_Math", "KaTeX_Size1", "KaTeX_Size2", "KaTeX_AMS"].map((f) => document.fonts.load(`16px ${f}`).catch(() => null)));
      for (const v of waiting) v.render();
      waiting.clear();
      afterLoad?.();
      return katex;
    });
  }
  return loading;
}

export function docHasMath(doc: PMNode): boolean {
  let found = false;
  doc.descendants((n) => { if (found) return false; if (n.type === schema.nodes.math) found = true; return !found; });
  return found;
}

/** Render TeX into `el`. Errors show the source in red (KaTeX's own error rendering). */
export function renderTex(el: HTMLElement, tex: string, display: boolean): boolean {
  if (!katex) { el.textContent = display ? `$$${tex}$$` : `$${tex}$`; return false; }
  try {
    katex.render(tex, el, { displayMode: display, throwOnError: false, output: "htmlAndMathml", strict: "ignore", trust: false });
  } catch (e) {
    el.textContent = tex;
    el.title = (e as Error).message;
  }
  return true;
}

export class MathView implements NodeView {
  dom: HTMLElement;
  constructor(private node: PMNode, private view: EditorView, private getPos: () => number | undefined) {
    this.dom = document.createElement("span");
    this.dom.className = "om-math" + (node.attrs.display ? " om-math-display" : "");
    this.dom.contentEditable = "false";
    this.dom.addEventListener("dblclick", (e) => { e.preventDefault(); this.edit(); });
    this.render();
  }
  render() {
    if (!renderTex(this.dom, this.node.attrs.tex, this.node.attrs.display)) { waiting.add(this); ensureKatex(); }
    this.dom.title = "Double-click to edit the equation";
  }
  edit() {
    const pos = this.getPos();
    if (pos === undefined) return;
    this.view.dispatch(this.view.state.tr.setSelection(NodeSelection.create(this.view.state.doc, pos)));
    editMath(this.view);
  }
  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    const changed = node.attrs.tex !== this.node.attrs.tex || node.attrs.display !== this.node.attrs.display;
    this.node = node;
    if (changed) { this.dom.className = "om-math" + (node.attrs.display ? " om-math-display" : ""); this.render(); }
    return true;
  }
  selectNode() { this.dom.classList.add("ProseMirror-selectednode"); }
  deselectNode() { this.dom.classList.remove("ProseMirror-selectednode"); }
  stopEvent(e: Event) { return e.type === "dblclick"; }
  ignoreMutation() { return true; }
  destroy() { waiting.delete(this); }
}

// ---------------------------------------------------------------------------
// Equation dialog

type DialogFn = (tex: string, display: boolean) => Promise<{ tex: string; display: boolean } | null>;
let dialogImpl: DialogFn | null = null;
/** main.ts provides the dialog (it owns the dialog widgets). */
export function setMathDialog(fn: DialogFn) { dialogImpl = fn; }

/** Edit the selected formula, or insert a new one (selected text becomes its source). */
export function editMath(view: EditorView, forceDisplay?: boolean): void {
  if (!dialogImpl) return;
  const { state } = view;
  const sel = state.selection;
  const existing = sel instanceof NodeSelection && sel.node.type === schema.nodes.math ? sel.node : null;
  const initial = existing ? existing.attrs.tex : state.doc.textBetween(sel.from, sel.to, " ");
  const $from = sel.$from;
  const emptyPara = !existing && $from.parent.type === schema.nodes.paragraph && $from.parent.content.size === 0;
  const display = existing ? existing.attrs.display : forceDisplay ?? emptyPara;
  const from = sel.from, to = sel.to;
  dialogImpl(initial, display).then((res) => {
    view.focus();
    if (!res) return;
    const tex = res.tex.trim();
    let tr = view.state.tr;
    if (!tex) { if (existing) view.dispatch(tr.delete(from, to)); return; }
    const fence = existing && existing.attrs.display === res.display ? existing.attrs.fence : null;
    const node = schema.nodes.math.create({ tex, display: res.display, fence }, null, view.state.storedMarks || $from.marks());
    tr = tr.replaceRangeWith(from, to, node);
    if (res.display) {
      // Display math is centered on its own line, like LaTeX's \[ \].
      const $p = tr.doc.resolve(tr.mapping.map(from));
      for (let d = $p.depth; d > 0; d--) {
        const n = $p.node(d);
        if (n.type === schema.nodes.paragraph) {
          if (n.childCount === 1) tr.setNodeMarkup($p.before(d), undefined, { ...n.attrs, props: { ...n.attrs.props, jc: "center" } });
          break;
        }
      }
    }
    view.dispatch(tr.scrollIntoView());
  });
}

export const insertMathCmd = (display?: boolean): Command => (_state, dispatch, view) => {
  if (!dispatch || !view) return true;
  editMath(view, display);
  return true;
};

/** Enter on a selected formula opens it for editing. */
export const editSelectedMath: Command = (state, dispatch, view) => {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection) || sel.node.type !== schema.nodes.math) return false;
  if (dispatch && view) editMath(view);
  return true;
};
