// Formatting and editing commands. Formatting lives in a single `rpr` mark
// (character properties) and in paragraph attrs, mirroring WordprocessingML.
import { Command, EditorState, TextSelection, NodeSelection, Transaction } from "prosemirror-state";
import { Node as PMNode, Mark, Fragment } from "prosemirror-model";
import { schema, effectiveRunProps, paragraphStyle } from "../schema";
import { RunProps, ParaProps, merge } from "../docx/props";
import { ctx } from "../docx/styles";
import { CellSelection } from "prosemirror-tables";

const rprType = schema.marks.rpr;
const linkType = schema.marks.link;

// ---------------------------------------------------------------------------
// Inspecting the selection

/** Paragraph node containing the selection head, with its position. */
export function selectionParagraph(state: EditorState): { node: PMNode; pos: number } | null {
  const $p = state.selection.$from;
  for (let d = $p.depth; d >= 0; d--) {
    const n = $p.node(d);
    if (n.type === schema.nodes.paragraph) return { node: n, pos: $p.before(d) };
  }
  return null;
}

/** Effective run props (style + mark) at the selection, for toolbar state. */
export function selectionRunProps(state: EditorState): RunProps {
  const para = selectionParagraph(state);
  let base: RunProps = para ? paragraphStyle(para.node).rPr : { ...ctx.docRPr };
  const marks = state.storedMarks || markAtSelection(state);
  const m = rprType.isInSet(marks);
  if (m) base = merge(base, effectiveRunProps(m.attrs.props as RunProps));
  return base;
}

function markAtSelection(state: EditorState): readonly Mark[] {
  const sel = state.selection;
  if (sel instanceof NodeSelection) return sel.node.marks;
  const $from = sel.$from;
  if (sel.empty) return $from.marks();
  // For a range use the marks of the first text node inside it.
  let found: readonly Mark[] | null = null;
  state.doc.nodesBetween(sel.from, sel.to, (node) => {
    if (found) return false;
    if (node.isText) { found = node.marks; return false; }
    return true;
  });
  return found || $from.marks();
}

export function selectionParaProps(state: EditorState): { direct: ParaProps; eff: ParaProps } {
  const para = selectionParagraph(state);
  if (!para) return { direct: {}, eff: {} };
  return { direct: para.node.attrs.props as ParaProps, eff: paragraphStyle(para.node).pPr };
}

/** Does every text node in the selection have prop `key` truthy (resolved against style)? */
export function selectionHasRunProp(state: EditorState, key: keyof RunProps): boolean {
  const sel = state.selection;
  if (sel.empty) return !!selectionRunProps(state)[key];
  let all = true, any = false;
  for (const range of sel.ranges) {
    state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos, parent) => {
      if (!all) return false;
      if (node.isText) {
        any = true;
        const m = rprType.isInSet(node.marks);
        let v: any = m ? effectiveRunProps(m.attrs.props as RunProps)[key] : undefined;
        if (v === undefined && parent) v = paragraphStyle(parent).rPr[key];
        if (!v) all = false;
      }
      return true;
    });
  }
  return any && all;
}

// ---------------------------------------------------------------------------
// Run formatting

export function applyRunProps(patch: Partial<RunProps>): Command {
  return (state, dispatch) => {
    const sel = state.selection;
    if (sel.empty && !(sel instanceof NodeSelection)) {
      const marks = state.storedMarks || sel.$from.marks();
      const existing = rprType.isInSet(marks);
      const mark = rprType.create({ xml: existing?.attrs.xml ?? null, props: merge(existing?.attrs.props || {}, patch) });
      if (dispatch) dispatch(state.tr.setStoredMarks(mark.addToSet(marks)));
      return true;
    }
    const tr = state.tr;
    for (const range of sel.ranges) {
      const from = range.$from.pos, to = range.$to.pos;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (node.isInline && (node.isText || node.type === schema.nodes.tab || node.type === schema.nodes.image)) {
          const existing = rprType.isInSet(node.marks);
          const mark = rprType.create({ xml: existing?.attrs.xml ?? null, props: merge(existing?.attrs.props || {}, patch) });
          tr.addMark(Math.max(from, pos), Math.min(to, pos + node.nodeSize), mark);
        }
        return true;
      });
    }
    // Also record the change for typing at the selection edge.
    if (dispatch) dispatch(tr.scrollIntoView());
    return true;
  };
}

function toggleRun(key: keyof RunProps): Command {
  return (state, dispatch, view) => {
    const on = selectionHasRunProp(state, key);
    return applyRunProps({ [key]: !on } as Partial<RunProps>)(state, dispatch, view);
  };
}

export const toggleBold: Command = toggleRun("b");
export const toggleItalic: Command = toggleRun("i");
export const toggleStrike: Command = toggleRun("strike");
export const toggleUnderline: Command = (state, dispatch, view) => {
  const cur = selectionRunProps(state).u;
  const on = !!cur && cur !== "none";
  return applyRunProps({ u: on ? "none" : "single" })(state, dispatch, view);
};
export const toggleSuperscript: Command = (state, dispatch, view) => {
  const on = selectionRunProps(state).vertAlign === "superscript";
  return applyRunProps({ vertAlign: on ? "baseline" : "superscript" })(state, dispatch, view);
};
export const toggleSubscript: Command = (state, dispatch, view) => {
  const on = selectionRunProps(state).vertAlign === "subscript";
  return applyRunProps({ vertAlign: on ? "baseline" : "subscript" })(state, dispatch, view);
};
export const toggleCaps: Command = toggleRun("caps");
export const toggleSmallCaps: Command = toggleRun("smallCaps");

export function setFont(name: string): Command { return applyRunProps({ font: name }); }
export function setFontSize(pt: number): Command { return applyRunProps({ size: Math.round(pt * 2) }); }
export function setColor(hex: string | null): Command { return applyRunProps({ color: hex ? hex.replace("#", "").toUpperCase() : "auto" }); }
export function setHighlight(name: string | null): Command { return applyRunProps({ highlight: name || "none" }); }

const SIZE_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
export function fontSizeStep(dir: 1 | -1, onePoint = false): Command {
  return (state, dispatch, view) => {
    const cur = (selectionRunProps(state).size || 22) / 2;
    let next: number;
    if (onePoint) next = cur + dir;
    else if (dir > 0) next = SIZE_STEPS.find((s) => s > cur) ?? cur + 2;
    else next = [...SIZE_STEPS].reverse().find((s) => s < cur) ?? cur - 1;
    next = Math.max(1, Math.min(400, next));
    return setFontSize(next)(state, dispatch, view);
  };
}

/** Ctrl+Space: remove direct character formatting (keep the style). */
export const clearFormatting: Command = (state, dispatch) => {
  const sel = state.selection;
  const tr = state.tr;
  if (sel.empty) {
    const marks = (state.storedMarks || sel.$from.marks()).filter((m) => m.type !== rprType);
    if (dispatch) dispatch(tr.setStoredMarks(marks));
    return true;
  }
  for (const range of sel.ranges) tr.removeMark(range.$from.pos, range.$to.pos, rprType);
  if (dispatch) dispatch(tr);
  return true;
};

// ---------------------------------------------------------------------------
// Paragraph formatting

/** Iterate paragraphs touched by the selection. */
export function forEachParagraph(state: EditorState, fn: (node: PMNode, pos: number) => void) {
  const seen = new Set<number>();
  for (const range of state.selection.ranges) {
    state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos) => {
      if (node.type === schema.nodes.paragraph) { if (!seen.has(pos)) { seen.add(pos); fn(node, pos); } return false; }
      return true;
    });
  }
  if (!seen.size) { const p = selectionParagraph(state); if (p) fn(p.node, p.pos); }
}

export function setParaProps(patch: Partial<ParaProps> | ((node: PMNode) => Partial<ParaProps> | null)): Command {
  return (state, dispatch) => {
    const tr = state.tr;
    let changed = false;
    forEachParagraph(state, (node, pos) => {
      const p = typeof patch === "function" ? patch(node) : patch;
      if (!p) return;
      // `undefined` in the patch deletes a direct-formatting key; null sets it explicitly.
      const props: any = { ...(node.attrs.props as ParaProps) };
      for (const k of Object.keys(p)) { const v = (p as any)[k]; if (v === undefined) delete props[k]; else props[k] = v; }
      tr.setNodeMarkup(tr.mapping.map(pos), undefined, { ...node.attrs, props });
      changed = true;
    });
    if (changed && dispatch) dispatch(tr.scrollIntoView());
    return changed;
  };
}

export function setAlign(jc: "left" | "center" | "right" | "both"): Command { return setParaProps({ jc }); }

export function setLineSpacing(mult: number): Command {
  return setParaProps({ spLine: Math.round(mult * 240), spLineRule: "auto" });
}
export function setSpaceBefore(pt: number): Command { return setParaProps({ spBefore: Math.round(pt * 20), spBeforeAuto: false }); }
export function setSpaceAfter(pt: number): Command { return setParaProps({ spAfter: Math.round(pt * 20), spAfterAuto: false }); }

export function setParaStyle(styleId: string | null): Command {
  return setParaProps({ pStyle: styleId, numId: null, ilvl: null });
}

/** Increase/decrease indent (Tab/Shift-Tab in lists, Ctrl+M / Ctrl+Shift+M). */
export function indentParagraphs(dir: 1 | -1): Command {
  return setParaProps((node) => {
    const eff = paragraphStyle(node).pPr;
    if (eff.numId) {
      const lvl = Math.max(0, Math.min(8, (eff.ilvl || 0) + dir));
      if (lvl === (eff.ilvl || 0)) return null;
      // Level indents come from the numbering definition; drop direct overrides.
      return { numId: eff.numId, ilvl: lvl, indLeft: undefined as any, indHanging: undefined as any, indFirstLine: undefined as any };
    }
    const left = Math.max(0, (eff.indLeft || 0) + dir * 720);
    return { indLeft: left };
  });
}

export function toggleList(kind: "bullet" | "decimal"): Command {
  return (state, dispatch) => {
    // Determine whether all selected paragraphs already are lists of this kind.
    const paras: { node: PMNode; pos: number }[] = [];
    forEachParagraph(state, (node, pos) => paras.push({ node, pos }));
    if (!paras.length) return false;
    const isKind = (n: PMNode) => {
      const eff = paragraphStyle(n).pPr;
      if (!eff.numId) return false;
      return ctx.isBulletList(eff.numId) === (kind === "bullet");
    };
    const allOn = paras.every((p) => isKind(p.node));
    const tr = state.tr;
    if (allOn) {
      for (const p of paras) {
        const props = { ...(p.node.attrs.props as ParaProps) };
        props.numId = null; props.ilvl = null;
        if (props.pStyle && listStyleId() && props.pStyle === listStyleId()) props.pStyle = null;
        tr.setNodeMarkup(p.pos, undefined, { ...p.node.attrs, props });
      }
      if (dispatch) dispatch(tr);
      return true;
    }
    // Continue the list of the previous paragraph if it is the same kind, else create one.
    let numId: number | null = null;
    const $first = state.doc.resolve(paras[0].pos);
    const parent = $first.parent, index = $first.index();
    for (let i = index - 1; i >= 0; i--) {
      const prev = parent.child(i);
      if (prev.type !== schema.nodes.paragraph) break;
      const eff = paragraphStyle(prev).pPr;
      if (eff.numId) { if (ctx.isBulletList(eff.numId) === (kind === "bullet")) numId = eff.numId; break; }
      if (prev.content.size > 0) break;
    }
    if (!numId) numId = ctx.createList(kind);
    const ls = listStyleId();
    for (const p of paras) {
      const props = { ...(p.node.attrs.props as ParaProps) };
      props.numId = numId; props.ilvl = props.ilvl ?? 0;
      delete props.indLeft; delete props.indHanging; delete props.indFirstLine;
      if (ls && (!props.pStyle || props.pStyle === ctx.defaultPara)) props.pStyle = ls;
      tr.setNodeMarkup(p.pos, undefined, { ...p.node.attrs, props });
    }
    if (dispatch) dispatch(tr);
    return true;
  };
}

function listStyleId(): string | null { return ctx.styleIdByName("List Paragraph"); }

// ---------------------------------------------------------------------------
// Insertion

export const insertLineBreak: Command = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.hard_break.create({ kind: "textWrapping" })).scrollIntoView());
  return true;
};

export const insertPageBreak: Command = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.hard_break.create({ kind: "page" })).scrollIntoView());
  return true;
};

export const insertTab: Command = (state, dispatch) => {
  const marks = state.storedMarks || state.selection.$from.marks();
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.tab.create(null, undefined, marks)).scrollIntoView());
  return true;
};

export function insertTextCmd(text: string): Command {
  return (state, dispatch) => { if (dispatch) dispatch(state.tr.insertText(text).scrollIntoView()); return true; };
}

/** Enter: split the paragraph keeping its attributes and marks (Word behaviour). */
export const splitParagraph: Command = (state, dispatch) => {
  const { $from, $to } = state.selection;
  if (state.selection instanceof NodeSelection && state.selection.node.isBlock) {
    if (!$from.parentOffset && $from.parent.type !== schema.nodes.paragraph) return false;
  }
  const para = selectionParagraph(state);
  if (!para) return false;
  const node = para.node;
  const props = node.attrs.props as ParaProps;
  const eff = paragraphStyle(node).pPr;
  const tr = state.tr;
  const marks = state.storedMarks || $from.marks();
  // Empty list item: Enter ends the list (Word).
  if (node.content.size === 0 && eff.numId && $from.parent === node) {
    const np = { ...props, numId: null, ilvl: null };
    const ls = listStyleId();
    if (ls && np.pStyle === ls) np.pStyle = null;
    delete np.indLeft; delete np.indHanging; delete np.indFirstLine;
    tr.setNodeMarkup(para.pos, undefined, { ...node.attrs, props: np });
    if (dispatch) dispatch(tr.scrollIntoView());
    return true;
  }
  tr.deleteSelection();
  const $pos = tr.selection.$from;
  const atEnd = $pos.parentOffset === $pos.parent.content.size;
  let attrs: Record<string, any> = { ...node.attrs, sectPr: null, pPr: node.attrs.pPr };
  const newProps = { ...props };
  // Word: after a heading, the next paragraph uses the style's "next" style (usually Normal).
  const style = ctx.style(props.pStyle);
  if (atEnd && style && style.next && style.next !== style.id) {
    newProps.pStyle = style.next === ctx.defaultPara ? null : style.next;
    attrs = { ...attrs, pPr: null };
  }
  if (newProps.pageBreakBefore) newProps.pageBreakBefore = false;
  attrs = { ...attrs, props: newProps };
  // Keep the original paragraph's attrs for the first half, new attrs for the second.
  tr.split($pos.pos, 1, [{ type: schema.nodes.paragraph, attrs }]);
  if (marks.length) tr.ensureMarks(marks);
  if (dispatch) dispatch(tr.scrollIntoView());
  return true;
};

/** Backspace at the start of a list paragraph removes the numbering first (Word). */
export const backspaceListStart: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  if (!empty || $from.parentOffset !== 0) return false;
  const node = $from.parent;
  if (node.type !== schema.nodes.paragraph) return false;
  const eff = paragraphStyle(node).pPr;
  if (!eff.numId) return false;
  const props = { ...(node.attrs.props as ParaProps), numId: null, ilvl: null };
  if (dispatch) dispatch(state.tr.setNodeMarkup($from.before(), undefined, { ...node.attrs, props }));
  return true;
};

// ---------------------------------------------------------------------------
// Links

export function setLink(href: string): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    const tr = state.tr;
    const mark = linkType.create({ href, rId: null, anchor: href.startsWith("#") ? href.slice(1) : null, tooltip: null, raw: null });
    if (empty) {
      tr.insertText(href, from, to);
      tr.addMark(from, from + href.length, mark);
    } else {
      tr.addMark(from, to, mark);
    }
    if (dispatch) dispatch(tr);
    return true;
  };
}

export const unsetLink: Command = (state, dispatch) => {
  const { $from, from, to, empty } = state.selection;
  const tr = state.tr;
  if (empty) {
    // Remove the link mark spanning the cursor.
    const mark = linkType.isInSet($from.marks());
    if (!mark) return false;
    let start = from, end = from;
    const parent = $from.parent, base = $from.start();
    parent.forEach((child, offset) => {
      const s = base + offset, e = s + child.nodeSize;
      if (child.marks.some((m) => m.eq(mark))) { if (s <= from && e >= from) { start = Math.min(start, s); end = Math.max(end, e); } }
    });
    // Extend over adjacent nodes with the same mark
    parent.forEach((child, offset) => {
      const s = base + offset, e = s + child.nodeSize;
      if (child.marks.some((m) => m.eq(mark)) && (e === start || s === end)) { start = Math.min(start, s); end = Math.max(end, e); }
    });
    tr.removeMark(start, end, linkType);
  } else tr.removeMark(from, to, linkType);
  if (dispatch) dispatch(tr);
  return true;
};

export function linkAtSelection(state: EditorState): Mark | null {
  const { $from } = state.selection;
  return linkType.isInSet(state.storedMarks || $from.marks()) || null;
}

// ---------------------------------------------------------------------------
// Selection helpers

export const selectParagraph: Command = (state, dispatch) => {
  const p = selectionParagraph(state);
  if (!p) return false;
  if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, p.pos + 1, p.pos + p.node.nodeSize - 1)));
  return true;
};

export function isInTable(state: EditorState): boolean {
  const $p = state.selection.$from;
  for (let d = $p.depth; d > 0; d--) if ($p.node(d).type === schema.nodes.table_cell) return true;
  return state.selection instanceof CellSelection;
}

/** Word count of the document body. */
export function countWords(doc: PMNode): { words: number; chars: number } {
  let words = 0, chars = 0;
  doc.descendants((node) => {
    if (node.isText) {
      const t = node.text || "";
      chars += t.length;
      const m = t.match(/[^\s ]+/g);
      if (m) words += m.length;
    }
    return true;
  });
  return { words, chars };
}

export function docHasContent(doc: PMNode): boolean {
  return doc.textContent.trim().length > 0 || doc.childCount > 1;
}

export type { Transaction, Fragment };
