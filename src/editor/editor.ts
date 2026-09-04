// EditorView construction and document-level editor behaviour.
import { EditorState, Plugin, PluginKey, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { Node as PMNode, DOMParser as PMDOMParser, Slice, Fragment } from "prosemirror-model";
import { history } from "prosemirror-history";
import { dropCursor } from "prosemirror-dropcursor";
import { gapCursor } from "prosemirror-gapcursor";
import { tableEditing, columnResizing, TableView, updateColumnsOnResize } from "prosemirror-tables";

/** prosemirror-tables' TableView plus the document's table styling from the schema toDOM. */
class OMTableView extends TableView {
  private minW: number;
  constructor(node: PMNode, cellMinWidth: number) {
    super(node, cellMinWidth);
    this.minW = cellMinWidth;
    this.applyStyle(node);
  }
  private applyStyle(node: PMNode) {
    const spec = schema.nodes.table.spec.toDOM!(node) as unknown as any[];
    const attrs = spec[1] || {};
    this.table.className = attrs.class || "om-tbl";
    this.table.setAttribute("style", attrs.style || "");
    updateColumnsOnResize(node, this.colgroup, this.table, this.minW);
    // Fixed-width documents: keep the grid width, not the auto width computed by the resizer.
    if (node.attrs.grid && /table-layout:fixed/.test(attrs.style || "")) { /* width already in style */ }
  }
  update(node: PMNode): boolean {
    const ok = super.update(node);
    if (ok) this.applyStyle(node);
    return ok;
  }
}
import { schema } from "../schema";
import { buildKeymap, shortcuts, AppActions, Shortcut } from "./keymap";
import { numberingPlugin } from "./lists";
import { findPlugin } from "./find";
import { paginationPlugin } from "./pagination";
import { ImageView } from "./nodeviews";
import { normalizeHtml } from "../md/markdown";
import { openExternal } from "../files";

export interface EditorOptions {
  mount: HTMLElement;
  pagesEl: HTMLElement;
  actions: AppActions;
  onUpdate: (view: EditorView) => void;
  onPages: (count: number) => void;
  onContextMenu: (view: EditorView, event: MouseEvent) => boolean;
}

export const pastePlainKey = new PluginKey("pastePlain");
let pastePlainNext = false;
export function requestPlainPaste() { pastePlainNext = true; }

/** Plugin: Ctrl+click opens links; plain click keeps the caret behaviour. */
function linkClickPlugin() {
  return new Plugin({
    props: {
      handleClick(view, _pos, event) {
        if (!(event.ctrlKey || event.metaKey)) return false;
        const a = (event.target as HTMLElement).closest?.("a.om-link") as HTMLAnchorElement | null;
        if (!a) return false;
        const href = a.getAttribute("href") || "";
        if (href && href !== "#") { if (!href.startsWith("#")) openExternal(href); return true; }
        return false;
      },
      handleDOMEvents: {
        // Never let the webview navigate away when a link is clicked.
        click(_view, e) { const a = (e.target as HTMLElement).closest?.("a"); if (a) { e.preventDefault(); } return false; },
        auxclick(_view, e) { const a = (e.target as HTMLElement).closest?.("a"); if (a && e.button === 1) { e.preventDefault(); return true; } return false; },
      },
    },
  });
}

function pastePlugin() {
  return new Plugin({
    props: {
      handlePaste(view, event) {
        const cd = event.clipboardData;
        if (!cd) return false;
        const plain = pastePlainNext;
        pastePlainNext = false;
        // Internal copies carry a ProseMirror slice; let PM handle them unless plain paste was requested.
        const html = cd.getData("text/html");
        if (!plain && html && /data-pm-slice/.test(html)) return false;
        const text = cd.getData("text/plain");
        if (plain || !html) {
          if (!text) return false;
          insertPlainText(view, text);
          return true;
        }
        return false; // fall through to transformPastedHTML
      },
      transformPastedHTML(html) {
        return normalizeHtml(html);
      },
    },
  });
}

export function insertPlainText(view: EditorView, text: string) {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  const marks = view.state.storedMarks || view.state.selection.$from.marks();
  const paraAttrs = view.state.selection.$from.parent.type === schema.nodes.paragraph ? view.state.selection.$from.parent.attrs : {};
  const nodes: PMNode[] = lines.map((line) => {
    const inline: PMNode[] = [];
    line.split("\t").forEach((seg, i) => {
      if (i > 0) inline.push(schema.nodes.tab.create(null, undefined, marks));
      if (seg) inline.push(schema.text(seg, marks));
    });
    return schema.nodes.paragraph.create({ ...paraAttrs, sectPr: null }, inline);
  });
  const slice = new Slice(Fragment.from(nodes), 1, 1);
  view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView());
}

// ---------------------------------------------------------------------------
// Smart quotes and dashes while typing (Word-style AutoCorrect, toggleable).

let smartTyping = true;
export function setSmartTyping(on: boolean) { smartTyping = on; }

function smartTypingPlugin() {
  return new Plugin({
    props: {
      handleTextInput(view, from, to, text) {
        if (!smartTyping || text.length !== 1) return false;
        const { state } = view;
        // Never inside code-like fonts.
        const marks = state.storedMarks || state.selection.$from.marks();
        const rpr = marks.find((m) => m.type === schema.marks.rpr);
        if (rpr && /consolas|courier|mono/i.test(rpr.attrs.props?.font || "")) return false;
        const before = state.doc.textBetween(Math.max(0, from - 2), from, "\n", "￼");
        const prev = before.slice(-1);
        const opening = !prev || /[\s(\[{‘“ \-–—]/.test(prev);
        let repl: string | null = null;
        if (text === '"') repl = opening ? "“" : "”";
        else if (text === "'") repl = opening ? "‘" : "’";
        else if (text === "-" && /\p{L}-$/u.test(before)) { view.dispatch(state.tr.insertText("—", from - 1, to)); return true; }
        if (repl === null) return false;
        view.dispatch(state.tr.insertText(repl, from, to));
        return true;
      },
      // " - " between words becomes an en dash once the next word starts (Word behaviour).
      handleKeyDown(view, event) {
        if (!smartTyping || event.key !== " " || event.ctrlKey || event.altKey || event.metaKey) return false;
        const { state } = view;
        if (!state.selection.empty) return false;
        const pos = state.selection.from;
        const before = state.doc.textBetween(Math.max(0, pos - 3), pos, "\n", "￼");
        if (/\S -$/.test(before) === false) return false;
        // Look back further: only when a word precedes " -".
        const ctx2 = state.doc.textBetween(Math.max(0, pos - 12), pos, "\n", "￼");
        if (!/\p{L}\s-$/u.test(ctx2)) return false;
        view.dispatch(state.tr.insertText("– ", pos - 1, pos));
        return true;
      },
    },
  });
}

export interface EditorHandle {
  view: EditorView;
  shortcuts: Shortcut[];
  setDocument(doc: PMNode): void;
}

export function createEditor(doc: PMNode, opts: EditorOptions): EditorHandle {
  const list = shortcuts(opts.actions);
  const plugins = () => [
    ...buildKeymap(list),
    history({ depth: 500, newGroupDelay: 400 }),
    columnResizing({ cellMinWidth: 16, defaultCellMinWidth: 40, View: OMTableView as any }),
    tableEditing({ allowTableNodeSelection: true }),
    numberingPlugin(),
    findPlugin(),
    paginationPlugin({ pagesEl: opts.pagesEl, onPages: opts.onPages }),
    dropCursor({ width: 2, color: "#2f6fed" }),
    gapCursor(),
    linkClickPlugin(),
    pastePlugin(),
    smartTypingPlugin(),
  ];
  const state = EditorState.create({ doc, plugins: plugins() });
  const view = new EditorView(opts.mount, {
    state,
    attributes: { spellcheck: "true", class: "om-editor" },
    nodeViews: {
      image: (node, view, getPos) => new ImageView(node, view, getPos as () => number | undefined),
    },
    dispatchTransaction(tr) {
      const newState = view.state.apply(tr);
      view.updateState(newState);
      opts.onUpdate(view);
    },
    handleDOMEvents: {
      contextmenu(v, e) { return opts.onContextMenu(v, e as MouseEvent); },
      // Middle click: on Linux the webview pastes the primary selection natively; on Windows
      // Chromium autoscrolls. Both are native; we only make sure focus lands in the editor.
      mousedown(v, e) { if ((e as MouseEvent).button === 1) v.focus(); return false; },
    },
    scrollMargin: 40,
    scrollThreshold: 40,
  });
  return {
    view,
    shortcuts: list,
    setDocument(newDoc: PMNode) {
      const st = EditorState.create({ doc: newDoc, plugins: plugins() });
      view.updateState(st);
      opts.onUpdate(view);
    },
  };
}

/** Force a full re-render (used when the theme changes: node styles are computed in toDOM). */
export function redrawView(view: EditorView) {
  view.setProps({ nodeViews: { image: (node, v, getPos) => new ImageView(node, v, getPos as () => number | undefined) } });
}

/** Parse arbitrary HTML into a document fragment using the schema's parse rules. */
export function parseHtml(html: string): PMNode {
  const dom = new DOMParser().parseFromString(normalizeHtml(html), "text/html");
  return PMDOMParser.fromSchema(schema).parse(dom.body);
}

export function selectPos(view: EditorView, pos: number) {
  const $pos = view.state.doc.resolve(Math.max(0, Math.min(pos, view.state.doc.content.size)));
  const sel = TextSelection.near($pos);
  view.dispatch(view.state.tr.setSelection(sel).scrollIntoView());
  view.focus();
}
