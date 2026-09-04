// Right-click menu: context aware (text, link, image, table cell).
import { EditorView } from "prosemirror-view";
import { NodeSelection } from "prosemirror-state";
import { showMenu, MenuItem } from "../ui/widgets";
import * as C from "./commands";
import * as T from "./tables";
import { AppActions, Shortcut, keyLabel } from "./keymap";
import { schema } from "../schema";
import { undo, redo } from "prosemirror-history";
import { undoDepth, redoDepth } from "prosemirror-history";
import { requestPlainPaste, insertPlainText } from "./editor";
import { openExternal } from "../files";

export interface ContextActions extends AppActions {
  paragraphDialog(): void;
  editLink(): void;
  cellShading(): void;
  insertImageFromBytes(bytes: Uint8Array, ext: string): void;
}

export async function pasteFromClipboard(view: EditorView, plain: boolean, insertImage?: (bytes: Uint8Array, ext: string) => void) {
  try {
    if (!plain && navigator.clipboard && (navigator.clipboard as any).read) {
      const items: ClipboardItem[] = await (navigator.clipboard as any).read();
      for (const item of items) {
        if (item.types.includes("text/html")) {
          const html = await (await item.getType("text/html")).text();
          view.focus();
          view.pasteHTML(html);
          return;
        }
        const imgType = item.types.find((t) => t.startsWith("image/"));
        if (imgType && insertImage) {
          const blob = await item.getType(imgType);
          insertImage(new Uint8Array(await blob.arrayBuffer()), imgType.split("/")[1].replace("jpeg", "jpeg"));
          return;
        }
      }
    }
    const text = await navigator.clipboard.readText();
    view.focus();
    if (text) insertPlainText(view, text);
  } catch {
    view.focus();
    if (plain) requestPlainPaste();
    document.execCommand("paste");
  }
}

export function showContextMenu(view: EditorView, event: MouseEvent, app: ContextActions, shortcuts: Shortcut[]): boolean {
  event.preventDefault();
  const state = view.state;
  const key = (id: string) => { const s = shortcuts.find((x) => x.id === id); return s ? keyLabel(s.keys[0]) : undefined; };
  const run = (cmd: (s: any, d: any, v: any) => boolean) => { view.focus(); cmd(view.state, view.dispatch, view); };
  const sel = state.selection;
  const hasSel = !sel.empty;
  const items: MenuItem[] = [];
  const onImage = sel instanceof NodeSelection && sel.node.type === schema.nodes.image;
  const link = C.linkAtSelection(state);
  const inTable = C.isInTable(state);
  const rp = C.selectionRunProps(state);
  const pp = C.selectionParaProps(state).eff;

  if (onImage) {
    const node = (sel as NodeSelection).node;
    const pos = sel.from;
    const setWrap = (float: string | null) => {
      const wrap = float ? { type: float === "center" ? "topAndBottom" : "square", float } : null;
      view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, kind: float ? "anchor" : "inline", wrap }));
    };
    items.push(
      { label: "Cut", key: key("cut") || "Ctrl+X", action: () => { view.focus(); document.execCommand("cut"); } },
      { label: "Copy", key: "Ctrl+C", action: () => { view.focus(); document.execCommand("copy"); } },
      { sep: true },
      { label: "Original size", disabled: !node.attrs.origW, action: () => view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, w: node.attrs.origW, h: node.attrs.origH })) },
      { label: "Fit to page width", action: () => {
        const s = view.state.doc.attrs.sect; const cw = (s.pgW - s.marL - s.marR) / 15;
        const ratio = node.attrs.h / node.attrs.w;
        view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, w: Math.round(cw), h: Math.round(cw * ratio) }));
      } },
      { label: "Text wrapping", submenu: [
        { label: "In line with text", checked: node.attrs.kind !== "anchor" || !node.attrs.wrap?.float, action: () => setWrap(null) },
        { label: "Left of text", checked: node.attrs.wrap?.float === "left", action: () => setWrap("left") },
        { label: "Right of text", checked: node.attrs.wrap?.float === "right", action: () => setWrap("right") },
        { label: "Centered, text above and below", checked: node.attrs.wrap?.float === "center", action: () => setWrap("center") },
      ] },
      { sep: true },
      { label: "Delete image", key: "Del", action: () => view.dispatch(view.state.tr.deleteSelection()) },
    );
    showMenu({ x: event.clientX, y: event.clientY }, items);
    return true;
  }

  items.push(
    { label: "Undo", key: key("undo"), disabled: undoDepth(state) === 0, action: () => run(undo) },
    { label: "Redo", key: key("redo"), disabled: redoDepth(state) === 0, action: () => run(redo) },
    { sep: true },
    { label: "Cut", key: "Ctrl+X", disabled: !hasSel, action: () => { view.focus(); document.execCommand("cut"); } },
    { label: "Copy", key: "Ctrl+C", disabled: !hasSel, action: () => { view.focus(); document.execCommand("copy"); } },
    { label: "Paste", key: "Ctrl+V", action: () => pasteFromClipboard(view, false, app.insertImageFromBytes) },
    { label: "Paste without formatting", key: key("pasteplain"), action: () => pasteFromClipboard(view, true) },
    { sep: true },
  );
  if (link) {
    items.push(
      { label: "Open link", action: () => { const h = link.attrs.href; if (h && !h.startsWith("#")) openExternal(h); } },
      { label: "Edit link…", key: key("link"), action: () => app.editLink() },
      { label: "Remove link", action: () => run(C.unsetLink) },
      { sep: true },
    );
  } else {
    items.push({ label: hasSel ? "Link…" : "Insert link…", key: key("link"), action: () => app.insertLink() });
  }
  items.push(
    { label: "Bold", key: key("bold"), checked: !!rp.b, action: () => run(C.toggleBold) },
    { label: "Italic", key: key("italic"), checked: !!rp.i, action: () => run(C.toggleItalic) },
    { label: "Underline", key: key("underline"), checked: !!rp.u && rp.u !== "none", action: () => run(C.toggleUnderline) },
    { label: "Font…", key: key("fontbox"), action: () => app.focusFontFamily() },
    { label: "Paragraph…", action: () => app.paragraphDialog() },
    { label: "Align", submenu: [
      { label: "Left", key: key("left"), checked: !pp.jc || pp.jc === "left" || pp.jc === "start", action: () => run(C.setAlign("left")) },
      { label: "Center", key: key("center"), checked: pp.jc === "center", action: () => run(C.setAlign("center")) },
      { label: "Right", key: key("right"), checked: pp.jc === "right" || pp.jc === "end", action: () => run(C.setAlign("right")) },
      { label: "Justify", key: key("justify"), checked: pp.jc === "both", action: () => run(C.setAlign("both")) },
    ] },
    { label: "List", submenu: [
      { label: "Bullets", key: key("bullets"), action: () => run(C.toggleList("bullet")) },
      { label: "Numbering", key: key("numbers"), action: () => run(C.toggleList("decimal")) },
      { sep: true },
      { label: "Increase indent", key: key("indent"), action: () => run(C.indentParagraphs(1)) },
      { label: "Decrease indent", key: key("outdent"), action: () => run(C.indentParagraphs(-1)) },
    ] },
    { label: "Clear formatting", key: key("clear"), action: () => run(C.clearFormatting) },
  );
  if (inTable) {
    items.push({ sep: true }, { label: "Table", submenu: tableMenu(view, app) });
  } else {
    items.push({ sep: true }, { label: "Insert table…", key: key("table"), action: () => app.insertTable() }, { label: "Insert image…", key: key("image"), action: () => app.insertImage() });
  }
  items.push({ sep: true }, { label: "Select paragraph", action: () => run(C.selectParagraph) }, { label: "Select all", key: key("selectall"), action: () => app.selectAllCmd() });
  showMenu({ x: event.clientX, y: event.clientY }, items);
  return true;
}

export function tableMenu(view: EditorView, app: ContextActions): MenuItem[] {
  const run = (cmd: (s: any, d: any, v?: any) => boolean) => { view.focus(); cmd(view.state, view.dispatch, view); };
  return [
    { label: "Insert row above", action: () => run(T.addRowBefore) },
    { label: "Insert row below", action: () => run(T.addRowAfter) },
    { label: "Insert column left", action: () => run(T.addColumnBefore) },
    { label: "Insert column right", action: () => run(T.addColumnAfter) },
    { sep: true },
    { label: "Delete row", action: () => run(T.deleteRow) },
    { label: "Delete column", action: () => run(T.deleteColumn) },
    { label: "Delete table", action: () => run(T.deleteTable) },
    { sep: true },
    { label: "Merge cells", action: () => run(T.mergeCells) },
    { label: "Split cell", action: () => run(T.splitCell) },
    { sep: true },
    { label: "Cell shading…", action: () => app.cellShading() },
    { label: "Vertical alignment", submenu: [
      { label: "Top", action: () => run(T.setCellVAlign("top")) },
      { label: "Middle", action: () => run(T.setCellVAlign("center")) },
      { label: "Bottom", action: () => run(T.setCellVAlign("bottom")) },
    ] },
    { label: "Borders", submenu: [
      { label: "All borders", action: () => run(T.setTableBorders(true)) },
      { label: "No borders", action: () => run(T.setTableBorders(false)) },
    ] },
    { label: "Table alignment", submenu: [
      { label: "Left", action: () => run(T.setTableAlign("left")) },
      { label: "Center", action: () => run(T.setTableAlign("center")) },
      { label: "Right", action: () => run(T.setTableAlign("right")) },
    ] },
    { sep: true },
    { label: "Select row", action: () => run(T.selectRow) },
    { label: "Select column", action: () => run(T.selectColumn) },
    { label: "Select table", action: () => run(T.selectTable) },
  ];
}
