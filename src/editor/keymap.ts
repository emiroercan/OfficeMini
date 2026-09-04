// Single source of truth for keyboard shortcuts: builds the ProseMirror keymap,
// the menu/tooltip labels and the shortcut help dialog.
import { keymap } from "prosemirror-keymap";
import { Command, EditorState, Transaction } from "prosemirror-state";
import { undo, redo } from "prosemirror-history";
import { baseKeymap, selectAll, chainCommands, deleteSelection, joinBackward, selectNodeBackward, joinForward, selectNodeForward, exitCode } from "prosemirror-commands";
import { goToNextCell, deleteTable } from "prosemirror-tables";
import { EditorView } from "prosemirror-view";
import * as C from "./commands";
import { schema } from "../schema";

export const isMac = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform);

/** Application-level actions the editor keymap needs (implemented by main.ts). */
export interface AppActions {
  newDocument(): void;
  openFile(): void;
  save(): void;
  saveAs(): void;
  print(): void;
  closeWindow(): void;
  find(): void;
  replace(): void;
  findNext(): void;
  findPrev(): void;
  closeFind(): boolean;
  zoomIn(): void;
  zoomOut(): void;
  zoomReset(): void;
  toggleMarks(): void;
  toggleFullscreen(): void;
  showShortcuts(): void;
  focusFontFamily(): void;
  focusFontSize(): void;
  insertLink(): void;
  insertImage(): void;
  insertTable(): void;
  goToPage(): void;
  pastePlain(): void;
  selectAllCmd(): void;
  toggleTheme(): void;
  toggleSource(): void;
}

export interface Shortcut {
  id: string;
  keys: string[];        // ProseMirror key names; first one is shown in the UI
  label: string;
  group: "File" | "Edit" | "Find" | "Format" | "Paragraph" | "Insert" | "View" | "Table" | "Navigation";
  command?: Command;     // editor command
  action?: keyof AppActions; // or an app action
  hidden?: boolean;      // not shown in help
}

const app = (fn: keyof AppActions): Command => (_s, _d, _v) => { void fn; return false; };
void app;

function tabCommand(actions: AppActions): Command {
  return (state, dispatch, view) => {
    if (C.isInTable(state) && !inListParagraph(state)) return goToNextCell(1)(state, dispatch);
    if (inListParagraph(state) && atParagraphStart(state)) return C.indentParagraphs(1)(state, dispatch, view);
    if (C.isInTable(state)) return goToNextCell(1)(state, dispatch);
    return C.insertTab(state, dispatch, view);
  };
}
function shiftTabCommand(): Command {
  return (state, dispatch, view) => {
    if (inListParagraph(state)) return C.indentParagraphs(-1)(state, dispatch, view);
    if (C.isInTable(state)) return goToNextCell(-1)(state, dispatch);
    return C.indentParagraphs(-1)(state, dispatch, view);
  };
}
function inListParagraph(state: EditorState): boolean {
  const p = C.selectionParagraph(state);
  return !!p && !!C.selectionParaProps(state).eff.numId;
}
function atParagraphStart(state: EditorState): boolean {
  return state.selection.empty && state.selection.$from.parentOffset === 0;
}

export function shortcuts(actions: AppActions): Shortcut[] {
  const A = (action: keyof AppActions): Command => () => { (actions[action] as () => void)(); return true; };
  return [
    // File
    { id: "new", keys: ["Mod-n"], label: "New document", group: "File", command: A("newDocument") },
    { id: "open", keys: ["Mod-o"], label: "Open…", group: "File", command: A("openFile") },
    { id: "save", keys: ["Mod-s"], label: "Save", group: "File", command: A("save") },
    { id: "saveas", keys: ["Mod-Shift-s"], label: "Save as…", group: "File", command: A("saveAs") },
    { id: "print", keys: ["Mod-p"], label: "Print…", group: "File", command: A("print") },
    { id: "close", keys: ["Mod-w", "Mod-q"], label: "Close window", group: "File", command: A("closeWindow") },
    // Edit
    { id: "undo", keys: ["Mod-z"], label: "Undo", group: "Edit", command: undo },
    { id: "redo", keys: ["Mod-Shift-z", "Mod-y"], label: "Redo", group: "Edit", command: redo },
    { id: "selectall", keys: ["Mod-a"], label: "Select all", group: "Edit", command: selectAll },
    { id: "pasteplain", keys: ["Mod-Shift-v"], label: "Paste without formatting", group: "Edit", command: A("pastePlain"), hidden: true },
    // Find
    { id: "find", keys: ["Mod-f"], label: "Find", group: "Find", command: A("find") },
    { id: "replace", keys: ["Mod-h"], label: "Find and replace", group: "Find", command: A("replace") },
    { id: "findnext", keys: ["F3", "Mod-g"], label: "Find next", group: "Find", command: A("findNext") },
    { id: "findprev", keys: ["Shift-F3", "Mod-Shift-g"], label: "Find previous", group: "Find", command: A("findPrev") },
    { id: "escape", keys: ["Escape"], label: "Close find bar / deselect", group: "Find", command: () => actions.closeFind(), hidden: true },
    // Format
    { id: "bold", keys: ["Mod-b"], label: "Bold", group: "Format", command: C.toggleBold },
    { id: "italic", keys: ["Mod-i"], label: "Italic", group: "Format", command: C.toggleItalic },
    { id: "underline", keys: ["Mod-u"], label: "Underline", group: "Format", command: C.toggleUnderline },
    { id: "strike", keys: ["Mod-Shift-x", "Alt-Shift-5"], label: "Strikethrough", group: "Format", command: C.toggleStrike },
    { id: "sub", keys: ["Mod-="], label: "Subscript", group: "Format", command: C.toggleSubscript },
    { id: "sup", keys: ["Mod-Shift-="], label: "Superscript", group: "Format", command: C.toggleSuperscript },
    { id: "sizeup", keys: ["Mod-Shift-.", "Mod-Shift->"], label: "Grow font", group: "Format", command: C.fontSizeStep(1) },
    { id: "sizedown", keys: ["Mod-Shift-,", "Mod-Shift-<"], label: "Shrink font", group: "Format", command: C.fontSizeStep(-1) },
    { id: "sizeup1", keys: ["Mod-]"], label: "Grow font 1 pt", group: "Format", command: C.fontSizeStep(1, true) },
    { id: "sizedown1", keys: ["Mod-["], label: "Shrink font 1 pt", group: "Format", command: C.fontSizeStep(-1, true) },
    { id: "clear", keys: ["Mod-Space"], label: "Clear character formatting", group: "Format", command: C.clearFormatting },
    { id: "fontbox", keys: ["Mod-Shift-f", "Mod-d"], label: "Font name box", group: "Format", command: A("focusFontFamily") },
    { id: "sizebox", keys: ["Mod-Shift-p"], label: "Font size box", group: "Format", command: A("focusFontSize") },
    // Paragraph
    { id: "left", keys: ["Mod-l"], label: "Align left", group: "Paragraph", command: C.setAlign("left") },
    { id: "center", keys: ["Mod-e"], label: "Center", group: "Paragraph", command: C.setAlign("center") },
    { id: "right", keys: ["Mod-r"], label: "Align right", group: "Paragraph", command: C.setAlign("right") },
    { id: "justify", keys: ["Mod-j"], label: "Justify", group: "Paragraph", command: C.setAlign("both") },
    { id: "bullets", keys: ["Mod-Shift-l", "Mod-Shift-8"], label: "Bulleted list", group: "Paragraph", command: C.toggleList("bullet") },
    { id: "numbers", keys: ["Mod-Shift-7"], label: "Numbered list", group: "Paragraph", command: C.toggleList("decimal") },
    { id: "indent", keys: ["Mod-m"], label: "Increase indent", group: "Paragraph", command: C.indentParagraphs(1) },
    { id: "outdent", keys: ["Mod-Shift-m"], label: "Decrease indent", group: "Paragraph", command: C.indentParagraphs(-1) },
    { id: "single", keys: ["Mod-1"], label: "Single line spacing", group: "Paragraph", command: C.setLineSpacing(1) },
    { id: "onehalf", keys: ["Mod-5"], label: "1.5 line spacing", group: "Paragraph", command: C.setLineSpacing(1.5) },
    { id: "double", keys: ["Mod-2"], label: "Double line spacing", group: "Paragraph", command: C.setLineSpacing(2) },
    { id: "normal", keys: ["Mod-Shift-n"], label: "Normal style", group: "Paragraph", command: C.setParaStyle(null) },
    { id: "h1", keys: ["Mod-Alt-1"], label: "Heading 1", group: "Paragraph", command: C.setParaStyle("Heading1") },
    { id: "h2", keys: ["Mod-Alt-2"], label: "Heading 2", group: "Paragraph", command: C.setParaStyle("Heading2") },
    { id: "h3", keys: ["Mod-Alt-3"], label: "Heading 3", group: "Paragraph", command: C.setParaStyle("Heading3") },
    // Insert
    { id: "linebreak", keys: ["Shift-Enter"], label: "Line break", group: "Insert", command: C.insertLineBreak },
    { id: "pagebreak", keys: ["Mod-Enter"], label: "Page break", group: "Insert", command: C.insertPageBreak },
    { id: "link", keys: ["Mod-k"], label: "Insert link…", group: "Insert", command: A("insertLink") },
    { id: "image", keys: ["Mod-Shift-i"], label: "Insert image…", group: "Insert", command: A("insertImage") },
    { id: "table", keys: ["Mod-Shift-t"], label: "Insert table…", group: "Insert", command: A("insertTable") },
    { id: "nbsp", keys: ["Mod-Shift-Space"], label: "Non-breaking space", group: "Insert", command: C.insertTextCmd(String.fromCharCode(0xa0)) },
    { id: "nbhyphen", keys: ["Mod-Shift--"], label: "Non-breaking hyphen", group: "Insert", command: C.insertTextCmd(String.fromCharCode(0x2011)) },
    // View
    { id: "zoomin", keys: ["Mod-=", "Mod-+"], label: "Zoom in", group: "View", command: A("zoomIn"), hidden: true },
    { id: "zoominx", keys: ["Mod-Shift-+", "Mod-Shift-="], label: "Zoom in", group: "View", command: A("zoomIn") },
    { id: "zoomout", keys: ["Mod--"], label: "Zoom out", group: "View", command: A("zoomOut") },
    { id: "zoom0", keys: ["Mod-0"], label: "Zoom 100%", group: "View", command: A("zoomReset") },
    { id: "marks", keys: ["Mod-Shift-*", "Mod-Shift-8"], label: "Show formatting marks", group: "View", command: A("toggleMarks"), hidden: true },
    { id: "fullscreen", keys: ["F11"], label: "Full screen", group: "View", command: A("toggleFullscreen") },
    { id: "theme", keys: ["Mod-Shift-d"], label: "Dark mode on/off", group: "View", command: A("toggleTheme") },
    { id: "source", keys: ["Mod-Alt-m"], label: "Markdown source / rendered view", group: "View", command: A("toggleSource") },
    { id: "help", keys: ["Mod-/", "F1"], label: "Keyboard shortcuts", group: "View", command: A("showShortcuts") },
    { id: "goto", keys: ["Mod-Shift-g"], label: "Go to page…", group: "Navigation", command: A("goToPage"), hidden: true },
    // Table
    { id: "nextcell", keys: ["Tab"], label: "Next cell / indent list / insert tab", group: "Table", command: tabCommand(actions) },
    { id: "prevcell", keys: ["Shift-Tab"], label: "Previous cell / outdent list", group: "Table", command: shiftTabCommand() },
    // Editing keys
    { id: "enter", keys: ["Enter"], label: "New paragraph", group: "Edit", command: C.splitParagraph, hidden: true },
    { id: "backspace", keys: ["Backspace", "Mod-Backspace"], label: "Backspace", group: "Edit", hidden: true,
      command: chainCommands(C.backspaceListStart, deleteSelection, joinBackward, selectNodeBackward) },
    { id: "delete", keys: ["Delete", "Mod-Delete"], label: "Delete", group: "Edit", hidden: true,
      command: chainCommands(deleteSelection, joinForward, selectNodeForward) },
    { id: "exitcode", keys: ["Mod-Shift-Enter"], label: "Exit block", group: "Edit", hidden: true, command: exitCode },
  ];
}

/** Notes shown in the help dialog that are not keymap entries (native / mouse behaviour). */
export const EXTRA_HELP: { group: string; rows: [string, string][] }[] = [
  { group: "Navigation", rows: [
    ["Home / End", "Start / end of line"], ["Ctrl+Home / Ctrl+End", "Start / end of document"],
    ["Ctrl+← / Ctrl+→", "Previous / next word"], ["Ctrl+↑ / Ctrl+↓", "Previous / next paragraph"],
    ["Page Up / Page Down", "Scroll one screen"], ["Shift + any of the above", "Extend selection"],
  ] },
  { group: "Mouse", rows: [
    ["Click", "Place cursor"], ["Double-click", "Select word"], ["Triple-click", "Select paragraph"],
    ["Shift+click", "Extend selection"], ["Drag", "Select text / move image"], ["Ctrl+click on link", "Open link"],
    ["Right-click", "Context menu (formatting, table, image, link)"], ["Middle-click", "Paste selection (Linux) / auto-scroll (Windows)"],
    ["Ctrl+wheel", "Zoom"], ["Shift+wheel", "Horizontal scroll"], ["Drag image corner", "Resize image"],
    ["Drag cell border", "Resize table column"], ["Drag across cells", "Select cells"],
  ] },
];

/** Human readable key label, e.g. "Ctrl+Shift+S". */
export function keyLabel(key: string): string {
  // A trailing "-" key is written as "Mod--" / "Mod-Shift--".
  let parts: string[];
  if (key.endsWith("--")) parts = [...key.slice(0, -2).split("-"), "-"];
  else parts = key.split("-");
  const shifted = parts.includes("Shift");
  return parts.map((part) => {
    switch (part) {
      case "Mod": return isMac ? "⌘" : "Ctrl";
      case "Alt": return isMac ? "⌥" : "Alt";
      case "Shift": return isMac ? "⇧" : "Shift";
      case "Ctrl": return isMac ? "⌃" : "Ctrl";
      case "Escape": return "Esc";
      case ".": return shifted ? ">" : ".";
      case ",": return shifted ? "<" : ",";
      case "=": return shifted ? "+" : "=";
      default: return part.length === 1 ? part.toUpperCase() : part;
    }
  }).join("+");
}

export function shortcutLabelFor(list: Shortcut[], id: string): string {
  const s = list.find((x) => x.id === id);
  return s ? keyLabel(s.keys[0]) : "";
}

export function buildKeymap(list: Shortcut[]) {
  const map: Record<string, Command> = {};
  for (const s of list) {
    if (!s.command) continue;
    for (const k of s.keys) if (!map[k]) map[k] = s.command;
  }
  return [keymap(map), keymap(baseKeymap)];
}

export type { Command, Transaction, EditorView };
void deleteTable;
void schema;
