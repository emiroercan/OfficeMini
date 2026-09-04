// Table insertion and cell property commands on top of prosemirror-tables.
import { Command, EditorState, TextSelection } from "prosemirror-state";
import { Node as PMNode } from "prosemirror-model";
import {
  CellSelection, TableMap, addRowBefore, addRowAfter, addColumnBefore, addColumnAfter, deleteRow, deleteColumn,
  deleteTable, mergeCells, splitCell, selectedRect, findTable,
} from "prosemirror-tables";
import { schema } from "../schema";
import { ctx } from "../docx/styles";
import { CellProps, TableProps } from "../docx/props";
import { twipsToPx } from "../docx/units";

export { addRowBefore, addRowAfter, addColumnBefore, addColumnAfter, deleteRow, deleteColumn, deleteTable, mergeCells, splitCell };

function emptyPara(tblStyle: string | null): PMNode {
  return schema.nodes.paragraph.create({ pPr: null, props: {}, inTable: true, tblStyle, sdt: null, sectPr: null });
}

export function insertTable(rows: number, cols: number): Command {
  return (state, dispatch) => {
    const sect = state.doc.attrs.sect;
    const contentW = Math.max(1440, sect.pgW - sect.marL - sect.marR);
    const colW = Math.floor(contentW / cols);
    const grid = new Array(cols).fill(colW);
    const tblStyle = ctx.styleIdByName("Table Grid") || ctx.styleIdByName("TableGrid");
    const props: TableProps = { tblStyle, width: { w: 0, type: "auto" }, look: "04A0" };
    if (!tblStyle) {
      const b = { val: "single", sz: 4, color: "auto", space: 0 };
      props.borders = { top: b, bottom: b, left: b, right: b, insideH: b, insideV: b };
    }
    const rowNodes: PMNode[] = [];
    for (let r = 0; r < rows; r++) {
      const cells: PMNode[] = [];
      for (let c = 0; c < cols; c++) {
        const cprops: CellProps = { width: { w: colW, type: "dxa" } };
        cells.push(schema.nodes.table_cell.create({ colspan: 1, rowspan: 1, colwidth: [Math.round(twipsToPx(colW))], tcPr: null, props: cprops, vmergeTcPr: [] }, emptyPara(tblStyle)));
      }
      rowNodes.push(schema.nodes.table_row.create({ trPr: null, props: {} }, cells));
    }
    const table = schema.nodes.table.create({ tblPr: null, props, grid, sdt: null }, rowNodes);
    const tr = state.tr;
    const { $from } = state.selection;
    // Insert after the current paragraph (or split it if the cursor is mid-paragraph).
    const para = $from.parent;
    let insertPos: number;
    if (para.type === schema.nodes.paragraph && para.content.size === 0) {
      insertPos = $from.before();
      tr.replaceWith(insertPos, insertPos + para.nodeSize, [table, emptyParaLike(para)]);
    } else {
      insertPos = $from.after();
      tr.insert(insertPos, [table, emptyParaLike(para)]);
    }
    tr.setSelection(TextSelection.create(tr.doc, insertPos + 4));
    if (dispatch) dispatch(tr.scrollIntoView());
    return true;
  };
}

function emptyParaLike(para: PMNode): PMNode {
  if (para.type === schema.nodes.paragraph) return schema.nodes.paragraph.create({ ...para.attrs, sectPr: null, pPr: null, props: { pStyle: (para.attrs.props as any).pStyle || null } });
  return schema.nodes.paragraph.create();
}

/** Apply a patch to the props of every selected cell. */
export function setCellProps(patch: Partial<CellProps>): Command {
  return (state, dispatch) => {
    const sel = state.selection;
    const tr = state.tr;
    const apply = (pos: number, node: PMNode) => {
      const props = { ...(node.attrs.props as CellProps), ...patch };
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, props });
    };
    if (sel instanceof CellSelection) {
      sel.forEachCell((node, pos) => apply(pos, node));
    } else {
      const $p = sel.$from;
      let found = false;
      for (let d = $p.depth; d > 0; d--) {
        if ($p.node(d).type === schema.nodes.table_cell) { apply($p.before(d), $p.node(d)); found = true; break; }
      }
      if (!found) return false;
    }
    if (dispatch) dispatch(tr);
    return true;
  };
}

export function setCellShading(hex: string | null): Command { return setCellProps({ shd: hex ? hex.replace("#", "").toUpperCase() : null }); }
export function setCellVAlign(v: "top" | "center" | "bottom"): Command { return setCellProps({ vAlign: v }); }

/** Toggle all table borders on/off for the table containing the selection. */
export function setTableBorders(on: boolean): Command {
  return (state, dispatch) => {
    const t = findTable(state.selection.$from);
    if (!t) return false;
    const props = { ...(t.node.attrs.props as TableProps) };
    const b = on ? { val: "single", sz: 4, color: "auto", space: 0 } : { val: "nil", sz: 0, color: "auto", space: 0 };
    props.borders = { top: b, bottom: b, left: b, right: b, insideH: b, insideV: b };
    if (dispatch) dispatch(state.tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, props }));
    return true;
  };
}

export function setTableAlign(jc: "left" | "center" | "right"): Command {
  return (state, dispatch) => {
    const t = findTable(state.selection.$from);
    if (!t) return false;
    const props = { ...(t.node.attrs.props as TableProps), jc };
    if (dispatch) dispatch(state.tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, props }));
    return true;
  };
}

export function tableInfo(state: EditorState): { rows: number; cols: number } | null {
  const t = findTable(state.selection.$from);
  if (!t) return null;
  const map = TableMap.get(t.node);
  return { rows: map.height, cols: map.width };
}

export const selectTable: Command = (state, dispatch) => {
  const t = findTable(state.selection.$from);
  if (!t) return false;
  const map = TableMap.get(t.node);
  const first = t.pos + 1 + map.map[0];
  const last = t.pos + 1 + map.map[map.map.length - 1];
  if (dispatch) dispatch(state.tr.setSelection(new CellSelection(state.doc.resolve(first), state.doc.resolve(last))));
  return true;
};

export const selectRow: Command = (state, dispatch) => {
  try {
    const rect = selectedRect(state);
    const map = rect.map;
    const first = rect.tableStart + map.map[rect.top * map.width];
    const last = rect.tableStart + map.map[rect.top * map.width + map.width - 1];
    if (dispatch) dispatch(state.tr.setSelection(new CellSelection(state.doc.resolve(first), state.doc.resolve(last))));
    return true;
  } catch { return false; }
};

export const selectColumn: Command = (state, dispatch) => {
  try {
    const rect = selectedRect(state);
    const map = rect.map;
    const first = rect.tableStart + map.map[rect.left];
    const last = rect.tableStart + map.map[(map.height - 1) * map.width + rect.left];
    if (dispatch) dispatch(state.tr.setSelection(new CellSelection(state.doc.resolve(first), state.doc.resolve(last))));
    return true;
  } catch { return false; }
};

export { CellSelection, findTable };
