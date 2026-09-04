// Find & replace: matches are computed per text block (so a word split across
// runs still matches) and highlighted with inline decorations. Supports plain
// text, whole word, regular expressions, a selection scope, Turkish-friendly
// case folding and case-preserving replacement.
import { Plugin, PluginKey, EditorState, Transaction, TextSelection } from "prosemirror-state";
import { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import { Node as PMNode } from "prosemirror-model";

export interface FindOptions { caseSensitive: boolean; wholeWord: boolean; regex: boolean; preserveCase: boolean; }
export interface Match { from: number; to: number; text: string; groups: string[]; }
export interface FindState {
  query: string;
  opts: FindOptions;
  scope: { from: number; to: number } | null;
  matches: Match[];
  current: number; // index into matches, -1 none
  active: boolean;
  decos: DecorationSet;
  error: string | null;
}

export const findKey = new PluginKey<FindState>("find");
export const DEFAULT_FIND_OPTIONS: FindOptions = { caseSensitive: false, wholeWord: false, regex: false, preserveCase: false };

function escapeRe(s: string) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

/** Length-preserving case fold: lower-cases single characters and unifies the Turkish i variants. */
export function foldChar(ch: string): string {
  const c = ch.codePointAt(0)!;
  if (c === 0x130 || c === 0x131 || c === 0x49) return "i"; // İ ı I -> i
  if (c < 128) return c >= 65 && c <= 90 ? String.fromCharCode(c + 32) : ch;
  const l = ch.toLowerCase();
  return l.length === ch.length ? l : ch;
}
export function foldString(s: string): string {
  let out = "";
  for (const ch of s) out += foldChar(ch);
  return out;
}

/** Build the search regex. Throws on an invalid regular expression. */
export function buildRegex(query: string, opts: FindOptions): RegExp {
  let src = opts.regex ? query : escapeRe(opts.caseSensitive ? query : foldString(query));
  if (opts.wholeWord) src = "(?<![\\p{L}\\p{N}_])(?:" + src + ")(?![\\p{L}\\p{N}_])";
  const flags = "gu" + (opts.regex && !opts.caseSensitive ? "i" : "");
  return new RegExp(src, flags);
}

export function findMatches(doc: PMNode, query: string, opts: FindOptions, scope: { from: number; to: number } | null): { matches: Match[]; error: string | null } {
  const out: Match[] = [];
  if (!query) return { matches: out, error: null };
  let re: RegExp;
  try { re = buildRegex(query, opts); } catch (e) { return { matches: out, error: (e as Error).message.replace(/^Invalid regular expression: /, "") }; }
  const fold = !opts.regex && !opts.caseSensitive;
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    if (scope && (pos + node.nodeSize <= scope.from || pos >= scope.to)) return true;
    let text = "";
    const map: number[] = [];
    node.forEach((child, offset) => {
      const base = pos + 1 + offset;
      if (child.isText) {
        const t = child.text || "";
        for (let i = 0; i < t.length; i++) map.push(base + i);
        text += fold ? foldString(t) : t;
      } else if (child.type.name === "tab") { map.push(base); text += "\t"; }
      else if (child.type.name === "hard_break") { map.push(base); text += "\n"; }
      else { map.push(base); text += "￼"; }
    });
    if (!text) return false;
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    let guard = 0;
    while ((m = re.exec(text)) && guard++ < 100000) {
      if (m[0].length === 0) { re.lastIndex++; continue; }
      const s = m.index, e = m.index + m[0].length;
      const from = map[s], to = map[e - 1] + 1;
      if (scope && (from < scope.from || to > scope.to)) continue;
      out.push({ from, to, text: node.textBetween(s, e, "\n", "￼") || m[0], groups: m.slice(1) });
    }
    return false;
  });
  return { matches: out, error: null };
}

/** Expand $1..$9 / $& in regex mode and adapt letter case when requested. */
export function expandReplacement(m: Match, replacement: string, opts: FindOptions, original: string): string {
  let r = replacement;
  if (opts.regex) {
    r = r.replace(/\$(\d|&|\$)/g, (_s, g) => g === "$" ? "$" : g === "&" ? original : (m.groups[parseInt(g, 10) - 1] ?? ""));
  }
  if (opts.preserveCase && !opts.caseSensitive) {
    const letters = original.replace(/[^\p{L}]/gu, "");
    if (letters && letters === letters.toUpperCase() && letters.length > 1) r = r.toUpperCase();
    else if (letters && letters[0] === letters[0].toUpperCase() && letters.slice(1) === letters.slice(1).toLowerCase()) r = r.charAt(0).toUpperCase() + r.slice(1);
  }
  return r;
}

function decorate(doc: PMNode, matches: Match[], current: number): DecorationSet {
  if (!matches.length) return DecorationSet.empty;
  return DecorationSet.create(doc, matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === current ? "om-find cur" : "om-find" })));
}

export function findPlugin(): Plugin<FindState> {
  return new Plugin<FindState>({
    key: findKey,
    state: {
      init: () => ({ query: "", opts: { ...DEFAULT_FIND_OPTIONS }, scope: null, matches: [], current: -1, active: false, decos: DecorationSet.empty, error: null }),
      apply(tr, prev) {
        const meta = tr.getMeta(findKey) as Partial<FindState> | undefined;
        let next: FindState = meta ? { ...prev, ...meta } : prev;
        if (tr.docChanged && next.scope) next = { ...next, scope: { from: tr.mapping.map(next.scope.from, -1), to: tr.mapping.map(next.scope.to, 1) } };
        if (!next.active) {
          if (prev.active || prev.matches.length) return { ...next, matches: [], current: -1, decos: DecorationSet.empty, error: null };
          return next;
        }
        const recompute = !!meta && ("query" in meta || "opts" in meta || "active" in meta || "scope" in meta) || tr.docChanged;
        if (recompute) {
          const { matches, error } = findMatches(tr.doc, next.query, next.opts, next.scope);
          let current: number;
          if (meta && "current" in meta && meta.current !== undefined) current = meta.current;
          else if (!matches.length) current = -1;
          else if (tr.docChanged && !meta) current = Math.min(Math.max(0, prev.current), matches.length - 1);
          else {
            const selFrom = tr.selection.from;
            let best = matches.findIndex((m) => m.from >= selFrom);
            if (best < 0) best = 0;
            current = best;
          }
          return { ...next, matches, current, error, decos: decorate(tr.doc, matches, current) };
        }
        if (meta && "current" in meta) return { ...next, decos: decorate(tr.doc, next.matches, next.current) };
        return next;
      },
    },
    props: {
      decorations(state) { return findKey.getState(state)?.decos || null; },
    },
  });
}

export function getFind(state: EditorState): FindState { return findKey.getState(state)!; }

export function setFindQuery(view: EditorView, patch: Partial<FindState>) {
  view.dispatch(view.state.tr.setMeta(findKey, { active: true, ...patch }).setMeta("addToHistory", false));
}

export function closeFind(view: EditorView) {
  const st = getFind(view.state);
  if (!st.active) return false;
  view.dispatch(view.state.tr.setMeta(findKey, { active: false, scope: null }).setMeta("addToHistory", false));
  return true;
}

/** Move to the next/previous match (wrapping), select it and scroll it into view. Returns whether it wrapped. */
export function findStep(view: EditorView, dir: 1 | -1): { ok: boolean; wrapped: boolean } {
  const st = getFind(view.state);
  if (!st.matches.length) return { ok: false, wrapped: false };
  let idx: number;
  let wrapped = false;
  const sel = view.state.selection;
  const cur = st.current >= 0 ? st.matches[st.current] : null;
  const selOnCurrent = !!cur && sel.from === cur.from && sel.to === cur.to;
  if (selOnCurrent) {
    idx = st.current + dir;
    if (idx >= st.matches.length) { idx = 0; wrapped = true; }
    if (idx < 0) { idx = st.matches.length - 1; wrapped = true; }
  } else {
    const p = dir > 0 ? sel.to : sel.from;
    idx = dir > 0 ? st.matches.findIndex((m) => m.from >= p) : st.matches.map((m) => m.to <= p).lastIndexOf(true);
    if (idx < 0) { idx = dir > 0 ? 0 : st.matches.length - 1; wrapped = true; }
  }
  const m = st.matches[idx];
  view.dispatch(view.state.tr.setMeta(findKey, { current: idx }).setMeta("addToHistory", false).setSelection(TextSelection.create(view.state.doc, m.from, m.to)).scrollIntoView());
  return { ok: true, wrapped };
}

/** Replace the current match (or jump to the next one if none is selected) and advance. */
export function replaceCurrent(view: EditorView, replacement: string): boolean {
  const st = getFind(view.state);
  if (!st.matches.length) return false;
  const sel = view.state.selection;
  const cur = st.current >= 0 ? st.matches[st.current] : null;
  if (!cur || sel.from !== cur.from || sel.to !== cur.to) { findStep(view, 1); return true; }
  const text = expandReplacement(cur, replacement, st.opts, view.state.doc.textBetween(cur.from, cur.to, "\n"));
  const marks = view.state.doc.resolve(cur.from).marks();
  const tr: Transaction = view.state.tr.replaceWith(cur.from, cur.to, text ? view.state.schema.text(text, marks) : []);
  view.dispatch(tr);
  const after = getFind(view.state);
  if (after.matches.length) {
    let next = after.matches.findIndex((m) => m.from >= cur.from + text.length);
    if (next < 0) next = 0;
    const m = after.matches[next];
    view.dispatch(view.state.tr.setMeta(findKey, { current: next }).setMeta("addToHistory", false).setSelection(TextSelection.create(view.state.doc, m.from, m.to)).scrollIntoView());
  }
  return true;
}

export function replaceAll(view: EditorView, replacement: string): number {
  const st = getFind(view.state);
  if (!st.matches.length) return 0;
  const tr = view.state.tr;
  for (let i = st.matches.length - 1; i >= 0; i--) {
    const m = st.matches[i];
    const text = expandReplacement(m, replacement, st.opts, view.state.doc.textBetween(m.from, m.to, "\n"));
    const marks = view.state.doc.resolve(m.from).marks();
    tr.replaceWith(m.from, m.to, text ? view.state.schema.text(text, marks) : []);
  }
  view.dispatch(tr);
  return st.matches.length;
}

// ---------------------------------------------------------------------------
// Plain-text search used by the Markdown source editor (textarea).

export function textareaFind(ta: HTMLTextAreaElement, query: string, opts: FindOptions, dir: 1 | -1): { count: number; index: number; wrapped: boolean; error: string | null } {
  if (!query) return { count: 0, index: -1, wrapped: false, error: null };
  let re: RegExp;
  try { re = buildRegex(query, opts); } catch (e) { return { count: 0, index: -1, wrapped: false, error: (e as Error).message }; }
  const fold = !opts.regex && !opts.caseSensitive;
  const hay = fold ? foldString(ta.value) : ta.value;
  const all: { s: number; e: number }[] = [];
  let m: RegExpExecArray | null;
  re.lastIndex = 0;
  while ((m = re.exec(hay))) { if (!m[0].length) { re.lastIndex++; continue; } all.push({ s: m.index, e: m.index + m[0].length }); }
  if (!all.length) return { count: 0, index: -1, wrapped: false, error: null };
  const selS = ta.selectionStart, selE = ta.selectionEnd;
  const onCurrent = all.findIndex((x) => x.s === selS && x.e === selE);
  let idx: number, wrapped = false;
  if (onCurrent >= 0) { idx = onCurrent + dir; if (idx >= all.length) { idx = 0; wrapped = true; } if (idx < 0) { idx = all.length - 1; wrapped = true; } }
  else { idx = dir > 0 ? all.findIndex((x) => x.s >= selE) : all.map((x) => x.e <= selS).lastIndexOf(true); if (idx < 0) { idx = dir > 0 ? 0 : all.length - 1; wrapped = true; } }
  const t = all[idx];
  ta.focus();
  ta.setSelectionRange(t.s, t.e);
  // scroll the selection into view
  const before = ta.value.slice(0, t.s);
  const line = before.split("\n").length - 1;
  const lh = parseFloat(getComputedStyle(ta).lineHeight) || 20;
  ta.scrollTop = Math.max(0, line * lh - ta.clientHeight / 2);
  return { count: all.length, index: idx, wrapped, error: null };
}

export function textareaReplace(ta: HTMLTextAreaElement, query: string, replacement: string, opts: FindOptions, all: boolean): number {
  let re: RegExp;
  try { re = buildRegex(query, opts); } catch { return 0; }
  const fold = !opts.regex && !opts.caseSensitive;
  if (all) {
    const hay = fold ? foldString(ta.value) : ta.value;
    let out = "", last = 0, n = 0, m: RegExpExecArray | null;
    re.lastIndex = 0;
    while ((m = re.exec(hay))) {
      if (!m[0].length) { re.lastIndex++; continue; }
      const orig = ta.value.slice(m.index, m.index + m[0].length);
      out += ta.value.slice(last, m.index) + expandReplacement({ from: 0, to: 0, text: orig, groups: m.slice(1) }, replacement, opts, orig);
      last = m.index + m[0].length; n++;
    }
    out += ta.value.slice(last);
    if (n) { ta.value = out; ta.dispatchEvent(new Event("input", { bubbles: true })); }
    return n;
  }
  const s = ta.selectionStart, e = ta.selectionEnd;
  if (e <= s) return 0;
  const orig = ta.value.slice(s, e);
  const hay = fold ? foldString(orig) : orig;
  re.lastIndex = 0;
  const m = re.exec(hay);
  if (!m || m.index !== 0 || m[0].length !== hay.length) return 0;
  const rep = expandReplacement({ from: s, to: e, text: orig, groups: m.slice(1) }, replacement, opts, orig);
  ta.setRangeText(rep, s, e, "end");
  ta.dispatchEvent(new Event("input", { bubbles: true }));
  return 1;
}
