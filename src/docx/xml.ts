// Minimal XML helpers over the browser DOMParser. DOCX uses stable prefixes,
// but we always match on namespace + localName to be safe.

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  asvg: "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
  xml: "http://www.w3.org/XML/1998/namespace",
};

export const NO_BREAK_HYPHEN = String.fromCharCode(0x2011);
export const SOFT_HYPHEN = String.fromCharCode(0xad);

const parser = new DOMParser();
const serializer = new XMLSerializer();

export function parseXml(text: string): Document {
  const doc = parser.parseFromString(text, "application/xml");
  const err = doc.getElementsByTagName("parsererror")[0];
  if (err) throw new Error("XML parse error: " + (err.textContent || "").slice(0, 200));
  return doc;
}

export function isEl(n: Node | null | undefined, ns: string, name: string): n is Element {
  return !!n && n.nodeType === 1 && (n as Element).localName === name && (n as Element).namespaceURI === ns;
}

export function child(el: Element | null | undefined, ns: string, name: string): Element | null {
  if (!el) return null;
  for (let n = el.firstChild; n; n = n.nextSibling) if (isEl(n, ns, name)) return n;
  return null;
}

export function children(el: Element | null | undefined, ns?: string, name?: string): Element[] {
  const out: Element[] = [];
  if (!el) return out;
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const e = n as Element;
    if (ns && e.namespaceURI !== ns) continue;
    if (name && e.localName !== name) continue;
    out.push(e);
  }
  return out;
}

export function descendant(el: Element | null | undefined, ns: string, name: string): Element | null {
  if (!el) return null;
  const list = el.getElementsByTagNameNS(ns, name);
  return list.length ? list[0] : null;
}

export function descendants(el: Element | null | undefined, ns: string, name: string): Element[] {
  if (!el) return [];
  return Array.from(el.getElementsByTagNameNS(ns, name));
}

export function attr(el: Element | null | undefined, ns: string | null, name: string): string | null {
  if (!el) return null;
  const v = ns ? el.getAttributeNS(ns, name) : el.getAttribute(name);
  if (v !== null) return v;
  // Some producers write un-namespaced attributes (rare) - fall back by local name.
  for (const a of Array.from(el.attributes)) if (a.localName === name) return a.value;
  return null;
}

export function wattr(el: Element | null | undefined, name: string): string | null {
  return attr(el, NS.w, name);
}

export function wint(el: Element | null | undefined, name: string): number | null {
  const v = wattr(el, name);
  if (v === null) return null;
  const n = parseInt(v, 10);
  return isNaN(n) ? null : n;
}

/** OOXML on/off value: absent element => null; present w/o val => true. */
export function onOff(el: Element | null): boolean | null {
  if (!el) return null;
  const v = wattr(el, "val");
  if (v === null) return true;
  return !(v === "0" || v === "false" || v === "off");
}

/**
 * Serialize an element to a string, dropping namespace declarations that are
 * identical to ones declared on the document root (they are redundant inside the
 * document and would bloat the output).
 */
export function serialize(el: Element, rootDecls?: Map<string, string>): string {
  let s = serializer.serializeToString(el);
  if (rootDecls && rootDecls.size) {
    const end = s.indexOf(">");
    if (end > 0) {
      let open = s.slice(0, end);
      open = open.replace(/\s+xmlns(?::([\w.-]+))?="([^"]*)"/g, (m, prefix, uri) => {
        const key = prefix ? prefix : "";
        return rootDecls.get(key) === uri ? "" : m;
      });
      s = open + s.slice(end);
    }
  }
  return s;
}

export function rootNamespaceDecls(root: Element): Map<string, string> {
  const m = new Map<string, string>();
  for (const a of Array.from(root.attributes)) {
    if (a.name === "xmlns") m.set("", a.value);
    else if (a.name.startsWith("xmlns:")) m.set(a.name.slice(6), a.value);
  }
  return m;
}

export function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Escape text for an XML text node, dropping characters illegal in XML 1.0. */
export function escapeXmlText(s: string): string {
  let needsClean = false;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if ((c < 32 && c !== 9 && c !== 10 && c !== 13) || c === 0xfffe || c === 0xffff) { needsClean = true; break; }
  }
  if (needsClean) {
    let clean = "";
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      if ((c < 32 && c !== 9 && c !== 10 && c !== 13) || c === 0xfffe || c === 0xffff) continue;
      clean += s[i];
    }
    s = clean;
  }
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Text content of all w:t descendants (tabs/breaks as whitespace). */
export function runText(el: Element): string {
  let out = "";
  const walk = (n: Node) => {
    if (n.nodeType === 1) {
      const e = n as Element;
      if (e.namespaceURI === NS.w) {
        if (e.localName === "t" || e.localName === "delText") { out += e.textContent || ""; return; }
        if (e.localName === "tab") { out += "\t"; return; }
        if (e.localName === "br" || e.localName === "cr") { out += "\n"; return; }
        if (e.localName === "noBreakHyphen") { out += NO_BREAK_HYPHEN; return; }
        if (e.localName === "softHyphen") { out += SOFT_HYPHEN; return; }
      }
      for (let c = n.firstChild; c; c = c.nextSibling) walk(c);
    }
  };
  walk(el);
  return out;
}
