import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";

/** A DOCX package: all parts kept verbatim so untouched parts round-trip byte-exact. */
export class Package {
  parts: Map<string, Uint8Array>;
  /** Entries left compressed in `raw` until first use (large media). */
  deferred = new Set<string>();
  private raw: Uint8Array | null = null;
  constructor(parts: Map<string, Uint8Array>) { this.parts = parts; }

  /** `defer` decides which entries stay compressed until requested (default: none). */
  static fromBytes(bytes: Uint8Array, defer?: (name: string, size: number) => boolean): Package {
    const deferred = new Set<string>();
    const files = unzipSync(bytes, defer ? { filter: (f) => { if (defer(f.name, f.originalSize)) { deferred.add(f.name); return false; } return true; } } : undefined);
    const parts = new Map<string, Uint8Array>();
    for (const [name, data] of Object.entries(files)) {
      if (name.endsWith("/")) continue;
      parts.set(name, data);
    }
    const pkg = new Package(parts);
    pkg.deferred = deferred;
    pkg.raw = deferred.size ? bytes : null;
    return pkg;
  }

  static empty(): Package { return new Package(new Map()); }

  has(name: string) { return this.parts.has(name) || this.deferred.has(name); }
  isDeferred(name: string) { return this.deferred.has(name); }
  get(name: string): Uint8Array | undefined {
    const p = this.parts.get(name);
    if (p) return p;
    if (this.deferred.has(name) && this.raw) {
      const files = unzipSync(this.raw, { filter: (f) => f.name === name });
      const data = files[name];
      if (data) { this.parts.set(name, data); this.deferred.delete(name); return data; }
    }
    return undefined;
  }
  /** Inflate everything that is still deferred (needed before writing the package). */
  materialize() { for (const name of Array.from(this.deferred)) this.get(name); }
  text(name: string): string | undefined {
    const d = this.parts.get(name);
    return d ? strFromU8(d) : undefined;
  }
  setText(name: string, text: string) { this.parts.set(name, strToU8(text)); }
  set(name: string, data: Uint8Array) { this.parts.set(name, data); }
  delete(name: string) { this.parts.delete(name); }

  toBytes(): Uint8Array {
    this.materialize();
    const obj: Record<string, [Uint8Array, { level: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 }]> = {};
    // [Content_Types].xml first for maximum compatibility.
    const names = Array.from(this.parts.keys()).sort((a, b) => {
      if (a === "[Content_Types].xml") return -1;
      if (b === "[Content_Types].xml") return 1;
      return a < b ? -1 : a > b ? 1 : 0;
    });
    for (const name of names) {
      const data = this.parts.get(name)!;
      // Media is already compressed; store it to keep saving fast.
      const level = /\.(png|jpe?g|gif|emf|wmf|bmp|tiff?|webp|mp4|mp3|bin)$/i.test(name) ? 0 : 6;
      obj[name] = [data, { level }];
    }
    return zipSync(obj);
  }
}

export interface Relationship { id: string; type: string; target: string; mode: string | null; }

/** Parse a .rels part. Targets are returned as given (relative to the source part's folder). */
export function parseRels(xml: string | undefined): Map<string, Relationship> {
  const out = new Map<string, Relationship>();
  if (!xml) return out;
  const re = /<Relationship\b([^>]*)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const a = m[1];
    const id = /\bId="([^"]*)"/.exec(a)?.[1];
    const type = /\bType="([^"]*)"/.exec(a)?.[1] || "";
    const target = /\bTarget="([^"]*)"/.exec(a)?.[1] || "";
    const mode = /\bTargetMode="([^"]*)"/.exec(a)?.[1] || null;
    if (id) out.set(id, { id, type, target: decodeXmlAttr(target), mode });
  }
  return out;
}

export function decodeXmlAttr(s: string): string {
  return s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

/** Resolve a relationship target relative to a source part path ("word/document.xml"). */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const base = sourcePart.split("/").slice(0, -1);
  const segs = target.split("/");
  for (const s of segs) {
    if (s === "..") base.pop();
    else if (s === "." || s === "") continue;
    else base.push(s);
  }
  return base.join("/");
}

export function mimeForExt(ext: string): string {
  switch (ext.toLowerCase()) {
    case "png": return "image/png";
    case "jpg": case "jpeg": return "image/jpeg";
    case "gif": return "image/gif";
    case "bmp": return "image/bmp";
    case "svg": return "image/svg+xml";
    case "webp": return "image/webp";
    case "tif": case "tiff": return "image/tiff";
    case "emf": return "image/x-emf";
    case "wmf": return "image/x-wmf";
    default: return "application/octet-stream";
  }
}

export const REL_TYPES = {
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
};
