// Image helpers: placeholders, and extracting embedded bitmaps from EMF/WMF
// (browsers cannot draw EMF; most EMFs produced by Word for pasted screenshots
// contain a PNG/JPEG or a DIB we can pull out).

export function placeholderImage(w: number, h: number, label: string): string {
  const W = Math.max(24, Math.round(w)), H = Math.max(16, Math.round(h));
  const fs = Math.max(9, Math.min(14, Math.floor(H / 6)));
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">` +
    `<rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" fill="#f3f3f3" stroke="#c9c9c9" stroke-dasharray="4 3"/>` +
    `<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-family="sans-serif" font-size="${fs}" fill="#888">${escapeSvg(label)}</text></svg>`;
  return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
}

function escapeSvg(s: string) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

function indexOfBytes(hay: Uint8Array, needle: number[], from = 0): number {
  // Native indexOf for the first byte keeps this fast on multi-megabyte EMFs.
  let i = from;
  const last = hay.length - needle.length;
  while (i <= last) {
    i = hay.indexOf(needle[0], i);
    if (i < 0 || i > last) return -1;
    let ok = true;
    for (let j = 1; j < needle.length; j++) if (hay[i + j] !== needle[j]) { ok = false; break; }
    if (ok) return i;
    i++;
  }
  return -1;
}

/** Try to obtain a browser-renderable image URL from EMF/WMF bytes. */
export function emfToDataUrl(bytes: Uint8Array, ext: string): string | null {
  try {
    // 1) Embedded PNG (EMF+ image records embed the original PNG for pasted images).
    const png = indexOfBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    if (png >= 0) {
      const iend = indexOfBytes(bytes, [0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82], png);
      if (iend > 0) {
        const slice = bytes.slice(png, iend + 8);
        return URL.createObjectURL(new Blob([slice as BlobPart], { type: "image/png" }));
      }
    }
    // 2) Embedded JPEG.
    const jpg = indexOfBytes(bytes, [0xff, 0xd8, 0xff]);
    if (jpg >= 0) {
      const eoi = lastIndexOfBytes(bytes, [0xff, 0xd9]);
      if (eoi > jpg) {
        const slice = bytes.slice(jpg, eoi + 2);
        return URL.createObjectURL(new Blob([slice as BlobPart], { type: "image/jpeg" }));
      }
    }
    // 3) Largest DIB in EMR_STRETCHDIBITS / EMR_BITBLT / EMR_SETDIBITSTODEVICE.
    if (ext === "emf") {
      const dib = largestDib(bytes);
      if (dib) return URL.createObjectURL(new Blob([dib as BlobPart], { type: "image/bmp" }));
    }
  } catch { /* fall through */ }
  return null;
}

function lastIndexOfBytes(hay: Uint8Array, needle: number[]): number {
  let i = hay.length - needle.length;
  while (i >= 0) {
    i = hay.lastIndexOf(needle[0], i);
    if (i < 0) return -1;
    let ok = true;
    for (let j = 1; j < needle.length; j++) if (hay[i + j] !== needle[j]) { ok = false; break; }
    if (ok) return i;
    i--;
  }
  return -1;
}

/** Walk EMF records and build a .bmp from the largest device-independent bitmap found. */
function largestDib(bytes: Uint8Array): Uint8Array | null {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 88 || dv.getUint32(0, true) !== 1 /* EMR_HEADER */) return null;
  let off = 0;
  let best: { bmiOff: number; bmiLen: number; bitsOff: number; bitsLen: number; area: number } | null = null;
  while (off + 8 <= bytes.length) {
    const type = dv.getUint32(off, true);
    const size = dv.getUint32(off + 4, true);
    if (size < 8 || off + size > bytes.length) break;
    let rec: [number, number, number, number] | null = null;
    // Offsets (from record start) of offBmiSrc, cbBmiSrc, offBitsSrc, cbBitsSrc
    if (type === 81 /* STRETCHDIBITS */) rec = [48, 52, 56, 60];
    else if (type === 80 /* SETDIBITSTODEVICE */) rec = [48, 52, 56, 60];
    else if (type === 76 /* BITBLT */) rec = [84, 88, 92, 96];
    else if (type === 77 /* STRETCHBLT */) rec = [92, 96, 100, 104];
    if (rec && off + rec[3] + 4 <= bytes.length) {
      const bmiOff = dv.getUint32(off + rec[0], true), bmiLen = dv.getUint32(off + rec[1], true);
      const bitsOff = dv.getUint32(off + rec[2], true), bitsLen = dv.getUint32(off + rec[3], true);
      if (bmiLen >= 40 && bitsLen > 0 && off + bmiOff + bmiLen <= bytes.length && off + bitsOff + bitsLen <= bytes.length) {
        const w = dv.getInt32(off + bmiOff + 4, true), h = Math.abs(dv.getInt32(off + bmiOff + 8, true));
        const area = w * h;
        if (!best || area > best.area) best = { bmiOff: off + bmiOff, bmiLen, bitsOff: off + bitsOff, bitsLen, area };
      }
    }
    if (type === 14 /* EOF */) break;
    off += size;
  }
  if (!best) return null;
  const fileSize = 14 + best.bmiLen + best.bitsLen;
  const out = new Uint8Array(fileSize);
  const o = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d;
  o.setUint32(2, fileSize, true);
  o.setUint32(10, 14 + best.bmiLen, true);
  out.set(bytes.subarray(best.bmiOff, best.bmiOff + best.bmiLen), 14);
  out.set(bytes.subarray(best.bitsOff, best.bitsOff + best.bitsLen), 14 + best.bmiLen);
  return out;
}

/** Read an image file's natural size (px). */
export function imageSize(url: string): Promise<{ w: number; h: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => reject(new Error("Cannot load image"));
    img.src = url;
  });
}
